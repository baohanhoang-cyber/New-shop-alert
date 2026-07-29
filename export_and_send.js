// export_and_send.js (CommonJS)
//
// Flow:
// 1. Read "Tracker" sheet, find rows where column T ("Need alert?") = "New" or "Update".
// 2. Map matched rows (New first, then Update) into "New Shop Template" sheet, starting row 3.
// 3. Mark those Tracker rows' column T as processed (DONE_VALUE, default "No").
// 4. Send a SeaTalk text message: fixed "Dear DSM, WH, Opex team," template with
//    role-based mention tags (DSM_EMAILS / WH_EMAILS / OPEX_EMAILS).
// 5. Screenshot "New Shop Template" (A1:M{lastRow}) -> PDF -> PNG -> trim whitespace -> send to SeaTalk.

const { execSync } = require("node:child_process");
const { writeFileSync, readFileSync, existsSync } = require("node:fs");
const { JWT } = require("google-auth-library");

// ---------------------------------------------------------------------------
// ENV
// ---------------------------------------------------------------------------
const {
  SA_JSON_BASE64,
  SHEET_ID,
  SEA_URL,                              // SeaTalk webhook (secret)
  PNG_NAME = "Report.png",
  PORTRAIT = "true",
  FITW = "true",
  GRIDLINES = "false",
  MAX_BYTES_MB = "5",
  SCALE_TO_PX = "1600",
  USE_LOCAL_IMAGE = "0",                // set to "1" to bypass export flow and use a local PNG
  LOCAL_IMAGE_PATH = "",
  TRACKER_SHEET_NAME = "Tracker",
  TEMPLATE_SHEET_NAME = "New Shop Template",
  DONE_VALUE = "No",                    // value written back into Tracker!T after processing
} = process.env;

function need(v, name) {
  if (!v) { console.error(`Missing env: ${name}`); process.exit(1); }
}
need(SA_JSON_BASE64, "SA_JSON_BASE64");
need(SHEET_ID, "SHEET_ID");
need(SEA_URL, "SEA_URL");

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function colLetterToIndex(letter) {
  // "A" -> 1, "T" -> 20, "AB" -> 28, ...
  let n = 0;
  for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}

function cellValue(row, letter) {
  const idx = colLetterToIndex(letter) - 1;
  const v = row[idx];
  return v == null ? "" : String(v).trim();
}

// Tracker column -> field name (see mapping screenshot)
const TRACKER_COLS = {
  needAlert: "T",       // Need alert? (New/Update/No)
  shopName: "F",
  shopId: "E",
  mainCAT: "I",
  ado: "J",
  inboundWH: "W",
  noSKUs: "Q",          // default "To be Updated" if empty
  itemQty: "R",         // default "To be Updated" if empty
  estIBdate: "S",
  inventoryCAT: "AB",
  bulky: "AC",
  inventoryZone: "AD",
  allowShopeeMTF: "AF",
};

// Column order A..M in "New Shop Template"
const TEMPLATE_FIELD_ORDER = [
  "status", "shopName", "shopId", "mainCAT", "ado", "inboundWH",
  "noSKUs", "itemQty", "estIBdate", "inventoryCAT", "bulky",
  "inventoryZone", "allowShopeeMTF",
];

const TEMPLATE_HEADER_ROWS = 2;   // row1 = group header, row2 = column header
const TEMPLATE_MIN_LAST_ROW = 10; // default screenshot range is A1:M10, extend if more shops

const DSM_EMAILS = [
  "luu.phamnhat@shopee.com",
  "duy.phunganh@shopee.com",
  "trang.vungoc@shopee.com",
];
const WH_EMAILS = [
  "thanhngan.nguyenngoc@spxexpress.com",
  "manhquan.tran@shopee.com",
  "hongnhung.phannguyen@spxexpress.com",
];
const OPEX_EMAILS = [
  "ngoc.nguyenhongbao@shopee.com",
];
function buildMentionTags(emails) {
  return emails.map(e => `<mention-tag target="seatalk://user?email=${e}"/>`).join(" ");
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
(async () => {
  try {
    // --- Auth ---
    const sa = JSON.parse(Buffer.from(SA_JSON_BASE64, "base64").toString("utf8"));
    const jwt = new JWT({
      email: sa.client_email,
      key: sa.private_key,
      scopes: [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets",
      ],
    });
    const tokenObj = await jwt.getAccessToken();
    const token = tokenObj && tokenObj.token;
    if (!token) { console.error("Failed to obtain access token"); process.exit(1); }
    const authHeaders = { Authorization: `Bearer ${token}` };

    // --- Sheet metadata: name -> sheetId ---
    const metaResp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?fields=sheets.properties`,
      { headers: authHeaders }
    );
    if (!metaResp.ok) {
      console.error("Failed to fetch spreadsheet metadata:", await metaResp.text());
      process.exit(1);
    }
    const meta = await metaResp.json();
    const sheetIdByName = {};
    for (const s of meta.sheets) sheetIdByName[s.properties.title] = s.properties.sheetId;

    const templateSheetId = sheetIdByName[TEMPLATE_SHEET_NAME];
    if (templateSheetId == null) {
      console.error(`Sheet "${TEMPLATE_SHEET_NAME}" not found in spreadsheet.`);
      process.exit(1);
    }

    // --- Read Tracker sheet (columns A:AF) ---
    const trackerResp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(`${TRACKER_SHEET_NAME}!A:AF`)}`,
      { headers: authHeaders }
    );
    if (!trackerResp.ok) {
      console.error("Failed to read Tracker sheet:", await trackerResp.text());
      process.exit(1);
    }
    const trackerJson = await trackerResp.json();
    const trackerValues = trackerJson.values || [];

    // Row 1 = header, data starts at row 2 (index 1)
    const newShops = [];
    const updateShops = [];
    for (let idx = 1; idx < trackerValues.length; idx++) {
      const row = trackerValues[idx];
      const rowNumber = idx + 1;
      const needAlert = cellValue(row, TRACKER_COLS.needAlert).toLowerCase();
      if (needAlert !== "new" && needAlert !== "update") continue;

      const shop = {
        rowNumber,
        status: needAlert === "new" ? "New" : "Update",
        shopName: cellValue(row, TRACKER_COLS.shopName),
        shopId: cellValue(row, TRACKER_COLS.shopId),
        mainCAT: cellValue(row, TRACKER_COLS.mainCAT),
        ado: cellValue(row, TRACKER_COLS.ado),
        inboundWH: cellValue(row, TRACKER_COLS.inboundWH),
        noSKUs: cellValue(row, TRACKER_COLS.noSKUs) || "To be Updated",
        itemQty: cellValue(row, TRACKER_COLS.itemQty) || "To be Updated",
        estIBdate: cellValue(row, TRACKER_COLS.estIBdate),
        inventoryCAT: cellValue(row, TRACKER_COLS.inventoryCAT),
        bulky: cellValue(row, TRACKER_COLS.bulky),
        inventoryZone: cellValue(row, TRACKER_COLS.inventoryZone),
        allowShopeeMTF: cellValue(row, TRACKER_COLS.allowShopeeMTF),
      };
      if (needAlert === "new") newShops.push(shop);
      else updateShops.push(shop);
    }

    const combinedShops = [...newShops, ...updateShops];
    console.log(`Found ${newShops.length} New + ${updateShops.length} Update = ${combinedShops.length} shop(s) to map.`);

    if (combinedShops.length === 0) {
      console.log("No shops marked New/Update. Nothing to do. Exiting.");
      process.exit(0);
    }

    // --- Write mapped rows into "New Shop Template" (starting row 3) ---
    const dataStartRow = TEMPLATE_HEADER_ROWS + 1; // row 3
    const dataEndRow = dataStartRow + combinedShops.length - 1;
    const lastRow = Math.max(TEMPLATE_MIN_LAST_ROW, dataEndRow);

    // Clear previous data body first, so stale rows from earlier runs don't linger
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(`${TEMPLATE_SHEET_NAME}!A${dataStartRow}:M2000`)}:clear`,
      { method: "POST", headers: { ...authHeaders, "Content-Type": "application/json" }, body: JSON.stringify({}) }
    );

    const templateRows = combinedShops.map(shop => TEMPLATE_FIELD_ORDER.map(f => shop[f]));
    const writeResp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(`${TEMPLATE_SHEET_NAME}!A${dataStartRow}:M${dataEndRow}`)}?valueInputOption=RAW`,
      {
        method: "PUT",
        headers: { ...authHeaders, "Content-Type": "application/json" },
        body: JSON.stringify({ values: templateRows }),
      }
    );
    if (!writeResp.ok) {
      console.error("Failed to write New Shop Template rows:", await writeResp.text());
      process.exit(1);
    }
    console.log(`Wrote ${templateRows.length} row(s) into ${TEMPLATE_SHEET_NAME}!A${dataStartRow}:M${dataEndRow}`);

    // --- Mark processed Tracker rows: column T = DONE_VALUE ---
    const trackerUpdateData = combinedShops.map(shop => ({
      range: `${TRACKER_SHEET_NAME}!T${shop.rowNumber}`,
      values: [[DONE_VALUE]],
    }));
    const markResp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values:batchUpdate`,
      {
        method: "POST",
        headers: { ...authHeaders, "Content-Type": "application/json" },
        body: JSON.stringify({ valueInputOption: "RAW", data: trackerUpdateData }),
      }
    );
    if (!markResp.ok) {
      console.warn("Failed to mark Tracker rows as processed:", await markResp.text());
    } else {
      console.log(`Marked ${trackerUpdateData.length} Tracker row(s) column T as "${DONE_VALUE}".`);
    }

    // --- Build text message: fixed template + role-based mention tags ---
    const dsmMentionTags = buildMentionTags(DSM_EMAILS);
    const whMentionTags = buildMentionTags(WH_EMAILS);
    const opexMentionTags = buildMentionTags(OPEX_EMAILS);

    const finalText =
      "Dear DSM, WH, Opex team,\n" +
      `DSM: ${dsmMentionTags}\n` +
      `WH: ${whMentionTags}\n` +
      `Opex: ${opexMentionTags}\n` +
      "Welcome onboard the new premium shop with the following infomation as attachment below";

    try {
      const textPayload = { tag: "text", text: { content: finalText } };
      const tResp = await fetch(SEA_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(textPayload),
      });
      console.log("Sent text to SeaTalk, status:", tResp.status);
      console.log("SeaTalk text response:", await tResp.text());
    } catch (e) {
      console.warn("Failed to send text to SeaTalk:", e);
    }

    // --- Screenshot "New Shop Template" (A1:M{lastRow}) -> PNG ---
    let pngBuffer = null;
    let tempSheetId = null;
    let createdTemp = false;

    if (String(USE_LOCAL_IMAGE) === "1" && LOCAL_IMAGE_PATH) {
      if (!existsSync(LOCAL_IMAGE_PATH)) {
        console.error("Local image not found at path:", LOCAL_IMAGE_PATH);
        process.exit(1);
      }
      pngBuffer = readFileSync(LOCAL_IMAGE_PATH);
      console.log("Read local PNG bytes:", pngBuffer.length);
    } else {
      // 1) Duplicate the Template sheet
      const dupName = `tmp_export_${Date.now()}`;
      const dupBody = {
        requests: [{ duplicateSheet: { sourceSheetId: templateSheetId, insertSheetIndex: 0, newSheetName: dupName } }],
      };
      let resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
        method: "POST",
        headers: { ...authHeaders, "Content-Type": "application/json" },
        body: JSON.stringify(dupBody),
      });
      if (!resp.ok) {
        console.error("Failed to duplicate sheet:", resp.status, await resp.text());
        process.exit(1);
      }
      const dupData = await resp.json();
      tempSheetId = dupData.replies[0].duplicateSheet.properties.sheetId;
      const gridRows = dupData.replies[0].duplicateSheet.properties.gridProperties.rowCount;
      const gridCols = dupData.replies[0].duplicateSheet.properties.gridProperties.columnCount;
      createdTemp = true;

      // 2) Crop to A1:M{lastRow}
      const endIndexRowExclusive = lastRow;
      const endIndexColExclusive = 13; // A..M

      const cropRequests = [];
      if (endIndexRowExclusive < gridRows) {
        cropRequests.push({
          deleteDimension: { range: { sheetId: tempSheetId, dimension: "ROWS", startIndex: endIndexRowExclusive, endIndex: gridRows } },
        });
      }
      if (endIndexColExclusive < gridCols) {
        cropRequests.push({
          deleteDimension: { range: { sheetId: tempSheetId, dimension: "COLUMNS", startIndex: endIndexColExclusive, endIndex: gridCols } },
        });
      }
      if (cropRequests.length > 0) {
        resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
          method: "POST",
          headers: { ...authHeaders, "Content-Type": "application/json" },
          body: JSON.stringify({ requests: cropRequests }),
        });
        if (!resp.ok) {
          console.error("Failed to crop temp sheet:", resp.status, await resp.text());
          await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
            method: "POST",
            headers: { ...authHeaders, "Content-Type": "application/json" },
            body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] }),
          }).catch(() => {});
          process.exit(1);
        }
      }

      // 3) Export temp sheet as PDF
      const exportUrl =
        `https://docs.google.com/spreadsheets/d/${encodeURIComponent(SHEET_ID)}/export` +
        `?exportFormat=pdf&gid=${encodeURIComponent(tempSheetId)}` +
        `&portrait=${PORTRAIT}&fitw=${FITW}&gridlines=${GRIDLINES}` +
        `&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0`;
      const pdfResp = await fetch(exportUrl, { headers: authHeaders });
      if (!pdfResp.ok) {
        console.error("Export PDF failed:", await pdfResp.text());
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
          method: "POST",
          headers: { ...authHeaders, "Content-Type": "application/json" },
          body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] }),
        }).catch(() => {});
        process.exit(1);
      }
      const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
      writeFileSync("report.pdf", pdfBuf);

      // 4) PDF -> PNG (poppler)
      const scale = Number(SCALE_TO_PX) || 1600;
      execSync(`pdftoppm -png -singlefile -scale-to ${scale} report.pdf report`, { stdio: "inherit" });

      // 5) Trim whitespace (ImageMagick)
      try {
        execSync(`convert report.png -fuzz 4% -trim +repage report_trim.png`, { stdio: "inherit" });
        pngBuffer = readFileSync("report_trim.png");
      } catch (err) {
        console.warn("Trim failed, falling back to original report.png:", err);
        pngBuffer = readFileSync("report.png");
      }

      // 6) Shrink if too large
      const maxBytes = (Number(MAX_BYTES_MB) || 5) * 1024 * 1024;
      if (pngBuffer.length > maxBytes) {
        const scale2 = Math.max(600, Math.floor(scale * 0.75));
        execSync(`pdftoppm -png -singlefile -scale-to ${scale2} report.pdf report_small`, { stdio: "inherit" });
        try {
          execSync(`convert report_small.png -fuzz 4% -trim +repage report_small_trim.png`, { stdio: "inherit" });
          pngBuffer = readFileSync("report_small_trim.png");
        } catch {
          pngBuffer = readFileSync("report_small.png");
        }
      }
    }

    // --- Send PNG to SeaTalk ---
    if (!pngBuffer) { console.error("No PNG buffer prepared."); process.exit(1); }
    const filePayload = { tag: "file", file: { filename: PNG_NAME, content: pngBuffer.toString("base64") } };
    const fileResp = await fetch(SEA_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filePayload),
    });
    console.log("SeaTalk file status:", fileResp.status);
    console.log("SeaTalk file response:", await fileResp.text());

    // --- Clear the data just mapped into New Shop Template (reset for next run) ---
    const clearAfterResp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(`${TEMPLATE_SHEET_NAME}!A${dataStartRow}:M2000`)}:clear`,
      { method: "POST", headers: { ...authHeaders, "Content-Type": "application/json" }, body: JSON.stringify({}) }
    );
    if (!clearAfterResp.ok) {
      console.warn("Failed to clear New Shop Template after sending:", await clearAfterResp.text());
    } else {
      console.log(`Cleared ${TEMPLATE_SHEET_NAME}!A${dataStartRow}:M2000 after sending.`);
    }

    // --- Cleanup temp sheet ---
    if (createdTemp && tempSheetId) {
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
        method: "POST",
        headers: { ...authHeaders, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] }),
      }).catch(err => console.warn("Failed to delete temp sheet:", err));
      console.log("Temp sheet cleanup attempted.");
    }

    console.log("All done.");
    process.exit(0);
  } catch (e) {
    console.error("Fatal error:", e);
    process.exit(1);
  }
})();
