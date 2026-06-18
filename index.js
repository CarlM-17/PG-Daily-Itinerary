const express = require("express");
const { google } = require("googleapis");

const app = express();
const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID =
  process.env.SPREADSHEET_ID || "13hgFrJPHbzrfXb2-tMHJ5lZbtQhSVvyJwRwYaTi9aII";
const DAILY_SHEET_NAME = process.env.DAILY_SHEET_NAME || "Daily Sheet";
const ITINERARY_SHEET_NAME = process.env.ITINERARY_SHEET_NAME || "Itinerary";
const TIMESHEET_SHEET_NAME = process.env.TIMESHEET_SHEET_NAME || "TimeSheet";
const STORE_SHEET_NAME = process.env.STORE_SHEET_NAME || "Store";

app.use(express.json());

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getPrivateKey() {
  const jsonCredentials = getServiceAccountJson();

  if (jsonCredentials && jsonCredentials.private_key) {
    return normalizePrivateKey(jsonCredentials.private_key);
  }

  if (process.env.GOOGLE_PRIVATE_KEY_BASE64) {
    return normalizePrivateKey(
      Buffer.from(process.env.GOOGLE_PRIVATE_KEY_BASE64, "base64").toString("utf8")
    );
  }

  const key =
    process.env.GOOGLE_PRIVATE_KEY ||
    process.env.Google_Private_Key ||
    process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY;

  if (key) {
    return normalizePrivateKey(key);
  }

  return "";
}

function normalizePrivateKey(value) {
  let key = String(value || "").trim();

  if (
    (key.startsWith('"') && key.endsWith('"')) ||
    (key.startsWith("'") && key.endsWith("'"))
  ) {
    key = key.slice(1, -1);
  }

  return key.replace(/\\n/g, "\n");
}

function getClientEmail() {
  const jsonCredentials = getServiceAccountJson();

  if (jsonCredentials && jsonCredentials.client_email) {
    return jsonCredentials.client_email;
  }

  const email =
    process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL ||
    process.env.Google_Service_Account_email ||
    process.env.GOOGLE_CLIENT_EMAIL ||
    process.env.CLIENT_EMAIL;

  if (email) {
    return String(email).trim();
  }

  return "";
}

function getServiceAccountJson() {
  const json =
    process.env.GOOGLE_SERVICE_ACCOUNT_JSON ||
    process.env.Google_Service_Account_JSON ||
    process.env.GOOGLE_CREDENTIALS_JSON;

  if (json) {
    return JSON.parse(json);
  }

  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64) {
    const decoded = Buffer.from(
      process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64,
      "base64"
    ).toString("utf8");
    return JSON.parse(decoded);
  }

  return null;
}

function getSheetsClient() {
  const clientEmail = getClientEmail();
  const privateKey = getPrivateKey();

  if (!clientEmail || !privateKey) {
    throw new Error(
      "Missing Google credentials. Set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY, or GOOGLE_SERVICE_ACCOUNT_JSON."
    );
  }

  if (!privateKey.includes("-----BEGIN PRIVATE KEY-----")) {
    throw new Error(
      "Invalid Google private key. Use the service account private_key value, not private_key_id."
    );
  }

  const auth = new google.auth.JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
}

function quoteSheetName(sheetName) {
  return `'${String(sheetName).replaceAll("'", "''")}'`;
}

function normalizeSheetTitle(sheetName) {
  return String(sheetName || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

async function resolveSheetInfo(sheets, preferredName) {
  const metadata = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties(sheetId,title)",
  });

  const sheetInfos = (metadata.data.sheets || [])
    .map((sheet) => sheet.properties)
    .filter((properties) => properties && properties.title);
  const titles = sheetInfos.map((sheet) => sheet.title);
  const preferred = String(preferredName || "").trim();
  const exact = sheetInfos.find((sheet) => sheet.title === preferred);

  if (exact) return exact;

  const normalizedPreferred = normalizeSheetTitle(preferred);
  const normalized = sheetInfos.find(
    (sheet) => normalizeSheetTitle(sheet.title) === normalizedPreferred
  );

  if (normalized) return normalized;

  const error = new Error(
    `Sheet tab "${preferred}" was not found. Available tabs: ${titles.join(", ")}`
  );
  error.status = 400;
  throw error;
}

async function resolveSheetName(sheets, preferredName) {
  const sheetInfo = await resolveSheetInfo(sheets, preferredName);
  return sheetInfo.title;
}

function todayInLocalInputFormat() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  const local = new Date(now.getTime() - offset * 60 * 1000);
  return local.toISOString().slice(0, 10);
}

function calculateDuration(timeIn, timeOut) {
  if (!timeIn || !timeOut) return "";

  const [inHours, inMinutes] = timeIn.split(":").map(Number);
  const [outHours, outMinutes] = timeOut.split(":").map(Number);

  if (
    Number.isNaN(inHours) ||
    Number.isNaN(inMinutes) ||
    Number.isNaN(outHours) ||
    Number.isNaN(outMinutes)
  ) {
    return "";
  }

  let start = inHours * 60 + inMinutes;
  let end = outHours * 60 + outMinutes;

  if (end < start) end += 24 * 60;

  const totalMinutes = end - start;
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;

  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function parseTimeToMinutes(value) {
  const text = String(value || "").trim();

  if (!text) return null;

  const amPmMatch = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);

  if (amPmMatch) {
    let hours = Number(amPmMatch[1]);
    const minutes = Number(amPmMatch[2]);
    const period = amPmMatch[3].toUpperCase();

    if (period === "PM" && hours < 12) hours += 12;
    if (period === "AM" && hours === 12) hours = 0;

    return hours * 60 + minutes;
  }

  const hourMinuteMatch = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);

  if (hourMinuteMatch) {
    return Number(hourMinuteMatch[1]) * 60 + Number(hourMinuteMatch[2]);
  }

  return null;
}

function minutesToClock(totalMinutes) {
  const normalized = ((totalMinutes % 1440) + 1440) % 1440;
  const hours = Math.floor(normalized / 60);
  const minutes = normalized % 60;

  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function minutesToDuration(startMinutes, endMinutes) {
  let adjustedEnd = endMinutes;

  if (adjustedEnd < startMinutes) {
    adjustedEnd += 24 * 60;
  }

  const totalMinutes = adjustedEnd - startMinutes;
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;

  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function formatItineraryDate(inputDate) {
  const [year, month, day] = String(inputDate || "").split("-").map(Number);

  if (!year || !month || !day) {
    return inputDate;
  }

  return new Intl.DateTimeFormat("en-US", {
    month: "long",
    day: "numeric",
    year: "numeric",
  }).format(new Date(year, month - 1, day));
}

function normalizeDateKey(value) {
  const text = String(value || "").trim();

  if (!text) return "";

  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    return [
      isoMatch[1],
      isoMatch[2].padStart(2, "0"),
      isoMatch[3].padStart(2, "0"),
    ].join("-");
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    return [
      slashMatch[3],
      slashMatch[1].padStart(2, "0"),
      slashMatch[2].padStart(2, "0"),
    ].join("-");
  }

  const parsed = new Date(text);

  if (!Number.isNaN(parsed.getTime())) {
    return [
      parsed.getFullYear(),
      String(parsed.getMonth() + 1).padStart(2, "0"),
      String(parsed.getDate()).padStart(2, "0"),
    ].join("-");
  }

  return text.toLowerCase();
}

async function readStores() {
  try {
    return await readStoresWithSheetsApi();
  } catch (error) {
    return readStoresFromPublicCsv();
  }
}

async function readStoresWithSheetsApi() {
  const sheets = getSheetsClient();
  const sheetName = await resolveSheetName(sheets, STORE_SHEET_NAME);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetName)}!A1:A`,
  });

  const rows = response.data.values || [];
  return cleanStoreRows(rows);
}

async function readStoresFromPublicCsv() {
  const url =
    `https://docs.google.com/spreadsheets/d/${encodeURIComponent(SPREADSHEET_ID)}` +
    `/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(STORE_SHEET_NAME)}`;
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error(
      "Unable to load Store sheet. Check credentials or make the spreadsheet visible to anyone with the link."
    );
  }

  const csv = await response.text();
  const rows = parseCsv(csv);
  return cleanStoreRows(rows);
}

function cleanStoreRows(rows) {
  const stores = rows
    .flat()
    .map((store) => String(store || "").trim())
    .filter(Boolean);

  if (stores[0] && stores[0].toLowerCase() === "store") {
    stores.shift();
  }

  return [...new Set(stores)];
}

function parseCsv(csv) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < csv.length; index += 1) {
    const char = csv[index];
    const next = csv[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      value += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else {
      value += char;
    }
  }

  if (value || row.length) {
    row.push(value);
    rows.push(row);
  }

  return rows;
}

async function appendVisit({ date, store, timeIn, timeOut, note }) {
  const sheets = getSheetsClient();
  const sheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const duration = calculateDuration(timeIn, timeOut);

  if (!date || !store || !timeIn || !timeOut) {
    const error = new Error("Date, Store, Time In, and Time Out are required.");
    error.status = 400;
    throw error;
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A1:F1`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[date, store, timeIn, timeOut, duration, note || ""]],
    },
  });

  await applyDailySheetFormatting(sheets, sheetInfo.sheetId);
  await updateItinerarySheet(sheets, date, store);
  await updateTimeSheet(sheets, date);

  return { date, store, timeIn, timeOut, duration, note: note || "" };
}

function normalizeDailyRows(rows) {
  return rows
    .map((row) => [
      String(row.date ?? row[0] ?? "").trim(),
      String(row.store ?? row[1] ?? "").trim(),
      String(row.timeIn ?? row[2] ?? "").trim(),
      String(row.timeOut ?? row[3] ?? "").trim(),
      String(row.duration ?? row[4] ?? "").trim(),
      String(row.note ?? row[5] ?? "").trim(),
    ])
    .filter((row) => row.some(Boolean));
}

async function readDailyRows() {
  const sheets = getSheetsClient();
  const sheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A2:F`,
  });

  return normalizeDailyRows(response.data.values || []).map((row) => ({
    date: row[0],
    store: row[1],
    timeIn: row[2],
    timeOut: row[3],
    duration: row[4],
    note: row[5],
  }));
}

async function replaceDailyRows(rows) {
  const sheets = getSheetsClient();
  const dailySheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const normalizedRows = normalizeDailyRows(rows).map((row) => [
    row[0],
    row[1],
    row[2],
    row[3],
    calculateDuration(row[2], row[3]) || row[4],
    row[5],
  ]);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(dailySheetInfo.title)}!A2:F`,
  });

  if (normalizedRows.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(dailySheetInfo.title)}!A2:F${normalizedRows.length + 1}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: normalizedRows,
      },
    });
  }

  await applyDailySheetFormatting(sheets, dailySheetInfo.sheetId);
  await rebuildDerivedSheetsFromDailyRows(sheets, normalizedRows);

  return normalizedRows.map((row) => ({
    date: row[0],
    store: row[1],
    timeIn: row[2],
    timeOut: row[3],
    duration: row[4],
    note: row[5],
  }));
}

async function rebuildDerivedSheetsFromDailyRows(sheets, rows) {
  await rebuildItinerarySheetFromDailyRows(sheets, rows);
  await rebuildTimeSheetFromDailyRows(sheets, rows);
}

async function rebuildItinerarySheetFromDailyRows(sheets, rows) {
  const sheetInfo = await resolveSheetInfo(sheets, ITINERARY_SHEET_NAME);
  const byDate = new Map();

  for (const row of rows) {
    const date = row[0];
    const store = row[1];
    const dateKey = normalizeDateKey(date);

    if (!date || !store || !dateKey) continue;

    if (!byDate.has(dateKey)) {
      byDate.set(dateKey, { date, stores: [] });
    }

    const group = byDate.get(dateKey);

    if (!group.stores.includes(store) && group.stores.length < 8) {
      group.stores.push(store);
    }
  }

  const values = Array.from(byDate.values()).map((group) => {
    const row = [formatItineraryDate(group.date), ...group.stores];

    while (row.length < 9) row.push("");

    return row;
  });

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A2:I`,
  });

  if (values.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(sheetInfo.title)}!A2:I${values.length + 1}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values },
    });
  }

  await applyItinerarySheetFormatting(sheets, sheetInfo.sheetId);
}

async function rebuildTimeSheetFromDailyRows(sheets, rows) {
  const sheetInfo = await resolveSheetInfo(sheets, TIMESHEET_SHEET_NAME);
  const byDate = new Map();

  for (const row of rows) {
    const date = row[0];
    const dateKey = normalizeDateKey(date);
    const timeIn = parseTimeToMinutes(row[2]);
    const timeOut = parseTimeToMinutes(row[3]);

    if (!date || !dateKey || timeIn === null || timeOut === null) continue;

    if (!byDate.has(dateKey)) {
      byDate.set(dateKey, {
        date,
        firstTimeIn: timeIn,
        lastTimeOut: timeOut,
      });
    }

    const group = byDate.get(dateKey);
    group.firstTimeIn = Math.min(group.firstTimeIn, timeIn);
    group.lastTimeOut = Math.max(group.lastTimeOut, timeOut);
  }

  const values = Array.from(byDate.values()).map((group) => [
    group.date,
    "",
    minutesToClock(group.firstTimeIn),
    minutesToClock(group.lastTimeOut),
    minutesToDuration(group.firstTimeIn, group.lastTimeOut),
  ]);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A2:E`,
  });

  if (values.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(sheetInfo.title)}!A2:E${values.length + 1}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values },
    });
  }

  await applyTimeSheetFormatting(sheets, sheetInfo.sheetId);
}

async function updateTimeSheet(sheets, date) {
  const dailySheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const timeSheetInfo = await resolveSheetInfo(sheets, TIMESHEET_SHEET_NAME);
  const targetDateKey = normalizeDateKey(date);
  const dailyResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(dailySheetInfo.title)}!A:E`,
  });
  const dailyRows = dailyResponse.data.values || [];
  const sameDateLogs = dailyRows
    .slice(1)
    .filter((row) => normalizeDateKey(row[0]) === targetDateKey)
    .map((row) => ({
      timeIn: parseTimeToMinutes(row[2]),
      timeOut: parseTimeToMinutes(row[3]),
    }))
    .filter((log) => log.timeIn !== null && log.timeOut !== null);

  if (!sameDateLogs.length) return;

  const firstTimeIn = Math.min(...sameDateLogs.map((log) => log.timeIn));
  const lastTimeOut = Math.max(...sameDateLogs.map((log) => log.timeOut));
  const rowValues = [
    date,
    "",
    minutesToClock(firstTimeIn),
    minutesToClock(lastTimeOut),
    minutesToDuration(firstTimeIn, lastTimeOut),
  ];
  const timeSheetResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(timeSheetInfo.title)}!A:E`,
  });
  const timeSheetRows = timeSheetResponse.data.values || [];
  const targetRowIndex = timeSheetRows.findIndex(
    (row, index) => index > 0 && normalizeDateKey(row[0]) === targetDateKey
  );

  if (targetRowIndex === -1) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(timeSheetInfo.title)}!A1:E1`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [rowValues],
      },
    });
  } else {
    const rowNumber = targetRowIndex + 1;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(timeSheetInfo.title)}!A${rowNumber}:E${rowNumber}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [rowValues],
      },
    });
  }

  await applyTimeSheetFormatting(sheets, timeSheetInfo.sheetId);
}

async function applyTimeSheetFormatting(sheets, sheetId) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 5,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.09, green: 0.38, blue: 0.13 },
                horizontalAlignment: "LEFT",
                textFormat: {
                  foregroundColor: { red: 1, green: 1, blue: 1 },
                  bold: true,
                },
              },
            },
            fields:
              "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 5,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 1, green: 1, blue: 1 },
                horizontalAlignment: "LEFT",
                textFormat: {
                  foregroundColor: { red: 0, green: 0, blue: 0 },
                  bold: false,
                },
              },
            },
            fields:
              "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
          },
        },
      ],
    },
  });
}

async function updateItinerarySheet(sheets, date, store) {
  const sheetInfo = await resolveSheetInfo(sheets, ITINERARY_SHEET_NAME);
  const itineraryDate = formatItineraryDate(date);
  const itineraryDateKey = normalizeDateKey(date);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A:I`,
  });
  const rows = response.data.values || [];
  const matchingRowIndexes = rows
    .map((row, index) => ({ row, index }))
    .filter(
      ({ row, index }) =>
        index > 0 && normalizeDateKey(row[0]) === itineraryDateKey
    )
    .map(({ index }) => index);

  if (!matchingRowIndexes.length) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(sheetInfo.title)}!A1:I1`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[itineraryDate, store]],
      },
    });
    await applyItinerarySheetFormatting(sheets, sheetInfo.sheetId);
    return;
  }

  const targetRowIndex = matchingRowIndexes[0];
  const existingStores = [];

  for (const rowIndex of matchingRowIndexes) {
    const row = rows[rowIndex] || [];

    for (const value of row.slice(1, 9)) {
      const storeName = String(value || "").trim();

      if (storeName && !existingStores.includes(storeName)) {
        existingStores.push(storeName);
      }
    }
  }

  if (!existingStores.includes(store)) {
    existingStores.push(store);
  }

  if (existingStores.length > 8) {
    const error = new Error(
      `Itinerary already has 8 stores for ${itineraryDate}.`
    );
    error.status = 400;
    throw error;
  }

  const rowNumber = targetRowIndex + 1;
  const rowValues = [itineraryDate, ...existingStores];

  while (rowValues.length < 9) {
    rowValues.push("");
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A${rowNumber}:I${rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [rowValues],
    },
  });

  for (const duplicateRowIndex of matchingRowIndexes.slice(1)) {
    const duplicateRowNumber = duplicateRowIndex + 1;
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(sheetInfo.title)}!A${duplicateRowNumber}:I${duplicateRowNumber}`,
    });
  }

  await applyItinerarySheetFormatting(sheets, sheetInfo.sheetId);
}

async function applyItinerarySheetFormatting(sheets, sheetId) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 9,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.22, green: 0.43, blue: 0.14 },
                horizontalAlignment: "CENTER",
                textFormat: {
                  foregroundColor: { red: 1, green: 1, blue: 1 },
                  bold: true,
                },
              },
            },
            fields:
              "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 9,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 1, green: 1, blue: 1 },
                horizontalAlignment: "LEFT",
                textFormat: {
                  foregroundColor: { red: 0, green: 0, blue: 0 },
                  bold: false,
                },
              },
            },
            fields:
              "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
          },
        },
      ],
    },
  });
}

async function applyDailySheetFormatting(sheets, sheetId) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 6,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.09, green: 0.38, blue: 0.13 },
                horizontalAlignment: "LEFT",
                textFormat: {
                  foregroundColor: { red: 1, green: 1, blue: 1 },
                  bold: true,
                },
              },
            },
            fields:
              "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 6,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 1, green: 1, blue: 1 },
                horizontalAlignment: "LEFT",
                textFormat: {
                  foregroundColor: { red: 0, green: 0, blue: 0 },
                  bold: false,
                },
              },
            },
            fields:
              "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
          },
        },
      ],
    },
  });
}

app.get("/api/stores", async (req, res) => {
  try {
    const stores = await readStores();
    res.json({ stores });
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

app.get("/api/daily", async (req, res) => {
  try {
    const rows = await readDailyRows();
    res.json({ rows });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.put("/api/daily", async (req, res) => {
  try {
    const rows = await replaceDailyRows(Array.isArray(req.body.rows) ? req.body.rows : []);
    res.json({ rows });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.post("/api/visits", async (req, res) => {
  try {
    const saved = await appendVisit({
      date: String(req.body.date || "").trim(),
      store: String(req.body.store || "").trim(),
      timeIn: String(req.body.timeIn || "").trim(),
      timeOut: String(req.body.timeOut || "").trim(),
      note: String(req.body.note || "").trim(),
    });
    res.status(201).json({ visit: saved });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

function formatApiError(error) {
  const message = String(error.message || "");

  if (
    message.includes("invalid_grant") ||
    message.includes("Invalid JWT Signature")
  ) {
    return (
      "Google credential error: invalid JWT signature. The service account email " +
      "and private key do not match, or the key was deleted/disabled. Use client_email " +
      "and private_key from the same service account JSON file."
    );
  }

  if (
    message.includes("DECODER routines::unsupported") ||
    message.includes("PEM routines") ||
    message.includes("private key")
  ) {
    return (
      "Google credential error: the private key is invalid. In Railway, set " +
      "GOOGLE_SERVICE_ACCOUNT_EMAIL to the service account email and GOOGLE_PRIVATE_KEY " +
      "to the matching private_key value including BEGIN/END PRIVATE KEY. Do not use private_key_id."
    );
  }

  return message || "Unable to save visit.";
}

app.get("/", (req, res) => {
  const today = escapeHtml(todayInLocalInputFormat());

  res.type("html").send(`<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PG Daily Itinerary Store Visits</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #1b2633;
      --muted: #5d6b7a;
      --line: #d8dee6;
      --field: #f7f9fc;
      --page: #edf1f6;
      --panel: #ffffff;
      --accent: #0f766e;
      --accent-dark: #0b5f59;
      --danger: #b42318;
      --success: #067647;
      --shadow: 0 20px 50px rgba(31, 43, 59, 0.12);
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      min-height: 100vh;
      font-family: Arial, Helvetica, sans-serif;
      background:
        linear-gradient(135deg, rgba(15, 118, 110, 0.10), transparent 34%),
        linear-gradient(315deg, rgba(217, 119, 6, 0.12), transparent 36%),
        var(--page);
      color: var(--ink);
    }

    main {
      width: min(960px, calc(100% - 32px));
      margin: 0 auto;
      padding: 36px 0;
    }

    .topbar {
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      gap: 16px;
      margin-bottom: 18px;
    }

    h1 {
      margin: 0;
      font-size: clamp(28px, 4vw, 42px);
      line-height: 1.08;
      letter-spacing: 0;
    }

    .sheet-link {
      color: var(--accent-dark);
      font-weight: 700;
      text-decoration: none;
      white-space: nowrap;
    }

    .shell {
      display: grid;
      grid-template-columns: 1.4fr 0.8fr;
      gap: 18px;
      align-items: start;
    }

    .panel {
      background: var(--panel);
      border: 1px solid rgba(216, 222, 230, 0.85);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }

    form {
      padding: 22px;
    }

    .grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
    }

    label {
      display: grid;
      gap: 7px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 700;
    }

    input,
    select,
    textarea {
      width: 100%;
      min-height: 44px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--field);
      color: var(--ink);
      font: inherit;
      font-size: 15px;
      padding: 10px 12px;
      outline: none;
    }

    input:focus,
    select:focus,
    textarea:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(15, 118, 110, 0.15);
    }

    textarea {
      min-height: 114px;
      resize: vertical;
    }

    .span-2 {
      grid-column: 1 / -1;
    }

    .duration-box {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      min-height: 44px;
      padding: 10px 12px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: #fff8eb;
      color: #7a4f01;
      font-weight: 800;
    }

    .actions {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 12px;
      margin-top: 18px;
    }

    button {
      min-height: 44px;
      border: 0;
      border-radius: 7px;
      background: var(--accent);
      color: white;
      font: inherit;
      font-weight: 800;
      padding: 0 18px;
      cursor: pointer;
    }

    button:hover {
      background: var(--accent-dark);
    }

    button:disabled {
      cursor: not-allowed;
      opacity: 0.7;
    }

    .capture-button {
      width: 100%;
      min-height: 54px;
      font-size: 18px;
    }

    .capture-button.secondary {
      background: #315f8c;
    }

    .capture-button.secondary:hover {
      background: #25496c;
    }

    .compact-actions {
      justify-content: flex-start;
    }

    .status {
      min-height: 20px;
      font-size: 14px;
      color: var(--muted);
    }

    .status.error {
      color: var(--danger);
    }

    .status.success {
      color: var(--success);
    }

    .summary {
      padding: 20px;
    }

    .summary h2 {
      margin: 0 0 14px;
      font-size: 18px;
      letter-spacing: 0;
    }

    .summary dl {
      display: grid;
      gap: 12px;
      margin: 0;
    }

    .summary div {
      display: grid;
      gap: 3px;
      padding-bottom: 12px;
      border-bottom: 1px solid var(--line);
    }

    .summary div:last-child {
      border-bottom: 0;
      padding-bottom: 0;
    }

    dt {
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      text-transform: uppercase;
    }

    dd {
      margin: 0;
      overflow-wrap: anywhere;
      font-weight: 700;
    }

    .table-panel {
      margin-top: 18px;
      overflow: hidden;
    }

    .table-toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
    }

    .table-toolbar h2 {
      margin: 0;
      font-size: 18px;
    }

    .table-actions {
      display: flex;
      gap: 10px;
    }

    .table-wrap {
      overflow-x: auto;
      background: #ffffff;
      -webkit-overflow-scrolling: touch;
    }

    table {
      width: 100%;
      min-width: 760px;
      border-collapse: collapse;
      color: #000000;
      font-size: 14px;
    }

    th {
      background: #176122;
      color: #ffffff;
      font-weight: 800;
      text-align: left;
      border: 1px solid #d9d9d9;
      padding: 4px 6px;
    }

    td {
      min-width: 110px;
      height: 26px;
      background: #ffffff;
      color: #000000;
      border: 1px solid #d9d9d9;
      padding: 3px 6px;
      outline: none;
    }

    td:focus {
      box-shadow: inset 0 0 0 2px #1a73e8;
    }

    .table-input,
    .table-select {
      width: 100%;
      min-height: 24px;
      border: 0;
      border-radius: 0;
      background: #ffffff;
      color: #000000;
      font: inherit;
      padding: 0;
      outline: none;
    }

    .table-input:focus,
    .table-select:focus {
      box-shadow: inset 0 0 0 2px #1a73e8;
    }

    .duration-cell {
      display: block;
      min-height: 24px;
      line-height: 24px;
    }

    @media (max-width: 760px) {
      body {
        background: var(--page);
      }

      main {
        width: min(100% - 12px, 960px);
        padding: 12px 0 18px;
      }

      .topbar,
      .actions,
      .table-toolbar {
        align-items: stretch;
        flex-direction: column;
      }

      .topbar {
        gap: 10px;
        margin-bottom: 12px;
      }

      h1 {
        font-size: 30px;
        line-height: 1.04;
      }

      .sheet-link,
      button {
        width: 100%;
        text-align: center;
      }

      .shell,
      .grid {
        grid-template-columns: 1fr;
      }

      form,
      .summary {
        padding: 16px;
      }

      input,
      select,
      textarea {
        min-height: 48px;
        font-size: 16px;
      }

      .capture-button {
        min-height: 58px;
      }

      .summary h2,
      .table-toolbar h2 {
        font-size: 17px;
      }

      .table-panel {
        margin-top: 12px;
      }

      .table-toolbar {
        padding: 12px;
      }

      .table-wrap {
        border-top: 1px solid var(--line);
      }

      table {
        min-width: 680px;
        font-size: 13px;
      }

      th,
      td {
        min-width: 104px;
        padding: 4px;
      }

      th:first-child,
      td:first-child {
        position: sticky;
        left: 0;
        z-index: 2;
      }

      td:first-child {
        background: #ffffff;
      }

      th:first-child {
        z-index: 3;
      }

      .table-input,
      .table-select,
      .duration-cell {
        min-height: 34px;
        font-size: 16px;
      }

      .table-input,
      .table-select {
        padding: 4px 2px;
      }
    }

    @media (max-width: 420px) {
      h1 {
        font-size: 25px;
      }

      form,
      .summary {
        padding: 12px;
      }

      .grid {
        gap: 12px;
      }

      textarea {
        min-height: 92px;
      }

      table {
        min-width: 620px;
      }

      th,
      td {
        min-width: 96px;
      }

      th:nth-child(6),
      td:nth-child(6) {
        min-width: 132px;
      }
    }
  </style>
</head>
<body>
  <main>
    <div class="topbar">
      <h1>PG Daily Itinerary Store Visits</h1>
      <a class="sheet-link" href="https://docs.google.com/spreadsheets/d/${escapeHtml(SPREADSHEET_ID)}/edit?gid=0#gid=0" target="_blank" rel="noreferrer">Open Google Sheet</a>
    </div>

    <section class="shell">
      <form class="panel" id="visitForm">
        <div class="grid">
          <label>
            Date
            <input id="date" name="date" type="date" value="${today}" required>
          </label>

          <label>
            Store
            <select id="store" name="store" required>
              <option value="">Loading stores...</option>
            </select>
          </label>

          <label>
            Time In
            <button id="timeInButton" class="capture-button" type="button">Time In</button>
          </label>

          <label>
            Time Out
            <button id="timeOutButton" class="capture-button secondary" type="button" disabled>Time Out</button>
          </label>

          <label>
            Duration
            <div class="duration-box"><span id="duration">00:00</span><span>HH:MM</span></div>
          </label>

          <label class="span-2">
            Note optional
            <textarea id="note" name="note" placeholder="Add visit notes here"></textarea>
          </label>
        </div>

        <div class="actions compact-actions">
          <div class="status" id="status" role="status" aria-live="polite"></div>
        </div>
      </form>

      <aside class="panel summary">
        <h2>Latest Entry</h2>
        <dl>
          <div><dt>Date</dt><dd id="previewDate">-</dd></div>
          <div><dt>Store</dt><dd id="previewStore">-</dd></div>
          <div><dt>Time</dt><dd id="previewTime">-</dd></div>
          <div><dt>Duration</dt><dd id="previewDuration">-</dd></div>
          <div><dt>Note</dt><dd id="previewNote">-</dd></div>
        </dl>
      </aside>
    </section>

    <section class="panel table-panel">
      <div class="table-toolbar">
        <h2>Daily Sheet</h2>
        <div class="status" id="tableStatus" role="status" aria-live="polite"></div>
      </div>
      <div class="table-wrap">
        <table id="dailyTable" aria-label="Daily Sheet editable table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Store</th>
              <th>Time_In</th>
              <th>Time_Out</th>
              <th>Duration</th>
              <th>Note</th>
            </tr>
          </thead>
          <tbody id="dailyTableBody">
            <tr><td colspan="6">Loading Daily Sheet...</td></tr>
          </tbody>
        </table>
      </div>
    </section>
  </main>

  <script>
    const form = document.getElementById("visitForm");
    const dateInput = document.getElementById("date");
    const storeInput = document.getElementById("store");
    const timeInButton = document.getElementById("timeInButton");
    const timeOutButton = document.getElementById("timeOutButton");
    const noteInput = document.getElementById("note");
    const durationOutput = document.getElementById("duration");
    const statusOutput = document.getElementById("status");
    const tableStatusOutput = document.getElementById("tableStatus");
    const dailyTableBody = document.getElementById("dailyTableBody");
    let currentEntry = null;
    let dailyRows = [];
    let storeOptions = [];
    let tableSaveTimer = null;

    function localDateValue() {
      const now = new Date();
      const offset = now.getTimezoneOffset();
      return new Date(now.getTime() - offset * 60000).toISOString().slice(0, 10);
    }

    function localTimeValue() {
      const now = new Date();
      return String(now.getHours()).padStart(2, "0") + ":" + String(now.getMinutes()).padStart(2, "0");
    }

    function setAutomaticDate() {
      dateInput.value = localDateValue();
    }

    function setStatus(message, type = "") {
      statusOutput.textContent = message;
      statusOutput.className = "status" + (type ? " " + type : "");
    }

    function setTableStatus(message, type = "") {
      tableStatusOutput.textContent = message;
      tableStatusOutput.className = "status" + (type ? " " + type : "");
    }

    function durationBetween(timeIn, timeOut) {
      if (!timeIn || !timeOut) return "00:00";
      const [inHours, inMinutes] = timeIn.split(":").map(Number);
      const [outHours, outMinutes] = timeOut.split(":").map(Number);
      let start = inHours * 60 + inMinutes;
      let end = outHours * 60 + outMinutes;
      if (end < start) end += 24 * 60;
      const total = end - start;
      const hours = Math.floor(total / 60);
      const minutes = total % 60;
      return String(hours).padStart(2, "0") + ":" + String(minutes).padStart(2, "0");
    }

    function refreshPreview() {
      const timeIn = currentEntry ? currentEntry.timeIn : "";
      const timeOut = currentEntry ? currentEntry.timeOut : "";
      const duration = durationBetween(timeIn, timeOut);
      durationOutput.textContent = duration;
      document.getElementById("previewDate").textContent = currentEntry ? currentEntry.date : dateInput.value || "-";
      document.getElementById("previewStore").textContent = currentEntry ? currentEntry.store : storeInput.value || "-";
      document.getElementById("previewTime").textContent =
        timeIn && timeOut ? timeIn + " - " + timeOut : timeIn ? timeIn + " - pending" : "-";
      document.getElementById("previewDuration").textContent = duration;
      document.getElementById("previewNote").textContent = currentEntry ? currentEntry.note || "-" : noteInput.value || "-";
    }

    function setPunchState(hasOpenEntry) {
      timeInButton.disabled = hasOpenEntry;
      timeOutButton.disabled = !hasOpenEntry;
      dateInput.disabled = hasOpenEntry;
      storeInput.disabled = hasOpenEntry;
    }

    function renderDailyTable() {
      const rows = currentEntry ? dailyRows.concat([currentEntry]) : dailyRows;

      if (!rows.length) {
        dailyTableBody.innerHTML = '<tr><td colspan="6">No rows yet.</td></tr>';
        return;
      }

      dailyTableBody.innerHTML = "";

      rows.forEach((row, rowIndex) => {
        const isPending = Boolean(currentEntry && row === currentEntry);
        const tr = document.createElement("tr");
        const editableIndex = isPending ? -1 : rowIndex;
        const dateCell = document.createElement("td");
        const dateControl = document.createElement("input");
        dateControl.className = "table-input";
        dateControl.type = "date";
        dateControl.value = row.date || "";
        dateControl.disabled = isPending;
        dateControl.dataset.rowIndex = editableIndex;
        dateControl.dataset.field = "date";
        dateCell.appendChild(dateControl);
        tr.appendChild(dateCell);

        const storeCell = document.createElement("td");
        const storeControl = document.createElement("select");
        storeControl.className = "table-select";
        storeControl.disabled = isPending;
        storeControl.dataset.rowIndex = editableIndex;
        storeControl.dataset.field = "store";
        appendStoreOptions(storeControl, row.store || "");
        storeCell.appendChild(storeControl);
        tr.appendChild(storeCell);

        for (const field of ["timeIn", "timeOut"]) {
          const timeCell = document.createElement("td");
          const timeControl = document.createElement("input");
          timeControl.className = "table-input";
          timeControl.type = "time";
          timeControl.value = row[field] || "";
          timeControl.disabled = isPending;
          timeControl.dataset.rowIndex = editableIndex;
          timeControl.dataset.field = field;
          timeCell.appendChild(timeControl);
          tr.appendChild(timeCell);
        }

        const durationCell = document.createElement("td");
        const durationText = document.createElement("span");
        durationText.className = "duration-cell";
        durationText.textContent = durationBetween(row.timeIn, row.timeOut);
        durationCell.appendChild(durationText);
        tr.appendChild(durationCell);

        const noteCell = document.createElement("td");
        const noteControl = document.createElement("input");
        noteControl.className = "table-input";
        noteControl.type = "text";
        noteControl.value = row.note || "";
        noteControl.disabled = isPending;
        noteControl.dataset.rowIndex = editableIndex;
        noteControl.dataset.field = "note";
        noteCell.appendChild(noteControl);
        tr.appendChild(noteCell);

        dailyTableBody.appendChild(tr);
      });
    }

    function appendStoreOptions(select, selectedStore) {
      const blankOption = document.createElement("option");
      blankOption.value = "";
      blankOption.textContent = "Select store";
      select.appendChild(blankOption);

      const options = storeOptions.slice();

      if (selectedStore && !options.includes(selectedStore)) {
        options.push(selectedStore);
      }

      for (const store of options) {
        const option = document.createElement("option");
        option.value = store;
        option.textContent = store;
        option.selected = store === selectedStore;
        select.appendChild(option);
      }
    }

    function getEditedTableRows() {
      return dailyRows
        .map((row) => ({
          date: row.date || "",
          store: row.store || "",
          timeIn: row.timeIn || "",
          timeOut: row.timeOut || "",
          duration: durationBetween(row.timeIn, row.timeOut),
          note: row.note || ""
        }))
        .filter((row) => row.date || row.store || row.timeIn || row.timeOut || row.note);
    }

    function queueTableAutoSave() {
      clearTimeout(tableSaveTimer);
      setTableStatus("Auto-saving...");
      tableSaveTimer = setTimeout(saveTableChanges, 700);
    }

    async function saveTableChanges() {
      try {
        const response = await fetch("/api/daily", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ rows: getEditedTableRows() })
        });
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to save table changes.");

        dailyRows = data.rows || [];
        setTableStatus("Auto-saved.", "success");
        renderDailyTable();
      } catch (error) {
        setTableStatus(error.message, "error");
      }
    }

    async function loadDailyRows() {
      try {
        const response = await fetch("/api/daily");
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to load Daily Sheet.");

        dailyRows = data.rows || [];
        renderDailyTable();
      } catch (error) {
        dailyTableBody.innerHTML = '<tr><td colspan="6">' + error.message + '</td></tr>';
      }
    }

    async function loadStores() {
      try {
        const response = await fetch("/api/stores");
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to load stores.");

        storeInput.innerHTML = '<option value="">Select store</option>';

        if (!data.stores.length) {
          storeInput.innerHTML = '<option value="">No stores found on Store sheet</option>';
          return;
        }

        storeOptions = data.stores;

        for (const store of data.stores) {
          const option = document.createElement("option");
          option.value = store;
          option.textContent = store;
          storeInput.appendChild(option);
        }

        renderDailyTable();
      } catch (error) {
        storeInput.innerHTML = '<option value="">Store list unavailable</option>';
        setStatus(error.message, "error");
      }
    }

    noteInput.addEventListener("input", () => {
      if (currentEntry) {
        currentEntry.note = noteInput.value;
        renderDailyTable();
      }

      refreshPreview();
    });

    form.addEventListener("submit", (event) => {
      event.preventDefault();
    });

    timeInButton.addEventListener("click", () => {
      if (!dateInput.value || !storeInput.value) {
        setStatus("Select a date and store before Time In.", "error");
        return;
      }

      currentEntry = {
        date: dateInput.value,
        store: storeInput.value,
        timeIn: localTimeValue(),
        timeOut: "",
        duration: "",
        note: noteInput.value.trim()
      };

      setPunchState(true);
      setStatus("Time In captured. Press Time Out when visit is complete.", "success");
      renderDailyTable();
      refreshPreview();
    });

    timeOutButton.addEventListener("click", async () => {
      if (!currentEntry) return;

      currentEntry.timeOut = localTimeValue();
      currentEntry.duration = durationBetween(currentEntry.timeIn, currentEntry.timeOut);
      currentEntry.note = noteInput.value.trim();
      timeOutButton.disabled = true;
      setStatus("Saving...");
      renderDailyTable();
      refreshPreview();

      try {
        const response = await fetch("/api/visits", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            date: currentEntry.date,
            store: currentEntry.store,
            timeIn: currentEntry.timeIn,
            timeOut: currentEntry.timeOut,
            note: currentEntry.note
          })
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.message || "Unable to save visit.");

        setStatus("Saved to Daily Sheet.", "success");
        currentEntry = null;
        setPunchState(false);
        setAutomaticDate();
        noteInput.value = "";
        await loadDailyRows();
        refreshPreview();
      } catch (error) {
        setStatus(error.message, "error");
        setPunchState(true);
      } finally {
        renderDailyTable();
      }
    });

    dailyTableBody.addEventListener("input", (event) => {
      const control = event.target.closest("[data-row-index]");
      if (!control) return;

      const rowIndex = Number(control.dataset.rowIndex);
      const field = control.dataset.field;

      if (rowIndex < 0 || !dailyRows[rowIndex] || !field) return;

      dailyRows[rowIndex][field] = control.value;

      if (field === "timeIn" || field === "timeOut") {
        dailyRows[rowIndex].duration = durationBetween(
          dailyRows[rowIndex].timeIn,
          dailyRows[rowIndex].timeOut
        );
        renderDailyTable();
      }

      queueTableAutoSave();
    });

    dailyTableBody.addEventListener("change", (event) => {
      const control = event.target.closest("[data-row-index]");
      if (!control) return;

      const rowIndex = Number(control.dataset.rowIndex);
      const field = control.dataset.field;

      if (rowIndex < 0 || !dailyRows[rowIndex] || !field) return;

      dailyRows[rowIndex][field] = control.value;

      if (field === "timeIn" || field === "timeOut") {
        dailyRows[rowIndex].duration = durationBetween(
          dailyRows[rowIndex].timeIn,
          dailyRows[rowIndex].timeOut
        );
        renderDailyTable();
      }

      queueTableAutoSave();
    });

    setAutomaticDate();
    setPunchState(false);
    loadStores();
    loadDailyRows();
    refreshPreview();
  </script>
</body>
</html>`);
});

app.listen(PORT, () => {
  console.log(`PG Daily Itinerary Store Visits is running on port ${PORT}`);
});
