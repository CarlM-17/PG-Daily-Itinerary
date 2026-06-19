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
const STORE_LIST_SHEET_NAME = process.env.STORE_LIST_SHEET_NAME || "StoreList";
const ITINERARY_SUMMARY_SHEET_NAME =
  process.env.ITINERARY_SUMMARY_SHEET_NAME || "Itinerary Summary";
const DAILY_BACKUP_SHEET_NAME = process.env.DAILY_BACKUP_SHEET_NAME || "Daily Backup";
const BACKUP_MIN_INTERVAL_MS = 10 * 60 * 1000;
let lastDailyBackupAt = 0;

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

async function ensureSheetInfo(sheets, preferredName) {
  try {
    return await resolveSheetInfo(sheets, preferredName);
  } catch (error) {
    if (error.status !== 400) throw error;
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: {
              title: preferredName,
            },
          },
        },
      ],
    },
  });

  return resolveSheetInfo(sheets, preferredName);
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

function durationToMinutes(value) {
  const match = String(value || "").trim().match(/^(\d+):(\d{2})$/);

  if (!match) return null;

  return Number(match[1]) * 60 + Number(match[2]);
}

function minutesToDurationText(totalMinutes) {
  const safeMinutes = Math.max(Number(totalMinutes) || 0, 0);
  const hours = Math.floor(safeMinutes / 60);
  const minutes = safeMinutes % 60;

  if (hours && minutes) return `${hours}h ${minutes}m`;
  if (hours) return `${hours}h`;
  return `${minutes}m`;
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

  if (stores[0] && ["store", "stores", "storelist", "store list"].includes(stores[0].toLowerCase())) {
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

  await backupDailySheet(sheets);

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A1:F1`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[date, store, timeIn, timeOut, duration, note || ""]],
    },
  });

  await updateItinerarySheet(sheets, date, store);
  await updateTimeSheet(sheets, date);
  await rebuildItinerarySummaryFromSheets(sheets);

  return { date, store, timeIn, timeOut, duration, note: note || "" };
}

function getUpdatedRowNumber(updatedRange) {
  const match = String(updatedRange || "").match(/![A-Z]+(\d+):/);
  return match ? Number(match[1]) : null;
}

async function getOpenVisit(sheets = getSheetsClient()) {
  const sheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A2:F`,
  });
  const rows = response.data.values || [];

  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index] || [];
    const visit = {
      rowNumber: index + 2,
      date: String(row[0] || "").trim(),
      store: String(row[1] || "").trim(),
      timeIn: String(row[2] || "").trim(),
      timeOut: String(row[3] || "").trim(),
      duration: String(row[4] || "").trim(),
      note: String(row[5] || "").trim(),
    };

    if (visit.date && visit.store && visit.timeIn && !visit.timeOut) {
      return visit;
    }
  }

  return null;
}

async function startOpenVisit({ date, store, timeIn, note }) {
  const sheets = getSheetsClient();
  const existingOpenVisit = await getOpenVisit(sheets);

  if (existingOpenVisit) {
    return { visit: existingOpenVisit, alreadyOpen: true };
  }

  const sheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);

  if (!date || !store || !timeIn) {
    const error = new Error("Date, Store, and Time In are required.");
    error.status = 400;
    throw error;
  }

  await backupDailySheet(sheets);

  const response = await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A1:F1`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[date, store, timeIn, "", "", note || ""]],
    },
  });
  const rowNumber = getUpdatedRowNumber(response.data.updates && response.data.updates.updatedRange);
  const visit = {
    rowNumber,
    date,
    store,
    timeIn,
    timeOut: "",
    duration: "",
    note: note || "",
  };

  return { visit, alreadyOpen: false };
}

async function endOpenVisit({ rowNumber, timeOut, note }) {
  const sheets = getSheetsClient();
  const sheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const openVisit = rowNumber
    ? await readVisitAtRow(sheets, sheetInfo.title, Number(rowNumber))
    : await getOpenVisit(sheets);

  if (!openVisit || !openVisit.timeIn || openVisit.timeOut) {
    const error = new Error("No open Time In session was found.");
    error.status = 400;
    throw error;
  }

  if (!timeOut) {
    const error = new Error("Time Out is required.");
    error.status = 400;
    throw error;
  }

  const duration = calculateDuration(openVisit.timeIn, timeOut);
  const finalNote = note ?? openVisit.note ?? "";

  await backupDailySheet(sheets);

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!D${openVisit.rowNumber}:F${openVisit.rowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[timeOut, duration, finalNote]],
    },
  });

  await updateItinerarySheet(sheets, openVisit.date, openVisit.store);
  await updateTimeSheet(sheets, openVisit.date);
  await rebuildItinerarySummaryFromSheets(sheets);

  return {
    ...openVisit,
    timeOut,
    duration,
    note: finalNote,
  };
}

async function readVisitAtRow(sheets, sheetName, rowNumber) {
  if (!rowNumber || rowNumber < 2) return null;

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetName)}!A${rowNumber}:F${rowNumber}`,
  });
  const row = (response.data.values || [])[0] || [];

  return {
    rowNumber,
    date: String(row[0] || "").trim(),
    store: String(row[1] || "").trim(),
    timeIn: String(row[2] || "").trim(),
    timeOut: String(row[3] || "").trim(),
    duration: String(row[4] || "").trim(),
    note: String(row[5] || "").trim(),
  };
}

async function updateVisitNote({ rowNumber, note }) {
  const sheets = getSheetsClient();
  const sheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const targetRowNumber = Number(rowNumber);

  if (!targetRowNumber || targetRowNumber < 2) {
    const error = new Error("A valid Daily Sheet row is required to update the note.");
    error.status = 400;
    throw error;
  }

  const visit = await readVisitAtRow(sheets, sheetInfo.title, targetRowNumber);

  if (!visit || !visit.date || !visit.store || !visit.timeIn) {
    const error = new Error("The selected Daily Sheet row was not found.");
    error.status = 404;
    throw error;
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!F${targetRowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[note || ""]],
    },
  });

  return { ...visit, note: note || "" };
}

async function updateDailyRow({ rowNumber, row }) {
  const sheets = getSheetsClient();
  const sheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const targetRowNumber = Number(rowNumber);

  if (!targetRowNumber || targetRowNumber < 2) {
    const error = new Error("A valid Daily Sheet row is required.");
    error.status = 400;
    throw error;
  }

  const normalized = normalizeDailyRows([row])[0];

  if (!normalized || !normalized.some(Boolean)) {
    const error = new Error("Refusing to save an empty Daily Sheet row.");
    error.status = 400;
    throw error;
  }

  const values = [
    normalized[0],
    normalized[1],
    normalized[2],
    normalized[3],
    calculateDuration(normalized[2], normalized[3]) || normalized[4],
    normalized[5],
  ];

  await backupDailySheet(sheets);

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A${targetRowNumber}:F${targetRowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [values],
    },
  });

  if (values[0] && values[1] && values[2] && values[3]) {
    await updateItinerarySheet(sheets, values[0], values[1]);
    await updateTimeSheet(sheets, values[0]);
  }

  return {
    rowNumber: targetRowNumber,
    date: values[0],
    store: values[1],
    timeIn: values[2],
    timeOut: values[3],
    duration: values[4],
    note: values[5],
  };
}

async function backupDailySheet(sheets) {
  const now = Date.now();

  if (now - lastDailyBackupAt < BACKUP_MIN_INTERVAL_MS) {
    return;
  }

  const dailySheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const backupSheetInfo = await ensureSheetInfo(sheets, DAILY_BACKUP_SHEET_NAME);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(dailySheetInfo.title)}!A:F`,
  });
  const values = [
    ["Backup created", new Date().toISOString(), "", "", "", ""],
    ...(response.data.values || []),
  ];

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(backupSheetInfo.title)}!A:F`,
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(backupSheetInfo.title)}!A1:F${values.length}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values },
  });
  lastDailyBackupAt = now;
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

  return (response.data.values || [])
    .map((row, index) => ({
      rowNumber: index + 2,
      date: String(row[0] || "").trim(),
      store: String(row[1] || "").trim(),
      timeIn: String(row[2] || "").trim(),
      timeOut: String(row[3] || "").trim(),
      duration: String(row[4] || "").trim(),
      note: String(row[5] || "").trim(),
    }))
    .filter((row) =>
      row.date || row.store || row.timeIn || row.timeOut || row.duration || row.note
    );
}

async function readTimeSheetSummary({ startDate, endDate }) {
  const sheets = getSheetsClient();
  const sheetInfo = await resolveSheetInfo(sheets, TIMESHEET_SHEET_NAME);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A2:E`,
  });
  const rows = response.data.values || [];
  const requiredMinutes = 9 * 60;
  const records = rows
    .map((row) => {
      const date = String(row[0] || "").trim();
      const timeIn = String(row[2] || row[1] || "").trim();
      const timeOut = String(row[3] || row[2] || "").trim();
      const duration = String(row[4] || row[3] || "").trim();
      const durationMinutes = durationToMinutes(duration);
      const lackingMinutes =
        durationMinutes === null ? 0 : Math.max(requiredMinutes - durationMinutes, 0);
      const normalizedDate = normalizeDateKey(date);

      return {
        date,
        normalizedDate,
        timeIn,
        timeOut,
        duration,
        durationMinutes,
        lackingMinutes,
        remarks: lackingMinutes
          ? `Undertime - lacking ${minutesToDurationText(lackingMinutes)}`
          : "Complete",
      };
    })
    .filter((record) => record.date)
    .filter((record) => {
      if (startDate && record.normalizedDate < startDate) return false;
      if (endDate && record.normalizedDate > endDate) return false;
      return true;
    });
  const undertimeRecords = records.filter((record) => record.lackingMinutes > 0);
  const totalLackingMinutes = undertimeRecords.reduce(
    (total, record) => total + record.lackingMinutes,
    0
  );

  return {
    records,
    totalDates: records.length,
    undertimeDates: undertimeRecords.length,
    completeDates: records.length - undertimeRecords.length,
    totalLackingMinutes,
    totalLackingText: minutesToDurationText(totalLackingMinutes),
  };
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

  if (!normalizedRows.length) {
    const error = new Error(
      "Refusing to save an empty Daily Sheet. Reload the page and try again."
    );
    error.status = 400;
    throw error;
  }

  await backupDailySheet(sheets);

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(dailySheetInfo.title)}!A2:F${normalizedRows.length + 1}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: normalizedRows,
    },
  });

  await applyDailySheetFormatting(sheets, dailySheetInfo.sheetId);
  await rebuildDerivedSheetsFromDailyRows(sheets, normalizedRows);

  return normalizedRows.map((row, index) => ({
    rowNumber: index + 2,
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
  await rebuildItinerarySummarySheet(sheets, rows);
}

function monthKeyFromDate(value) {
  const key = normalizeDateKey(value);

  if (!key || !key.includes("-")) return "";

  return key.slice(0, 7);
}

function formatMonthLabel(monthKey) {
  const [year, month] = String(monthKey || "").split("-").map(Number);

  if (!year || !month) return monthKey;

  return new Intl.DateTimeFormat("en-US", {
    month: "long",
    year: "numeric",
  }).format(new Date(year, month - 1, 1));
}

function currentMonthKey() {
  return todayInLocalInputFormat().slice(0, 7);
}

async function readStoreListRows(sheets) {
  const sheetName = await resolveSheetName(sheets, STORE_LIST_SHEET_NAME);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetName)}!A1:A`,
  });

  return cleanStoreRows(response.data.values || []);
}

async function rebuildItinerarySummarySheet(sheets, dailyRows) {
  const sheetInfo = await ensureSheetInfo(sheets, ITINERARY_SUMMARY_SHEET_NAME);
  const summary = await buildItinerarySummaryData(sheets, dailyRows);
  const storeRows = summary.stores;
  const values = [
    ["Itinerary Summary", "", "", "", "", "", "", "", "", "", ""],
    ["Month", summary.monthLabel, "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", "", ""],
    ["Metric", "Count", "", "Store", "Status", "Visit Count", "Last Visit", "", "Chart Data", "Count", ""],
    summaryMonitorRow("Visited", summary.visitedCount, storeRows[0], "Visited", summary.visitedCount),
    summaryMonitorRow("Not Visited", summary.notVisitedCount, storeRows[1], "Not Visited", summary.notVisitedCount),
    summaryMonitorRow("Total Stores", summary.totalStores, storeRows[2], "", ""),
  ];

  for (let index = 3; index < storeRows.length; index += 1) {
    values.push(summaryMonitorRow("", "", storeRows[index], "", ""));
  }

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A:K`,
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(sheetInfo.title)}!A1:K${values.length}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values },
  });

  await applyItinerarySummaryFormatting(sheets, sheetInfo.sheetId, values.length);
}

async function buildItinerarySummaryData(sheets, dailyRows, targetMonthKey = currentMonthKey()) {
  const requiredStores = await readStoreListRows(sheets);
  const monthLabel = formatMonthLabel(targetMonthKey);
  const visitCounts = new Map();
  const lastVisitDates = new Map();

  for (const row of dailyRows) {
    const date = row[0];
    const store = String(row[1] || "").trim();

    if (!store || monthKeyFromDate(date) !== targetMonthKey) continue;

    visitCounts.set(store, (visitCounts.get(store) || 0) + 1);
    lastVisitDates.set(store, date);
  }

  const storeNames = requiredStores.length
    ? requiredStores
    : Array.from(new Set(dailyRows.map((row) => String(row[1] || "").trim()).filter(Boolean)));
  const stores = storeNames.map((store) => {
    const count = visitCounts.get(store) || 0;

    return {
      store,
      status: count > 0 ? "Visited" : "Not Visited",
      visitCount: count,
      lastVisit: lastVisitDates.get(store) || "",
    };
  });
  const visitedCount = stores.filter((store) => store.visitCount > 0).length;
  const notVisitedCount = Math.max(stores.length - visitedCount, 0);

  return {
    monthLabel,
    visitedCount,
    notVisitedCount,
    totalStores: stores.length,
    stores,
  };
}

function summaryMonitorRow(metric, count, storeRow, chartLabel, chartCount) {
  return [
    metric,
    count,
    "",
    ...summaryStoreRow(storeRow),
    "",
    chartLabel,
    chartCount,
    "",
  ];
}

function summaryStoreRow(storeRow) {
  if (!storeRow) return ["", "", "", ""];

  return [
    storeRow.store,
    storeRow.status,
    storeRow.visitCount,
    storeRow.lastVisit,
  ];
}

async function rebuildItinerarySummaryFromSheets(sheets) {
  const dailySheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(dailySheetInfo.title)}!A2:F`,
  });
  const rows = normalizeDailyRows(response.data.values || []);

  await rebuildItinerarySummarySheet(sheets, rows);
}

async function applyItinerarySummaryFormatting(sheets, sheetId, rowCount) {
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties(sheetId),sheets.charts.chartId",
  });
  const summarySheet = (spreadsheet.data.sheets || []).find(
    (sheet) => sheet.properties && sheet.properties.sheetId === sheetId
  );
  const deleteChartRequests = (summarySheet.charts || []).map((chart) => ({
    deleteEmbeddedObject: {
      objectId: chart.chartId,
    },
  }));

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        ...deleteChartRequests,
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: Math.max(rowCount, 8),
              startColumnIndex: 0,
              endColumnIndex: 11,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 1, green: 1, blue: 1 },
                textFormat: {
                  foregroundColor: { red: 0, green: 0, blue: 0 },
                  bold: false,
                },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 11,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.09, green: 0.38, blue: 0.13 },
                textFormat: {
                  foregroundColor: { red: 1, green: 1, blue: 1 },
                  bold: true,
                  fontSize: 14,
                },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 3,
              endRowIndex: 4,
              startColumnIndex: 0,
              endColumnIndex: 10,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.09, green: 0.38, blue: 0.13 },
                textFormat: {
                  foregroundColor: { red: 1, green: 1, blue: 1 },
                  bold: true,
                },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        },
        {
          autoResizeDimensions: {
            dimensions: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: 0,
              endIndex: 10,
            },
          },
        },
        {
          addChart: {
            chart: {
              spec: {
                title: "Monthly Store Visit Coverage",
                pieChart: {
                  legendPosition: "RIGHT_LEGEND",
                  domain: {
                    sourceRange: {
                      sources: [
                        {
                          sheetId,
                          startRowIndex: 4,
                          endRowIndex: 6,
                          startColumnIndex: 8,
                          endColumnIndex: 9,
                        },
                      ],
                    },
                  },
                  series: {
                    sourceRange: {
                      sources: [
                        {
                          sheetId,
                          startRowIndex: 4,
                          endRowIndex: 6,
                          startColumnIndex: 9,
                          endColumnIndex: 10,
                        },
                      ],
                    },
                  },
                },
              },
              position: {
                overlayPosition: {
                  anchorCell: {
                    sheetId,
                    rowIndex: 1,
                    columnIndex: 10,
                  },
                  widthPixels: 430,
                  heightPixels: 280,
                },
              },
            },
          },
        },
      ],
    },
  });
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

app.patch("/api/daily/row", async (req, res) => {
  try {
    const row = await updateDailyRow({
      rowNumber: req.body.rowNumber,
      row: req.body.row || {},
    });
    res.json({ row });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.get("/api/visits/active", async (req, res) => {
  try {
    const visit = await getOpenVisit();
    res.json({ visit });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.post("/api/visits/start", async (req, res) => {
  try {
    const result = await startOpenVisit({
      date: String(req.body.date || "").trim(),
      store: String(req.body.store || "").trim(),
      timeIn: String(req.body.timeIn || "").trim(),
      note: String(req.body.note || "").trim(),
    });
    res.status(result.alreadyOpen ? 200 : 201).json(result);
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.post("/api/visits/end", async (req, res) => {
  try {
    const visit = await endOpenVisit({
      rowNumber: req.body.rowNumber,
      timeOut: String(req.body.timeOut || "").trim(),
      note: String(req.body.note || "").trim(),
    });
    res.json({ visit });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.patch("/api/visits/note", async (req, res) => {
  try {
    const visit = await updateVisitNote({
      rowNumber: req.body.rowNumber,
      note: String(req.body.note || "").trim(),
    });
    res.json({ visit });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.post("/api/summary/rebuild", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    await rebuildItinerarySummaryFromSheets(sheets);
    res.json({ message: "Itinerary Summary refreshed." });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.get("/api/summary", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const requestedMonth = String(req.query.month || "").trim();
    const monthKey = requestedMonth.match(/^\d{4}-\d{2}$/)
      ? requestedMonth
      : currentMonthKey();
    const dailySheetInfo = await resolveSheetInfo(sheets, DAILY_SHEET_NAME);
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetName(dailySheetInfo.title)}!A2:F`,
    });
    const rows = normalizeDailyRows(response.data.values || []);
    const summary = await buildItinerarySummaryData(sheets, rows, monthKey);

    res.json({ summary });
  } catch (error) {
    const message = formatApiError(error);
    res.status(error.status || 500).json({ message });
  }
});

app.get("/api/timesheet", async (req, res) => {
  try {
    const startDate = String(req.query.start || "").trim();
    const endDate = String(req.query.end || "").trim();
    const isDate = (value) => !value || /^\d{4}-\d{2}-\d{2}$/.test(value);

    if (!isDate(startDate) || !isDate(endDate)) {
      const error = new Error("Use YYYY-MM-DD format for Time Sheet date filters.");
      error.status = 400;
      throw error;
    }

    const data = await readTimeSheetSummary({ startDate, endDate });
    res.json(data);
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
    message.toLowerCase().includes("quota") ||
    message.toLowerCase().includes("rate limit") ||
    message.includes("429")
  ) {
    return "Google Sheets quota exceeded. Wait a minute, then try again. Your Daily Sheet will not be cleared by this app.";
  }

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
      --ink: #132018;
      --muted: #667085;
      --line: #d7ded8;
      --field: #f8faf8;
      --page: #eef3ef;
      --panel: #ffffff;
      --accent: #145c2a;
      --accent-dark: #0b3f1b;
      --accent-soft: #e8f3ec;
      --accent-line: #b9d7c3;
      --danger: #a33a32;
      --danger-soft: #fde8e6;
      --success: #067647;
      --warning: #946200;
      --shadow: 0 16px 38px rgba(19, 32, 24, 0.10);
      --shadow-soft: 0 8px 22px rgba(19, 32, 24, 0.08);
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      min-height: 100vh;
      font-family: Arial, Helvetica, sans-serif;
      background:
        linear-gradient(180deg, rgba(11, 63, 27, 0.10), transparent 300px),
        linear-gradient(135deg, rgba(20, 92, 42, 0.10), transparent 34%),
        var(--page);
      color: var(--ink);
    }

    main {
      width: min(1200px, calc(100% - 32px));
      margin: 0 auto;
      padding: 28px 0 40px;
    }

    .topbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      margin-bottom: 20px;
    }

    h1 {
      margin: 0;
      font-size: clamp(30px, 4vw, 46px);
      line-height: 1.08;
      letter-spacing: 0;
    }

    .eyebrow {
      margin: 0 0 6px;
      color: var(--accent-dark);
      font-size: 12px;
      font-weight: 900;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }

    .subtitle {
      margin: 8px 0 0;
      color: var(--muted);
      font-size: 15px;
      font-weight: 700;
    }

    .sheet-link {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      padding: 0 14px;
      border: 1px solid var(--accent-line);
      border-radius: 999px;
      background: #ffffff;
      color: var(--accent-dark);
      font-weight: 700;
      text-decoration: none;
      white-space: nowrap;
      box-shadow: var(--shadow-soft);
    }

    .shell {
      display: grid;
      grid-template-columns: minmax(0, 1.65fr) minmax(280px, 0.75fr);
      gap: 18px;
      align-items: start;
    }

    .panel {
      background: var(--panel);
      border: 1px solid rgba(215, 222, 216, 0.95);
      border-radius: 14px;
      box-shadow: var(--shadow);
    }

    form {
      padding: 24px;
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
      font-weight: 800;
    }

    .required-mark {
      color: var(--danger);
      font-weight: 900;
    }

    input,
    select,
    textarea {
      width: 100%;
      min-height: 46px;
      border: 1px solid var(--line);
      border-radius: 10px;
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
      background: #ffffff;
      box-shadow: 0 0 0 4px rgba(20, 92, 42, 0.14);
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
      border-radius: 10px;
      background: #fff8e6;
      color: var(--warning);
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
      border-radius: 10px;
      background: var(--accent);
      color: white;
      font: inherit;
      font-weight: 800;
      padding: 0 18px;
      cursor: pointer;
      box-shadow: 0 8px 18px rgba(20, 92, 42, 0.18);
      transition: transform 0.15s ease, box-shadow 0.15s ease, background 0.15s ease;
    }

    button:hover {
      background: var(--accent-dark);
      box-shadow: 0 10px 22px rgba(20, 92, 42, 0.24);
      transform: translateY(-1px);
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
      background: #f3f6f4;
      color: var(--accent-dark);
      border: 1px solid var(--accent-line);
      box-shadow: none;
    }

    .capture-button.secondary:hover {
      background: #e6eee8;
    }

    .compact-actions {
      justify-content: flex-start;
    }

    .status {
      min-height: 24px;
      border-radius: 999px;
      padding: 4px 10px;
      font-size: 13px;
      font-weight: 800;
      color: var(--muted);
    }

    .status.error {
      background: var(--danger-soft);
      color: var(--danger);
    }

    .status.success {
      background: #dcfae6;
      color: var(--success);
    }

    .summary {
      padding: 22px;
    }

    .summary h2 {
      margin: 0 0 14px;
      font-size: 20px;
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
      font-size: 11px;
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

    .tabs {
      display: flex;
      gap: 8px;
      margin-top: 18px;
      padding: 6px;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.76);
      box-shadow: var(--shadow-soft);
    }

    .tab-button {
      width: auto;
      min-height: 42px;
      background: transparent;
      color: var(--accent-dark);
      border: 0;
      box-shadow: none;
    }

    .tab-button.active {
      background: var(--accent);
      color: #ffffff;
      border-color: var(--accent);
    }

    .tab-panel[hidden] {
      display: none;
    }

    .table-toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      padding: 16px 18px;
      border-bottom: 1px solid var(--line);
      background: #fbfdfb;
    }

    .table-toolbar h2 {
      margin: 0;
      font-size: 20px;
    }

    .filter-row {
      display: flex;
      align-items: end;
      flex-wrap: wrap;
      gap: 10px;
    }

    .filter-row label {
      min-width: 150px;
      gap: 5px;
    }

    .filter-row input,
    .filter-row select {
      min-height: 38px;
      background: #ffffff;
    }

    .clear-filter-button {
      min-height: 38px;
      background: #f6f7f6;
      color: var(--accent-dark);
      border: 1px solid var(--line);
      padding: 0 12px;
      box-shadow: none;
    }

    .clear-filter-button:hover {
      background: #eef7f6;
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
      border-collapse: separate;
      border-spacing: 0;
      color: #000000;
      font-size: 14px;
    }

    th {
      background: #176122;
      color: #ffffff;
      font-weight: 800;
      text-align: left;
      border-bottom: 1px solid #0f4618;
      padding: 9px 10px;
      position: sticky;
      top: 0;
      z-index: 4;
    }

    td {
      min-width: 110px;
      height: 38px;
      background: #ffffff;
      color: #000000;
      border-bottom: 1px solid #e5e8e5;
      padding: 6px 10px;
      outline: none;
    }

    tbody tr:nth-child(even) td {
      background: #f8fbf8;
    }

    tbody tr:hover td {
      background: #eef7f0;
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
      background: transparent;
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

    .summary-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      padding: 18px;
      border-bottom: 1px solid var(--line);
      background: #fbfdfb;
    }

    .metric {
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 14px;
      background: #ffffff;
      box-shadow: var(--shadow-soft);
    }

    .metric span {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      text-transform: uppercase;
    }

    .metric strong {
      display: block;
      margin-top: 4px;
      font-size: 30px;
      line-height: 1;
    }

    .chart-box {
      padding: 16px;
      border-bottom: 1px solid var(--line);
    }

    .chart-bar {
      display: flex;
      height: 34px;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: #f8d7da;
    }

    .chart-visited {
      width: var(--visited-width, 0%);
      background: #176122;
    }

    .chart-undertime {
      width: var(--undertime-width, 0%);
      background: #a33a32;
    }

    .chart-legend {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-top: 10px;
      font-size: 13px;
      color: var(--muted);
      font-weight: 700;
    }

    .legend-dot {
      display: inline-block;
      width: 10px;
      height: 10px;
      margin-right: 5px;
      border-radius: 50%;
      background: #176122;
    }

    .legend-dot.missing {
      background: #b42318;
    }

    .status-pill {
      display: inline-block;
      border-radius: 999px;
      padding: 3px 8px;
      font-size: 12px;
      font-weight: 800;
      background: #f8d7da;
      color: #8a1f17;
    }

    .status-pill.visited {
      background: #d1fadf;
      color: #05603a;
    }

    .status-pill.complete {
      background: #d1fadf;
      color: #05603a;
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

      .filter-row {
        align-items: stretch;
        flex-direction: column;
      }

      .filter-row label {
        min-width: 0;
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
      button,
      .tab-button {
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

      .tabs {
        flex-direction: column;
        margin-top: 12px;
      }

      .summary-grid {
        grid-template-columns: 1fr;
        padding: 12px;
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
        background: inherit;
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
      <div>
        <p class="eyebrow">Retail Operations</p>
        <h1>PG Daily Itinerary Store Visits</h1>
        <p class="subtitle">Track store coverage, visit times, and monthly itinerary completion.</p>
      </div>
      <a class="sheet-link" href="https://docs.google.com/spreadsheets/d/${escapeHtml(SPREADSHEET_ID)}/edit?gid=0#gid=0" target="_blank" rel="noreferrer">Open Google Sheet</a>
    </div>

    <section class="shell">
      <form class="panel" id="visitForm">
        <div class="grid">
          <label>
            Date <span class="required-mark">*</span>
            <input id="date" name="date" type="date" value="${today}" required>
          </label>

          <label>
            Store <span class="required-mark">*</span>
            <select id="store" name="store" required>
              <option value="">Loading stores...</option>
            </select>
          </label>

          <label>
            Time In <span class="required-mark">*</span>
            <button id="timeInButton" class="capture-button" type="button">Time In</button>
          </label>

          <label>
            Time Out <span class="required-mark">*</span>
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

    <div class="tabs" role="tablist" aria-label="Workbook views">
      <button id="dailyTabButton" class="tab-button active" type="button" role="tab" aria-selected="true" aria-controls="dailyTabPanel">Daily Sheet</button>
      <button id="summaryTabButton" class="tab-button" type="button" role="tab" aria-selected="false" aria-controls="summaryTabPanel">Itinerary Summary</button>
      <button id="timeSheetTabButton" class="tab-button" type="button" role="tab" aria-selected="false" aria-controls="timeSheetTabPanel">Time Sheet</button>
    </div>

    <section id="dailyTabPanel" class="panel table-panel tab-panel" role="tabpanel">
      <div class="table-toolbar">
        <h2>Daily Sheet</h2>
        <div class="filter-row">
          <label>
            Filter Date
            <input id="dailyDateFilter" type="date">
          </label>
          <label>
            Filter Store
            <select id="dailyStoreFilter">
              <option value="">All stores</option>
            </select>
          </label>
          <button id="clearDailyFiltersButton" class="clear-filter-button" type="button">Clear</button>
        </div>
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

    <section id="summaryTabPanel" class="panel table-panel tab-panel" role="tabpanel" hidden>
      <div class="table-toolbar">
        <h2>Itinerary Summary</h2>
        <div class="filter-row">
          <label>
            Review Month
            <input id="summaryMonthFilter" type="month">
          </label>
        </div>
        <div class="status" id="summaryStatus" role="status" aria-live="polite"></div>
      </div>
      <div class="summary-grid">
        <div class="metric"><span>Month</span><strong id="summaryMonth">-</strong></div>
        <div class="metric"><span>Visited</span><strong id="summaryVisited">0</strong></div>
        <div class="metric"><span>Not Visited</span><strong id="summaryNotVisited">0</strong></div>
      </div>
      <div class="chart-box">
        <div id="summaryChartBar" class="chart-bar" style="--visited-width: 0%">
          <div class="chart-visited"></div>
          <div></div>
        </div>
        <div class="chart-legend">
          <span><span class="legend-dot"></span>Visited</span>
          <span><span class="legend-dot missing"></span>Not Visited</span>
        </div>
      </div>
      <div class="table-wrap">
        <table aria-label="Itinerary Summary table">
          <thead>
            <tr>
              <th>Store</th>
              <th>Status</th>
              <th>Visit Count</th>
              <th>Last Visit</th>
            </tr>
          </thead>
          <tbody id="summaryTableBody">
            <tr><td colspan="4">Loading Itinerary Summary...</td></tr>
          </tbody>
        </table>
      </div>
    </section>

    <section id="timeSheetTabPanel" class="panel table-panel tab-panel" role="tabpanel" hidden>
      <div class="table-toolbar">
        <h2>Time Sheet</h2>
        <div class="filter-row">
          <label>
            Start Date
            <input id="timeSheetStartFilter" type="date">
          </label>
          <label>
            End Date
            <input id="timeSheetEndFilter" type="date">
          </label>
          <button id="clearTimeSheetFiltersButton" class="clear-filter-button" type="button">Current Month</button>
        </div>
        <div class="status" id="timeSheetStatus" role="status" aria-live="polite"></div>
      </div>
      <div class="summary-grid">
        <div class="metric"><span>Total Dates</span><strong id="timeSheetTotalDates">0</strong></div>
        <div class="metric"><span>Undertime Dates</span><strong id="timeSheetUndertimeDates">0</strong></div>
        <div class="metric"><span>Total Lacking</span><strong id="timeSheetTotalLacking">0m</strong></div>
      </div>
      <div class="chart-box">
        <div id="timeSheetChartBar" class="chart-bar" style="--undertime-width: 0%">
          <div class="chart-undertime"></div>
          <div></div>
        </div>
        <div class="chart-legend">
          <span><span class="legend-dot missing"></span>Undertime</span>
          <span><span class="legend-dot"></span>Complete</span>
        </div>
      </div>
      <div class="table-wrap">
        <table aria-label="Time Sheet table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Time In</th>
              <th>Time Out</th>
              <th>Duration</th>
              <th>Remarks</th>
            </tr>
          </thead>
          <tbody id="timeSheetTableBody">
            <tr><td colspan="5">Loading Time Sheet...</td></tr>
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
    const summaryStatusOutput = document.getElementById("summaryStatus");
    const dailyTableBody = document.getElementById("dailyTableBody");
    const summaryTableBody = document.getElementById("summaryTableBody");
    const timeSheetTableBody = document.getElementById("timeSheetTableBody");
    const dailyTabButton = document.getElementById("dailyTabButton");
    const summaryTabButton = document.getElementById("summaryTabButton");
    const timeSheetTabButton = document.getElementById("timeSheetTabButton");
    const dailyTabPanel = document.getElementById("dailyTabPanel");
    const summaryTabPanel = document.getElementById("summaryTabPanel");
    const timeSheetTabPanel = document.getElementById("timeSheetTabPanel");
    const dailyDateFilter = document.getElementById("dailyDateFilter");
    const dailyStoreFilter = document.getElementById("dailyStoreFilter");
    const clearDailyFiltersButton = document.getElementById("clearDailyFiltersButton");
    const summaryMonthFilter = document.getElementById("summaryMonthFilter");
    const timeSheetStartFilter = document.getElementById("timeSheetStartFilter");
    const timeSheetEndFilter = document.getElementById("timeSheetEndFilter");
    const clearTimeSheetFiltersButton = document.getElementById("clearTimeSheetFiltersButton");
    const timeSheetStatusOutput = document.getElementById("timeSheetStatus");
    let currentEntry = null;
    let dailyRows = [];
    let storeOptions = [];
    let tableSaveTimer = null;
    let noteSaveTimer = null;

    function localDateValue() {
      const now = new Date();
      const offset = now.getTimezoneOffset();
      return new Date(now.getTime() - offset * 60000).toISOString().slice(0, 10);
    }

    function currentMonthValue() {
      return localDateValue().slice(0, 7);
    }

    function currentMonthStartValue() {
      return currentMonthValue() + "-01";
    }

    function currentMonthEndValue() {
      const [year, month] = currentMonthValue().split("-").map(Number);
      const end = new Date(year, month, 0);
      return [
        end.getFullYear(),
        String(end.getMonth() + 1).padStart(2, "0"),
        String(end.getDate()).padStart(2, "0")
      ].join("-");
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

    function setSummaryStatus(message, type = "") {
      summaryStatusOutput.textContent = message;
      summaryStatusOutput.className = "status" + (type ? " " + type : "");
    }

    function setTimeSheetStatus(message, type = "") {
      timeSheetStatusOutput.textContent = message;
      timeSheetStatusOutput.className = "status" + (type ? " " + type : "");
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

    function matchesDailyFilters(row) {
      const dateFilter = dailyDateFilter.value;
      const storeFilter = dailyStoreFilter.value;
      const hasSearchFilter = Boolean(dateFilter || storeFilter);

      if (dateFilter && row.date !== dateFilter) return false;
      if (storeFilter && row.store !== storeFilter) return false;
      if (!hasSearchFilter && String(row.date || "").slice(0, 7) !== currentMonthValue()) return false;

      return true;
    }

    function renderDailyTable() {
      const baseRows = dailyRows
        .map((row, index) => ({ row, index }))
        .filter(({ row }) => matchesDailyFilters(row));
      const rows = baseRows;

      if (!rows.length) {
        const message = dailyDateFilter.value || dailyStoreFilter.value
          ? "No rows match the current filters."
          : "No Daily Sheet rows found for the current month.";
        dailyTableBody.innerHTML = '<tr><td colspan="6">' + message + '</td></tr>';
        return;
      }

      dailyTableBody.innerHTML = "";

      rows.forEach(({ row, index }) => {
        const isPending = Boolean(
          currentEntry &&
          Number(row.rowNumber) === Number(currentEntry.rowNumber) &&
          !row.timeOut
        );
        const tr = document.createElement("tr");
        const editableIndex = isPending ? index : index;
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
        noteControl.disabled = false;
        noteControl.dataset.rowIndex = editableIndex;
        noteControl.dataset.field = "note";
        if (isPending) {
          noteControl.dataset.pendingNote = "true";
        }
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

    function queueTableAutoSave(rowIndex) {
      clearTimeout(tableSaveTimer);
      setTableStatus("Auto-saving...");
      tableSaveTimer = setTimeout(() => saveTableChanges(rowIndex), 2000);
    }

    function queueOpenNoteSave() {
      if (!currentEntry || !currentEntry.rowNumber) return;

      clearTimeout(noteSaveTimer);
      setStatus("Saving note...");
      noteSaveTimer = setTimeout(saveOpenVisitNote, 2000);
    }

    async function saveOpenVisitNote() {
      if (!currentEntry || !currentEntry.rowNumber) return;

      try {
        const response = await fetch("/api/visits/note", {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            rowNumber: currentEntry.rowNumber,
            note: currentEntry.note || ""
          })
        });
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to save note.");

        currentEntry.note = data.visit.note || "";
        setStatus("Note saved.", "success");
      } catch (error) {
        setStatus(error.message, "error");
      }
    }

    async function saveTableChanges(rowIndex) {
      const row = dailyRows[rowIndex];

      if (!row || !row.rowNumber) {
        setTableStatus("Unable to auto-save this row. Reload the page and try again.", "error");
        return;
      }

      try {
        const response = await fetch("/api/daily/row", {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            rowNumber: row.rowNumber,
            row
          })
        });
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to save table changes.");

        dailyRows[rowIndex] = data.row;
        setTableStatus("Auto-saved.", "success");
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

    async function restoreActiveVisit() {
      try {
        const response = await fetch("/api/visits/active");
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to check active visit.");

        currentEntry = data.visit || null;

        if (currentEntry) {
          dateInput.value = currentEntry.date || localDateValue();
          noteInput.value = currentEntry.note || "";
          if (currentEntry.store) {
            storeInput.value = currentEntry.store;
          }
          setPunchState(true);
          setStatus("Active Time In restored. Press Time Out when this store visit is complete.", "success");
        } else {
          setPunchState(false);
        }

        renderDailyTable();
        refreshPreview();
      } catch (error) {
        setStatus(error.message, "error");
      }
    }

    async function refreshItinerarySummary() {
      try {
        const response = await fetch("/api/summary/rebuild", {
          method: "POST",
          headers: { "Content-Type": "application/json" }
        });
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to refresh Itinerary Summary.");

        setTableStatus("Itinerary Summary refreshed.", "success");
        await loadSummary();
      } catch (error) {
        setTableStatus(error.message, "error");
        setSummaryStatus(error.message, "error");
      }
    }

    async function loadSummary() {
      try {
        const monthQuery = summaryMonthFilter.value
          ? "?month=" + encodeURIComponent(summaryMonthFilter.value)
          : "";
        const response = await fetch("/api/summary" + monthQuery);
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to load Itinerary Summary.");

        renderSummary(data.summary);
        setSummaryStatus("Updated.", "success");
      } catch (error) {
        setSummaryStatus(error.message, "error");
        summaryTableBody.innerHTML = '<tr><td colspan="4">' + error.message + '</td></tr>';
      }
    }

    function renderSummary(summary) {
      const total = summary.totalStores || 0;
      const visited = summary.visitedCount || 0;
      const visitedWidth = total ? Math.round((visited / total) * 100) : 0;
      document.getElementById("summaryMonth").textContent = summary.monthLabel || "-";
      document.getElementById("summaryVisited").textContent = String(visited);
      document.getElementById("summaryNotVisited").textContent = String(summary.notVisitedCount || 0);
      document.getElementById("summaryChartBar").style.setProperty("--visited-width", visitedWidth + "%");
      summaryTableBody.innerHTML = "";

      if (!summary.stores || !summary.stores.length) {
        summaryTableBody.innerHTML = '<tr><td colspan="4">No stores found in StoreList.</td></tr>';
        return;
      }

      for (const store of summary.stores) {
        const tr = document.createElement("tr");
        const statusClass = store.status === "Visited" ? "visited" : "";
        tr.innerHTML =
          "<td>" + escapeForHtml(store.store) + "</td>" +
          '<td><span class="status-pill ' + statusClass + '">' + escapeForHtml(store.status) + "</span></td>" +
          "<td>" + escapeForHtml(store.visitCount) + "</td>" +
          "<td>" + escapeForHtml(store.lastVisit || "-") + "</td>";
        summaryTableBody.appendChild(tr);
      }
    }

    async function loadTimeSheet() {
      try {
        setTimeSheetStatus("Loading...");
        const params = new URLSearchParams();
        if (timeSheetStartFilter.value) params.set("start", timeSheetStartFilter.value);
        if (timeSheetEndFilter.value) params.set("end", timeSheetEndFilter.value);
        const query = params.toString() ? "?" + params.toString() : "";
        const response = await fetch("/api/timesheet" + query);
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to load Time Sheet.");

        renderTimeSheet(data);
        setTimeSheetStatus("Updated.", "success");
      } catch (error) {
        setTimeSheetStatus(error.message, "error");
        timeSheetTableBody.innerHTML = '<tr><td colspan="5">' + escapeForHtml(error.message) + '</td></tr>';
      }
    }

    function renderTimeSheet(data) {
      const totalDates = data.totalDates || 0;
      const undertimeDates = data.undertimeDates || 0;
      const undertimeWidth = totalDates ? Math.round((undertimeDates / totalDates) * 100) : 0;
      document.getElementById("timeSheetTotalDates").textContent = String(totalDates);
      document.getElementById("timeSheetUndertimeDates").textContent = String(undertimeDates);
      document.getElementById("timeSheetTotalLacking").textContent = data.totalLackingText || "0m";
      document.getElementById("timeSheetChartBar").style.setProperty("--undertime-width", undertimeWidth + "%");
      timeSheetTableBody.innerHTML = "";

      if (!data.records || !data.records.length) {
        timeSheetTableBody.innerHTML = '<tr><td colspan="5">No Time Sheet rows found for the selected period.</td></tr>';
        return;
      }

      for (const record of data.records) {
        const tr = document.createElement("tr");
        const isComplete = record.lackingMinutes === 0;
        tr.innerHTML =
          "<td>" + escapeForHtml(record.date) + "</td>" +
          "<td>" + escapeForHtml(record.timeIn || "-") + "</td>" +
          "<td>" + escapeForHtml(record.timeOut || "-") + "</td>" +
          "<td>" + escapeForHtml(record.duration || "-") + "</td>" +
          '<td><span class="status-pill ' + (isComplete ? "complete" : "") + '">' + escapeForHtml(record.remarks) + "</span></td>";
        timeSheetTableBody.appendChild(tr);
      }
    }

    function escapeForHtml(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
    }

    function setActiveTab(tabName) {
      const showSummary = tabName === "summary";
      const showTimeSheet = tabName === "timesheet";
      const showDaily = !showSummary && !showTimeSheet;
      dailyTabButton.classList.toggle("active", showDaily);
      summaryTabButton.classList.toggle("active", showSummary);
      timeSheetTabButton.classList.toggle("active", showTimeSheet);
      dailyTabButton.setAttribute("aria-selected", String(showDaily));
      summaryTabButton.setAttribute("aria-selected", String(showSummary));
      timeSheetTabButton.setAttribute("aria-selected", String(showTimeSheet));
      dailyTabPanel.hidden = !showDaily;
      summaryTabPanel.hidden = !showSummary;
      timeSheetTabPanel.hidden = !showTimeSheet;

      if (showSummary) {
        loadSummary();
      }

      if (showTimeSheet) {
        loadTimeSheet();
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
        dailyStoreFilter.innerHTML = '<option value="">All stores</option>';

        for (const store of data.stores) {
          const option = document.createElement("option");
          option.value = store;
          option.textContent = store;
          storeInput.appendChild(option);

          const filterOption = document.createElement("option");
          filterOption.value = store;
          filterOption.textContent = store;
          dailyStoreFilter.appendChild(filterOption);
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
        const rowIndex = dailyRows.findIndex(
          (row) => Number(row.rowNumber) === Number(currentEntry.rowNumber)
        );

        if (rowIndex !== -1) {
          dailyRows[rowIndex].note = currentEntry.note;
        }

        queueOpenNoteSave();
        renderDailyTable();
      }

      refreshPreview();
    });

    form.addEventListener("submit", (event) => {
      event.preventDefault();
    });

    timeInButton.addEventListener("click", async () => {
      if (!dateInput.value || !storeInput.value) {
        setStatus("Select a date and store before Time In.", "error");
        return;
      }

      const startedAt = localTimeValue();

      timeInButton.disabled = true;
      setStatus("Saving Time In...");

      try {
        const response = await fetch("/api/visits/start", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            date: dateInput.value,
            store: storeInput.value,
            timeIn: startedAt,
            note: noteInput.value.trim()
          })
        });
        const data = await response.json();

        if (!response.ok) throw new Error(data.message || "Unable to save Time In.");

        currentEntry = data.visit;
        setPunchState(true);
        setStatus(
          data.alreadyOpen
            ? "Active Time In restored. Press Time Out when this store visit is complete."
            : "Time In saved. Press Time Out when this store visit is complete.",
          "success"
        );
        await loadDailyRows();
        renderDailyTable();
        refreshPreview();
      } catch (error) {
        setStatus(error.message, "error");
        setPunchState(false);
      }
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
        const response = await fetch("/api/visits/end", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            rowNumber: currentEntry.rowNumber,
            timeOut: currentEntry.timeOut,
            note: currentEntry.note
          })
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.message || "Unable to save visit.");

        setStatus("Time Out saved. Visit complete.", "success");
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

      if (control.dataset.pendingNote === "true" && field === "note") {
        currentEntry.note = control.value;
        noteInput.value = control.value;
        queueOpenNoteSave();
        refreshPreview();
        return;
      }

      if (field === "timeIn" || field === "timeOut") {
        dailyRows[rowIndex].duration = durationBetween(
          dailyRows[rowIndex].timeIn,
          dailyRows[rowIndex].timeOut
        );
        renderDailyTable();
      }

      queueTableAutoSave(rowIndex);
    });

    dailyTableBody.addEventListener("change", (event) => {
      const control = event.target.closest("[data-row-index]");
      if (!control) return;

      const rowIndex = Number(control.dataset.rowIndex);
      const field = control.dataset.field;

      if (rowIndex < 0 || !dailyRows[rowIndex] || !field) return;

      dailyRows[rowIndex][field] = control.value;

      if (control.dataset.pendingNote === "true" && field === "note") {
        currentEntry.note = control.value;
        noteInput.value = control.value;
        queueOpenNoteSave();
        refreshPreview();
        return;
      }

      if (field === "timeIn" || field === "timeOut") {
        dailyRows[rowIndex].duration = durationBetween(
          dailyRows[rowIndex].timeIn,
          dailyRows[rowIndex].timeOut
        );
        renderDailyTable();
      }

      queueTableAutoSave(rowIndex);
    });

    dailyTabButton.addEventListener("click", () => setActiveTab("daily"));
    summaryTabButton.addEventListener("click", () => setActiveTab("summary"));
    timeSheetTabButton.addEventListener("click", () => setActiveTab("timesheet"));
    dailyDateFilter.addEventListener("change", renderDailyTable);
    dailyStoreFilter.addEventListener("change", renderDailyTable);
    clearDailyFiltersButton.addEventListener("click", () => {
      dailyDateFilter.value = "";
      dailyStoreFilter.value = "";
      renderDailyTable();
    });
    summaryMonthFilter.addEventListener("change", loadSummary);
    timeSheetStartFilter.addEventListener("change", loadTimeSheet);
    timeSheetEndFilter.addEventListener("change", loadTimeSheet);
    clearTimeSheetFiltersButton.addEventListener("click", () => {
      timeSheetStartFilter.value = currentMonthStartValue();
      timeSheetEndFilter.value = currentMonthEndValue();
      loadTimeSheet();
    });

    setAutomaticDate();
    summaryMonthFilter.value = currentMonthValue();
    timeSheetStartFilter.value = currentMonthStartValue();
    timeSheetEndFilter.value = currentMonthEndValue();
    setPunchState(false);
    Promise.all([loadStores(), loadDailyRows()])
      .then(restoreActiveVisit);
    refreshPreview();
  </script>
</body>
</html>`);
});

app.listen(PORT, () => {
  console.log(`PG Daily Itinerary Store Visits is running on port ${PORT}`);
});
