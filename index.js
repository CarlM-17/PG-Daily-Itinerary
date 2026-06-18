const express = require("express");
const { google } = require("googleapis");

const app = express();
const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID =
  process.env.SPREADSHEET_ID || "13hgFrJPHbzrfXb2-tMHJ5lZbtQhSVvyJwRwYaTi9aII";
const DAILY_SHEET_NAME = process.env.DAILY_SHEET_NAME || "Daily Sheet";
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

  return { date, store, timeIn, timeOut, duration, note: note || "" };
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
      cursor: wait;
      opacity: 0.7;
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

    @media (max-width: 760px) {
      main {
        width: min(100% - 20px, 960px);
        padding: 20px 0;
      }

      .topbar,
      .actions {
        align-items: stretch;
        flex-direction: column;
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
            <input id="timeIn" name="timeIn" type="time" required>
          </label>

          <label>
            Time Out
            <input id="timeOut" name="timeOut" type="time" required>
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

        <div class="actions">
          <div class="status" id="status" role="status" aria-live="polite"></div>
          <button id="saveButton" type="submit">Save Visit</button>
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
  </main>

  <script>
    const form = document.getElementById("visitForm");
    const dateInput = document.getElementById("date");
    const storeInput = document.getElementById("store");
    const timeInInput = document.getElementById("timeIn");
    const timeOutInput = document.getElementById("timeOut");
    const noteInput = document.getElementById("note");
    const durationOutput = document.getElementById("duration");
    const statusOutput = document.getElementById("status");
    const saveButton = document.getElementById("saveButton");

    function localDateValue() {
      const now = new Date();
      const offset = now.getTimezoneOffset();
      return new Date(now.getTime() - offset * 60000).toISOString().slice(0, 10);
    }

    function localTimeValue() {
      const now = new Date();
      return String(now.getHours()).padStart(2, "0") + ":" + String(now.getMinutes()).padStart(2, "0");
    }

    function setAutomaticDateAndTime() {
      const currentTime = localTimeValue();
      dateInput.value = localDateValue();
      timeInInput.value = currentTime;
      timeOutInput.value = currentTime;
    }

    function setStatus(message, type = "") {
      statusOutput.textContent = message;
      statusOutput.className = "status" + (type ? " " + type : "");
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
      const duration = durationBetween(timeInInput.value, timeOutInput.value);
      durationOutput.textContent = duration;
      document.getElementById("previewDate").textContent = dateInput.value || "-";
      document.getElementById("previewStore").textContent = storeInput.value || "-";
      document.getElementById("previewTime").textContent =
        timeInInput.value && timeOutInput.value ? timeInInput.value + " - " + timeOutInput.value : "-";
      document.getElementById("previewDuration").textContent = duration;
      document.getElementById("previewNote").textContent = noteInput.value || "-";
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

        for (const store of data.stores) {
          const option = document.createElement("option");
          option.value = store;
          option.textContent = store;
          storeInput.appendChild(option);
        }
      } catch (error) {
        storeInput.innerHTML = '<option value="">Store list unavailable</option>';
        setStatus(error.message, "error");
      }
    }

    form.addEventListener("input", refreshPreview);
    form.addEventListener("change", refreshPreview);

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      saveButton.disabled = true;
      setStatus("Saving...");

      try {
        const response = await fetch("/api/visits", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            date: dateInput.value,
            store: storeInput.value,
            timeIn: timeInInput.value,
            timeOut: timeOutInput.value,
            note: noteInput.value
          })
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.message || "Unable to save visit.");

        setStatus("Saved to Daily Sheet.", "success");
        setAutomaticDateAndTime();
        noteInput.value = "";
        refreshPreview();
      } catch (error) {
        setStatus(error.message, "error");
      } finally {
        saveButton.disabled = false;
      }
    });

    setAutomaticDateAndTime();
    loadStores();
    refreshPreview();
  </script>
</body>
</html>`);
});

app.listen(PORT, () => {
  console.log(`PG Daily Itinerary Store Visits is running on port ${PORT}`);
});
