const { google } = require('googleapis');

const DEFAULT_SHEET_ID = '1dBolcUEme_3AXNnPMxWj-QtXtbt1Yc05vbmfWMTElCs';
const DEFAULT_SHEET_NAME = 'Contact ช่างซ่อม';

const REQUIRED_HEADERS = [
  'id',
  'Company',
  'Area',
  'Scope',
  'Name 1',
  'Phone 1',
  'Name 2',
  'Phone 2',
  'บิล',
  'ชื่อบัญชี',
  'ธนาคาร',
  'เลขบัญชี',
  'notes',
  'createdAt',
  'updatedAt',
  'isDeleted',
  'deletedAt',
];

function normalizeHeader(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[._-]/g, ' ');
}

function nowIso() {
  return new Date().toISOString();
}

function readBody(req) {
  const contentType = String(req.headers['content-type'] || '').toLowerCase();
  if (!req.body) return {};
  if (typeof req.body === 'object') return req.body;
  if (typeof req.body === 'string') {
    if (contentType.includes('application/json')) {
      try {
        return JSON.parse(req.body);
      } catch (e) {
        return {};
      }
    }
  }
  return {};
}

function json(res, status, data) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(data));
}

function getEnv(name, fallback) {
  const v = process.env[name];
  if (v == null || String(v).trim() === '') return fallback;
  return v;
}

function assertConfig() {
  const clientEmail = getEnv('GOOGLE_SERVICE_ACCOUNT_EMAIL', '');
  const privateKey = getEnv('GOOGLE_PRIVATE_KEY', '');
  if (!clientEmail || !privateKey) {
    const err = new Error('Missing GOOGLE_SERVICE_ACCOUNT_EMAIL or GOOGLE_PRIVATE_KEY env vars');
    err.statusCode = 500;
    throw err;
  }
}

function getSheetConfig() {
  return {
    spreadsheetId: getEnv('SHEET_ID', DEFAULT_SHEET_ID),
    sheetName: getEnv('SHEET_NAME', DEFAULT_SHEET_NAME),
  };
}

function getAuth() {
  const clientEmail = getEnv('GOOGLE_SERVICE_ACCOUNT_EMAIL', '');
  const privateKeyRaw = getEnv('GOOGLE_PRIVATE_KEY', '');
  const privateKey = privateKeyRaw.replace(/\\n/g, '\n');

  return new google.auth.JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getSheetsClient() {
  const auth = getAuth();
  await auth.authorize();
  return google.sheets({ version: 'v4', auth });
}

function buildHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((h, idx) => {
    const n = normalizeHeader(h);
    if (!n) return;
    if (map[n] == null) map[n] = idx;
  });
  return map;
}

function a1(sheetName, col1, row1, col2, row2) {
  const colToLetter = (col) => {
    let temp = col + 1;
    let letter = '';
    while (temp > 0) {
      const mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - 1) / 26);
    }
    return letter;
  };
  const start = colToLetter(col1) + row1;
  const end = colToLetter(col2) + row2;
  return sheetName + '!' + start + ':' + end;
}

async function ensureHeaders(sheets, spreadsheetId, sheetName) {
  const range = sheetName + '!1:1';
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  const headerRow = (resp.data.values && resp.data.values[0]) ? resp.data.values[0] : [];
  const normalized = headerRow.map(normalizeHeader);

  if (headerRow.length === 0 || headerRow.every((v) => !String(v || '').trim())) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: sheetName + '!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [REQUIRED_HEADERS] },
    });
    return { headerRow: REQUIRED_HEADERS.slice(), map: buildHeaderMap(REQUIRED_HEADERS) };
  }

  const toAppend = [];
  for (const h of REQUIRED_HEADERS) {
    const hn = normalizeHeader(h);
    if (!normalized.includes(hn)) toAppend.push(h);
  }

  if (toAppend.length) {
    const startColIdx = headerRow.length;
    const endColIdx = startColIdx + toAppend.length - 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: a1(sheetName, startColIdx, 1, endColIdx, 1),
      valueInputOption: 'RAW',
      requestBody: { values: [toAppend] },
    });
    const merged = headerRow.concat(toAppend);
    return { headerRow: merged, map: buildHeaderMap(merged) };
  }

  return { headerRow, map: buildHeaderMap(headerRow) };
}

function getCell(row, map, headerKey) {
  const idx = map[normalizeHeader(headerKey)];
  if (idx == null) return '';
  return row[idx] == null ? '' : row[idx];
}

function setCell(row, map, headerKey, value) {
  const idx = map[normalizeHeader(headerKey)];
  if (idx == null) return;
  row[idx] = value;
}

function toContact(row, map, rowNumber) {
  const id = String(getCell(row, map, 'id') || '').trim();
  if (!id) return null;

  const isDeletedRaw = getCell(row, map, 'isDeleted');
  const isDeleted = String(isDeletedRaw || '').toLowerCase() === 'true' || String(isDeletedRaw || '').toLowerCase() === 'yes' || String(isDeletedRaw || '') === '1';

  return {
    id,
    company: String(getCell(row, map, 'Company') || getCell(row, map, 'company') || '').trim(),
    area: String(getCell(row, map, 'Area') || getCell(row, map, 'area') || '').trim(),
    scope: String(getCell(row, map, 'Scope') || getCell(row, map, 'scope') || '').trim(),
    name1: String(getCell(row, map, 'Name 1') || getCell(row, map, 'name 1') || '').trim(),
    phone1: String(getCell(row, map, 'Phone 1') || getCell(row, map, 'phone 1') || '').trim(),
    name2: String(getCell(row, map, 'Name 2') || getCell(row, map, 'name 2') || '').trim(),
    phone2: String(getCell(row, map, 'Phone 2') || getCell(row, map, 'phone 2') || '').trim(),
    bill: String(getCell(row, map, 'บิล') || getCell(row, map, 'bill') || '').trim(),
    account_name: String(getCell(row, map, 'ชื่อบัญชี') || getCell(row, map, 'account_name') || '').trim(),
    bank: String(getCell(row, map, 'ธนาคาร') || getCell(row, map, 'bank') || '').trim(),
    account_no: String(getCell(row, map, 'เลขบัญชี') || getCell(row, map, 'account_no') || '').trim(),
    notes: String(getCell(row, map, 'notes') || '').trim(),
    createdAt: String(getCell(row, map, 'createdAt') || '').trim(),
    updatedAt: String(getCell(row, map, 'updatedAt') || '').trim(),
    isDeleted,
    deletedAt: String(getCell(row, map, 'deletedAt') || '').trim(),
    rowNumber,
  };
}

function normalizeInput(obj) {
  const input = obj && typeof obj === 'object' ? obj : {};
  const pick = (keys) => {
    for (const k of keys) {
      if (Object.prototype.hasOwnProperty.call(input, k)) return input[k];
    }
    return '';
  };

  const out = {
    company: pick(['company', 'Company']),
    area: pick(['area', 'Area']),
    scope: pick(['scope', 'Scope']),
    name1: pick(['name1', 'name 1', 'Name 1']),
    phone1: pick(['phone1', 'phone 1', 'Phone 1']),
    name2: pick(['name2', 'name 2', 'Name 2']),
    phone2: pick(['phone2', 'phone 2', 'Phone 2']),
    bill: pick(['bill', 'บิล']),
    account_name: pick(['account_name', 'ชื่อบัญชี']),
    bank: pick(['bank', 'ธนาคาร']),
    account_no: pick(['account_no', 'เลขบัญชี']),
    notes: pick(['notes']),
  };

  const cleaned = {};
  for (const [k, v] of Object.entries(out)) {
    if (v == null) continue;
    if (typeof v === 'string' && v.trim() === '') continue;
    cleaned[k] = v;
  }
  return cleaned;
}

function rowFromContact(contact, headerRow, map) {
  const row = new Array(headerRow.length).fill('');
  setCell(row, map, 'id', contact.id || '');
  setCell(row, map, 'Company', contact.company || '');
  setCell(row, map, 'Area', contact.area || '');
  setCell(row, map, 'Scope', contact.scope || '');
  setCell(row, map, 'Name 1', contact.name1 || '');
  setCell(row, map, 'Phone 1', contact.phone1 || '');
  setCell(row, map, 'Name 2', contact.name2 || '');
  setCell(row, map, 'Phone 2', contact.phone2 || '');
  setCell(row, map, 'บิล', contact.bill || '');
  setCell(row, map, 'ชื่อบัญชี', contact.account_name || '');
  setCell(row, map, 'ธนาคาร', contact.bank || '');
  setCell(row, map, 'เลขบัญชี', contact.account_no || '');
  setCell(row, map, 'notes', contact.notes || '');
  setCell(row, map, 'createdAt', contact.createdAt || '');
  setCell(row, map, 'updatedAt', contact.updatedAt || '');
  setCell(row, map, 'isDeleted', contact.isDeleted ? 'TRUE' : 'FALSE');
  setCell(row, map, 'deletedAt', contact.deletedAt || '');
  return row;
}

async function readAllRows(sheets, spreadsheetId, sheetName) {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: sheetName,
  });
  return resp.data.values || [];
}

async function findRowById(sheets, spreadsheetId, sheetName, header, id) {
  const rows = await readAllRows(sheets, spreadsheetId, sheetName);
  if (rows.length < 2) return { rows, rowNumber: null };

  const idIdx = header.map[normalizeHeader('id')];
  if (idIdx == null) return { rows, rowNumber: null };

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const cell = row[idIdx] == null ? '' : String(row[idIdx]).trim();
    if (cell === id) return { rows, rowNumber: i + 1 };
  }
  return { rows, rowNumber: null };
}

async function backfillIdsIfNeeded(sheets, spreadsheetId, sheetName, header, rows) {
  if (!rows || rows.length < 2) return;
  const idIdx = header.map[normalizeHeader('id')];
  if (idIdx == null) return;

  const createdIdx = header.map[normalizeHeader('createdAt')];
  const updatedIdx = header.map[normalizeHeader('updatedAt')];
  const now = nowIso();

  const updates = [];
  for (let i = 1; i < rows.length; i += 1) {
    const rowNumber = i + 1;
    const row = rows[i] || [];
    const hasContent = row.some((v) => String(v || '').trim() !== '');
    const existingId = row[idIdx] == null ? '' : String(row[idIdx]).trim();
    if (!hasContent || existingId) continue;

    const newId = require('crypto').randomUUID();
    updates.push({ rowNumber, colIdx: idIdx, value: newId });
    if (createdIdx != null) updates.push({ rowNumber, colIdx: createdIdx, value: now });
    if (updatedIdx != null) updates.push({ rowNumber, colIdx: updatedIdx, value: now });
  }

  for (const u of updates) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: a1(sheetName, u.colIdx, u.rowNumber, u.colIdx, u.rowNumber),
      valueInputOption: 'RAW',
      requestBody: { values: [[u.value]] },
    });
  }
}

module.exports = async (req, res) => {
  try {
    assertConfig();
    const { spreadsheetId, sheetName } = getSheetConfig();
    const sheets = await getSheetsClient();

    const header = await ensureHeaders(sheets, spreadsheetId, sheetName);

    if (req.method === 'GET') {
      const includeDeleted = String(req.query && req.query.includeDeleted || 'false').toLowerCase() === 'true';
      const rows = await readAllRows(sheets, spreadsheetId, sheetName);
      await backfillIdsIfNeeded(sheets, spreadsheetId, sheetName, header, rows);
      const contacts = (rows.slice(1) || [])
        .map((r, i) => toContact(r, header.map, i + 2))
        .filter(Boolean);
      const data = includeDeleted ? contacts : contacts.filter((c) => !c.isDeleted);
      return json(res, 200, { ok: true, data });
    }

    if (req.method === 'POST') {
      const payload = normalizeInput(readBody(req));
      const id = payload.id ? String(payload.id).trim() : '';
      const rows = await readAllRows(sheets, spreadsheetId, sheetName);
      await backfillIdsIfNeeded(sheets, spreadsheetId, sheetName, header, rows);

      if (!id) {
        const newId = require('crypto').randomUUID();
        const now = nowIso();
        const contact = Object.assign(
          {
            id: newId,
            createdAt: now,
            updatedAt: now,
            isDeleted: false,
            deletedAt: '',
          },
          payload
        );

        const rowValues = rowFromContact(contact, header.headerRow, header.map);
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: sheetName,
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS',
          requestBody: { values: [rowValues] },
        });
        return json(res, 200, { ok: true, data: contact });
      }

      const found = await findRowById(sheets, spreadsheetId, sheetName, header, id);
      if (!found.rowNumber) return json(res, 404, { ok: false, error: 'Not found' });

      const existingRow = found.rows[found.rowNumber - 1] || [];
      const existing = toContact(existingRow, header.map, found.rowNumber);
      if (!existing) return json(res, 404, { ok: false, error: 'Not found' });

      const merged = Object.assign({}, existing, payload, {
        id,
        updatedAt: nowIso(),
        createdAt: existing.createdAt || nowIso(),
      });

      const rowValues = rowFromContact(merged, header.headerRow, header.map);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: a1(sheetName, 0, found.rowNumber, header.headerRow.length - 1, found.rowNumber),
        valueInputOption: 'RAW',
        requestBody: { values: [rowValues] },
      });

      return json(res, 200, { ok: true, data: merged });
    }

    if (req.method === 'DELETE') {
      const id = String((req.query && req.query.id) || '').trim();
      if (!id) return json(res, 400, { ok: false, error: 'Missing id' });

      const found = await findRowById(sheets, spreadsheetId, sheetName, header, id);
      if (!found.rowNumber) return json(res, 404, { ok: false, error: 'Not found' });

      const now = nowIso();
      const updates = [
        { key: 'isDeleted', value: 'TRUE' },
        { key: 'deletedAt', value: now },
        { key: 'updatedAt', value: now },
      ];

      for (const u of updates) {
        const colIdx = header.map[normalizeHeader(u.key)];
        if (colIdx == null) continue;
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: a1(sheetName, colIdx, found.rowNumber, colIdx, found.rowNumber),
          valueInputOption: 'RAW',
          requestBody: { values: [[u.value]] },
        });
      }

      return json(res, 200, { ok: true });
    }

    res.setHeader('Allow', 'GET,POST,DELETE');
    return json(res, 405, { ok: false, error: 'Method not allowed' });
  } catch (err) {
    const status = err && err.statusCode ? err.statusCode : 500;
    return json(res, status, { ok: false, error: err && err.message ? err.message : String(err) });
  }
};

