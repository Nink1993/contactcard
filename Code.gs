const CONFIG = {
  SPREADSHEET_ID: '1dBolcUEme_3AXNnPMxWj-QtXtbt1Yc05vbmfWMTElCs',
  SHEET_NAME: 'Contact ช่างซ่อม',
  REQUIRED_HEADERS: [
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
  ],
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Contact List');
}

function getAllContacts(options) {
  return listContacts_(options);
}

function saveContact(contact) {
  if (!contact || typeof contact !== 'object') throw new Error('Invalid contact payload');
  if (contact.id) return updateContact_(contact.id, contact);
  return createContact_(contact);
}

function deleteContact(id) {
  return deleteContact_(id);
}

function installTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const already = triggers.some((t) => t.getHandlerFunction() === 'onEdit');
  if (!already) {
    ScriptApp.newTrigger('onEdit').forSpreadsheet(CONFIG.SPREADSHEET_ID).onEdit().create();
  }
  return { ok: true };
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    if (range.getRow() === 1) return;

    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    try {
      const header = ensureHeaders_(sheet);
      const row = range.getRow();
      const rowValues = sheet.getRange(row, 1, 1, header.lastColumn).getValues()[0];
      const existingId = valueAt_(rowValues, header.map, 'id');
      const now = new Date();

      const updates = {};
      if (!existingId) updates.id = Utilities.getUuid();
      updates.updatedAt = now.toISOString();
      if (!valueAt_(rowValues, header.map, 'createdAt')) updates.createdAt = now.toISOString();

      const hasContent = Object.keys(header.map).some((k) => {
        if (k === 'id' || k === 'createdat' || k === 'updatedat' || k === 'isdeleted' || k === 'deletedat') return false;
        const v = valueAt_(rowValues, header.map, k);
        return String(v || '').trim() !== '';
      });
      if (!hasContent) return;

      applyFieldUpdates_(sheet, row, header, updates);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err && err.stack ? err.stack : String(err));
  }
}

function listContacts_(options) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const header = ensureHeaders_(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(2, 1, lastRow - 1, header.lastColumn).getValues();
    const includeDeleted = Boolean(options && options.includeDeleted);

    const idCol = header.map[normalizeHeader_('id')];
    const createdAtCol = header.map[normalizeHeader_('createdAt')];
    const updatedAtCol = header.map[normalizeHeader_('updatedAt')];
    const nowIso = new Date().toISOString();

    const idUpdates = [];
    values.forEach((row, i) => {
      const rowNumber = i + 2;
      const existingId = idCol ? String(row[idCol - 1] || '').trim() : '';
      if (existingId) return;
      const hasContent = Object.keys(header.map).some((k) => {
        if (k === 'id' || k === 'createdat' || k === 'updatedat' || k === 'isdeleted' || k === 'deletedat') return false;
        const v = valueAt_(row, header.map, k);
        return String(v || '').trim() !== '';
      });
      if (!hasContent) return;
      const newId = Utilities.getUuid();
      if (idCol) idUpdates.push({ rowNumber, col: idCol, value: newId });
      if (createdAtCol) idUpdates.push({ rowNumber, col: createdAtCol, value: nowIso });
      if (updatedAtCol) idUpdates.push({ rowNumber, col: updatedAtCol, value: nowIso });
      row[idCol - 1] = newId;
      if (createdAtCol) row[createdAtCol - 1] = nowIso;
      if (updatedAtCol) row[updatedAtCol - 1] = nowIso;
    });

    if (idUpdates.length) {
      idUpdates.forEach((u) => sheet.getRange(u.rowNumber, u.col).setValue(u.value));
    }

    const rows = values
      .map((row, idx) => ({ rowNumber: idx + 2, row }))
      .map((entry) => fromRow_(entry.row, header.map, entry.rowNumber))
      .filter((c) => c && c.id);

    if (includeDeleted) return rows;
    return rows.filter((c) => !c.isDeleted);
  } finally {
    lock.releaseLock();
  }
}

function createContact_(input) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const header = ensureHeaders_(sheet);
    const now = new Date().toISOString();
    const contact = normalizeContact_(input);
    contact.id = Utilities.getUuid();
    contact.createdAt = now;
    contact.updatedAt = now;
    contact.isDeleted = false;
    contact.deletedAt = '';

    const row = toRow_(contact, header);
    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();
    return Object.assign({}, contact, { rowNumber });
  } finally {
    lock.releaseLock();
  }
}

function updateContact_(id, input) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const header = ensureHeaders_(sheet);
    const rowNumber = findRowById_(sheet, header, id);
    if (!rowNumber) throw new Error('Not found');

    const existingRow = sheet.getRange(rowNumber, 1, 1, header.lastColumn).getValues()[0];
    const existing = fromRow_(existingRow, header.map, rowNumber);
    const patch = normalizeContact_(input);

    const nowIso = new Date().toISOString();
    const merged = Object.assign({}, existing, patch, {
      id,
      updatedAt: nowIso,
      createdAt: existing.createdAt || nowIso,
    });

    const row = toRow_(merged, header);
    sheet.getRange(rowNumber, 1, 1, header.lastColumn).setValues([row]);
    return merged;
  } finally {
    lock.releaseLock();
  }
}

function deleteContact_(id) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const header = ensureHeaders_(sheet);
    const rowNumber = findRowById_(sheet, header, id);
    if (!rowNumber) throw new Error('Not found');

    const nowIso = new Date().toISOString();
    const updates = { isDeleted: true, deletedAt: nowIso, updatedAt: nowIso };
    applyFieldUpdates_(sheet, rowNumber, header, updates);

    const afterRow = sheet.getRange(rowNumber, 1, 1, header.lastColumn).getValues()[0];
    return fromRow_(afterRow, header.map, rowNumber);
  } finally {
    lock.releaseLock();
  }
}

function ensureHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  headerRow = headerRow.map((v) => String(v || '').trim());
  const isEmptyHeader = headerRow.every((h) => !h);
  if (isEmptyHeader) {
    sheet.getRange(1, 1, 1, lastColumn).clearContent();
    sheet.getRange(1, 1, 1, CONFIG.REQUIRED_HEADERS.length).setValues([CONFIG.REQUIRED_HEADERS]);
    const map0 = buildHeaderMap_(CONFIG.REQUIRED_HEADERS);
    return { headerRow: CONFIG.REQUIRED_HEADERS.slice(), map: map0, lastColumn: CONFIG.REQUIRED_HEADERS.length };
  }

  const existingNorm = headerRow.map(normalizeHeader_);
  const required = CONFIG.REQUIRED_HEADERS.slice();
  const toAppend = [];

  required.forEach((h) => {
    const hn = normalizeHeader_(h);
    if (!existingNorm.includes(hn)) toAppend.push(h);
  });

  if (toAppend.length) {
    const writeCol = headerRow.length + 1;
    sheet.getRange(1, writeCol, 1, toAppend.length).setValues([toAppend]);
    headerRow = headerRow.concat(toAppend);
  }

  const map = buildHeaderMap_(headerRow);
  return { headerRow, map, lastColumn: headerRow.length };
}

function buildHeaderMap_(headerRow) {
  const map = {};
  headerRow.forEach((h, idx) => {
    const hn = normalizeHeader_(h);
    if (!hn) return;
    if (map[hn] == null) map[hn] = idx + 1;
  });
  return map;
}

function normalizeHeader_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[._-]/g, ' ');
}

function normalizeContact_(input) {
  const obj = input && typeof input === 'object' ? input : {};
  const out = {};
  out.company = pick_(obj, ['company', 'Company', 'บริษัท']);
  out.area = pick_(obj, ['area', 'Area']);
  out.scope = pick_(obj, ['scope', 'Scope']);
  out.name1 = pick_(obj, ['name1', 'name 1', 'Name 1']);
  out.phone1 = pick_(obj, ['phone1', 'phone 1', 'Phone 1']);
  out.name2 = pick_(obj, ['name2', 'name 2', 'Name 2']);
  out.phone2 = pick_(obj, ['phone2', 'phone 2', 'Phone 2']);
  out.bill = pick_(obj, ['bill', 'บิล']);
  out.account_name = pick_(obj, ['account_name', 'account name', 'ชื่อบัญชี']);
  out.bank = pick_(obj, ['bank', 'ธนาคาร']);
  out.account_no = pick_(obj, ['account_no', 'account no', 'เลขบัญชี']);
  out.notes = pick_(obj, ['notes', 'Notes']);
  if (obj.isDeleted != null) out.isDeleted = Boolean(obj.isDeleted);
  return pruneEmpty_(out);
}

function pruneEmpty_(obj) {
  const out = {};
  Object.keys(obj).forEach((k) => {
    const v = obj[k];
    if (v == null) return;
    if (typeof v === 'string' && v.trim() === '') return;
    out[k] = v;
  });
  return out;
}

function pick_(obj, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const k = keys[i];
    if (Object.prototype.hasOwnProperty.call(obj, k)) return obj[k];
  }
  return '';
}

function toRow_(contact, header) {
  const row = new Array(header.lastColumn).fill('');

  setByHeader_(row, header.map, 'id', contact.id || '');
  setByHeader_(row, header.map, 'company', contact.company || '', ['company', 'Company']);
  setByHeader_(row, header.map, 'area', contact.area || '', ['area', 'Area']);
  setByHeader_(row, header.map, 'scope', contact.scope || '', ['scope', 'Scope']);
  setByHeader_(row, header.map, 'name 1', contact.name1 || '', ['name 1', 'Name 1']);
  setByHeader_(row, header.map, 'phone 1', contact.phone1 || '', ['phone 1', 'Phone 1']);
  setByHeader_(row, header.map, 'name 2', contact.name2 || '', ['name 2', 'Name 2']);
  setByHeader_(row, header.map, 'phone 2', contact.phone2 || '', ['phone 2', 'Phone 2']);
  setByHeader_(row, header.map, 'บิล', contact.bill || '', ['บิล', 'bill']);
  setByHeader_(row, header.map, 'ชื่อบัญชี', contact.account_name || '', ['ชื่อบัญชี', 'account name', 'account_name']);
  setByHeader_(row, header.map, 'ธนาคาร', contact.bank || '', ['ธนาคาร', 'bank']);
  setByHeader_(row, header.map, 'เลขบัญชี', contact.account_no || '', ['เลขบัญชี', 'account no', 'account_no']);
  setByHeader_(row, header.map, 'notes', contact.notes || '');
  setByHeader_(row, header.map, 'createdat', contact.createdAt || '');
  setByHeader_(row, header.map, 'updatedat', contact.updatedAt || '');
  setByHeader_(row, header.map, 'isdeleted', contact.isDeleted ? 'TRUE' : 'FALSE');
  setByHeader_(row, header.map, 'deletedat', contact.deletedAt || '');

  return row;
}

function setByHeader_(row, map, headerKey, value, aliases) {
  const keys = [headerKey].concat(aliases || []);
  for (let i = 0; i < keys.length; i += 1) {
    const k = normalizeHeader_(keys[i]);
    const col = map[k];
    if (col) {
      row[col - 1] = value;
      return;
    }
  }
}

function valueAt_(row, map, headerKey) {
  const col = map[normalizeHeader_(headerKey)];
  if (!col) return '';
  return row[col - 1];
}

function fromRow_(row, map, rowNumber) {
  const id = String(valueAt_(row, map, 'id') || '').trim();
  if (!id) return null;

  const company = String(valueAt_(row, map, 'Company') || valueAt_(row, map, 'company') || '').trim();
  const area = String(valueAt_(row, map, 'Area') || valueAt_(row, map, 'area') || '').trim();
  const scope = String(valueAt_(row, map, 'Scope') || valueAt_(row, map, 'scope') || '').trim();
  const name1 = String(valueAt_(row, map, 'Name 1') || valueAt_(row, map, 'name 1') || '').trim();
  const phone1 = String(valueAt_(row, map, 'Phone 1') || valueAt_(row, map, 'phone 1') || '').trim();
  const name2 = String(valueAt_(row, map, 'Name 2') || valueAt_(row, map, 'name 2') || '').trim();
  const phone2 = String(valueAt_(row, map, 'Phone 2') || valueAt_(row, map, 'phone 2') || '').trim();
  const bill = String(valueAt_(row, map, 'บิล') || valueAt_(row, map, 'bill') || '').trim();
  const account_name = String(valueAt_(row, map, 'ชื่อบัญชี') || valueAt_(row, map, 'account name') || valueAt_(row, map, 'account_name') || '').trim();
  const bank = String(valueAt_(row, map, 'ธนาคาร') || valueAt_(row, map, 'bank') || '').trim();
  const account_no = String(valueAt_(row, map, 'เลขบัญชี') || valueAt_(row, map, 'account no') || valueAt_(row, map, 'account_no') || '').trim();
  const notes = String(valueAt_(row, map, 'notes') || '').trim();

  const createdAt = String(valueAt_(row, map, 'createdAt') || valueAt_(row, map, 'createdat') || '').trim();
  const updatedAt = String(valueAt_(row, map, 'updatedAt') || valueAt_(row, map, 'updatedat') || '').trim();
  const deletedAt = String(valueAt_(row, map, 'deletedAt') || valueAt_(row, map, 'deletedat') || '').trim();

  const isDeletedRaw = valueAt_(row, map, 'isDeleted') || valueAt_(row, map, 'isdeleted');
  const isDeleted = String(isDeletedRaw || '').toLowerCase() === 'true' || String(isDeletedRaw || '').toLowerCase() === 'yes' || String(isDeletedRaw || '') === '1';

  return {
    id,
    company,
    area,
    scope,
    name1,
    phone1,
    name2,
    phone2,
    bill,
    account_name,
    bank,
    account_no,
    notes,
    createdAt,
    updatedAt,
    isDeleted,
    deletedAt,
    rowNumber,
  };
}

function findRowById_(sheet, header, id) {
  const idCol = header.map[normalizeHeader_('id')];
  if (!idCol) throw new Error('Missing id column');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const tf = sheet.getRange(2, idCol, lastRow - 1, 1).createTextFinder(id).matchEntireCell(true);
  const found = tf.findNext();
  return found ? found.getRow() : null;
}

function applyFieldUpdates_(sheet, rowNumber, header, updates) {
  const keys = Object.keys(updates || {});
  if (!keys.length) return;

  keys.forEach((k) => {
    const col = header.map[normalizeHeader_(k)];
    if (!col) return;
    sheet.getRange(rowNumber, col).setValue(updates[k]);
  });
}
