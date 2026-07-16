const WEBAPP_CONFIG = Object.freeze({
  spreadsheetId: '1XE-ZsuxcRExU1Z9jymjbwt34abTSeQn2NvbKA74u9wg',
  historySheet: 'Historie',
  ordersSheet: 'Moje jidlo',
  timezone: 'Europe/Prague',
  maxItemsPerOrder: 20
});

const ORDER_HEADERS = Object.freeze([
  'Datum a čas výběru',
  'Datum menu',
  'Sekce',
  'Položka',
  'Alergeny',
  'Cena (Kč)',
  'Klíč položky',
  'ID výběru'
]);

/**
 * Spusťte jednou před nasazením webové aplikace.
 */
function setup() {
  const spreadsheet = SpreadsheetApp.openById(WEBAPP_CONFIG.spreadsheetId);
  spreadsheet.setSpreadsheetTimeZone(WEBAPP_CONFIG.timezone);
  ensureOrdersSheet_(spreadsheet);
}

/**
 * Veřejné JSONP API pro načtení menu a ověření zápisu.
 */
function doGet(event) {
  const parameters = event && event.parameter ? event.parameter : {};
  const action = parameters.action || 'menu';

  try {
    if (action === 'menu') {
      return jsonp_(parameters.callback, getTodayMenu_());
    }

    if (action === 'status') {
      return jsonp_(parameters.callback, getOrderStatus_(parameters.requestId));
    }

    return jsonp_(parameters.callback, {
      ok: false,
      error: 'Neznámá akce.'
    });
  } catch (error) {
    return jsonp_(parameters.callback, {
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

/**
 * Zápis probíhá přes text/plain POST, aby prohlížeč nemusel provádět CORS preflight.
 * Frontend následně ověří uložení přes JSONP endpoint action=status.
 */
function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents || '{}');
    if (payload.action !== 'save') throw new Error('Neznámá akce.');

    const result = saveOrder_(payload);
    return json_(result);
  } catch (error) {
    return json_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function getTodayMenu_() {
  const spreadsheet = SpreadsheetApp.openById(WEBAPP_CONFIG.spreadsheetId);
  const history = spreadsheet.getSheetByName(WEBAPP_CONFIG.historySheet);
  if (!history || history.getLastRow() < 2) {
    return {ok: true, date: todayKey_(), items: []};
  }

  const values = history.getRange(2, 1, history.getLastRow() - 1, 9).getValues();
  const today = todayKey_();
  const items = values
    .map(historyRowToItem_)
    .filter(item => item.date === today && item.name && item.price !== '')
    .sort((a, b) => a.order - b.order);

  return {
    ok: true,
    date: today,
    updatedAt: new Date().toISOString(),
    items
  };
}

function saveOrder_(payload) {
  const requestId = String(payload.requestId || '');
  const selectedIds = Array.isArray(payload.selectedIds) ? payload.selectedIds.map(String) : [];

  if (!/^[a-zA-Z0-9_-]{12,80}$/.test(requestId)) {
    throw new Error('Neplatné ID výběru.');
  }
  if (!selectedIds.length) throw new Error('Nebylo vybráno žádné jídlo.');
  if (selectedIds.length > WEBAPP_CONFIG.maxItemsPerOrder) {
    throw new Error('Najednou lze uložit nejvýše ' + WEBAPP_CONFIG.maxItemsPerOrder + ' položek.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const spreadsheet = SpreadsheetApp.openById(WEBAPP_CONFIG.spreadsheetId);
    const orders = ensureOrdersSheet_(spreadsheet);

    if (findRequest_(orders, requestId)) {
      return {ok: true, requestId, duplicate: true};
    }

    const currentMenu = getTodayMenu_().items;
    const menuById = new Map(currentMenu.map(item => [item.id, item]));
    const uniqueIds = [...new Set(selectedIds)];
    const selectedItems = uniqueIds.map(id => menuById.get(id)).filter(Boolean);

    if (selectedItems.length !== uniqueIds.length) {
      throw new Error('Některé vybrané položky již nejsou v dnešním menu. Obnovte nabídku.');
    }

    const selectedAt = new Date();
    const rows = selectedItems.map(item => [
      selectedAt,
      parseDateKey_(item.date),
      item.section,
      item.name,
      item.allergens,
      item.price,
      item.id,
      requestId
    ]);

    orders.getRange(orders.getLastRow() + 1, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
    orders.getRange(2, 1, orders.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    orders.getRange(2, 2, orders.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd');

    return {ok: true, requestId, savedCount: rows.length};
  } finally {
    lock.releaseLock();
  }
}

function getOrderStatus_(requestId) {
  const safeRequestId = String(requestId || '');
  if (!/^[a-zA-Z0-9_-]{12,80}$/.test(safeRequestId)) {
    return {ok: false, saved: false, error: 'Neplatné ID výběru.'};
  }

  const spreadsheet = SpreadsheetApp.openById(WEBAPP_CONFIG.spreadsheetId);
  const orders = ensureOrdersSheet_(spreadsheet);
  return {
    ok: true,
    requestId: safeRequestId,
    saved: findRequest_(orders, safeRequestId)
  };
}

function findRequest_(sheet, requestId) {
  if (sheet.getLastRow() < 2) return false;
  return Boolean(
    sheet
      .getRange(2, 8, sheet.getLastRow() - 1, 1)
      .createTextFinder(requestId)
      .matchEntireCell(true)
      .findNext()
  );
}

function ensureOrdersSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(WEBAPP_CONFIG.ordersSheet);
  if (!sheet) sheet = spreadsheet.insertSheet(WEBAPP_CONFIG.ordersSheet);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ORDER_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#e8eaed');
    sheet.autoResizeColumns(1, ORDER_HEADERS.length);
  }

  return sheet;
}

function historyRowToItem_(row) {
  const date = dateKey_(row[0]);
  return {
    id: String(row[8] || [date, row[1], row[2], row[4]].join('|').toLowerCase()),
    date,
    section: String(row[1] || 'Ostatní'),
    name: String(row[2] || ''),
    allergens: String(row[3] || ''),
    price: row[4] === '' ? '' : Number(row[4]),
    order: Number(row[5]) || 0
  };
}

function todayKey_() {
  return Utilities.formatDate(new Date(), WEBAPP_CONFIG.timezone, 'yyyy-MM-dd');
}

function dateKey_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, WEBAPP_CONFIG.timezone, 'yyyy-MM-dd');
  }

  const text = String(value || '').trim();
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];

  const local = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (local) {
    return local[3] + '-' + local[2].padStart(2, '0') + '-' + local[1].padStart(2, '0');
  }

  return text;
}

function parseDateKey_(key) {
  const parts = key.split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, payload) {
  const safeCallback = /^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(callback || '')
    ? callback
    : 'console.log';
  return ContentService
    .createTextOutput(safeCallback + '(' + JSON.stringify(payload) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
