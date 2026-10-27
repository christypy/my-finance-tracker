const BACKEND_VERSION_ = '2026-10-27-projects-v1';

const ADVANCE_CATEGORY_NAME_ = '代墊';

const SHEET_NAMES = {
  accounts: 'Accounts',
  liabilities: 'Liabilities',
  transactions: 'Transactions',
  recurring: 'RecurringTemplates',
  categories: 'Categories',
  orders: 'Orders',
  projects: 'Projects',
  projectItems: 'ProjectItems'
};

const SHEET_HEADERS = {
  
  accounts: ['id', 'name', 'type', 'balance', 'interestRate', 'note', 'updatedAt', 'favorite', 'interestCap', 'normalRate', 'interestFreqMonths'],
  liabilities: ['id', 'name', 'amount', 'dueDate', 'paid', 'note', 'updatedAt'],
  
  transactions: ['id', 'date', 'type', 'category', 'amount', 'accountId', 'note', 'updatedAt', 'payer', 'splitMode', 'settled', 'recurringId', 'subcategory', 'toAccountId', 'cashFunded', 'projectId'],
  
  recurring: ['id', 'name', 'type', 'category', 'amount', 'accountId', 'dayOfMonth', 'note', 'active', 'updatedAt', 'subcategory', 'payer', 'splitMode', 'toAccountId'],
  
  categories: ['id', 'type', 'mainName', 'subName', 'updatedAt'],
  // 待付款、尚未付款的項目：先計入月預算，但不影響帳戶餘額、也不算進記帳支出。
  // 付款確認後會轉成一筆正式支出（Transactions）並從這張表移除；取消訂單則直接刪除這列。
  orders: ['id', 'name', 'amount', 'category', 'subcategory', 'orderDate', 'note', 'updatedAt', 'payer', 'splitMode', 'recurringId'],
  // 專案（爸爸／媽媽給的專款）：專案進行中完全獨立，不碰 Transactions、不動帳戶餘額，
  // 只有「結案」時才會一次整合進記帳（見 closeProjectTx_），「取消結案」可完整還原。
  projects: ['id', 'name', 'source', 'fundAmount', 'fundAccountId', 'startDate', 'status', 'closedAt', 'leftoverMode', 'leftoverAccountId', 'reimburse', 'result', 'note', 'updatedAt'],
  projectItems: ['id', 'projectId', 'date', 'name', 'amount', 'category', 'payer', 'accountId', 'note', 'updatedAt']
};


const DEFAULT_CATEGORY_TREE_ = {
  expense: [
    { name: '餐飲', subs: ['早餐', '午餐', '晚餐', '飲料', '消夜'] },
    { name: '交通', subs: ['大眾運輸', '加油', '停車', '計程車', '保養'] },
    { name: '購物', subs: ['日用品', '防曬', '保養品', '化妝品', '服飾', '電器', '其他'] },
    { name: '娛樂', subs: ['電影', '遊戲', '訂閱', '旅遊'] },
    { name: '醫療', subs: ['看診', '藥品', '保健品'] },
    { name: '教育', subs: ['書籍', '課程', '學費'] },
    { name: '居家', subs: ['房租', '水電', '瓦斯', '網路', '家具'] },
    { name: '保險', subs: ['保費'] },
    { name: '旅遊', subs: ['機票', '住宿', '餐飲', '交通'] },
    { name: '其他', subs: [] }
  ],
  income: [
    { name: '薪資', subs: ['本薪', '加班費'] },
    { name: '獎金', subs: ['年終獎金', '績效獎金'] },
    { name: '投資收益', subs: ['股息', '利息', '資本利得'] },
    { name: '兼職', subs: [] },
    { name: '其他', subs: [] }
  ]
};
function seedDefaultCategories_(sheet) {
  const headers = SHEET_HEADERS.categories;
  const now = new Date().toISOString();
  const matrix = [];
  ['expense', 'income'].forEach(type => {
    DEFAULT_CATEGORY_TREE_[type].forEach(m => {
      if (!m.subs || !m.subs.length) {
        matrix.push([Utilities.getUuid(), type, m.name, '', now]);
      } else {
        m.subs.forEach(s => matrix.push([Utilities.getUuid(), type, m.name, s, now]));
      }
    });
  });
  if (matrix.length) sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
}



const sheetCache_ = {};
let cachedSpreadsheet_ = null;
function getSpreadsheet_() {
  if (!cachedSpreadsheet_) cachedSpreadsheet_ = SpreadsheetApp.getActiveSpreadsheet();
  return cachedSpreadsheet_;
}
function getSheet_(key) {
  if (sheetCache_[key]) return sheetCache_[key];
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAMES[key]);
  let headerRow;
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES[key]);
    sheet.appendRow(SHEET_HEADERS[key]);
    setDateColumnsAsPlainText_(sheet, key);
    if (key === 'categories') seedDefaultCategories_(sheet);
    headerRow = SHEET_HEADERS[key]; // 新建立的表，表頭就是這份定義，不用再讀一次
  } else {
    headerRow = ensureHeaders_(sheet, SHEET_HEADERS[key]); // 效能：讓 ensureHeaders_ 把它剛剛讀到的表頭陣列回傳出來
  }
  sheetCache_[key] = sheet;

  if (!colMapCache_[key]) {
    const map = {};
    headerRow.forEach((h, i) => {
      const name = String(h || '').trim();
      if (name && map[name] === undefined) map[name] = i + 1; // 1-based 欄號；重複表頭取第一個
    });
    colMapCache_[key] = map;
  }
  return sheet;
}


function setDateColumnsAsPlainText_(sheet, key) {
  const lastCol = sheet.getLastColumn();
  const headerRow = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim()) : [];
  DATE_ONLY_FIELDS_.forEach(col => {
    const idx = headerRow.indexOf(col);
    if (idx !== -1) {
      sheet.getRange(1, idx + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
    }
  });
}


function fixExistingDateColumns() {
  ['transactions', 'liabilities'].forEach(key => {
    const sheet = getSheet_(key);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { setDateColumnsAsPlainText_(sheet, key); return; }
    DATE_ONLY_FIELDS_.forEach(col => {
      const idx = getColMap_(key)[col];
      if (!idx) return;
      const range = sheet.getRange(2, idx, lastRow - 1, 1);
      const values = range.getValues().map(row => [normalizeCellValue_(col, row[0])]);
      range.setNumberFormat('@').setValues(values);
    });
  });
}


function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim()) : [];
  const existing = {};
  currentHeaders.forEach(h => { if (h) existing[h] = true; });
  const missing = headers.filter(h => !existing[h]);
  if (missing.length) {
    sheet.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing]);
  }

  return missing.length ? currentHeaders.concat(missing) : currentHeaders;
}


const colMapCache_ = {};
function getColMap_(key) {
  if (colMapCache_[key]) return colMapCache_[key];

  getSheet_(key);
  return colMapCache_[key] || {};
}

function colIndex_(key, headerName) {
  const idx = getColMap_(key)[headerName];
  if (!idx) throw new Error('試算表「' + SHEET_NAMES[key] + '」找不到欄位: ' + headerName);
  return idx;
}


function getActualHeaderRow_(key) {
  const map = getColMap_(key);
  const arr = [];
  Object.keys(map).forEach(h => { arr[map[h] - 1] = h; });
  return arr;
}


function buildRowArray_(key, data) {
  const colMap = getColMap_(key);
  let width = 0;
  Object.keys(colMap).forEach(h => { if (colMap[h] > width) width = colMap[h]; });
  const row = new Array(width).fill('');
  SHEET_HEADERS[key].forEach(h => {
    const idx = colMap[h];
    if (idx) row[idx - 1] = (data[h] !== undefined ? data[h] : '');
  });
  return row;
}

function checkToken_(token) {
  const real = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN');
  return real && token && real === token;
}


const DATE_ONLY_FIELDS_ = ['date', 'dueDate', 'orderDate', 'startDate', 'closedAt'];
function normalizeCellValue_(header, value) {
  if (value instanceof Date) {
    const tz = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    if (DATE_ONLY_FIELDS_.indexOf(header) !== -1) {
      return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
    }

    return Utilities.formatDate(value, tz, "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  }
  return value;
}

function readAll_(key) {
  const sheet = getSheet_(key);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = normalizeCellValue_(h, r[i])));
      return obj;
    });
}


function findRowIndexById_(key, id) {
  const sheet = getSheet_(key);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const idCol = colIndex_(key, 'id');
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  const target = String(id);
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === target) return i + 2; // 1-based row number
  }
  return -1;
}


function readRowObj_(key, rowIndex) {
  const sheet = getSheet_(key);
  const lastCol = sheet.getLastColumn();
  const headerRow = getActualHeaderRow_(key);
  const values = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const obj = {};
  headerRow.forEach((h, i) => { if (h) obj[h] = normalizeCellValue_(h, values[i]); });
  return obj;
}

function addRow_(key, data) {
  const sheet = getSheet_(key);
  data.id = data.id || Utilities.getUuid(); // 前端樂觀更新時會先產生 id，直接沿用
  data.updatedAt = new Date().toISOString();
  const row = buildRowArray_(key, data); // 依試算表目前實際的欄位順序組列，不假設固定順序
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);

  return data;
}


function updateRow_(key, data, knownRowIndex) {
  const sheet = getSheet_(key);
  const rowIndex = knownRowIndex || findRowIndexById_(key, data.id);
  if (rowIndex === -1) throw new Error('找不到這筆資料: ' + data.id);
  data.updatedAt = new Date().toISOString();
  const row = buildRowArray_(key, data); // 依試算表目前實際的欄位順序組列，不假設固定順序
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  // 效能考量同 addRow_，故意不再同步更新下拉驗證，見上方註解。
  return data;
}

function deleteRow_(key, id) {
  const sheet = getSheet_(key);
  const rowIndex = findRowIndexById_(key, id);
  if (rowIndex === -1) throw new Error('找不到這筆資料: ' + id);
  sheet.deleteRow(rowIndex);
  return { id: id };
}




// 另一半／妹妹先付的支出，錢不是從我的帳戶出的：帳戶一律清空，避免扣到我的餘額。
// 只影響之後新增／編輯的紀錄；
// 編輯舊紀錄時，舊的扣款仍會用原本的帳戶正確還原。
function normalizePayerAccount_(d) {
  if (d && d.type === 'expense' && (d.payer === 'partner' || d.payer === 'sister')) {
    d.accountId = '';
  }
  return d;
}

// 後端最後一道防呆：就算前端漏擋，也不讓「沒日期／金額不是正數／轉帳帳戶有問題」的紀錄寫進試算表，
// 避免帳戶餘額被算錯。只檢查絕對不合理的情況，不會擋正常（含舊資料）的操作。
function validateTransactionData_(d) {
  if (!d) throw new Error('沒有收到記帳資料');
  if (['expense', 'income', 'transfer', 'settlement', 'project'].indexOf(d.type) === -1) throw new Error('記帳類型不正確: ' + d.type);
  if (!String(normalizeCellValue_('date', d.date) || '').trim()) throw new Error('缺少日期，這筆沒有儲存');
  const amt = Number(d.amount);
  if (!isFinite(amt) || amt <= 0) throw new Error('金額必須大於 0，這筆沒有儲存');
  if (d.type === 'transfer') {
    if (!d.accountId || !d.toAccountId) throw new Error('轉帳必須同時有轉出與轉入帳戶，這筆沒有儲存');
    if (String(d.accountId) === String(d.toAccountId)) throw new Error('轉出與轉入帳戶不能相同，這筆沒有儲存');
  }
}

function txDelta_(type, amount, payer) {
  const amt = Number(amount) || 0;
  if (type === 'expense') return -amt;
  if (type === 'income') return amt;
  if (type === 'settlement') return payer === 'partner' ? amt : -amt;
  if (type === 'project') return payer === 'in' ? amt : -amt; // 專案結案產生的帳戶增減，不計入收支統計
  return 0;
}


function adjustAccountBalance_(accountId, delta) {
  if (!accountId || !delta) return null;
  const rowIndex = findRowIndexById_('accounts', accountId);
  if (rowIndex === -1) return null;
  const sheet = getSheet_('accounts');
  const obj = readRowObj_('accounts', rowIndex);
  obj.balance = (Number(obj.balance) || 0) + delta;
  obj.updatedAt = new Date().toISOString();
  const row = buildRowArray_('accounts', obj);
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return obj;
}


function accumulateBalanceDelta_(data, sign, deltas) {
  if (data.type === 'transfer') {
    const amt = Number(data.amount) || 0;
    if (data.accountId) deltas[data.accountId] = (deltas[data.accountId] || 0) + (-amt * sign);
    if (data.toAccountId) deltas[data.toAccountId] = (deltas[data.toAccountId] || 0) + (amt * sign);
  } else if (data.accountId) {
    deltas[data.accountId] = (deltas[data.accountId] || 0) + txDelta_(data.type, data.amount, data.payer) * sign;
  }
}

// 把算好的「每個帳戶淨變化多少」實際讀寫進 Accounts 表。
// 效能：帳戶表很小，整張只讀一次，記憶體裡找出要改的列，每個帳戶只寫一次
// （以前每個帳戶要 讀id欄 + 讀整列 + 寫整列 共三次試算表呼叫）。
function applyBalanceDeltas_(deltas) {
  const ids = Object.keys(deltas).filter(id => deltas[id]);
  if (!ids.length) return [];
  const sheet = getSheet_('accounts');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headerRow = getActualHeaderRow_('accounts');
  const idIdx = headerRow.indexOf('id');
  const values = sheet.getRange(2, 1, lastRow - 1, headerRow.length).getValues();
  const rowOf = {};
  for (let i = 0; i < values.length; i++) rowOf[String(values[i][idIdx])] = i;
  const touched = [];
  const now = new Date().toISOString();
  ids.forEach(accountId => {
    const i = rowOf[String(accountId)];
    if (i === undefined) return;
    const obj = {};
    headerRow.forEach((h, hi) => { if (h) obj[h] = normalizeCellValue_(h, values[i][hi]); });
    obj.balance = (Number(obj.balance) || 0) + deltas[accountId];
    obj.updatedAt = now;
    const row = buildRowArray_('accounts', obj);
    sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
    touched.push(obj);
  });
  return touched;
}

function applyBalanceEffect_(data, sign) {
  const deltas = {};
  accumulateBalanceDelta_(data, sign, deltas);
  return applyBalanceDeltas_(deltas);
}


function txDupKey_(data) {
  const date = normalizeCellValue_('date', data.date);
  const amt = Math.round((Number(data.amount) || 0) * 100);
  return [date, data.type, data.category || '', data.subcategory || '', amt].join('|');
}


function findDuplicateTransaction_(data, excludeId) {
  if (data.type === 'transfer' || data.type === 'settlement' || data.type === 'project' || data.projectId) return null;
  const sheet = getSheet_('transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const headerRow = getActualHeaderRow_('transactions'); // 依實際表頭順序，不假設固定欄號
  const idIdx = headerRow.indexOf('id');
  const dateIdx = headerRow.indexOf('date');
  if (dateIdx === -1) return null; // 表頭異常找不到日期欄，安全起見放棄比對，不擋記帳

  const numRows = lastRow - 1;
  const targetDate = normalizeCellValue_('date', data.date);
  const dateValues = sheet.getRange(2, dateIdx + 1, numRows, 1).getValues();
  const matchedRows = [];
  for (let i = 0; i < dateValues.length; i++) {
    if (normalizeCellValue_('date', dateValues[i][0]) === targetDate) matchedRows.push(i + 2); // 轉成 1-based 列號
  }
  if (!matchedRows.length) return null;

  const lastCol = sheet.getLastColumn();
  const targetKey = txDupKey_(data);
  for (let i = 0; i < matchedRows.length; i++) {
    const row = matchedRows[i];
    const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    if (values[idIdx] === '' || values[idIdx] === null) continue;
    if (excludeId && String(values[idIdx]) === String(excludeId)) continue;
    const obj = {};
    headerRow.forEach((h, hi) => { if (h) obj[h] = normalizeCellValue_(h, values[hi]); });
    if (txDupKey_(obj) === targetKey) return obj;
  }
  return null;
}

function buildTransactionKeySet_() {
  const set = {};
  const sheet = getSheet_('transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;
  const lastCol = sheet.getLastColumn();
  const headerRow = getActualHeaderRow_('transactions'); // 依實際表頭順序，不假設固定欄號
  const idIdx = headerRow.indexOf('id');
  const dateIdx = headerRow.indexOf('date');
  const typeIdx = headerRow.indexOf('type');
  const catIdx = headerRow.indexOf('category');
  const subIdx = headerRow.indexOf('subcategory');
  const amtIdx = headerRow.indexOf('amount');
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  values.forEach(r => {
    if (r[idIdx] === '' || r[idIdx] === null) return;
    const date = normalizeCellValue_('date', r[dateIdx]);
    const amt = Math.round((Number(r[amtIdx]) || 0) * 100);
    set[[date, r[typeIdx], r[catIdx] || '', r[subIdx] || '', amt].join('|')] = true;
  });
  return set;
}


function addTransactionTx_(data) {
  const force = !!data.force;
  const isRetry = !!data.retry;
  delete data.force;
  delete data.retry;
  normalizePayerAccount_(data);
  if (!(isRetry && data.id && findRowIndexById_('transactions', data.id) !== -1)) validateTransactionData_(data);
  // 前端是先在本機存好才背景送出，網路逾時後會「重送」同一筆（帶同一個 id）。
  // 重送時先確認這個 id 是否其實已經寫進去了，避免重複記帳／重複扣餘額。
  // 第一次送出（絕大多數情況）不做這個檢查，維持最快速度。
  if (isRetry && data.id && findRowIndexById_('transactions', data.id) !== -1) {
    return { transaction: data, account: null, accounts: [], replayed: true };
  }
  if (!force) {
    const dup = findDuplicateTransaction_(data);
    if (dup) return { duplicate: dup };
  }
  const tx = addRow_('transactions', data);
  const accountsTouched = applyBalanceEffect_(tx, 1);
  return { transaction: tx, account: accountsTouched[0] || null, accounts: accountsTouched };
}


function updateTransactionTx_(data) {
  const sheet = getSheet_('transactions');
  const rowIndex = findRowIndexById_('transactions', data.id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳: ' + data.id);

  const force = !!data.force;
  delete data.force;
  delete data.retry;
  normalizePayerAccount_(data);
  if (!force) validateTransactionData_(data); // 遷移工具等內部呼叫帶 force，不受此限
  if (!force) {
    const dup = findDuplicateTransaction_(data, data.id);
    if (dup) return { duplicate: dup };
  }

  const old = readRowObj_('transactions', rowIndex);

  // 後端保險：已結算的紀錄，編輯時不能被改回未結算（結算只能由 settleLedger 設定）
  if (String(old.settled) === 'true') data.settled = true;
  const tx = updateRow_('transactions', data, rowIndex);


  const deltas = {};
  accumulateBalanceDelta_(old, -1, deltas);
  accumulateBalanceDelta_(tx, 1, deltas);
  const accountsTouched = applyBalanceDeltas_(deltas);
  return { transaction: tx, accounts: accountsTouched };
}

// 待付款項目「確認付款」：一次完成「新增一筆正式支出（含扣帳戶餘額）＋移除這筆待付款」，
// 放在同一個後端呼叫裡，不會出現「記帳寫進去了但待付款還留著」造成預算重複計算的情況。
// 前端網路逾時重送時（retry），若這筆記帳的 id 已經寫入，就只確保待付款被移除，不會重複記帳／重複扣餘額。
function payOrderTx_(data) {
  const orderId = data.orderId;
  delete data.orderId;
  delete data.force;
  delete data.retry;
  normalizePayerAccount_(data);
  const orderRow = orderId ? findRowIndexById_('orders', orderId) : -1;
  if (data.id && findRowIndexById_('transactions', data.id) !== -1) {
    if (orderRow !== -1) getSheet_('orders').deleteRow(orderRow);
    return { transaction: data, account: null, accounts: [], replayed: true };
  }
  const tx = addRow_('transactions', data);
  const accountsTouched = applyBalanceEffect_(tx, 1);
  if (orderRow !== -1) getSheet_('orders').deleteRow(orderRow);
  return { transaction: tx, account: accountsTouched[0] || null, accounts: accountsTouched };
}

// 把一筆已記好的「支出」改成「待付款」：同一個後端呼叫內
// 先新增待付款、再刪除這筆記帳並還原帳戶餘額，不會出現兩邊都存在而重複計算預算。
// 每月固定繳款：把「模板」裡填了「每月幾號繳」的支出模板，每個月自動新增一筆「待付款」。
// 冪等：同一個模板同一個月只會產生一次——已經有（或曾經有、後來被取消的）待付款、
// 或這個月已經有帶同一個 recurringId 的正式記帳，就不會再新增，多台裝置同時開啟也不會重複。
function generateRecurringOrders_(month) {
  month = String(month || '');
  if (!/^\d{4}-\d{2}$/.test(month)) {
    const tz = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    month = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  }
  const props = PropertiesService.getScriptProperties();
  let keys = {};
  try { keys = JSON.parse(props.getProperty('RECURRING_ORDER_KEYS') || '{}') || {}; } catch (err) { keys = {}; }
  const done = {};
  (keys[month] || []).forEach(id => { done[String(id)] = true; });
  readAll_('orders').forEach(o => {
    if (o.recurringId && String(o.orderDate || '').indexOf(month) === 0) done[String(o.recurringId)] = true;
  });
  readTransactionsSince_(month + '-01').rows.forEach(t => {
    if (t.recurringId && String(t.date || '').indexOf(month) === 0) done[String(t.recurringId)] = true;
  });

  const added = [];
  readAll_('recurring').forEach(t => {
    if (t.type !== 'expense') return;
    if (t.active === false || String(t.active) === 'false') return;
    const day = parseInt(t.dayOfMonth, 10);
    if (!day || day < 1 || day > 31) return;
    if (done[String(t.id)]) return;
    const shared = t.splitMode === 'split' || t.splitMode === 'advance';
    const dd = ('0' + Math.min(day, 28)).slice(-2);
    const order = addRow_('orders', {
      name: t.name || t.note || t.category || '',
      amount: Number(t.amount) || 0,
      category: t.category || '其他',
      subcategory: t.subcategory || '',
      orderDate: month + '-' + dd,
      note: t.note || '',
      payer: shared ? (t.payer || 'me') : 'me',
      splitMode: shared ? t.splitMode : 'personal',
      recurringId: t.id
    });
    done[String(t.id)] = true;
    (keys[month] = keys[month] || []).push(t.id);
    added.push(order);
  });

  if (added.length) {
    const kept = {};
    Object.keys(keys).forEach(m => { if (m >= month) kept[m] = keys[m]; });
    props.setProperty('RECURRING_ORDER_KEYS', JSON.stringify(kept));
  }
  return { added: added, month: month };
}

function convertTxToOrderTx_(data) {
  const rowIndex = findRowIndexById_('transactions', data.id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳（可能已被刪除或轉換）');
  const old = readRowObj_('transactions', rowIndex);
  if (old.type !== 'expense') throw new Error('只有支出可以改成待付款');
  const oldShared = old.splitMode === 'split' || old.splitMode === 'advance';
  if (oldShared && String(old.settled) === 'true') throw new Error('這筆共同帳本項目已經結算過，不能改成待付款');
  // 前端會帶入編輯表單上目前的內容；沒帶的欄位就沿用原本那筆記帳
  const splitMode = data.splitMode !== undefined ? data.splitMode : (oldShared ? old.splitMode : 'personal');
  const isShared = splitMode === 'split' || splitMode === 'advance';
  const amount = data.amount !== undefined ? Number(data.amount) : Number(old.amount);
  const order = addRow_('orders', {
    name: data.name || old.note || old.category || '',
    amount: amount || 0,
    category: data.category || old.category || '其他',
    subcategory: data.subcategory || '',
    orderDate: data.orderDate || old.date,
    note: '',
    payer: isShared ? (data.payer || old.payer || 'me') : 'me',
    splitMode: isShared ? splitMode : 'personal',
    recurringId: old.recurringId || '' // 保留固定支出標記，轉成待付款再付款後統計才不會變彈性
  });
  getSheet_('transactions').deleteRow(rowIndex);
  const accountsTouched = applyBalanceEffect_(old, -1);
  return { order: order, id: data.id, accounts: accountsTouched };
}

function deleteTransactionTx_(id) {
  const sheet = getSheet_('transactions');
  const rowIndex = findRowIndexById_('transactions', id);

  if (rowIndex === -1) return { id: id, account: null, accounts: [] };
  const old = readRowObj_('transactions', rowIndex);
  sheet.deleteRow(rowIndex);
  const accountsTouched = applyBalanceEffect_(old, -1);
  return { id: id, account: accountsTouched[0] || null, accounts: accountsTouched };
}


// 效能：以前是每個 id 各自呼叫 deleteTransactionTx_ ——每一筆都要「先讀整欄 id 找列號、
// 再讀那一整列、再刪那一列」，同一個帳戶如果被好幾筆命中，帳戶餘額也會被讀寫好幾次。
// 勾選很多筆一次刪除時（批次刪除功能就是為了這個情境設計的），這樣做會變成好幾十次
// 來回的試算表操作，感覺特別慢。改成：一次讀整張表找出要刪的列＋加總每個帳戶的
// 淨變化，帳戶餘額每個帳戶只讀寫一次，最後才由下往上實際刪列（由下往上刪是因為
// 刪掉前面的列之後，後面待刪的列號會往上位移，由下往上刪才不會刪錯列）。
function batchDeleteTransactionsTx_(ids) {
  const idList = (ids || []).filter(function (id) { return id !== null && id !== undefined && id !== ''; });
  if (!idList.length) return { ids: [], accounts: [] };

  const sheet = getSheet_('transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ids: [], accounts: [] };

  const headerRow = getActualHeaderRow_('transactions');
  const idIdx = headerRow.indexOf('id');
  if (idIdx === -1) return { ids: [], accounts: [] };

  const lastCol = sheet.getLastColumn();
  const numRows = lastRow - 1;
  const values = sheet.getRange(2, 1, numRows, lastCol).getValues();

  const idSet = {};
  idList.forEach(function (id) { idSet[String(id)] = true; });

  const deltas = {};
  const deletedIds = [];
  const rowsToDelete = []; // 1-based 列號，稍後由後往前刪

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rid = row[idIdx];
    if (rid === '' || rid === null || !idSet[String(rid)]) continue;
    const obj = {};
    headerRow.forEach(function (h, hi) { if (h) obj[h] = normalizeCellValue_(h, row[hi]); });
    accumulateBalanceDelta_(obj, -1, deltas);
    deletedIds.push(rid);
    rowsToDelete.push(i + 2);
  }

  // 每個帳戶的淨變化一次算好、一次讀寫，不管這次刪幾筆，同一個帳戶最多只會讀寫一次
  const accountsTouched = applyBalanceDeltas_(deltas);

  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (row) { sheet.deleteRow(row); });

  return { ids: deletedIds, accounts: accountsTouched };
}


function batchAdd_(key, rows) {
  const sheet = getSheet_(key);
  const now = new Date().toISOString();
  const isTx = key === 'transactions';
  const existingKeys = isTx ? buildTransactionKeySet_() : null;
  const seenKeys = {};
  let skipped = 0;
  const matrix = [];
  (rows || []).forEach(r => {
    if (isTx && r.type !== 'transfer' && r.type !== 'settlement' && r.type !== 'project') {
      const k = txDupKey_(r);
      if (existingKeys[k] || seenKeys[k]) { skipped++; return; }
      seenKeys[k] = true;
    }
    const data = Object.assign({}, r);
    if (isTx) normalizePayerAccount_(data);
    data.id = Utilities.getUuid();
    data.updatedAt = now;
    matrix.push(buildRowArray_(key, data)); // 依實際欄位順序組列，不假設固定順序
  });
  if (matrix.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, matrix.length, matrix[0].length).setValues(matrix);
  }
  return { count: matrix.length, skipped: skipped };
}


// 效能：以前用 getDataRange() 把整張表「每一欄」都讀出來，其實只需要 id 跟 settled
// 兩欄；而且命中的每一列都各自呼叫一次 setValue，勾選很多筆一次結算時，等於好幾十次
// 各自獨立的試算表寫入。改成只讀需要的兩欄，命中的列直接在記憶體裡的陣列上改成
// true，最後不管改了幾列，一次 setValues 寫回整欄，讀寫都只各做一次。
function settleLedger_(ids, settlementData) {
  const sheet = getSheet_('transactions');
  const idCol = colIndex_('transactions', 'id');
  const settledCol = colIndex_('transactions', 'settled');
  const idSet = {};
  (ids || []).forEach(id => (idSet[String(id)] = true));

  const lastRow = sheet.getLastRow();
  let updated = 0;
  if (lastRow >= 2) {
    const numRows = lastRow - 1;
    const idValues = sheet.getRange(2, idCol, numRows, 1).getValues();
    const settledValues = sheet.getRange(2, settledCol, numRows, 1).getValues();
    for (let i = 0; i < numRows; i++) {
      if (idSet[String(idValues[i][0])]) {
        settledValues[i][0] = true;
        updated++;
      }
    }
    if (updated) sheet.getRange(2, settledCol, numRows, 1).setValues(settledValues);
  }

  let settlement = null;
  let account = null;
  if (settlementData) {
    const result = addTransactionTx_(settlementData);
    settlement = result.transaction;
    account = result.account;
  }
  return { updated: updated, settlement: settlement, account: account };
}

function categorySheetRows_() {
  const sheet = getSheet_('categories');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheet.getLastColumn();
  const headerRow = getActualHeaderRow_('categories'); // 依實際表頭順序，不假設固定欄號
  const idIdx = headerRow.indexOf('id');
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];
  values.forEach(v => {
    if (v[idIdx] === '' || v[idIdx] === null) return; // 空列跳過
    const obj = {};
    headerRow.forEach((h, i) => { if (h) obj[h] = normalizeCellValue_(h, v[i]); });
    rows.push(obj);
  });
  return rows;
}

function rewriteCategorySheet_(rows) {
  const sheet = getSheet_('categories');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (rows.length) {
    const matrix = rows.map(r => buildRowArray_('categories', r)); // 依實際欄位順序組列
    sheet.getRange(2, 1, matrix.length, matrix[0].length).setValues(matrix);
  }
}

function buildCategoryTree_(rows) {
  const tree = { expense: [], income: [] };
  const mainIndex = { expense: {}, income: {} };
  rows.forEach(r => {
    const type = r.type === 'income' ? 'income' : 'expense';
    const mainName = String(r.mainName || '').trim();
    if (!mainName) return;
    let m = mainIndex[type][mainName];
    if (!m) {
      m = { name: mainName, subs: [] };
      mainIndex[type][mainName] = m;
      tree[type].push(m);
    }
    const subName = String(r.subName || '').trim();
    if (subName && m.subs.indexOf(subName) === -1) m.subs.push(subName);
  });
  return tree;
}

function getCategoryTreeData_() {
  return buildCategoryTree_(categorySheetRows_());
}

function addMainCategory_(type, name) {
  name = String(name || '').trim();
  if (!name) throw new Error('名稱不能為空');
  const rows = categorySheetRows_();
  if (rows.some(r => r.type === type && String(r.mainName) === name)) throw new Error('主類別已存在');
  getSheet_('categories').appendRow([Utilities.getUuid(), type, name, '', new Date().toISOString()]);

  return { name: name };
}

function addSubCategory_(type, mainName, subName) {
  subName = String(subName || '').trim();
  if (!subName) throw new Error('名稱不能為空');
  const rows = categorySheetRows_();
  if (!rows.some(r => r.type === type && String(r.mainName) === mainName)) throw new Error('主類別不存在');
  if (rows.some(r => r.type === type && String(r.mainName) === mainName && String(r.subName) === subName)) throw new Error('子類別已存在');
  getSheet_('categories').appendRow([Utilities.getUuid(), type, mainName, subName, new Date().toISOString()]);
  // 效能考量同 addMainCategory_，故意不自動重建整表下拉驗證，見上方註解。
  return { name: subName };
}

function removeMainCategory_(type, name) {
  const rows = categorySheetRows_();
  const kept = rows.filter(r => !(r.type === type && String(r.mainName) === name));
  rewriteCategorySheet_(kept);

  return { removed: rows.length - kept.length };
}

function removeSubCategory_(type, mainName, subName) {
  const rows = categorySheetRows_();
  const kept = rows.filter(r => !(r.type === type && String(r.mainName) === mainName && String(r.subName) === subName));
  rewriteCategorySheet_(kept);
  // 效能考量同 addMainCategory_，故意不自動重建整表下拉驗證，見上方註解。
  return { removed: rows.length - kept.length };
}


function renameMainCategory_(type, oldName, newName) {
  newName = String(newName || '').trim();
  if (!newName) throw new Error('名稱不能為空');
  const rows = categorySheetRows_();
  if (oldName !== newName && rows.some(r => r.type === type && String(r.mainName) === newName)) {
    throw new Error('已有相同名稱的主類別');
  }
  const now = new Date().toISOString();
  let changed = 0;
  rows.forEach(r => {
    if (r.type === type && String(r.mainName) === oldName) {
      r.mainName = newName;
      r.updatedAt = now;
      changed++;
    }
  });
  if (!changed) throw new Error('找不到這個主類別');
  rewriteCategorySheet_(rows);
  const txResult = renameCategoryEverywhere_({ txType: type, level: 'main', oldValue: oldName, newValue: newName });

  return { changed: changed, transactionsUpdated: txResult.updated };
}

function renameSubCategory_(type, mainName, oldSub, newSub) {
  newSub = String(newSub || '').trim();
  if (!newSub) throw new Error('名稱不能為空');
  const rows = categorySheetRows_();
  if (oldSub !== newSub && rows.some(r => r.type === type && String(r.mainName) === mainName && String(r.subName) === newSub)) {
    throw new Error('已有相同名稱的子類別');
  }
  const now = new Date().toISOString();
  let changed = 0;
  rows.forEach(r => {
    if (r.type === type && String(r.mainName) === mainName && String(r.subName) === oldSub) {
      r.subName = newSub;
      r.updatedAt = now;
      changed++;
    }
  });
  if (!changed) throw new Error('找不到這個子類別');
  rewriteCategorySheet_(rows);
  const txResult = renameCategoryEverywhere_({ txType: type, level: 'sub', mainName: mainName, oldValue: oldSub, newValue: newSub });
  // 效能考量同 renameMainCategory_，故意不自動重建整表下拉驗證，見上方註解。
  return { changed: changed, transactionsUpdated: txResult.updated };
}


function renameCategoryEverywhere_(payload) {
  const txType = payload && payload.txType;
  const level = payload && payload.level;
  const oldValue = String((payload && payload.oldValue) || '');
  const newValue = String((payload && payload.newValue) || '');
  const mainName = (payload && payload.mainName) || '';
  if (!txType || (level !== 'main' && level !== 'sub') || !oldValue || !newValue) {
    throw new Error('缺少必要參數');
  }

  let totalUpdated = 0;
  ['transactions', 'recurring'].forEach(key => {
    const sheet = getSheet_(key);
    const headerRow = getActualHeaderRow_(key); // 依實際表頭順序，不假設固定欄號
    const idCol = headerRow.indexOf('id');
    const typeCol = headerRow.indexOf('type');
    const catCol = headerRow.indexOf('category');
    const subCol = headerRow.indexOf('subcategory');
    const updatedAtCol = headerRow.indexOf('updatedAt');
    if (typeCol === -1 || catCol === -1) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const lastCol = sheet.getLastColumn();
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues();
    const now = new Date().toISOString();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (row[idCol] === '' || row[idCol] === null) continue;
      if (row[typeCol] !== txType) continue;

      if (level === 'main') {
        if (row[catCol] !== oldValue) continue;
        row[catCol] = newValue;
      } else {
        if (subCol === -1) continue;
        if (row[catCol] !== mainName || row[subCol] !== oldValue) continue;
        row[subCol] = newValue;
      }
      if (updatedAtCol !== -1) row[updatedAtCol] = now;
      changed = true;
      totalUpdated++;
    }
    if (changed) range.setValues(values);
  });

  return { updated: totalUpdated };
}


function reorderMainCategories_(type, order) {
  const rows = categorySheetRows_();
  const rowsOfType = rows.filter(r => r.type === type);
  const rowsOtherType = rows.filter(r => r.type !== type);

  const groups = {};
  const groupOrder = [];
  rowsOfType.forEach(r => {
    const name = String(r.mainName);
    if (!groups[name]) { groups[name] = []; groupOrder.push(name); }
    groups[name].push(r);
  });

  const requested = (Array.isArray(order) ? order : []).map(String);
  const finalOrder = [];
  requested.forEach(n => { if (groups[n] && finalOrder.indexOf(n) === -1) finalOrder.push(n); });
  groupOrder.forEach(n => { if (finalOrder.indexOf(n) === -1) finalOrder.push(n); });

  const reordered = [];
  finalOrder.forEach(n => { reordered.push.apply(reordered, groups[n]); });

  rewriteCategorySheet_(rowsOtherType.concat(reordered));
  // 純粹調整順序，不會動到任何一列的文字內容，完全不需要重建下拉驗證。
  return { order: finalOrder };
}


function replaceCategoryTree_(type, tree) {
  const rows = categorySheetRows_();
  const kept = rows.filter(r => r.type !== type);
  const now = new Date().toISOString();
  (tree || []).forEach(m => {
    const mainName = String((m && m.name) || '').trim();
    if (!mainName) return;
    const subs = (m && m.subs) || [];
    if (!subs.length) {
      kept.push({ id: Utilities.getUuid(), type: type, mainName: mainName, subName: '', updatedAt: now });
    } else {
      subs.forEach(s => {
        const subName = String(s || '').trim();
        if (!subName) return;
        kept.push({ id: Utilities.getUuid(), type: type, mainName: mainName, subName: subName, updatedAt: now });
      });
    }
  });
  rewriteCategorySheet_(kept);
  rebuildAllTransactionValidations();
  return { count: kept.filter(r => r.type === type).length };
}






function keepWarm_() {

  return 'ok';
}

function setupKeepWarmTrigger() {
  removeKeepWarmTrigger();
  ScriptApp.newTrigger('keepWarm_')
    .timeBased()
    .everyMinutes(10)
    .create();
  return '已設定：每 10 分鐘會自動保溫一次，減少冷啟動等待。';
}

function removeKeepWarmTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'keepWarm_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}


function readTransactionsSince_(since) {
  const sheet = getSheet_('transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], fullCount: 0 };

  const headerRow = getActualHeaderRow_('transactions'); // 依實際表頭順序，不假設固定欄號
  const idIdx = headerRow.indexOf('id');
  const dateIdx = headerRow.indexOf('date');
  const idColIndex = idIdx + 1;
  const dateColIndex = dateIdx + 1;
  const minCol = Math.min(idColIndex, dateColIndex);
  const span = Math.abs(idColIndex - dateColIndex) + 1;
  const idxValues = sheet.getRange(2, minCol, lastRow - 1, span).getValues();
  const idOffset = idColIndex - minCol;
  const dateOffset = dateColIndex - minCol;

  let fullCount = 0;
  let minMatchIdx = -1;
  for (let i = 0; i < idxValues.length; i++) {
    const idVal = idxValues[i][idOffset];
    if (idVal === '' || idVal === null) continue; // 空列（刪除後留下的）跳過
    fullCount++;
    if (minMatchIdx === -1) {
      const d = normalizeCellValue_('date', idxValues[i][dateOffset]);
      if (!since || (d && d >= since)) minMatchIdx = i;
    }
  }
  if (minMatchIdx === -1) return { rows: [], fullCount: fullCount };

  const startRow = 2 + minMatchIdx;
  const numRows = lastRow - startRow + 1;
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (r[idIdx] === '' || r[idIdx] === null) continue;
    const obj = {};
    headerRow.forEach((h, hi) => { if (h) obj[h] = normalizeCellValue_(h, r[hi]); });
    if (!since || (obj.date || '') >= since) rows.push(obj);
  }
  return { rows: rows, fullCount: fullCount };
}


function txTypeCol_() { return colIndex_('transactions', 'type'); }
function txCategoryCol_() { return colIndex_('transactions', 'category'); }
function txSubcategoryCol_() { return colIndex_('transactions', 'subcategory'); }

// 把 Categories 樹整理成 'type||主類別名稱' -> 子類別陣列 的查詢表，方便查詢。
function buildCategorySubsMap_(tree) {
  const subsMap = {};
  Object.keys(tree).forEach(type => {
    (tree[type] || []).forEach(m => { subsMap[type + '||' + m.name] = m.subs || []; });
  });
  return subsMap;
}

function applyValidationRange_(sheet, startRow, typeValues, mainNameValues, tree) {
  try {
    const numRows = typeValues.length;
    if (!numRows) return;
    const subsMap = buildCategorySubsMap_(tree);
    const catRules = new Array(numRows);
    const subRules = new Array(numRows);
    for (let i = 0; i < numRows; i++) {
      const type = typeValues[i];
      const mainName = mainNameValues[i];
      const mainNames = (tree[type] || []).map(m => m.name);
      catRules[i] = [mainNames.length
        ? SpreadsheetApp.newDataValidation().requireValueInList(mainNames, true).setAllowInvalid(true).build()
        : null];
      const subs = subsMap[type + '||' + mainName] || [];
      subRules[i] = [subs.length
        ? SpreadsheetApp.newDataValidation().requireValueInList(subs, true).setAllowInvalid(true).build()
        : null];
    }
    sheet.getRange(startRow, txCategoryCol_(), numRows, 1).setDataValidations(catRules);
    sheet.getRange(startRow, txSubcategoryCol_(), numRows, 1).setDataValidations(subRules);
  } catch (err) {
    console.error('applyValidationRange_ 失敗（不影響記帳/類別本身的資料）: ' + err);
  }
}

function checkCategoryMismatches() {
  const sheet = getSheet_('transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('沒有記帳資料。'); return '沒有記帳資料。'; }

  const numRows = lastRow - 1;
  const typeCol = txTypeCol_();
  const catCol = txCategoryCol_();
  const minCol = Math.min(typeCol, catCol);
  const maxCol = Math.max(typeCol, catCol);
  const values = sheet.getRange(2, minCol, numRows, maxCol - minCol + 1).getValues();
  const typeOffset = typeCol - minCol;
  const catOffset = catCol - minCol;

  const tree = getCategoryTreeData_();
  const validMainNames = {}; // 'type||mainName' -> true
  Object.keys(tree).forEach(type => {
    (tree[type] || []).forEach(m => { validMainNames[type + '||' + m.name] = true; });
  });

  const mismatchCounts = {}; // 'type||mainName' -> 筆數
  let matchedCount = 0;
  for (let i = 0; i < values.length; i++) {
    const type = String(values[i][typeOffset] || '').trim();
    const mainName = String(values[i][catOffset] || '').trim();
    if (type !== 'expense' && type !== 'income') continue; // 轉帳/結算不適用主類別下拉
    const key = type + '||' + mainName;
    if (validMainNames[key]) {
      matchedCount++;
    } else {
      mismatchCounts[key] = (mismatchCounts[key] || 0) + 1;
    }
  }

  const lines = [];
  lines.push('對得上 Categories 目前清單的記帳列：' + matchedCount + ' 筆。');
  const mismatchKeys = Object.keys(mismatchCounts);
  if (!mismatchKeys.length) {
    lines.push('沒有任何一筆的主類別文字對不上 Categories 清單——如果子類別下拉還是沒出現，' +
      '很可能是那個主類別本身在 Categories 裡就沒有設定任何子類別（例如「其他」）。');
  } else {
    lines.push('對不上 Categories 清單的主類別文字，共 ' + mismatchKeys.length + ' 種：');
    mismatchKeys
      .sort((a, b) => mismatchCounts[b] - mismatchCounts[a])
      .forEach(key => {
        const parts = key.split('||');
        lines.push('  type=' + parts[0] + '，主類別="' + parts[1] + '"，' + mismatchCounts[key] + ' 筆');
      });
    lines.push('這些文字要嘛是打錯字/多空格，要嘛是 Categories 裡已經改名或刪掉了但沒有回頭改這些舊紀錄。');
  }
  const report = lines.join('\n');
  Logger.log(report);
  return report;
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAMES.transactions) return;
    const row = e.range.getRow();
    if (row < 2) return; // 表頭不處理

    const typeCol = txTypeCol_();
    const catCol = txCategoryCol_();
    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    // 貼上/拖曳填滿可能一次改到多欄多列，只要範圍有碰到 type 或 category 欄，
    // 就把「整個被編輯到的這段列」一起重新算一次驗證。
    const touchesRelevantCol = (startCol <= typeCol && typeCol <= endCol) ||
      (startCol <= catCol && catCol <= endCol);
    if (!touchesRelevantCol) return;

    const numRows = e.range.getNumRows();
    const minCol = Math.min(typeCol, catCol);
    const maxCol = Math.max(typeCol, catCol);
    const values = sheet.getRange(row, minCol, numRows, maxCol - minCol + 1).getValues();
    const typeOffset = typeCol - minCol;
    const catOffset = catCol - minCol;
    const typeValues = values.map(r => String(r[typeOffset] || '').trim());
    const mainNameValues = values.map(r => String(r[catOffset] || '').trim());
    applyValidationRange_(sheet, row, typeValues, mainNameValues, getCategoryTreeData_());
  } catch (err) {

  }
}


function applyTransactionRowValidation_(sheet, row) {
  try {
    const type = String(sheet.getRange(row, txTypeCol_()).getValue() || '').trim();
    const mainName = String(sheet.getRange(row, txCategoryCol_()).getValue() || '').trim();
    applyValidationRange_(sheet, row, [type], [mainName], getCategoryTreeData_());
  } catch (err) {
    console.error('applyTransactionRowValidation_ 失敗（不影響這筆記帳本身的儲存）: ' + err);
  }
}


function rebuildAllTransactionValidations() {
  try {
    const sheet = getSheet_('transactions');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return '沒有資料列，不需要處理。';
    const numRows = lastRow - 1;
    const typeCol = txTypeCol_();
    const catCol = txCategoryCol_();
    const minCol = Math.min(typeCol, catCol);
    const maxCol = Math.max(typeCol, catCol);
    const values = sheet.getRange(2, minCol, numRows, maxCol - minCol + 1).getValues();
    const typeOffset = typeCol - minCol;
    const catOffset = catCol - minCol;
    const typeValues = values.map(r => String(r[typeOffset] || '').trim());
    const mainNameValues = values.map(r => String(r[catOffset] || '').trim());
    applyValidationRange_(sheet, 2, typeValues, mainNameValues, getCategoryTreeData_());
    return '已重新套用 ' + numRows + ' 列的下拉選單驗證。';
  } catch (err) {
    console.error('rebuildAllTransactionValidations 失敗（不影響 Categories/Transactions 本身的資料）: ' + err);
    return '下拉選單驗證更新失敗（不影響資料本身）：' + err;
  }
}


// ============================================================
// 專案（爸爸／媽媽給的專款）
// 進行中：Projects / ProjectItems 兩張表自成一格，完全不碰 Transactions、不動帳戶餘額。
// 結案：後端重新讀專案與支出明細自己算（不信任前端數字），一次整合進記帳＋帳戶餘額：
//   ‧ 帳戶增減：同一帳戶淨額合併成一筆 type:'project'（不算收入／支出）
//   ‧ 剩餘歸自己 → 一筆收入「專案結餘」（不指定帳戶，只供統計，帳戶增減已含在上面）
//   ‧ 超支要自己補 → 一筆支出「專案補貼」（同上）
//   ‧ 剩餘放進共同錢包 → 該錢包帳戶 +剩餘，沒有收入／支出
// 所有寫進去的列都帶 projectId，「取消結案」會整批移除並還原餘額。
//
// 支出明細的 payer：
//   fund    直接用收到的錢付（從收款帳戶扣）
//   me      我先用自己的帳戶付（accountId；錢由專款補回，所以我保留收款帳戶裡的錢）
//   partner 另一半先付 → 結案時用專款還他（從 reimburse.partner 帳戶出）
//   sister  妹妹先付   → 結案時用專款還她（從 reimburse.sister 帳戶出）
// ============================================================
function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function readProjectRowsByProject_(key, projectId) {
  return readAll_(key).filter(r => String(r.projectId) === String(projectId));
}

function computeProjectClose_(project, items, opts) {
  const F = round2_(project.fundAmount);
  const byAcc = {};   // accountId -> 淨增減
  const add = (accId, delta) => { if (!accId || !delta) return; byAcc[accId] = round2_((byAcc[accId] || 0) + delta); };
  const owed = { partner: 0, sister: 0 };
  let spent = 0;
  const fundAcc = project.fundAccountId;
  add(fundAcc, F);
  items.forEach(it => {
    const amt = round2_(it.amount);
    spent += amt;
    if (it.payer === 'me') add(it.accountId || fundAcc, -amt);
    else if (it.payer === 'partner') owed.partner += amt;
    else if (it.payer === 'sister') owed.sister += amt;
    else add(fundAcc, -amt); // fund
  });
  const rb = opts.reimburse || {};
  ['partner', 'sister'].forEach(k => { if (owed[k] > 0) add(rb[k] || fundAcc, -owed[k]); });
  spent = round2_(spent);
  const remaining = round2_(F - spent);
  const mode = opts.leftoverMode;
  let incomeRow = null, expenseRow = null;
  if (remaining > 0.005) {
    if (mode === 'shared') {
      if (!opts.leftoverAccountId) throw new Error('剩餘要放進共同錢包，請選擇錢包帳戶');
      add(fundAcc, -remaining);
      add(opts.leftoverAccountId, remaining);
    } else { // pocket：錢留在自己手上，記一筆收入供統計
      incomeRow = remaining;
    }
  } else if (remaining < -0.005) {
    expenseRow = -remaining; // 超支，已由上面的實際付款流程自己出了這筆錢
  }
  return { F: F, spent: spent, remaining: remaining, byAcc: byAcc, owed: owed, incomeRow: incomeRow, expenseRow: expenseRow };
}

function closeProjectTx_(data) {
  const rowIndex = findRowIndexById_('projects', data.id);
  if (rowIndex === -1) throw new Error('找不到這個專案');
  const project = readRowObj_('projects', rowIndex);
  if (project.status === 'closed') return { project: project, rows: [], accounts: [], replayed: true };
  const items = readProjectRowsByProject_('projectItems', data.id);
  const opts = {
    leftoverMode: data.leftoverMode === 'shared' ? 'shared' : 'pocket',
    leftoverAccountId: data.leftoverAccountId || '',
    reimburse: data.reimburse || {}
  };
  if (!project.fundAccountId) throw new Error('專案沒有設定收款帳戶，請先編輯專案');
  const calc = computeProjectClose_(project, items, opts);
  const closedAt = String(normalizeCellValue_('date', data.closedAt) || '').trim() ||
    Utilities.formatDate(new Date(), getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const rows = [];
  const makeTx = t => addRow_('transactions', Object.assign({
    date: closedAt, subcategory: '', toAccountId: '', payer: 'me', splitMode: 'personal',
    settled: true, recurringId: '', projectId: project.id
  }, t));
  Object.keys(calc.byAcc).forEach(accId => {
    const d = calc.byAcc[accId];
    if (!d) return;
    rows.push(makeTx({ type: 'project', category: '專案', amount: Math.abs(d), accountId: accId, payer: d > 0 ? 'in' : 'out', note: '專案「' + project.name + '」結案' }));
  });
  if (calc.incomeRow) rows.push(makeTx({ type: 'income', category: '專案結餘', amount: calc.incomeRow, accountId: '', note: '專案「' + project.name + '」剩餘歸自己' }));
  if (calc.expenseRow) rows.push(makeTx({ type: 'expense', category: '專案補貼', amount: calc.expenseRow, accountId: '', note: '專案「' + project.name + '」超支自己補貼' }));

  const deltas = {};
  rows.forEach(r => accumulateBalanceDelta_(r, 1, deltas));
  const accounts = applyBalanceDeltas_(deltas);

  project.status = 'closed';
  project.closedAt = closedAt;
  project.leftoverMode = opts.leftoverMode;
  project.leftoverAccountId = opts.leftoverAccountId;
  project.reimburse = JSON.stringify(opts.reimburse);
  project.result = JSON.stringify({ spent: calc.spent, remaining: calc.remaining, owed: calc.owed });
  const saved = updateRow_('projects', project, rowIndex);
  return { project: saved, rows: rows, accounts: accounts };
}

function removeProjectTransactions_(projectId) {
  const sheet = getSheet_('transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headerRow = getActualHeaderRow_('transactions');
  const idIdx = headerRow.indexOf('id');
  const pIdx = headerRow.indexOf('projectId');
  if (pIdx === -1) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const deltas = {};
  const rowsToDelete = [];
  const ids = [];
  for (let i = 0; i < values.length; i++) {
    if (values[i][idIdx] === '' || String(values[i][pIdx]) !== String(projectId)) continue;
    const obj = {};
    headerRow.forEach((h, hi) => { if (h) obj[h] = normalizeCellValue_(h, values[i][hi]); });
    accumulateBalanceDelta_(obj, -1, deltas);
    rowsToDelete.push(i + 2);
    ids.push(values[i][idIdx]);
  }
  const accounts = applyBalanceDeltas_(deltas);
  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  return { ids: ids, accounts: accounts };
}

function reopenProjectTx_(id) {
  const rowIndex = findRowIndexById_('projects', id);
  if (rowIndex === -1) throw new Error('找不到這個專案');
  const project = readRowObj_('projects', rowIndex);
  const removed = removeProjectTransactions_(id);
  project.status = 'open';
  project.closedAt = '';
  project.result = '';
  const saved = updateRow_('projects', project, rowIndex);
  return { project: saved, removedIds: removed.ids || [], accounts: removed.accounts || [] };
}

function deleteProjectTx_(id) {
  const rowIndex = findRowIndexById_('projects', id);
  if (rowIndex === -1) return { id: id, removedIds: [], accounts: [] };
  const removed = removeProjectTransactions_(id);
  const sheet = getSheet_('projectItems');
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const pCol = colIndex_('projectItems', 'projectId');
    const vals = sheet.getRange(2, pCol, lastRow - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(id)) sheet.deleteRow(i + 2);
    }
  }
  getSheet_('projects').deleteRow(findRowIndexById_('projects', id));
  return { id: id, removedIds: removed.ids || [], accounts: removed.accounts || [] };
}

function jsonOut_(obj) {
  obj.version = BACKEND_VERSION_;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const token = e.parameter.token;
    if (!checkToken_(token)) return jsonOut_({ ok: false, error: '密鑰錯誤' });

    const sheetKey = e.parameter.sheet; // accounts | liabilities | transactions | recurring | all


    if (sheetKey === 'all') {
      let since = null;
      const recentMonths = parseInt(e.parameter.recentMonths, 10);
      if (recentMonths && recentMonths > 0) {
        const tz = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
        const cutoff = new Date();
        cutoff.setMonth(cutoff.getMonth() - recentMonths);
        since = Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd');
      }
      const txResult = readTransactionsSince_(since);
      const truncated = txResult.rows.length < txResult.fullCount;
      return jsonOut_({
        ok: true,
        data: {
          accounts: readAll_('accounts'),
          liabilities: readAll_('liabilities'),
          transactions: txResult.rows,
          recurring: readAll_('recurring'),
          orders: readAll_('orders'),
          projects: readAll_('projects'),
          projectItems: readAll_('projectItems'),
          categories: getCategoryTreeData_(),
          transactionsTruncated: truncated,
          transactionsSince: since
        }
      });
    }

    if (sheetKey === 'categories') return jsonOut_({ ok: true, data: getCategoryTreeData_() });

    if (!SHEET_NAMES[sheetKey]) return jsonOut_({ ok: false, error: '未知的資料表' });

    return jsonOut_({ ok: true, data: readAll_(sheetKey) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}



function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkToken_(body.token)) return jsonOut_({ ok: false, error: '密鑰錯誤' });

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      return jsonOut_({ ok: false, error: '系統忙碌中，請稍後再試一次。' });
    }
    try {
      const action = body.action;

      if (action === 'addTransactionTx') return jsonOut_({ ok: true, data: addTransactionTx_(body.data) });
      if (action === 'payOrderTx') return jsonOut_({ ok: true, data: payOrderTx_(body.data) });
      if (action === 'generateRecurringOrders') return jsonOut_({ ok: true, data: generateRecurringOrders_(body.data && body.data.month) });
      if (action === 'convertTxToOrderTx') return jsonOut_({ ok: true, data: convertTxToOrderTx_(body.data) });
      if (action === 'updateTransactionTx') return jsonOut_({ ok: true, data: updateTransactionTx_(body.data) });
      if (action === 'deleteTransactionTx') return jsonOut_({ ok: true, data: deleteTransactionTx_(body.data.id) });
      if (action === 'batchDeleteTransactionsTx') return jsonOut_({ ok: true, data: batchDeleteTransactionsTx_((body.data && body.data.ids) || []) });
      if (action === 'batchAddTransactions') return jsonOut_({ ok: true, data: batchAdd_('transactions', (body.data && body.data.rows) || []) });
      if (action === 'closeProjectTx') return jsonOut_({ ok: true, data: closeProjectTx_(body.data) });
      if (action === 'reopenProjectTx') return jsonOut_({ ok: true, data: reopenProjectTx_(body.data && body.data.id) });
      if (action === 'deleteProjectTx') return jsonOut_({ ok: true, data: deleteProjectTx_(body.data && body.data.id) });
      if (action === 'settleLedger') return jsonOut_({ ok: true, data: settleLedger_((body.data && body.data.ids) || [], body.data && body.data.settlement) });

      if (action === 'addMainCategory') return jsonOut_({ ok: true, data: addMainCategory_(body.data.type, body.data.name) });
      if (action === 'addSubCategory') return jsonOut_({ ok: true, data: addSubCategory_(body.data.type, body.data.mainName, body.data.subName) });
      if (action === 'removeMainCategory') return jsonOut_({ ok: true, data: removeMainCategory_(body.data.type, body.data.name) });
      if (action === 'removeSubCategory') return jsonOut_({ ok: true, data: removeSubCategory_(body.data.type, body.data.mainName, body.data.subName) });
      if (action === 'renameMainCategory') return jsonOut_({ ok: true, data: renameMainCategory_(body.data.type, body.data.oldName, body.data.newName) });
      if (action === 'renameSubCategory') return jsonOut_({ ok: true, data: renameSubCategory_(body.data.type, body.data.mainName, body.data.oldSub, body.data.newSub) });
      if (action === 'replaceCategoryTree') return jsonOut_({ ok: true, data: replaceCategoryTree_(body.data.type, body.data.tree) });
      if (action === 'reorderMainCategories') return jsonOut_({ ok: true, data: reorderMainCategories_(body.data.type, body.data.order) });

      const sheetKey = body.sheet;
      if (!SHEET_NAMES[sheetKey]) return jsonOut_({ ok: false, error: '未知的資料表' });

      let result;
      if (action === 'add') {
        result = addRow_(sheetKey, body.data);
      } else if (action === 'update') {
        result = updateRow_(sheetKey, body.data);
      } else if (action === 'delete') {
        result = deleteRow_(sheetKey, body.data.id);
      } else {
        return jsonOut_({ ok: false, error: '未知的動作' });
      }
      return jsonOut_({ ok: true, data: result });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
