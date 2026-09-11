/**
 * 個人存款管理系統 - Google Apps Script 後端
 * -------------------------------------------------
 * 用法：
 * 1. 開一份新的 Google 試算表（隨便命名，例如「我的存款管理」）
 * 2. 上方選單「擴充功能」→「Apps Script」，打開 Apps Script 編輯器
 * 3. 把這個檔案的全部內容貼進去，覆蓋預設的 Code.gs
 * 4. 上方「專案設定」→「指令碼屬性」，新增一筆屬性：
 *      SECRET_TOKEN = 你自己設定的一串密碼（英數混合，越長越安全）
 * 5. 點「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *      執行身分：我 (你的帳號)
 *      誰能存取：所有人
 *    部署後會拿到一個網址，形如：
 *      https://script.google.com/macros/s/XXXXXXXX/exec
 *    這個網址 + SECRET_TOKEN 就是前端要用的兩個設定值。
 * 6. 之後每次修改這份程式碼，都要重新「管理部署作業」→ 編輯 → 新版本，
 *    網址才會套用最新程式碼。
 *
 * 效能設計重點：
 * - doGet 支援 sheet=all，一次回傳四張表，前端首次載入只需 1 次 HTTP 請求。
 * - 記帳（新增/編輯/刪除）若有連結帳戶，會在同一次 exec 內同時寫入記帳列與
 *   調整帳戶餘額，前端不需要再額外呼叫一次「更新帳戶」。
 * - CSV 匯入、共同帳本結算都是一次 exec 內用陣列批次寫入，不會一筆一筆來回。
 */

const SHEET_NAMES = {
  accounts: 'Accounts',
  liabilities: 'Liabilities',
  interest: 'InterestRecords',
  transactions: 'Transactions'
};

const SHEET_HEADERS = {
  accounts: ['id', 'name', 'type', 'balance', 'interestRate', 'note', 'updatedAt'],
  liabilities: ['id', 'name', 'amount', 'dueDate', 'paid', 'note', 'updatedAt'],
  interest: ['id', 'accountId', 'date', 'amount', 'note', 'createdAt'],
  // 記帳紀錄：
  //   type       'expense'（支出）| 'income'（收入）| 'settlement'（共同帳本結算轉帳，記錄用不算收支）
  //   category   品項（餐飲/交通/薪資...），CSV 匯入的品項可以是任意文字
  //   payer      共同帳本用：'me'（我）| 'partner'（對方）先付款
  //   splitMode  共同帳本用：'personal'（個人，不分攤）| 'split'（平分50/50）| 'advance'（全額代墊算對方的）
  //   settled    共同帳本用：是否已經結算轉帳完成
  // 注意：新增欄位一律加在陣列「最後面」，不要插在中間，
  // 這樣舊表格用 ensureHeaders_ 自動補欄位時，既有資料的欄位對應才不會跑掉。
  transactions: ['id', 'date', 'type', 'category', 'amount', 'accountId', 'note', 'updatedAt', 'payer', 'splitMode', 'settled']
};

function getSheet_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES[key]);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES[key]);
    sheet.appendRow(SHEET_HEADERS[key]);
  } else {
    ensureHeaders_(sheet, SHEET_HEADERS[key]);
  }
  return sheet;
}

// 如果表格是舊版本、欄位比目前定義的少（例如舊的 Transactions 表沒有 payer/splitMode/settled），
// 自動把缺少的欄位標題補到最後面，資料不會受影響。
function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headers.length > currentHeaders.length) {
    sheet.getRange(1, currentHeaders.length + 1, 1, headers.length - currentHeaders.length)
      .setValues([headers.slice(currentHeaders.length)]);
  }
}

function checkToken_(token) {
  const real = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN');
  return real && token && real === token;
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
      headers.forEach((h, i) => (obj[h] = r[i]));
      return obj;
    });
}

function findRowIndexById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // 1-based row number
  }
  return -1;
}

function readRowObj_(sheet, rowIndex, headers) {
  const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => (obj[h] = values[i]));
  return obj;
}

function addRow_(key, data) {
  const sheet = getSheet_(key);
  const headers = SHEET_HEADERS[key];
  data.id = Utilities.getUuid();
  data.updatedAt = new Date().toISOString();
  if (key === 'interest') data.createdAt = new Date().toISOString();
  const row = headers.map(h => (data[h] !== undefined ? data[h] : ''));
  sheet.appendRow(row);
  return data;
}

function updateRow_(key, data) {
  const sheet = getSheet_(key);
  const headers = SHEET_HEADERS[key];
  const rowIndex = findRowIndexById_(sheet, data.id);
  if (rowIndex === -1) throw new Error('找不到這筆資料: ' + data.id);
  data.updatedAt = new Date().toISOString();
  const row = headers.map(h => (data[h] !== undefined ? data[h] : ''));
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  return data;
}

function deleteRow_(key, id) {
  const sheet = getSheet_(key);
  const rowIndex = findRowIndexById_(sheet, id);
  if (rowIndex === -1) throw new Error('找不到這筆資料: ' + id);
  sheet.deleteRow(rowIndex);
  return { id: id };
}

// ---------- 記帳＋帳戶餘額 複合操作（一次 exec 內完成，減少來回） ----------

function txDelta_(type, amount) {
  const amt = Number(amount) || 0;
  return type === 'expense' ? -amt : (type === 'income' ? amt : 0);
}

function adjustAccountBalance_(accountId, delta) {
  if (!accountId || !delta) return null;
  const sheet = getSheet_('accounts');
  const headers = SHEET_HEADERS.accounts;
  const rowIndex = findRowIndexById_(sheet, accountId);
  if (rowIndex === -1) return null;
  const obj = readRowObj_(sheet, rowIndex, headers);
  obj.balance = (Number(obj.balance) || 0) + delta;
  obj.updatedAt = new Date().toISOString();
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  return obj;
}

function addTransactionTx_(data) {
  const tx = addRow_('transactions', data);
  const account = tx.accountId ? adjustAccountBalance_(tx.accountId, txDelta_(tx.type, tx.amount)) : null;
  return { transaction: tx, account: account };
}

function updateTransactionTx_(data) {
  const sheet = getSheet_('transactions');
  const headers = SHEET_HEADERS.transactions;
  const rowIndex = findRowIndexById_(sheet, data.id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳: ' + data.id);
  const old = readRowObj_(sheet, rowIndex, headers);

  const accountsTouched = [];
  if (old.accountId) {
    const acc = adjustAccountBalance_(old.accountId, -txDelta_(old.type, old.amount));
    if (acc) accountsTouched.push(acc);
  }
  const tx = updateRow_('transactions', data);
  if (tx.accountId) {
    const acc = adjustAccountBalance_(tx.accountId, txDelta_(tx.type, tx.amount));
    if (acc) accountsTouched.push(acc);
  }
  return { transaction: tx, accounts: accountsTouched };
}

function deleteTransactionTx_(id) {
  const sheet = getSheet_('transactions');
  const headers = SHEET_HEADERS.transactions;
  const rowIndex = findRowIndexById_(sheet, id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳: ' + id);
  const old = readRowObj_(sheet, rowIndex, headers);
  sheet.deleteRow(rowIndex);
  const account = old.accountId ? adjustAccountBalance_(old.accountId, -txDelta_(old.type, old.amount)) : null;
  return { id: id, account: account };
}

// 批次新增（CSV 匯入用）：一次 exec 內用陣列寫入所有列，不逐筆來回
function batchAdd_(key, rows) {
  const sheet = getSheet_(key);
  const headers = SHEET_HEADERS[key];
  const now = new Date().toISOString();
  const matrix = (rows || []).map(r => {
    const data = Object.assign({}, r);
    data.id = Utilities.getUuid();
    data.updatedAt = now;
    return headers.map(h => (data[h] !== undefined ? data[h] : ''));
  });
  if (matrix.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, matrix.length, headers.length).setValues(matrix);
  }
  return { count: matrix.length };
}

// 共同帳本結算：把指定的記帳列標記為已結算，一次 exec 內完成
function settleLedger_(ids, settlementData) {
  const sheet = getSheet_('transactions');
  const headers = SHEET_HEADERS.transactions;
  const idColIndex = headers.indexOf('id');
  const settledColIndex = headers.indexOf('settled');
  const idSet = {};
  (ids || []).forEach(id => (idSet[String(id)] = true));

  const values = sheet.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < values.length; i++) {
    if (idSet[String(values[i][idColIndex])]) {
      sheet.getRange(i + 1, settledColIndex + 1).setValue(true);
      updated++;
    }
  }
  let settlement = null;
  if (settlementData) {
    settlement = addRow_('transactions', settlementData);
  }
  return { updated: updated, settlement: settlement };
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const token = e.parameter.token;
    if (!checkToken_(token)) return jsonOut_({ ok: false, error: '密鑰錯誤' });

    const sheetKey = e.parameter.sheet; // accounts | liabilities | interest | transactions | all

    // 一次回傳四張表，前端首次載入只需 1 次 HTTP 請求，減少等待時間
    if (sheetKey === 'all') {
      return jsonOut_({
        ok: true,
        data: {
          accounts: readAll_('accounts'),
          liabilities: readAll_('liabilities'),
          interest: readAll_('interest'),
          transactions: readAll_('transactions')
        }
      });
    }

    if (!SHEET_NAMES[sheetKey]) return jsonOut_({ ok: false, error: '未知的資料表' });

    return jsonOut_({ ok: true, data: readAll_(sheetKey) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// 前端用 text/plain 送出 JSON，避免瀏覽器觸發 CORS 預檢請求（preflight）
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkToken_(body.token)) return jsonOut_({ ok: false, error: '密鑰錯誤' });

    const action = body.action;

    // 複合操作：記帳＋帳戶餘額、批次匯入、共同帳本結算，都在同一次 exec 內完成
    if (action === 'addTransactionTx') return jsonOut_({ ok: true, data: addTransactionTx_(body.data) });
    if (action === 'updateTransactionTx') return jsonOut_({ ok: true, data: updateTransactionTx_(body.data) });
    if (action === 'deleteTransactionTx') return jsonOut_({ ok: true, data: deleteTransactionTx_(body.data.id) });
    if (action === 'batchAddTransactions') return jsonOut_({ ok: true, data: batchAdd_('transactions', (body.data && body.data.rows) || []) });
    if (action === 'settleLedger') return jsonOut_({ ok: true, data: settleLedger_((body.data && body.data.ids) || [], body.data && body.data.settlement) });

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
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
