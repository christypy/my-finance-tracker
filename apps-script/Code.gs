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
 *    網址才會套用最新程式碼。**光按編輯器上方的「儲存」不會生效**，
 *    已經發布出去的網址仍然執行「上一次部署當下」的舊程式碼；如果前端
 *    已經改成新版、但忘記在這裡部署新版本，就會出現「連線失敗／未知的
 *    動作／未知的資料表」這類錯誤。改完程式碼後，記得同時更新下面的
 *    BACKEND_VERSION 字串（例如改成今天的日期），前端「設定」頁會自動
 *    比對前後端版本號，一眼就能看出是不是忘記部署新版本。
 *
 * 效能設計重點：
 * - doGet 支援 sheet=all，一次回傳四張表，前端首次載入只需 1 次 HTTP 請求。
 * - 記帳（新增/編輯/刪除）若有連結帳戶，會在同一次 exec 內同時寫入記帳列與
 *   調整帳戶餘額，前端不需要再額外呼叫一次「更新帳戶」。
 * - CSV 匯入、共同帳本結算都是一次 exec 內用陣列批次寫入，不會一筆一筆來回。
 * - sheet=all 帶 recentMonths 時，只會先讀 Transactions 的「日期」這一欄
 *   （單欄讀取很快）找出符合的資料從第幾列開始，再只讀那個範圍，不會像
 *   之前一樣每次都把整張記帳表全部讀出來才篩選，記帳筆數愈多、效果愈明顯。
 *
 * 連線速度的提醒：Apps Script 網頁應用程式沒有常駐伺服器，每次呼叫都可能
 * 要重新啟動執行環境（尤其是閒置一段時間後的第一次呼叫），這是 Google
 * 平台本身的限制，跟這份程式碼無關，沒辦法完全避免，只能盡量減少每次
 * 讀寫的資料量（見上面的效能設計重點）。
 */

// 後端版本號：每次修改這份 Code.gs、並且重新部署「新版本」時，記得順手
// 更新這個字串（例如改成今天的日期），前端「設定」頁會拿這個值跟前端
// FRONTEND_VERSION 比對，用來提醒「忘記部署新版本」這種最常見的連線失敗原因。
const BACKEND_VERSION_ = '2026-09-13';

// 全額代墊的記帳／固定項目統一存成這個主類別名稱，跟前端 ADVANCE_CATEGORY_NAME 保持一致，
// 這樣不管是使用者手動記帳、還是固定項目自動加入，代墊品項在清單上都長得一樣。
const ADVANCE_CATEGORY_NAME_ = '代墊';

const SHEET_NAMES = {
  accounts: 'Accounts',
  liabilities: 'Liabilities',
  interest: 'InterestRecords',
  transactions: 'Transactions',
  recurring: 'RecurringTemplates'
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
  //   recurringId 若這筆是由「固定項目」自動產生的，記錄來源範本 id，用來判斷某個月是否已經加過
  // 注意：新增欄位一律加在陣列「最後面」，不要插在中間，
  // 這樣舊表格用 ensureHeaders_ 自動補欄位時，既有資料的欄位對應才不會跑掉。
  // subcategory  子類別（選填，配合前端「主類別/子類別」管理，新增於陣列最後面）
  // toAccountId  轉帳專用（type:'transfer'）：accountId 是轉出帳戶、toAccountId 是轉入帳戶，
  //              兩個帳戶的餘額會同時調整，且不計入收入／支出統計。新增於陣列最後面。
  transactions: ['id', 'date', 'type', 'category', 'amount', 'accountId', 'note', 'updatedAt', 'payer', 'splitMode', 'settled', 'recurringId', 'subcategory', 'toAccountId'],
  // 固定項目範本（房租、健保費、訂閱費用...）：
  //   type       'expense' | 'income'
  //   dayOfMonth 每月幾號要繳（1-28，僅供提醒顯示用，不會自動觸發）
  //   active     是否啟用，停用的範本不會出現在「本月待加入」清單
  //   subcategory 子類別（選填，新增於陣列最後面）
  //   payer      共同帳本用：'me'（我）| 'partner'（對方）先付款，跟記帳同義（新增於陣列最後面）
  //   splitMode  共同帳本用：'personal'（個人，不分攤）| 'split'（平分50/50）| 'advance'（全額代墊算對方的）
  //              （新增於陣列最後面）；這筆範本被加入記帳時，payer/splitMode 會原封不動帶到
  //              產生出來的記帳紀錄，等同「保留這筆紀錄、只有日期會變」
  recurring: ['id', 'name', 'type', 'category', 'amount', 'accountId', 'dayOfMonth', 'note', 'active', 'updatedAt', 'subcategory', 'payer', 'splitMode']
};

// 效能：同一次 doGet/doPost 執行內，常常會對同一張表重複呼叫 getSheet_
// （例如 updateTransactionTx_ 跟它內部呼叫的 updateRow_ 都要用到 Transactions
// 表）。快取起來可以省掉重複的 getSheetByName / ensureHeaders_（會讀一次表頭）
// 呼叫；SpreadsheetApp.getActiveSpreadsheet() 也一起快取，理由相同。
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
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES[key]);
    sheet.appendRow(SHEET_HEADERS[key]);
    setDateColumnsAsPlainText_(sheet, key);
  } else {
    ensureHeaders_(sheet, SHEET_HEADERS[key]);
  }
  sheetCache_[key] = sheet;
  return sheet;
}

// 把 date/dueDate 欄位整欄格式設成「純文字」，這樣之後不管是程式寫入
// 還是使用者在試算表上手動輸入 "2025-08-19"，都不會被 Sheets 自動轉成
// 日期型別（也就不會再有雙擊跳出日期選擇器、但前端日曆讀不到的落差）。
// 只有新建立的表會自動套用；既有的表格如果想順便修正，
// 可以到 Apps Script 編輯器手動執行一次 fixExistingDateColumns() 函式。
function setDateColumnsAsPlainText_(sheet, key) {
  const headers = SHEET_HEADERS[key];
  DATE_ONLY_FIELDS_.forEach(col => {
    const idx = headers.indexOf(col);
    if (idx !== -1) {
      sheet.getRange(1, idx + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
    }
  });
}

// 手動一次性修正：把既有表格中已經被 Sheets 自動轉成「日期」型別的
// date/dueDate 欄位，改回純文字（不影響顯示出來的日期內容，只是型別）。
// 用法：Apps Script 編輯器上方選這個函式名稱，按「執行」一次即可，
// 不需要每次讀取都跑，跑一次之後欄位格式就會固定是文字。
function fixExistingDateColumns() {
  ['transactions', 'liabilities'].forEach(key => {
    const sheet = getSheet_(key);
    const headers = SHEET_HEADERS[key];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { setDateColumnsAsPlainText_(sheet, key); return; }
    DATE_ONLY_FIELDS_.forEach(col => {
      const idx = headers.indexOf(col);
      if (idx === -1) return;
      const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
      const values = range.getValues().map(row => [normalizeCellValue_(col, row[0])]);
      range.setNumberFormat('@').setValues(values);
    });
  });
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

// Google 試算表有個很容易踩到的陷阱：即使是程式（setValues）寫入的純文字，
// 只要字串長得像日期（例如 "2025-08-19"），只要那個儲存格格式是「自動」，
// Sheets 就會自動幫你轉成「日期」型別，跟在 UI 上手動輸入日期、雙擊儲存格
// 會跳出日期選擇器是同一件事。updatedAt 因為存的是完整 ISO 時間戳記
// （2026-09-11T08:01:44.374Z），格式不吻合日期格式，才會維持文字。
// 這會導致 getValues() 讀回來的 date/dueDate 不是 "2025-08-19" 字串，
// 而是一個 Date 物件；之後 JSON.stringify 會把它轉成
// "2025-08-18T16:00:00.000Z" 這種帶時區、甚至日期會位移一天的格式，
// 跟前端日曆用字串比對 (t.date === '2025-08-19') 當然對不起來，
// 這就是「試算表明明有資料，日曆卻是空的」的真正原因。
// 這裡統一在讀取時把日期欄位轉回試算表所在時區的 yyyy-MM-dd 純文字，
// 不管儲存格底層是文字還是日期型別，前端拿到的永遠是一致的格式。
const DATE_ONLY_FIELDS_ = ['date', 'dueDate'];
function normalizeCellValue_(header, value) {
  if (value instanceof Date) {
    const tz = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    if (DATE_ONLY_FIELDS_.indexOf(header) !== -1) {
      return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
    }
    // 其他欄位（理論上不該是日期型別，但保險起見）轉成 ISO 字串，避免
    // JSON.stringify 直接輸出 Date 物件造成奇怪格式。
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

// 效能重點：只讀第一欄（id）拿來比對就好，不要把整張表所有欄位都讀出來。
// 記帳（Transactions）欄位多、資料量又會一直增加，原本每次「找這筆資料在
// 第幾列」都要整張表（所有欄位）讀一次，資料筆數一多，光是找列號就會變成
// 拖慢新增/編輯/刪除的主因；找列號其實只需要 id 這一欄。
function findRowIndexById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const target = String(id);
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === target) return i + 2; // 1-based row number
  }
  return -1;
}

function readRowObj_(sheet, rowIndex, headers) {
  const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => (obj[h] = normalizeCellValue_(h, values[i])));
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

// knownRowIndex：如果呼叫端已經知道這筆資料在第幾列（例如
// updateTransactionTx_ 為了先讀出舊資料，已經找過一次列號），就直接傳進來，
// 不用再重新掃一次整張表找列號。
function updateRow_(key, data, knownRowIndex) {
  const sheet = getSheet_(key);
  const headers = SHEET_HEADERS[key];
  const rowIndex = knownRowIndex || findRowIndexById_(sheet, data.id);
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

// type 'settlement'（共同帳本結算轉帳）用 payer 決定方向：
// payer 'partner' = 對方轉錢給你 → 帳戶餘額增加；payer 'me'（或未指定）= 你轉給對方 → 帳戶餘額減少。
// 跟一般 expense/income 走同一套 applyBalanceEffect_/addTransactionTx_/updateTransactionTx_/
// deleteTransactionTx_ 複合流程，這樣之後編輯或刪除這筆結算紀錄時，餘額也會正確地一起還原/調整，
// 不需要另外寫一套專屬的餘額調整邏輯。
function txDelta_(type, amount, payer) {
  const amt = Number(amount) || 0;
  if (type === 'expense') return -amt;
  if (type === 'income') return amt;
  if (type === 'settlement') return payer === 'partner' ? amt : -amt;
  return 0;
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

// 計算一筆記帳對帳戶餘額的影響，累加進 deltas（key 是 accountId，value 是
// 淨變化金額），先不實際讀寫表格：
// - expense/income/settlement：只影響 accountId 一個帳戶
// - transfer：同時影響 accountId（轉出，扣款）與 toAccountId（轉入，入帳）兩個帳戶
// sign = 1 表示套用這筆紀錄的效果；sign = -1 表示反向還原（編輯前/刪除時用）
function accumulateBalanceDelta_(data, sign, deltas) {
  if (data.type === 'transfer') {
    const amt = Number(data.amount) || 0;
    if (data.accountId) deltas[data.accountId] = (deltas[data.accountId] || 0) + (-amt * sign);
    if (data.toAccountId) deltas[data.toAccountId] = (deltas[data.toAccountId] || 0) + (amt * sign);
  } else if (data.accountId) {
    deltas[data.accountId] = (deltas[data.accountId] || 0) + txDelta_(data.type, data.amount, data.payer) * sign;
  }
}

// 把算好的「每個帳戶淨變化多少」實際讀寫進 Accounts 表，一個帳戶只讀寫一次。
function applyBalanceDeltas_(deltas) {
  const touched = [];
  Object.keys(deltas).forEach(accountId => {
    if (!deltas[accountId]) return; // 淨變化剛好是 0，這個帳戶完全不用動
    const acc = adjustAccountBalance_(accountId, deltas[accountId]);
    if (acc) touched.push(acc);
  });
  return touched;
}

// 新增／刪除只有「一筆資料」要套用效果，直接算完就寫入。
function applyBalanceEffect_(data, sign) {
  const deltas = {};
  accumulateBalanceDelta_(data, sign, deltas);
  return applyBalanceDeltas_(deltas);
}

function addTransactionTx_(data) {
  const tx = addRow_('transactions', data);
  const accountsTouched = applyBalanceEffect_(tx, 1);
  return { transaction: tx, account: accountsTouched[0] || null, accounts: accountsTouched };
}

// 編輯記帳時，同一個帳戶常常同時出現在「舊資料要還原」跟「新資料要套用」兩邊
// （例如只是改名稱/備註，帳戶、金額、類型完全沒變）。原本的作法是先把舊資料
// 的效果反向套用（讀一次、寫一次），再把新資料的效果套用一次（又讀一次、寫
// 一次），同一個帳戶等於重複讀寫兩趟。這裡改成把新舊兩筆資料的效果合併算成
// 「每個帳戶淨變化多少」，淨變化是 0 的帳戶（最常見：只改了名稱／備註）就完
// 全不用碰 Accounts 表，一般情況也是每個帳戶只讀寫一次，減少一半的表格存取。
function updateTransactionTx_(data) {
  const sheet = getSheet_('transactions');
  const headers = SHEET_HEADERS.transactions;
  const rowIndex = findRowIndexById_(sheet, data.id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳: ' + data.id);
  const old = readRowObj_(sheet, rowIndex, headers);

  const tx = updateRow_('transactions', data, rowIndex);

  const deltas = {};
  accumulateBalanceDelta_(old, -1, deltas);
  accumulateBalanceDelta_(tx, 1, deltas);
  const accountsTouched = applyBalanceDeltas_(deltas);
  return { transaction: tx, accounts: accountsTouched };
}

function deleteTransactionTx_(id) {
  const sheet = getSheet_('transactions');
  const headers = SHEET_HEADERS.transactions;
  const rowIndex = findRowIndexById_(sheet, id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳: ' + id);
  const old = readRowObj_(sheet, rowIndex, headers);
  sheet.deleteRow(rowIndex);
  const accountsTouched = applyBalanceEffect_(old, -1);
  return { id: id, account: accountsTouched[0] || null, accounts: accountsTouched };
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

// 共同帳本結算：把指定的記帳列標記為已結算，一次 exec 內完成；
// 如果有帶 settlementData（使用者選了「轉帳帳戶」），會額外新增一筆
// type:'settlement' 的記帳紀錄，並透過 addTransactionTx_ 走一般記帳的
// 複合流程，自動把這筆金額計入所選帳戶的餘額——之後如果要編輯/刪除這筆
// 結算紀錄，餘額也會透過同一套流程正確地跟著調整/還原。
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
  let account = null;
  if (settlementData) {
    const result = addTransactionTx_(settlementData);
    settlement = result.transaction;
    account = result.account;
  }
  return { updated: updated, settlement: settlement, account: account };
}

// ---------- 固定項目（房租/健保費/訂閱費用...）自動加入本月記帳 ----------
// 每次呼叫都是「檢查 + 補上」：對每個啟用中的範本，檢查該月份的 Transactions
// 是否已經有 recurringId = 範本id 的紀錄，沒有的話才新增一筆（同時走
// addTransactionTx_ 複合流程，連帳戶餘額一起更新）。這樣不管使用者是自己按
// 按鈕、或用時間驅動觸發器自動呼叫，都不會重複新增。
function pendingRecurringTemplates_(month) {
  const templates = readAll_('recurring').filter(t => String(t.active) === 'true' || t.active === true);
  const monthTx = readAll_('transactions').filter(t => (t.date || '').slice(0, 7) === month);
  const doneIds = {};
  monthTx.forEach(t => { if (t.recurringId) doneIds[String(t.recurringId)] = true; });
  return templates.filter(t => !doneIds[String(t.id)]);
}

function runRecurringTemplates_(month, ids) {
  if (!month) throw new Error('缺少月份參數');
  const pending = pendingRecurringTemplates_(month);
  const idSet = ids && ids.length ? {} : null;
  if (idSet) ids.forEach(id => (idSet[String(id)] = true));
  const toRun = idSet ? pending.filter(t => idSet[String(t.id)]) : pending;

  const created = [];
  const accountsTouched = [];
  toRun.forEach(tpl => {
    const day = Math.min(Math.max(parseInt(tpl.dayOfMonth, 10) || 1, 1), 28);
    const date = month + '-' + String(day).padStart(2, '0');
    // 固定項目也可能是共同帳本項目：payer/splitMode 直接沿用範本設定，
    // 等同「保留這筆紀錄，只有日期會變，其他都不變」，不管是手動按「加入」
    // 還是每月自動觸發都一樣。舊範本沒有設定過 payer/splitMode 時，預設
    // 跟以前一樣是「我先付／個人不分攤」，行為不會改變。
    const payer = tpl.payer === 'partner' ? 'partner' : 'me';
    const splitMode = (tpl.splitMode === 'split' || tpl.splitMode === 'advance') ? tpl.splitMode : 'personal';
    // 全額代墊只有在「我先付」時，才代表整筆其實是對方的花費，才統一存成
    // 「代墊」；「對方先付」代表整筆其實是我的花費，要保留範本原本設定的類別，
    // 跟前端記帳表單／固定項目表單同一套規則（isAdvanceMode_ / isRecAdvanceMode_）。
    const isAdvance = splitMode === 'advance' && payer !== 'partner';
    const result = addTransactionTx_({
      date: date,
      type: tpl.type || 'expense',
      // 全額代墊的品項跟一般記帳表單一樣，主類別統一存成「代墊」
      category: isAdvance ? ADVANCE_CATEGORY_NAME_ : (tpl.category || tpl.name),
      subcategory: isAdvance ? '' : (tpl.subcategory || ''),
      amount: Number(tpl.amount) || 0,
      accountId: tpl.accountId || '',
      note: tpl.note || tpl.name,
      payer: payer,
      splitMode: splitMode,
      settled: false,
      recurringId: tpl.id
    });
    created.push(result.transaction);
    if (result.account) accountsTouched.push(result.account);
  });
  return { created: created, accounts: accountsTouched, skipped: pending.length - toRun.length };
}

// ---------- 固定項目「勾選啟用＝每月一開始自動新增」的自動觸發 ----------
// 前端「固定項目」頁面的手動加入按鈕，需要使用者自己打開網頁才會執行；
// 如果希望勾選「啟用」的固定項目每個月一開始就自動出現在記帳紀錄，不用
// 手動點，需要靠 Google 的時間驅動觸發器，在背景自動呼叫這支函式。
//
// 設定方式（只需要做一次）：
// 1. 在 Apps Script 編輯器上方的函式下拉選單，選擇 setupMonthlyRecurringTrigger
// 2. 按「執行」，第一次會跳出 Google 授權視窗，照著同意即可
// 3. 之後每個月 1 號凌晨，Google 會自動在背景執行 autoAddMonthlyRecurring，
//    把所有啟用中、當月還沒加入的固定項目自動加進記帳紀錄（重複執行也不會
//    重複新增，邏輯跟手動按「一鍵加入」是共用的）
// 如果之後想取消自動新增，執行 removeMonthlyRecurringTrigger 即可。
function autoAddMonthlyRecurring() {
  const tz = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const month = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  return runRecurringTemplates_(month, null);
}

function setupMonthlyRecurringTrigger() {
  removeMonthlyRecurringTrigger();
  ScriptApp.newTrigger('autoAddMonthlyRecurring')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();
  return '已設定：每月 1 號凌晨會自動把啟用中的固定項目加入記帳。';
}

function removeMonthlyRecurringTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoAddMonthlyRecurring') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ---------- 選用：定時「保溫」，減少「一開始的載入」要等很久的狀況 ----------
// Apps Script 網頁應用程式沒有常駐伺服器：閒置一段時間後的第一次呼叫，
// Google 平台要重新啟動執行環境，這一段「冷啟動」時間往往比實際讀寫試算表
// 的時間長很多，是「打開網頁第一次載入特別久」最主要的原因，光靠這份程式碼
// 本身沒辦法完全消除。這裡提供一個可選的做法：设定時間驅動觸發器，每 10
// 分鐘自動呼叫一次極輕量的 keepWarm_()（幾乎不讀寫試算表），讓執行環境
// 比較常保持「熱」的狀態，減少使用者真的打開網頁時遇到冷啟動的機率。
// 效果會因為 Google 平台當下的資源調度而有所不同，不是每次都一定有感，
// 但成本很低（不會多耗用什麼配額），值得開著。
// 設定方式（只需要做一次）：跟 setupMonthlyRecurringTrigger 一樣，在 Apps
// Script 編輯器選 setupKeepWarmTrigger，按「執行」；不想要了就執行
// removeKeepWarmTrigger。
function keepWarm_() {
  // 刻意不讀寫任何試算表資料，只是讓這個 Apps Script 專案的執行環境被叫醒、
  // 保持熱機，所以幾乎不花時間、也不占讀寫配額。
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

// 效能優化：sheet=all 帶 recentMonths 時，以前的作法是 readAll_('transactions')
// 把整張記帳表所有欄位都讀出來，才在記憶體裡篩選「最近 N 個月」，記帳筆數一多
// （幾百上千筆之後）就會愈讀愈慢。這裡改成：
// 1) 先只讀 id + date 這兩欄（通常相鄰，一次 Range 呼叫就能讀完，比讀全部欄位快很多），
//    用來跳過空列（刪除後留下的空白列），並且找出「日期 >= since」最早出現在第幾列；
// 2) 只針對那之後的列，才真的把全部欄位讀出來組成物件。
// 這樣「最近 3 個月」在資料量很大時只需要讀最後一小段，而不是每次都整張表全讀。
// since 為 null（沒有帶 recentMonths，例如背景補齊完整歷史）時，等同讀全部。
function readTransactionsSince_(since) {
  const sheet = getSheet_('transactions');
  const headers = SHEET_HEADERS.transactions;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], fullCount: 0 };

  const idColIndex = headers.indexOf('id') + 1;
  const dateColIndex = headers.indexOf('date') + 1;
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
  const values = sheet.getRange(startRow, 1, numRows, headers.length).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (r[0] === '' || r[0] === null) continue;
    const obj = {};
    headers.forEach((h, hi) => { obj[h] = normalizeCellValue_(h, r[hi]); });
    if (!since || (obj.date || '') >= since) rows.push(obj);
  }
  return { rows: rows, fullCount: fullCount };
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

    const sheetKey = e.parameter.sheet; // accounts | liabilities | interest | transactions | all

    // 一次回傳四張表，前端首次載入只需 1 次 HTTP 請求，減少等待時間。
    // 記帳紀錄（Transactions）累積久了資料量會很大，如果帶了 recentMonths 參數，
    // 就只回傳最近 N 個月的記帳，讓第一次載入的畫面能更快顯示出來；
    // 前端之後會在背景另外用 sheet=transactions 補齊完整歷史紀錄。
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
          interest: readAll_('interest'),
          transactions: txResult.rows,
          recurring: readAll_('recurring'),
          transactionsTruncated: truncated,
          transactionsSince: since
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
    if (action === 'runRecurringTemplates') return jsonOut_({ ok: true, data: runRecurringTemplates_(body.data && body.data.month, (body.data && body.data.ids) || null) });

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
