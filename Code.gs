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
const BACKEND_VERSION_ = '2026-09-15-1';

// 全額代墊的記帳／固定項目統一存成這個主類別名稱，跟前端 ADVANCE_CATEGORY_NAME 保持一致，
// 這樣不管是使用者手動記帳、還是固定項目自動加入，代墊品項在清單上都長得一樣。
const ADVANCE_CATEGORY_NAME_ = '代墊';

const SHEET_NAMES = {
  accounts: 'Accounts',
  liabilities: 'Liabilities',
  interest: 'InterestRecords',
  transactions: 'Transactions',
  recurring: 'RecurringTemplates',
  categories: 'Categories'
};

const SHEET_HEADERS = {
  // favorite  是否為記帳時常用的支付工具（選填，新增於陣列最後面）：
  //           空字串／未設定＝視為常用（相容舊資料，既有帳戶不會突然消失）；
  //           'false' 或 false＝不常用，記帳/模板/共同帳本結算的帳戶選單預設不顯示，
  //           但如果某筆舊紀錄已經用了這個帳戶，編輯那筆紀錄時還是看得到、選得到它。
  // interestCap / normalRate 新增於陣列最後面（原因同上：ensureHeaders_ 只會在
  // 「最後面」補齊缺少的欄位，插在中間會讓既有欄位對不起來，寫入/讀取整個錯位）。
  // interestFreqMonths 同樣新增於陣列最後面：帳戶撥息週期的月數（1=每月、
  // 3=每季、6=每半年、12=每年），空值／未設定時前端會當成 1（每月），相容舊資料。
  accounts: ['id', 'name', 'type', 'balance', 'interestRate', 'note', 'updatedAt', 'favorite', 'interestCap', 'normalRate', 'interestFreqMonths'],
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
  //   toAccountId 轉帳專用（type:'transfer'）：accountId 是轉出帳戶、toAccountId 是轉入帳戶，
  //              跟記帳的轉帳同義，加入記帳時會產生一筆 type:'transfer' 的紀錄（新增於陣列最後面）
  recurring: ['id', 'name', 'type', 'category', 'amount', 'accountId', 'dayOfMonth', 'note', 'active', 'updatedAt', 'subcategory', 'payer', 'splitMode', 'toAccountId'],
  // 類別設定（主類別／子類別）：以前存在前端瀏覽器的 localStorage，跟試算表資料完全分開，
  // 換裝置、清瀏覽器快取都會不見，也沒辦法在試算表上直接看到/管理。改成存在這張表：
  //   type      'expense' | 'income'
  //   mainName  主類別名稱
  //   subName   子類別名稱；空字串代表這一列本身就是「這個主類別」（就算它底下還沒有
  //             任何子類別，也需要至少一列才能讓這個主類別「存在」）
  // 每個子類別各佔一列，同一個主類別會有多列（mainName 相同、subName 不同）。
  categories: ['id', 'type', 'mainName', 'subName', 'updatedAt']
};

// 類別表第一次建立（試算表裡原本沒有 Categories 這張表）時，用這份預設清單
// 灌進去，讓舊使用者原本習慣的預設主類別/子類別在新的儲存方式下維持一樣，
// 跟前端原本的 DEFAULT_CATEGORY_TREE 內容一致。
const DEFAULT_CATEGORY_TREE_ = {
  expense: [
    { name: '餐飲', subs: ['早餐', '午餐', '晚餐', '飲料', '消夜'] },
    { name: '交通', subs: ['大眾運輸', '加油', '停車', '計程車', '保養'] },
    { name: '購物', subs: ['日用品', '服飾', '電器', '其他'] },
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
    if (key === 'categories') seedDefaultCategories_(sheet);
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
// 注意：這裡故意直接用 SHEET_HEADERS[key] 的欄位名稱去讀「現有的表頭那一列」
// 找欄號，而不是呼叫 getColMap_（getColMap_ 內部會呼叫 getSheet_，而
// getSheet_ 建立新表時就是呼叫這個函式，會變成互相呼叫）。這裡只在
// 「剛建立好表頭」或「表頭已經穩定存在」的情況下呼叫，直接讀表頭那一列
// 找欄位名稱在哪一欄即可，一樣不假設固定順序。
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

// 手動一次性修正：把既有表格中已經被 Sheets 自動轉成「日期」型別的
// date/dueDate 欄位，改回純文字（不影響顯示出來的日期內容，只是型別）。
// 用法：Apps Script 編輯器上方選這個函式名稱，按「執行」一次即可，
// 不需要每次讀取都跑，跑一次之後欄位格式就會固定是文字。
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

// 如果表格是舊版本、缺少目前定義的某些欄位（例如舊的 Transactions 表沒有
// payer/splitMode/settled），自動把缺少的欄位標題補到最後面，資料不會受影響。
//
// 重要：這裡改成用「欄位名稱」比對缺少哪些欄位，而不是單純比較欄位數量。
// 原本的寫法是「新定義的欄位數量 > 現有欄位數量，才補上多出來的那幾欄」，
// 這在試算表整份被搬移、或使用者自己在 Sheets 上手動調整過欄位順序時，
// 還是「看起來」欄位數量一樣，但其實順序不同，補欄位的邏輯不會出錯，
// 但後面讀寫如果還假設「第 N 欄一定是某個欄位」就會全部對錯欄位——這正是
// 「試算表搬移過，讀取失敗」的根本原因。真正的修法在下面的 getColMap_()：
// 統一改成每次都用「目前實際的表頭文字」去比對欄位名稱找出真正欄號，
//不再假設任何固定順序。這裡的 ensureHeaders_ 只負責「名稱比對後，
// 真的完全找不到的欄位」才補到最後面，既有欄位不管順序如何都不會被動到。
function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim()) : [];
  const existing = {};
  currentHeaders.forEach(h => { if (h) existing[h] = true; });
  const missing = headers.filter(h => !existing[h]);
  if (missing.length) {
    sheet.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing]);
  }
}

// ---------- 欄位對應（不假設試算表的實際欄位順序） ----------
// 使用者可能自行在 Google Sheet 上拖曳搬移過欄位順序、複製/搬移過整份試算表、
// 或在既有欄位中間插入/刪除過欄位。以前很多地方直接假設「試算表的實際欄位
// 順序」跟上面 SHEET_HEADERS 定義的順序完全一致（例如用
// SHEET_HEADERS.xxx.indexOf('id') + 1 當作實際欄號），一旦順序對不上，
// 讀寫就會全部錯位（讀到別的欄位的值、寫到別的欄位去），表現出來就是
// 「明明有資料卻讀取失敗／資料兜不起來／帳戶餘額對不上」。
//
// 這裡統一改成：每次都讀「試算表目前實際的表頭那一列」文字，用欄位名稱
// 比對找出真正的欄號，不管實際順序是什麼、有沒有被搬過、有沒有被插入過
// 其他欄位，永遠讀寫到正確的欄位。colMapCache_ 只在單次執行（一次
// doGet/doPost 呼叫）內快取，不會跨執行殘留舊的欄位對應。
const colMapCache_ = {};
function getColMap_(key) {
  if (colMapCache_[key]) return colMapCache_[key];
  const sheet = getSheet_(key); // 內部已呼叫 ensureHeaders_，確保欄位標題齊全
  const lastCol = sheet.getLastColumn();
  const headerRow = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const map = {};
  headerRow.forEach((h, i) => {
    const name = String(h || '').trim();
    if (name && map[name] === undefined) map[name] = i + 1; // 1-based 欄號；重複表頭取第一個
  });
  colMapCache_[key] = map;
  return map;
}

// 取得某個欄位名稱目前實際的欄號（1-based）；找不到就丟例外（理論上不會
// 發生，因為 ensureHeaders_ 保證 SHEET_HEADERS 定義的欄位一定存在）。
function colIndex_(key, headerName) {
  const idx = getColMap_(key)[headerName];
  if (!idx) throw new Error('試算表「' + SHEET_NAMES[key] + '」找不到欄位: ' + headerName);
  return idx;
}

// 依實際欄號由小到大排出目前的表頭陣列，例如 ['name','id','balance',...]，
// 用來把「一整列」的值依實際順序對應回欄位名稱（讀取時用）。
function getActualHeaderRow_(key) {
  const map = getColMap_(key);
  const arr = [];
  Object.keys(map).forEach(h => { arr[map[h] - 1] = h; });
  return arr;
}

// 把 data 物件依「試算表目前實際的欄位順序」組成要寫入的一整列陣列，
// 而不是照 SHEET_HEADERS 定義的順序——這樣不管實際欄位被搬到哪裡，
// 每個值都會準確落在正確的欄位。回傳陣列長度＝目前試算表的欄位數，
// SHEET_HEADERS 以外、但試算表上額外存在的欄位（理論上不該有）會保留空白。
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

// 效能重點：只讀 id 那一欄拿來比對就好，不要把整張表所有欄位都讀出來。
// 記帳（Transactions）欄位多、資料量又會一直增加，原本每次「找這筆資料在
// 第幾列」都要整張表（所有欄位）讀一次，資料筆數一多，光是找列號就會變成
// 拖慢新增/編輯/刪除的主因；找列號其實只需要 id 這一欄。
// 注意：id 欄的實際欄號用 colIndex_ 動態找出來（見上面 getColMap_ 的說明），
// 不假設 id 一定在第 1 欄——試算表欄位被搬移過的話，id 也可能不在第 1 欄。
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

// 讀「一整列」目前實際存在的所有欄位，依實際表頭文字對應回欄位名稱
// （不假設欄位順序跟 SHEET_HEADERS 一致）。
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
  data.id = Utilities.getUuid();
  data.updatedAt = new Date().toISOString();
  if (key === 'interest') data.createdAt = new Date().toISOString();
  const row = buildRowArray_(key, data); // 依試算表目前實際的欄位順序組列，不假設固定順序
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  // 記帳（Transactions）新增一列後，順手幫這一列的 category / subcategory
  // 欄位套上「目前 Categories 表」的下拉選單驗證，這樣不管是網站寫入、還是
  // 之後使用者直接在試算表上手動編輯這一列，都能跟 Categories 保持一致。
  if (key === 'transactions') {
    applyTransactionRowValidation_(sheet, sheet.getLastRow());
  }
  return data;
}

// knownRowIndex：如果呼叫端已經知道這筆資料在第幾列（例如
// updateTransactionTx_ 為了先讀出舊資料，已經找過一次列號），就直接傳進來，
// 不用再重新掃一次整張表找列號。
function updateRow_(key, data, knownRowIndex) {
  const sheet = getSheet_(key);
  const rowIndex = knownRowIndex || findRowIndexById_(key, data.id);
  if (rowIndex === -1) throw new Error('找不到這筆資料: ' + data.id);
  data.updatedAt = new Date().toISOString();
  const row = buildRowArray_(key, data); // 依試算表目前實際的欄位順序組列，不假設固定順序
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  if (key === 'transactions') {
    applyTransactionRowValidation_(sheet, rowIndex);
  }
  return data;
}

function deleteRow_(key, id) {
  const sheet = getSheet_(key);
  const rowIndex = findRowIndexById_(key, id);
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

// 帳戶即使餘額調整到剛好 0 元，這一列本身完全不會被刪除或跳過——這裡只是
// 「更新既有那一列的 balance 值」，帳戶本身（連同 id/name/type 等其他欄位）
// 永遠繼續存在於 Accounts 表，跟餘額是不是 0 完全無關。
// 如果 findRowIndexById_ 找不到這個帳戶（rowIndex === -1），代表這個
// accountId 在 Accounts 表裡真的不存在（例如帳戶已被刪除），這種情況下
// 才會回傳 null；只要帳戶還在，不管餘額多少都一定會被正確更新。
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

// ---------- 防止重複記帳：同一天、同類型、同主/子類別、同金額視為「疑似重複」 ----------
// 只擋 expense/income 這種使用者一筆一筆手動輸入、最容易「手滑點兩下」或
// 「網路卡頓重送」的類型；transfer（轉帳）、settlement（結算）本來就可能
// 同一天同金額出現好幾筆合理的紀錄（例如分兩次轉一樣的金額），不擋。
// 金額用四捨五入到分（*100）比較，避免浮點數誤差誤判。
function txDupKey_(data) {
  const date = normalizeCellValue_('date', data.date);
  const amt = Math.round((Number(data.amount) || 0) * 100);
  return [date, data.type, data.category || '', data.subcategory || '', amt].join('|');
}

// 掃描 Transactions 表找出跟 data 完全同 key 的既有紀錄（若有）。
// excludeId：編輯某一筆時，要排除自己這一筆，不要跟自己比對成重複。
function findDuplicateTransaction_(data, excludeId) {
  if (data.type === 'transfer' || data.type === 'settlement') return null;
  const sheet = getSheet_('transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const lastCol = sheet.getLastColumn();
  const headerRow = getActualHeaderRow_('transactions'); // 依實際表頭順序，不假設固定欄號
  const idIdx = headerRow.indexOf('id');
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const targetKey = txDupKey_(data);
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (r[idIdx] === '' || r[idIdx] === null) continue;
    if (excludeId && String(r[idIdx]) === String(excludeId)) continue;
    const obj = {};
    headerRow.forEach((h, hi) => { if (h) obj[h] = normalizeCellValue_(h, r[hi]); });
    if (txDupKey_(obj) === targetKey) return obj;
  }
  return null;
}

// 讀出整張 Transactions 表現有的 key 集合，CSV 批次匯入用來一次比對，
// 不用每一列都重新掃一次整張表（掃一次表、查表都是 O(1)）。
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

// data.force === true 時略過重複檢查，直接新增（前端在跳出「這筆看起來
// 重複了，確定要新增嗎？」的確認視窗、使用者按下確定之後，會帶著
// force:true 再送一次）。回傳 { duplicate: 既有那筆紀錄 } 而不是丟例外，
// 這樣前端可以直接把既有那筆的內容顯示給使用者看，比純文字錯誤訊息更清楚。
function addTransactionTx_(data) {
  if (!data.force) {
    const dup = findDuplicateTransaction_(data);
    if (dup) return { duplicate: dup };
  }
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
  const rowIndex = findRowIndexById_('transactions', data.id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳: ' + data.id);

  if (!data.force) {
    const dup = findDuplicateTransaction_(data, data.id);
    if (dup) return { duplicate: dup };
  }

  const old = readRowObj_('transactions', rowIndex);

  const tx = updateRow_('transactions', data, rowIndex);

  // 轉帳（type:'transfer'）編輯時，新舊兩筆資料都會各自影響「轉出／轉入」
  // 兩個帳戶（見 accumulateBalanceDelta_），這裡把舊資料反向還原、新資料
  // 套用的效果合併計算，不管是改金額、改轉出/轉入帳戶、還是把一般記帳改
  // 成轉帳，相關的帳戶都會正確連動更新餘額。
  const deltas = {};
  accumulateBalanceDelta_(old, -1, deltas);
  accumulateBalanceDelta_(tx, 1, deltas);
  const accountsTouched = applyBalanceDeltas_(deltas);
  return { transaction: tx, accounts: accountsTouched };
}

function deleteTransactionTx_(id) {
  const sheet = getSheet_('transactions');
  const rowIndex = findRowIndexById_('transactions', id);
  if (rowIndex === -1) throw new Error('找不到這筆記帳: ' + id);
  const old = readRowObj_('transactions', rowIndex);
  sheet.deleteRow(rowIndex);
  const accountsTouched = applyBalanceEffect_(old, -1);
  return { id: id, account: accountsTouched[0] || null, accounts: accountsTouched };
}

// 批次刪除（記帳頁面「批次刪除」勾選功能用）：一次 exec 內把多筆記帳依序刪掉，
// 每一筆都走跟單筆刪除一樣的 deleteTransactionTx_（含帳戶餘額還原邏輯）；
// 找不到的 id（例如已經被刪過）直接略過，不中斷其他筆的刪除。
// 回傳這批動作總共影響到的帳戶（同一個帳戶被多筆記帳影響時，只保留最後一次
// 的餘額——因為每一筆都是依序套用完成，最後一次就是所有異動套用完的最終結果）。
function batchDeleteTransactionsTx_(ids) {
  const accountsById = {};
  const deletedIds = [];
  (ids || []).forEach(function (id) {
    try {
      const result = deleteTransactionTx_(id);
      deletedIds.push(id);
      (result.accounts || []).forEach(function (acc) {
        if (acc && acc.id) accountsById[acc.id] = acc;
      });
    } catch (err) {
      // 找不到的（可能已經被刪過）直接跳過，繼續刪下一筆
    }
  });
  return { ids: deletedIds, accounts: Object.keys(accountsById).map(function (k) { return accountsById[k]; }) };
}

// 批次新增（CSV 匯入用）：一次 exec 內用陣列寫入所有列，不逐筆來回。
// key 為 'transactions' 時，會順便擋掉「日期＋類型＋主/子類別＋金額」都
// 跟既有紀錄一樣的重複列（例如同一份 CSV 不小心匯入兩次），以及 CSV 檔案
// 裡本身就重複的列，一律略過不寫入，並回傳 skipped 筆數讓前端顯示提醒。
function batchAdd_(key, rows) {
  const sheet = getSheet_(key);
  const now = new Date().toISOString();
  const isTx = key === 'transactions';
  const existingKeys = isTx ? buildTransactionKeySet_() : null;
  const seenKeys = {};
  let skipped = 0;
  const matrix = [];
  (rows || []).forEach(r => {
    if (isTx && r.type !== 'transfer' && r.type !== 'settlement') {
      const k = txDupKey_(r);
      if (existingKeys[k] || seenKeys[k]) { skipped++; return; }
      seenKeys[k] = true;
    }
    const data = Object.assign({}, r);
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

// 共同帳本結算：把指定的記帳列標記為已結算，一次 exec 內完成；
// 如果有帶 settlementData（使用者選了「轉帳帳戶」），會額外新增一筆
// type:'settlement' 的記帳紀錄，並透過 addTransactionTx_ 走一般記帳的
// 複合流程，自動把這筆金額計入所選帳戶的餘額——之後如果要編輯/刪除這筆
// 結算紀錄，餘額也會透過同一套流程正確地跟著調整/還原。
function settleLedger_(ids, settlementData) {
  const sheet = getSheet_('transactions');
  const idColIndex = colIndex_('transactions', 'id') - 1; // 轉成 0-based，配合下面用 getDataRange 讀出的陣列索引
  const settledColIndex = colIndex_('transactions', 'settled') - 1;
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

// ---------- 類別管理（主類別／子類別），改存在 Categories 表 ----------
// 每個類別（不管是主類別本身、還是某個子類別）各佔一列；同一個主類別會有
// 好幾列（mainName 相同），只是 subName 不同。刪除／改名一律採「整批讀出、
// 在記憶體裡篩選或修改、一次寫回」的作法（跟 settleLedger_ 同一套手法），
// 避免逐列 deleteRow 時列號一直往前位移、一不小心刪錯列或漏刪的問題。
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

// 把目前記憶體裡的 rows 陣列整批寫回 Categories 表，並清掉表格裡原本多出來的舊列。
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

// 把扁平的列資料，組回前端要用的 { expense:[{name,subs:[]}], income:[...] } 樹狀結構。
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
  rebuildAllTransactionValidations();
  return { name: name };
}

function addSubCategory_(type, mainName, subName) {
  subName = String(subName || '').trim();
  if (!subName) throw new Error('名稱不能為空');
  const rows = categorySheetRows_();
  if (!rows.some(r => r.type === type && String(r.mainName) === mainName)) throw new Error('主類別不存在');
  if (rows.some(r => r.type === type && String(r.mainName) === mainName && String(r.subName) === subName)) throw new Error('子類別已存在');
  getSheet_('categories').appendRow([Utilities.getUuid(), type, mainName, subName, new Date().toISOString()]);
  rebuildAllTransactionValidations();
  return { name: subName };
}

function removeMainCategory_(type, name) {
  const rows = categorySheetRows_();
  const kept = rows.filter(r => !(r.type === type && String(r.mainName) === name));
  rewriteCategorySheet_(kept);
  rebuildAllTransactionValidations();
  return { removed: rows.length - kept.length };
}

function removeSubCategory_(type, mainName, subName) {
  const rows = categorySheetRows_();
  const kept = rows.filter(r => !(r.type === type && String(r.mainName) === mainName && String(r.subName) === subName));
  rewriteCategorySheet_(kept);
  rebuildAllTransactionValidations();
  return { removed: rows.length - kept.length };
}

// 重新命名主類別：連同底下所有子類別列的 mainName 一起改，並呼叫
// renameCategoryEverywhere_（定義在下面）把過去已經記過的 Transactions／
// RecurringTemplates 資料也一併更新成新名稱。
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
  rebuildAllTransactionValidations();
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
  rebuildAllTransactionValidations();
  return { changed: changed, transactionsUpdated: txResult.updated };
}

// 把類別改名的效果，同步套用到 Transactions／RecurringTemplates 裡已經記過的舊資料，
// 讓「類別清單」跟「舊紀錄」的名稱保持一致，不會變成清單改了名字、舊資料卻對不起來。
// 做法跟 settleLedger_ 一樣：整張表讀進記憶體、就地修改符合條件的儲存格，最後一次寫回去。
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

// 一次性搬移用：把前端瀏覽器原本存在 localStorage 的自訂類別樹整批寫進這張表，
// 取代目前這個 type 底下的所有列。只有前端偵測到「這個瀏覽器有自訂過類別、
// 但這是第一次接上這個試算表」時才會呼叫一次，之後不會再用到。
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

// ---------- 固定項目（房租/健保費/訂閱費用...）自動加入本月記帳 ----------
// 每次呼叫都是「檢查 + 補上」：對每個啟用中的範本，檢查該月份的 Transactions
// 是否已經有 recurringId = 範本id 的紀錄，沒有的話才新增一筆（同時走
// addTransactionTx_ 複合流程，連帳戶餘額一起更新）。這樣不管使用者是自己按
// 按鈕、或用時間驅動觸發器自動呼叫，都不會重複新增。
function pendingRecurringTemplates_(month) {
  const templates = readAll_('recurring').filter(t => String(t.active) === 'true' || t.active === true);
  const doneIds = recurringDoneIdsForMonth_(month);
  return templates.filter(t => doneIds.indexOf(String(t.id)) === -1);
}

// 只針對「某個月份」找出哪些固定項目已經有對應的記帳紀錄（recurringId）。
// 前端首頁載入時帶的 sheet=all&recentMonths=3 為了效能只回傳「近 3 個月」的
// 精簡記帳資料，用那份資料判斷「本月待加入」理論上沒問題（本月一定落在
// 近 3 個月內），但如果讀取當下剛好遇到背景自動觸發器（autoAddMonthlyRecurring）
// 同時在寫入、或 Apps Script 冷啟動造成的讀取時間差，就可能讀到還沒寫入
// 完成的暫時性狀態，等前端背景補齊完整歷史後才發現其實已經加過，畫面因此
// 「跳」一次。這裡改成：只針對「這個月」直接、獨立地從試算表重新讀一次
// 最新資料（用 readTransactionsSince_ 只讀當月範圍，不必整張表全讀，效能
// 跟精簡讀取一樣快），當作最準確的答案，讓 doGet(sheet=all) 可以直接把這份
// 「已完成清單」一起回傳，前端載入的第一次畫面就已經是正確答案。
function recurringDoneIdsForMonth_(month) {
  const since = month + '-01';
  const txResult = readTransactionsSince_(since);
  const doneIds = {};
  txResult.rows.forEach(t => {
    if ((t.date || '').slice(0, 7) === month && t.recurringId) doneIds[String(t.recurringId)] = true;
  });
  return Object.keys(doneIds);
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
    const isTransfer = tpl.type === 'transfer';
    // 轉帳範本（type:'transfer'）：跟記帳表單的轉帳同一套規則，主類別固定存
    // 「轉帳」、不分主/子類別，也不算共同帳本（payer/splitMode 一律用預設值），
    // accountId 是轉出帳戶、toAccountId 是轉入帳戶。
    const payload = isTransfer ? {
      date: date,
      type: 'transfer',
      category: '轉帳',
      subcategory: '',
      amount: Number(tpl.amount) || 0,
      accountId: tpl.accountId || '',
      toAccountId: tpl.toAccountId || '',
      note: tpl.note || tpl.name,
      payer: 'me',
      splitMode: 'personal',
      settled: false,
      recurringId: tpl.id,
      force: true
    } : {
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
      recurringId: tpl.id,
      force: true
    };
    // force:true：固定項目本來就已經靠 recurringDoneIds（見
    // recurringDoneIds­ForMonth_）確保同一個範本同一個月不會被重複加入，
    // 不需要再套用「日期＋類型＋類別＋金額」的通用重複檢查——不然剛好跟
    // 使用者自己手動記的某一筆撞上（例如金額、類別剛好一樣），固定項目
    // 就會被誤判成重複而悄悄加不進去。
    const result = addTransactionTx_(payload);
    created.push(result.transaction);
    // 轉帳範本一次會影響「轉出＋轉入」兩個帳戶，這裡改用 result.accounts（完整清單）
    // 而不是只取 result.account（單一帳戶），不然轉入那個帳戶的餘額不會回傳給前端。
    (result.accounts || []).forEach(acc => accountsTouched.push(acc));
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

// ============ Transactions 主類別／子類別下拉選單同步 ============
// 目的：讓 Transactions 表的 category（主類別）／subcategory（子類別）欄位，
// 不管是網站寫入、還是使用者直接在 Google Sheets 上手動編輯，永遠只能從
// Categories 表「目前」有的清單裡選，避免打錯字、或選到已經被刪除/改名的
// 舊類別，導致試算表跟網站看到的類別對不起來。
//
// 做法：
// 1. onEdit(e) 簡易觸發器：使用者手動改動 Transactions 的 type 或 category
//    欄位時，即時重新計算「同一列」category / subcategory 的下拉選單內容
//    （category 依 type 決定可選哪些主類別；subcategory 依已選的主類別決定
//    可選哪些子類別，兩層是連動的）。
// 2. addRow_ / updateRow_ 寫入 Transactions 時（不管是網站的「新增記帳」
//    還是「編輯記帳」），也會順手幫那一列套用最新的驗證，讓網站寫入的資料
//    一樣看得到、選得到跟 Categories 一致的下拉選單。
// 3. Categories 表本身有異動（新增/刪除/改名主類別或子類別）時，最後都會
//    呼叫 rebuildAllTransactionValidations()，讓「所有既有列」的下拉選單
//    清單同步更新成最新的 Categories 內容（例如：某個子類別被刪掉後，
//    原本用到它的那些列，下拉選單也要重新反映「現在還剩哪些子類別可選」）。
//
// 注意「setAllowInvalid(true)」：這裡刻意允許儲存格目前的值不在清單裡
// （只會在儲存格右上角顯示一個小提示三角形、不會擋掉編輯、也不會清空原本
// 的文字）。原因是 CSV 匯入、或是舊資料裡本來就可能有跟目前 Categories
// 清單對不起來的文字（見 SHEET_HEADERS.transactions 上面的註解：
// 「CSV 匯入的品項可以是任意文字」），如果改成 setAllowInvalid(false)
// （完全擋死、不准跟清單不符的值存在），舊資料或匯入資料反而會被卡住無法
// 儲存。想要「過往紀錄也強制只能是清單內的值」的話，把下面兩處
// .setAllowInvalid(true) 改成 .setAllowInvalid(false) 即可，但建議先跑一次
// rebuildAllTransactionValidations()、自己在試算表上確認清單裡沒有出現
// 「三角形警告」的舊資料，再切換成完全擋死，避免不小心卡住既有紀錄。
// 效能重點（避免「已超過執行階段時間上限」）：Apps Script 每次執行有時間
// 上限，如果記帳筆數有幾百到上千筆，逐格呼叫 setDataValidation()（每一格
// 都是一次跟 Sheets 後端的來回）很容易把時間吃光。
// 這裡改用 Range.setDataValidations(rules) —— 一次把「一整欄、每一列各自
// 的規則」組成一個陣列，寫入一次搞定：category 欄一次呼叫、subcategory 欄
// 一次呼叫，總共固定 2 次，不會再隨記帳筆數線性增加，資料再多也不會逾時。
// （備註：RangeList 沒有 setDataValidation 方法，只有 setDataValidations()
// 這個「整段連續範圍、規則陣列」的寫法才是 Apps Script 支援的批次寫法。）
// 改成函式（而不是模組載入當下就算好的常數），因為欄號現在要依「試算表
// 目前實際的表頭」動態查出來（見 getColMap_），不能再假設固定等於
// SHEET_HEADERS 定義的順序；模組載入當下也還沒有機會讀到試算表內容。
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

// 對「連續一段列」(startRow ~ startRow+numRows-1) 一次套用資料驗證。
// typeValues / mainNameValues：跟列數一一對應、依序排好的 type / 主類別 陣列，
// 由呼叫端準備好（不在這裡另外讀表，方便重複利用呼叫端已經讀出來的資料）。
// null 規則＝清除該儲存格的資料驗證（type 或主類別是空的/查無對應清單時）。
//
// 重要：這個函式內部把任何例外都吞掉（只記錄到執行紀錄，不會往外丟）。
// 原因：這個函式會被 addRow_ / updateRow_ 掛勾呼叫（新增/編輯記帳的過程中），
// 也會被 addMainCategory_ 等類別管理函式掛勾呼叫（新增/刪除/改名類別的過程
// 中）。如果這裡任何一步出錯又沒接住，會導致「本來應該成功的記帳寫入／
// 類別異動」整個回傳失敗給前端，即使實際上資料已經寫進試算表——網站會
// 顯示錯誤、畫面可能沒更新，統計分析等後續畫面就會看起來「壞掉」或
// 「無法顯示」，但其實只是下拉選單這個附加功能出了問題，不應該連累到
// 核心的記帳/類別功能。想除錯的話，到 Apps Script 編輯器左側「執行項目」
// 找失敗紀錄，或手動執行 rebuildAllTransactionValidations() 看回傳訊息。
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

// ---------- 診斷工具：檢查 Transactions 裡的主類別文字，有哪些對不到 Categories ----------
// 子類別下拉選單「依賴」該列目前的主類別文字要跟 Categories 表裡的某個主類別
// 完全一樣（一字不差），才查得到子類別清單；對不上的話子類別欄就不會有下拉
// （這是正常行為，不是錯誤——但代表那些列的主類別文字跟目前的 Categories
// 清單不一致，通常是手動輸入的錯字、CSV 匯入的舊文字、或類別改名時沒有
// 一起更新造成的）。
// 用法：到 Apps Script 編輯器選這個函式、按「執行」，執行完到「執行項目」
// 或按 Ctrl+Enter 看 Logger 輸出，也可以看函式回傳值。會列出「主類別完全
// 對不上 Categories 清單」的文字，以及各自有幾筆記帳用到，方便你判斷要
// 修 Categories 表還是修 Transactions 裡的文字。
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
    // onEdit 觸發器裡的例外不會顯示給使用者看，這裡刻意吞掉避免打斷正常編輯；
    // 如果要除錯，到 Apps Script 編輯器左側「執行項目」可以看到失敗紀錄。
  }
}

// 幫「單一一列」的 category / subcategory 欄位重新套用資料驗證。
// 給 addRow_ / updateRow_ 這種一次只處理一列的情境使用。
function applyTransactionRowValidation_(sheet, row) {
  try {
    const type = String(sheet.getRange(row, txTypeCol_()).getValue() || '').trim();
    const mainName = String(sheet.getRange(row, txCategoryCol_()).getValue() || '').trim();
    applyValidationRange_(sheet, row, [type], [mainName], getCategoryTreeData_());
  } catch (err) {
    console.error('applyTransactionRowValidation_ 失敗（不影響這筆記帳本身的儲存）: ' + err);
  }
}

// 一次幫「整張 Transactions 表」現有每一列重新套用資料驗證清單。
// 使用時機：
//  1) 第一次要啟用這個功能時，到 Apps Script 編輯器手動執行一次，
//     幫所有既有的記帳列補上下拉選單。
//  2) Categories 表有異動時，程式碼會自動呼叫（見 addMainCategory_ 等函式）。
// 只做 1 次讀取（type/category 連續範圍一起讀）＋固定 2 次
// setDataValidations() 寫入（category 欄一次、subcategory 欄一次），
// 記帳筆數再多也是固定成本，不會逾時。
// 整個函式包在 try/catch 裡、永遠不會往外丟例外：這是「幫下拉選單補資料驗證」
// 的附加功能，就算失敗也絕對不能連累呼叫它的 addMainCategory_ 等類別管理
// 函式，讓類別新增/改名/刪除本身失敗。
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
      const tzNow = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
      const currentMonth = Utilities.formatDate(new Date(), tzNow, 'yyyy-MM');
      // 效能重點：本月固定項目完成清單直接從剛剛已經讀出來的 txResult.rows 篩選，
      // 不要再另外呼叫 recurringDoneIdsForMonth_（那個函式內部會重新完整掃一次
      // Transactions 的 id+date 欄，等於每次「載入首頁」都把同一張表的 id/date
      // 欄整欄讀兩遍，記帳筆數愈多、首次載入愈慢，是原本最大的浪費）。
      // since 一定 ≤ 這個月 1 號（或 since 為 null 代表已經拿到全部資料），
      // 所以 txResult.rows 保證涵蓋這個月全部資料，這裡篩選出來的結果
      // 跟原本呼叫 recurringDoneIdsForMonth_ 完全一樣，只是不用重讀一次表。
      const doneIdsSet_ = {};
      txResult.rows.forEach(function (t) {
        if ((t.date || '').slice(0, 7) === currentMonth && t.recurringId) doneIdsSet_[String(t.recurringId)] = true;
      });
      return jsonOut_({
        ok: true,
        data: {
          accounts: readAll_('accounts'),
          liabilities: readAll_('liabilities'),
          interest: readAll_('interest'),
          transactions: txResult.rows,
          recurring: readAll_('recurring'),
          categories: getCategoryTreeData_(),
          transactionsTruncated: truncated,
          transactionsSince: since,
          recurringDoneIds: Object.keys(doneIdsSet_),
          recurringDoneMonth: currentMonth
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
    if (action === 'batchDeleteTransactionsTx') return jsonOut_({ ok: true, data: batchDeleteTransactionsTx_((body.data && body.data.ids) || []) });
    if (action === 'batchAddTransactions') return jsonOut_({ ok: true, data: batchAdd_('transactions', (body.data && body.data.rows) || []) });
    if (action === 'settleLedger') return jsonOut_({ ok: true, data: settleLedger_((body.data && body.data.ids) || [], body.data && body.data.settlement) });
    if (action === 'runRecurringTemplates') return jsonOut_({ ok: true, data: runRecurringTemplates_(body.data && body.data.month, (body.data && body.data.ids) || null) });

    // 類別管理：新增/改名/刪除主類別、子類別，改名時會一併更新過去的記帳／固定項目資料
    if (action === 'addMainCategory') return jsonOut_({ ok: true, data: addMainCategory_(body.data.type, body.data.name) });
    if (action === 'addSubCategory') return jsonOut_({ ok: true, data: addSubCategory_(body.data.type, body.data.mainName, body.data.subName) });
    if (action === 'removeMainCategory') return jsonOut_({ ok: true, data: removeMainCategory_(body.data.type, body.data.name) });
    if (action === 'removeSubCategory') return jsonOut_({ ok: true, data: removeSubCategory_(body.data.type, body.data.mainName, body.data.subName) });
    if (action === 'renameMainCategory') return jsonOut_({ ok: true, data: renameMainCategory_(body.data.type, body.data.oldName, body.data.newName) });
    if (action === 'renameSubCategory') return jsonOut_({ ok: true, data: renameSubCategory_(body.data.type, body.data.mainName, body.data.oldSub, body.data.newSub) });
    if (action === 'replaceCategoryTree') return jsonOut_({ ok: true, data: replaceCategoryTree_(body.data.type, body.data.tree) });

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

