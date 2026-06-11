// ============================================================
//  Recruit Dashboard — Google Apps Script API
//  貼到 Apps Script 後，部署為 Web App
//  執行身分：我（你的 Google 帳號）
//  存取權限：任何人
// ============================================================

const SHEET_ID = '1R8TKDQlN3VhG7NP6w9aZtD_qIFkk2b2WCp-yhDAvBos'; // Dashboard DB Sheet ID

// 外部缺額表 Sheet ID（唯讀，讀 4 個部門分頁，篩選 Status）
const VACANCY_SHEET_ID = '1hGbY4FyRjRMpGb1NVj59D7epk7eas7dW-FL0PVHXXas';
const VACANCY_TABS     = ['AI技術應用組', '工程部', '行銷運營部', '業務部']; // 逐分頁讀取再合併
const VACANCY_STATUS   = ['錄取', '補缺', '內轉補缺']; // 只取這些 Status
const VAC_OPTIONS_TAB  = '選項清單'; // 缺額表內維護「管道(A欄)/來源(B欄)」下拉選項的分頁

// ── 使用者權限：全部在 Users 分頁設定，前端只在「設定」填一次密碼 ──
// Users 分頁欄位：account（標籤用）| password（明碼）| role（viewer/editor/admin）| name | active（yes/no）
// 每次請求帶 pw，後端即時比對 → 對應角色；無效密碼一律擋下（含讀取）。
const ROLE_RANK = { viewer: 1, editor: 2, admin: 3 };  // 角色階級（數字越大權限越高）

// ── 允許 CORS，讓前端 Dashboard 可以呼叫 ──
function doGet(e) {
  // 若是直接從編輯器點「執行」誤觸，e 會是 undefined，直接回友善提示
  if (!e || !e.parameter) {
    return ContentService
      .createTextOutput('Recruit Dashboard API is alive. 請從前端 Dashboard 呼叫，不是直接開啟此 URL。')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  const action = e.parameter.action;
  // 讀取也需有效密碼（前端主要走 POST；此處為直接打 URL 的防線）
  if (!roleForPw(e.parameter.pw)) {
    return ContentService.createTextOutput(JSON.stringify({ error: '存取密碼無效' })).setMimeType(ContentService.MimeType.JSON);
  }
  let result;
  try {
    if      (action === 'getTasks')    result = getTasks(e.parameter);
    else if (action === 'getOKR')      result = getOKR(e.parameter);
    else if (action === 'getVac')      result = getVacancies();
    else if (action === 'getAll')      result = getAllData();
    else result = { error: 'Unknown action' };
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData) return jsonResponse({ error: '請從前端 Dashboard 發 POST 請求，不要在編輯器直接執行' });
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON' });
  }

  const action = body.action;

  // ── 以密碼解析角色（明碼比對 Users 分頁）──
  const who = roleForPw(body.pw);

  // whoami：回傳目前密碼對應的角色（前端開頁判斷用）
  if (action === 'whoami') {
    return jsonResponse(who ? { success: true, role: who.role, name: who.name } : { error: '存取密碼無效或未設定' });
  }

  // 其餘動作一律需有效密碼（含讀取 getAll）
  if (!who) return jsonResponse({ error: '存取密碼無效，請於右上「設定」填入正確密碼' });
  const rank = ROLE_RANK[who.role] || 0;

  // 寫入類動作需「編輯」以上權限
  const EDITOR_ACTIONS = ['upsertTask','deleteTask','upsertVac','deleteVac','upsertPipeline','deletePipeline','upsertLink','updateOKR'];
  if (EDITOR_ACTIONS.indexOf(action) >= 0 && rank < ROLE_RANK.editor) {
    return jsonResponse({ error: '權限不足（此操作需要「編輯」以上權限）' });
  }

  let result;
  try {
    if      (action === 'getAll')         result = getAllData();
    else if (action === 'upsertTask')     result = upsertTask(body.data);
    else if (action === 'deleteTask')     result = deleteTask(body.data);
    else if (action === 'upsertVac')      result = upsertVac(body.data);
    else if (action === 'deleteVac')      result = deleteVac(body.data);
    else if (action === 'upsertPipeline') result = upsertPipeline(body.data);
    else if (action === 'deletePipeline') result = deletePipeline(body.data);
    else if (action === 'upsertLink')     result = upsertLink(body.data);
    else if (action === 'updateOKR')      result = updateOKR(body.data);
    else result = { error: 'Unknown action' };
  } catch (err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

// ============================================================
//  使用者權限（全部在 Users 分頁設定，明碼密碼）
// ============================================================

// Users 分頁欄位：account | password | role | name | active
function getUsers() {
  return sheetToObjects(SpreadsheetApp.openById(SHEET_ID), 'Users');
}

// 以明碼密碼解析使用者角色；無效或停用 → null（呼叫端據此完全擋下）
function roleForPw(pw) {
  if (!pw) return null;
  pw = pw.toString();
  const u = getUsers().filter(function (x) {
    return (x.password || '').toString() === pw && (x.active || '').toString().toLowerCase() !== 'no';
  })[0];
  if (!u) return null;
  return { role: (u.role || 'viewer').toString().trim(), name: (u.name || u.account || '').toString() };
}

// 初始化：建立 Users 分頁與預設管理者（在編輯器手動執行一次）
function setupUsersSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Users');
  if (!sheet) sheet = ss.insertSheet('Users');
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 5).setValues([['account', 'password', 'role', 'name', 'active']]);
  if (!getUsers().filter(function (u) { return (u.account || '') === 'admin'; })[0]) {
    sheet.appendRow(['admin', 'admin1234', 'admin', '管理者', 'yes']);
  }
  return 'Users 分頁已就緒。預設管理者密碼 admin1234（明碼，直接在 Users 分頁改）。角色填 viewer / editor / admin。';
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  READ — 主要資料庫
// ============================================================

function getAllData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const manualVacs = sheetToObjects(ss, 'Vacancies').map(r => ({...r, _source: 'manual'}));
  return {
    tasks:      sheetToObjects(ss, 'Tasks'),
    okr:        sheetToObjects(ss, 'OKR'),
    mbp:        sheetToObjects(ss, 'MBP'),
    members:    sheetToObjects(ss, 'Members'),
    settings:   settingsToObject(ss),
    pipeline:   sheetToObjects(ss, 'Pipeline'),
    links:      sheetToObjects(ss, 'Links'),
    snapshots:  sheetToObjects(ss, 'HourSnapshots'), // 每日工時快照（週對週/區間比較用）
    vacancies:  [...getExternalVacancies(), ...manualVacs], // 外部缺額表 + 手動新增
    vacFieldOpts: getVacFieldOptions(),                     // 管道/來源下拉選項（與缺額表同步）
  };
}

function getTasks(params) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return sheetToObjects(ss, 'Tasks');
}

function getOKR(params) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const type   = params.type   || 'okr';
  const quarter = params.quarter || '';
  const data   = sheetToObjects(ss, type === 'mbp' ? 'MBP' : 'OKR');
  return quarter ? data.filter(r => r.quarter === quarter) : data;
}

function getVacancies() {
  return getExternalVacancies();
}

function getSettings() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return settingsToObject(ss);
}

// ============================================================
//  外部缺額表串接（唯讀）
//  資料來源：4 個部門分頁（VACANCY_TABS），逐一讀取再合併
//  篩選條件：Status ∈ {錄取, 補缺, 內轉補缺}（VACANCY_STATUS）
//  欄位對應（缺額表表頭 → Dashboard 欄位）：
//    職能      = 職能
//    事業      = 事業
//    close day = close day
//    報到日    = 報到日
//    報到人員  = 報到人員
//    職等      = 職等   ← 表尾那個（report 區，lastIndexOf，非開缺區 col H）
//    職級      = 職級
//    管道      = 管道
//    來源      = 來源
//    部門      = 分頁名稱（tab name）
//    狀態      = Status
// ============================================================

function getExternalVacancies() {
  const results = [];
  try {
    const ss = SpreadsheetApp.openById(VACANCY_SHEET_ID);
    VACANCY_TABS.forEach(function (tabName) {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet) { Logger.log('缺額表找不到分頁：' + tabName); return; }
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const header = data[0];

      const col  = (name) => header.indexOf(name);
      const colL = (name) => header.lastIndexOf(name); // 重複欄位取最後一個
      const C = {
        func:      col('職能'),
        biz:       col('事業'),
        close_day: col('close day'),
        onboard:   col('報到日'),
        reporter:  col('報到人員'),
        level:     colL('職等'),   // 表尾的職等（接近職級那一欄）
        grade:     col('職級'),
        channel:   col('管道'),
        src:       col('來源'),
        status:    col('Status'),
        title:     col('中文職稱'),
        job_no:    col('單號'),
      };

      const get  = (row, idx) => idx >= 0 ? (row[idx] || '').toString().trim() : '';
      const date = (row, idx) => idx >= 0 ? formatDate(row[idx]) : '';

      for (let i = 1; i < data.length; i++) {
        const row    = data[i];
        const status = get(row, C.status).toString().trim();
        if (!VACANCY_STATUS.includes(status)) continue;

        results.push({
          vac_id:      'EXT-' + tabName + '-' + (i + 1),
          dept:        tabName,                       // 部門 = 分頁名稱
          func:        get(row, C.func),
          level:       get(row, C.level),
          grade:       get(row, C.grade),
          biz:         get(row, C.biz),
          close_day:   date(row, C.close_day),
          onboard_day: date(row, C.onboard),
          reporter:    get(row, C.reporter),
          channel:     get(row, C.channel),
          src:         get(row, C.src),
          status:      status,
          title:       get(row, C.title),
          job_no:      get(row, C.job_no).toString().trim(),
          _source:     'external',
        });
      }
    });
    return results;
  } catch (err) {
    // 若外部 Sheet 讀不到，回傳已收集到的部分，不整批中斷
    Logger.log('getExternalVacancies error: ' + err.message);
    return results;
  }
}

// 讀缺額表「選項清單」分頁的下拉選項，讓 dashboard 的下拉與缺額表自動同步。
// 結構：第 1 列為表頭（管道 / 來源），各選項依序列在其下方；保留表單順序、去重、略過空白。
// 這樣只要在「選項清單」改，dashboard 下次載入即自動跟著變，不需再改 code。
function getVacFieldOptions() {
  const FIELD_HEADER = { channel: '管道', src: '來源' };
  const out = { channel: [], src: [] };
  try {
    const ss = SpreadsheetApp.openById(VACANCY_SHEET_ID);
    const sheet = ss.getSheetByName(VAC_OPTIONS_TAB);
    if (!sheet) { Logger.log('缺額表找不到分頁：' + VAC_OPTIONS_TAB); return out; }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return out;
    const header = data[0].map(function (h) { return String(h).trim(); });
    Object.keys(FIELD_HEADER).forEach(function (key) {
      const cIdx = header.indexOf(FIELD_HEADER[key]);
      if (cIdx < 0) return;
      const seen = {}, list = [];
      for (let i = 1; i < data.length; i++) {
        const v = String(data[i][cIdx] || '').trim();
        if (v && !seen[v]) { seen[v] = true; list.push(v); }
      }
      out[key] = list;
    });
  } catch (err) {
    Logger.log('getVacFieldOptions error: ' + err.message);
  }
  return out;
}

// 缺額表覆寫的查找鍵：優先用 job_no；空白則用「中文職稱|報到人員|close_day|onboard_day」組合
// 這樣即便來源缺額表的「職缺編號」欄是空的，仍可用穩定的組合鍵儲存使用者編輯
function _vacOverrideKey(d) {
  const jn = (d.job_no || '').toString().trim();
  if (jn) return jn;
  return '_synth_' + [
    (d.title       || '').toString().trim(),
    (d.reporter    || '').toString().trim(),
    (d.close_day   || '').toString().trim(),
    (d.onboard_day || '').toString().trim(),
  ].join('|');
}

// 從 dashboard 自己的 DB Sheet 讀外部缺額表的覆寫欄位
// 回傳格式：{ [override_key]: { cm, src, status } }
// 覆寫表的第一欄（仍叫 job_no）儲存的是 _vacOverrideKey 的結果
function _readVacOverrides() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('VacOverrides');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {};
    const header = data[0];
    const jIdx  = header.indexOf('job_no');
    const cIdx  = header.indexOf('cm');
    const sIdx  = header.indexOf('src');
    const stIdx = header.indexOf('status');
    if (jIdx < 0) return {};
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][jIdx] || '').trim();
      if (!key) continue;
      map[key] = {
        cm:     cIdx  >= 0 ? String(data[i][cIdx]  || '') : '',
        src:    sIdx  >= 0 ? String(data[i][sIdx]  || '') : '',
        status: stIdx >= 0 ? String(data[i][stIdx] || '') : '',
      };
    }
    return map;
  } catch (err) {
    Logger.log('_readVacOverrides error: ' + err.message);
    return {};
  }
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return val.toString().trim();
}

// ============================================================
//  WRITE — 主要資料庫
// ============================================================

function upsertTask(data) {
  return upsertRow(SHEET_ID, 'Tasks', 'task_id', data);
}

function upsertVac(data) {
  // 外部缺額表本身唯讀（屬於團隊共用 Sheet），使用者在 dashboard 編輯的
  // cm/src/status 寫到本專案 DB 的 VacOverrides。
  // Key 規則：優先用 job_no；若缺則用 _vacOverrideKey 組合鍵
  if (data._source === 'external') {
    const key = data._override_key || _vacOverrideKey(data);
    if (!key || key === '_synth_|||') {
      return { success: false, error: '此職缺缺少可用識別欄位（job_no/中文職稱/報到人員/日期全空），無法儲存' };
    }
    return _upsertVacOverride(data, key);
  }
  return upsertRow(SHEET_ID, 'Vacancies', 'vac_id', data);
}

function _upsertVacOverride(data, key) {
  const HEADER = ['job_no', 'cm', 'src', 'status', 'updated_at'];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('VacOverrides');
  if (!sheet) {
    sheet = ss.insertSheet('VacOverrides');
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
  // 確保表頭齊全（若有人手改過 sheet）
  const lastCol   = Math.max(1, sheet.getLastColumn());
  const curHeader = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (HEADER.some(h => curHeader.indexOf(h) < 0)) {
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }

  const allData = sheet.getDataRange().getValues();
  const header  = allData[0];
  const jIdx    = header.indexOf('job_no');
  const now     = new Date().toISOString();

  const newRow = header.map(h => {
    if (h === 'job_no')     return key;  // 存放 override key（可能是 job_no 或 _synth_...）
    if (h === 'cm')         return data.cm     || '';
    if (h === 'src')        return data.src    || '';
    if (h === 'status')     return data.status || '';
    if (h === 'updated_at') return now;
    return '';
  });

  // 找現有列（key 對應 job_no 欄）
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][jIdx]).trim() === String(key).trim()) {
      sheet.getRange(i + 1, 1, 1, header.length).setValues([newRow]);
      return { success: true, action: 'updated', key: key };
    }
  }
  // 新增列
  sheet.appendRow(newRow);
  return { success: true, action: 'inserted', key: key };
}

function upsertPipeline(data) {
  return upsertRow(SHEET_ID, 'Pipeline', 'pip_id', data);
}

function upsertLink(data) {
  return upsertRow(SHEET_ID, 'Links', 'link_id', data);
}

const OKR_HEADERS = ['quarter','objective','kr_label','kr_text','kr_pct','task_name','task_desc'];
const MBP_HEADERS = ['quarter','member','objective','kr_label','kr_text','kr_pct','task_name','task_desc'];

function ensureHeaders(sheet, expected) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const current = lastCol > 0 && lastRow > 0
    ? sheet.getRange(1, 1, 1, Math.max(lastCol, expected.length)).getValues()[0]
    : [];
  const hasAll = expected.every(h => current.indexOf(h) >= 0);
  if (hasAll) return current;
  // 若表頭不完整：清空後寫入正確 header
  sheet.clear();
  sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
  return expected.slice();
}

function updateOKR(data) {
  const sheetName = data.type === 'mbp' ? 'MBP' : 'OKR';
  const expected  = data.type === 'mbp' ? MBP_HEADERS : OKR_HEADERS;
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  let   sheet  = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  // 確保表頭存在（新 sheet 或欄位不齊時自動建立）
  const header = ensureHeaders(sheet, expected);

  const rows    = data.rows || [];
  const quarter = data.quarter;
  const member  = data.member || null;

  // 取得現有資料，刪除同 quarter（+ 同 member，若是 MBP）的舊列
  const existing = sheet.getDataRange().getValues();
  const qIdx     = header.indexOf('quarter');
  const mIdx     = header.indexOf('member');

  // 從後往前刪，避免 row index 位移
  // 用 String() 比對以避免 Sheets 把 "2025" 讀成數字 2025 造成型別不符
  for (let i = existing.length - 1; i >= 1; i--) {
    const rowQ = qIdx >= 0 ? existing[i][qIdx] : null;
    const rowM = mIdx >= 0 ? existing[i][mIdx] : null;
    const matchQ = String(rowQ) === String(quarter);
    const matchM = !member || String(rowM) === String(member);
    if (matchQ && matchM) sheet.deleteRow(i + 1);
  }

  // 寫入新列
  if (rows.length > 0) {
    const newRows = rows.map(r => header.map(h => r[h] !== undefined ? r[h] : ''));
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, header.length).setValues(newRows);
  }
  return { success: true, written: rows.length };
}

// ============================================================
//  工具函式
// ============================================================

// 純日期欄位：讀出來會轉成 YYYY-MM-DD 字串（給 <input type="date"> 用）
const PURE_DATE_FIELDS = ['start_date','end_date','done_date','close_day','onboard_day','open_date','report_date','snap_date'];

function sheetToObjects(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];
  const data   = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const header = data[0];
  const tz = Session.getScriptTimeZone();
  return data.slice(1).map(row => {
    const obj = {};
    header.forEach((h, i) => {
      const v = row[i];
      if (v instanceof Date) {
        // 日期欄位轉 YYYY-MM-DD；其餘 Date（例如 updated_at）保留 ISO
        if (PURE_DATE_FIELDS.indexOf(h) >= 0) {
          obj[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        } else {
          obj[h] = v.toISOString();
        }
      } else {
        obj[h] = (v !== undefined && v !== null) ? v.toString() : '';
      }
    });
    return obj;
  });
}

function settingsToObject(ss) {
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const obj  = {};
  data.slice(1).forEach(row => {
    if (row[0]) obj[row[0].toString().trim()] = row[1] !== undefined ? row[1].toString() : '';
  });
  return obj;
}

function upsertRow(sheetId, tabName, idField, data) {
  const ss    = SpreadsheetApp.openById(sheetId);
  let   sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);

  // 加入 updated_at
  data['updated_at'] = new Date().toISOString();

  // 取得/建立表頭
  let header;
  if (sheet.getLastRow() === 0) {
    // 空表：用 data 的 key 自動生成 header（idField 第一、updated_at 最後、底線開頭欄位略過）
    const dataKeys = Object.keys(data).filter(k => k !== idField && k !== 'updated_at' && !k.startsWith('_'));
    header = [idField, ...dataKeys, 'updated_at'];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  } else {
    header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // 若 data 帶了新 key 而 header 沒有，於「現有最後一欄之後」真正新增欄位。
    // 注意：必須用 lastColumn+offset 直接寫到全新的空欄，不可用 updated_at 的位置 setValue，
    //       否則會覆蓋掉 updated_at 的標題、把舊資料誤標成新欄位（曾造成 done_date 顯示 updated_at 時間戳）。
    const missing = Object.keys(data).filter(k => !k.startsWith('_') && header.indexOf(k) < 0);
    if (missing.length) {
      const startCol = sheet.getLastColumn();
      missing.forEach((k, offset) => sheet.getRange(1, startCol + 1 + offset).setValue(k));
      header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
  }

  const allData = sheet.getDataRange().getValues();
  const idIdx   = header.indexOf(idField);
  const newRow  = header.map(h => data[h] !== undefined ? data[h] : '');

  if (idIdx >= 0 && data[idField]) {
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idIdx] && allData[i][idIdx].toString() === data[idField].toString()) {
        sheet.getRange(i + 1, 1, 1, header.length).setValues([newRow]);
        return { success: true, action: 'updated' };
      }
    }
  }
  sheet.appendRow(newRow);
  return { success: true, action: 'inserted' };
}

// 刪除單筆（用 idField 對應）
function deleteRow(sheetId, tabName, idField, idValue) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return { success: false, error: tabName + ' sheet 不存在' };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, error: '無資料可刪' };
  const idIdx = data[0].indexOf(idField);
  if (idIdx < 0) return { success: false, error: '缺 ' + idField + ' 欄位' };
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] && data[i][idIdx].toString() === idValue.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: '找不到 ' + idField + '=' + idValue };
}

function deleteVac(data) {
  if (!data || !data.vac_id) return { success: false, error: '缺 vac_id' };
  if (data._source === 'external') return { success: false, error: '外部資料不可刪除（請至團隊缺額表修改）' };
  return deleteRow(SHEET_ID, 'Vacancies', 'vac_id', data.vac_id);
}

function deletePipeline(data) {
  if (!data || !data.pip_id) return { success: false, error: '缺 pip_id' };
  return deleteRow(SHEET_ID, 'Pipeline', 'pip_id', data.pip_id);
}

function deleteTask(data) {
  if (!data || !data.task_id) return { success: false, error: '缺 task_id' };
  return deleteRow(SHEET_ID, 'Tasks', 'task_id', data.task_id);
}

// ============================================================
//  SETUP — 一次性自動補齊 Tasks sheet 必要欄位
//  使用方式：到 Apps Script 編輯器，函式下拉選 setupTasksSchema → ▶ 執行
//  作用：
//    - 確保 Tasks 表有 task_type, team_kr_ref, member_kr_ref, start_date, end_date
//    - 對既有資料的 task_type 預設填「專案任務」（你可手動改為「常態工作」）
//    - 已存在的欄位不會重複加，可重複執行
// ============================================================

function setupTasksSchema() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) { Logger.log('❌ 找不到 Tasks sheet，請先建立此分頁'); return; }

  const lastCol = Math.max(1, sheet.getLastColumn());
  const header  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const NEED    = ['task_type', 'team_kr_ref', 'member_kr_ref', 'start_date', 'end_date'];
  const missing = NEED.filter(c => header.indexOf(c) < 0);

  if (!missing.length) { Logger.log('✓ Tasks 已具備所有必要欄位'); return; }

  Logger.log('開始補齊 Tasks 欄位...');
  Logger.log('─────────────────────────────────────────────');

  // 1. 在最右邊加上缺少的欄位
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  Logger.log('✓ 新增 ' + missing.length + ' 個欄位：' + missing.join(', '));

  // 2. 既有資料的 task_type 預設「專案任務」
  if (missing.indexOf('task_type') >= 0) {
    const lastRow    = sheet.getLastRow();
    if (lastRow > 1) {
      const typeColNum = lastCol + missing.indexOf('task_type') + 1;
      const range      = sheet.getRange(2, typeColNum, lastRow - 1, 1);
      const values     = range.getValues();
      const newValues  = values.map(r => [r[0] || '專案任務']);
      range.setValues(newValues);
      Logger.log('  ↳ 既有 ' + (lastRow - 1) + ' 筆任務的 task_type 預設填「專案任務」');
    }
  }

  Logger.log('─────────────────────────────────────────────');
  Logger.log('✓ 完成。下一步可執行：');
  Logger.log('   migrateKrRefDryRun → 預覽舊 kr_ref 拆成 team/member_kr_ref');
}

// ============================================================
//  MIGRATION — 一次性將舊 kr_ref 拆成 team_kr_ref + member_kr_ref
//  使用方式：
//    1. 在 Apps Script 編輯器上方函式下拉選 migrateKrRefDryRun → ▶ 執行
//       （乾跑模式，只看 log 不寫入；確認結果無誤再執行下一步）
//    2. 選 migrateKrRef → ▶ 執行（實際寫入）
//
//  邏輯：
//    - 對每筆「專案任務」且 kr_ref 有值的列，將 kr_ref 同步寫到
//      team_kr_ref 和 member_kr_ref，格式為 "<TARGET_QUARTER> | <原值>"
//    - TARGET_QUARTER 預設為當前年份；若你的 OKR 用 Q1/Q2/Q3/Q4 命名，
//      請手動修改 TARGET_QUARTER 變數
//    - 略過：常態工作、kr_ref 為空、或已有 team/member_kr_ref 的列
//    - kr_ref 欄位本身不會被刪除（保留作為備份）
//    - 可重複執行（idempotent）：已遷移過的列會被略過
// ============================================================

function migrateKrRef()       { _migrateKrRef(false); }
function migrateKrRefDryRun() { _migrateKrRef(true);  }

function _migrateKrRef(dryRun) {
  // ── 可調整：遷移後加在 KR 前面的 quarter 標籤 ──
  const TARGET_QUARTER = new Date().getFullYear().toString(); // 例如 '2025'
  // 若你的 OKR 用 Q1/Q2/Q3/Q4，請改成：const TARGET_QUARTER = 'Q3';

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) { Logger.log('❌ 找不到 Tasks sheet'); return; }

  // Step 1：實際模式時，先補上不存在的欄位
  if (!dryRun) {
    const initHeader = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    const needCols   = ['team_kr_ref', 'member_kr_ref'].filter(c => initHeader.indexOf(c) < 0);
    if (needCols.length) {
      sheet.getRange(1, initHeader.length + 1, 1, needCols.length).setValues([needCols]);
      Logger.log('✓ 新增欄位：' + needCols.join(', '));
    }
  }

  // Step 2：重新讀取資料（含剛新增的欄位）
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log('⚠ Tasks sheet 沒有資料'); return; }

  const header        = data[0];
  const idxKr         = header.indexOf('kr_ref');
  const idxTeam       = header.indexOf('team_kr_ref');
  const idxMember     = header.indexOf('member_kr_ref');
  const idxType       = header.indexOf('task_type');
  const idxMemberName = header.indexOf('member');
  const idxTaskId     = header.indexOf('task_id');

  if (idxKr < 0)        { Logger.log('⚠ 找不到 kr_ref 欄位，無資料可遷移'); return; }
  if (idxTeam < 0 || idxMember < 0) {
    Logger.log('⚠ 目標欄位不存在（請先用 migrateKrRef 而非 DryRun，或手動加上 team_kr_ref / member_kr_ref 欄位）');
    return;
  }

  Logger.log((dryRun ? '【乾跑模式 — 不會寫入】' : '【實際執行】') + ' 開始掃描 Tasks sheet');
  Logger.log('共 ' + (data.length - 1) + ' 筆任務、quarter prefix = "' + TARGET_QUARTER + '"');
  Logger.log('─────────────────────────────────────────────');

  let migrated = 0, skipped = 0, already = 0;
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const taskId     = (row[idxTaskId]     || '(no id)').toString();
    const memberName = (row[idxMemberName] || '?'      ).toString();
    const oldKr      = (row[idxKr]         || ''       ).toString().trim();
    const taskType   = (row[idxType]       || ''       ).toString().trim();
    const curTeam    = (row[idxTeam]       || ''       ).toString().trim();
    const curMember  = (row[idxMember]     || ''       ).toString().trim();

    if (taskType === '常態工作') {
      Logger.log(taskId + ' (' + memberName + ')  ── 略過（常態工作）');
      skipped++; continue;
    }
    if (!oldKr) {
      Logger.log(taskId + ' (' + memberName + ')  ── 略過（kr_ref 為空）');
      skipped++; continue;
    }
    if (curTeam || curMember) {
      Logger.log(taskId + ' (' + memberName + ')  ── 略過（已遷移過，保留現值）');
      already++; continue;
    }

    const newVal = TARGET_QUARTER + ' | ' + oldKr;
    Logger.log(taskId + ' (' + memberName + ')  ' + oldKr + ' → ' + newVal + '  ✓');
    updates.push({ rowNum: i + 1, val: newVal });
    migrated++;
  }

  if (!dryRun && updates.length) {
    // 逐列寫入（對 < 200 筆任務性能足夠；超過建議改用 batch setValues）
    updates.forEach(u => {
      sheet.getRange(u.rowNum, idxTeam   + 1).setValue(u.val);
      sheet.getRange(u.rowNum, idxMember + 1).setValue(u.val);
    });
  }

  Logger.log('─────────────────────────────────────────────');
  Logger.log(
    (dryRun ? '✓ 乾跑完成（未寫入）' : '✓ 實際執行完成')
    + '：遷移 ' + migrated + ' 筆 / 略過 ' + skipped + ' 筆 / 已遷移 ' + already + ' 筆'
  );
  if (dryRun && migrated > 0) {
    Logger.log('→ 確認無誤後，請執行 migrateKrRef（不含 DryRun）寫入實際資料');
  }
}

// ============================================================
//  REVERT — 萬一 migrateKrRef 結果不滿意，用這個一鍵還原
//  使用方式：
//    1. 先執行 revertKrMigrationDryRun → ▶ 看 log 確認要清的列
//    2. 再執行 revertKrMigration → ▶ 實際清除
//
//  邏輯：
//    - 只清「值符合遷移腳本格式」的列（"<TARGET_QUARTER> | <kr_ref>"）
//    - 使用者「手動填」的 team_kr_ref / member_kr_ref 會被保留
//    - team_kr_ref 和 member_kr_ref 兩欄獨立判斷（只清符合的那邊）
//    - kr_ref 欄位的備份值不會被動到
//    - team_kr_ref / member_kr_ref 欄位本身不會被刪（避免下次遷移又要重建）
// ============================================================

function revertKrMigration()       { _revertKrMigration(false); }
function revertKrMigrationDryRun() { _revertKrMigration(true);  }

function _revertKrMigration(dryRun) {
  const TARGET_QUARTER = new Date().getFullYear().toString();
  // 若你上次遷移時改過 TARGET_QUARTER（例如 'Q3'），這裡也要對應改成同一個值

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) { Logger.log('❌ 找不到 Tasks sheet'); return; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log('⚠ Tasks sheet 沒有資料'); return; }

  const header        = data[0];
  const idxKr         = header.indexOf('kr_ref');
  const idxTeam       = header.indexOf('team_kr_ref');
  const idxMember     = header.indexOf('member_kr_ref');
  const idxTaskId     = header.indexOf('task_id');
  const idxMemberName = header.indexOf('member');

  if (idxTeam < 0 && idxMember < 0) {
    Logger.log('⚠ 找不到 team_kr_ref / member_kr_ref 欄位，無資料需清除');
    return;
  }
  if (idxKr < 0) {
    Logger.log('⚠ 找不到 kr_ref 備份欄，無法判斷哪些是遷移腳本填的（請確認沒被刪掉）');
    return;
  }

  Logger.log((dryRun ? '【乾跑模式 — 不會清除】' : '【實際執行】') + ' 開始掃描 Tasks sheet');
  Logger.log('清除目標：值等於 "' + TARGET_QUARTER + ' | <kr_ref>" 的列');
  Logger.log('─────────────────────────────────────────────');

  let reverted = 0, kept = 0;
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const taskId     = (row[idxTaskId]     || '(no id)').toString();
    const memberName = (row[idxMemberName] || '?'      ).toString();
    const oldKr      = (row[idxKr]         || ''       ).toString().trim();
    const curTeam    = idxTeam   >= 0 ? (row[idxTeam]   || '').toString().trim() : '';
    const curMember  = idxMember >= 0 ? (row[idxMember] || '').toString().trim() : '';

    if (!curTeam && !curMember) continue; // 兩邊都空，不用處理

    const expected         = oldKr ? (TARGET_QUARTER + ' | ' + oldKr) : null;
    const teamIsMigrated   = expected && curTeam   === expected;
    const memberIsMigrated = expected && curMember === expected;

    if (teamIsMigrated || memberIsMigrated) {
      const parts = [];
      if (teamIsMigrated)   parts.push('team_kr_ref="' + curTeam + '"');
      if (memberIsMigrated) parts.push('member_kr_ref="' + curMember + '"');
      Logger.log(taskId + ' (' + memberName + ')  清除：' + parts.join(' + '));
      updates.push({ rowNum: i + 1, clearTeam: teamIsMigrated, clearMember: memberIsMigrated });
      reverted++;
    } else {
      Logger.log(
        taskId + ' (' + memberName + ')  保留（不符遷移格式，可能是手動填的）：'
        + 'team="' + curTeam + '", member="' + curMember + '"'
      );
      kept++;
    }
  }

  if (!dryRun && updates.length) {
    updates.forEach(u => {
      if (u.clearTeam   && idxTeam   >= 0) sheet.getRange(u.rowNum, idxTeam   + 1).setValue('');
      if (u.clearMember && idxMember >= 0) sheet.getRange(u.rowNum, idxMember + 1).setValue('');
    });
  }

  Logger.log('─────────────────────────────────────────────');
  Logger.log(
    (dryRun ? '✓ 乾跑完成（未清除）' : '✓ 實際清除完成')
    + '：清除 ' + reverted + ' 筆 / 保留 ' + kept + ' 筆（手動填的）'
  );
  if (dryRun && reverted > 0) {
    Logger.log('→ 確認無誤後，請執行 revertKrMigration（不含 DryRun）實際清除');
  }
}

// ============================================================
//  UPGRADE — 將舊版 team_kr_ref / member_kr_ref 從 KR 層升級到「子任務指標」層
//  舊格式：「quarter | KR_label」                     例：「2025 | KR1」
//  新格式：「quarter | objective | KR_label | 任務指標」 例：「2025 | 拓展被動求才管道 | KR1 | LinkedIn Sourcing」
//
//  使用方式：
//    1. 先執行 upgradeKrRefFormatDryRun → ▶ 看 log 確認升級結果
//    2. 確認無誤再執行 upgradeKrRefFormat → ▶ 寫入
//
//  邏輯：
//    - 對每筆有舊格式的任務，依「quarter + KR_label」查 OKR/MBP sheet
//    - 用該 KR 的「第一個任務指標」當作預設升級對象
//    - 升級後請使用者到 Dashboard 任務工時頁微調（如需指定其他子任務指標）
// ============================================================

function upgradeKrRefFormat()       { _upgradeKrRefFormat(false); }
function upgradeKrRefFormatDryRun() { _upgradeKrRefFormat(true);  }

function _upgradeKrRefFormat(dryRun) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. 建立 KR → 第一個任務指標的對照
  const teamMap = _buildKRFirstTaskMap(ss, 'OKR', false);
  const mbpMap  = _buildKRFirstTaskMap(ss, 'MBP', true);

  Logger.log('Team OKR 對照建立完成，共 ' + Object.keys(teamMap).length + ' 個 KR');
  Logger.log('MBP 對照建立完成，共 ' + Object.keys(mbpMap).length + ' 個個人 KR');

  // 2. 讀取 Tasks
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) { Logger.log('❌ 找不到 Tasks sheet'); return; }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log('⚠ Tasks sheet 沒有資料'); return; }

  const header        = data[0];
  const idxTeam       = header.indexOf('team_kr_ref');
  const idxMember     = header.indexOf('member_kr_ref');
  const idxMemberName = header.indexOf('member');
  const idxTaskId     = header.indexOf('task_id');

  if (idxTeam < 0 && idxMember < 0) {
    Logger.log('⚠ 找不到 team_kr_ref / member_kr_ref 欄位，無資料需升級');
    return;
  }

  Logger.log((dryRun ? '【乾跑模式 — 不會寫入】' : '【實際執行】') + ' 開始掃描 Tasks sheet');
  Logger.log('─────────────────────────────────────────────');

  let upgraded = 0, skipped = 0, miss = 0;
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const taskId     = (row[idxTaskId]     || '(no id)').toString();
    const memberName = (row[idxMemberName] || ''       ).toString().trim();
    const curTeam    = idxTeam   >= 0 ? (row[idxTeam]   || '').toString().trim() : '';
    const curMember  = idxMember >= 0 ? (row[idxMember] || '').toString().trim() : '';

    const newTeam   = _upgradeOneRef(curTeam,   teamMap, null);
    const newMember = _upgradeOneRef(curMember, mbpMap,  memberName);

    const teamChanged   = newTeam   !== curTeam;
    const memberChanged = newMember !== curMember;

    if (!teamChanged && !memberChanged) {
      const teamFmt   = _refStatus(curTeam);
      const memberFmt = _refStatus(curMember);
      if (teamFmt === 'new' || memberFmt === 'new' || (!curTeam && !curMember)) {
        skipped++;
        continue;
      }
      // 舊格式但找不到對照
      Logger.log(taskId + '  ⚠ 舊格式但找不到對照（請手動處理）：team="' + curTeam + '" member="' + curMember + '"');
      miss++;
      continue;
    }

    const log = [];
    if (teamChanged)   log.push('team:「' + curTeam + '」→「' + newTeam + '」');
    if (memberChanged) log.push('member:「' + curMember + '」→「' + newMember + '」');
    Logger.log(taskId + '  ✓ ' + log.join(' / '));
    updates.push({ rowNum: i + 1, newTeam: teamChanged ? newTeam : null, newMember: memberChanged ? newMember : null });
    upgraded++;
  }

  if (!dryRun && updates.length) {
    updates.forEach(u => {
      if (u.newTeam   !== null && idxTeam   >= 0) sheet.getRange(u.rowNum, idxTeam   + 1).setValue(u.newTeam);
      if (u.newMember !== null && idxMember >= 0) sheet.getRange(u.rowNum, idxMember + 1).setValue(u.newMember);
    });
  }

  Logger.log('─────────────────────────────────────────────');
  Logger.log(
    (dryRun ? '✓ 乾跑完成（未寫入）' : '✓ 實際升級完成')
    + '：升級 ' + upgraded + ' 筆 / 略過 ' + skipped + ' 筆（已是新格式或空）/ 找不到對照 ' + miss + ' 筆'
  );
  if (dryRun && upgraded > 0) {
    Logger.log('→ 確認無誤後，請執行 upgradeKrRefFormat（不含 DryRun）寫入');
  }
}

// 內部工具：建立 KR → 第一個任務指標 的對照表
// 回傳 key 格式：
//   OKR: 「quarter|KR_label」
//   MBP: 「quarter|member|KR_label」
function _buildKRFirstTaskMap(ss, sheetName, withMember) {
  const map = {};
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;

  const header = data[0];
  const qIdx  = header.indexOf('quarter');
  const oIdx  = header.indexOf('objective');
  const klIdx = header.indexOf('kr_label');
  const tnIdx = header.indexOf('task_name');
  const mIdx  = withMember ? header.indexOf('member') : -1;

  if (qIdx < 0 || klIdx < 0 || tnIdx < 0) return map;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const q  = (row[qIdx]  || '').toString().trim();
    const o  = (row[oIdx]  || '').toString().trim();
    const kl = (row[klIdx] || '').toString().trim();
    const tn = (row[tnIdx] || '').toString().trim();
    const m  = mIdx >= 0 ? (row[mIdx] || '').toString().trim() : '';

    if (!q || !kl || !tn) continue;

    const key = withMember ? (q + '|' + m + '|' + kl) : (q + '|' + kl);
    if (!map[key]) map[key] = { q: q, o: o, kl: kl, tn: tn };
  }
  return map;
}

// 內部工具：升級單一 ref 字串
function _upgradeOneRef(ref, map, memberName) {
  if (!ref) return ref;
  const parts = ref.split('|').map(function(p){ return p.trim(); });
  if (parts.length >= 4) return ref;     // 已是新格式
  if (parts.length !== 2) return ref;    // 不認識的格式
  const q = parts[0], kl = parts[1];
  const key = memberName ? (q + '|' + memberName + '|' + kl) : (q + '|' + kl);
  const entry = map[key];
  if (!entry) return ref;
  return q + ' | ' + entry.o + ' | ' + kl + ' | ' + entry.tn;
}

function _refStatus(ref) {
  if (!ref) return 'empty';
  const parts = ref.split('|');
  return parts.length >= 4 ? 'new' : 'old';
}

// ============================================================
//  每日工時快照（週對週 / 任意區間比較的基礎）
//
//  原理：
//    - 每個任務的 actual_hours 是「一路往上疊加的累計總工時」（不歸零）
//    - 本函式每天 23:00 自動把「當下每個任務的累計 actual_hours」拍一張照，
//      連同當天日期存進 HourSnapshots 分頁
//    - 前端要算「某區間做了多少」時 = 區間結束端累計 − 區間開始前一天累計
//    - 常態工作（無起訖日）原本無法切區間，靠這套快照就能精準切出每週工時
//
//  HourSnapshots 欄位：snap_date | task_id | task_name | member | task_type | actual_hours
//  （task_name/member/task_type 只是方便人工查 sheet，前端計算只用 task_id + actual_hours）
//
//  同一天重複執行 → 先清掉當天舊快照再寫入，結果冪等（永遠是當天最新值）
// ============================================================

const SNAPSHOT_SHEET   = 'HourSnapshots';
const SNAPSHOT_HEADERS = ['snap_date', 'task_id', 'task_name', 'member', 'task_type', 'actual_hours'];

function snapshotTasks() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const tasks = sheetToObjects(ss, 'Tasks');
  if (!tasks.length) { Logger.log('⚠ Tasks 無資料，略過快照'); return; }

  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // 確保快照分頁與表頭存在
  let sheet = ss.getSheetByName(SNAPSHOT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SNAPSHOT_SHEET);
    sheet.getRange(1, 1, 1, SNAPSHOT_HEADERS.length).setValues([SNAPSHOT_HEADERS]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SNAPSHOT_HEADERS.length).setValues([SNAPSHOT_HEADERS]);
  }

  // 冪等：先刪掉「今天」已存在的快照列（從後往前刪避免位移）
  const existing = sheet.getDataRange().getValues();
  const header   = existing[0];
  const dIdx     = header.indexOf('snap_date');
  if (dIdx >= 0) {
    for (let i = existing.length - 1; i >= 1; i--) {
      const v = existing[i][dIdx];
      const d = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').substring(0, 10);
      if (d === today) sheet.deleteRow(i + 1);
    }
  }

  // 一次寫入今天所有任務的累計工時
  const rows = tasks
    .filter(t => t.task_id)
    .map(t => [
      today,
      t.task_id,
      t.task_name || '',
      t.member || '',
      t.task_type || '',
      (+t.actual_hours || 0),
    ]);

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SNAPSHOT_HEADERS.length).setValues(rows);
  }
  Logger.log('✓ ' + today + ' 已快照 ' + rows.length + ' 筆任務工時');
  return '✓ ' + today + ' 已快照 ' + rows.length + ' 筆任務工時';
}

// 一次性：建立「每天晚上 23:00 自動快照」的時間觸發器
// 使用方式：到 Apps Script 編輯器，函式下拉選 setupSnapshotTrigger → ▶ 執行（只需執行一次）
// 重複執行不會建立多個觸發器（會先清掉舊的同名觸發器）
function setupSnapshotTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'snapshotTasks') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('snapshotTasks')
    .timeBased()
    .everyDays(1)
    .atHour(23)        // 每天 23:00（依 Apps Script 專案時區，請確認設為 GMT+8 台北）
    .create();
  const tz = Session.getScriptTimeZone();
  Logger.log('✓ 已建立每日 23:00 自動快照觸發器（時區：' + tz + '）');
  Logger.log('  若時區不是台北，請到「專案設定 → 時區」改為 (GMT+08:00) Taipei 後重新執行本函式');
  return '✓ 已建立每日 23:00 自動快照觸發器（時區：' + tz + '）';
}
