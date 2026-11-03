// ========================
// LINE & Google Sheets 設定
// ========================
const ACCESS_TOKEN = SETTINGS_property_('ACCESS_TOKEN');
const LINE_ENDPOINT_REPLY = 'https://api.line.me/v2/bot/message/reply';
const LINE_ENDPOINT_PUSH = 'https://api.line.me/v2/bot/message/push';

const SHEET_ID = SETTINGS_property_('SHEET_ID');
// シートへのアクセスは必要な処理からだけ行う。onOpen/onEditでは開かない。
let SETTINGS_spreadsheetCache_;
const FUND_LINE_TO = SETTINGS_property_('FUND_LINE_TO');

// ========================
// 設定・定数
// ========================
const CONFIG = {
  HIMOKU_RANGE: 'A2:A31',
  INCOME_LABEL: '収入'
};

// 接続設定はスクリプトプロパティから取得する。空文字は任意設定として維持する。
function SETTINGS_property_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value === null) throw new Error('スクリプトプロパティが未登録です：' + key);
  return value;
}

function SETTINGS_spreadsheet_() {
  if (!SETTINGS_spreadsheetCache_) {
    SETTINGS_spreadsheetCache_ = SpreadsheetApp.openById(SHEET_ID);
  }
  return SETTINGS_spreadsheetCache_;
}

function SETTINGS_mainSheet_() {
  return SETTINGS_spreadsheet_().getSheetByName(new Date().getFullYear().toString());
}

function SETTINGS_detailSheet_() {
  return SETTINGS_spreadsheet_().getSheetByName('内訳');
}

// ========================
// メイン処理
// ========================
function doPost(e) {
  const event = JSON.parse(e.postData.contents).events[0];
  if (!event?.replyToken || !event.message?.text) return;

  const userMessage = event.message.text;
  const messageLines = userMessage.split(/\r\n|\n/);

  const validation = isValid(messageLines);
  const replyText = validation === 'OK'
    ? processEntry(messageLines)
    : validation;

  sendReply(event.replyToken, replyText);
}

// ========================
// 業務ロジック
// ========================
function processEntry(data) {
  const [dateStr, category, amountRaw, memo = ''] = data;
  const amount = cleanAmount(amountRaw);

  if (category === CONFIG.INCOME_LABEL) {
    return saveIncome(dateStr, amount, amountRaw);
  }

  return saveExpense(dateStr, category, amount, amountRaw, memo);
}

// 支出の保存
function saveExpense(dateStr, category, amount, amountRaw, memo) {
  const targetRow = SETTINGS_detailSheet_().getLastRow() + 1;
  const timeZone = SETTINGS_detailSheet_().getParent().getSpreadsheetTimeZone();
  const dateObj = DETAIL_balanceDate_(parseDate(dateStr), category, timeZone);

  SETTINGS_detailSheet_().getRange(targetRow, 1, 1, 4).setValues([[
    dateObj,
    category,
    Number(amount),
    memo
  ]]);

  const savedDate = category === '月初残高' || category === '月末残高'
    ? Utilities.formatDate(dateObj, timeZone, 'yyyy/MM/dd') : dateStr;
  return `${savedDate}に${category}${memo ? `(${memo})` : ''}として${amountRaw}を登録しました。`;
}

// 収入の保存
function saveIncome(dateStr, amount, amountRaw) {
  const dateObj = parseDate(dateStr);
  const year = dateObj.getFullYear().toString();
  const month = dateObj.getMonth() + 1;

  const sheet = SETTINGS_spreadsheet_().getSheetByName(year);
  if (!sheet) return `${year}年のシートが見つかりません`;

  const categories = sheet.getRange('A1:A100').getValues().flat();
  const rowIndex = categories.indexOf(CONFIG.INCOME_LABEL);
  if (rowIndex === -1) return '収入項目がシートにありません';

  const colIndex = month + 1;
  const currentVal = Number(sheet.getRange(rowIndex + 1, colIndex).getValue()) || 0;

  sheet.getRange(rowIndex + 1, colIndex).setValue(currentVal + Number(amount));

  return `収入に${amountRaw}を登録しました。`;
}

// ========================
// ユーティリティ
// ========================
function isValid(data) {
  const [dateStr, category, amountRaw] = data;
  if (!dateStr || !category || !amountRaw) return '入力が足りません';

  try {
    parseDate(dateStr);
  } catch (error) {
    return '日付を正しく入力してください';
  }

  if (category !== CONFIG.INCOME_LABEL) {
    const validCategories = SETTINGS_mainSheet_().getRange(CONFIG.HIMOKU_RANGE).getValues().flat();
    if (!validCategories.includes(category)) return `費目に存在しません：${category}`;
  }

  if (isNaN(Number(cleanAmount(amountRaw)))) return '金額を正しく入力してください';

  return 'OK';
}

function cleanAmount(raw) {
  return String(raw || '').replace(/[円,]/g, '');
}

function parseDate(dateStr) {
  const s = String(dateStr || '').trim();

  // YYYY年M月 → 当日の日付を使う
  const monthOnly = s.match(/^(\d{4})年\s*(\d{1,2})月$/);
  if (monthOnly) {
    const year = Number(monthOnly[1]);
    const month = Number(monthOnly[2]);
    const day = new Date().getDate();

    if (month < 1 || month > 12) {
      throw new Error('日付解析失敗: ' + s);
    }

    const d = new Date(year, month - 1, Math.min(day, new Date(year, month, 0).getDate()));
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // YYYY/MM/DD or YYYY-MM-DD
  const dateMatch = s.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (!dateMatch) {
    throw new Error('日付解析失敗: ' + s);
  }

  const year = Number(dateMatch[1]);
  const month = Number(dateMatch[2]);
  const day = Number(dateMatch[3]);

  if (month < 1 || month > 12 || day < 1 || day > 31) {
    throw new Error('日付解析失敗: ' + s);
  }

  const d = new Date(year, month - 1, day);
  d.setHours(0, 0, 0, 0);

  if (
    d.getFullYear() !== year ||
    d.getMonth() !== month - 1 ||
    d.getDate() !== day
  ) {
    throw new Error('日付解析失敗: ' + s);
  }

  return d;
}

// 同じスプレッドシートに紐づくGASでは、開く・編集する際に自動実行。
function onOpen(e) {
  DETAIL_markActivity_(e);
}

function onEdit(e) {
  DETAIL_markActivity_(e);
}

function DETAIL_markActivity_(e) {
  if (!e || !e.source) return;
  PropertiesService.getScriptProperties().setProperty(
    'DETAIL_LAST_ACTIVITY_' + e.source.getId(), String(Date.now())
  );
}

function DETAIL_recentActivity_() {
  const value = PropertiesService.getScriptProperties().getProperty(
    'DETAIL_LAST_ACTIVITY_' + SETTINGS_detailSheet_().getParent().getId()
  );
  const last = Number(value);
  return !!value && Number.isFinite(last) && Date.now() - last < 3 * 60 * 1000;
}

function sortAndFilterSheet() {
  if (DETAIL_recentActivity_()) {
    console.log('最後に開く・編集してから3分以内なのでソートをスキップしました。');
    return;
  }
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5 * 1000)) return;

  let updateVisibility = false;
  try {
    if (DETAIL_recentActivity_()) return;
    const sheet = SETTINGS_detailSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const lastColumn = sheet.getLastColumn();

    // 非表示の過去日付も含め、同じ日付・費目の明細を隣接させる。
    const startRow = 2;
    const rowCount = lastRow - startRow + 1;
    updateVisibility = true;

    // 手入力済みの残高も、指定されている月の月初・月末に補正する。
    const timeZone = sheet.getParent().getSpreadsheetTimeZone();
    const beforeSort = sheet.getRange(startRow, 1, rowCount, 2).getValues();
    DETAIL_correctBalanceDates_(sheet, startRow, beforeSort, timeZone);

    // 月初残高 → 日付・費目順の明細 → 月末残高。
    // 一時列で行全体を並べ替え、金額・メモ・数式・書式も一緒に移動する。
    DETAIL_sortEntries_(sheet, startRow, lastColumn, beforeSort, timeZone);

  } finally {
    try {
      // 並べ替えの後にグループと折りたたみ状態を確定する。
      // 途中でエラーになっても表示状態の更新を省略しない。
      if (updateVisibility) {
        DETAIL_refreshVisibility_(SETTINGS_detailSheet_());
      }
    } finally {
      lock.releaseLock();
    }
  }
}

function DETAIL_correctBalanceDates_(sheet, startRow, entries, timeZone) {
  entries.forEach(([dateValue, category], i) => {
    const corrected = DETAIL_balanceDate_(dateValue, category, timeZone);
    if (corrected instanceof Date && corrected.getTime() !== dateValue.getTime()) {
      sheet.getRange(startRow + i, 1).setValue(corrected);
    }
  });
}

function DETAIL_sortEntries_(sheet, startRow, lastColumn, entries, timeZone) {
  const rowCount = entries.length;
  const sortColumn = lastColumn + 1;
  sheet.insertColumnAfter(lastColumn);
  try {
    const sortKeys = entries.map(([dateValue, category]) =>
      [DETAIL_entrySortKey_(dateValue, category, timeZone)]);
    sheet.getRange(startRow, sortColumn, rowCount, 1).setValues(sortKeys);
    sheet.getRange(startRow, 1, rowCount, sortColumn).sort([
      { column: sortColumn, ascending: true },
      { column: 2, ascending: true }
    ]);
  } finally {
    sheet.deleteColumn(sortColumn);
  }
}

function DETAIL_refreshVisibility_(sheet) {
  const count = sheet.getLastRow() - 1;
  if (count > 0) {
    const entries = sheet.getRange(2, 1, count, 2).getValues();
    sheet.showRows(2, count);
    DETAIL_groupEntries_(sheet, 2, entries, new Date());
  }
}

function DETAIL_entrySortKey_(dateValue, category, timeZone) {
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime())) return '9999-99';
  const month = Utilities.formatDate(dateValue, timeZone, 'yyyy-MM');
  const priority = category === '月初残高' ? '0' : category === '月末残高' ? '2' : '1';
  // 残高は入力時の日付にかかわらず、その月の最初・最後に置く。
  const day = priority === '1' ? Utilities.formatDate(dateValue, timeZone, 'dd') : '00';
  return month + '-' + priority + '-' + day;
}

function DETAIL_balanceDate_(dateValue, category, timeZone) {
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime()) ||
      (category !== '月初残高' && category !== '月末残高')) return dateValue;
  const month = Utilities.formatDate(dateValue, timeZone, 'yyyy-MM');
  const [year, monthNumber] = month.split('-').map(Number);
  const day = category === '月初残高' ? 1 : new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
  return Utilities.parseDate(month + '-' + String(day).padStart(2, '0'), timeZone, 'yyyy-MM-dd');
}

// 当月以前をグループ化するが、当月の月初・月末残高は表示に残す。
function DETAIL_groupPastEntries_(sheet, startRow, entries, now) {
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const currentMonth = Utilities.formatDate(now, timeZone, 'yyyy-MM');
  let groupStart = -1;
  for (let i = 0; i <= entries.length; i++) {
    const [dateValue, category] = i < entries.length ? entries[i] : [null, ''];
    const month = dateValue instanceof Date && !isNaN(dateValue.getTime())
      ? Utilities.formatDate(dateValue, timeZone, 'yyyy-MM')
      : '';
    const keepBalance = month === currentMonth &&
      (category === '月初残高' || category === '月末残高');
    const group = month !== '' && month <= currentMonth && !keepBalance;
    if (group && groupStart === -1) groupStart = i;
    if (!group && groupStart !== -1) {
      const row = startRow + groupStart;
      sheet.getRange(row, 1, i - groupStart, 1).shiftRowGroupDepth(1);
      sheet.getRowGroup(row, 1).collapse();
      groupStart = -1;
    }
  }
}

// 当月以前と再来月以降を別々のグループにする。翌月と当月残高は表示に残す。
function DETAIL_groupEntries_(sheet, startRow, entries, now) {
  if (!entries.length) return;
  sheet.getRange(startRow, 1, entries.length, 1).shiftRowGroupDepth(-8);
  sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  DETAIL_groupPastEntries_(sheet, startRow, entries, now || new Date());
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const currentMonth = Utilities.formatDate(now || new Date(), timeZone, 'yyyy-MM');
  const [year, month] = currentMonth.split('-').map(Number);
  const nextMonth = new Date(Date.UTC(year, month, 1)).toISOString().slice(0, 7);
  const isFuture = ([value]) => value instanceof Date && !isNaN(value.getTime()) &&
    Utilities.formatDate(value, timeZone, 'yyyy-MM') > nextMonth;
  const first = entries.findIndex(isFuture);
  if (first === -1) return;
  let end = first + 1;
  while (end < entries.length && isFuture(entries[end])) end++;
  // BEFOREの操作ボタンはグループ直前の行に付くため、先頭行を除く。
  const count = end - first - 1;
  if (count > 0) {
    const row = startRow + first + 1;
    sheet.getRange(row, 1, count, 1).shiftRowGroupDepth(1);
    sheet.getRowGroup(row, 1).collapse();
  }
}

// ========================
// API・通信系
// ========================
function sendReply(token, text) {
  UrlFetchApp.fetch(LINE_ENDPOINT_REPLY, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${ACCESS_TOKEN}`,
    },
    payload: JSON.stringify({
      replyToken: token,
      messages: [{ type: 'text', text }]
    }),
  });
}

function doGet(e) {
  const list = getHimokuList(e.parameter.year);

  return ContentService
    .createTextOutput(JSON.stringify(list))
    .setMimeType(ContentService.MimeType.JSON);
}

function getHimokuList(year) {
  if (!year) return [];

  const sheet = SETTINGS_spreadsheet_().getSheetByName(year);
  if (!sheet) return [];

  const data = sheet.getRange('A1:A50').getValues().flat();
  const start = data.indexOf('娯楽');
  const end = data.indexOf('口座(生活費)');

  if (start === -1 || end === -1 || end <= start) return [];
  return data.slice(start, end);
}

// ============================================================
// 投資信託管理：2026年10月から積立を推計して加算
// FUNDSの口数・元本は9月積立までを含む初期残高。毎月書き換えない。
// 基準価額はGASで3銘柄取得成功済み。口座の買付成功は確認できない。
// 追加口数は四捨五入による推計。実績との一致は保証しない。
// 初期設定は setupFunds、自動更新は既存の recordFundNightly トリガーを使用。
// 資産推移に1日1行の合計のみを記録。口座区分ごとに四捨五入して合計。
// 旧「投信日次記録」の合計は初回に移行し、元シートは非表示で保存する。
// ============================================================
const FUNDS = [
  {name:'楽天・プラス・S&P500',code:'JP90C000Q2U6',units:186895,principal:316074,monthlyAmount:40000},
  {name:'eMAXIS Slim 米国株式（S&P500）',code:'JP90C000GKC6',units:109281,principal:402500,monthlyAmount:0},
  {name:'iFreeNEXT FANG+',code:'JP90C000FZD4',units:13769,principal:119655,monthlyAmount:10000}
];
// 9月積立までを含む初期口数を、楽天証券の実際の口座区分ごとに設定。
// 例: [{account:'旧つみたてNISA',units:実際の口数}, {account:'新NISAつみたて投資枠',units:実際の口数}]
// 未入力で止めることで、推測の内訳による誤記録を防ぐ。
const FUND_ACCOUNT_POSITIONS = {
  JP90C000Q2U6: [{account:'つみたて投資枠',units:185084},{account:'成長投資枠',units:1811}],
  JP90C000GKC6: [{account:'つみたて投資枠',units:106416},{account:'成長投資枠',units:2865}],
  JP90C000FZD4: [{account:'つみたて投資枠',units:12618},{account:'成長投資枠',units:1151}]
};
// 現在の毎月積立を加算する口座区分名。上のaccountと同じ名前を設定。
const FUND_MONTHLY_ACCOUNT = {
  JP90C000Q2U6: 'つみたて投資枠',
  JP90C000FZD4: 'つみたて投資枠'
};
// 添付コードの設定を維持。自動収集の利用許諾は別途確認が必要。
const FUND_ENABLE_RAKUTEN_FETCH = true;
const FUND_SPREADSHEET_ID = SETTINGS_property_('FUND_SPREADSHEET_ID');
const FUND_TZ = 'Asia/Tokyo';
const FUND_START_MONTH = '2026-10';
const FUND_BUY_DAY = 12;
const FUND_HISTORY_SHEET = '資産推移';
const FUND_LEGACY_SHEET = '投信日次記録';
const FUND_INPUT_SHEET = '基準価額入力';
const FUND_VERIFY_SHEET = '取得検証';
const FUND_PLAN_SHEET = '積立履歴';
const FUND_HEADERS = ['基準日','総評価額（円）','元本合計（円）','含み益（円）',
  '含み益率','前日比（円）','前日比率','積立入金（円）','運用損益（円）',
  '運用損益率','比較基準日','取得日時','備考'];
const FUND_PLAN_HEADERS = ['対象月','銘柄','銘柄コード','積立額（円）','買付予定日',
  '約定予定日','受渡予定日','保有反映日','約定基準価額','追加口数','状態','更新日時',
  '実際の約定日（任意）','実際の買付口数（任意）','買付取消（TRUEで取消）'];

function setupFunds() {
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    FUND_validateFunds_();const ss=FUND_spreadsheet_();ss.setSpreadsheetTimeZone(FUND_TZ);
    FUND_historySheet_(ss);
    FUND_ensureSheet_(ss,FUND_INPUT_SHEET,['基準日','銘柄コード','基準価額（円／1万口）']);
    FUND_ensureSheet_(ss,FUND_PLAN_SHEET,FUND_PLAN_HEADERS);
    FUND_migrateSummary_(ss);
    [FUND_LEGACY_SHEET,FUND_INPUT_SHEET,FUND_VERIFY_SHEET,FUND_PLAN_SHEET].forEach(name=>{
      const s=ss.getSheetByName(name);if(s&&!s.isSheetHidden())s.hideSheet();
    });
    console.log('資産推移の準備完了: '+ss.getUrl());
  } finally {lock.releaseLock();}
}
function recordFundNightly() {
  // 更新処理が失敗しても翌日の予約が残るよう、先に予約する。
  FUND_scheduleNextNight_();
  recordFundDaily();
}
function FUND_scheduleNextNight_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    let next = new Date(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd') + 'T21:55:00+09:00');
    if (next.getTime() <= now.getTime()) {
      next = new Date(next.getTime() + 24 * 60 * 60 * 1000);
    }
    const old = ScriptApp.getProjectTriggers().filter(t =>
      ['recordFundDaily', 'recordFundNightly'].includes(t.getHandlerFunction()) &&
      t.getEventType() === ScriptApp.EventType.CLOCK
    );
    ScriptApp.newTrigger('recordFundNightly').timeBased().at(next).create();
    old.forEach(t => ScriptApp.deleteTrigger(t));
    console.log('次回更新予約: ' + Utilities.formatDate(next, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  } finally {
    lock.releaseLock();
  }
}

function recordFundDaily() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    FUND_validateFunds_();
    const ss = FUND_spreadsheet_(), now = new Date(), today = FUND_dateKey_(now);
    const sheet = FUND_historySheet_(ss);
    const packets = FUND_ENABLE_RAKUTEN_FETCH ? FUNDS.map(FUND_fetchPacket_) : FUND_readInputs_(ss);
    packets.forEach(p => { p.history = p.history.filter(v => v.date<=today); });
    if (packets.some(p=>!p.history.length)) throw new Error('基準価額が3銘柄分必要です。');
    const latest = packets.map(p=>p.history[0].date);
    if (!latest.every(d=>d===latest[0])) throw new Error('基準日が揃わず記録保留: '+latest.join(', '));
    FUND_migrateSummary_(ss);
    const stored = FUND_loadStored_(sheet);
    const last = stored.length ? stored[stored.length-1].date : '';
    if (last && latest[0]<last) throw new Error('取得価格が最終記録より古いため停止しました。');
    const props = PropertiesService.getScriptProperties();
    const baseKey = JSON.stringify([FUND_START_MONTH,FUNDS.map(f=>[f.code,f.units,f.principal])]);
    if (props.getProperty('FUND_BASE_V2') && props.getProperty('FUND_BASE_V2')!==baseKey)
      throw new Error('初期口数・元本は変更しないでください。積立実績は積立履歴で補正します。');
    const calendar = FUND_holidays_();
    const plans = FUND_updatePlans_(ss,packets,today,calendar);
    props.setProperty('FUND_BASE_V2',baseKey);
    const common = packets[0].history.map(p=>p.date).filter(d=>packets.every(p=>p.history.some(v=>v.date===d)));
    if (last && !common.includes(last)) throw new Error('最終記録日が取得履歴の範囲外です。欠損を避けるため停止しました。');
    // 記録済みの取得範囲も再計算する。速報→確報や積立実績補正に対応。
    const targets = last
      ? common.filter(d=>d>=stored[0].date).sort()
      : [common[0]];
    const changes=[];
    let latestSummary;
    targets.forEach(date=>{
      const prior = common.find(d=>d<date) || '';
      const prices = packets.map(p=>p.history.find(v=>v.date===date));
      const oldPrices = prior ? packets.map(p=>p.history.find(v=>v.date===prior)) : null;
      const rows = FUND_makeRows_(date,prices,prior,oldPrices,plans,now,today);
      if (date === latest[0]) latestSummary = rows[0];
      const group = stored.find(g=>g.date===date);
      if (group) {
        const different=rows.some((r,i)=>r.some((v,j)=>j!==11 && FUND_cellKey_(v)!==FUND_cellKey_(group.rows[i][j])));
        if (different) group.indices.forEach((index,i)=>{
          FUND_writeHistory_(sheet,index,[rows[i]]);
          FUND_formatRows_(sheet,index,1);
        });
        if(different) changes.push(date+'更新');
      } else {
        if(last && date<last) throw new Error('記録途中に欠損日があります: '+date);
        const start=sheet.getLastRow()+1;
        FUND_writeHistory_(sheet,start,rows);
        FUND_formatRows_(sheet,start,rows.length);
        changes.push(date+'追加');
      }
    });
    SpreadsheetApp.flush();
    console.log(changes.length ? changes.join(', ') : '新しい基準価額・積立変更はありません。');
    if (latestSummary) FUND_notifyLine_(latestSummary);
  } finally { lock.releaseLock(); }
}

// スクリプトプロパティ FUND_LINE_TO（自分のLINEユーザーID）を使用する。
// 既存の家計簿用 ACCESS_TOKEN（同じMessaging APIチャネル）を使用する。
function FUND_notifyLine_(row) {
  const props = PropertiesService.getScriptProperties();
  const to = typeof FUND_LINE_TO === 'undefined' ? '' : String(FUND_LINE_TO || '').trim();
  if (!to) {
    console.log('LINE通知未設定: スクリプトプロパティ FUND_LINE_TO を設定してください。');
    return;
  }
  if (!/^U[0-9a-f]{32}$/i.test(to)) throw new Error('スクリプトプロパティ FUND_LINE_TO に自分のLINEユーザーIDを設定してください。');
  if (typeof ACCESS_TOKEN === 'undefined' || !ACCESS_TOKEN) throw new Error('LINEのACCESS_TOKENが未設定です。');
  const date = FUND_dateKey_(row[0]);
  const signature = JSON.stringify([to, date, row.slice(1, 11)]);
  if (props.getProperty('FUND_LINE_SENT') === signature) return;
  const yen = n => Number(n).toLocaleString('ja-JP') + '円';
  const signed = n => (n > 0 ? '+' : '') + yen(n);
  const pct = n => (Number(n) * 100).toFixed(2) + '%';
  const text = [
    '資産推移 ' + date.replace(/-/g, '/'),
    '総評価額: ' + yen(row[1]),
    '元本合計: ' + yen(row[2]),
    '含み益: ' + signed(row[3]) + (row[4] === '' ? '' : '（' + pct(row[4]) + '）'),
    '前日比: ' + (row[5] === '' ? '比較データなし' : signed(row[5]) + (row[6] === '' ? '' : '（' + pct(row[6]) + '）')),
    String(row[12]).includes('積立口数は推計') ? '※積立口数は推計です。' : '',
    date === FUND_dateKey_(new Date()) ? '※当日分の公開値（速報の可能性あり）' : ''
  ].filter(Boolean).join('\n');
  const payload = JSON.stringify({to: to, messages: [{type: 'text', text: text}]});
  let pending = JSON.parse(props.getProperty('FUND_LINE_PENDING') || 'null');
  if (!pending || pending.payload !== payload) {
    pending = {payload: payload, key: Utilities.getUuid(), created: Date.now()};
    props.setProperty('FUND_LINE_PENDING', JSON.stringify(pending));
  }
  if (Date.now() - pending.created >= 24 * 60 * 60 * 1000) {
    throw new Error('LINE通知の再送期限が切れました。届いているか確認し、FUND_LINE_PENDINGを削除して再実行してください。');
  }
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + ACCESS_TOKEN, 'X-Line-Retry-Key': pending.key},
    payload: pending.payload,
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const headers = response.getAllHeaders();
  const accepted = status === 409 && Object.keys(headers).some(k => k.toLowerCase() === 'x-line-accepted-request-id');
  if (status !== 200 && !accepted) throw new Error('シート更新済み・LINE通知失敗（HTTP ' + status + '）。設定を確認してrecordFundDailyを再実行してください。');
  props.setProperty('FUND_LINE_SENT', signature);
  props.deleteProperty('FUND_LINE_PENDING');
  console.log('LINE通知リクエスト受付完了: ' + date);
}

// 一か月×一銘柄をキーに管理。残高へ直接加算せず、履歴合計から毎回算出。
function FUND_updatePlans_(ss,packets,today,calendar) {
  const sheet=FUND_ensureSheet_(ss,FUND_PLAN_SHEET,FUND_PLAN_HEADERS);
  const data=sheet.getLastRow()>1 ? sheet.getRange(2,1,sheet.getLastRow()-1,FUND_PLAN_HEADERS.length).getValues() : [];
  const keys=new Set();
  data.forEach(r=>{
    const key=String(r[0])+'|'+r[2];
    if(keys.has(key)) throw new Error('積立履歴が重複しています: '+key);
    if(!/^\d{4}-\d{2}$/.test(String(r[0])) || String(r[0])<FUND_START_MONTH || !FUNDS.some(f=>f.code===r[2]))
      throw new Error('積立履歴の対象月・コードが不正です。2026年9月以前は初期残高に含まれています。');
    keys.add(key);
  });
  const month=today.slice(0,7);
  // 現在より前の月は、当月中に保存した予定が必要。過去の休日を推測しない。
  for(let m=FUND_START_MONTH;m<month;m=FUND_nextMonth_(m)) {
    FUNDS.filter(f=>f.monthlyAmount>0).forEach(f=>{
      if(!keys.has(m+'|'+f.code)) throw new Error('積立履歴の'+m+' '+f.name+'がありません。長期停止分は実績を入力して補完してください。');
    });
  }
  if(month>=FUND_START_MONTH) FUNDS.forEach((f,i)=>{
    if(!f.monthlyAmount || keys.has(month+'|'+f.code)) return;
    if(!FUND_ENABLE_RAKUTEN_FETCH) throw new Error('積立予定の新規作成には公開ページのファンド休日情報が必要です。');
    const schedule=FUND_schedule_(month,packets[i].closed,calendar);
    data.push([month,f.name,f.code,f.monthlyAmount,FUND_dateObject_(schedule.order),
      FUND_dateObject_(schedule.trade),FUND_dateObject_(schedule.settle),FUND_dateObject_(schedule.reflect),
      '','','予定',new Date(),'','','']);
    keys.add(month+'|'+f.code);
  });
  const plans=[];
  data.forEach(r=>{
    const i=FUNDS.findIndex(f=>f.code===r[2]);
    const amount=r[3];
    if(typeof amount!=='number' || !Number.isFinite(amount) || amount<0) throw new Error('積立額が不正です: '+r[0]);
    const actualTrade=r[12]!=='' ? FUND_inputDate_(r[12]) : '';
    const actualUnits=r[13]!=='' ? r[13] : null;
    if(actualUnits!==null && (!Number.isSafeInteger(actualUnits) || actualUnits<=0)) throw new Error('実際の買付口数は正の整数で入力してください。');
    const cancel=r[14]===true || String(r[14]).toUpperCase()==='TRUE';
    if(cancel || amount===0) { r[8]='';r[9]=0;r[10]='取消';r[11]=new Date();return; }
    if(r[10]==='取消') r[10]='予定';
    if(String(r[0])===month && ['予定','価額待ち'].includes(r[10]) && !actualTrade && FUND_ENABLE_RAKUTEN_FETCH) {
      const schedule=FUND_schedule_(month,packets[i].closed,calendar);
      r[4]=FUND_dateObject_(schedule.order);r[5]=FUND_dateObject_(schedule.trade);
      r[6]=FUND_dateObject_(schedule.settle);r[7]=FUND_dateObject_(schedule.reflect);
    }
    const trade=actualTrade || FUND_inputDate_(r[5]);
    if(trade<String(r[0])+'-01') throw new Error('約定日が対象月より前です: '+r[0]);
    // 楽天の保有表示に合わせ、約定日の翌暦日から加算。
    const reflect=FUND_addDays_(trade,1);
    r[7]=FUND_dateObject_(reflect);
    const price=packets[i].history.find(p=>p.date===trade);
    if(trade<=today) {
      if(price) { r[8]=price.nav;r[9]=Math.ceil(amount*10000/price.nav);r[10]='推計'; }
      else if(actualUnits===null && (actualTrade || r[10]!=='推計')) {
        if(trade<packets[i].history[packets[i].history.length-1].date)
          throw new Error('約定日の価額が履歴範囲外です: '+r[0]+' '+r[1]+'。実際の買付口数を入力してください。');
        r[10]='価額待ち';
      }
    }
    if(actualUnits!==null) { r[9]=actualUnits;r[10]='実績口数'; }
    // 日次記録対象日までに未確定の積立があれば、過少残高を記録しない。
    if(reflect<=packets[i].history[0].date && !['推計','実績口数'].includes(r[10]))
      throw new Error('積立口数が未確定です: '+r[0]+' '+r[1]+'。約定日・価額を確認してください。');
    if(['推計','実績口数'].includes(r[10])) {
      if(!Number.isSafeInteger(r[9]) || r[9]<=0) throw new Error('積立口数が不正です。');
      plans.push({code:r[2],reflect:reflect,units:r[9],amount:amount,status:r[10]});
    }
    r[11]=new Date();
  });
  if(data.length) {
    sheet.getRange(2,1,data.length,1).setNumberFormat('@');
    sheet.getRange(2,1,data.length,FUND_PLAN_HEADERS.length).setValues(data);
    sheet.getRange(2,5,data.length,4).setNumberFormat('yyyy/mm/dd');
    sheet.getRange(2,9,data.length,2).setNumberFormat('#,##0');
    sheet.getRange(2,12,data.length,1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sheet.getRange(2,13,data.length,1).setNumberFormat('yyyy/mm/dd');
  }
  return plans;
}

function FUND_balance_(fund,date,plans) {
  return plans.filter(p=>p.code===fund.code && p.reflect<=date).reduce((b,p)=>{
    b.units+=p.units;b.principal+=p.amount;return b;
  },{units:fund.units,principal:fund.principal});
}
function FUND_makeRows_(date,prices,prior,oldPrices,plans,now,today) {
  const principal=FUNDS.reduce((s,f)=>s+FUND_balance_(f,date,plans).principal,0);
  // 実際の口座区分ごとに円未満を四捨五入してから合計。
  const value=FUND_totalValue_(date,prices,plans);
  const oldValue=oldPrices ? FUND_totalValue_(prior,oldPrices,plans) : null;
  const oldPrincipal=prior ? FUNDS.reduce((s,f)=>s+FUND_balance_(f,prior,plans).principal,0) : null;
  const gain=value-principal,deposit=oldPrincipal===null ? 0 : principal-oldPrincipal;
  const delta=oldValue===null ? '' : value-oldValue;
  const profit=oldValue===null ? '' : delta-deposit;
  const note=(date===today?'当日値（速報の可能性あり）':'前日以前の公開値')+
    (plans.some(p=>p.reflect<=date&&p.status==='推計')?'／積立口数は推計':'')+'／口座区分ごとに四捨五入';
  return [[FUND_dateObject_(date),value,principal,gain,principal?gain/principal:'',
    delta,oldValue?delta/oldValue:'',deposit,profit,oldValue?profit/oldValue:'',
    prior?FUND_dateObject_(prior):'',now,note]];
}
function FUND_totalValue_(date,prices,plans) {
  FUND_validateAccounts_();
  return FUNDS.reduce((total,f,i)=>{
    const positions=FUND_ACCOUNT_POSITIONS[f.code].map(p=>({...p}));
    const additions=plans.filter(p=>p.code===f.code&&p.reflect<=date);
    if(additions.length) {
      const account=FUND_MONTHLY_ACCOUNT[f.code];
      let position=positions.find(p=>p.account===account);
      if(!position) {position={account:account,units:0};positions.push(position);}
      position.units+=additions.reduce((sum,p)=>sum+p.units,0);
    }
    return total+positions.reduce((sum,p)=>sum+Math.round(p.units*prices[i].nav/10000),0);
  },0);
}
function FUND_validateAccounts_() {
  FUNDS.forEach(f=>{
    const positions=FUND_ACCOUNT_POSITIONS[f.code];
    if(!Array.isArray(positions)||!positions.length)throw new Error(f.name+'の口座区分別口数をFUND_ACCOUNT_POSITIONSに設定してください。');
    const names=new Set();
    positions.forEach(p=>{
      if(typeof p.account!=='string'||!p.account.trim()||names.has(p.account)||!Number.isSafeInteger(p.units)||p.units<0)
        throw new Error(f.name+'の口座区分名・口数が不正または重複しています。');
      names.add(p.account);
    });
    if(positions.reduce((sum,p)=>sum+p.units,0)!==f.units)throw new Error(f.name+'の区分別口数合計がFUNDSの初期口数と一致しません。');
    if(f.monthlyAmount>0&&!(FUND_MONTHLY_ACCOUNT[f.code]||'').trim())throw new Error(f.name+'の積立先口座区分をFUND_MONTHLY_ACCOUNTに設定してください。');
  });
}
function FUND_holidays_() {
  const props=PropertiesService.getScriptProperties(), year=FUND_dateKey_(new Date()).slice(0,4);
  let saved=JSON.parse(props.getProperty('FUND_HOLIDAYS') || 'null');
  if(!saved || saved.fetched!==FUND_dateKey_(new Date())) {
    const r=UrlFetchApp.fetch('https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv',{muteHttpExceptions:true});
    if(r.getResponseCode()!==200) throw new Error('内閣府の祝日データ取得失敗: HTTP '+r.getResponseCode());
    const dates=Utilities.parseCsv(r.getContentText('Shift_JIS')).map(row=>{
      const m=String(row[0]).match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      return m ? m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0') : '';
    }).filter(d=>d && d.slice(0,4)>=FUND_START_MONTH.slice(0,4));
    if(!dates.some(d=>d.slice(0,4)===year)) throw new Error('当年の公式祝日データがありません。');
    saved={fetched:FUND_dateKey_(new Date()),dates:dates};
    props.setProperty('FUND_HOLIDAYS',JSON.stringify(saved));
  }
  return new Set(saved.dates);
}
function FUND_business_(date,closed,calendar) {
  if(!calendar.has(date.slice(0,4)+'-01-01')) throw new Error('祝日カレンダー未公表の年です: '+date.slice(0,4));
  const day=new Date(date+'T12:00:00Z').getUTCDay(), md=date.slice(5);
  return day!==0 && day!==6 && !calendar.has(date) && !closed.includes(date) &&
    !['01-01','01-02','01-03','12-31'].includes(md);
}
function FUND_nextBusiness_(date,closed,calendar,include) {
  let d=include?date:FUND_addDays_(date,1);
  for(let i=0;i<40;i++,d=FUND_addDays_(d,1)) if(FUND_business_(d,closed,calendar)) return d;
  throw new Error('営業日を計算できません。');
}
function FUND_schedule_(month,closed,calendar) {
  const order=FUND_nextBusiness_(month+'-'+String(FUND_BUY_DAY).padStart(2,'0'),closed,calendar,true);
  const trade=FUND_nextBusiness_(order,closed,calendar,false);
  let settle=trade;
  // 受渡予定は提示された実績の「約定から国内3営業日」を使用。
  for(let i=0;i<3;i++) settle=FUND_nextBusiness_(settle,[],calendar,false);
  return {order:order,trade:trade,settle:settle,reflect:FUND_addDays_(trade,1)};
}
function FUND_nextMonth_(month) { return FUND_addDays_(month+'-01',32).slice(0,7); }
function FUND_addDays_(key,n) {
  const d=new Date(key+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10);
}

function FUND_fetchPacket_(fund) {
  const detail=FUND_fetchHtml_('https://www.rakuten-sec.co.jp/web/fund/detail/?ID='+encodeURIComponent(fund.code));
  const args=Array.from(detail.html.matchAll(/show_price\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\s*\)/g))
    .find(m=>m[1]===fund.code && m[4]==='price');
  if(!args) throw new Error('過去データリンクなし: '+fund.code);
  const url='https://www.rakuten-sec.co.jp/web/fund/detail/price.html?Code='+encodeURIComponent(args[1])+
    '&fundStart='+encodeURIComponent(args[2])+'&date='+encodeURIComponent(args[3])+'&action=price';
  const result=FUND_fetchHtml_(url), closed=FUND_parseClosed_(detail.html);
  return {history:FUND_parseHistory_(result.html).map(p=>Object.assign(p,{url:url,contentType:result.contentType})),closed:closed};
}
function FUND_parseClosed_(html) {
  const row=(html.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi)||[]).find(r=>
    /<th\b[^>]*>[\s\S]*?ファンド休日[\s\S]*?<\/th>/i.test(r));
  const cell=row && row.match(/<td\b[^>]*>([\s\S]*?)<\/td>/i);
  if(!cell) throw new Error('ファンド休日を抽出できません。');
  const text=FUND_stripHtml_(cell[1]), now=FUND_dateKey_(new Date());
  if(!text || /^(?:-+|―|なし|該当なし)$/.test(text)) return [];
  const tokens=text.split(/[、,]/),year=Number(now.slice(0,4)),month=Number(now.slice(5,7));
  return tokens.map(t=>{
    const m=t.match(/^(\d{1,2})\/(\d{1,2})$/);
    if(!m) throw new Error('ファンド休日の形式変更: '+text);
    const y=Number(m[1])<month ? year+1 : year;
    const d=y+'-'+m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0');FUND_validateDate_(d);return d;
  });
}
function FUND_fetchHtml_(url) {
  if(!FUND_ENABLE_RAKUTEN_FETCH) throw new Error('楽天証券の自動取得は無効です。');
  const r=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:false});
  if(r.getResponseCode()!==200) throw new Error('HTTP '+r.getResponseCode()+': '+url);
  const h=r.getAllHeaders(), k=Object.keys(h).find(k=>k.toLowerCase()==='content-type');
  if(!k || !/text\/html/i.test(String(h[k]))) throw new Error('想定外のContent-Type');
  return {html:r.getContentText('UTF-8'),contentType:String(h[k])};
}
function FUND_parseHistory_(html) {
  const table=(html.match(/<table\b[^>]*>[\s\S]*?<\/table>/gi)||[]).find(t=>{
    const h=(t.match(/<thead\b[^>]*>([\s\S]*?)<\/thead>/i)||[,''])[1];
    return /年月日/.test(h)&&/基準価額/.test(h)&&/純資産総額/.test(h);
  });
  if(!table) throw new Error('基準価額表が見つかりません。');
  const data=[],seen=new Set();
  (table.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi)||[]).forEach(r=>{
    const c=(r.match(/<(?:th|td)\b[^>]*>[\s\S]*?<\/(?:th|td)>/gi)||[]).map(FUND_stripHtml_);
    const m=c[0]&&c[0].match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);if(!m)return;
    if(c.length!==5 || !/^[\d,]+(?:\.\d+)?円$/.test(c[1])) throw new Error('基準価額表の形式変更');
    const date=m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0'),nav=Number(c[1].replace(/[円,]/g,''));
    FUND_validateDate_(date);
    if(!Number.isFinite(nav)||nav<=0||seen.has(date)) throw new Error('不正・重複の基準価額');
    seen.add(date);data.push({date:date,nav:nav});
  });
  if(!data.length) throw new Error('HTTP200ですが基準価額表にデータがありません。');
  return data.sort((a,b)=>b.date.localeCompare(a.date));
}
function FUND_stripHtml_(s) { return s.replace(/<[^>]*>/g,'').replace(/&nbsp;|&#160;/gi,' ').replace(/&amp;/g,'&').replace(/\s+/g,'').trim(); }

function FUND_readInputs_(ss) {
  const sheet=FUND_ensureSheet_(ss,FUND_INPUT_SHEET,['基準日','銘柄コード','基準価額（円／1万口）']);
  if(sheet.getLastRow()<2) throw new Error('基準価額入力が空です。');
  const packets=FUNDS.map(()=>({history:[],closed:[]}));
  sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues().forEach(r=>{
    if(r.every(v=>v===''))return;
    const i=FUNDS.findIndex(f=>f.code===r[1]),date=FUND_inputDate_(r[0]);
    if(i<0||typeof r[2]!=='number'||!Number.isFinite(r[2])||r[2]<=0) throw new Error('基準価額入力が不正です。');
    if(packets[i].history.some(p=>p.date===date)) throw new Error('基準価額入力が重複しています。');
    packets[i].history.push({date:date,nav:r[2],url:'本人による基準価額入力'});
  });
  packets.forEach(p=>p.history.sort((a,b)=>b.date.localeCompare(a.date)));return packets;
}
function FUND_loadStored_(sheet) {
  if(sheet.getLastRow()<2)return [];
  const seen=new Set();
  return sheet.getRange(2,1,sheet.getLastRow()-1,FUND_HEADERS.length+1).getValues().map((displayRow,i)=>{
    const r=[displayRow[0],...displayRow.slice(2)];
    const date=FUND_dateKey_(r[0]);
    if(seen.has(date))throw new Error('資産推移の日付が重複しています: '+date);
    if(typeof r[1]!=='number'||!Number.isFinite(r[1]))throw new Error('資産推移の評価額が不正です。');
    seen.add(date);return {date:date,rows:[r],indices:[i+2]};
  }).sort((a,b)=>a.date.localeCompare(b.date));
}
function FUND_migrateSummary_(ss) {
  const target=ss.getSheetByName(FUND_HISTORY_SHEET),old=ss.getSheetByName(FUND_LEGACY_SHEET);
  if(!target||target.getLastRow()>1||!old||old.getLastRow()<2)return;
  const h=old.getRange(1,1,1,3).getValues()[0];
  if(h[0]!=='基準日'||h[2]!=='銘柄コード')throw new Error('旧投信シートの見出しが想定と異なります。');
  const data=old.getRange(2,1,old.getLastRow()-1,Math.min(old.getLastColumn(),18)).getValues();
  const seen=new Set(),rows=data.filter(r=>r[2]==='TOTAL').map(r=>{
    const key=FUND_dateKey_(r[0]);if(seen.has(key))throw new Error('旧記録の合計行が重複しています。');seen.add(key);
    const value=Math.ceil(r[5]),principal=r[6],gain=value-principal;
    const before=typeof r[9]==='number' ? Math.ceil(r[5]-r[9]) : null;
    const delta=before===null?'':value-before,deposit=typeof r[15]==='number'?r[15]:0;
    const profit=before===null?'':delta-deposit;
    return [r[0],value,principal,gain,principal?gain/principal:'',delta,before?delta/before:'',
      deposit,profit,before?profit/before:'',r[11]||'',r[12]||new Date(),'旧記録から移行／合計を円単位で切り上げ'];
  }).sort((a,b)=>a[0].getTime()-b[0].getTime());
  if(rows.length){FUND_writeHistory_(target,2,rows);FUND_formatRows_(target,2,rows.length);}
}
function FUND_validateFunds_() {
  FUND_validateAccounts_();
  FUNDS.forEach(f=>{
    if(!Number.isSafeInteger(f.units)||f.units<0||!Number.isFinite(f.principal)||f.principal<0||
      !Number.isSafeInteger(f.monthlyAmount)||f.monthlyAmount<0)throw new Error('口数・元本・積立額の設定が不正です: '+f.name);
  });
}
function FUND_spreadsheet_() {
  const ss=FUND_SPREADSHEET_ID ? SpreadsheetApp.openById(FUND_SPREADSHEET_ID) :
    SETTINGS_spreadsheet_();
  if(!ss)throw new Error('スプレッドシートが見つかりません。');return ss;
}
function FUND_ensureSheet_(ss,name,headers) {
  const s=ss.getSheetByName(name)||ss.insertSheet(name);
  if(s.getMaxColumns()<headers.length)s.insertColumnsAfter(s.getMaxColumns(),headers.length-s.getMaxColumns());
  if(s.getLastRow()&&JSON.stringify(s.getRange(1,1,1,headers.length).getValues()[0])!==JSON.stringify(headers))
    throw new Error('見出しが想定と異なります: '+name);
  s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eef5').setWrap(true);
  s.setFrozenRows(1);s.setColumnWidths(1,headers.length,150);s.setRowHeight(1,42);return s;
}
// 内部の13項目は維持し、シートだけB列に総資産を追加する。
function FUND_historySheet_(ss) {
  const headers=[FUND_HEADERS[0],'総資産（投信＋口座）（円）',...FUND_HEADERS.slice(1)];
  let sheet=ss.getSheetByName(FUND_HISTORY_SHEET);
  let migrated=false;
  if(sheet&&sheet.getLastRow()>0) {
    const old=sheet.getRange(1,1,1,FUND_HEADERS.length).getValues()[0];
    if(JSON.stringify(old)===JSON.stringify(FUND_HEADERS)) {
      sheet.insertColumnBefore(2);
      sheet.getRange(1,2).setValue(headers[1]);
      migrated=true;
    } else {
      // 列挿入直後にタイムアウトした場合にも見出しを復旧。
      const current=sheet.getRange(1,1,1,headers.length).getValues()[0];
      if(current[1]===''&&JSON.stringify([current[0],...current.slice(2)])===JSON.stringify(FUND_HEADERS)) {
        sheet.getRange(1,2).setValue(headers[1]);migrated=true;
      }
    }
  }
  sheet=FUND_ensureSheet_(ss,FUND_HISTORY_SHEET,headers);
  if(sheet.getLastRow()>1) {
    // まとめて設定。再実行しても列や行は増えない。
    FUND_assetFormulas_(sheet,2,sheet.getLastRow()-1);
    if(migrated)FUND_formatRows_(sheet,2,sheet.getLastRow()-1);
  }
  return sheet;
}
function FUND_writeHistory_(sheet,start,rows) {
  sheet.getRange(start,1,rows.length,FUND_HEADERS.length+1)
    .setValues(rows.map(row=>[row[0],'',...row.slice(1)]));
}
function FUND_assetFormulas_(sheet,start,count) {
  const formulas=Array.from({length:count},(_,i)=>{
    const r=start+i;
    // 現在の年・月の口座残高。1月=B42、10月=K42、12月=M42。
    // シートや残高が不正ならエラーを表示し、0円には置き換えない。
    const prefix=`"'"&YEAR(TODAY())&"'!"`;
    const annualYear='INDIRECT('+prefix+'&"A1")';
    const cash='INDIRECT('+prefix+'&ADDRESS(42,MONTH(TODAY())+1,4))';
    return ['=IF(AND('+annualYear+'=YEAR(TODAY()),ISNUMBER('+cash+')),$C'+r+'+'+cash+',NA())'];
  });
  sheet.getRange(start,2,count,1).setFormulas(formulas).setNumberFormat('#,##0');
}
function FUND_formatRows_(sheet,start,count) {
  sheet.getRange(start,1,count,1).setNumberFormat('yyyy/mm/dd');
  [3,4,5].forEach(c=>sheet.getRange(start,c,count,1).setNumberFormat('#,##0'));
  [6,8,11].forEach(c=>sheet.getRange(start,c,count,1).setNumberFormat('0.00%'));
  [7,9,10].forEach(c=>sheet.getRange(start,c,count,1).setNumberFormat('+#,##0;-#,##0;0'));
  sheet.getRange(start,12,count,1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(start,13,count,1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  FUND_assetFormulas_(sheet,start,count);
}
function FUND_cellKey_(v) { return v instanceof Date ? v.getTime() : v; }
function FUND_dateKey_(d) {
  if(!(d instanceof Date)||isNaN(d.getTime()))throw new Error('日付セルが不正です。');
  return Utilities.formatDate(d,FUND_TZ,'yyyy-MM-dd');
}
function FUND_dateObject_(key) { return new Date(key+'T00:00:00+09:00'); }
function FUND_inputDate_(v) {
  const key=v instanceof Date ? FUND_dateKey_(v) : String(v).trim().replace(/\//g,'-');
  FUND_validateDate_(key);return key;
}
function FUND_validateDate_(key) {
  if(!/^\d{4}-\d{2}-\d{2}$/.test(key)||isNaN(FUND_dateObject_(key).getTime())||FUND_dateKey_(FUND_dateObject_(key))!==key)
    throw new Error('基準日はyyyy-MM-ddで入力してください: '+key);
}
