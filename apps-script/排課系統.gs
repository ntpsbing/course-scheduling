/**
 * 楠梓國小科任排課系統：Google 試算表端程式（Apps Script）
 *
 * 安裝：試算表 →「擴充功能」→「Apps Script」→ 把這整份貼上 → 儲存 → 重新整理試算表
 * 之後試算表上方會多一個「排課系統」選單，第一次請先按「初次設定」。
 *
 * 流程：按「開始排課」→ 通知 GitHub → GitHub 來拿資料（老師姓名先換成代號）→ 排課
 *      → 把結果送回來 → 換回真名、寫進「排課結果」等工作表。
 * 真實姓名只留在這份試算表，不會送到 GitHub。
 */

const REPO = 'ntpsbing/course-scheduling';
const VIEWER_URL = 'https://raw.githubusercontent.com/' + REPO + '/main/web/' +
  encodeURIComponent('課表檢視.template.html');
const SHEETS = ['規則表', '科任配課', '本土語預排', '特教抽離', '共同時段', '職務限制', '班級教室', '老師需求'];
const OPTIONAL_SHEETS = ['班級需求']; // 沒有這張也能排
// 哪些欄放老師姓名（從 0 起算）、資料從第幾列開始 → 送出前換成代號
const NAME_COLS = {
  '科任配課': { cols: [0], from: 5 },
  '本土語預排': { cols: [4], from: 5 },
  '特教抽離': { cols: [5], from: 5 },
  '職務限制': { cols: [6], from: 4 },
  '老師需求': { cols: [0], from: 5 },
};
const DAYS = ['一', '二', '三', '四', '五'];
const PERIODS = [1, 2, 3, 4, 5, 6, 7];
const GRAY = { subject: '#ffffff', lock: '#d9d9d9', keep: '#d9d9d9', sped: '#bfbfbf', lang: '#efefef', off: '#808080', head: '#d9d9d9' };
const props = () => PropertiesService.getScriptProperties();

/* ---------------- 選單 ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('排課系統')
    .addItem('▶ 開始排課', 'startScheduling')
    .addItem('查看課表（網頁）', 'openViewer')
    .addItem('查看排課進度', 'showStatus')
    .addSeparator()
    .addItem('建立／重建排課控制台', 'buildConsole')
    .addItem('初次設定', 'openSetup')
    .addToUi();
}

function startScheduling() {
  const ui = SpreadsheetApp.getUi();
  const p = props();
  if (!p.getProperty('GITHUB_TOKEN') || !p.getProperty('SCHED_KEY')) {
    ui.alert('還沒完成設定', '請先按「排課系統 → 初次設定」，照步驟完成後再排課。', ui.ButtonSet.OK);
    return;
  }
  const missing = SHEETS.filter(n => !SpreadsheetApp.getActive().getSheetByName(n));
  if (missing.length) {
    ui.alert('找不到工作表', '這份試算表缺少：' + missing.join('、') + '\n請用「排課需求收集表」的格式。', ui.ButtonSet.OK);
    return;
  }
  const ok = ui.alert('開始排課', '電腦會依「規則表」排出全校科任課表，大約 1～3 分鐘。\n' +
    '排好後會覆蓋「排課結果」「老師課表」「規則體檢」三張工作表。要開始嗎？', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  const jobId = Utilities.getUuid().slice(0, 8);
  p.setProperty('JOB', JSON.stringify({ id: jobId, status: 'queued', at: Date.now() }));
  p.setProperty('SHEET_ID', SpreadsheetApp.getActive().getId());
  const res = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO + '/dispatches', {
    method: 'post', muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + p.getProperty('GITHUB_TOKEN'), Accept: 'application/vnd.github+json' },
    payload: JSON.stringify({ event_type: 'schedule' }),
  });
  const code = res.getResponseCode();
  if (code !== 204) {
    const why = code === 401 ? 'GitHub 金鑰無效或過期，請到「初次設定」換一把新的。'
      : code === 403 || code === 404 ? 'GitHub 金鑰沒有這個 repo 的權限（需要 Contents：Read and write）。'
        : 'GitHub 回應 ' + code + '，請稍後再試。';
    log_('失敗', why);
    ui.alert('沒有送出', why, ui.ButtonSet.OK);
    return;
  }
  log_('已送出', '編號 ' + jobId + '，等待 GitHub 計算');
  SpreadsheetApp.getActive().toast('已送出！大約 1～3 分鐘後結果會自動寫回，可以先做別的事。', '排課中', 10);
}

function showStatus() {
  const job = JSON.parse(props().getProperty('JOB') || '{}');
  const text = { queued: '已送出，等待 GitHub 開始計算', running: '計算中', done: '完成', error: '失敗' }[job.status] || '還沒排過';
  const mins = job.at ? Math.round((Date.now() - job.at) / 60000) : 0;
  SpreadsheetApp.getUi().alert('排課進度', text + (job.at ? '（' + mins + ' 分鐘前）' : '') +
    (job.message ? '\n' + job.message : '') +
    '\n\n詳細紀錄在「排課紀錄」工作表；GitHub 執行情形：https://github.com/' + REPO + '/actions',
  SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ---------------- GitHub 來拿資料／送結果 ---------------- */

function doGet(e) {
  const q = (e && e.parameter) || {};
  if (!q.key) return HtmlService.createHtmlOutput('這是排課系統的連線網址，給 GitHub 使用，不需要直接開啟。');
  if (q.key !== props().getProperty('SCHED_KEY')) return json_({ ok: false, error: '通關密語不對' });
  const p = props();
  let job = JSON.parse(p.getProperty('JOB') || '{}');
  if (!job.id || job.status === 'done' || job.status === 'error') job = { id: Utilities.getUuid().slice(0, 8) }; // GitHub 手動重跑
  const ss = SpreadsheetApp.openById(p.getProperty('SHEET_ID'));
  const built = buildTables_(ss);
  job.status = 'running';
  job.at = Date.now();
  p.setProperty('JOB', JSON.stringify(job));
  p.setProperty('NAME_MAP', JSON.stringify(built.back));
  log_('計算中', 'GitHub 已取得資料（' + Object.keys(built.back).length + ' 位老師已換成代號）', ss);
  return json_({ ok: true, jobId: job.id, source: ss.getName(), tables: built.tables, timeLimit: 60 });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.key !== props().getProperty('SCHED_KEY')) return json_({ ok: false, error: '通關密語不對' });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const p = props();
    const job = JSON.parse(p.getProperty('JOB') || '{}');
    const ss = SpreadsheetApp.openById(p.getProperty('SHEET_ID'));
    if (body.action === 'error') {
      job.status = 'error'; job.message = body.error;
      p.setProperty('JOB', JSON.stringify(job));
      log_('失敗', body.error, ss);
      return json_({ ok: true });
    }
    const result = restoreNames_(body.result, JSON.parse(p.getProperty('NAME_MAP') || '{}'));
    saveResultFile_(ss, result);
    writeResult_(ss, result);
    const m = result.meta;
    const msg = result.rules
      ? m.n_lessons + ' 節排好，' + result.rules.filter(r => r.ok === true).length + '/' + result.rules.length + ' 條規則完全做到（' + m.seconds + ' 秒）'
      : '排不出來，原因請看「規則體檢」';
    job.status = result.rules ? 'done' : 'error'; job.message = msg; job.at = Date.now();
    p.setProperty('JOB', JSON.stringify(job));
    log_(result.rules ? '完成' : '失敗', msg, ss);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 讀八張工作表；老師姓名換成 T001、T002…（對照表只存在這裡） */
function buildTables_(ss) {
  const to = {}, back = {};
  const code = name => {
    if (!to[name]) { const c = 'T' + ('00' + (Object.keys(to).length + 1)).slice(-3); to[name] = c; back[c] = name; }
    return to[name];
  };
  const tables = {};
  SHEETS.concat(OPTIONAL_SHEETS).forEach(n => {
    if (!ss.getSheetByName(n)) return;
    const values = ss.getSheetByName(n).getDataRange().getValues();
    const rule = NAME_COLS[n];
    if (rule) {
      values.forEach((row, i) => {
        if (i + 1 < rule.from || row.join('').indexOf('範例') >= 0) return;
        rule.cols.forEach(c => { const v = String(row[c] || '').trim(); if (v) row[c] = code(v); });
      });
    }
    tables[n] = values.map(r => r.map(v => (v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : v)));
  });
  return { tables: tables, back: back };
}

function restoreNames_(result, back) {
  const text = JSON.stringify(result).replace(/T\d{3}/g, c => (back[c] ? JSON.stringify(back[c]).slice(1, -1) : c));
  return JSON.parse(text);
}

/* ---------------- 寫回試算表 ---------------- */

function writeResult_(ss, r) {
  if (r.classes) {
    writeGrid_(ss, '排課結果', '班級', Object.keys(r.classes), name => r.classes[name].cells, c => {
      if (c.type === 'subject') return c.label + '\n' + c.who + (c.room && c.room !== '原班教室' ? '\n@' + c.room.slice(0, 2) : '');
      if (c.type === 'sped') return c.label + '\n抽離・導師';
      if (c.type === 'keep') return '導師\n' + (c.note || '');
      if (c.type === 'off') return '';
      return c.label + (c.who ? '\n' + c.who : '');
    }, name => { const c = r.classes[name]; return [c.subject_hours, c.keep, c.free]; }, ['科任節數', '保留給導師', '導師空格'], '導師');
    const ts = Object.keys(r.teachers).sort();
    writeGrid_(ss, '老師課表', '老師', ts, name => r.teachers[name].cells, c => {
      if (c.type === 'subject') return c.label + '\n' + c.subject + (c.conflict ? '\n⚠希望不排' : '');
      return c.label;
    }, name => { const t = r.teachers[name]; return [t.hours, Math.max.apply(null, t.per_day)]; }, ['每週節數', '單日最多'], '');
  }
  const sh = sheet_(ss, '規則體檢');
  const rows = [['編號', '層級', '規則積木', '對象', '內容', '結果', '說明']];
  (r.rules || []).forEach(x => rows.push([x.id, x.level, x.block, x.target, x.content,
    x.ok === true ? '○ 做到' : x.ok === false ? '△ 部分' : '— 略過', x.summary + (x.details && x.details.length ? '\n・' + x.details.slice(0, 10).join('\n・') : '')]));
  (r.insights || []).concat(r.issues || []).forEach(x => rows.push(['提醒', '', '', '', x.text, '', '']));
  rows.push(['', '', '', '', '產生時間：' + (r.meta.generated || '') + '　結果：' + r.meta.status + '（' + r.meta.seconds + ' 秒）', '', '']);
  sh.getRange(1, 1, rows.length, 7).setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, 1, 7).setBackground(GRAY.head).setFontWeight('bold');
  sh.setFrozenRows(1);
  [50, 50, 100, 110, 280, 70, 380].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function writeGrid_(ss, title, firstHead, names, cellsOf, textOf, extraOf, extraHeads, emptyText) {
  const sh = sheet_(ss, title);
  const slots = [];
  DAYS.forEach(d => PERIODS.forEach(p => slots.push(d + p)));
  const head = [firstHead].concat(slots.map(k => k.charAt(0) + '\n' + k.charAt(1)), extraHeads);
  const values = [head], bgs = [head.map(() => GRAY.head)], bold = [head.map(() => 'bold')];
  names.forEach(n => {
    const cells = cellsOf(n);
    const row = [n], bg = ['#ffffff'], bw = ['bold'];
    slots.forEach(k => {
      const c = cells[k];
      row.push(c ? textOf(c) : emptyText);
      bg.push(c ? (GRAY[c.type] || '#ffffff') : '#ffffff');
      bw.push(c && c.type === 'subject' ? 'bold' : 'normal');
    });
    extraOf(n).forEach(v => { row.push(v); bg.push('#ffffff'); bw.push('normal'); });
    values.push(row); bgs.push(bg); bold.push(bw);
  });
  const range = sh.getRange(1, 1, values.length, head.length);
  range.setValues(values).setBackgrounds(bgs).setFontWeights(bold).setWrap(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  range.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 70);
  sh.setColumnWidths(2, slots.length, 58);
  // 每天第 1 節左邊加粗線，方便看
  DAYS.forEach((d, i) => sh.getRange(1, 2 + i * 7, values.length, 1)
    .setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
}

function sheet_(ss, name) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  return sh;
}

function saveResultFile_(ss, result) {
  const p = props();
  const text = JSON.stringify(result);
  const id = p.getProperty('RESULT_FILE');
  try {
    if (id) { DriveApp.getFileById(id).setContent(text); return; }
  } catch (err) { /* 檔案被刪了就重建 */ }
  const parent = DriveApp.getFileById(ss.getId()).getParents();
  const folder = parent.hasNext() ? parent.next() : DriveApp.getRootFolder();
  p.setProperty('RESULT_FILE', folder.createFile('排課結果_最新.json', text, 'application/json').getId());
}

function log_(status, message, ss) {
  ss = ss || SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('排課紀錄') || ss.insertSheet('排課紀錄');
  if (sh.getLastRow() === 0) sh.appendRow(['時間', '狀態', '說明']).getRange(1, 1, 1, 3).setBackground(GRAY.head).setFontWeight('bold');
  sh.appendRow([Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'), status, message]);
  updateConsole_(ss, status, message);
}

/* ---------------- 排課控制台（工作表上的大按鈕） ---------------- */

const CONSOLE = '排課控制台';
const BUTTONS = [ // [圖片, 按下去執行的功能]；圖片原檔在 apps-script/按鈕/
  ['start', 'startScheduling'], ['view', 'openViewer'], ['status', 'showStatus'],
];

function buildConsole() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONSOLE) || ss.insertSheet(CONSOLE, 0);
  sh.clear();
  sh.getImages().forEach(img => img.remove());
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 24);
  sh.setColumnWidths(2, 3, 290);
  sh.getRange('B2').setValue('科任排課控制台').setFontSize(20).setFontWeight('bold');
  sh.getRange('B3').setValue('按下面的按鈕就能排課，不用打任何指令。第一次使用請先按上方「排課系統 → 初次設定」。')
    .setFontColor('#5d6b78');
  sh.setRowHeight(5, 84);
  BUTTONS.forEach(([name, fn], i) => {
    const blob = Utilities.newBlob(Utilities.base64Decode(BUTTON_PNG[name]), 'image/png', name + '.png');
    sh.insertImage(blob, 2 + i, 5, 8, 6).setWidth(270).setHeight(68).assignScript(fn);
  });
  const rows = [['目前狀態', '還沒排過'], ['最後更新', ''], ['結果', '']];
  sh.getRange('B7').setValue('排課狀態').setFontWeight('bold').setFontSize(13);
  sh.getRange(8, 2, 3, 2).setValues(rows);
  sh.getRange('B8:B10').setFontColor('#5d6b78');
  sh.getRange('C8:C10').setFontWeight('bold');
  [8, 9, 10].forEach(r => sh.getRange(r, 3, 1, 2).merge());
  sh.getRange('B12').setValue('怎麼用').setFontWeight('bold').setFontSize(13);
  sh.getRange('B13:B16').setValues([
    ['① 先把「規則表」「科任配課」「老師需求」等工作表填好'],
    ['② 按「▶ 開始排課」，約 1～3 分鐘，結果會寫進「排課結果」「老師課表」「規則體檢」'],
    ['③ 按「查看課表」看漂亮的網頁版；「規則體檢」會告訴你哪些規則沒做到、可以怎麼調'],
    ['④ 想調整就改規則表或資料，再按一次「開始排課」；每次紀錄都在「排課紀錄」'],
  ]);
  sh.getRange('B18').setValue('小提醒：手機版 Google 試算表不能按按鈕，請用電腦操作。').setFontColor('#5d6b78').setFontSize(9);
  const job = JSON.parse(props().getProperty('JOB') || '{}');
  if (job.status) updateConsole_(ss, { queued: '已送出', running: '計算中', done: '完成', error: '失敗' }[job.status], job.message || '');
  ss.setActiveSheet(sh);
}

function updateConsole_(ss, status, message) {
  const sh = ss.getSheetByName(CONSOLE);
  if (!sh) return;
  const color = { '完成': '#2f7d4f', '失敗': '#b3342b' }[status] || '#a3620a';
  sh.getRange('C8').setValue(status).setFontColor(color);
  sh.getRange('C9').setValue(Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'));
  sh.getRange('C10').setValue(message || '');
}

const BUTTON_PNG = {
  start: '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',
  view: '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',
  status: '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',
};

/* ---------------- 課表檢視（只有能開這份試算表的人看得到） ---------------- */

function openViewer() {
  const id = props().getProperty('RESULT_FILE');
  if (!id) { SpreadsheetApp.getUi().alert('還沒有排課結果，請先按「開始排課」。'); return; }
  const data = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8');
  const tpl = UrlFetchApp.fetch(VIEWER_URL).getContentText('UTF-8');
  const html = tpl.replace('__DATA__', data.replace(/<\//g, '<\\/'));
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1200).setHeight(800), '科任排課結果');
}

/* ---------------- 初次設定 ---------------- */

function openSetup() {
  const p = props();
  if (!p.getProperty('SCHED_KEY')) p.setProperty('SCHED_KEY', Utilities.getUuid().replace(/-/g, ''));
  p.setProperty('SHEET_ID', SpreadsheetApp.getActive().getId());
  if (!SpreadsheetApp.getActive().getSheetByName(CONSOLE)) buildConsole();
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(SETUP_HTML).setTitle('排課系統 初次設定'));
}

function getSetupInfo() {
  const p = props();
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (err) { /* 還沒部署 */ }
  return { key: p.getProperty('SCHED_KEY'), url: url, hasToken: !!p.getProperty('GITHUB_TOKEN') };
}

function saveToken(token) {
  token = String(token || '').trim();
  const res = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO, {
    muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
  });
  if (res.getResponseCode() !== 200) return { ok: false, msg: '這把金鑰不能用（GitHub 回應 ' + res.getResponseCode() + '），請確認有複製完整。' };
  props().setProperty('GITHUB_TOKEN', token);
  return { ok: true, msg: '金鑰正確，已儲存！' };
}

const SETUP_HTML = `
<style>
 body{font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif;font-size:13px;line-height:1.6;color:#1c2530;padding:4px 12px}
 h3{font-size:14px;margin:16px 0 4px} ol{padding-left:18px;margin:4px 0} li{margin:2px 0}
 .box{background:#f1f4f6;border-radius:6px;padding:6px 8px;word-break:break-all;font-family:monospace;font-size:12px}
 button{background:#1e5a7a;color:#fff;border:0;border-radius:6px;padding:6px 12px;cursor:pointer;margin-top:4px}
 input{width:100%;box-sizing:border-box;padding:6px;border:1px solid #c5ced6;border-radius:6px}
 .ok{color:#2f7d4f;font-weight:bold} .bad{color:#b3342b;font-weight:bold} small{color:#5d6b78}
</style>
<p>共 4 步，只要做一次。做完就能按「開始排課」。</p>

<h3>① 部署成網頁應用程式</h3>
<ol><li>Apps Script 編輯器右上角「部署」→「新增部署作業」</li>
<li>類型選「網頁應用程式」；執行身分：<b>我</b>；誰可以存取：<b>所有人</b></li>
<li>按「部署」，第一次會要求授權，按「允許」</li></ol>
<small>為什麼要「所有人」：GitHub 沒有學校帳號，要靠下面的通關密語進來；沒有密語的人什麼都拿不到。</small>

<h3>② 取得 GitHub 金鑰</h3>
<ol><li>開 <a href="https://github.com/settings/personal-access-tokens/new" target="_blank">GitHub 新增金鑰頁</a></li>
<li>名稱填「排課系統」，期限選最長（到期要換新的）</li>
<li>Repository access 選「Only select repositories」→ course-scheduling</li>
<li>Permissions → Repository permissions → <b>Contents：Read and write</b></li>
<li>按「Generate token」，複製貼到下面</li></ol>
<input id="tk" type="password" placeholder="github_pat_…"><button onclick="save()">儲存並測試</button>
<div id="tkmsg"></div>

<h3>③ 把兩個資料貼到 GitHub</h3>
<ol><li>開 <a href="https://github.com/${REPO}/settings/secrets/actions/new" target="_blank">GitHub Secrets 新增頁</a></li>
<li>Name 填 <b>SCHED_URL</b>，Secret 貼這個網址：<div class="box" id="url">（請先完成 ①，再按重新整理）</div></li>
<li>再新增一個：Name 填 <b>SCHED_KEY</b>，Secret 貼這組通關密語：<div class="box" id="key"></div></li></ol>
<small>網址結尾應該是 <b>/exec</b>；如果不是，請改貼「部署」視窗裡顯示的網頁應用程式網址。</small><br>
<button onclick="load()">重新整理</button>

<h3>④ 試試看</h3>
<p>關掉這個視窗，到最前面的「排課控制台」工作表按「▶ 開始排課」大按鈕。</p>

<script>
 function load(){google.script.run.withSuccessHandler(function(i){
   document.getElementById('key').textContent=i.key;
   if(i.url)document.getElementById('url').textContent=i.url;
   if(i.hasToken)document.getElementById('tkmsg').innerHTML='<span class="ok">✓ 已設定過金鑰</span>';
 }).getSetupInfo();}
 function save(){var m=document.getElementById('tkmsg');m.textContent='測試中…';
   google.script.run.withSuccessHandler(function(r){m.innerHTML='<span class="'+(r.ok?'ok':'bad')+'">'+r.msg+'</span>';})
   .saveToken(document.getElementById('tk').value);}
 load();
</script>`;
