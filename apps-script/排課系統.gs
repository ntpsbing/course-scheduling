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
const GRAY = { subject: '#ffffff', lock: '#d9d9d9', sped: '#bfbfbf', lang: '#efefef', off: '#808080', head: '#d9d9d9' };
const props = () => PropertiesService.getScriptProperties();

/* ---------------- 選單 ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('排課系統')
    .addItem('▶ 開始排課', 'startScheduling')
    .addItem('查看課表（網頁）', 'openViewer')
    .addItem('查看排課進度', 'showStatus')
    .addSeparator()
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
  SHEETS.forEach(n => {
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
}

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
<p>關掉這個視窗，按「排課系統 → ▶ 開始排課」。</p>

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
