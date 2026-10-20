/**
 * 第49期幹部續任調查・即時資料介面（Google 試算表 + Apps Script）
 *
 * 用法：在「幹部續任調查表（回覆）」試算表的「擴充功能 → Apps Script」貼上這整段程式碼，
 * 部署成「網頁應用程式」（執行身分：我；誰可以存取：所有人），
 * 再把部署網址填到 config.json 的 endpoint。
 *
 * 試算表本身維持不公開。這裡只回傳整理過的資料：姓名遮蔽中間字、同一人只留最後一次填寫，
 * 「原因」裡提到的「○○學長」也會遮蔽，原始回覆不會外流。
 */

const SHEET_NAME = '表單回覆 1';
const TZ = 'Asia/Taipei';
const CACHE_SECONDS = 60;

function doGet() {
  return json_(summary_());
}

function summary_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('summary');
  if (hit) return JSON.parse(hit);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sh.getDataRange().getValues();
  const head = values.shift().map(String);
  const col = key => head.findIndex(h => h.indexOf(key) >= 0);
  const c = {
    time: col('時間戳記'), name: col('姓名'), group: col('目前擔任'),
    stay: col('是否續留'), want: col('期望的職務'), reason: col('原因'),
  };

  const realNames = [];
  const byName = {};
  const days = {};
  let total = 0;
  values.forEach(r => {
    const name = String(r[c.name] || '').trim();
    const time = r[c.time];
    if (!name || !(time instanceof Date)) return;
    total++;
    const day = Utilities.formatDate(time, TZ, 'yyyy-MM-dd');
    days[day] = (days[day] || 0) + 1;
    const prev = byName[name];
    if (prev && prev.time > time) { prev.times++; return; }
    if (!prev) realNames.push(name);
    byName[name] = {
      time: time,
      times: prev ? prev.times + 1 : 1,
      group: String(r[c.group] || '').trim(),
      stay: String(r[c.stay] || '').trim(),
      want: String(r[c.want] || '').trim(),
      reason: String(r[c.reason] || '').trim(),
    };
  });

  const people = realNames.map(n => {
    const p = byName[n];
    return {
      name: mask_(n),
      group: p.group,
      stay: p.stay,
      want: p.want,
      reason: maskText_(p.reason, realNames),
      times: p.times,
      last: Utilities.formatDate(p.time, TZ, 'yyyy-MM-dd'),
    };
  });
  const timeline = Object.keys(days).sort().map(d => ({ date: d, n: days[d] }));

  const out = { ok: true, updated: new Date().toISOString(), total: total, people: people, timeline: timeline };
  cache.put('summary', JSON.stringify(out), CACHE_SECONDS);
  return out;
}

// 「巫宗翰」→「巫○翰」、「王明」→「王○」
function mask_(name) {
  const s = Array.from(name);
  if (s.length <= 1) return name;
  if (s.length === 2) return s[0] + '○';
  return s[0] + '○'.repeat(s.length - 2) + s[s.length - 1];
}

// 原因是自由填寫，遮蔽其中的幹部姓名與「○○學長／學姊」
function maskText_(text, realNames) {
  let t = text;
  realNames.forEach(n => { t = t.split(n).join(mask_(n)); });
  return t.replace(/([一-鿿])([一-鿿])(學長|學姊|學姐)/g, (m, a, b, suffix) =>
    (a + b === '輔導' || a + b === '各組') ? m : '○' + b + suffix);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 在編輯器按「執行」測試用：查看會回傳給網頁的資料
function 測試() {
  CacheService.getScriptCache().remove('summary');
  Logger.log(JSON.stringify(summary_(), null, 2));
}
