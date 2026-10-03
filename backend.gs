/**
 * Mbrmj "First Day Challenge" — Google Apps Script backend
 *
 * DEPLOY
 *  1. Create a Google Sheet, open Extensions → Apps Script, paste this file.
 *  2. Run setup() once (authorize). It creates Events, Sessions, Answers, Leads, Dashboard.
 *  3. Deploy → New deployment → type "Web app"
 *       Execute as: Me      ·      Who has access: Anyone
 *  4. Copy the /exec URL into CONFIG.APPS_SCRIPT_URL in index.html.
 *  After editing this file, use Deploy → Manage deployments → Edit → New version.
 *
 * The page sends POST bodies as text/plain (no CORS preflight):
 *   { type: "events", events: [ {ts, session_id, visitor_id, variant, event, step, data} ] }
 *   { type: "lead", name, whatsapp, email, consent, score, blind_spots, gaps, variant, ref, session_id, report_html }
 */

// Only needed if the script was created at script.google.com (not from inside the Sheet).
// Paste the Sheet ID: the long string in the Sheet URL between /d/ and /edit
var SHEET_ID = '';

function ss_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  if (!SHEET_ID) throw new Error('Set SHEET_ID at the top of backend.gs (or create the script from Extensions → Apps Script inside the Sheet).');
  return SpreadsheetApp.openById(SHEET_ID);
}

var HEADERS = {
  Events: ['ts', 'session_id', 'visitor_id', 'variant', 'event', 'step', 'data_json'],
  Sessions: ['session_id', 'visitor_id', 'first_seen', 'last_seen', 'variant', 'utm_source', 'utm_medium', 'utm_campaign', 'ref', 'device',
    'status', 'last_step', 'furthest_question', 'score', 'req', 'design', 'codebase', 'debug', 'test', 'total_time_sec', 'resumed', 'contact_left'],
  Answers: ['ts', 'session_id', 'variant', 'q', 'stage', 'type', 'selected_ids', 'missed_ids', 'extra_ids', 'correct', 'confidence', 'time_ms'],
  Leads: ['ts', 'session_id', 'name', 'whatsapp', 'email', 'consent', 'score', 'blind_spots', 'gaps', 'variant', 'ref']
};
var STATUS_ORDER = { viewed: 0, started: 1, in_progress: 2, completed: 3 };

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var body = JSON.parse(e.postData.contents);
    if (body.type === 'lead') handleLead_(body);
    else handleEvents_(body.events || (Array.isArray(body) ? body : []));
  } catch (err) {
    console.error(err);
  } finally {
    try { lock.releaseLock(); } catch (x) { /* not held */ }
  }
  return ContentService.createTextOutput('ok');
}

function doGet() { return ContentService.createTextOutput('mbrmj challenge backend is running'); }

function sheet_(name) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (HEADERS[name]) { sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold'); sh.setFrozenRows(1); }
  }
  return sh;
}

/* ---------------- events ---------------- */
function handleEvents_(events) {
  if (!events.length) return;
  var evRows = [], ansRows = [], bySession = {};
  events.forEach(function (ev) {
    var d = ev.data || {};
    evRows.push([ev.ts, ev.session_id, ev.visitor_id, ev.variant, ev.event, ev.step, JSON.stringify(d)]);
    if (ev.event === 'answer') {
      ansRows.push([ev.ts, ev.session_id, ev.variant, d.q, d.stage, d.type, (d.selected_ids || []).join(','), (d.missed_ids || []).join(','),
        (d.extra_ids || []).join(','), d.correct === true, d.confidence, d.time_ms]);
    }
    (bySession[ev.session_id] = bySession[ev.session_id] || []).push(ev);
  });
  append_('Events', evRows);
  if (ansRows.length) append_('Answers', ansRows);
  Object.keys(bySession).forEach(function (sid) { upsertSession_(sid, bySession[sid]); });
}

function append_(name, rows) {
  var sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function findRow_(sh, sid) {
  var cache = CacheService.getScriptCache(), key = 'row_' + sid, hit = cache.get(key);
  if (hit) {
    var r = Number(hit);
    if (sh.getRange(r, 1).getValue() === sid) return r;
  }
  var found = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).createTextFinder(sid).matchEntireCell(true).findNext();
  if (found) { cache.put(key, String(found.getRow()), 21600); return found.getRow(); }
  return 0;
}

function upsertSession_(sid, evs) {
  var sh = sheet_('Sessions'), cols = HEADERS.Sessions, row = findRow_(sh, sid), cur = {};
  if (row) {
    var vals = sh.getRange(row, 1, 1, cols.length).getValues()[0];
    cols.forEach(function (c, i) { cur[c] = vals[i]; });
  } else {
    cur = { session_id: sid, first_seen: evs[0].ts, status: 'viewed', furthest_question: 0, resumed: false, contact_left: false, total_time_sec: 0 };
  }
  evs.forEach(function (ev) {
    var d = ev.data || {};
    cur.visitor_id = cur.visitor_id || ev.visitor_id;
    cur.variant = cur.variant || ev.variant;
    cur.last_seen = ev.ts;
    if (ev.event === 'abandon') { cur.last_step = d.last_step || ev.step; }
    else cur.last_step = ev.step;
    if (ev.event === 'page_view') {
      ['utm_source', 'utm_medium', 'utm_campaign', 'ref'].forEach(function (k) { if (!cur[k]) cur[k] = d[k] || ''; });
      if (!cur.device) cur.device = d.device || '';
    }
    var st = { page_view: 'viewed', start: 'started', question_view: 'in_progress', answer: 'in_progress', complete: 'completed' }[ev.event];
    if (st && (STATUS_ORDER[st] > (STATUS_ORDER[cur.status] || 0))) cur.status = st;
    if ((ev.event === 'question_view' || ev.event === 'answer') && d.q > (cur.furthest_question || 0)) cur.furthest_question = d.q;
    if (ev.event === 'resume') cur.resumed = true;
    if (ev.event === 'complete') {
      cur.score = d.score;
      var s = d.stages || {};
      ['req', 'design', 'codebase', 'debug', 'test'].forEach(function (k) { cur[k] = s[k] || ''; });
      cur.furthest_question = 8;
    }
  });
  cur.total_time_sec = Math.max(0, Math.round((new Date(cur.last_seen) - new Date(cur.first_seen)) / 1000)) || 0;
  var out = cols.map(function (c) { return cur[c] === undefined ? '' : cur[c]; });
  if (row) sh.getRange(row, 1, 1, cols.length).setValues([out]);
  else {
    sh.appendRow(out);
    CacheService.getScriptCache().put('row_' + sid, String(sh.getLastRow()), 21600);
  }
}

/* ---------------- leads ---------------- */
function handleLead_(b) {
  append_('Leads', [[b.ts || new Date().toISOString(), b.session_id, b.name, "'" + b.whatsapp, b.email || '', b.consent === true, b.score, b.blind_spots || '', b.gaps || '', b.variant, b.ref || '']]);
  var sh = sheet_('Sessions'), row = findRow_(sh, b.session_id);
  if (row) sh.getRange(row, HEADERS.Sessions.indexOf('contact_left') + 1).setValue(true);
  if (b.email) {
    try {
      MailApp.sendEmail({ to: b.email, subject: 'تقريرك من تحدي مبرمج', htmlBody: buildEmail_(b), name: 'مبرمج' });
    } catch (err) { console.error('mail failed', err); }
  }
}

// Hosted copy of the white logo (publish logo-white.png next to index.html on GitHub Pages).
var EMAIL_LOGO = 'https://ahmadyacine.github.io/mbrmj-challenge/logo-white.png';
var EMAIL_FONT = "Tahoma,'Segoe UI',Arial,sans-serif";
var EMAIL_COL = {
  blind: { c: '#C8473A', bg: '#FBEDEA', tx: '#8E2D23' },
  gap: { c: '#D99A2B', bg: '#FCF4E2', tx: '#7A5410' },
  strong: { c: '#2F8F6B', bg: '#E9F5EF', tx: '#1D6247' }
};

function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

/** b.report = { name, score, rows:[{ar,st,label}], focus:{ar,st,label,m,n}|null, others:[{ar,st,label,m,n}],
 *               strong:[{ar,m}], story:{situation,did,result,tail}, plan:[[day,text]], closing, cta:{url,text} } */
function buildEmail_(b) {
  var r = b.report || {}, F = EMAIL_FONT;
  function stageCard(x, big) {
    var k = EMAIL_COL[x.st];
    return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 14px"><tr>' +
      '<td style="background:' + k.bg + ';border-right:4px solid ' + k.c + ';border-radius:10px;padding:' + (big ? '18px 20px' : '14px 18px') + ';font-family:' + F + ';text-align:right">' +
      '<div style="font-size:' + (big ? 19 : 17) + 'px;font-weight:bold;color:#1B3250;margin:0 0 8px">' + esc_(x.ar) +
      ' <span style="font-size:13px;font-weight:bold;color:' + k.tx + ';background:#ffffff;border-radius:99px;padding:2px 10px;margin-right:6px">' + esc_(x.label) + '</span></div>' +
      '<div style="font-size:15px;line-height:1.9;color:#0F1B2D;margin:0 0 6px"><b>وش يعني في الشغل:</b> ' + esc_(x.m) + '</div>' +
      '<div style="font-size:15px;line-height:1.9;color:#0F1B2D"><b>إيش تتعلم بعدها:</b> ' + esc_(x.n) + '</div></td></tr></table>';
  }
  function h2(t) { return '<div style="font-family:' + F + ';font-size:19px;font-weight:bold;color:#1B3250;margin:30px 0 12px;text-align:right">' + esc_(t) + '</div>'; }

  var rows = (r.rows || []).map(function (x) {
    var k = EMAIL_COL[x.st];
    return '<tr><td width="22" style="padding:9px 0;border-bottom:1px solid #EEEAE0"><span style="display:inline-block;width:14px;height:14px;border-radius:50%;background:' + k.c + '"></span></td>' +
      '<td style="padding:9px 8px;border-bottom:1px solid #EEEAE0;font-family:' + F + ';font-size:16px;color:#0F1B2D;text-align:right">' + esc_(x.ar) + '</td>' +
      '<td style="padding:9px 0;border-bottom:1px solid #EEEAE0;text-align:left"><span style="font-family:' + F + ';font-size:13px;font-weight:bold;color:' + k.tx + ';background:' + k.bg + ';border-radius:99px;padding:3px 12px">' + esc_(x.label) + '</span></td></tr>';
  }).join('');

  var body = '';
  body += '<div style="font-family:' + F + ';font-size:20px;font-weight:bold;color:#1B3250;margin:0 0 6px;text-align:right">أهلاً ' + esc_(r.name || b.name) + '،</div>' +
    '<div style="font-family:' + F + ';font-size:15px;line-height:1.9;color:#4B5563;margin:0 0 22px;text-align:right">خلصت مهمة سالم. هذي خريطتك: وين أنت جاهز، ووين تشتغل قبل أول وظيفة.</div>';
  body += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td style="background:#1B3250;border-radius:12px;padding:16px 20px;text-align:right">' +
    '<table role="presentation" width="100%"><tr><td style="font-family:' + F + ';font-size:15px;color:#C9D4E3;text-align:right">نتيجتك</td>' +
    '<td style="font-family:' + F + ';font-size:36px;font-weight:bold;color:#E8A33D;text-align:left">' + esc_(r.score) + '</td></tr></table></td></tr></table>';
  body += h2('خريطة الفجوات') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + rows + '</table>';
  if (r.focus) body += h2('أهم شي تشتغل عليه') + stageCard(r.focus, true);
  if (r.others && r.others.length) { body += h2('باقي الفجوات'); r.others.forEach(function (x) { body += stageCard(x, false); }); }
  if (r.strong && r.strong.length) {
    body += h2('نقاط قوتك');
    r.strong.forEach(function (x) {
      body += '<div style="font-family:' + F + ';font-size:15px;line-height:1.9;color:#0F1B2D;margin:0 0 8px;text-align:right"><b style="color:#1D6247">✓ ' + esc_(x.ar) + '</b> — ' + esc_(x.m) + '</div>';
    });
  }
  if (r.story) {
    var st = r.story;
    body += h2('قصتك للمقابلة') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td style="background:#1B3250;border-radius:12px;padding:18px 20px;font-family:' + F + ';font-size:15px;line-height:1.9;color:#EAF0F7;text-align:right">' +
      '<div style="margin:0 0 10px"><b style="color:#E8A33D">الموقف:</b> ' + esc_(st.situation) + '</div>' +
      '<div style="margin:0 0 10px"><b style="color:#E8A33D">إيش سويت:</b> ' + esc_(st.did) + '</div>' +
      '<div style="margin:0 0 12px"><b style="color:#E8A33D">النتيجة:</b> ' + esc_(st.result) + '</div>' +
      '<div style="font-weight:bold;color:#ffffff">' + esc_(st.tail) + '</div></td></tr></table>';
  }
  if (r.plan && r.plan.length) {
    body += h2('خطة ٧ أيام') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + r.plan.map(function (p) {
      return '<tr><td width="70" valign="top" style="padding:9px 0;border-bottom:1px solid #EEEAE0;font-family:' + F + ';font-size:15px;font-weight:bold;color:#1B3250;text-align:right">' + esc_(p[0]) + '</td>' +
        '<td style="padding:9px 8px;border-bottom:1px solid #EEEAE0;font-family:' + F + ';font-size:15px;line-height:1.8;color:#0F1B2D;text-align:right">' + esc_(p[1]) + '</td></tr>';
    }).join('') + '</table>';
  }
  if (r.checklist && r.checklist.url) {
    body += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:26px 0 0"><tr><td style="border:1.5px solid #1B3250;border-radius:10px;padding:14px 18px;text-align:right;font-family:' + F + ';font-size:15px;line-height:1.8;color:#0F1B2D">' +
      '<b style="color:#1B3250">قائمة مرفقة لك:</b> ما تبدأ أي مهمة حقيقية بدونها.<br><a href="' + esc_(r.checklist.url) + '" style="color:#1B3250;font-weight:bold">' + esc_(r.checklist.text) + '</a></td></tr></table>';
  }
  if (r.closing) {
    body += '<div style="font-family:' + F + ';font-size:16px;line-height:2;color:#0F1B2D;margin:30px 0 18px;padding-top:22px;border-top:2px solid #1B3250;text-align:right">' + esc_(r.closing) + '</div>';
    if (r.cta && r.cta.url) {
      body += '<table role="presentation" cellpadding="0" cellspacing="0" align="right"><tr><td style="background:#E8A33D;border-radius:10px"><a href="' + esc_(r.cta.url) + '" style="display:inline-block;padding:14px 26px;font-family:' + F + ';font-size:16px;font-weight:bold;color:#0F1B2D;text-decoration:none">' + esc_(r.cta.text) + '</a></td></tr></table><div style="clear:both"></div>';
    }
  }

  return '<!doctype html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light only"><title>تقريرك من تحدي مبرمج</title></head>' +
    '<body style="margin:0;padding:0;background:#F6F4EF"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F6F4EF"><tr><td align="center" style="padding:24px 12px">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" dir="rtl" style="width:100%;max-width:600px">' +
    '<tr><td style="background:#1B3250;border-radius:14px 14px 0 0;padding:24px 28px 20px;text-align:right">' +
    '<img src="' + EMAIL_LOGO + '" width="104" alt="مبرمج" style="display:block;border:0;height:auto;margin:0 0 14px">' +
    '<div style="font-family:Consolas,Menlo,monospace;font-size:13px;letter-spacing:1px;color:#E8A33D;direction:ltr;text-align:right">TASK #001 · DONE</div></td></tr>' +
    '<tr><td style="background:#E8A33D;height:4px;line-height:4px;font-size:0">&nbsp;</td></tr>' +
    '<tr><td style="background:#ffffff;border-radius:0 0 14px 14px;padding:28px 28px 30px">' + body + '</td></tr>' +
    '<tr><td style="padding:18px 8px 0;font-family:' + F + ';font-size:12px;line-height:1.8;color:#8A8F99;text-align:center">© مبرمج · بياناتك تستخدم فقط للتواصل معك بخصوص التقرير</td></tr>' +
    '</table></td></tr></table></body></html>';
}

/* ---------------- setup + dashboard ---------------- */
function setup() {
  Object.keys(HEADERS).forEach(function (n) { sheet_(n); });
  buildDashboard_();
  var d = ss_().getSheetByName('Sheet1');
  if (d && d.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(d);
}

var QUESTION_META = [
  { n: 1, stage: 'req', opts: { a: 1, b: 1, c: 0, d: 1, e: 0 } },
  { n: 2, stage: 'req', opts: { a: 1, b: 0, c: 1, d: 0, e: 1 } },
  { n: 3, stage: 'design', opts: { a: 0, b: 1, c: 0, d: 0 } },
  { n: 4, stage: 'codebase', opts: { a: 0, b: 1, c: 0, d: 0 } },
  { n: 5, stage: 'debug', opts: { a: 1, b: 0, c: 1, d: 0, e: 0 } },
  { n: 6, stage: 'debug', opts: { a: 1, b: 0, c: 0, d: 0 } },
  { n: 7, stage: 'design', opts: { a: 1, b: 0, c: 1, d: 0, e: 0 } },
  { n: 8, stage: 'test', opts: { a: 1, b: 1, c: 1, d: 0, e: 0 } }
];

function buildDashboard_() {
  var ss = ss_(), sh = ss.getSheetByName('Dashboard');
  if (sh) sh.clear(); else sh = ss.insertSheet('Dashboard');
  var crit = ['*', 'a', 'b', 'c'], colL = ['B', 'C', 'D', 'E'];
  var S = 'Sessions!', A = 'Answers!';
  var row = 1;
  function title(t) { sh.getRange(row, 1).setValue(t).setFontWeight('bold').setFontSize(13).setBackground('#1B3250').setFontColor('#ffffff'); sh.getRange(row, 2, 1, 4).setBackground('#1B3250'); row++; }
  function head() {
    sh.getRange(row, 1, 1, 5).setValues([['المقياس', 'الكل', 'نسخة a', 'نسخة b', 'نسخة c']]).setFontWeight('bold').setBackground('#EDE9DE');
    sh.getRange(row + 1, 1, 1, 5).setValues([['(فلتر النسخة)', '*', 'a', 'b', 'c']]).setFontColor('#888888');
    row += 2; return row - 1; // criteria row
  }
  function put(label, fn, fmtStr) {
    var r = row, f = [label];
    colL.forEach(function (c) { f.push(fn(c + '$' + critRow)); });
    sh.getRange(r, 1).setValue(label);
    sh.getRange(r, 2, 1, 4).setFormulas([f.slice(1)]);
    if (fmtStr) sh.getRange(r, 2, 1, 4).setNumberFormat(fmtStr);
    row++;
  }
  var critRow;
  // ---- Funnel
  title('القمع (Funnel) — جلسات'); critRow = head();
  put('زيارات (جلسات)', function (c) { return '=COUNTIFS(' + S + 'E2:E,' + c + ')'; });
  put('بدأوا التحدي', function (c) { return '=COUNTIFS(' + S + 'E2:E,' + c + ',' + S + 'K2:K,"<>viewed")'; });
  for (var n = 1; n <= 8; n++) (function (n) {
    put('وصلوا للسؤال ' + n, function (c) { return '=COUNTIFS(' + S + 'E2:E,' + c + ',' + S + 'M2:M,">=' + n + '")'; });
  })(n);
  put('أكملوا (النتيجة)', function (c) { return '=COUNTIFS(' + S + 'E2:E,' + c + ',' + S + 'K2:K,"completed")'; });
  put('تركوا بياناتهم', function (c) { return '=COUNTIFS(' + S + 'E2:E,' + c + ',' + S + 'V2:V,TRUE)'; });
  put('ضغطات المشاركة', function (c) { return '=COUNTIFS(Events!D2:D,' + c + ',Events!E2:E,"share_click")'; });
  put('نسبة الإكمال', function (c) { return '=IFERROR(COUNTIFS(' + S + 'E2:E,' + c + ',' + S + 'K2:K,"completed")/COUNTIFS(' + S + 'E2:E,' + c + ',' + S + 'K2:K,"<>viewed"),0)'; }, '0%');
  row++;
  // ---- Drop-off
  title('التسرب لكل سؤال (آخر سؤال وصلوا له ولم يكملوا)'); critRow = head();
  for (n = 1; n <= 8; n++) (function (n) {
    put('توقفوا عند السؤال ' + n, function (c) { return '=COUNTIFS(' + S + 'E2:E,' + c + ',' + S + 'M2:M,' + n + ',' + S + 'K2:K,"<>completed")'; });
  })(n);
  row++;
  // ---- Per-question quality
  title('أداء الأسئلة'); critRow = head();
  for (n = 1; n <= 8; n++) (function (n) {
    put('Q' + n + ' — % صح', function (c) { return '=IFERROR(COUNTIFS(' + A + 'C2:C,' + c + ',' + A + 'D2:D,' + n + ',' + A + 'J2:J,TRUE)/COUNTIFS(' + A + 'C2:C,' + c + ',' + A + 'D2:D,' + n + '),"")'; }, '0%');
  })(n);
  for (n = 1; n <= 8; n++) (function (n) {
    put('Q' + n + ' — % متأكد وغلط (blind)', function (c) { return '=IFERROR(COUNTIFS(' + A + 'C2:C,' + c + ',' + A + 'D2:D,' + n + ',' + A + 'J2:J,FALSE,' + A + 'K2:K,"sure")/COUNTIFS(' + A + 'C2:C,' + c + ',' + A + 'D2:D,' + n + '),"")'; }, '0%');
  })(n);
  for (n = 1; n <= 8; n++) (function (n) {
    put('Q' + n + ' — متوسط الوقت (ثانية)', function (c) { return '=IFERROR(AVERAGEIFS(' + A + 'L2:L,' + A + 'C2:C,' + c + ',' + A + 'D2:D,' + n + ')/1000,"")'; }, '0.0');
  })(n);
  row++;
  // ---- Stage blind spots
  title('تكرار حالات كل مرحلة'); critRow = head();
  [['req', 'O', 'تحليل المتطلبات'], ['design', 'P', 'التصميم'], ['codebase', 'Q', 'الكود الموجود'], ['debug', 'R', 'التصحيح'], ['test', 'S', 'الاختبار']].forEach(function (s) {
    ['blind', 'gap', 'strong'].forEach(function (st) {
      put(s[2] + ' — ' + st, function (c) { return '=COUNTIFS(' + S + 'E2:E,' + c + ',' + S + s[1] + '2:' + s[1] + ',"' + st + '")'; });
    });
  });
  row++;
  // ---- Trap rate per option
  title('معدل الفخ لكل خيار (% من اللي اختاروه) — الخيارات الغلط هي الفخاخ'); critRow = head();
  QUESTION_META.forEach(function (q) {
    Object.keys(q.opts).forEach(function (id) {
      (function (q, id) {
        put('Q' + q.n + ' خيار ' + id + (q.opts[id] ? ' (صح)' : ' (فخ)'), function (c) {
          return '=IFERROR(COUNTIFS(' + A + 'C2:C,' + c + ',' + A + 'D2:D,' + q.n + ',' + A + 'G2:G,"*' + id + '*")/COUNTIFS(' + A + 'C2:C,' + c + ',' + A + 'D2:D,' + q.n + '),"")';
        }, '0%');
      })(q, id);
    });
  });
  sh.setColumnWidth(1, 300); sh.setColumnWidths(2, 4, 90);
  sh.setRightToLeft(true); sh.setFrozenRows(0);
  ss.setActiveSheet(sh);
}
