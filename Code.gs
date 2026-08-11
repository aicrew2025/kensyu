/**
 * aicrew 生成AI研修 — 申込受付エンジン（support@aicrew.jp 名義で実行）
 *  ① シート記録 ② 回ごとのMeet予定を自動作成（REST API・拡張サービス不要）
 *  ③ 受講者を自動招待（非表示・通知なし） ④ 確認メール（Gmail・Meet URL記載・.ics添付）
 *
 * ※ 拡張サービスは不要。カレンダーは REST API（UrlFetchApp + OAuthトークン）で操作します。
 *   初回デプロイ/実行時に「カレンダー」権限の承認を求められたら許可してください。
 */

const SHEET_ID = '1ERVjiXnihNJfPsVb0vOTR4ECrDmaorZVa79hM_bUZdk';
const SHEET_GID = 1428018104;
const MEET_SHEET = 'Meet URL一覧';
const REMINDER_SHEET = '_リマインド';   // リマインド用インデックス（前日16:00送信）
const DETAIL_SHEET = '受講記録';        // 社労士連携：1人×1回=1行の明細（出勤簿・助成金証憑向き）
const SUMMARY_SHEET = '受講者サマリ';   // 社労士連携：1人1行の集計（合計時間・出席回数）
const SUPPORT_EMAIL = 'support@aicrew.jp';   // 問い合わせ窓口（日程変更・欠席連絡・トラブル・研修内容の質問）
const FROM_NAME = 'aicrew 研修事務局';
const RECDIST_SHEET = '_録画配信';           // 録画の自動配信ログ（公開日・視聴期限・配信先）
const RECORD_FOLDERS = ['Meet Recordings', 'Meet の録画'];  // 録画の保存先フォルダ名（言語設定による揺れに対応）
const VIEW_MONTHS = 2;                        // 録画の視聴期間（公開日から2ヶ月）
// Supabase キープアライブ（無料プランの自動一時停止を防ぐ。キーは公開キーで index.html と同一）
const SB_REST = 'https://gfbedmzdgjpczpfadaft.supabase.co/rest/v1/registrations?select=id&limit=1';
const SB_KEY = 'sb_publishable_ECc0koefc0NFUeXwyT2U1w_BW6KExj4';
const SB_DASHBOARD = 'https://supabase.com/dashboard/project/gfbedmzdgjpczpfadaft';
// 出席の自動記録（Meet監査ログ）
const ATTLOG_SHEET = '_入室ログ';        // 入退室の生ログ（証憑・手動確認用）
const ATTEND_MIN_MINUTES = 100;          // この分数以上の滞在で「出席」（2時間講義の目安。調整可）
const CIR = ['①', '②', '③', '④', '⑤'];
const COURSE_JP = { video: 'AI動画', sales: '営業', bo: 'BO' };

// 権限承認用：エディタでこの関数を一度「実行」し、カレンダー＋外部リクエスト＋Drive等を許可してください
// ※ 自動録画（Meet API）を使うため、appsscript.json（マニフェスト）に oauthScopes の追記が必要です（反映手順書参照）
function authorize() {
  CalendarApp.getDefaultCalendar(); // カレンダー権限
  DriveApp.getRootFolder();         // Drive権限（録画の自動配信に使用）
  UrlFetchApp.fetch('https://www.googleapis.com/calendar/v3/users/me/calendarList?maxResults=1',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }); // 外部リクエスト権限
  return 'authorized';
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const p = JSON.parse(e.postData.contents);
    let meetErr = null;
    try { ensureMeetUrls_(p); } catch (me) { meetErr = String(me); } // Meet失敗でもメールは送る
    appendRow_(p);
    try { appendManagement_(p); } catch (mge) { } // 社労士連携：明細＋サマリに記録
    try { logReminders_(p); } catch (re) { } // リマインド用インデックスに記録
    try { notifyAdmin_(p); } catch (ae) { } // 事務局へ申込通知（タイムリーに把握）
    sendMail_(p);
    return json_({ ok: true, meetErr: meetErr });
  } catch (err) {
    try { alertAdminError_(err, e); } catch (_) { } // 失敗時に事務局へアラート
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json_({ ok: true, service: 'aicrew-kensyu', ts: new Date().toISOString() }); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function md_(d) { return d.slice(5).replace('-', '/'); }
function jpDate_(ymd) { const a = String(ymd).split('-'); return a[0] + '年' + Number(a[1]) + '月' + Number(a[2]) + '日'; }
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function splitTime_(t) { const a = String(t).split(/[〜~]/); return [a[0].trim(), (a[1] || a[0]).trim()]; }
function hours_(time) { const [st, en] = splitTime_(time); const a = st.split(':').map(Number), b = en.split(':').map(Number); return ((b[0] * 60 + b[1]) - (a[0] * 60 + a[1])) / 60; }

function slotKey_(course, n, date) {
  if ((course === 'sales' || course === 'bo') && n <= 3) return 'common|' + n + '|' + date;
  return course + '|' + n + '|' + date;
}
function targetLabel_(course, n) {
  if ((course === 'sales' || course === 'bo') && n <= 3) return '営業・BO共通';
  return COURSE_JP[course] || course;
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}

// ---- カレンダー REST ヘルパー（拡張サービス不要） ----
function calFetch_(method, path, payload) {
  const opt = {
    method: method,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    contentType: 'application/json',
    muteHttpExceptions: true
  };
  if (payload) opt.payload = JSON.stringify(payload);
  const res = UrlFetchApp.fetch('https://www.googleapis.com/calendar/v3' + path, opt);
  const code = res.getResponseCode(), txt = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('cal_api ' + code + ': ' + txt.slice(0, 300));
  return JSON.parse(txt);
}

function createMeetEvent_(summary, date, st, en, desc) {
  const ev = {
    summary: summary, description: desc,
    start: { dateTime: date + 'T' + st + ':00', timeZone: 'Asia/Tokyo' },
    end: { dateTime: date + 'T' + en + ':00', timeZone: 'Asia/Tokyo' },
    guestsCanSeeOtherGuests: false, guestsCanInviteOthers: false,
    conferenceData: { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } }
  };
  const c = calFetch_('post', '/calendars/primary/events?conferenceDataVersion=1&sendUpdates=none', ev);
  let url = c.hangoutLink || '';
  if (!url && c.conferenceData && c.conferenceData.entryPoints) {
    const vp = c.conferenceData.entryPoints.filter(function (x) { return x.entryPointType === 'video'; })[0];
    if (vp) url = vp.uri;
  }
  return { eventId: c.id, meetUrl: url };
}

// ---- 自動録画（Google Meet REST API）----
// 主催アカウント（録画権限のある人）が入室した時点で録画が自動で始まるように、Meetスペースに設定を入れる。
function meetCodeFromUrl_(url) { const m = String(url).match(/meet\.google\.com\/([a-z0-9\-]+)/i); return m ? m[1] : ''; }
function enableAutoRecording_(meetUrl) {
  const code = meetCodeFromUrl_(meetUrl);
  if (!code) throw new Error('meet code not found: ' + meetUrl);
  const res = UrlFetchApp.fetch(
    'https://meet.googleapis.com/v2/spaces/' + code + '?updateMask=config.artifactConfig.recordingConfig.autoRecordingGeneration',
    {
      method: 'patch',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      contentType: 'application/json',
      payload: JSON.stringify({ config: { artifactConfig: { recordingConfig: { autoRecordingGeneration: 'ON' } } } }),
      muteHttpExceptions: true
    });
  const c = res.getResponseCode();
  if (c < 200 || c >= 300) throw new Error('meet_api ' + c + ': ' + res.getContentText().slice(0, 200));
}

function addGuest_(eventId, email) {
  const ev = calFetch_('get', '/calendars/primary/events/' + encodeURIComponent(eventId));
  const at = ev.attendees || [];
  for (let i = 0; i < at.length; i++) {
    if (String(at[i].email || '').toLowerCase() === String(email).toLowerCase()) return;
  }
  at.push({ email: email });
  calFetch_('patch', '/calendars/primary/events/' + encodeURIComponent(eventId) + '?sendUpdates=none', { attendees: at });
}

// 各回のMeet URLを用意（作成 or 再利用）し、受講者を招待
function ensureMeetUrls_(p) {
  const sh = ensureSheet_(MEET_SHEET, ['対象', '回', '日付', '曜日', '時間', 'テーマ', 'Meet URL', 'eventId', 'slotKey', '自動録画']);
  const data = sh.getDataRange().getValues();
  const idx = {};
  for (let r = 1; r < data.length; r++) { const key = data[r][8]; if (key) idx[key] = { url: data[r][6], eventId: data[r][7] }; }
  p.sessions.forEach(function (o) {
    const key = slotKey_(p.course, o.n, o.date);
    let slot = idx[key];
    if (!slot || !slot.url) {
      const tl = targetLabel_(p.course, o.n);
      const [st, en] = splitTime_(o.time);
      const summary = 'aicrew 生成AI研修 ' + CIR[o.n - 1] + ' ' + o.theme + '（' + tl + '）';
      const made = createMeetEvent_(summary, o.date, st, en, 'aicrew 生成AI研修／オンライン（Google Meet）');
      // 自動録画を設定（失敗しても申込処理は止めない。NGの場合は講師が手動で録画開始）
      let autoRec = 'OK';
      try { enableAutoRecording_(made.meetUrl); } catch (arErr) { autoRec = 'NG: ' + String(arErr).slice(0, 120); }
      sh.appendRow([tl, CIR[o.n - 1], md_(o.date), o.wd, o.time, o.theme, made.meetUrl, made.eventId, key, autoRec]);
      slot = { url: made.meetUrl, eventId: made.eventId };
      idx[key] = slot;
    }
    o.meetUrl = slot.url;
    try { addGuest_(slot.eventId, p.email); } catch (err) { /* 招待失敗は無視 */ }
  });
}

function appendRow_(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheets().filter(s => s.getSheetId() === SHEET_GID)[0] || ss.getSheets()[0];
  const s = p.sessions;
  const cells = s.map(o => md_(o.date) + '(' + o.wd + ') ' + o.time + ' ' + o.theme);
  sh.appendRow([new Date(), p.company, p.name, p.kana, p.email, p.course_label,
    cells[0], cells[1], cells[2], cells[3], cells[4], p.sid || '']);
}

// ---- 社労士連携：管理シート一元化（明細＋サマリ） ----
// 助成金（人材開発支援助成金 等）の証憑として、1人×1回=1行の明細と、1人1行の集計を同じスプレッドシートに自動生成。
function appendManagement_(p) {
  const now = new Date();

  // ① 受講記録（明細）：出勤簿・受講記録として提出しやすい1行=1回の形。「出席」列は当日に事務局が手入力。
  const detail = ensureSheet_(DETAIL_SHEET,
    ['申込日時', '会社', '氏名', 'カナ', 'メール', 'コース', '回', '対象',
      '日付', '曜日', '時間', '開始', '終了', '時間数(h)', 'テーマ', '出席', 'Meet URL', '申込ID']);
  let total = 0;
  const dates = [];
  p.sessions.forEach(function (o) {
    const [st, en] = splitTime_(o.time);
    const h = hours_(o.time);
    total += h; dates.push(o.date);
    detail.appendRow([now, p.company, p.name, p.kana, p.email, p.course_label,
      CIR[o.n - 1], targetLabel_(p.course, o.n), o.date, o.wd, o.time, st, en, h, o.theme, '', o.meetUrl || '', p.sid || '']);
  });

  // ② 受講者サマリ：1人1行。合計時間・予定回数は申込内容から、出席回数は明細の「出席」入力を数式で自動集計（メール一致）。
  const summary = ensureSheet_(SUMMARY_SHEET,
    ['申込日時', '会社', '氏名', 'カナ', 'メール', 'コース', '合計時間(h)', '予定回数', '出席回数', '初回日', '最終日', '申込ID']);
  dates.sort();
  const row = summary.getLastRow() + 1;
  summary.appendRow([now, p.company, p.name, p.kana, p.email, p.course_label,
    total, p.sessions.length, '', dates[0], dates[dates.length - 1], p.sid || '']);
  // 出席回数 = 明細(メール=自分 かつ 出席="出席")の件数。明細のメール列=E、出席列=Q。
  summary.getRange(row, 9).setFormula(
    "=COUNTIFS('" + DETAIL_SHEET + "'!$E:$E,$E" + row + ",'" + DETAIL_SHEET + "'!$Q:$Q,\"出席\")");
}

// 申込処理が失敗したとき、事務局へエラーアラート（取りこぼし防止）
function alertAdminError_(err, e) {
  let body = '申込の処理中にエラーが発生しました。下記をご確認のうえ、必要に応じて手動対応をお願いします。\n\n';
  body += '■ エラー内容\n' + String(err) + '\n\n';
  try { body += '■ 受信データ\n' + (e && e.postData ? e.postData.contents : '(取得できませんでした)') + '\n'; } catch (_) { }
  GmailApp.sendEmail(SUPPORT_EMAIL, '【申込エラー】処理に失敗しました（要確認）', body, { name: FROM_NAME });
}

// 事務局へ申込通知（顧客の申込状況をタイムリーに把握できるように）
function notifyAdmin_(p) {
  const rows = p.sessions.map(function (o) {
    return CIR[o.n - 1] + ' ' + o.date + '(' + o.wd + ') ' + o.time + '　' + o.theme;
  }).join('<br>');
  const html = '<div style="font-family:sans-serif;font-size:13px;line-height:1.6">' +
    '<p><b>新しい申込みを受け付けました。</b></p>' +
    '<p>会社：' + esc_(p.company) + '<br>' +
    '氏名：' + esc_(p.name) + '（' + esc_(p.kana) + '）<br>' +
    'メール：' + esc_(p.email) + '<br>' +
    'コース：' + esc_(p.course_label) + '</p>' +
    '<p>' + rows + '</p>' +
    '<p style="color:#888;font-size:12px">※ スプレッドシート（受講記録・受講者サマリ）にも自動記録済みです。</p></div>';
  GmailApp.sendEmail(SUPPORT_EMAIL,
    '【申込通知】' + p.company + ' ' + p.name + '様（' + p.course_label + '）',
    '新しい申込みがありました。HTMLメール対応のメーラーでご覧ください。',
    { name: FROM_NAME, htmlBody: html });
}

function sendMail_(p) {
  const s = p.sessions;
  const rows = s.map(function (o, i) {
    const meet = o.meetUrl ? '<a href="' + esc_(o.meetUrl) + '">Meetを開く</a>' : '追ってご案内';
    return '<tr>' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px"><b>' + CIR[i] + '</b> ' + esc_(o.theme) + '</td>' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px">' + esc_(md_(o.date)) + '(' + esc_(o.wd) + ')</td>' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px">' + esc_(o.time) + '</td>' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px">' + meet + '</td></tr>';
  }).join('');
  const html =
    '<div style="font-family:sans-serif;color:#1a1a1a;line-height:1.6">' +
    '<h2 style="color:#0f6e56">お申込みを受け付けました</h2>' +
    '<p>' + esc_(p.company) + '／<b>' + esc_(p.name) + '</b> 様<br>' + esc_(p.course_label) + '（全5回・計10.0時間／2ヶ月以内）</p>' +
    '<table style="border-collapse:collapse;font-size:13px">' +
    '<tr style="background:#f0fbf6">' +
    '<td style="border:1px solid #d7ece2;padding:6px 8px">研修内容</td>' +
    '<td style="border:1px solid #d7ece2;padding:6px 8px">日付</td>' +
    '<td style="border:1px solid #d7ece2;padding:6px 8px">時間</td>' +
    '<td style="border:1px solid #d7ece2;padding:6px 8px">参加</td></tr>' + rows + '</table>' +
    '<p style="font-size:13px">各回の <b>Google Meet</b> はリンクをクリックするだけで参加できます（アプリ不要・ブラウザでOK・Windows / Mac / スマホ対応）。</p>' +
    '<p style="font-size:13px">添付の <b>.ics</b> ファイルは、<b>Outlook や Apple カレンダー等に5回分をまとめて登録</b>するためのものです。ファイルを開いて「追加／インポート」を選ぶと登録できます（Googleカレンダーをお使いの方は招待が自動で届きます）。</p>' +
    '<p style="font-size:13px">複数人で1つの画面でご覧になる場合も、<b>必ずお一人ずつ、ご自身の端末からご参加ください</b>（マイク・カメラはオフのままでOKです）。出席を個人単位で記録するためのお願いです。</p>' +
    '<p style="font-size:13px">各回とも<b>開始5分前から入室いただけます</b>。出席記録のため、<b>開始5分前までのご入室</b>にご協力ください。</p>' +
    '<p style="font-size:12.5px;background:#fff8e6;border:1px solid #f0d98a;border-radius:8px;padding:10px 12px;color:#7a5c00">' +
    '【ご注意：受講日の変更・欠席について】' +
    '<b style="color:#c0392b">無断欠席・遅刻・早退は、助成金の受給に影響が出る場合があります。</b><br>' +
    '変更の際は研修サポート事務局 <b>' + SUPPORT_EMAIL + '</b> までお早めにご連絡ください。</p>' +
    '<p style="font-size:12.5px;background:#eef2fb;border:1px solid #d3def3;border-radius:8px;padding:10px 12px;color:#2a3556">' +
    '【録画について】各回は録画し、講義終了後に復習用として<b>自動でメールにてご案内</b>します（<b>各録画の公開日から2ヶ月間</b>視聴可能）。' +
    '録画は受講者ご本人の復習用です。第三者への送信・共有・SNS等への転載はご遠慮ください。</p></div>';
  GmailApp.sendEmail(p.email, '【aicrew 生成AI研修】お申込みを受け付けました（Meet URLのご案内）',
    'HTMLメール対応のメーラーでご覧ください。',
    { name: FROM_NAME, htmlBody: html, attachments: [Utilities.newBlob(buildIcs_(p), 'text/calendar', 'aicrew-kensyu.ics')] });
}

// ---- リマインド（前日16:00） ----
function logReminders_(p) {
  const sh = ensureSheet_(REMINDER_SHEET, ['日付', 'メール', '氏名', 'コース', 'テーマ', '時間', 'Meet URL']);
  p.sessions.forEach(function (o) {
    sh.appendRow([o.date, p.email, p.name, p.course_label, o.theme, o.time, o.meetUrl || '']);
  });
}

// 毎日16:00トリガーで実行 → 翌日開催分のリマインドを送信
function sendReminders() {
  const tz = 'Asia/Tokyo';
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 86400000), tz, 'yyyy-MM-dd');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(REMINDER_SHEET);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  const byEmail = {};
  for (let r = 1; r < data.length; r++) {
    const dc = data[r][0];
    const ds = (dc instanceof Date) ? Utilities.formatDate(dc, tz, 'yyyy-MM-dd') : String(dc).slice(0, 10);
    if (ds !== tomorrow) continue;
    const email = data[r][1];
    (byEmail[email] = byEmail[email] || { name: data[r][2], items: [] }).items.push(
      { course: data[r][3], theme: data[r][4], time: data[r][5], url: data[r][6] });
  }
  Object.keys(byEmail).forEach(function (email) {
    const g = byEmail[email];
    const rows = g.items.map(function (it) {
      const meet = it.url ? '<a href="' + esc_(it.url) + '">Meetを開く</a>' : '別途ご案内';
      return '<tr><td style="border:1px solid #d7ece2;padding:6px 8px">' + esc_(it.course) + '</td>' +
        '<td style="border:1px solid #d7ece2;padding:6px 8px">' + esc_(it.theme) + '</td>' +
        '<td style="border:1px solid #d7ece2;padding:6px 8px">' + esc_(it.time) + '</td>' +
        '<td style="border:1px solid #d7ece2;padding:6px 8px">' + meet + '</td></tr>';
    }).join('');
    const html = '<div style="font-family:sans-serif;color:#1a1a1a;line-height:1.6">' +
      '<h2 style="color:#0f6e56">明日の研修リマインド</h2>' +
      '<p><b>' + esc_(g.name) + '</b> 様<br>明日の研修のご案内です。<b>開始5分前から入室いただけます</b>。出席記録のため、開始5分前までのご入室にご協力ください。</p>' +
      '<table style="border-collapse:collapse;font-size:13px"><tr style="background:#f0fbf6">' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px">コース</td>' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px">テーマ</td>' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px">時間</td>' +
      '<td style="border:1px solid #d7ece2;padding:6px 8px">参加</td></tr>' + rows + '</table>' +
      '<p style="font-size:12.5px;color:#777">複数人で受講の場合も、お一人ずつご自身の端末からご参加ください。<br>' +
      '当日の録画は、講義終了後に復習用として自動でメールにてご案内します（公開日から2ヶ月間視聴可・第三者への共有はご遠慮ください）。</p></div>';
    GmailApp.sendEmail(email, '【aicrew 生成AI研修】明日の研修リマインド', '明日の研修のご案内です。',
      { name: FROM_NAME, htmlBody: html });
  });
}

// 一度だけ実行：毎日16:00の自動リマインドを設定
function setupReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(16).everyDays(1).inTimezone('Asia/Tokyo').create();
  return 'reminder trigger set (daily 16:00 JST)';
}

// ---- 録画の自動配信（講義終了後、そのコマの登録者へ閲覧権限＋案内メール）----

// Drive REST ヘルパー
function driveFetch_(method, path, payload) {
  const opt = {
    method: method,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    contentType: 'application/json',
    muteHttpExceptions: true
  };
  if (payload) opt.payload = JSON.stringify(payload);
  const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3' + path, opt);
  const code = res.getResponseCode(), txt = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('drive_api ' + code + ': ' + txt.slice(0, 200));
  return txt ? JSON.parse(txt) : {};
}

// 「Meet URL一覧」からコマ情報を読み込む（キー＝Meet URL）
function loadSlotsByUrl_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(MEET_SHEET);
  const map = {};
  if (!sh) return map;
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const url = String(data[r][6] || '').trim();
    if (!url) continue;
    const key = String(data[r][8] || '');           // 例 video|3|2026-08-19
    map[url] = {
      target: data[r][0], kai: data[r][1], theme: data[r][5],
      url: url, slotKey: key, date: key.split('|')[2] || ''
    };
  }
  return map;
}

// _リマインドから「そのコマ（日付×Meet URL）」の登録者を取得
function slotRecipients_(date, meetUrl) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(REMINDER_SHEET);
  const out = [];
  if (!sh) return out;
  const tz = 'Asia/Tokyo';
  const data = sh.getDataRange().getValues();
  const seen = {};
  for (let r = 1; r < data.length; r++) {
    const dc = data[r][0];
    const ds = (dc instanceof Date) ? Utilities.formatDate(dc, tz, 'yyyy-MM-dd') : String(dc).slice(0, 10);
    if (ds !== date) continue;
    if (String(data[r][6] || '').trim() !== meetUrl) continue;
    const email = String(data[r][1] || '').trim().toLowerCase();
    if (!email || seen[email]) continue;
    seen[email] = true;
    out.push({ email: email, name: data[r][2] });
  }
  return out;
}

// 毎時トリガー：Meet Recordings フォルダを走査し、新しい録画を該当コマの登録者へ配信
function distributeRecordings() {
  const tz = 'Asia/Tokyo';
  const log = ensureSheet_(RECDIST_SHEET,
    ['公開日', '視聴期限', 'fileId', 'ファイル名', 'slotKey', '日付', '配信数', '状態', '配信先']);
  const logData = log.getDataRange().getValues();
  const done = {};
  for (let r = 1; r < logData.length; r++) { if (logData[r][2]) done[logData[r][2]] = true; }

  const slots = loadSlotsByUrl_();
  const now = new Date();
  const weekAgo = new Date(now.getTime() - 7 * 86400000);  // 初回導入時に過去の古いファイルを拾わない保険

  RECORD_FOLDERS.forEach(function (fname) {
    const folders = DriveApp.getFoldersByName(fname);
    while (folders.hasNext()) {
      const files = folders.next().getFiles();
      while (files.hasNext()) {
        const f = files.next();
        try {
          if (done[f.getId()]) continue;
          if (String(f.getMimeType()).indexOf('video/') !== 0) continue;  // 動画のみ（チャットログ等は除外）
          if (f.getDateCreated() < weekAgo) continue;
          processRecording_(f, slots, log, tz);
          done[f.getId()] = true;
        } catch (err) {
          try {
            GmailApp.sendEmail(SUPPORT_EMAIL, '【録画配信エラー】自動配信に失敗しました（要確認）',
              'ファイル: ' + f.getName() + '\nURL: https://drive.google.com/file/d/' + f.getId() + '/view\n\nエラー: ' + String(err) +
              '\n\nお手数ですが、該当コマの受講者へ手動で共有をお願いします（閲覧のみ・DL不可）。',
              { name: FROM_NAME });
          } catch (_) { }
        }
      }
    }
  });
}

// 1ファイル分の配信処理
function processRecording_(f, slots, log, tz) {
  const name = f.getName();
  // 録画ファイル名は「<予定タイトル> (日時...)」形式 → タイトル前方一致＋録画作成日でコマを特定
  const createdYmd = Utilities.formatDate(f.getDateCreated(), tz, 'yyyy-MM-dd');
  let slot = null;
  Object.keys(slots).forEach(function (u) {
    const s = slots[u];
    const title = 'aicrew 生成AI研修 ' + s.kai + ' ' + s.theme + '（' + s.target + '）';
    if (name.indexOf(title) === 0 && s.date === createdYmd) slot = s;
  });
  const pubYmd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const expire = new Date(); expire.setMonth(expire.getMonth() + VIEW_MONTHS);
  const expYmd = Utilities.formatDate(expire, tz, 'yyyy-MM-dd');

  if (!slot) {
    // どのコマか特定できない録画（テスト録画など）→ ログに記録し事務局へ連絡、配信はしない
    log.appendRow([pubYmd, '', f.getId(), name, '(未特定)', createdYmd, 0, '未特定', '']);
    GmailApp.sendEmail(SUPPORT_EMAIL, '【録画配信】コマを特定できない録画があります',
      'ファイル: ' + name + '\nURL: https://drive.google.com/file/d/' + f.getId() + '/view\n作成日: ' + createdYmd +
      '\n\n研修の録画であれば、該当コマの受講者へ手動で共有してください（閲覧のみ・DL不可）。テスト録画等であれば対応不要です。',
      { name: FROM_NAME });
    return;
  }

  const recipients = slotRecipients_(slot.date, slot.url);
  if (recipients.length === 0) {
    log.appendRow([pubYmd, '', f.getId(), name, slot.slotKey, slot.date, 0, '受講者なし', '']);
    return;
  }

  // ダウンロード・コピー・再共有を禁止（閲覧のみ）
  try { driveFetch_('patch', '/files/' + f.getId() + '?fields=id', { copyRequiresWriterPermission: true, writersCanShare: false }); } catch (_) { }

  // 受講者ごとに閲覧権限を付与（Googleからの通知メールは送らず、下の案内メールに一本化）
  const sent = [];
  recipients.forEach(function (rc) {
    try {
      driveFetch_('post', '/files/' + f.getId() + '/permissions?sendNotificationEmail=false',
        { role: 'reader', type: 'user', emailAddress: rc.email });
      sendRecordingMail_(rc, slot, f.getId(), expYmd);
      sent.push(rc.email);
    } catch (permErr) {
      GmailApp.sendEmail(SUPPORT_EMAIL, '【録画配信】一部の受講者への配信に失敗しました',
        '受講者: ' + rc.name + '（' + rc.email + '）\nファイル: ' + name + '\nエラー: ' + String(permErr) +
        '\n\n手動での共有をお願いします。', { name: FROM_NAME });
    }
  });
  log.appendRow([pubYmd, expYmd, f.getId(), name, slot.slotKey, slot.date, sent.length, '配信済', sent.join(', ')]);
}

// 受講者への録画案内メール
function sendRecordingMail_(rc, slot, fileId, expYmd) {
  const link = 'https://drive.google.com/file/d/' + fileId + '/view';
  const html = '<div style="font-family:sans-serif;color:#1a1a1a;line-height:1.6">' +
    '<h2 style="color:#0f6e56">研修録画のご案内（復習用）</h2>' +
    '<p><b>' + esc_(rc.name) + '</b> 様</p>' +
    '<p>ご受講いただいた下記の回の録画を、復習用にご案内します。</p>' +
    '<p style="font-size:13px;background:#f0fbf6;border:1px solid #d7ece2;border-radius:8px;padding:10px 12px">' +
    '<b>' + esc_(slot.kai) + ' ' + esc_(slot.theme) + '（' + esc_(slot.target) + '）</b>　' + esc_(md_(slot.date)) + '<br>' +
    '<a href="' + link + '">録画を視聴する</a>（ご登録のGoogleアカウント/メールアドレスでご覧いただけます）</p>' +
    '<p style="font-size:12.5px;background:#fff8e6;border:1px solid #f0d98a;border-radius:8px;padding:10px 12px;color:#7a5c00">' +
    '【視聴期限】<b>' + esc_(jpDate_(expYmd)) + ' まで</b>（公開日から2ヶ月間）。期限を過ぎると自動的に視聴できなくなります。<br>' +
    '【ご注意】録画は<b>受講者ご本人の復習用</b>です。第三者への送信・共有・SNS等への転載はご遠慮ください。</p>' +
    '<p style="font-size:12px;color:#777">視聴できない場合は ' + SUPPORT_EMAIL + ' までご連絡ください。</p></div>';
  GmailApp.sendEmail(rc.email, '【aicrew 生成AI研修】録画のご案内（' + slot.kai + ' ' + slot.theme + '）',
    'HTMLメール対応のメーラーでご覧ください。', { name: FROM_NAME, htmlBody: html });
}

// 毎日1回トリガー：視聴期限（公開日から2ヶ月）を過ぎた録画の閲覧権限を自動で外す
function revokeExpiredRecordings() {
  const tz = 'Asia/Tokyo';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(RECDIST_SHEET);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][7]) !== '配信済') continue;
    const exp = (data[r][1] instanceof Date) ? Utilities.formatDate(data[r][1], tz, 'yyyy-MM-dd') : String(data[r][1]).slice(0, 10);
    if (!exp || exp >= today) continue;
    const fileId = data[r][2];
    const emails = String(data[r][8] || '').split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
    try {
      const perms = driveFetch_('get', '/files/' + fileId + '/permissions?fields=permissions(id,emailAddress,role)');
      (perms.permissions || []).forEach(function (pm) {
        const em = String(pm.emailAddress || '').toLowerCase();
        if (pm.role === 'reader' && emails.indexOf(em) >= 0) {
          try { driveFetch_('delete', '/files/' + fileId + '/permissions/' + pm.id); } catch (_) { }
        }
      });
      sh.getRange(r + 1, 8).setValue('期限済');
    } catch (err) {
      GmailApp.sendEmail(SUPPORT_EMAIL, '【録画配信】視聴期限の権限解除に失敗しました',
        'fileId: ' + fileId + '\nエラー: ' + String(err) + '\n\nDriveで手動で共有解除をお願いします。', { name: FROM_NAME });
    }
  }
}

// 一度だけ実行：録画配信（毎時）＋期限失効（毎日3時）のトリガーを設定
function setupRecordingTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const h = t.getHandlerFunction();
    if (h === 'distributeRecordings' || h === 'revokeExpiredRecordings') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('distributeRecordings').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('revokeExpiredRecordings').timeBased().atHour(3).everyDays(1).inTimezone('Asia/Tokyo').create();
  return 'recording triggers set (distribute hourly / revoke daily 3:00 JST)';
}

// ---- Supabase キープアライブ（無料プランの自動停止防止＋停止検知アラート）----
// 毎日1回、DBに軽いクエリを投げて「利用中」を維持する。応答がない＝停止の疑い → 事務局へ即アラート。
function keepAliveSupabase() {
  let ok = false, detail = '';
  try {
    const res = UrlFetchApp.fetch(SB_REST, {
      headers: { apikey: SB_KEY, Authorization: 'Bearer ' + SB_KEY },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    ok = (code >= 200 && code < 300);
    detail = 'HTTP ' + code + ' / ' + res.getContentText().slice(0, 100);
  } catch (err) {
    detail = String(err);  // DNS消失（＝一時停止）の場合はここに来る
  }
  if (!ok) {
    GmailApp.sendEmail(SUPPORT_EMAIL,
      '【要対応】Supabaseが応答しません（申込フォームが停止している恐れ）',
      '申込フォームの送信基盤（Supabase）への疎通確認に失敗しました。\n' +
      'この状態では、受講者がフォームから申込みできません（「送信に失敗しました」になります）。\n\n' +
      '■ 確認結果\n' + detail + '\n\n' +
      '■ 対応方法\n' +
      '下記URLを開き、「一時停止中」と表示されていたら「Restore project（プロジェクトを再開）」を押してください（2〜5分で復旧）。\n' +
      SB_DASHBOARD + '\n',
      { name: FROM_NAME });
  }
  return ok ? 'alive: ' + detail : 'NG: ' + detail;
}

// ---- 出席の自動記録（Meet監査ログ：誰が・どの会議に・何分いたか）----
// 前提：このスクリプトの実行アカウント（support@）に、管理コンソールの「レポート」閲覧権限（管理者ロール）が必要。

// Meetの会議コードを突合用に正規化（監査ログは「ABCDEFGHIJ」形式、URLは「abc-defg-hij」形式のため）
function normMeetCode_(s) { return String(s || '').replace(/-/g, '').toLowerCase(); }

// 監査ログ（call_ended）を取得し、{code|email: 滞在分} と生ログ配列を返す
function fetchMeetAudit_() {
  const start = new Date(Date.now() - 2 * 86400000).toISOString();  // 直近2日分
  const agg = {}, raw = [];
  let pageToken = '';
  for (let page = 0; page < 5; page++) {
    const url = 'https://admin.googleapis.com/admin/reports/v1/activity/users/all/applications/meet' +
      '?eventName=call_ended&maxResults=1000&startTime=' + encodeURIComponent(start) +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) throw new Error('reports_api ' + code + ': ' + res.getContentText().slice(0, 300));
    const body = JSON.parse(res.getContentText());
    (body.items || []).forEach(function (item) {
      const when = item.id && item.id.time ? String(item.id.time) : '';
      (item.events || []).forEach(function (ev) {
        let mc = '', email = '', dur = 0, disp = '';
        (ev.parameters || []).forEach(function (pm) {
          if (pm.name === 'meeting_code') mc = pm.value || '';
          if (pm.name === 'identifier') email = pm.value || '';
          if (pm.name === 'display_name') disp = pm.value || '';
          if (pm.name === 'duration_seconds') dur = Number(pm.intValue || pm.value || 0);
        });
        if (!mc) return;
        const mins = Math.round(dur / 60);
        raw.push({ time: when, code: mc, email: email, disp: disp, mins: mins });
        if (email) {
          const key = normMeetCode_(mc) + '|' + String(email).toLowerCase();
          agg[key] = (agg[key] || 0) + mins;
        }
      });
    });
    pageToken = body.nextPageToken || '';
    if (!pageToken) break;
  }
  return { agg: agg, raw: raw };
}

// 毎日22時トリガー：当日・前日の受講記録の「出席」列を監査ログから自動入力
function markAttendance() {
  const tz = 'Asia/Tokyo';
  const days = [Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
                Utilities.formatDate(new Date(Date.now() - 86400000), tz, 'yyyy-MM-dd')];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(DETAIL_SHEET);
  if (!sh) return 'no detail sheet';
  const data = sh.getDataRange().getValues();
  // 対象行：日付が当日/前日 かつ 出席が空欄 or 「要確認」（ログ遅延に備えて再判定）
  const targets = [];
  for (let r = 1; r < data.length; r++) {
    const dc = data[r][8];  // I列: 日付
    const ds = (dc instanceof Date) ? Utilities.formatDate(dc, tz, 'yyyy-MM-dd') : String(dc).slice(0, 10);
    if (days.indexOf(ds) < 0) continue;
    const att = String(data[r][16] || '');  // Q列: 出席
    if (att && att.indexOf('要確認') !== 0) continue;
    const url = String(data[r][17] || '');  // R列: Meet URL
    if (!url) continue;
    targets.push({ row: r + 1, email: String(data[r][4] || '').toLowerCase(), code: normMeetCode_(meetCodeFromUrl_(url)), cur: att });
  }
  if (targets.length === 0) return 'no target rows';

  let audit;
  try {
    audit = fetchMeetAudit_();
  } catch (err) {
    GmailApp.sendEmail(SUPPORT_EMAIL, '【出席自動記録】Meetの入室ログを取得できませんでした',
      'エラー: ' + String(err) + '\n\n' +
      'このスクリプトの実行アカウントに、Google Workspace 管理コンソールの「レポート」閲覧権限（管理者ロール）が付与されているかご確認ください。\n' +
      '管理コンソール → アカウント → 管理者ロール で、レポートを閲覧できるロールを support@ のアカウントに割り当ててください。\n' +
      'それまでの間、出席は従来どおり受講記録シートへの手入力でお願いします。', { name: FROM_NAME });
    return 'audit fetch failed';
  }

  // 生ログを _入室ログ に追記（同じ日×会議×人は1行、証憑・手動確認用）
  const logSh = ensureSheet_(ATTLOG_SHEET, ['日時', '会議コード', 'メール', '表示名', '滞在分']);
  const logData = logSh.getDataRange().getValues();
  const seen = {};
  for (let r = 1; r < logData.length; r++) {
    seen[String(logData[r][0]).slice(0, 10) + '|' + normMeetCode_(logData[r][1]) + '|' + String(logData[r][2] || logData[r][3]).toLowerCase()] = true;
  }
  audit.raw.forEach(function (e) {
    const k = String(e.time).slice(0, 10) + '|' + normMeetCode_(e.code) + '|' + String(e.email || e.disp).toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    logSh.appendRow([e.time, e.code, e.email, e.disp, e.mins]);
  });

  // 突合して出席列を更新
  let marked = 0;
  targets.forEach(function (t) {
    if (!t.email || !t.code) return;
    const mins = audit.agg[t.code + '|' + t.email] || 0;
    let val = '';
    if (mins >= ATTEND_MIN_MINUTES) val = '出席';
    else if (mins > 0) val = '要確認(' + mins + '分)';
    if (val && val !== t.cur) { sh.getRange(t.row, 17).setValue(val); marked++; }
  });
  return 'marked ' + marked + ' / targets ' + targets.length;
}

// 一度だけ実行：すべての自動処理トリガーをまとめて登録（リマインド／録画配信／期限失効／Supabase維持／出席記録）
function setupAllTriggers() {
  const a = setupReminderTrigger();
  const b = setupRecordingTriggers();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const h = t.getHandlerFunction();
    if (h === 'keepAliveSupabase' || h === 'markAttendance') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('keepAliveSupabase').timeBased().atHour(8).everyDays(1).inTimezone('Asia/Tokyo').create();
  ScriptApp.newTrigger('markAttendance').timeBased().atHour(22).everyDays(1).inTimezone('Asia/Tokyo').create();
  return a + ' / ' + b + ' / keepalive set (daily 8:00 JST) / attendance set (daily 22:00 JST)';
}

function icsUtc_(date, hm) {
  const a = date.split('-').map(Number), b = hm.split(':').map(Number);
  const d = new Date(Date.UTC(a[0], a[1] - 1, a[2], b[0] - 9, b[1]));
  const p = n => ('0' + n).slice(-2);
  return d.getUTCFullYear() + p(d.getUTCMonth() + 1) + p(d.getUTCDate()) + 'T' + p(d.getUTCHours()) + p(d.getUTCMinutes()) + '00Z';
}

function buildIcs_(p) {
  const d = new Date(), z = n => ('0' + n).slice(-2);
  const now = d.getUTCFullYear() + z(d.getUTCMonth() + 1) + z(d.getUTCDate()) + 'T' + z(d.getUTCHours()) + z(d.getUTCMinutes()) + z(d.getUTCSeconds()) + 'Z';
  const L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//aicrew//kensyu//JA', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH'];
  p.sessions.forEach(function (o, i) {
    const [st, en] = splitTime_(o.time);
    const loc = o.meetUrl ? o.meetUrl : 'オンライン(Google Meet)';
    L.push('BEGIN:VEVENT',
      'UID:kensyu-' + o.date + '-' + i + '-' + Utilities.getUuid() + '@aicrew.jp',
      'DTSTAMP:' + now, 'DTSTART:' + icsUtc_(o.date, st), 'DTEND:' + icsUtc_(o.date, en),
      'SUMMARY:aicrew 生成AI研修 ' + CIR[i] + ' ' + o.theme + ' (' + p.course_label + ')',
      'LOCATION:' + loc,
      'DESCRIPTION:オンライン(Google Meet)' + (o.meetUrl ? ' ' + o.meetUrl : ''),
      'END:VEVENT');
  });
  L.push('END:VCALENDAR');
  return L.join('\r\n');
}
