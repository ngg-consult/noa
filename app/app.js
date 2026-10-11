/*
 * ממשק "סדר היום": מושך יומן ומיילים מ-Microsoft 365 דרך המחבר של claude.ai,
 * מסווג בעזרת Claude, משבץ עם MorningEngine ושומר מצב פרטי במסד הנתונים של הדף.
 */
(function () {
  'use strict';
  var E = window.MorningEngine;
  var SERVER = 'Microsoft 365';
  var CFG = E.DEFAULTS;

  var cap = { db: null, mcp: null, sample: null, downloads: null, user: null, permissions: null };
  var state = {
    uid: null, me: null, plan: null, rules: [], handled: {},
    fetchedAt: null, busy: '', pendingRule: null,
    view: null, preview: null, previewFetchedAt: null,
    src: { calendar: 'idle', mail: 'idle', ai: 'idle' }, srcMsg: {}
  };

  var $ = function (id) { return document.getElementById(id); };

  // ---------- עזרי זמן ----------

  function nowLocal() { return E.localParts(new Date(), CFG.tz); }

  var HE_DATE = new Intl.DateTimeFormat('he-IL', { timeZone: CFG.tz, weekday: 'long', day: 'numeric', month: 'long' });
  var HE_TIME = new Intl.DateTimeFormat('he-IL', { timeZone: CFG.tz, hour: '2-digit', minute: '2-digit' });

  function relAge(iso) {
    if (!iso) return '';
    var h = (Date.now() - new Date(iso).getTime()) / 3600000;
    if (h < 1) return 'לפני ' + Math.max(1, Math.round(h * 60)) + ' דק\'';
    if (h < 24) return 'לפני ' + Math.round(h) + ' שע\'';
    return 'לפני ' + Math.round(h / 24) + ' ימים';
  }

  // אחרי סוף יום העבודה מוצג אוטומטית יום העבודה הבא
  function dayCfg() { return E.settingsFromRules(state.rules, CFG); }
  function nextDate() { return E.nextWorkday(nowLocal().date); }
  function isEvening() { return nowLocal().min >= E.hm(dayCfg().dayEnd); }
  function activeView() { return state.view || (isEvening() ? 'next' : 'today'); }
  var HE_WEEKDAY = new Intl.DateTimeFormat('he-IL', { timeZone: CFG.tz, weekday: 'long' });
  function nextLabel() {
    var d = nextDate();
    return d === E.addDays(nowLocal().date, 1) ? 'מחר' : HE_WEEKDAY.format(new Date(d + 'T12:00:00Z'));
  }

  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

  // ---------- אחסון (פרטי לכל משתמשת) ----------

  function docRef(name) { return cap.db.collection('data/users/' + state.uid).doc(name); }

  async function loadState() {
    if (!cap.db || !state.uid) return;
    var res = await Promise.allSettled([docRef('plan').get(), docRef('rules').get(), docRef('handled').get(), docRef('preview').get()]);
    if (res[0].status === 'fulfilled' && res[0].value.exists) {
      var d = res[0].value.data();
      state.plan = d.plan || null;
      state.fetchedAt = d.fetchedAt || null;
    }
    if (res[1].status === 'fulfilled' && res[1].value.exists) state.rules = res[1].value.data().rules || [];
    if (res[2].status === 'fulfilled' && res[2].value.exists) state.handled = res[2].value.data().ids || {};
    if (res[3].status === 'fulfilled' && res[3].value.exists) {
      state.preview = res[3].value.data().preview || null;
      state.previewFetchedAt = res[3].value.data().fetchedAt || null;
    }
  }

  var saving = {};
  function save(name, body) {
    if (!cap.db || !state.uid) return Promise.resolve();
    // כתיבה אחת בכל פעם לכל מסמך
    var prev = saving[name] || Promise.resolve();
    saving[name] = prev.then(function () { return docRef(name).set(body); }).catch(function (e) {
      toast('השמירה נכשלה (' + (e && e.code || 'שגיאה') + '). השינוי יישמר בפעולה הבאה.');
    });
    return saving[name];
  }

  function savePlan() { return save('plan', { plan: state.plan, fetchedAt: state.fetchedAt }); }
  function savePreview() { return save('preview', { preview: state.preview, fetchedAt: state.previewFetchedAt }); }
  function saveRules() { return save('rules', { rules: state.rules, updatedAt: new Date().toISOString() }); }
  function saveHandled() {
    var cutoff = E.addDays(nowLocal().date, -30), ids = {};
    Object.keys(state.handled).forEach(function (k) { if (state.handled[k] >= cutoff) ids[k] = state.handled[k]; });
    state.handled = ids;
    return save('handled', { ids: ids });
  }

  // ---------- Microsoft 365 ----------

  function mcpMessage(e) {
    var code = e && e.code;
    if (code === 'needs_reauth') return 'החיבור ל-Microsoft 365 פג. יש לחבר מחדש ב-claude.ai ← הגדרות ← Connectors.';
    if (code === 'server_not_connected' || code === 'server_not_found') return 'Microsoft 365 לא מחובר. יש להוסיף אותו ב-claude.ai ← הגדרות ← Connectors.';
    if (code === 'selection_required') return 'יש יותר מחיבור Microsoft 365 אחד. בחרי אחד בהודעה שהופיעה מעל הדף.';
    if (code === 'not_in_manifest' || code === 'approval_required' || code === 'consent_required') return 'הדף עוד לא קיבל אישור לקרוא מ-Microsoft 365. אשרי בתפריט ההרשאות.';
    if (code === 'blocked_by_policy') return 'מדיניות הארגון חוסמת את הקריאה הזו. יש לפנות למנהל המערכת.';
    if (code === 'not_granted' || code === 'capability_disabled') return 'החיבור לנתונים זמין רק בתוך claude.ai.';
    if (code === 'server_unavailable') return 'Microsoft 365 לא הגיב. ננסה שוב ברענון הבא.';
    return 'שגיאה בקריאה מ-Microsoft 365' + (e && e.message ? ': ' + e.message : '');
  }

  async function call(tool, input) {
    try {
      return await cap.mcp.callTool(SERVER, tool, input, { cache: false });
    } catch (e) {
      if (e && e.retryable) {
        await new Promise(function (r) { setTimeout(r, (e.retryAfterMs || 1500) + Math.random() * 800); });
        return await cap.mcp.callTool(SERVER, tool, input, { cache: false });
      }
      throw e;
    }
  }

  async function callPaged(tool, input, maxPages) {
    var all = [], offset = 0;
    for (var i = 0; i < maxPages; i++) {
      var r = await call(tool, Object.assign({}, input, { limit: 25, offset: offset }));
      var p = E.parseToolBlocks(r);
      all = all.concat(p.items);
      if (p.nextOffset == null || !p.items.length) break;
      offset = p.nextOffset;
    }
    return all;
  }

  async function whoAmI() {
    if (state.me) return state.me;
    try {
      var r = await call('get_me', {});
      var p = r.payload || {};
      state.me = { email: p.mail || p.userPrincipalName || '', name: p.displayName || '' };
    } catch (e) { state.me = { email: '', name: '' }; }
    return state.me;
  }

  async function fetchCalendar(today) {
    return callPaged('outlook_calendar_search', {
      query: '*',
      afterDateTime: E.addDays(today, -1) + 'T00:00:00',
      beforeDateTime: E.addDays(today, 1) + 'T23:59:00',
      order: 'oldest'
    }, 4);
  }

  async function fetchMail(today) {
    return callPaged('outlook_email_search', {
      afterDateTime: E.addDays(today, -CFG.mailLookbackDays) + 'T00:00:00'
    }, 3);
  }

  // ---------- Claude (סיווג, הכנה לפגישות, פירוש כללים) ----------

  function aiNote(ok, msg) {
    state.src.ai = ok ? 'ok' : 'fallback';
    if (msg) state.srcMsg.ai = msg;
  }

  function sampleErr(e) {
    var c = e && e.code;
    if (c === 'not_granted') return 'הסיווג החכם לא אושר לדף. מוצג סיווג בסיסי.';
    if (c === 'rate_limited') return 'הגעת למגבלת השימוש ב-Claude. מוצג סיווג בסיסי.';
    return 'הסיווג החכם לא זמין כרגע. מוצג סיווג בסיסי.';
  }

  var CLASSIFY_PROMPT = [
    'You are the executive assistant of {who}, a consultant at NGG. Triage each incoming item with the Eisenhower matrix and David Allen\'s GTD.',
    'For every item return:',
    '- actionable: false when it is informational, an automated notification, a newsletter, or needs nothing from the user.',
    '- nextAction: the very next physical action, in Hebrew, starting with an infinitive verb (e.g. "להשיב לרחל עם לוח זמנים מעודכן"), at most 12 words.',
    '- minutes: realistic minutes for that next action only (2-120).',
    '- context: exactly one of "מחשב", "טלפון", "פגישה", "בחוץ".',
    '- urgent: true when there is a deadline today or tomorrow, someone is waiting on the user, or delay causes damage.',
    '- important: true when it advances client work, revenue, management commitments or the user\'s goals.',
    '- delegateTo: a person or role (from the item) who could do it instead of the user, or null when the user must do it personally.',
    '- client: the client or organisation the item concerns (from the sender domain or text), or null.',
    'Reply with only JSON: {"items":[{"id":"...","actionable":true,"nextAction":"...","minutes":15,"context":"מחשב","urgent":false,"important":true,"delegateTo":null,"client":null}]} containing every input id.',
    'Items:'
  ].join('\n');

  async function classify(items) {
    var out = {};
    if (!items.length) return out;
    var who = state.me && state.me.email ? (state.me.name + ' <' + state.me.email + '>') : 'the user';
    if (!cap.sample) {
      items.forEach(function (i) { out[i.id] = E.heuristicClassify(i); });
      aiNote(false, 'הסיווג החכם לא זמין בתצוגה הזו. מוצג סיווג בסיסי.');
      return out;
    }
    for (var b = 0; b < items.length; b += 12) {
      var batch = items.slice(b, b + 12);
      var data = batch.map(function (i) {
        return { id: i.id, from: i.from, subject: i.title, snippet: i.snippet, receivedAt: i.receivedAt, importance: i.importance };
      });
      try {
        setBusy('מתעדף ' + Math.min(b + 12, items.length) + ' מתוך ' + items.length + ' פריטים…');
        var r = await cap.sample.json(CLASSIFY_PROMPT.replace('{who}', who) + '\n' + JSON.stringify(data), { modelTier: 'default' });
        var byId = {};
        ((r && r.items) || []).forEach(function (x) { if (x && x.id) byId[x.id] = x; });
        batch.forEach(function (i) { out[i.id] = E.sanitizeClassification(byId[i.id], i); });
        aiNote(true);
      } catch (e) {
        batch.forEach(function (i) { out[i.id] = E.heuristicClassify(i); });
        aiNote(false, sampleErr(e));
      }
    }
    return out;
  }

  async function prepMeetings(meetings, items) {
    var need = meetings.filter(function (m) { return !m.prep; });
    if (!need.length) return;
    need.forEach(function (m) { m.prep = E.heuristicPrep(m, items); });
    if (!cap.sample) return;
    var data = need.map(function (m) {
      return { id: m.id, title: m.title, organizer: m.organizer, attendees: m.attendees.slice(0, 8), location: m.location, time: E.fmt(m.start) };
    });
    var mail = items.slice(0, 40).map(function (i) { return { from: i.from, subject: i.title, snippet: i.snippet.slice(0, 120) }; });
    try {
      setBusy('מכין שורת הכנה לכל פגישה…');
      var r = await cap.sample.json([
        'For each meeting today write ONE short Hebrew line (max 14 words) telling the user what to prepare beforehand.',
        'Be concrete: name the document, figure, decision or open question to bring, using the recent emails when they relate to the meeting\'s subject or attendees.',
        'For a routine internal sync, name the one update to bring. Do not invent facts that are not in the data.',
        'Reply with only JSON: {"preps":[{"id":"...","prep":"..."}]}',
        'Meetings: ' + JSON.stringify(data),
        'Recent emails: ' + JSON.stringify(mail)
      ].join('\n'), { modelTier: 'default' });
      ((r && r.preps) || []).forEach(function (p) {
        var m = need.find(function (x) { return x.id === p.id; });
        if (m && typeof p.prep === 'string' && p.prep.trim()) m.prep = p.prep.trim().slice(0, 140);
      });
    } catch (e) { aiNote(false, sampleErr(e)); }
  }

  var TIME = /^\d{2}:\d{2}$/;
  function cleanRule(r) {
    if (!r || typeof r !== 'object') return null;
    var strs = function (a) { return Array.isArray(a) ? a.filter(function (x) { return typeof x === 'string' && x.trim(); }).map(function (x) { return x.trim(); }).slice(0, 10) : undefined; };
    var m = r.match || {}, e = r.effect || {};
    var match = {
      keywords: strs(m.keywords), senders: strs(m.senders),
      contexts: (strs(m.contexts) || []).filter(function (c) { return E.CONTEXTS.indexOf(c) >= 0; }),
      quadrants: (strs(m.quadrants) || []).filter(function (q) { return /^Q[1-4]$/.test(q); }),
      maxMinutes: Number(m.maxMinutes) || undefined, minMinutes: Number(m.minMinutes) || undefined,
      all: m.all === true || undefined
    };
    var effect = {};
    if (typeof e.priority === 'number' && isFinite(e.priority)) effect.priority = Math.max(-100, Math.min(300, e.priority));
    if (typeof e.urgent === 'boolean') effect.urgent = e.urgent;
    if (typeof e.important === 'boolean') effect.important = e.important;
    if (e.window && (TIME.test(e.window.start) || TIME.test(e.window.end))) {
      effect.window = { start: TIME.test(e.window.start) ? e.window.start : null, end: TIME.test(e.window.end) ? e.window.end : null };
    }
    if (Number(e.minutes) > 0) effect.minutes = Math.min(240, Math.round(Number(e.minutes)));
    if (E.CONTEXTS.indexOf(e.context) >= 0) effect.context = e.context;
    if (typeof e.delegateTo === 'string' && e.delegateTo.trim()) effect.delegateTo = e.delegateTo.trim().slice(0, 60);
    if (e.skip === true) effect.skip = true;
    if (e.settings && typeof e.settings === 'object') {
      var s = {};
      if (TIME.test(e.settings.dayStart)) s.dayStart = e.settings.dayStart;
      if (TIME.test(e.settings.dayEnd)) s.dayEnd = e.settings.dayEnd;
      if (Number(e.settings.bufferMin) >= 0 && e.settings.bufferMin !== null && e.settings.bufferMin !== undefined) s.bufferMin = Math.min(30, Number(e.settings.bufferMin));
      if (Object.keys(s).length) effect.settings = s;
    }
    Object.keys(match).forEach(function (k) { if (match[k] === undefined || (Array.isArray(match[k]) && !match[k].length)) delete match[k]; });
    var rule = { match: match, effect: effect };
    rule.explain = (typeof r.explain === 'string' && r.explain.trim()) ? r.explain.trim().slice(0, 160) : E.explainRule(rule);
    return E.isUsableRule(rule) ? rule : null;
  }

  async function interpretRule(text) {
    if (cap.sample) {
      try {
        var r = await cap.sample.json([
          'Convert the user\'s scheduling instruction (Hebrew) into a JSON rule for her daily planner.',
          'Schema: {"match":{"keywords":[string],"senders":[string],"contexts":["מחשב"|"טלפון"|"פגישה"|"בחוץ"],"quadrants":["Q1"|"Q2"|"Q3"|"Q4"],"maxMinutes":number,"minMinutes":number,"all":boolean},',
          '"effect":{"priority":number,"urgent":boolean,"important":boolean,"window":{"start":"HH:MM","end":"HH:MM"},"minutes":number,"context":string,"delegateTo":string,"skip":boolean,"settings":{"dayStart":"HH:MM","dayEnd":"HH:MM","bufferMin":number}},',
          '"explain":"one short Hebrew sentence saying what the rule does"}',
          'Omit fields that do not apply. priority ranges -100..300; 200 means top of the day, 40 a moderate boost, negative pushes later.',
          'Day parts: בוקר 08:00-12:00, צהריים 12:00-15:00, אחה"צ 14:00-18:00, ערב 17:00-21:00.',
          'keywords are short distinctive terms (client or project name); add the English spelling or email domain when obvious. Quadrants: Q1 urgent+important, Q2 important, Q3 urgent only, Q4 neither.',
          'For working hours use effect.settings with match.all=true.',
          'Instruction: ' + JSON.stringify(text),
          'Reply with only the JSON object.'
        ].join('\n'), { modelTier: 'default', cache: false });
        var c = cleanRule(r);
        if (c) return c;
      } catch (e) { /* ננסה פירוש בסיסי */ }
    }
    var h = E.parseRuleText(text);
    return E.isUsableRule(h) ? cleanRule(h) : null;
  }

  async function extractPasted(text) {
    var lines = text.split(/\n+/).map(function (l) { return l.trim(); }).filter(Boolean);
    if (cap.sample) {
      try {
        var r = await cap.sample.json([
          'The user pasted messages (WhatsApp, Teams or notes). Extract only the action items that are hers to do.',
          'For each: {"title": short Hebrew title, "from": who asked or null, "nextAction": Hebrew infinitive next action, "minutes": 2-120, "context": "מחשב"|"טלפון"|"פגישה"|"בחוץ", "urgent": boolean, "important": boolean, "delegateTo": string|null, "client": string|null}.',
          'Reply with only JSON: {"tasks":[...]}. Return {"tasks":[]} if nothing is actionable.',
          'Text:\n' + text.slice(0, 12000)
        ].join('\n'), { modelTier: 'default', cache: false });
        if (r && Array.isArray(r.tasks)) {
          return r.tasks.slice(0, 30).map(function (t) {
            var item = { id: 'manual:' + uid(), source: 'paste', title: String(t.title || t.nextAction || 'משימה').slice(0, 120), from: t.from || '', snippet: '', receivedAt: new Date().toISOString(), link: '' };
            return E.makeTask(item, E.sanitizeClassification(t, item));
          });
        }
      } catch (e) { /* גיבוי: שורה = משימה */ }
    }
    return lines.slice(0, 30).map(function (l) {
      var item = { id: 'manual:' + uid(), source: 'manual', title: l.slice(0, 120), from: '', snippet: '', receivedAt: new Date().toISOString(), link: '' };
      var c = E.heuristicClassify(item);
      c.nextAction = l.slice(0, 160);
      c.important = true;
      return E.makeTask(item, c);
    });
  }

  // ---------- בניית היום ----------

  function replanNow() {
    if (!state.plan) return;
    var n = nowLocal();
    var from = state.plan.date === n.date ? n.min : 0;
    state.plan = E.replan(state.plan, state.rules, { nowMs: Date.now(), fromMin: from, cfg: CFG });
    previewLocal();
  }

  // לו"ז יום העבודה הבא: הפגישות שלו + כל מה שעוד פתוח היום, מתחילת היום
  function previewLocal() {
    if (!state.preview || !state.plan) return;
    var base = { date: state.preview.date, meetings: state.preview.meetings || [], allDay: state.preview.allDay || [], tasks: E.carryOver(state.plan) };
    state.preview = E.replan(base, state.rules, { nowMs: Date.now(), fromMin: 0, cfg: CFG });
  }

  async function buildPreview() {
    if (state.busy || !state.plan) return;
    var date = nextDate();
    var old = state.preview && state.preview.date === date ? state.preview : null;
    var meetings = old ? old.meetings : [], allDay = old ? (old.allDay || []) : [];
    setBusy('מושך את היומן של ' + nextLabel() + '…');
    if (cap.mcp) {
      try {
        var raw = await fetchCalendar(date);
        var ev = E.normalizeEvents(raw, date, CFG);
        var oldPrep = {};
        meetings.forEach(function (m) { if (m.prep) oldPrep[m.id] = m.prep; });
        ev.meetings.forEach(function (m) { if (oldPrep[m.id]) m.prep = oldPrep[m.id]; });
        meetings = ev.meetings; allDay = ev.allDay;
        state.src.calendar = 'ok'; delete state.srcMsg.calendar;
      } catch (e) { state.src.calendar = 'error'; state.srcMsg.calendar = mcpMessage(e); }
    }
    var mailItems = state.plan.tasks.filter(function (t) { return t.source === 'mail'; });
    await prepMeetings(meetings, mailItems);
    state.preview = { date: date, meetings: meetings, allDay: allDay, tasks: [] };
    state.previewFetchedAt = new Date().toISOString();
    previewLocal();
    setBusy('');
    render();
    savePreview();
  }

  async function buildDay() {
    if (state.busy) return;
    var today = nowLocal().date;
    var prev = state.plan;
    var sameDay = prev && prev.date === today;
    var base = sameDay ? prev.tasks.slice() : E.carryOver(prev);
    setBusy('מושך את היומן והמיילים מ-Outlook…');

    var meetings = sameDay ? prev.meetings : [], allDay = sameDay ? (prev.allDay || []) : [];
    var emails = [];
    if (cap.mcp) {
      await whoAmI();
      var res = await Promise.allSettled([fetchCalendar(today), fetchMail(today)]);
      if (res[0].status === 'fulfilled') {
        var ev = E.normalizeEvents(res[0].value, today, CFG);
        var oldPrep = {};
        var prevMeetings = (meetings || []).concat(state.preview && state.preview.date === today ? state.preview.meetings || [] : []);
        prevMeetings.forEach(function (m) { if (m.prep) oldPrep[m.id] = m.prep; });
        ev.meetings.forEach(function (m) { if (oldPrep[m.id]) m.prep = oldPrep[m.id]; });
        meetings = ev.meetings; allDay = ev.allDay;
        state.src.calendar = 'ok'; delete state.srcMsg.calendar;
      } else { state.src.calendar = 'error'; state.srcMsg.calendar = mcpMessage(res[0].reason); }
      if (res[1].status === 'fulfilled') {
        emails = E.normalizeEmails(res[1].value, state.me && state.me.email);
        state.src.mail = 'ok'; delete state.srcMsg.mail;
      } else { state.src.mail = 'error'; state.srcMsg.mail = mcpMessage(res[1].reason); }
    } else {
      state.src.calendar = state.src.mail = 'error';
      state.srcMsg.calendar = state.srcMsg.mail = 'החיבור ל-Outlook זמין רק כשהדף פתוח בתוך claude.ai.';
    }

    var known = {};
    base.forEach(function (t) { known[t.id] = true; });
    var fresh = emails.filter(function (m) { return !known[m.id] && !state.handled[m.id]; });
    var cls = await classify(fresh);
    var newTasks = fresh.map(function (m) { return E.makeTask(m, cls[m.id] || E.heuristicClassify(m)); });

    await prepMeetings(meetings, emails);

    state.plan = { date: today, meetings: meetings, allDay: allDay, tasks: base.concat(newTasks) };
    state.fetchedAt = new Date().toISOString();
    replanNow();
    setBusy('');
    render();
    savePlan();
    if (activeView() === 'next') await buildPreview();
  }

  // ---------- פעולות משתמשת ----------

  function mark(id, status) {
    var n = nowLocal();
    state.plan = E.markTask(state.plan, id, status, n.min);
    if (status === 'done' || status === 'dismissed') { state.handled[id] = n.date; saveHandled(); }
    if (status === 'pending') { delete state.handled[id]; saveHandled(); }
    replanNow();
    render();
    savePlan();
    var t = state.plan.tasks.find(function (x) { return x.id === id; });
    if (activeView() === 'next') { if (status === 'done') toast('סומן כבוצע.'); return; }
    if (status === 'missed' && t) {
      toast(t.placement === 'today' ? 'הועבר ל-' + E.fmt(t.start) + '. שאר היום סודר מחדש.' : 'אין חלון פנוי היום. המשימה עברה למחר.');
    } else if (status === 'done') toast('סומן כבוצע. שאר היום סודר מחדש.');
  }

  function force(id) {
    state.plan.tasks = state.plan.tasks.map(function (t) { return t.id === id ? Object.assign({}, t, { forced: true }) : t; });
    replanNow(); render(); savePlan();
  }

  async function onRuleSubmit(ev) {
    ev.preventDefault();
    var text = $('rule-text').value.trim();
    if (!text) return;
    $('rule-submit').disabled = true;
    $('rule-preview').hidden = false;
    $('rule-preview').innerHTML = '<p class="muted">מפרש את הכלל…</p>';
    var rule = await interpretRule(text);
    $('rule-submit').disabled = false;
    if (!rule) {
      state.pendingRule = null;
      $('rule-preview').innerHTML = '<p class="warn">לא הצלחתי להבין את הכלל. נסי לנסח עם נושא ופעולה, למשל: "כל מה שקשור ללקוח X הכי דחוף" או "שיחות טלפון רק אחה"צ".</p>';
      return;
    }
    rule.text = text;
    state.pendingRule = rule;
    var box = $('rule-preview');
    box.innerHTML = '';
    var p = el('p', 'rule-explain', 'הבנתי: ' + rule.explain);
    var row = el('div', 'row');
    var ok = el('button', 'btn primary', 'שמירת הכלל');
    ok.type = 'button';
    ok.onclick = confirmRule;
    var no = el('button', 'btn ghost', 'ביטול');
    no.type = 'button';
    no.onclick = function () { state.pendingRule = null; box.hidden = true; };
    row.append(ok, no);
    box.append(p, row);
  }

  function confirmRule() {
    var r = state.pendingRule;
    if (!r) return;
    r.id = 'rule-' + uid();
    r.active = true;
    r.createdAt = new Date().toISOString();
    state.rules = state.rules.concat([r]);
    state.pendingRule = null;
    $('rule-text').value = '';
    $('rule-preview').hidden = true;
    saveRules();
    replanNow(); render(); savePlan();
    toast('הכלל נשמר והוחל על היום. הוא יופעל מחדש בכל בוקר.');
  }

  function toggleRule(id) {
    // כללים לא נמחקים – רק מושבתים, כדי לשמור תיעוד מלא
    state.rules = state.rules.map(function (r) { return r.id === id ? Object.assign({}, r, { active: r.active === false }) : r; });
    saveRules();
    replanNow(); render(); savePlan();
  }

  async function exportRules() {
    if (!cap.downloads) return;
    try {
      await cap.downloads.save({ filename: 'morning-rules.json', data: JSON.stringify({ rules: state.rules }, null, 2) });
    } catch (e) {
      if (e && e.code !== 'declined') toast('ההורדה לא זמינה כרגע.');
    }
  }

  async function onPasteSubmit(ev) {
    ev.preventDefault();
    var text = $('paste-text').value.trim();
    if (!text || !state.plan) return;
    $('paste-submit').disabled = true;
    setBusy('מחלץ משימות מהטקסט…');
    var tasks = await extractPasted(text);
    setBusy('');
    $('paste-submit').disabled = false;
    if (!tasks.length) { toast('לא נמצאו משימות בטקסט.'); return; }
    $('paste-text').value = '';
    state.plan.tasks = state.plan.tasks.concat(tasks);
    replanNow(); render(); savePlan();
    toast('נוספו ' + tasks.length + ' משימות ושובצו בלו"ז.');
  }

  function setView(v) {
    state.view = v;
    if (v === 'next' && (!state.preview || state.preview.date !== nextDate())) buildPreview();
    render();
  }

  // ---------- תצוגה ----------

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  function chip(text, cls) { return el('span', 'chip ' + (cls || ''), text); }

  function safeLink(href, label) {
    if (!/^https:\/\//.test(href || '')) return null;
    var a = el('a', 'open', label);
    a.href = href; a.target = '_blank'; a.rel = 'noopener noreferrer';
    return a;
  }

  var toastTimer;
  function toast(msg) {
    var t = $('toast');
    t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, 4200);
  }

  function setBusy(msg) {
    state.busy = msg;
    $('busy').hidden = !msg;
    $('busy-text').textContent = msg;
    $('refresh').disabled = !!msg;
  }

  function taskRow(t, n) {
    var li = el('li', 'item task q-' + t.quadrant + (t.status === 'done' ? ' is-done' : ''));
    var time = el('div', 'time');
    time.append(el('b', '', t.start != null ? E.fmt(t.start) : '—'), el('span', '', t.end != null ? E.fmt(t.end) : ''));
    var body = el('div', 'body');
    var title = el('div', 'title', t.nextAction || t.title);
    var sub = el('div', 'sub');
    var srcLabel = t.source === 'mail' ? 'מייל' : t.source === 'paste' ? 'הודבק' : 'ידני';
    sub.textContent = srcLabel + (t.from ? ' · ' + t.from : '') + (t.title && t.nextAction ? ' · ' + t.title : '') + (t.receivedAt && t.source === 'mail' ? ' · ' + relAge(t.receivedAt) : '');
    var chips = el('div', 'chips');
    chips.append(chip(E.QUADRANTS[t.quadrant].label, 'q'), chip(t.context), chip(t.minutes + ' דק\''));
    if (t.minutes <= 2) chips.append(chip('פחות משתי דקות', 'note'));
    if (t.delegateTo) chips.append(chip('להאציל ל' + t.delegateTo, 'deleg'));
    if (t.client) chips.append(chip(t.client));
    if (t.carried) chips.append(chip(t.missedCount ? 'נדחה ' + t.missedCount + '×' : 'מאתמול', 'note'));
    if (t.appliedRules && t.appliedRules.length) chips.append(chip('כלל', 'rule'));
    body.append(title, sub, chips);
    var link = safeLink(t.link, 'פתיחה ב-Outlook');
    if (link) body.append(link);

    var acts = el('div', 'acts');
    if (t.status === 'done') {
      var undo = el('button', 'btn ghost small', 'ביטול סימון');
      undo.onclick = function () { mark(t.id, 'pending'); };
      acts.append(el('span', 'done-mark', '✓ בוצע'), undo);
    } else {
      var d = el('button', 'btn ok small', 'בוצע');
      d.setAttribute('aria-label', 'סימון כבוצע: ' + (t.nextAction || t.title));
      d.onclick = function () { mark(t.id, 'done'); };
      var m = el('button', 'btn miss small', 'לא בוצע');
      m.setAttribute('aria-label', 'סימון כלא בוצע: ' + (t.nextAction || t.title));
      m.onclick = function () { mark(t.id, 'missed'); };
      var x = el('button', 'btn ghost small', 'לא רלוונטי');
      x.onclick = function () { mark(t.id, 'dismissed'); };
      if (activeView() === 'next') acts.append(d, x); else acts.append(d, m, x);
      if (activeView() === 'today' && t.start != null && t.end <= n.min && state.plan.date === n.date) li.classList.add('overdue');
    }
    li.append(time, body, acts);
    return li;
  }

  function meetingRow(m) {
    var li = el('li', 'item meeting' + (m.tentative ? ' tentative' : ''));
    var time = el('div', 'time');
    time.append(el('b', '', E.fmt(m.start)), el('span', '', E.fmt(m.end)));
    var body = el('div', 'body');
    body.append(el('div', 'title', m.title));
    var prep = el('div', 'prep');
    prep.append(el('span', 'prep-k', 'להכנה'), document.createTextNode(m.prep || 'להגדיר מטרה ותוצר צפוי'));
    body.append(prep);
    var sub = el('div', 'sub', (m.attendees.length ? m.attendees.length + ' משתתפים' : '') + (m.location ? ' · ' + m.location : '') + (m.tentative ? ' · טנטטיבי' : ''));
    body.append(sub);
    var link = safeLink(m.link, 'פתיחה ביומן');
    if (link) body.append(link);
    li.append(time, body, el('div', 'acts'));
    return li;
  }

  function render() {
    var n = nowLocal();
    var view = activeView();
    var isNext = view === 'next';
    var plan = isNext ? state.preview : state.plan;
    var shownDate = isNext ? nextDate() : n.date;
    $('today').textContent = HE_DATE.format(new Date(shownDate + 'T12:00:00Z'));
    $('view-label').textContent = isNext ? 'הכנה ל' + nextLabel() : 'היום';
    $('v-next').textContent = nextLabel();
    $('v-today').setAttribute('aria-pressed', String(!isNext));
    $('v-next').setAttribute('aria-pressed', String(isNext));
    var stamp = isNext ? state.previewFetchedAt : state.fetchedAt;
    $('updated').textContent = stamp ? 'עודכן ' + HE_TIME.format(new Date(stamp)) : 'עוד לא עודכן';
    $('next-note').hidden = !isNext;
    $('l-tomorrow').textContent = isNext ? 'עוברות ליום שאחריו' : 'עוברות למחר';
    $('h-tomorrow').textContent = isNext ? 'לא נכנס ל' + nextLabel() : 'עובר למחר';

    renderSources();
    renderRules();

    var list = $('timeline');
    list.innerHTML = '';
    if (!plan) {
      if (isNext) $('empty-text').textContent = state.busy ? 'בונה את הלו"ז של ' + nextLabel() + '…' : 'הלו"ז של ' + nextLabel() + ' ייבנה בלחיצה על "רענון".';
      $('empty').hidden = false;
      $('summary').hidden = true;
      $('later').hidden = true;
      return;
    }
    $('empty').hidden = true;

    var today = plan.tasks.filter(function (t) { return t.start != null && t.status !== 'dismissed'; });
    var tomorrow = plan.tasks.filter(function (t) { return t.placement === 'tomorrow' && t.status === 'pending'; });
    var deferred = plan.tasks.filter(function (t) { return t.placement === 'deferred' && t.status === 'pending'; });
    var done = plan.tasks.filter(function (t) { return t.status === 'done'; }).length;

    // סיכום + מטריצה
    $('summary').hidden = false;
    $('s-meetings').textContent = plan.meetings.length;
    $('s-tasks').textContent = today.length;
    $('s-done').textContent = done;
    $('s-tomorrow').textContent = tomorrow.length;
    var counts = { Q1: 0, Q2: 0, Q3: 0, Q4: 0 };
    plan.tasks.forEach(function (t) { if (t.status === 'pending' && counts[t.quadrant] != null) counts[t.quadrant]++; });
    Object.keys(counts).forEach(function (q) { $('m-' + q).textContent = counts[q]; });
    var cfg = plan.cfg || CFG;
    $('hours').textContent = cfg.dayStart + '–' + cfg.dayEnd;

    if (plan.allDay && plan.allDay.length) {
      var ad = el('li', 'item allday');
      ad.append(el('div', 'time', 'כל היום'), el('div', 'body', plan.allDay.map(function (a) { return a.title; }).join(' · ')), el('div', 'acts'));
      list.append(ad);
    }

    var rows = plan.meetings.map(function (m) { return { at: m.start, end: m.end, node: function () { return meetingRow(m); } }; })
      .concat(today.map(function (t) { return { at: t.start, end: t.end, node: function () { return taskRow(t, n); } }; }))
      .sort(function (a, b) { return a.at - b.at || a.end - b.end; });

    var nowPlaced = isNext || plan.date !== n.date;
    rows.forEach(function (r) {
      if (!nowPlaced && r.at > n.min) {
        list.append(nowLine(n.min));
        nowPlaced = true;
      }
      list.append(r.node());
    });
    if (!nowPlaced) list.append(nowLine(n.min));
    if (!rows.length) list.append(el('li', 'item none', 'אין פגישות או משימות משובצות להיום.'));

    // מחר ונדחה
    $('later').hidden = !(tomorrow.length || deferred.length);
    fillSimple('tomorrow-list', tomorrow, false);
    fillSimple('deferred-list', deferred, true);
    $('tomorrow-wrap').hidden = !tomorrow.length;
    $('deferred-wrap').hidden = !deferred.length;
  }

  function nowLine(min) {
    var li = el('li', 'now');
    li.append(el('span', '', 'עכשיו ' + E.fmt(min)));
    return li;
  }

  function fillSimple(id, tasks, canForce) {
    var ul = $(id);
    ul.innerHTML = '';
    tasks.slice(0, 40).forEach(function (t) {
      var li = el('li', 'mini');
      var txt = el('div', 'mini-t');
      txt.append(el('b', '', t.nextAction || t.title), el('span', 'muted', (t.from ? t.from + ' · ' : '') + E.QUADRANTS[t.quadrant].label + ' · ' + t.minutes + ' דק\''));
      var acts = el('div', 'acts');
      if (canForce) {
        var f = el('button', 'btn ghost small', 'לשבץ היום');
        f.onclick = function () { force(t.id); };
        acts.append(f);
      }
      var x = el('button', 'btn ghost small', canForce ? 'סגירה' : 'לא רלוונטי');
      x.onclick = function () { mark(t.id, 'dismissed'); };
      acts.append(x);
      li.append(txt, acts);
      ul.append(li);
    });
  }

  function renderSources() {
    var S = [
      ['calendar', 'יומן Outlook'],
      ['mail', 'מייל Outlook'],
      ['ai', 'תעדוף חכם (Claude)']
    ];
    var ul = $('sources');
    ul.innerHTML = '';
    S.forEach(function (s) {
      var st = state.src[s[0]];
      var li = el('li', 'src st-' + st);
      var label = st === 'ok' ? 'מחובר' : st === 'error' ? 'לא זמין' : st === 'fallback' ? 'בסיסי' : 'ממתין';
      li.append(el('span', 'dot'), el('b', '', s[1]), el('span', 'st', label));
      if (state.srcMsg[s[0]]) li.append(el('p', 'src-msg', state.srcMsg[s[0]]));
      ul.append(li);
    });
    var wa = el('li', 'src st-manual');
    wa.append(el('span', 'dot'), el('b', '', 'וואטסאפ'), el('span', 'st', 'הדבקה ידנית'));
    wa.append(el('p', 'src-msg', 'אין חיבור רשמי לקריאת וואטסאפ אישי. הדביקי הודעות בתיבה "הוספת משימות".'));
    var todo = el('li', 'src st-manual');
    todo.append(el('span', 'dot'), el('b', '', 'רשימת משימות'), el('span', 'st', 'בתוך האפליקציה'));
    todo.append(el('p', 'src-msg', 'המחבר של Microsoft 365 לא חושף את Microsoft To Do. משימות נוספות כאן.'));
    ul.append(wa, todo);
    $('perm-btn').hidden = !(cap.permissions && (state.src.calendar === 'error' || state.src.mail === 'error' || state.src.ai === 'fallback'));
  }

  function renderRules() {
    var ul = $('rules');
    ul.innerHTML = '';
    $('rules-empty').hidden = state.rules.length > 0;
    state.rules.slice().reverse().forEach(function (r) {
      var li = el('li', 'rule' + (r.active === false ? ' off' : ''));
      var txt = el('div', 'rule-body');
      txt.append(el('b', '', '"' + r.text + '"'), el('span', 'muted', r.explain || E.explainRule(r)));
      var b = el('button', 'btn ghost small', r.active === false ? 'הפעלה' : 'השבתה');
      b.onclick = function () { toggleRule(r.id); };
      li.append(txt, b);
      ul.append(li);
    });
    $('export').hidden = !cap.downloads || !state.rules.length;
  }

  // ---------- הפעלה ----------

  async function boot() {
    document.documentElement.setAttribute('dir', 'rtl');
    document.documentElement.setAttribute('lang', 'he');
    render();
    $('rule-form').addEventListener('submit', onRuleSubmit);
    $('paste-form').addEventListener('submit', onPasteSubmit);
    $('refresh').addEventListener('click', function () { buildDay(); });
    $('v-today').addEventListener('click', function () { setView('today'); });
    $('v-next').addEventListener('click', function () { setView('next'); });
    $('export').addEventListener('click', exportRules);
    $('perm-btn').addEventListener('click', function () { cap.permissions && cap.permissions.manage().catch(function () {}); });

    if (!window.claude || !window.claude.use) {
      $('empty-text').textContent = 'האפליקציה עובדת רק כשהיא פתוחה בתוך claude.ai, שם יש לה גישה ל-Outlook.';
      return;
    }
    var names = ['db', 'mcp', 'sample', 'downloads', 'user', 'permissions'];
    var got = await Promise.all(names.map(function (n) { return window.claude.use(n).catch(function () { return null; }); }));
    names.forEach(function (n, i) { cap[n] = got[i]; });

    if (cap.user) { try { state.uid = await cap.user.id(); } catch (e) { state.uid = null; } }
    if (!cap.db || !state.uid) toast('השמירה לא זמינה בתצוגה הזו. הלו"ז יחושב, אבל סימונים וכללים לא יישמרו.');
    await loadState();
    render();

    var n = nowLocal();
    if (!state.plan || state.plan.date !== n.date) await buildDay();
    else {
      // אותו יום: לרענן אם עברה יותר משעה מהמשיכה האחרונה
      var stale = !state.fetchedAt || Date.now() - new Date(state.fetchedAt).getTime() > 3600000;
      if (stale) await buildDay(); else { replanNow(); render(); }
    }
    if (activeView() === 'next') {
      var pStale = !state.previewFetchedAt || Date.now() - new Date(state.previewFetchedAt).getTime() > 3600000;
      if (!state.preview || state.preview.date !== nextDate() || pStale) await buildPreview();
    }

    // בדיקה כל דקה: קו "עכשיו", ובניית יום חדש אוטומטית בבוקר
    setInterval(function () {
      var cur = nowLocal();
      if (state.plan && state.plan.date !== cur.date && cur.min >= E.hm('05:00') && !document.hidden) buildDay();
      else if (!state.busy && activeView() === 'next' && !document.hidden &&
        (!state.preview || state.preview.date !== nextDate())) buildPreview();
      else if (!state.busy) render();
    }, 60000);
    document.addEventListener('visibilitychange', function () {
      if (!document.hidden && state.plan && state.plan.date !== nowLocal().date) buildDay();
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
