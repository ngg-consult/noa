/*
 * MorningEngine – לוגיקת התעדוף והשיבוץ של "סדר היום".
 * פונקציות טהורות בלבד (ללא DOM וללא רשת), כדי שאפשר יהיה לבדוק אותן ב-Node.
 * כל הזמנים בתוך היום נשמרים כדקות מחצות לפי אזור הזמן של המשתמשת.
 */
(function (root) {
  'use strict';

  var DEFAULTS = {
    tz: 'Asia/Jerusalem',
    dayStart: '08:30',
    dayEnd: '18:00',
    bufferMin: 5,      // מרווח לפני ואחרי כל פגישה
    minSlot: 10,       // חלון קצר מזה לא משמש לשיבוץ
    mailLookbackDays: 3
  };

  var CONTEXTS = ['מחשב', 'טלפון', 'פגישה', 'בחוץ'];

  var QUADRANTS = {
    Q1: { label: 'דחוף וחשוב', action: 'לבצע היום', base: 100 },
    Q2: { label: 'חשוב, לא דחוף', action: 'לתכנן ולבצע', base: 70 },
    Q3: { label: 'דחוף, לא חשוב', action: 'להאציל', base: 45 },
    Q4: { label: 'לא דחוף ולא חשוב', action: 'לוותר או לדחות', base: 10 }
  };

  // ---------- זמן ----------

  function hm(s) {
    var p = String(s).split(':');
    return Number(p[0]) * 60 + Number(p[1] || 0);
  }

  function fmt(min) {
    min = Math.max(0, Math.round(min));
    var h = Math.floor(min / 60), m = min % 60;
    return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
  }

  function roundUp(n, step) { return Math.ceil(n / step) * step; }

  var dtfCache = {};
  function dtf(tz) {
    if (!dtfCache[tz]) {
      dtfCache[tz] = new Intl.DateTimeFormat('en-CA', {
        timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', hourCycle: 'h23', weekday: 'short'
      });
    }
    return dtfCache[tz];
  }

  // תאריך מקומי (YYYY-MM-DD), דקה ביום ויום בשבוע (0=ראשון) של רגע נתון
  function localParts(date, tz) {
    var parts = {};
    dtf(tz || DEFAULTS.tz).formatToParts(date).forEach(function (p) { parts[p.type] = p.value; });
    var wd = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].indexOf(parts.weekday);
    return {
      date: parts.year + '-' + parts.month + '-' + parts.day,
      min: Number(parts.hour) * 60 + Number(parts.minute),
      weekday: wd
    };
  }

  function addDays(dateStr, n) {
    var d = new Date(dateStr + 'T12:00:00Z');
    d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
  }

  // Graph מחזיר {dateTime, timeZone}; כשהאזור הוא UTC מוסיפים Z.
  // אזור אחר מתפרש כשעון-קיר באזור הזמן של האפליקציה.
  function graphTime(t, tz) {
    if (!t || !t.dateTime) return null;
    var s = String(t.dateTime).slice(0, 19);
    var zone = String(t.timeZone || 'UTC');
    if (/^(UTC|GMT|Etc\/UTC|Coordinated Universal Time)$/i.test(zone)) {
      return localParts(new Date(s + 'Z'), tz);
    }
    return { date: s.slice(0, 10), min: hm(s.slice(11, 16)), weekday: new Date(s.slice(0, 10) + 'T12:00:00Z').getUTCDay() };
  }

  // ---------- קריאת תוצאות המחבר ----------

  // מחבר Microsoft 365 מחזיר כל פריט כבלוק טקסט JSON נפרד, והאחרון מכיל nextOffset
  function parseToolBlocks(result) {
    var items = [], nextOffset = null;
    var blocks = (result && result.content) || [];
    blocks.forEach(function (b) {
      if (!b || b.type !== 'text' || typeof b.text !== 'string') return;
      var obj;
      try { obj = JSON.parse(b.text); } catch (e) { return; }
      var list = Array.isArray(obj) ? obj : [obj];
      list.forEach(function (o) {
        if (!o || typeof o !== 'object') return;
        if (o.id || o.uri) items.push(o);
        else if (o.nextOffset != null) nextOffset = o.nextOffset;
      });
    });
    if (!items.length && result && Array.isArray(result.payload)) {
      result.payload.forEach(function (o) { if (o && (o.id || o.uri)) items.push(o); });
    }
    return { items: items, nextOffset: nextOffset };
  }

  // ---------- נרמול ----------

  function normalizeEvents(raw, dateStr, cfg) {
    cfg = cfg || DEFAULTS;
    var out = [], allDay = [];
    (raw || []).forEach(function (e) {
      if (!e || e.isCancelled) return;
      var s = graphTime(e.start, cfg.tz), en = graphTime(e.end, cfg.tz);
      if (!s || !en) return;
      if (e.isAllDay) {
        if (s.date <= dateStr && en.date >= dateStr) allDay.push({ id: e.id, title: e.subject || 'אירוע' });
        return;
      }
      if (s.date > dateStr || en.date < dateStr) return;
      var start = s.date < dateStr ? 0 : s.min;
      var end = en.date > dateStr ? 24 * 60 : en.min;
      if (end <= start) return;
      out.push({
        id: e.id,
        kind: 'meeting',
        title: e.subject || 'פגישה',
        start: start,
        end: end,
        organizer: e.organizer || '',
        attendees: e.attendees || [],
        location: e.location || '',
        link: e.webLink || '',
        tentative: e.showAs === 'tentative',
        free: e.showAs === 'free',
        prep: ''
      });
    });
    out.sort(function (a, b) { return a.start - b.start || a.end - b.end; });
    return { meetings: out, allDay: allDay };
  }

  var INVITE_SUBJECT = /^(accepted|declined|tentative|canceled|cancelled|updated invitation|invitation|התקבל|נדחה|מבוטל|בוטל|הזמנה|אושר|מקובל|טנטטיבי)\b/i;
  var INVITE_BODY = /(Microsoft Teams Need help|Join the meeting now|פגישת Microsoft Teams|Microsoft Teams meeting|להצטרף אל: https:\/\/teams)/i;

  function isInvite(m) {
    return INVITE_SUBJECT.test(String(m.subject || '').trim()) || INVITE_BODY.test(String(m.summary || ''));
  }

  function normalizeEmails(raw, me) {
    var mine = String(me || '').toLowerCase();
    var seen = {};
    var out = [];
    (raw || []).forEach(function (m) {
      if (!m || !m.id || seen[m.id]) return;
      seen[m.id] = true;
      var from = String(m.sender || '');
      if (mine && from.toLowerCase() === mine) return;
      if (isInvite(m)) return;
      out.push({
        id: 'mail:' + m.id,
        source: 'mail',
        title: String(m.subject || '(ללא נושא)').trim(),
        from: from,
        receivedAt: m.receivedDateTime || m.sentDateTime || null,
        snippet: String(m.summary || '').replace(/\s+/g, ' ').trim().slice(0, 300),
        link: m.webLink || '',
        importance: m.importance || 'normal',
        isRead: !!m.isRead
      });
    });
    return out;
  }

  // ---------- סיווג ללא AI (גיבוי) ----------

  var URGENT_WORDS = /(דחוף|בהקדם|היום|עד מחר|מיידי|ASAP|urgent|immediately|EOD|deadline|דדליין|תזכורת)/i;
  var NOISE_WORDS = /(newsletter|unsubscribe|ניוזלטר|להסרה|הסר מרשימת|no-?reply|noreply|webinar|וובינר|מבצע|promotion)/i;

  function heuristicClassify(item) {
    var text = [item.title, item.snippet, item.from].join(' ');
    var noise = NOISE_WORDS.test(text);
    var urgent = item.importance === 'high' || URGENT_WORDS.test(text);
    var phone = /(להתקשר|שיחה|טלפון|call me|תתקשרי|תחזרי אליי)/i.test(text);
    return {
      actionable: !noise,
      nextAction: (phone ? 'להתקשר בנושא: ' : 'לעבור ולהשיב: ') + item.title,
      minutes: 15,
      context: phone ? 'טלפון' : 'מחשב',
      urgent: urgent,
      important: !noise,
      delegateTo: null,
      client: null,
      classifiedBy: 'heuristic'
    };
  }

  // מנקה ומגביל תשובה שהגיעה מ-AI, כך שערכים לא תקינים לא ישברו את השיבוץ
  function sanitizeClassification(c, item) {
    var h = heuristicClassify(item);
    if (!c || typeof c !== 'object') return h;
    var minutes = Number(c.minutes);
    if (!isFinite(minutes) || minutes <= 0) minutes = h.minutes;
    minutes = Math.min(240, Math.max(2, Math.round(minutes)));
    return {
      actionable: c.actionable !== false,
      nextAction: typeof c.nextAction === 'string' && c.nextAction.trim() ? c.nextAction.trim().slice(0, 160) : h.nextAction,
      minutes: minutes,
      context: CONTEXTS.indexOf(c.context) >= 0 ? c.context : h.context,
      urgent: !!c.urgent,
      important: c.important === undefined ? h.important : !!c.important,
      delegateTo: typeof c.delegateTo === 'string' && c.delegateTo.trim() ? c.delegateTo.trim().slice(0, 60) : null,
      client: typeof c.client === 'string' && c.client.trim() ? c.client.trim().slice(0, 60) : null,
      classifiedBy: 'ai'
    };
  }

  function makeTask(item, cls) {
    var t = {};
    Object.keys(item).forEach(function (k) { t[k] = item[k]; });
    Object.keys(cls).forEach(function (k) { t[k] = cls[k]; });
    t.kind = 'task';
    t.status = t.status || 'pending';
    t.missedCount = t.missedCount || 0;
    return t;
  }

  // ---------- כללים ----------

  function ruleMatches(rule, task) {
    var m = rule.match || {};
    var hay = [task.title, task.from, task.snippet, task.client, task.nextAction, task.delegateTo]
      .join(' ').toLowerCase();
    var groups = 0;
    if (m.keywords && m.keywords.length) {
      groups++;
      if (!m.keywords.some(function (k) { return k && hay.indexOf(String(k).toLowerCase()) >= 0; })) return false;
    }
    if (m.senders && m.senders.length) {
      groups++;
      var from = String(task.from || '').toLowerCase();
      if (!m.senders.some(function (s) { return s && from.indexOf(String(s).toLowerCase()) >= 0; })) return false;
    }
    if (m.contexts && m.contexts.length) {
      groups++;
      if (m.contexts.indexOf(task.context) < 0) return false;
    }
    if (m.sources && m.sources.length) {
      groups++;
      if (m.sources.indexOf(task.source) < 0) return false;
    }
    if (m.quadrants && m.quadrants.length) {
      groups++;
      if (m.quadrants.indexOf(quadrantOf(task)) < 0) return false;
    }
    if (m.maxMinutes) {
      groups++;
      if (!(task.minutes <= m.maxMinutes)) return false;
    }
    if (m.minMinutes) {
      groups++;
      if (!(task.minutes >= m.minMinutes)) return false;
    }
    return groups > 0 || m.all === true;
  }

  // הגדרות יום (שעת התחלה/סיום, מרווח) שנקבעו בכללים – הכלל האחרון גובר
  function settingsFromRules(rules, base) {
    var s = {};
    Object.keys(base || DEFAULTS).forEach(function (k) { s[k] = (base || DEFAULTS)[k]; });
    (rules || []).forEach(function (r) {
      if (r.active === false || !r.effect || !r.effect.settings) return;
      var e = r.effect.settings;
      if (e.dayStart) s.dayStart = e.dayStart;
      if (e.dayEnd) s.dayEnd = e.dayEnd;
      if (e.bufferMin != null) s.bufferMin = Number(e.bufferMin);
    });
    return s;
  }

  function applyRules(tasks, rules) {
    return tasks.map(function (t0) {
      var t = {};
      Object.keys(t0).forEach(function (k) { t[k] = t0[k]; });
      t.ruleBoost = 0;
      t.window = null;
      t.appliedRules = [];
      t.skipped = false;
      (rules || []).forEach(function (r) {
        if (r.active === false || !r.effect || !ruleMatches(r, t)) return;
        var e = r.effect;
        t.appliedRules.push(r.id);
        if (typeof e.priority === 'number') t.ruleBoost += e.priority;
        if (e.urgent != null) t.urgent = !!e.urgent;
        if (e.important != null) t.important = !!e.important;
        if (e.minutes) t.minutes = Number(e.minutes);
        if (e.context && CONTEXTS.indexOf(e.context) >= 0) t.context = e.context;
        if (e.delegateTo) t.delegateTo = e.delegateTo;
        if (e.skip) t.skipped = true;
        if (e.window && (e.window.start || e.window.end)) {
          t.window = { start: e.window.start || null, end: e.window.end || null };
        }
      });
      return t;
    });
  }

  // ---------- תעדוף ----------

  function quadrantOf(t) {
    if (t.urgent && t.important) return 'Q1';
    if (t.important) return 'Q2';
    if (t.urgent) return 'Q3';
    return 'Q4';
  }

  // ככל שהפריט חדש יותר – תוספת גבוהה יותר (דעיכה של 24 שעות)
  function recencyBoost(receivedAt, nowMs) {
    if (!receivedAt) return 0;
    var age = (nowMs - new Date(receivedAt).getTime()) / 3600000;
    if (!isFinite(age)) return 0;
    if (age < 0) age = 0;
    return 30 * Math.exp(-age / 24);
  }

  function scoreTask(t, nowMs) {
    var q = quadrantOf(t);
    var s = QUADRANTS[q].base + recencyBoost(t.receivedAt, nowMs) + (t.ruleBoost || 0);
    if (t.minutes <= 2) s += 8;               // GTD: פחות משתי דקות – לעשות מיד
    if (t.carried) s += 10 + 5 * Math.min(3, t.missedCount || 0); // לא להזניח את מה שנדחה
    return Math.round(s * 10) / 10;
  }

  // משך בפועל ביומן: GTD – משימה לא דחופה-לא חשובה שמאצילים לוקחת רק 5 דקות של העברה
  function slotMinutes(t) {
    var m = t.minutes || 15;
    if (quadrantOf(t) === 'Q3' && t.delegateTo) m = 5;
    return Math.max(5, roundUp(m, 5));
  }

  // ---------- חלונות פנויים ושיבוץ ----------

  function mergeBusy(intervals) {
    var list = intervals.slice().sort(function (a, b) { return a[0] - b[0]; });
    var out = [];
    list.forEach(function (iv) {
      var last = out[out.length - 1];
      if (last && iv[0] <= last[1]) last[1] = Math.max(last[1], iv[1]);
      else out.push([iv[0], iv[1]]);
    });
    return out;
  }

  function freeWindows(meetings, cfg, fromMin) {
    var dayStart = hm(cfg.dayStart), dayEnd = hm(cfg.dayEnd);
    var start = Math.max(dayStart, roundUp(fromMin || 0, 5));
    var buf = cfg.bufferMin || 0;
    var busy = mergeBusy(meetings.filter(function (m) { return !m.free; }).map(function (m) {
      var b = m.kind === 'meeting' ? buf : 0; // מרווח רק סביב פגישות
      return [m.start - b, m.end + b];
    }));
    var wins = [], cur = start;
    busy.forEach(function (b) {
      if (b[1] <= cur) return;
      if (b[0] > cur) wins.push({ start: cur, end: Math.min(b[0], dayEnd) });
      cur = Math.max(cur, b[1]);
    });
    if (cur < dayEnd) wins.push({ start: cur, end: dayEnd });
    return wins.filter(function (w) { return w.end - w.start >= (cfg.minSlot || 10); });
  }

  function byPriority(a, b) {
    return (b.score - a.score) ||
      (new Date(b.receivedAt || 0) - new Date(a.receivedAt || 0)) ||
      String(a.id).localeCompare(String(b.id));
  }

  // שיבוץ חמדני: המשימה בעדיפות הגבוהה ביותר מקבלת את החלון המוקדם ביותר שמתאים לה
  function schedule(tasks, windows) {
    var wins = windows.map(function (w) { return { start: w.start, end: w.end }; });
    var placed = [], overflow = [];
    tasks.slice().sort(byPriority).forEach(function (t) {
      var dur = slotMinutes(t);
      var lo = t.window && t.window.start ? hm(t.window.start) : 0;
      var hi = t.window && t.window.end ? hm(t.window.end) : 24 * 60;
      if (t.notBefore) lo = Math.max(lo, t.notBefore);
      for (var i = 0; i < wins.length; i++) {
        var w = wins[i];
        var s = Math.max(w.start, lo), e = Math.min(w.end, hi);
        if (e - s >= dur) {
          placed.push({ id: t.id, start: s, end: s + dur });
          var pieces = [];
          if (s - w.start >= 5) pieces.push({ start: w.start, end: s });
          if (w.end - (s + dur) >= 5) pieces.push({ start: s + dur, end: w.end });
          wins.splice.apply(wins, [i, 1].concat(pieces));
          return;
        }
      }
      overflow.push(t.id);
    });
    return { placed: placed, overflow: overflow };
  }

  // ---------- תכנון מלא ----------

  /*
   * plan: { date, meetings, tasks } – tasks כוללות status: pending|done|missed|dismissed.
   * מחשבת מחדש ציון ושיבוץ לכל מה שלא בוצע, החל מ-fromMin (לרוב "עכשיו").
   * משימות שבוצעו נשארות במקומן לתיעוד.
   */
  function replan(plan, rules, opts) {
    var cfg = settingsFromRules(rules, opts.cfg || DEFAULTS);
    var nowMs = opts.nowMs;
    var fromMin = opts.fromMin || 0;
    var ruled = applyRules(plan.tasks, rules);
    var open = [], keep = [], deferred = [];
    ruled.forEach(function (t) {
      t.quadrant = quadrantOf(t);
      t.score = scoreTask(t, nowMs);
      if (t.status === 'done' || t.status === 'dismissed') { keep.push(t); return; }
      if (!t.forced && (t.actionable === false || t.skipped || t.quadrant === 'Q4')) {
        t.start = t.end = null;
        t.placement = 'deferred';
        deferred.push(t);
        return;
      }
      open.push(t);
    });

    // חלונות פנויים: גם משימות שכבר בוצעו תופסות את זמנן
    var blocks = plan.meetings.concat(keep.filter(function (t) { return t.start != null; }).map(function (t) {
      return { start: t.start, end: t.end };
    }));
    var wins = freeWindows(blocks, cfg, fromMin);
    var res = schedule(open, wins);
    var at = {};
    res.placed.forEach(function (p) { at[p.id] = p; });
    open.forEach(function (t) {
      if (at[t.id]) { t.start = at[t.id].start; t.end = at[t.id].end; t.placement = 'today'; }
      else { t.start = t.end = null; t.placement = 'tomorrow'; }
    });

    var tasks = keep.concat(open, deferred).sort(function (a, b) {
      var as = a.start == null ? 1e9 : a.start, bs = b.start == null ? 1e9 : b.start;
      return as - bs || byPriority(a, b);
    });
    return {
      date: plan.date,
      meetings: plan.meetings,
      allDay: plan.allDay || [],
      tasks: tasks,
      cfg: cfg,
      plannedAt: new Date(nowMs).toISOString()
    };
  }

  function markTask(plan, id, status, nowMin) {
    var tasks = plan.tasks.map(function (t) {
      if (t.id !== id) return t;
      var c = {};
      Object.keys(t).forEach(function (k) { c[k] = t[k]; });
      c.status = status;
      if (status === 'done') {
        // בוצעה: נרשמת בזמן שבו סומנה (או בחלון ששובץ לה, אם עדיין לא עבר)
        var dur = c.end != null && c.start != null ? c.end - c.start : slotMinutes(c);
        if (c.start == null || c.start > nowMin) { c.end = roundUp(nowMin, 5); c.start = Math.max(0, c.end - dur); }
        else if (nowMin < c.end) c.end = Math.max(c.start + 5, roundUp(nowMin, 5)); // הסתיימה מוקדם – שאר הזמן מתפנה
        c.doneAt = new Date().toISOString();
      }
      if (status === 'missed') {
        c.missedCount = (c.missedCount || 0) + 1;
        c.notBefore = c.end != null ? Math.max(c.end, nowMin) : nowMin; // עוברת לשעה פנויה מאוחרת יותר
        c.status = 'pending';
        c.carried = true;
      }
      if (status === 'pending') { c.doneAt = null; }
      return c;
    });
    return { date: plan.date, meetings: plan.meetings, allDay: plan.allDay, tasks: tasks };
  }

  // מה עובר מהתוכנית של אתמול לבוקר הבא
  function carryOver(prevPlan) {
    if (!prevPlan || !prevPlan.tasks) return [];
    return prevPlan.tasks.filter(function (t) {
      return t.status !== 'done' && t.status !== 'dismissed' && t.actionable !== false;
    }).map(function (t) {
      var c = {};
      Object.keys(t).forEach(function (k) { c[k] = t[k]; });
      c.carried = true;
      c.status = 'pending';
      c.start = c.end = null;
      c.notBefore = null;
      return c;
    });
  }

  // ---------- פירוש כלל ללא AI (גיבוי) ----------

  var DAYPARTS = [
    { re: /(בבוקר|בוקר)/, start: '08:00', end: '12:00', label: 'בוקר' },
    { re: /(בצהריים|צהריים|בצהרים|צהרים)/, start: '12:00', end: '15:00', label: 'צהריים' },
    { re: /(אחה"צ|אחר הצהריים|אחרי הצהריים|אחה״צ)/, start: '14:00', end: '18:00', label: 'אחה"צ' },
    { re: /(בערב|ערב)/, start: '17:00', end: '21:00', label: 'ערב' },
    { re: /(סוף היום|בסוף היום)/, start: '16:00', end: '18:00', label: 'סוף היום' }
  ];

  function parseRuleText(text) {
    var t = String(text || '').trim();
    var rule = { text: t, match: {}, effect: {}, explain: '' };
    var mTime = t.match(/(?:מתחיל[הת]?|להתחיל)[^0-9]*(\d{1,2}(?::\d{2})?)/);
    var eTime = t.match(/(?:מסיימ[הת]?|לסיים|עד)\s*(?:ב-?|בשעה\s*)?(\d{1,2}(?::\d{2})?)/);
    if (mTime || (eTime && /יום/.test(t))) {
      var norm = function (x) { return /:/.test(x) ? x.padStart(5, '0') : String(x).padStart(2, '0') + ':00'; };
      rule.effect.settings = {};
      if (mTime) rule.effect.settings.dayStart = norm(mTime[1]);
      if (eTime) rule.effect.settings.dayEnd = norm(eTime[1]);
      rule.match.all = true;
      rule.explain = 'שעות העבודה: ' + (rule.effect.settings.dayStart || DEFAULTS.dayStart) + '–' + (rule.effect.settings.dayEnd || DEFAULTS.dayEnd);
      return rule;
    }
    var about = t.match(/(?:קשור|שקשור|הקשור|שקשורים|קשורות)\s*ל[-־]?\s*(.+?)(?:\s+(?:הכי|הוא|היא|זה|יהיה|תמיד|רק|קודם|בעדיפות)|[,.]|$)/);
    var subject = about ? about[1].replace(/^(לקוח|הלקוח|פרויקט|הפרויקט)\s+/, '').trim() : null;
    if (subject) rule.match.keywords = [subject];
    CONTEXTS.forEach(function (c) {
      var re = c === 'טלפון' ? /(טלפון|שיחות|שיחה|להתקשר)/ : c === 'פגישה' ? /(פגישה|פגישות)/ : c === 'בחוץ' ? /(בחוץ|סידורים)/ : /(מחשב)/;
      if (re.test(t) && !subject) { rule.match.contexts = (rule.match.contexts || []).concat([c]); }
    });
    if (/(קצרות|קצרים|קטנות)/.test(t) && !subject) rule.match.maxMinutes = 15;
    if (/(ארוכות|עמוקה|ריכוז)/.test(t) && !subject) rule.match.minMinutes = 45;
    if (/(הכי דחוף|בראש|קודם כל|עדיפות עליונה|הכי חשוב)/.test(t)) {
      rule.effect.priority = 200; rule.effect.urgent = true; rule.effect.important = true;
    } else if (/(דחוף|חשוב|עדיפות גבוהה)/.test(t) && !/לא\s+(דחוף|חשוב)/.test(t)) {
      rule.effect.priority = 40;
    } else if (/(לא דחוף|לא חשוב|בסוף|אחרון|עדיפות נמוכה)/.test(t)) {
      rule.effect.priority = -40;
    }
    if (/(לדלג|להתעלם|לא להציג|לא לשבץ)/.test(t)) rule.effect.skip = true;
    DAYPARTS.forEach(function (d) {
      if (!rule.effect.window && d.re.test(t) && /(רק|תמיד|לשבץ|ב)/.test(t)) {
        rule.effect.window = { start: d.start, end: d.end };
      }
    });
    var range = t.match(/בין\s*(\d{1,2}(?::\d{2})?)\s*(?:ל|-|עד)\s*-?(\d{1,2}(?::\d{2})?)/);
    if (range) {
      var n = function (x) { return /:/.test(x) ? x.padStart(5, '0') : String(x).padStart(2, '0') + ':00'; };
      rule.effect.window = { start: n(range[1]), end: n(range[2]) };
    }
    var dele = t.match(/(?:להאציל|להעביר|לתת)\s+ל[-־]?\s*([^\s,.]+)/);
    if (dele) rule.effect.delegateTo = dele[1];
    rule.explain = explainRule(rule);
    return rule;
  }

  function explainRule(rule) {
    var m = rule.match || {}, e = rule.effect || {}, who = [];
    if (m.keywords && m.keywords.length) who.push('כל מה שקשור ל"' + m.keywords.join('" / "') + '"');
    if (m.senders && m.senders.length) who.push('מהשולחים ' + m.senders.join(', '));
    if (m.contexts && m.contexts.length) who.push('משימות ' + m.contexts.join('/'));
    if (m.maxMinutes) who.push('משימות של עד ' + m.maxMinutes + ' דק\'');
    if (m.minMinutes) who.push('משימות של ' + m.minMinutes + ' דק\' ומעלה');
    if (m.quadrants && m.quadrants.length) who.push('משימות ברבע ' + m.quadrants.join('/'));
    var what = [];
    if (e.priority >= 150) what.push('בעדיפות עליונה');
    else if (e.priority > 0) what.push('בעדיפות גבוהה יותר');
    else if (e.priority < 0) what.push('בעדיפות נמוכה יותר');
    if (e.window) what.push('רק בין ' + (e.window.start || 'תחילת היום') + ' ל-' + (e.window.end || 'סוף היום'));
    if (e.minutes) what.push(e.minutes + ' דקות');
    if (e.delegateTo) what.push('להאציל ל' + e.delegateTo);
    if (e.skip) what.push('לא לשבץ');
    if (e.settings) what.push('שעות עבודה ' + (e.settings.dayStart || '') + '–' + (e.settings.dayEnd || ''));
    if (!who.length && !what.length) return '';
    return (who.join(', ') || 'כל המשימות') + ' → ' + (what.join(', ') || 'ללא שינוי');
  }

  function isUsableRule(rule) {
    if (!rule || !rule.effect) return false;
    var e = rule.effect, m = rule.match || {};
    var hasEffect = typeof e.priority === 'number' || e.window || e.minutes || e.delegateTo || e.skip ||
      e.settings || e.urgent != null || e.important != null || e.context;
    var hasMatch = m.all || (m.keywords && m.keywords.length) || (m.senders && m.senders.length) ||
      (m.contexts && m.contexts.length) || (m.sources && m.sources.length) ||
      (m.quadrants && m.quadrants.length) || m.maxMinutes || m.minMinutes;
    return !!(hasEffect && hasMatch);
  }

  // הכנה לפגישה בלי AI: מבוסס על המשתתפים והמיילים האחרונים מהם
  function heuristicPrep(meeting, items) {
    var people = (meeting.attendees || []).map(function (a) { return String(a).toLowerCase(); });
    var related = (items || []).filter(function (i) {
      return i.from && people.indexOf(String(i.from).toLowerCase()) >= 0;
    });
    if (related.length) return 'לעבור על "' + related[0].title + '" לפני הפגישה';
    var ext = (meeting.attendees || []).filter(function (a) { return !/nggconsult\.com$/i.test(a); });
    if (ext.length) return 'לרענן את ההתכתבות האחרונה עם ' + ext[0].split('@')[1];
    return 'להגדיר מטרה אחת ותוצר צפוי לפגישה';
  }

  var api = {
    DEFAULTS: DEFAULTS, CONTEXTS: CONTEXTS, QUADRANTS: QUADRANTS,
    hm: hm, fmt: fmt, localParts: localParts, addDays: addDays, graphTime: graphTime,
    parseToolBlocks: parseToolBlocks, normalizeEvents: normalizeEvents, normalizeEmails: normalizeEmails,
    isInvite: isInvite, heuristicClassify: heuristicClassify, sanitizeClassification: sanitizeClassification,
    makeTask: makeTask, ruleMatches: ruleMatches, applyRules: applyRules, settingsFromRules: settingsFromRules,
    quadrantOf: quadrantOf, recencyBoost: recencyBoost, scoreTask: scoreTask, slotMinutes: slotMinutes,
    freeWindows: freeWindows, schedule: schedule, replan: replan, markTask: markTask, carryOver: carryOver,
    parseRuleText: parseRuleText, explainRule: explainRule, isUsableRule: isUsableRule, heuristicPrep: heuristicPrep
  };
  root.MorningEngine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
