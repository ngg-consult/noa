// בדיקות ללוגיקת התעדוף והשיבוץ. הרצה: node --test test/
const test = require('node:test');
const assert = require('node:assert/strict');
const E = require('../app/engine.js');

const CFG = Object.assign({}, E.DEFAULTS, { dayStart: '08:30', dayEnd: '18:00', bufferMin: 5 });
const NOW = Date.parse('2026-10-11T05:00:00Z'); // 08:00 שעון ישראל

function task(id, o) {
  return Object.assign({
    id, kind: 'task', source: 'mail', title: id, from: 'x@client.com', snippet: '',
    receivedAt: '2026-10-11T04:00:00Z', minutes: 30, context: 'מחשב',
    urgent: false, important: true, actionable: true, status: 'pending'
  }, o);
}

test('graphTime converts UTC to Israel wall clock', () => {
  const p = E.graphTime({ dateTime: '2026-10-11T07:00:00.0000000', timeZone: 'UTC' }, 'Asia/Jerusalem');
  assert.equal(p.date, '2026-10-11');
  assert.equal(E.fmt(p.min), '10:00');
});

test('parseToolBlocks reads one JSON object per text block plus the paging block', () => {
  const r = E.parseToolBlocks({ content: [
    { type: 'text', text: JSON.stringify({ id: 'a', subject: 's' }) },
    { type: 'text', text: JSON.stringify({ id: 'b', subject: 't' }) },
    { type: 'text', text: JSON.stringify({ moreResults: true, nextOffset: 2 }) }
  ] });
  assert.deepEqual(r.items.map(i => i.id), ['a', 'b']);
  assert.equal(r.nextOffset, 2);
});

test('normalizeEmails drops my own mail and meeting invitations', () => {
  const out = E.normalizeEmails([
    { id: '1', subject: 'הצעת מחיר', sender: 'a@b.com', summary: 'נא לשלוח' },
    { id: '2', subject: 'פגישה', sender: 'c@d.com', summary: 'Microsoft Teams Need help? Join the meeting now' },
    { id: '3', subject: 'שלי', sender: 'me@ngg.com', summary: '' }
  ], 'me@ngg.com');
  assert.deepEqual(out.map(o => o.id), ['mail:1']);
});

test('normalizeEvents skips cancelled and keeps the day only', () => {
  const r = E.normalizeEvents([
    { id: 'e1', subject: 'A', start: { dateTime: '2026-10-11T07:00:00', timeZone: 'UTC' }, end: { dateTime: '2026-10-11T08:00:00', timeZone: 'UTC' }, showAs: 'busy' },
    { id: 'e2', subject: 'B', isCancelled: true, start: { dateTime: '2026-10-11T09:00:00', timeZone: 'UTC' }, end: { dateTime: '2026-10-11T10:00:00', timeZone: 'UTC' } },
    { id: 'e3', subject: 'C', start: { dateTime: '2026-10-12T07:00:00', timeZone: 'UTC' }, end: { dateTime: '2026-10-12T08:00:00', timeZone: 'UTC' } }
  ], '2026-10-11', CFG);
  assert.deepEqual(r.meetings.map(m => [m.id, E.fmt(m.start), E.fmt(m.end)]), [['e1', '10:00', '11:00']]);
});

test('quadrants follow the Eisenhower matrix', () => {
  assert.equal(E.quadrantOf({ urgent: true, important: true }), 'Q1');
  assert.equal(E.quadrantOf({ urgent: false, important: true }), 'Q2');
  assert.equal(E.quadrantOf({ urgent: true, important: false }), 'Q3');
  assert.equal(E.quadrantOf({ urgent: false, important: false }), 'Q4');
});

test('newer items score higher within the same quadrant', () => {
  const newer = E.scoreTask(task('n', { receivedAt: '2026-10-11T04:30:00Z' }), NOW);
  const older = E.scoreTask(task('o', { receivedAt: '2026-10-09T04:30:00Z' }), NOW);
  assert.ok(newer > older);
});

test('free windows sit between meetings with buffers', () => {
  const wins = E.freeWindows([{ kind: 'meeting', start: E.hm('10:00'), end: E.hm('11:00') }], CFG, E.hm('08:00'));
  assert.deepEqual(wins.map(w => [E.fmt(w.start), E.fmt(w.end)]), [['08:30', '09:55'], ['11:05', '18:00']]);
});

test('replan places the most important task first and defers Q4', () => {
  const plan = {
    date: '2026-10-11',
    meetings: [{ id: 'm', kind: 'meeting', start: E.hm('09:30'), end: E.hm('10:30') }],
    tasks: [
      task('low', { important: false, urgent: false }),
      task('q2', {}),
      task('q1', { urgent: true })
    ]
  };
  const out = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  const byId = Object.fromEntries(out.tasks.map(t => [t.id, t]));
  assert.equal(E.fmt(byId.q1.start), '08:30');
  assert.equal(E.fmt(byId.q2.start), '10:35'); // 09:00–09:25 קצר מדי ל-30 דקות
  assert.equal(byId.low.placement, 'deferred');
});

test('a free-text rule boosts everything about a client', () => {
  const rule = E.parseRuleText('כל מה שקשור ללקוח רשות שדות התעופה הכי דחוף');
  assert.ok(E.isUsableRule(rule));
  assert.deepEqual(rule.match.keywords, ['רשות שדות התעופה']);
  const plan = {
    date: '2026-10-11', meetings: [],
    tasks: [task('a', { urgent: true }), task('b', { title: 'מסמך עבור רשות שדות התעופה', important: false })]
  };
  const out = E.replan(plan, [Object.assign({ id: 'r1' }, rule)], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  assert.equal(out.tasks.find(t => t.start === E.hm('08:30')).id, 'b');
});

test('a time-window rule keeps phone calls to lunchtime', () => {
  const rule = Object.assign({ id: 'r2' }, E.parseRuleText('שיחות טלפון רק בצהריים'));
  assert.deepEqual(rule.match.contexts, ['טלפון']);
  assert.deepEqual(rule.effect.window, { start: '12:00', end: '15:00' });
  const plan = { date: '2026-10-11', meetings: [], tasks: [task('call', { context: 'טלפון', urgent: true })] };
  const out = E.replan(plan, [rule], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  assert.equal(E.fmt(out.tasks[0].start), '12:00');
});

test('working-hours rule changes the day start', () => {
  const rule = Object.assign({ id: 'r3' }, E.parseRuleText('אני מתחילה לעבוד ב-9'));
  const cfg = E.settingsFromRules([rule], CFG);
  assert.equal(cfg.dayStart, '09:00');
});

test('missed task moves to a later free slot, then to tomorrow when the day is full', () => {
  let plan = { date: '2026-10-11', meetings: [], tasks: [task('t', { minutes: 60 })] };
  plan = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  assert.equal(E.fmt(plan.tasks[0].start), '08:30');
  plan = E.markTask(plan, 't', 'missed', E.hm('09:00'));
  plan = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('09:00'), cfg: CFG });
  assert.equal(E.fmt(plan.tasks[0].start), '09:30');
  plan = E.markTask(plan, 't', 'missed', E.hm('17:30'));
  plan = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('17:30'), cfg: CFG });
  assert.equal(plan.tasks[0].placement, 'tomorrow');
});

test('done task frees the rest of the day and stays logged', () => {
  let plan = { date: '2026-10-11', meetings: [], tasks: [task('a', { urgent: true }), task('b', {})] };
  plan = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  plan = E.markTask(plan, 'a', 'done', E.hm('08:45'));
  plan = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('08:45'), cfg: CFG });
  const b = plan.tasks.find(t => t.id === 'b');
  assert.equal(E.fmt(b.start), '08:45');
  assert.equal(plan.tasks.find(t => t.id === 'a').status, 'done');
});

test('carryOver brings unfinished work to the next morning', () => {
  const carried = E.carryOver({ tasks: [task('a', { status: 'done' }), task('b', { status: 'pending', start: 500, end: 530 })] });
  assert.deepEqual(carried.map(t => [t.id, t.carried, t.start]), [['b', true, null]]);
});

test('sanitizeClassification clamps bad AI output', () => {
  const c = E.sanitizeClassification({ minutes: 9999, context: 'ירח', urgent: 'yes' }, { title: 'x', snippet: '', from: '' });
  assert.equal(c.minutes, 240);
  assert.equal(c.context, 'מחשב');
  assert.equal(c.urgent, true);
});

test('nextWorkday skips Friday and Saturday', () => {
  assert.equal(E.nextWorkday('2026-10-11'), '2026-10-12'); // ראשון → שני
  assert.equal(E.nextWorkday('2026-10-15'), '2026-10-18'); // חמישי → ראשון
});

test('a near deadline raises priority and makes the task urgent', () => {
  const plan = { date: '2026-10-11', meetings: [], tasks: [
    task('far', { important: true, dueDate: '2026-10-30' }),
    task('soon', { important: true, dueDate: '2026-10-12', receivedAt: '2026-10-01T00:00:00Z' })
  ] };
  const out = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  const soon = out.tasks.find(t => t.id === 'soon');
  assert.equal(soon.quadrant, 'Q1');
  assert.equal(E.fmt(soon.start), '08:30');
});

test('a task due today with no free time is flagged at risk', () => {
  const plan = { date: '2026-10-11', meetings: [{ kind: 'meeting', start: E.hm('08:30'), end: E.hm('18:00') }],
    tasks: [task('x', { dueDate: '2026-10-11' })] };
  const out = E.replan(plan, [], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  assert.equal(out.tasks[0].placement, 'tomorrow');
  assert.equal(out.tasks[0].atRisk, true);
});

test('a rule about personal tasks matches the personal source', () => {
  const rule = Object.assign({ id: 'rp' }, E.parseRuleText('משימות אישיות רק בין 17:00 ל-18:00'));
  assert.deepEqual(rule.match.sources, ['personal']);
  const plan = { date: '2026-10-11', meetings: [], tasks: [task('p', { source: 'personal' })] };
  const out = E.replan(plan, [rule], { nowMs: NOW, fromMin: E.hm('08:00'), cfg: CFG });
  assert.equal(E.fmt(out.tasks[0].start), '17:00');
});
