// رفيق — application UI (mobile-first, RTL/LTR, light by default)
import { state, onChange, boot, enterLocal, signInGoogle, signOut, cloudReady, putTask, putPerson, putFile, putUnit, newTask, newPerson, newFile, newUnit, saveProfile, syncNow, uuid, nowISO, bulkReplace, localLeftovers, adoptLocalLeftovers, authError, connectDrive, unlinkDrive, refreshDriveStatus, invokeRefine, invokeAssist, sb, pushSupported, pushCurrent, pushEnable, pushDisable, pushTest } from './data.js';
import { reportHTML, ensureReportCSS, monthName } from './report.js';
import { uploadFile, trashFile, renameFile, fileBlob, viewUrl, previewUrl, rootFolderUrl, kindOf, extLabel, fmtSize, MAX_BYTES } from './drive.js';
import { monthOnly, parseCapture, timeFrom, splitTasks, isLogText, splitLogs, logDate, recurNext, cleanRecur, setWeekend, weekendDays, weekStartDay, today, addDays, diffDays, pd, ds, nextWeekday, addWorkdays, isWeekend, validDate, findPeople, nameTokens, normAr, lastOfMonth } from './parse.js';
import { DICT } from './i18n.js';
import { xlsxBlob } from './xlsx.js';
import { readSheet, serialDate } from './xlsxread.js';

/* ================= i18n & formatting ================= */
const L = () => state.profile.lang === 'en' ? 'en' : 'ar';
export function t(k, v) { let s = DICT[L()][k] ?? DICT.ar[k] ?? k; if (v) for (const x in v) s = s.split('{' + x + '}').join(v[x]); return s; }
const LOC = () => L() === 'ar' ? 'ar-u-ca-gregory-nu-latn' : 'en-GB';
const DTF = new Map();
const dtf = (loc, o) => { const k = loc + JSON.stringify(o); let f = DTF.get(k); if (!f) { f = new Intl.DateTimeFormat(loc, o); DTF.set(k, f); } return f; };
const fmt = (s, o = { day: 'numeric', month: 'long' }) => dtf(LOC(), o).format(pd(s));
const fmtShort = s => fmt(s, { day: 'numeric', month: 'short' });
const wd = s => dtf(LOC(), { weekday: 'long' }).format(pd(s));
const T = () => today();
function rel(s) { if (!s) return t('noDue'); const d = diffDays(s, T()); if (d === 0) return t('today'); if (d === 1) return t('tomorrow'); if (d === -1) return t('yesterday'); if (d > 1 && d < 7) return wd(s); return fmtShort(s); }
function plural(n, base) { // Arabic-aware counts
  if (L() === 'en') return t(base + (n === 1 ? '_1' : '_n'), { n });
  if (n === 1) return t(base + '_1'); if (n === 2) return t(base + '_2'); if (n >= 3 && n <= 10) return t(base + '_few', { n }); return t(base + '_n', { n });
}
const lateTxt = s => t('lateBy', { d: plural(diffDays(T(), s), 'days') });
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const $ = s => document.querySelector(s);

/* ================= icons ================= */
const I = (p, a = '') => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" ${a}>${p}</svg>`;
const ic = {
  today: I('<path d="M3 10h18M8 3v4M16 3v4"/><rect x="3" y="5" width="18" height="16" rx="3"/><path d="M8 14h3v3H8z"/>'),
  list: I('<path d="M9 6h11M9 12h11M9 18h11"/><path d="M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2"/>'),
  plus: I('<path d="M12 5v14M5 12h14"/>', 'stroke-width="2.4"'),
  users: I('<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.8-3.6 3.4-5.5 6.5-5.5s5.7 1.9 6.5 5.5M16 4.5a3.5 3.5 0 0 1 0 7M18.5 14.8c1.6.8 2.6 2.5 3 5.2"/>'),
  user: I('<circle cx="12" cy="8" r="4"/><path d="M4 21c1-4 4.5-6 8-6s7 2 8 6"/>'),
  more: I('<circle cx="6" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="18" cy="12" r="1.6"/>'),
  phone: I('<path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2"/>'),
  mic: I('<rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0 0 14 0M12 18v3"/>'),
  check: I('<path d="M5 12.5l4.5 4.5L19 7.5"/>', 'stroke-width="3"'),
  clock: I('<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>'),
  cal: I('<rect x="3" y="5" width="18" height="16" rx="3"/><path d="M3 10h18M8 3v4M16 3v4"/>'),
  back: I('<path d="M9 6l6 6-6 6"/>'),
  x: I('<path d="M6 6l12 12M18 6L6 18"/>'),
  search: I('<circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/>'),
  flag: I('<path d="M5 21V4M5 4h12l-2 4 2 4H5"/>'),
  later: I('<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l3 2"/>'),
  edit: I('<path d="M4 20h4L19 9l-4-4L4 16z"/><path d="M14 6l4 4"/>'),
  trash: I('<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>'),
  archive: I('<rect x="3" y="4" width="18" height="5" rx="1.5"/><path d="M5 9v10h14V9M10 13h4"/>'),
  reopen: I('<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/>'),
  gear: I('<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1"/>'),
  globe: I('<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3c2.5 2.6 3.8 5.6 3.8 9s-1.3 6.4-3.8 9c-2.5-2.6-3.8-5.6-3.8-9S9.5 5.6 12 3z"/>'),
  moon: I('<path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5z"/>'),
  cloud: I('<path d="M7 18h10a4 4 0 0 0 .5-8 6 6 0 0 0-11.5 1.5A3.5 3.5 0 0 0 7 18z"/>'),
  download: I('<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>'),
  upload: I('<path d="M12 20V9M7 14l5-5 5 5M5 4h14"/>'),
  logout: I('<path d="M15 4h4v16h-4M10 8l-4 4 4 4M6 12h10"/>'),
  info: I('<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.5"/>'),
  minus: I('<path d="M5 12h14"/>', 'stroke-width="2.4"'),
  note: I('<path d="M5 4h14v16H5z"/><path d="M8 9h8M8 13h8M8 17h5"/>'),
  spark: I('<path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/>'),
  org: I('<rect x="9" y="3" width="6" height="5" rx="1.2"/><rect x="3" y="16" width="6" height="5" rx="1.2"/><rect x="15" y="16" width="6" height="5" rx="1.2"/><path d="M12 8v4M6 16v-2h12v2"/>'),
  folder: I('<path d="M3 7a2 2 0 0 1 2-2h4l2 2.5h8a2 2 0 0 1 2 2V18a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>'),
  table: I('<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/>'),
  clip: I('<path d="M20 11.5l-8.2 8.2a5 5 0 0 1-7.1-7.1l8.5-8.5a3.3 3.3 0 0 1 4.7 4.7l-8.5 8.5a1.7 1.7 0 0 1-2.4-2.4l7.8-7.8"/>'),
  ext: I('<path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>'),
  link: I('<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>'),
  retry: I('<path d="M20 12a8 8 0 1 1-2.3-5.7L20 8.5M20 3v5.5h-5.5"/>'),
  send: I('<path d="M21 3L10 14M21 3l-7 18-4-7-7-4z"/>'),
  repeat: I('<path d="M17 2l4 4-4 4"/><path d="M3 11v-1a4 4 0 0 1 4-4h14"/><path d="M7 22l-4-4 4-4"/><path d="M21 13v1a4 4 0 0 1-4 4H3"/>'),
  bell: I('<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>'),
};
const sic = (k, s = 14) => ic[k].replace('<svg', `<svg width="${s}" height="${s}"`);
const gLogo = '<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><path fill="#4285F4" d="M22.5 12.3c0-.8-.1-1.5-.2-2.2H12v4.2h5.9a5 5 0 0 1-2.2 3.3v2.7h3.6c2.1-1.9 3.2-4.8 3.2-8z"/><path fill="#34A853" d="M12 23c3 0 5.5-1 7.3-2.7l-3.6-2.7c-1 .7-2.2 1-3.7 1-2.9 0-5.3-1.9-6.2-4.5H2.1v2.8A11 11 0 0 0 12 23z"/><path fill="#FBBC05" d="M5.8 14.1a6.6 6.6 0 0 1 0-4.2V7.1H2.1a11 11 0 0 0 0 9.8z"/><path fill="#EA4335" d="M12 5.4c1.6 0 3.1.6 4.2 1.7l3.2-3.2A11 11 0 0 0 2.1 7.1l3.7 2.8C6.7 7.3 9.1 5.4 12 5.4z"/></svg>';

/* ================= model helpers ================= */
const live = () => state.tasks.filter(x => !x.deleted);
const COLL = new Intl.Collator('ar');
let peopleMemo = { rev: -1, list: [] };
const people = () => { if (peopleMemo.rev !== state.rev) peopleMemo = { rev: state.rev, list: state.people.filter(p => !p.deleted).sort((a, b) => COLL.compare(a.name, b.name)) }; return peopleMemo.list; };
const taskById = id => state.tasks.find(x => x.id === id);
const personById = id => state.people.find(p => p.id === id);
const pname = id => { const p = personById(id); return p && !p.deleted ? p.name : t('deletedPerson'); };
const files = () => state.files.filter(f => !f.deleted).sort((a, b) => (b.created_at || '') > (a.created_at || '') ? 1 : -1);
const fileById = id => state.files.find(f => f.id === id);
const filesOfTask = id => files().filter(f => f.task_id === id);
const filesOfPerson = id => files().filter(f => f.person_id === id);
/* departments (kind 'dept') and their sections (kind 'sec', parent_id = the department) */
const orgOn = () => !(state.mode === 'cloud' && state.schemaOld); // the cloud database must be updated (009_org.sql) first
let unitsMemo = { rev: -1, list: [] };
const units = () => { if (unitsMemo.rev !== state.rev) unitsMemo = { rev: state.rev, list: (state.units || []).filter(u => !u.deleted).sort((a, b) => (a.sort || 0) - (b.sort || 0) || COLL.compare(a.name, b.name)) }; return unitsMemo.list; };
const depts = () => units().filter(u => u.kind === 'dept');
const secsOf = id => units().filter(u => u.kind === 'sec' && u.parent_id === id);
const unitById = id => id ? (state.units || []).find(u => u.id === id && !u.deleted) || null : null;
const deptOf = id => { const u = unitById(id); return !u ? null : u.kind === 'dept' ? u : unitById(u.parent_id); };
const unitLabel = (id, short) => { const u = unitById(id); if (!u) return ''; if (u.kind === 'dept') return u.name; const d = unitById(u.parent_id); return d && !short ? `${d.name} › ${u.name}` : u.name; };
const unitOf = x => orgOn() ? unitById(x?.unit_id) : null;
/** A department/section named in free text: exact names first, then a close one («التميز» → «إدارة التميز المؤسسي»). Returns its id or null. */
const UNIT_STOP = /^(اداره|ادارة|قسم|قطاع|مكتب|وحده|وحدة|ال)$/;
function matchUnit(text, strict) {
  if (/[›>]/.test(String(text || ''))) { const [a, b] = String(text).split(/[›>]/).map(x => x.trim()); const d = deptOf(matchUnit(a)); if (!d) return null; const sc = secsOf(d.id).find(x => normAr(x.name) === normAr(b)) || secsOf(d.id).find(x => normAr(x.name).includes(normAr(b)) || normAr(b).includes(normAr(x.name))); return sc ? sc.id : d.id; } // «الإدارة › القسم» as the app writes it
  const n = ' ' + normAr(String(text || '')).replace(/[«»"'().,،؛:?؟!]/g, ' ').replace(/\s+/g, ' ').trim() + ' '; if (n.trim().length < 2) return null;
  const list = units();
  const rx = s => new RegExp('\\s(?:[وفبل]|ال|لل)?' + s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s');
  const core = nm => normAr(nm).replace(/^(اداره|قسم|قطاع|وحده|مكتب)\s+/, '');
  const whole = list.filter(u => rx(normAr(u.name)).test(n) || (core(u.name) !== normAr(u.name) && core(u.name).split(' ').length >= 1 && new RegExp('(اداره|قسم|قطاع|وحده|مكتب)\\s+(?:ال)?' + core(u.name).replace(/^ال/, '(?:ال)?') + '\\s').test(n))).sort((a, b) => b.name.length - a.name.length)[0];
  if (whole) return whole.id;
  if (strict) return null; // inside a task's own words only a named department counts («تقرير الأداء» is not «قسم قياس الأداء»)
  const sig = s => normAr(s).split(/\s+/).map(w => w.replace(/^ال/, '')).filter(w => w.length > 2 && !UNIT_STOP.test(w));
  const words = new Set(sig(n)); let best = null, bs = 0;
  list.forEach(u => { const ws = sig(u.name); if (!ws.length) return; const hit = ws.filter(w => words.has(w)).length; const sc = hit / ws.length + (u.kind === 'sec' ? 0.01 : 0); if (hit && sc > bs) { bs = sc; best = u.id; } });
  return bs >= 0.5 ? best : null;
}
/** The unit a person belongs to (set once on the person). */
const unitOfPerson = pid => { const p = personById(pid); return p && unitById(p.unit_id) ? p.unit_id : null; };
const localDay = iso => ds(new Date(iso || Date.now()));
const initials = n => { const s = String(n || '').replace(/^(?:م|د|أ|ا|Eng|Dr|Mr|Mrs|Ms)\.?\s+/i, '').replace(/^ال(?=\S{2,})/, '').trim(); const m = s.match(/[\p{L}\p{N}]/u); return m ? m[0].toUpperCase() : '?'; };
const isOpen = x => ['todo', 'prog', 'wait', 'review', 'hold'].includes(x.status);
const isWaitLike = x => x.status === 'wait' || x.status === 'review'; // waiting on someone else: a reply, or a review/approval of finished work
const isLog = x => x.kind === 'log';
const isLate = x => isOpen(x) && x.status !== 'review' && x.due && x.due < T(); // under review = the user's part is done
const openFus = x => (x.followups || []).filter(f => f.status === 'open');
const fuLate = f => f.status === 'open' && f.due && f.due < T();
function addLog(x, kind, text = '', data = {}) { x.log = x.log || []; x.log.push({ id: uuid(), at: nowISO(), kind, text, data }); }
function getOrCreatePerson(name) {
  const n = (name || '').trim().replace(/\s+/g, ' ').slice(0, 200); if (!n) return null;
  const id = matchPerson(n); if (id) return id;
  const p = newPerson(n); putPerson(p); return p.id;
}
function sortTasks(list) {
  return list.slice().sort((a, b) => (isLate(b) - isLate(a)) || ((a.due || '9999') < (b.due || '9999') ? -1 : (a.due || '9999') > (b.due || '9999') ? 1 : 0) || ((a.due_time || '99') < (b.due_time || '99') ? -1 : (a.due_time || '99') > (b.due_time || '99') ? 1 : 0) || ({ hi: 0, mid: 1, lo: 2 }[a.priority] - { hi: 0, mid: 1, lo: 2 }[b.priority]) || (a.created_at < b.created_at ? -1 : 1));
}
const nextWorkday = () => addWorkdays(T(), 1);
const wdName = d => dtf(LOC(), { weekday: 'long' }).format(new Date(2026, 9, 4 + d)); // 4 Oct 2026 is a Sunday
const fmtTime = s => { const m = /^(\d{2}):(\d{2})$/.exec(s || ''); return m ? dtf(LOC(), { hour: 'numeric', minute: '2-digit' }).format(new Date(2026, 0, 1, +m[1], +m[2])) : ''; };
function recurLabel(r) {
  if (!r) return '';
  if (r.f === 'daily') return t('rcDaily');
  if (r.f === 'weekly') { const days = r.days || []; if (!days.length) return t('rcWeeklyOpt', { d: '' }); return t('rcWeekly', { d: days.map(d => L() === 'ar' ? wdName(d).replace(/^ال/, '') : wdName(d)).join(L() === 'ar' ? ' و' : ', ') }); }
  if (r.f === 'monthly') return r.dom === 'last' ? t('rcMonthLast') : t('rcMonthly', { n: r.dom });
  if (r.f === 'quarterly') return r.dom === 'last' ? t('rcQuarterLast') : t('rcQuarterly', { n: r.dom });
  return '';
}
const remindLabel = m => m == null ? t('rmDefault') : m < 0 ? t('rmNone') : m === 0 ? t('rmAt') : m === 60 ? t('rmHour') : m === 120 ? t('rmHours2') : t('rmBefore', { m });
/** Recurring task: after one occurrence is completed, create the next one. */
function spawnNext(x) {
  if (!x.recur || isLog(x)) return null;
  let nd = recurNext(x.recur, x.due || x.completed_on || T()), guard = 0;
  while (nd && nd <= T() && guard++ < 400) nd = recurNext(x.recur, nd);
  if (!nd) return null;
  const y = newTask({ title: x.title, details: x.details, priority: x.priority, role: x.role, source: x.source, project: x.project, due: nd, due_time: x.due_time || null, remind_min: x.remind_min ?? null, recur: x.recur, steps_total: x.steps_total || 0 });
  const seen = new Set();
  y.followups = (x.followups || []).filter(f => !seen.has(f.person_id) && seen.add(f.person_id)).map(f => ({ id: uuid(), person_id: f.person_id, what: f.what || '', due: diffDays(nd, T()) > 1 ? addWorkdays(nd, -1) : nd, time: f.time || null, status: 'open', created_at: nowISO(), closed_on: null, log: [] }));
  addLog(y, 'recurred', '', { from: x.id });
  x.recur = null; addLog(x, 'spawned', '', { next: y.id, due: nd });
  putTask(y);
  return y;
}

/* ================= UI state ================= */
const UI = { tab: 'today', seg: 'active', personF: null, q: '', layers: [], theme: null, booting: true, lq: '', ltype: 'all', lmonth: '' };

/* ================= toast ================= */
let toastTimer;
function toast(msg, undo, kind, ms) {
  const el = $('#toast');
  el.className = 'toast' + (kind ? ' ' + kind : '');
  el.innerHTML = `<span>${esc(msg)}</span>${undo ? `<button type="button" id="undoBtn">${t('undo')}</button>` : ''}`;
  requestAnimationFrame(() => el.classList.add('on'));
  clearTimeout(toastTimer);
  if (undo) $('#undoBtn').onclick = () => { undo(); el.classList.remove('on'); };
  toastTimer = setTimeout(() => el.classList.remove('on'), ms || (undo ? 4500 : 2600));
}

/* ================= shell ================= */
function applyLangTheme() {
  const shell = $('#shell');
  const lang = L();
  document.documentElement.lang = lang; document.documentElement.dir = lang === 'ar' ? 'rtl' : 'ltr';
  shell.dir = lang === 'ar' ? 'rtl' : 'ltr';
  const dark = (state.profile.settings || {}).theme === 'dark';
  if (dark) shell.setAttribute('data-mode', 'dark'); else shell.removeAttribute('data-mode');
  document.querySelector('meta[name=theme-color]')?.setAttribute('content', dark ? '#0F1522' : '#EEF1F5');
  document.title = lang === 'ar' ? 'رفيق' : 'Rafeeq';
  document.querySelector('nav[aria-label]')?.setAttribute('aria-label', t('navLabel'));
}
function renderAll() {
  applyLangTheme();
  const signedIn = !!state.mode;
  $('#splash').hidden = !(UI.booting && !signedIn);
  if (UI.booting && !signedIn) { $('#login').hidden = true; $('#app').hidden = true; return; }
  $('#login').hidden = signedIn;
  $('#app').hidden = !signedIn;
  if (!signedIn) { renderLogin(); return; }
  if (!state.profile.display_name && !UI.onboarded) { renderOnboarding(); return; }
  renderNav(); renderMain(); UI.layers.forEach(Lr => Lr.rerender());
  if (location.hash.startsWith('#t=')) routeHash();
}

function renderLogin() {
  $('#login').innerHTML = `
    <div class="login-in">
      <div class="brand-mark">${L() === 'ar' ? 'ر' : 'R'}</div>
      <h1>${t('appName')}</h1>
      <p class="lead">${t('tagline')}</p>
      <ul class="feat">
        <li><span class="fi">${ic.mic}</span><span><b>${t('feat1t')}</b>${t('feat1d')}</span></li>
        <li><span class="fi">${ic.users}</span><span><b>${t('feat2t')}</b>${t('feat2d')}</span></li>
        <li><span class="fi">${ic.cloud}</span><span><b>${t('feat3t')}</b>${t('feat3d')}</span></li>
      </ul>
      ${authError ? `<div class="notice err" role="alert"><b>${t('authErrTitle')}</b><span>${t(authErrKey())}</span></div>` : ''}
      <div class="login-actions">
        ${cloudReady ? `<button type="button" class="gbtn" data-act="google">${gLogo}${t('continueGoogle')}</button>` : `<div class="notice">${t('cloudNotReady')}</div>`}
        <button type="button" class="btn ghost block" data-act="local">${t('useLocal')}</button>
        <button type="button" class="linkbtn" data-act="lang">${L() === 'ar' ? 'English' : 'العربية'}</button>
      </div>
      <p class="legal">${t('loginLegal')} <a href="privacy.html">${t('privacy')}</a></p>
    </div>`;
}
function authErrKey() {
  const d = (authError?.detail || '') + ' ' + (authError?.code || '');
  if (/access_denied|cancel/i.test(d)) return 'authErrDenied';
  if (/exchange external code|invalid_client|redirect_uri|unauthorized_client/i.test(d)) return 'authErrSetup';
  return 'authErrGeneric';
}
function renderOnboarding() {
  $('#login').hidden = true; $('#app').hidden = false;
  $('#nav').hidden = true; $('#app').classList.add('bare'); closeAllLayers();
  $('#main').innerHTML = `<div class="onb">
      <div class="brand-mark">${L() === 'ar' ? 'ر' : 'R'}</div>
      <h1>${t('onbTitle')}</h1><p class="lead">${t('onbLead')}</p>
      <form id="onbForm" class="form" autocomplete="off">
        <label class="fld"><span>${t('yourName')}</span><input id="onbName" name="name" required maxlength="80" value="${esc(state.user?.name || '')}" dir="auto"></label>
        <label class="fld"><span>${t('jobTitle')} <em>${t('optional')}</em></span><input id="onbJob" name="job" maxlength="80" placeholder="${esc(t('jobTitlePh'))}" dir="auto"></label>
        <div class="fld"><span>${t('language')}</span><div class="seg2"><button type="button" class="${L() === 'ar' ? 'on' : ''}" data-act="setlang" data-v="ar">العربية</button><button type="button" class="${L() === 'en' ? 'on' : ''}" data-act="setlang" data-v="en">English</button></div></div>
        <button class="btn primary block lg" type="submit">${t('start')}</button>
      </form></div>`;
}
function renderNav() {
  $('#nav').hidden = false; $('#app').classList.remove('bare');
  const tab = (k, icon, label) => `<button type="button" class="tab ${UI.tab === k ? 'on' : ''}" data-tab="${k}" ${UI.tab === k ? 'aria-current="page"' : ''}>${icon}<span>${label}</span>${k === 'today' && badgeCount() ? `<i class="dot num">${badgeCount()}</i>` : ''}</button>`;
  $('#nav').innerHTML = `<div class="nav-brand"><span class="brand-mark sm">${L() === 'ar' ? 'ر' : 'R'}</span><b>${t('appName')}</b></div>
    ${tab('today', ic.today, t('navToday'))}${tab('tasks', ic.list, t('navTasks'))}
    <button type="button" class="addbtn" data-act="chat" aria-label="${esc(t('assistant'))}">${ic.spark}<span class="lbl">${t('askRafeeq')}</span></button>
    ${tab('people', ic.users, t('navPeople'))}${tab('library', ic.folder, t('navLibrary'))}${tab('more', ic.more, t('navMore')).replace('class="tab', 'class="tab deskonly')}
    <div class="nav-foot">${syncChip()}</div>`;
}
const meBtn = () => `<button type="button" class="me mob" data-tab="more" aria-label="${esc(t('navMore'))}"><span class="av">${esc(initials(state.profile.display_name || '?'))}</span><span class="gb" aria-hidden="true">${sic('more', 12)}</span></button>`;
const badgeCount = () => { const act = live().filter(x => isOpen(x) && !x.archived); return act.filter(isLate).length + act.reduce((n, x) => n + openFus(x).filter(fuLate).length, 0); };
function syncChip() {
  const s = state.sync;
  if (state.mode === 'local') return `<button type="button" class="sync local" data-tab="more" title="${esc(t('syncLocal'))}" aria-label="${esc(t('syncLocal'))}">${sic('info')}<span class="sl">${t('syncLocal')}</span></button>`;
  if (s.status === 'offline') return `<span class="sync warn">${sic('cloud')}${t('syncOffline')}${s.pending ? ` · <b class="num">${s.pending}</b>` : ''}</span>`;
  if (s.status === 'error' || (s.rejected || []).length) return `<button type="button" class="sync err" data-tab="more">${sic('cloud')}${t('syncError')}</button>`;
  if (s.status === 'syncing' || s.pending) return `<span class="sync" title="${esc(t('syncing'))}">${sic('cloud')}<span class="sl">${t('syncing')}</span></span>`;
  return `<span class="sync ok" title="${esc(t('synced'))}">${sic('check')}<span class="sl">${t('synced')}</span></span>`;
}
/** Re-render only the results of a search box, so the box (and the phone keyboard) stays put. */
function renderListOnly(view, sel) {
  const box = $(sel); if (!box || UI.tab !== view) return renderMain();
  UI.listOnly = true; try { box.innerHTML = VIEWS[view](); } finally { UI.listOnly = false; }
  bindSwipes(box);
}
function renderMain() {
  setWeekend((state.profile.settings || {}).wk || [5, 6]);
  const m = $('#main');
  const y = window.scrollY;
  m.innerHTML = VIEWS[UI.tab]();
  m.classList.toggle('wide', UI.tab === 'tasks' && UI.tview === 'table');
  bindSwipes(m);
  if (UI.keepScroll) window.scrollTo(0, y);
  UI.keepScroll = false;
}
function go(tab) { UI.tab = tab; closeAllLayers(); renderNav(); renderMain(); window.scrollTo(0, 0); }

/* ================= components ================= */
function statusPill(x) {
  if (isLog(x)) return `<span class="pill log">${sic('check', 12)}${t('logPill')}</span>`;
  const m = { todo: ['', 'stTodo'], prog: ['prog', 'stProg'], wait: ['wait', 'stWait'], review: ['review', 'stReview'], hold: ['', 'stHold'], done: ['done', 'stDone'], cancelled: ['', 'stCancelled'] }[x.status] || ['', 'stTodo'];
  return `<span class="pill ${m[0]}">${x.status === 'wait' ? sic('clock', 13) : ''}${t(m[1])}</span>`;
}
/** The status on a card, tappable: one tap opens the status choices. */
function statusBtn(x) {
  const m = { prog: 'prog', wait: 'wait', review: 'review' }[x.status] || '';
  return `<button type="button" class="pill stbtn ${m}" data-act="stmenu" data-id="${x.id}" aria-label="${esc(t('changeStatusOf', { s: t(stKey(x.status)) }))}">${x.status === 'wait' ? sic('clock', 12) : ''}${t(stKey(x.status))}<svg viewBox="0 0 24 24" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.6" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg></button>`;
}
function openStatusMenu(x) {
  openSheet(t('changeStatus'), `<p class="note" dir="auto" style="margin-top:0">«${esc(x.title.slice(0, 80))}»</p>
    <div class="st-pick">${OPEN_ST().map(s => `<button type="button" class="st-opt ${x.status === s ? 'on' : ''}" data-act="stpick" data-id="${x.id}" data-v="${s}"><span class="pill ${{ prog: 'prog', wait: 'wait', review: 'review' }[s] || ''}">${t(stKey(s))}</span><small>${t('stHint_' + s)}</small></button>`).join('')}
      <button type="button" class="st-opt" data-act="stpick" data-id="${x.id}" data-v="done"><span class="pill done">${sic('check', 12)}${t('stDone')}</span><small>${t('stHint_done')}</small></button></div>`);
}
function duePill(x) {
  if (!x.due) return isOpen(x) ? '' : '';
  if (isLate(x)) return `<span class="pill late">${lateTxt(x.due)}</span>`;
  if (x.due === T() && isOpen(x)) return `<span class="pill wait">${t('dueToday')}</span>`;
  return `<span class="num">${sic('cal', 13)} ${rel(x.due)}</span>`;
}
function taskCard(x, opts = {}) {
  const nf = openFus(x).sort((a, b) => (a.due || '9') < (b.due || '9') ? -1 : 1)[0];
  const stripe = isLate(x) ? 'late' : isWaitLike(x) ? 'wait' : '';
  const done = !isOpen(x);
  // one line, two facts at most: when, and the single thing that matters most about it (details live in the task page)
  const when = done ? (x.completed_on ? `<span>${t('doneOn', { d: fmtShort(x.completed_on) })}</span>` : '')
    : (x.due || x.due_time ? (isLate(x) ? `<span class="pill late">${lateTxt(x.due)}</span>` : `<span class="num">${x.due ? rel(x.due) : ''}${x.due_time ? `${x.due ? ' · ' : ''}${fmtTime(x.due_time)}` : ''}${x.recur ? ` ${sic('repeat', 12)}` : ''}</span>`) : '');
  const key = x.status === 'wait' && x.waiting_on ? `<span>${sic('clock', 13)} ${t('waitingFrom', { w: t('w_' + (x.waiting_what || 'reply')), p: esc(pname(x.waiting_on)) })}</span>`
    : x.priority === 'hi' && !done ? `<span class="pill hi">${t('prioHi')}</span>`
    : nf ? `<span>${sic('users', 13)} ${esc(pname(nf.person_id).split(' ').slice(0, 2).join(' '))}${openFus(x).length > 1 ? ` +${openFus(x).length - 1}` : ''}</span>`
    : x.project && !opts.noProject ? `<span class="projlink" data-proj="${esc(x.project)}" role="link" dir="auto">${esc(x.project)}</span>` : '';
  const un = unitOf(x);
  const stEl = isLog(x) ? statusPill(x) : isOpen(x) && !opts.noStatusBtn ? statusBtn(x) : opts.status ? statusPill(x) : '';
  const meta = [stEl, when, key, un && !opts.noUnit ? `<span class="unit" dir="auto">${sic('org', 12)} ${esc(unitLabel(un.id, true))}</span>` : ''].filter(Boolean).join('');
  const card = `<div class="tcard ${done ? 'is-done' : ''} ${isLog(x) ? 'is-log' : ''}" data-open="${x.id}" role="button" tabindex="0">
      ${stripe ? `<span class="stripe ${stripe}"></span>` : ''}
      ${x.status === 'cancelled' ? '<span class="chk" aria-hidden="true"></span>' : `<button type="button" class="chk ${x.status === 'done' ? 'on' : ''}" ${isOpen(x) ? `data-act="quickdone" data-id="${x.id}" aria-label="${esc(t('complete'))}"` : 'tabindex="-1" aria-hidden="true"'}>${x.status === 'done' ? ic.check : ''}</button>`}
      <div class="body"><div class="t" dir="auto">${esc(x.title)}</div>${meta ? `<div class="meta">${meta}</div>` : ''}</div></div>`;
  if (!isOpen(x) || opts.noSwipe) return card;
  return `<div class="swipe" data-id="${x.id}"><span class="lbl done">${sic('check', 18)}${t('swDone')}</span><span class="lbl later">${sic('later', 18)}${t('swLater')}</span>${card}</div>`;
}
function fuRow(x, f) {
  return `<div class="fu-row ${fuLate(f) ? 'late' : ''}"><span class="av">${esc(initials(pname(f.person_id)))}</span>
    <div class="body" data-open="${x.id}" role="button" tabindex="0"><div class="t" dir="auto">${esc(pname(f.person_id))}</div>
      <div class="meta"><span dir="auto">${esc(f.what || x.title)}</span>${f.time ? `<span class="num">${sic('clock', 12)} ${fmtTime(f.time)}</span>` : ''}${fuLate(f) ? `<span class="pill late">${lateTxt(f.due)}</span>` : ''}</div></div>
    <button type="button" class="iconbtn sm" data-act="remind" data-id="${x.id}" data-fu="${esc(f.id)}" aria-label="${esc(t('remind'))}" title="${esc(t('remind'))}">${sic('send', 17)}</button><button type="button" class="btn sm" data-act="furesult" data-id="${x.id}" data-fu="${esc(f.id)}">${t('logResult')}</button></div>`;
}
/* «متابعاتي»: everyone with something open between you, most urgent first */
function myFollowups() {
  const td = T(), by = new Map();
  const add = (pid, it) => { if (!pid) return; if (!by.has(pid)) by.set(pid, []); by.get(pid).push(it); };
  live().filter(x => isOpen(x) && !x.archived).forEach(x => {
    openFus(x).forEach(f => add(f.person_id, { x, f, due: f.due || null, late: fuLate(f) }));
    if (x.status === 'wait' && x.waiting_on) add(x.waiting_on, { x, f: null, due: x.due || null, late: isLate(x), wait: true });
  });
  const lateDays = it => it.late && it.due ? diffDays(td, it.due) : 0;
  return [...by].map(([pid, items]) => {
    items.sort((a, b) => (b.late - a.late) || ((a.due || '9999') < (b.due || '9999') ? -1 : 1));
    const late = items.filter(i => i.late).length, worst = Math.max(0, ...items.map(lateDays));
    const next = items.map(i => i.due).filter(Boolean).sort()[0] || null;
    return { pid, items, late, worst, next, today: items.some(i => i.due === td), waitOnly: items.every(i => i.wait) };
  }).sort((a, b) => (b.worst - a.worst) || (b.late - a.late) || (b.today - a.today) || ((a.next || '9999') < (b.next || '9999') ? -1 : (a.next || '9999') > (b.next || '9999') ? 1 : 0) || pname(a.pid).localeCompare(pname(b.pid), 'ar'));
}
function mfCard(g) {
  const p = personById(g.pid), name = pname(g.pid), phone = p?.contact && /\d{7,}/.test(p.contact) ? p.contact : '';
  const badge = g.late ? `<span class="pill late">${t('mfLate', { d: plural(g.worst, 'days') })}</span>` : g.today ? `<span class="pill prog">${t('today')}</span>` : g.waitOnly ? `<span class="pill wait">${t('mfWaitingReply')}</span>` : g.next ? `<span class="num muted">${rel(g.next)}</span>` : '';
  const row = it => `<div class="mf-item ${it.late ? 'late' : ''}">
      <span class="body" data-open="${it.x.id}" role="button" tabindex="0"><span class="t" dir="auto">${esc(it.wait ? it.x.title : (it.f.what || it.x.title))}</span>
        <span class="meta">${it.f && it.f.what ? `<span dir="auto">${esc(it.x.title)}</span>` : ''}${it.due ? `<span class="due">${it.late ? lateTxt(it.due) : rel(it.due)}</span>` : `<span class="due">${t('noDue')}</span>`}${it.wait ? `<span>${sic('clock', 12)} ${t('mfWaitingFor', { w: t('w_' + (it.x.waiting_what || 'reply')) })}</span>` : ''}</span></span>
      ${it.f ? `<button type="button" class="iconbtn sm" data-act="remind" data-id="${it.x.id}" data-fu="${esc(it.f.id)}" aria-label="${esc(t('remind'))}" title="${esc(t('remind'))}">${sic('send', 17)}</button><button type="button" class="btn sm" data-act="furesult" data-id="${it.x.id}" data-fu="${esc(it.f.id)}">${t('logResult')}</button>` : ''}
    </div>`;
  return `<section class="mf-card ${g.late ? 'late' : ''}">
      <div class="mf-h"><button type="button" class="mf-who" data-person="${g.pid}"><span class="av">${esc(initials(name))}</span><span class="grow"><b dir="auto">${esc(name)}</b>${p?.org ? `<small dir="auto">${esc(p.org)}</small>` : ''}</span></button>${badge}</div>
      <div class="mf-items">${g.items.slice(0, 6).map(row).join('')}${g.items.length > 6 ? `<button type="button" class="linkbtn" data-person="${g.pid}">${t('showMore', { n: g.items.length - 6 })}</button>` : ''}</div>
      ${phone ? `<div class="mf-acts"><a class="btn sm" href="tel:${esc(phone.replace(/[^\d+]/g, ''))}">${sic('phone', 15)}${t('call')}</a><a class="btn sm" href="https://wa.me/${esc(waNumber(phone))}" target="_blank" rel="noopener">${sic('send', 15)}WhatsApp</a></div>` : ''}
    </section>`;
}
/** Departments with open work, as filter chips with counts (late ones flagged). */
function unitRow() {
  if (!orgOn() || !depts().length) return '';
  const open = live().filter(x => isOpen(x) && !x.archived);
  const rows = depts().map(d => { const ids = new Set([d.id, ...secsOf(d.id).map(s => s.id)]); const ts = open.filter(x => ids.has(x.unit_id)); return { d, n: ts.length, late: ts.filter(isLate).length }; }).filter(r => r.n || UI.unitF === r.d.id);
  const none = open.filter(x => !unitById(x.unit_id)).length;
  if (!rows.length) return '';
  return `<div class="people-row unit-row">${rows.map(r => `<button type="button" class="pchip uchip ${UI.unitF && deptOf(UI.unitF)?.id === r.d.id ? 'on' : ''}" data-unitf="${r.d.id}">${sic('org', 13)}<span class="nm" dir="auto">${esc(r.d.name)}</span><i class="num">${r.n}</i>${r.late ? `<i class="num late">${r.late} ${t('secLate')}</i>` : ''}</button>`).join('')}${none ? `<button type="button" class="pchip uchip ${UI.unitF === '-' ? 'on' : ''}" data-unitf="-"><span class="nm">${t('noUnit')}</span><i class="num">${none}</i></button>` : ''}</div>`;
}
const inUnit = (x, uid) => uid === '-' ? !unitById(x.unit_id) : !!uid && (x.unit_id === uid || unitById(x.unit_id)?.parent_id === uid);
function tasksHead(mode) {
  const full = mode === 'list', tbl = mode === 'table';
  const pc = id => live().filter(x => isOpen(x) && (x.waiting_on === id || openFus(x).some(f => f.person_id === id))).length;
  const ppl = people().filter(p => pc(p.id)).slice(0, 12);
  const openN = live().filter(isOpen).length;
  return `<header class="hdr"><div class="grow"><h1>${t('navTasks')}</h1><div class="sub">${plural(openN, 'openTasks')}</div></div><button type="button" class="btn sm" data-tview="${tbl ? 'cards' : 'table'}" aria-pressed="${tbl}">${tbl ? sic('list', 16) : sic('table', 16)}<span class="lbl-sm">${t(tbl ? 'viewCards' : 'viewTable')}</span></button><button type="button" class="btn primary sm desk" data-act="chat">${ic.plus}${t('newTask')}</button><button type="button" class="iconbtn addman mob" data-act="chat" aria-label="${esc(t('newTask'))}">${ic.plus}</button>${meBtn()}</header>
      ${full ? `<label class="search">${ic.search}<input type="search" id="taskQ" value="${esc(UI.q)}" placeholder="${esc(t('searchTasks'))}" aria-label="${esc(t('searchTasks'))}" dir="auto"></label>` : ''}
      ${full ? unitRow() : ''}
      ${full && ppl.length ? `<div class="people-row">${ppl.map(p => `<button type="button" class="pchip ${UI.personF === p.id ? 'on' : ''}" data-pf="${p.id}"><span class="av">${esc(initials(p.name))}</span><span class="nm" dir="auto">${esc(p.name.replace(/^(?:م|د|أ|ا)\.\s*/, '').split(' ')[0])}</span><i class="num">${pc(p.id)}</i></button>`).join('')}</div>` : ''}
      ${tbl ? '' : `<div class="seg four" role="tablist">${[['active', 'segActive'], ['wait', 'segWaitShort'], ['week', 'segWeek'], ['done', 'segDone']].map(([k, l]) => `<button type="button" role="tab" aria-selected="${UI.seg === k}" class="${UI.seg === k ? 'on' : ''}" data-seg="${k}">${t(l)}</button>`).join('')}</div>`}`;
}
/* table: every task in one sheet, filterable and sortable — for review with a manager, and as an Excel file */
const PRIO_L = p => t(p === 'hi' ? 'tpHi' : p === 'lo' ? 'tpLo' : 'tpMid');
const peopleOf = x => [...new Set([x.status === 'wait' ? x.waiting_on : null, ...(isOpen(x) ? openFus(x) : (x.followups || [])).map(f => f.person_id)].filter(Boolean).map(pname))];
const lastTouch = x => { const at = (x.log || []).map(l => l.at).filter(Boolean).sort().pop() || x.client_ts || x.created_at; return at ? localDay(at) : ''; };
const TCOLS = [
  { k: 'title', h: 'fTitle', key: true, w: 42, get: x => x.title },
  { k: 'due', h: 'fDue', key: true, w: 12, date: true, get: x => x.due || '' },
  { k: 'time', h: 'fTime', w: 9, get: x => x.due_time || '' },
  { k: 'status', h: 'fStatus', key: true, w: 14, get: x => t(stKey(x.status)) + (isLate(x) ? ` · ${t('secLate')}` : ''), sort: x => (isLate(x) ? 0 : 1) + ({ wait: 1, review: 1.5, prog: 2, todo: 3, hold: 4, done: 5, cancelled: 6 }[x.status] || 9) / 10 },
  { k: 'unit', h: 'fUnit', key: true, w: 22, get: x => orgOn() ? unitLabel(x.unit_id) : '', on: () => orgOn() && depts().length > 0 },
  { k: 'prio', h: 'fPriority', key: true, w: 11, get: x => PRIO_L(x.priority), sort: x => ({ hi: 0, mid: 1, lo: 2 }[x.priority] ?? 1) },
  { k: 'role', h: 'fRole', w: 13, get: x => t('role_' + (x.role || 'exec')) },
  { k: 'source', h: 'fSource', w: 16, get: x => srcLabel(x.source) },
  { k: 'people', h: 'tcPeople', key: true, w: 22, get: x => peopleOf(x).join(L() === 'ar' ? '، ' : ', ') },
  { k: 'project', h: 'tcProject', key: true, w: 18, get: x => x.project || '' },
  { k: 'recur', h: 'fRecur', w: 16, get: x => recurLabel(x.recur) },
  { k: 'updated', h: 'tcUpdated', w: 12, date: true, get: lastTouch },
  { k: 'completed', h: 'tcCompleted', w: 12, date: true, get: x => x.completed_on || '' },
];
const tcols = () => TCOLS.filter(c => !c.on || c.on());
const XCOLS = () => [...tcols(), { k: 'result', h: 'result', w: 30, get: x => x.result || '' }, { k: 'details', h: 'fDetails', w: 34, get: x => x.details || '' }];
function tableRows() {
  const f = UI.tf || {}, st = f.st || 'open';
  let list = live().filter(x => !x.archived && !isLog(x));
  list = list.filter(x => st === 'all' ? true : st === 'open' ? isOpen(x) : st === 'late' ? isLate(x) : st === 'wait' ? isOpen(x) && isWaitLike(x) : !isOpen(x));
  if (f.person) list = list.filter(x => x.waiting_on === f.person || (x.followups || []).some(y => y.person_id === f.person));
  if (f.project) list = list.filter(x => projKey(x.project) === projKey(f.project));
  if (f.unit && orgOn()) list = list.filter(x => inUnit(x, f.unit));
  if (f.from || f.to) list = list.filter(x => { const d = !isOpen(x) ? x.completed_on : x.due; return d && (!f.from || d >= f.from) && (!f.to || d <= f.to); });
  const col = TCOLS.find(c => c.k === (UI.tsort?.k || 'due')) || TCOLS[1], dir = UI.tsort?.d || 1;
  const val = x => col.sort ? col.sort(x) : (col.get(x) || (col.date ? '9999' : '\uffff'));
  return list.sort((a, b) => { const va = val(a), vb = val(b); return (va < vb ? -1 : va > vb ? 1 : 0) * dir || (a.created_at < b.created_at ? -1 : 1); });
}
function tableView() {
  const f = UI.tf || {}, rows = tableRows(), all = UI.tcols === 'all';
  const projects = [...new Set(live().map(x => (x.project || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ar'));
  const ppl = people().filter(p => live().some(x => x.waiting_on === p.id || (x.followups || []).some(y => y.person_id === p.id)));
  const sel = (id, cur, opts) => `<select id="${id}" class="tf">${opts.map(([v, l]) => `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>`;
  const sortK = UI.tsort?.k || 'due', sortD = UI.tsort?.d || 1;
  const cellOf = (c, x) => { const v = c.get(x); if (!v) return '<span class="muted">—</span>'; if (c.date) return `<span class="num">${fmtShort(v)}</span>`; if (c.k === 'time') return `<span class="num">${fmtTime(v)}</span>`; if (c.k === 'title') return `<span dir="auto">${esc(v)}</span>`; if (c.k === 'project') return `<span class="projlink" data-proj="${esc(v)}" dir="auto">${esc(v)}</span>`; return `<span dir="auto">${esc(v)}</span>`; };
  return `<div class="tf-bar">
      ${sel('tfSt', f.st || 'open', [['open', t('tfOpen')], ['late', t('secLate')], ['wait', t('segWaitShort')], ['done', t('segDone')], ['all', t('tfAll')]])}
      ${sel('tfPerson', f.person || '', [['', t('tfAnyPerson')], ...ppl.map(p => [p.id, p.name])])}
      ${sel('tfProject', f.project || '', [['', t('tfAnyProject')], ...projects.map(p => [p, p])])}
      ${orgOn() && depts().length ? sel('tfUnit', f.unit || '', [['', t('tfAnyUnit')], ...depts().flatMap(d => [[d.id, d.name], ...secsOf(d.id).map(sc => [sc.id, '   ↳ ' + sc.name])]), ['-', t('noUnit')]]) : ''}
      <span class="tf-dates"><label class="tf-date"><span>${t('tfFrom')}</span><input type="date" id="tfFrom" value="${esc(f.from || '')}" aria-label="${t('tfFrom')}"></label><label class="tf-date"><span>${t('tfTo')}</span><input type="date" id="tfTo" value="${esc(f.to || '')}" aria-label="${t('tfTo')}"></label></span>
      ${f.person || f.project || f.unit || f.from || f.to || (f.st && f.st !== 'open') ? `<button type="button" class="linkbtn" data-act="tfreset">${t('tfReset')}</button>` : ''}
    </div>
    <div class="tf-sum"><span>${plural(rows.length, 'tasks')}</span><span class="grow"></span><button type="button" class="btn sm mob" data-act="tcols">${t(all ? 'tcKeyCols' : 'tcAllCols')}</button><button type="button" class="btn sm primary" data-act="xlsx" ${rows.length ? '' : 'disabled'}>${sic('download', 16)}${t('exportXlsx')}</button></div>
    ${rows.length ? `<div class="tbl-wrap ${all ? '' : 'compact'}"><table class="tbl"><thead><tr>${tcols().map(c => `<th class="${c.key ? '' : 'opt'} c-${c.k}"><button type="button" data-tsort="${c.k}" aria-sort="${sortK === c.k ? (sortD > 0 ? 'ascending' : 'descending') : 'none'}">${t(c.h)}${sortK === c.k ? `<i>${sortD > 0 ? '▲' : '▼'}</i>` : ''}</button></th>`).join('')}</tr></thead>
      <tbody>${rows.map(x => `<tr data-open="${x.id}" class="${isLate(x) ? 'late' : ''} ${isOpen(x) ? '' : 'done'}" tabindex="0" role="button">${tcols().map(c => `<td class="${c.key ? '' : 'opt'} c-${c.k}">${cellOf(c, x)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`
      : `<div class="empty small"><p>${t('noMatch')}</p></div>`}`;
}
function exportXlsx() {
  const rows = tableRows(); if (!rows.length) return;
  const blob = xlsxBlob({ sheet: t('xlsxSheet'), rtl: L() === 'ar', columns: XCOLS().map(c => ({ title: t(c.h), width: c.w, type: c.date ? 'date' : 'text' })), rows: rows.map(x => XCOLS().map(c => c.k === 'time' ? (x.due_time ? fmtTime(x.due_time) : '') : c.get(x))) });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `rafeeq-tasks-${T()}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  toast(t('xlsxDone', { n: plural(rows.length, 'tasks') }));
}
/* weekly agenda: workdays only, items ordered by time — what the week looks like at a glance */
function weekHTML() {
  const off = UI.wk || 0, td = T();
  let base = weekRange()[0]; if (isWeekend(td) && [0, 1, 2, 3, 4, 5, 6].map(i => addDays(base, i)).filter(d => !isWeekend(d)).every(d => d < td)) base = addDays(base, 7); // on the weekend, «this week» is the coming one
  const start = addDays(base, off * 7);
  const days = [0, 1, 2, 3, 4, 5, 6].map(i => addDays(start, i)).filter(d => !isWeekend(d));
  const open = live().filter(x => isOpen(x) && !x.archived && !isLog(x));
  const late = off === 0 ? open.filter(isLate).length : 0;
  const itemsOf = d => {
    const out = [];
    open.forEach(x => {
      if (x.due === d) out.push({ x, time: x.due_time || '', title: x.title, who: x.status === 'wait' && x.waiting_on ? pname(x.waiting_on) : '' });
      openFus(x).forEach(f => { if (f.due === d) out.push({ x, time: f.time || '', title: f.what || x.title, who: pname(f.person_id), fu: true }); });
    });
    return out.sort((a, b) => (a.time || '99') < (b.time || '99') ? -1 : (a.time || '99') > (b.time || '99') ? 1 : 0);
  };
  const row = it => `<button type="button" class="wk-item ${it.fu ? 'fu' : ''}" data-open="${it.x.id}"><span class="tm num">${it.time ? fmtTime(it.time) : ''}</span><span class="t" dir="auto">${esc(it.title)}</span>${it.who ? `<span class="who" dir="auto">${sic(it.fu ? 'users' : 'clock', 12)} ${esc(it.who.split(' ').slice(0, 2).join(' '))}</span>` : ''}</button>`;
  const dayBox = d => { const its = itemsOf(d); const [dd, mm] = [pd(d).getDate(), pd(d).getMonth() + 1];
    return `<section class="wk-day ${d === td ? 'today' : ''} ${d < td ? 'past' : ''}"><div class="wk-h"><b>${wd(d)}</b><span class="num">${fmtShort(d)}</span>${its.length ? `<span class="cnt">${its.length}</span>` : ''}<span class="grow"></span>${d >= td ? `<button type="button" class="iconbtn sm" data-act="chatfill" data-q="${dd}/${mm} " aria-label="${esc(t('newTask'))}" title="${esc(t('newTask'))}">${sic('plus', 16)}</button>` : ''}</div>
      ${its.length ? its.map(row).join('') : `<div class="wk-empty">${t('wkFree')}</div>`}</section>`; };
  return `<div class="wk-nav"><button type="button" class="iconbtn" data-wk="${off - 1}" aria-label="${esc(t('prevWeek'))}">${sic('back', 18)}</button><b class="num">${fmtShort(days[0] || start)} – ${fmtShort(days[days.length - 1] || addDays(start, 6))}</b><button type="button" class="iconbtn" data-wk="${off + 1}" aria-label="${esc(t('nextWeek'))}">${sic('back', 18).replace('<svg', '<svg style="transform:scaleX(-1)"')}</button></div>
    ${off !== 0 ? `<button type="button" class="linkbtn center" data-wk="0">${t('thisWeek')}</button>` : ''}
    ${late ? `<button type="button" class="wk-late" data-seg="active">${sic('flag', 15)} ${t('wkLate', { n: plural(late, 'tasks') })}</button>` : ''}
    <div class="wk">${days.map(dayBox).join('')}</div>`;
}
/* project page: everything under one project name, plus a project report */
const projKey = s => normAr(String(s || '').trim());
function projectItems(name) {
  const k = projKey(name), ym = thisYM();
  const all = live().filter(x => !x.archived && projKey(x.project) === k);
  const open = sortTasks(all.filter(x => isOpen(x) && !isLog(x)));
  const done = all.filter(x => !isOpen(x) && (x.completed_on || '').startsWith(ym)).sort((a, b) => (b.completed_on || '') > (a.completed_on || '') ? 1 : -1);
  return { all, open, late: open.filter(isLate), wait: open.filter(isWaitLike), rest: open.filter(x => !isLate(x) && !isWaitLike(x)), done, ym };
}
function openProject(name) {
  if (!name) return;
  pushLayer(() => {
    const p = projectItems(name);
    const sub = [t('projOpen', { n: p.open.length }), p.late.length ? t('projLate', { n: p.late.length }) : '', t('projDoneMonth', { n: p.done.length })].filter(Boolean).join(' · ');
    const block = (title, list, cls = '') => list.length ? sec(title, list.length, list.map(x => taskCard(x, { noSwipe: true, status: !isOpen(x), noProject: true })).join(''), cls) : '';
    return topBar(t('projTitle')) + `<div class="layer-b">
      <header class="hdr"><div class="grow"><h1 dir="auto">${esc(name)}</h1><div class="sub">${sub}</div></div></header>
      ${p.all.length ? '' : `<div class="empty small"><p>${t('noTasksHere')}</p></div>`}
      ${block(t('secLate'), p.late, 'late')}${block(t('segWait'), p.wait)}${block(t('projNext'), p.rest)}${block(t('projDoneThisMonth'), p.done)}
      ${p.open.length || p.done.length ? `<div class="save-row"><button type="button" class="btn primary block" data-act="projreport" data-v="${esc(name)}">${ic.spark}${t('projReport')}</button></div>` : ''}
    </div>`;
  });
}
function projectSnapshot(name) {
  const p = projectItems(name);
  const withDue = p.done.filter(x => x.due && !isLog(x)), onTime = withDue.filter(x => x.completed_on <= x.due);
  let fuClosed = 0; p.all.forEach(x => (x.followups || []).forEach(f => { if (f.status === 'done' && (f.closed_on || '').startsWith(p.ym)) fuClosed++; }));
  const item = (x, kind) => ({ kind, title: x.title, project: '', priority: x.priority, status: kind === 'open' ? stKey(x.status) : '', due: x.due, completed_on: x.completed_on, onTime: !!(x.due && x.completed_on && x.completed_on <= x.due), result: (x.result || '').slice(0, 600), details: (x.details || '').slice(0, 600), update: kind === 'open' ? lastNote(x) : '', steps_done: x.steps_done || 0, steps_total: x.steps_total || 0, people: [...new Set([x.waiting_on, ...(x.followups || []).map(f => f.person_id)].filter(Boolean).map(pname))].slice(0, 6), files: filesOfTask(x.id).map(f => ({ name: f.name, kind: kindOf(f) })) });
  const doneT = p.done.filter(x => !isLog(x)), logs = p.done.filter(isLog);
  return { v: 1, lang: L(), period: p.ym, project: name.slice(0, 120), generated_at: nowISO(), generated_on: T(), owner: { name: state.profile.display_name || '', title: state.profile.job_title || '' },
    kpis: { done: doneT.length, onTimeRate: withDue.length ? onTime.length / withDue.length : null, fuClosed, open: p.open.length, logs: logs.length }, byProject: [],
    items: [...doneT.map(x => item(x, 'done')), ...logs.map(x => item(x, 'log')), ...p.open.map(x => item(x, 'open'))].slice(0, 300) };
}
const sec = (title, n, body, cls = '') => `<section class="sec"><div class="sec-h"><h3>${title}</h3>${n != null ? `<span class="cnt ${cls}">${n}</span>` : ''}</div><div class="stack">${body}</div></section>`;
const emptyBox = (title, text, btn = '') => `<div class="empty"><div class="empty-ic">${ic.note}</div><h3>${title}</h3><p>${text}</p>${btn}</div>`;

/* ---------- library components ---------- */
const monthLabel = ym => new Intl.DateTimeFormat(LOC(), { month: 'long', year: 'numeric' }).format(pd(ym + '-01'));
const ftile = f => `<span class="ftile k-${kindOf(f)}" aria-hidden="true">${extLabel(f)}</span>`;
function fileRow(f, { ctx } = {}) {
  const tk = f.task_id && ctx !== 'task' ? taskById(f.task_id) : null;
  const meta = [
    f.size ? `<span class="num"><bdi>${fmtSize(f.size)}</bdi></span>` : '',
    `<span class="num">${fmtShort(localDay(f.created_at))}</span>`,
    tk && !tk.deleted ? `<span class="lnk" dir="auto">${sic('list', 13)} ${esc(tk.title.length > 40 ? tk.title.slice(0, 40) + '…' : tk.title)}</span>` : '',
    f.person_id && ctx !== 'person' ? `<span dir="auto">${sic('user', 13)} ${esc(pname(f.person_id))}</span>` : '',
  ].filter(Boolean).join('');
  return `<button type="button" class="frow" data-file="${f.id}">${ftile(f)}<span class="body"><span class="t" dir="auto">${esc(f.name)}</span><span class="meta">${meta}</span>${f.note ? `<span class="fnote" dir="auto">${esc(f.note)}</span>` : ''}</span></button>`;
}
function driveCard() {
  if (state.drive.linked === null && navigator.onLine) return `<div class="card drive-card muted">${t('driveChecking')}</div>`;
  return `<div class="card drive-card"><div class="dc-top"><span class="mi">${ic.folder}</span><div><b>${t('connectTitle')}</b><p>${t('connectText')}</p></div></div><button type="button" class="gbtn sm" data-act="connectdrive">${gLogo}${t('connectBtn')}</button><p class="note">${t('connectNote')}</p></div>`;
}

/* ---------- uploads ---------- */
const UPQ = []; let upRunning = false;
function uploadsBox(ctx) {
  const items = UPQ.filter(u => (!ctx.task || u.task_id === ctx.task) && (!ctx.person || u.person_id === ctx.person));
  const inner = items.map(u => {
    const st = u.st === 'up' ? t('upUp', { p: Math.round(u.p * 100) }) : u.st === 'wait' ? t('upWait') : u.st === 'done' ? t('upDone') : t(u.err === 'offline' ? 'upOffline' : u.err === 'not_linked' ? 'upNotLinked' : u.err === 'missing_secret' ? 'upSetup' : 'upErr');
    return `<div class="uprow ${u.st}" data-up="${u.id}">${ftile({ name: u.name, mime: u.mime })}<div class="body"><div class="t" dir="auto">${esc(u.name)}</div><div class="bar"><i data-upbar="${u.id}" style="width:${Math.round((u.st === 'done' ? 1 : u.p) * 100)}%"></i></div><div class="meta"><span data-upst="${u.id}">${st}</span><span class="num"><bdi>${fmtSize(u.size)}</bdi></span></div></div>
      ${u.st === 'err' ? (u.err === 'not_linked' ? `<button type="button" class="btn sm primary" data-act="connectdrive">${t('connectShort')}</button>` : `<button type="button" class="iconbtn sm" data-act="retryup" data-up="${u.id}" aria-label="${esc(t('retry'))}">${ic.retry}</button>`) : ''}
      ${u.st !== 'up' && u.st !== 'done' ? `<button type="button" class="iconbtn sm" data-act="dropup" data-up="${u.id}" aria-label="${esc(t('remove'))}">${ic.x}</button>` : ''}</div>`;
  }).join('');
  return `<div class="ups" data-ups="${esc(JSON.stringify(ctx))}">${inner}</div>`;
}
function drawUps() {
  document.querySelectorAll('[data-ups]').forEach(el => { let ctx = {}; try { ctx = JSON.parse(el.dataset.ups); } catch { } el.outerHTML = uploadsBox(ctx); });
}
function pickFiles(ctx = {}) {
  if (state.mode !== 'cloud') { toast(t('libNeedsAccount')); return; }
  if (state.drive.linked !== true) { openConnectSheet(); return; }
  const inp = document.createElement('input');
  inp.type = 'file'; inp.multiple = true; inp.hidden = true;
  inp.onchange = () => { queueFiles([...inp.files], ctx); inp.remove(); };
  document.body.appendChild(inp); inp.click();
  setTimeout(() => inp.isConnected && !inp.files?.length && inp.remove(), 120000);
}
function queueFiles(list, ctx) {
  if (!list.length) return;
  list.forEach(file => {
    if (file.size > MAX_BYTES) { toast(t('tooBig', { n: file.name }), null, 'err'); return; }
    UPQ.push({ id: uuid(), file, name: file.name || 'file', size: file.size, mime: file.type || '', task_id: ctx.task || null, person_id: ctx.person || (ctx.task ? null : null), p: 0, st: 'wait', err: '' });
  });
  drawUps(); runUploads();
}
async function runUploads() {
  if (upRunning) return; upRunning = true;
  let done = 0;
  try {
    for (;;) {
      const u = UPQ.find(x => x.st === 'wait'); if (!u) break;
      u.st = 'up'; u.p = 0; drawUps();
      try {
        const r = await uploadFile(u.file, { onProgress: p => { u.p = p; const b = document.querySelector(`[data-upbar="${u.id}"]`); if (b) b.style.width = Math.round(p * 100) + '%'; const st = document.querySelector(`[data-upst="${u.id}"]`); if (st) st.textContent = t('upUp', { p: Math.round(p * 100) }); } });
        const fr = newFile({ drive_id: r.id, name: r.name || u.name, mime: r.mimeType || u.mime, size: Number(r.size) || u.size, task_id: u.task_id, person_id: u.person_id });
        UPQ.splice(UPQ.indexOf(u), 1); done++;
        putFile(fr);
        const x = u.task_id ? taskById(u.task_id) : null;
        if (x) { addLog(x, 'file_added', fr.name); putTask(x); }
      } catch (e) {
        u.st = 'err'; u.err = e?.code === 'server' && e.detail === 'missing_secret' ? 'missing_secret' : (e?.code || 'drive');
        console.warn('upload failed', e);
        if (u.err === 'not_linked' || u.err === 'offline' || u.err === 'missing_secret') UPQ.forEach(x => { if (x.st === 'wait') { x.st = 'err'; x.err = u.err; } });
      }
      drawUps();
    }
  } finally { upRunning = false; }
  if (done && !UPQ.length) toast(plural(done, 'uploaded'));
}
function openConnectSheet() {
  openSheet(t('connectTitle'), `<div class="connect"><span class="mi lg">${ic.folder}</span><p>${t('connectText')}</p><ul class="ticks"><li>${sic('check', 15)}${t('connectP1')}</li><li>${sic('check', 15)}${t('connectP2')}</li><li>${sic('check', 15)}${t('connectP3')}</li></ul>
    <button type="button" class="gbtn" data-act="connectgo">${gLogo}${t('connectBtn')}</button><p class="note">${t('connectNote')}</p></div>`);
}

/* ---------- file details ---------- */
const blobCache = new Map();
function openFile(id) {
  const Lr = pushLayer(() => {
    const f = fileById(id);
    if (!f || f.deleted) return topBar('') + `<div class="layer-b"><div class="empty small"><p>${t('notFound')}</p></div></div>`;
    const tk = f.task_id ? taskById(f.task_id) : null;
    return topBar(`<span dir="auto">${esc(f.name)}</span>`, `<button type="button" class="iconbtn" data-act="editfile" data-fid="${f.id}" aria-label="${esc(t('edit'))}">${ic.edit}</button>`) + `<div class="layer-b">
      <div class="fprev k-${kindOf(f)}" data-keep="pv-${f.id}" data-pv="${f.id}"></div>
      <div class="save-row"><a class="btn primary block" href="${esc(viewUrl(f.drive_id))}" target="_blank" rel="noopener">${sic('ext', 18)}${t('openInDrive')}</a><button type="button" class="btn block" data-act="dlfile" data-fid="${f.id}">${sic('download', 18)}${t('download')}</button></div>
      <div class="kv num">
        <div><small>${t('fType')}</small><b>${extLabel(f)}</b></div>
        <div><small>${t('fSize')}</small><b><bdi>${fmtSize(f.size) || '—'}</bdi></b></div>
        <div><small>${t('fUploaded')}</small><b>${wd(localDay(f.created_at))} ${fmtShort(localDay(f.created_at))}</b></div>
        <div><small>${t('fPerson')}</small><b dir="auto">${f.person_id ? `<button type="button" class="linkbtn inl" data-person="${esc(f.person_id)}">${esc(pname(f.person_id))}</button>` : t('noLink')}</b></div>
        <div class="wide"><small>${t('fTask')}</small><b dir="auto">${tk && !tk.deleted ? `<button type="button" class="linkbtn inl" data-open="${tk.id}">${esc(tk.title)}</button>` : t('noLink')}</b></div>
        ${f.note ? `<div class="wide"><small>${t('fNote')}</small><b class="d-text" dir="auto" style="margin:0">${esc(f.note)}</b></div>` : ''}
      </div>
      <div class="danger-row"><button type="button" class="btn sm ghost danger" data-act="delfile" data-fid="${f.id}">${sic('trash', 16)}${t('delete')}</button></div>
    </div>`;
  });
  const f = fileById(id); if (f) mountPreview(Lr.el, f);
}
async function mountPreview(root, f) {
  const box = root.querySelector(`[data-pv="${f.id}"]`); if (!box || box.dataset.loaded) return;
  box.dataset.loaded = '1';
  const fallback = () => { box.innerHTML = `<iframe src="${esc(previewUrl(f.drive_id))}" title="${esc(f.name)}" loading="lazy" allow="autoplay" referrerpolicy="no-referrer"></iframe>`; };
  if (kindOf(f) !== 'img') { fallback(); return; }
  box.innerHTML = `<div class="pv-wait"><span class="spin"></span></div>`;
  try {
    let url = blobCache.get(f.drive_id);
    if (!url) { url = URL.createObjectURL(await fileBlob(f.drive_id)); blobCache.set(f.drive_id, url); }
    box.innerHTML = `<img src="${url}" alt="${esc(f.name)}">`;
  } catch (e) { console.warn(e); fallback(); }
}
async function downloadFile(f) {
  toast(t('downloading'));
  try {
    const blob = await fileBlob(f.drive_id);
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = f.name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } catch (e) {
    console.warn(e);
    if (e?.code === 'not_linked') return openConnectSheet();
    window.open(`https://drive.google.com/uc?export=download&id=${encodeURIComponent(f.drive_id)}`, '_blank', 'noopener');
  }
}
function taskOptions(sel) {
  const open = sortTasks(live().filter(isOpen)), rest = live().filter(x => !isOpen(x)).sort((a, b) => (b.completed_on || '') > (a.completed_on || '') ? 1 : -1).slice(0, 60);
  const cut = s => s.length > 70 ? s.slice(0, 70) + '…' : s;
  return `<option value="">${t('noLink')}</option>${open.length ? `<optgroup label="${esc(t('segActive'))}">${open.map(x => `<option value="${x.id}" ${sel === x.id ? 'selected' : ''}>${esc(cut(x.title))}</option>`).join('')}</optgroup>` : ''}${rest.length ? `<optgroup label="${esc(t('segDone'))}">${rest.map(x => `<option value="${x.id}" ${sel === x.id ? 'selected' : ''}>${esc(cut(x.title))}</option>`).join('')}</optgroup>` : ''}`;
}
function openFileForm(f) {
  openSheet(t('editFile'), `<form id="fForm" class="form" autocomplete="off">
      <label class="fld"><span>${t('fName')}</span><input name="name" required maxlength="300" dir="auto" value="${esc(f.name)}"></label>
      <div class="fld"><span>${t('fTask')}</span><select name="task">${taskOptions(f.task_id)}</select>${live().length ? '' : `<small class="note" style="margin:0">${t('noTasksYet')}</small>`}<button type="button" class="linkbtn start" data-act="filetask" data-fid="${f.id}">${sic('plus', 15)} ${t('newTaskFromFile')}</button></div>
      <div class="fld"><span>${t('fPerson')}</span><select name="person"><option value="">${t('noLink')}</option>${people().map(p => `<option value="${p.id}" ${f.person_id === p.id ? 'selected' : ''}>${esc(p.name)}</option>`).join('')}</select>${people().length ? '' : `<small class="note" style="margin:0">${t('noPeopleYet')}</small>`}</div>
      <label class="fld"><span>${t('fNote')} <em>${t('optional')}</em></span><textarea name="note" rows="2" dir="auto" placeholder="${esc(t('fNotePh'))}">${esc(f.note)}</textarea></label>
      <div class="save-row"><button class="btn primary block" type="submit">${ic.check}${t('save')}</button></div></form>`, {
    onMount: sh => sh.querySelector('#fForm').onsubmit = async e => {
      e.preventDefault(); const fd = new FormData(e.target);
      const name = (fd.get('name') || '').trim().replace(/\s+/g, ' '); if (!name) return;
      const prevName = f.name, prevTask = f.task_id;
      f.name = name; f.task_id = fd.get('task') || null; f.person_id = fd.get('person') || null; f.note = (fd.get('note') || '').trim();
      putFile(f); closeSheet(true); toast(t('saved'));
      if (f.task_id && f.task_id !== prevTask) { const x = taskById(f.task_id); if (x) { addLog(x, 'file_added', f.name); putTask(x); } }
      if (name !== prevName) { try { await renameFile(f.drive_id, name); } catch (err) { console.warn(err); toast(t('renameDriveFail'), null, 'err'); } }
    }
  });
  sheetState.guard = true;
}
function deleteFile(f) {
  confirmSheet(t('delFileQ'), t('delete'), async () => {
    try { await trashFile(f.drive_id); }
    catch (e) { if (e?.code !== 'not_found') { toast(e?.code === 'offline' ? t('needOnline') : t('delFileFail'), null, 'err'); return; } }
    f.deleted = true; putFile(f); popLayer(); toast(t('fileTrashed'));
  });
}

/* ================= views ================= */
const VIEWS = {
  today() {
    const td = T();
    const act = live().filter(isOpen).filter(x => !x.archived);
    const late = sortTasks(act.filter(isLate));
    const dueToday = sortTasks(act.filter(x => x.due === td && !isWaitLike(x)));
    const fus = []; act.forEach(x => openFus(x).forEach(f => { if (f.due && f.due <= td) fus.push([x, f]); }));
    fus.sort((a, b) => a[1].due < b[1].due ? -1 : 1);
    const waiting = sortTasks(act.filter(x => isWaitLike(x) && !isLate(x)));
    const soon = sortTasks(act.filter(x => x.due && x.due > td && x.due <= addDays(td, 7) && !isWaitLike(x)));
    const nodue = sortTasks(act.filter(x => !x.due && x.status === 'prog'));
    const actionable = act.filter(x => x.status === 'todo' || x.status === 'prog');
    const pick = sortTasks(actionable.filter(isLate))[0] || sortTasks(actionable.filter(x => x.due === td))[0] || sortTasks(actionable.filter(x => x.priority === 'hi'))[0] || sortTasks(actionable.filter(x => x.status === 'prog'))[0] || sortTasks(actionable)[0];
    const why = !pick ? '' : isLate(pick) ? lateTxt(pick.due) : pick.due === td ? t('dueToday') : pick.priority === 'hi' ? t('prioHi') : pick.status === 'prog' ? t('stProg') : pick.due ? t('nearestDue') : '';
    let hijri = ''; try { hijri = new Intl.DateTimeFormat((L() === 'ar' ? 'ar-SA' : 'en') + '-u-ca-islamic-umalqura-nu-latn', { day: 'numeric', month: 'long', year: 'numeric' }).format(new Date()); } catch { }
    const name = (state.profile.display_name || '').split(' ')[0];
    let h = `<header class="hdr"><div class="grow"><div class="sub">${esc(hijri)}</div><h1>${wd(td)}${L() === 'ar' ? '،' : ','} ${fmt(td)}</h1>${name ? `<div class="hello">${t('hello', { n: esc(name) })}</div>` : ''}</div><div class="hdr-side">${syncChip()}</div><button type="button" class="iconbtn addman mob" data-act="chat" aria-label="${esc(t('newTask'))}">${ic.plus}</button>${meBtn()}</header>`;
    if (!live().length) {
      return h + emptyBox(t('emptyTitle'), t('emptyText'), `<button type="button" class="btn primary lg" data-act="chat">${ic.spark}${t('addFirst')}</button>`) + leftoversBanner();
    }
    h += leftoversBanner();
    h += `<div class="counters num"><button type="button" class="counter late" data-jump="s-late"><b>${late.length}</b><span>${t('cLate')}</span></button><button type="button" class="counter" data-jump="s-today"><b>${dueToday.length}</b><span>${t('cToday')}</span></button><button type="button" class="counter fu" data-jump="s-fu"><b>${fus.length}</b><span>${t('cFu')}</span></button></div>`;
    if (pick) h += `<div class="start"><div class="eyebrow">${sic('flag', 15)}${t('startWith')}</div><h2 dir="auto">${esc(pick.title)}</h2>${why ? `<div class="why ${isLate(pick) ? 'late' : ''}">${why}</div>` : ''}<div class="row"><button type="button" class="btn primary" data-act="quickdone" data-id="${pick.id}">${ic.check}${t('complete')}</button><button type="button" class="btn" data-open="${pick.id}">${t('openDetails')}</button></div></div>`;
    if (late.length || dueToday.length || soon.length) h += `<p class="hint-swipe">${t('swipeHint')}</p>`;
    if (late.length) h += `<div id="s-late">${sec(t('secLate'), late.length, late.map(x => taskCard(x)).join(''), 'late')}</div>`;
    if (dueToday.length) h += `<div id="s-today">${sec(t('secToday'), dueToday.length, dueToday.map(x => taskCard(x)).join(''))}</div>`;
    if (fus.length) h += `<div id="s-fu">${sec(t('secFu'), fus.length, fus.map(([x, f]) => fuRow(x, f)).join(''))}</div>`;
    if (waiting.length) h += sec(t('secWaiting'), waiting.length, waiting.map(x => taskCard(x)).join(''));
    if (soon.length) h += sec(t('secSoon'), soon.length, soon.map(x => taskCard(x)).join(''));
    if (nodue.length) h += sec(t('secNoDue'), nodue.length, nodue.map(x => taskCard(x)).join(''));
    if (!late.length && !dueToday.length && !fus.length) h += `<div class="allclear">${sic('check', 18)}${t('allClear')}</div>`;
    const doneToday = live().filter(x => x.status === 'done' && x.completed_on === td).sort((a, b) => (b.client_ts || '') > (a.client_ts || '') ? 1 : -1);
    h += `<section class="sec" id="s-done"><div class="sec-h"><h3>${t('secDoneToday')}</h3>${doneToday.length ? `<span class="cnt">${doneToday.length}</span>` : ''}<span class="grow"></span><button type="button" class="btn sm ghost" data-act="newlog">${sic('plus', 16)}${t('addLog')}</button></div><div class="stack">${doneToday.length ? doneToday.map(x => taskCard(x)).join('') : `<div class="box"><div class="fu muted">${t('noDoneToday')}</div></div>`}</div></section>`;
    return h;
  },
  tasks() {
    if (UI.tview === 'table') return tasksHead('table') + tableView();
    if (UI.seg === 'week') return tasksHead('week') + weekHTML();
    let list = live();
    if (UI.seg === 'active') list = list.filter(x => isOpen(x) && !isWaitLike(x) && !x.archived);
    if (UI.seg === 'wait') list = list.filter(x => isOpen(x) && !x.archived && isWaitLike(x));
    if (UI.seg === 'done') list = list.filter(x => !isOpen(x) && !x.archived);
    if (UI.seg === 'archived') list = list.filter(x => x.archived);
    if (UI.personF) list = list.filter(x => x.waiting_on === UI.personF || (x.followups || []).some(f => f.person_id === UI.personF));
    if (UI.unitF && orgOn()) list = list.filter(x => inUnit(x, UI.unitF));
    if (UI.q.trim()) { const q = normAr(UI.q.trim()); list = list.filter(x => normAr([x.title, x.details, x.project, x.source, x.notes, x.result, ...(x.followups || []).map(f => f.what + ' ' + pname(f.person_id)), ...(x.log || []).map(l => l.text)].join(' ')).includes(q)); }
    list = UI.seg === 'done' ? list.sort((a, b) => (b.completed_on || '') > (a.completed_on || '') ? 1 : -1) : sortTasks(list);
    const shown = list.slice(0, UI.listMax || 80);
    const listHTML = `${shown.length ? shown.map(x => taskCard(x, { status: UI.seg !== 'active' })).join('') : `<div class="empty small"><p>${UI.q.trim() || UI.personF || UI.unitF ? t('noMatch') : t('noTasksHere')}</p></div>`}${list.length > shown.length ? `<button type="button" class="linkbtn center" data-act="morelist">${t('showMore', { n: list.length - shown.length })}</button>` : ''}`;
    if (UI.listOnly) return listHTML;
    return tasksHead('list') + `
      ${UI.personF || (UI.unitF && orgOn()) ? `<div class="chips filter">${UI.personF ? `<button type="button" class="chip on" data-pf="${UI.personF}" dir="auto">${esc(pname(UI.personF))} ${sic('x', 14)}</button>` : ''}${UI.unitF && orgOn() ? `<button type="button" class="chip on" data-unitf="${UI.unitF}" dir="auto">${sic('org', 13)} ${esc(UI.unitF === '-' ? t('noUnit') : unitLabel(UI.unitF))} ${sic('x', 14)}</button>` : ''}${UI.unitF && UI.unitF !== '-' && secsOf(deptOf(UI.unitF)?.id).length ? secsOf(deptOf(UI.unitF).id).map(sc => `<button type="button" class="chip ${UI.unitF === sc.id ? 'on' : ''}" data-unitf="${UI.unitF === sc.id ? deptOf(sc.id).id : sc.id}" data-sec="1" dir="auto">${esc(sc.name)}</button>`).join('') : ''}</div>` : ''}
      <div class="stack" id="taskList">${listHTML}</div>
      ${UI.seg === 'done' || UI.seg === 'archived' ? `<button type="button" class="linkbtn center" data-seg="${UI.seg === 'archived' ? 'done' : 'archived'}">${UI.seg === 'archived' ? t('backToDone') : t('showArchived')}</button>` : ''}`;
  },
  people() {
    const ppl = people();
    const mine = myFollowups();
    const view = UI.pview || (mine.length ? 'mine' : 'all');
    const lateN = mine.filter(g => g.late).length;
    const sub = view === 'mine' ? (mine.length ? `${plural(mine.length, 'people')}${lateN ? ` · ${t('nLateShort', { n: lateN })}` : ''}` : t('peopleSub')) : t('peopleSub');
    const head = `<header class="hdr"><div class="grow"><h1>${t('navPeople')}</h1><div class="sub">${sub}</div></div><button type="button" class="btn primary sm" data-act="newperson">${ic.plus}${t('add')}</button>${meBtn()}</header>
      <div class="seg two" role="tablist">${[['mine', 'pvMine'], ['all', 'pvAll']].map(([k, l]) => `<button type="button" role="tab" aria-selected="${view === k}" class="${view === k ? 'on' : ''}" data-pview="${k}">${t(l)}${k === 'mine' && mine.length ? ` <span class="cnt ${lateN ? 'late' : ''}">${mine.length}</span>` : ''}</button>`).join('')}</div>`;
    if (view === 'mine') return head + (mine.length ? `<div class="stack">${mine.map(mfCard).join('')}</div>` : emptyBox(t('mfEmptyTitle'), t('mfEmptyText'), ppl.length ? `<button type="button" class="btn" data-pview="all">${t('pvAll')}</button>` : ''));
    const stats = id => { let open = 0, late = 0, last = null; live().forEach(x => { (x.followups || []).forEach(f => { if (f.person_id !== id) return; if (f.status === 'open' && isOpen(x)) { open++; if (fuLate(f)) late++; } (f.log || []).forEach(l => { if (!last || l.date > last) last = l.date; }); }); if (x.waiting_on === id && x.status === 'wait') open++; }); return { open, late, last }; };
    return head + `
      ${ppl.length ? `<div class="stack">${ppl.map(p => { const s = stats(p.id); return `<button type="button" class="prow" data-person="${p.id}"><span class="av lg">${esc(initials(p.name))}</span><span class="body"><span class="t" dir="auto">${esc(p.name)}</span>${p.org ? `<span class="sub" dir="auto">${esc(p.org)}</span>` : ''}<span class="meta">${s.open ? `<span class="pill ${s.late ? 'late' : 'prog'}">${plural(s.open, 'openFu')}</span>` : ''}<span>${s.last ? t('lastContact', { d: rel(s.last) }) : t('noContact')}</span></span></span>${sic('back', 18).replace('<svg', '<svg class="chev"')}</button>`; }).join('')}</div>` : emptyBox(t('noPeople'), t('noPeopleText'), `<button type="button" class="btn primary" data-act="newperson">${ic.plus}${t('addPerson')}</button>`)}`;
  },
  library() {
    const all = files();
    const cloud = state.mode === 'cloud';
    const linked = state.drive.linked === true;
    const head = `<header class="hdr"><div class="grow"><h1>${t('navLibrary')}</h1><div class="sub">${cloud && all.length ? plural(all.length, 'files') : t('libSub')}</div></div>${cloud ? `<button type="button" class="btn primary sm" data-act="upload">${ic.upload}${t('uploadFile')}</button>` : ''}${meBtn()}</header>`;
    if (!cloud) return head + emptyBox(t('libLocalTitle'), t('libLocalText'), cloudReady ? `<button type="button" class="gbtn sm" style="max-width:320px;margin:0 auto" data-act="google">${gLogo}${t('continueGoogle')}</button>` : '');
    let h = head;
    if (!linked) h += driveCard();
    h += uploadsBox({});
    if (!all.length) return h + (linked ? emptyBox(t('libEmptyTitle'), t('libEmptyText'), `<button type="button" class="btn primary lg" data-act="upload">${ic.upload}${t('uploadFile')}</button>`) : '');
    const kinds = [['all', 'tAll'], ['pdf', 'tPdf'], ['sheet', 'tSheet'], ['doc', 'tDoc'], ['img', 'tImg'], ['other', 'tOther']];
    const kindMatch = f => UI.ltype === 'all' || (UI.ltype === 'other' ? ['other', 'slides'].includes(kindOf(f)) : kindOf(f) === UI.ltype);
    const months = [...new Set(all.map(f => localDay(f.created_at).slice(0, 7)))].sort().reverse();
    if (UI.lmonth && !months.includes(UI.lmonth)) UI.lmonth = '';
    let list = all.filter(kindMatch);
    if (UI.lmonth) list = list.filter(f => localDay(f.created_at).startsWith(UI.lmonth));
    if (UI.lq.trim()) { const q = normAr(UI.lq.trim()); list = list.filter(f => normAr([f.name, f.note, f.task_id ? taskById(f.task_id)?.title : '', f.person_id ? pname(f.person_id) : ''].join(' ')).includes(q)); }
    const counts = Object.fromEntries(kinds.map(([k]) => [k, k === 'all' ? all.length : all.filter(f => k === 'other' ? ['other', 'slides'].includes(kindOf(f)) : kindOf(f) === k).length]));
    h += `<label class="search">${ic.search}<input type="search" id="libQ" value="${esc(UI.lq)}" placeholder="${esc(t('searchFiles'))}" aria-label="${esc(t('searchFiles'))}" dir="auto"></label>
      <div class="lib-filters"><div class="people-row">${kinds.filter(([k]) => k === 'all' || counts[k]).map(([k, l]) => `<button type="button" class="chip ${UI.ltype === k ? 'on' : ''}" data-ltype="${k}">${t(l)} <i class="num">${counts[k]}</i></button>`).join('')}</div>
      ${months.length > 1 ? `<select class="msel" id="libMonth" aria-label="${esc(t('month'))}"><option value="">${t('allMonths')}</option>${months.map(m => `<option value="${m}" ${UI.lmonth === m ? 'selected' : ''}>${monthLabel(m)}</option>`).join('')}</select>` : ''}</div>`;
    const groups = new Map(); list.forEach(f => { const m = localDay(f.created_at).slice(0, 7); if (!groups.has(m)) groups.set(m, []); groups.get(m).push(f); });
    const listHTML = list.length ? [...groups].map(([m, fs]) => sec(monthLabel(m), fs.length, fs.map(f => fileRow(f)).join(''))).join('') : `<div class="empty small"><p>${t('noMatch')}</p></div>`;
    if (UI.listOnly) return listHTML;
    h += `<div id="libList">${listHTML}</div>`;
    if (linked) h += `<a class="linkbtn center folder-link" href="${rootFolderUrl()}" target="_blank" rel="noopener">${sic('ext', 15)} ${t('openFolder')}</a>`;
    return h;
  },
  more() {
    const p = state.profile, u = state.user;
    if (UI.pushOn === undefined && state.mode === 'cloud' && pushSupported()) { UI.pushOn = null; pushCurrent().then(sb2 => { UI.pushOn = !!sb2; if (UI.tab === 'more') renderMain(); }).catch(() => { UI.pushOn = false; }); }
    return `<header class="hdr"><div class="grow"><h1>${t('navMore')}</h1></div></header>
      <button type="button" class="profile" data-act="profile"><span class="av xl">${esc(initials(p.display_name || '?'))}</span><span class="body"><b dir="auto">${esc(p.display_name || t('yourName'))}</b><span class="sub" dir="auto">${esc(p.job_title || t('addJobTitle'))}</span>${u ? `<span class="sub mail">${esc(u.email || '')}</span>` : `<span class="sub">${t('localAccount')}</span>`}</span>${sic('edit', 18)}</button>
      ${(() => { const h = harvest(thisYM()); const r = h.kpis.onTimeRate == null ? t('noRateYet') : t('onTimeOf', { r: Math.round(h.kpis.onTimeRate * 100) + '%' }); return `<button type="button" class="hcard" data-act="harvest"><span class="mi lg">${ic.spark}</span><span class="grow"><b>${t('rpTitle', { m: monthName(L(), thisYM()) })}</b><small>${t('harvestCardSub', { n: plural(h.kpis.done, 'doneN'), r })}</small></span><span class="hcta">${t('present')} · PDF${state.mode === 'cloud' ? ' · ' + t('shareLink') : ''}</span></button>`; })()}
      <div class="card sync-card">${state.mode === 'cloud' ? `<div class="row-between"><span>${syncChip()}</span><button type="button" class="btn sm" data-act="sync">${t('syncNow')}</button></div><p class="note">${state.sync.last ? t('lastSync', { d: new Intl.DateTimeFormat(LOC(), { hour: 'numeric', minute: '2-digit' }).format(new Date(state.sync.last)) }) : ''}${state.sync.error ? ` · <span class="err-txt">${esc(state.sync.error === 'storage' ? t('storageFull') : state.sync.error)}</span>` : ''}</p>${(state.sync.rejected || []).length ? `<p class="note err-txt">${t('syncRejected', { n: state.sync.rejected.length })} <span dir="ltr">${esc(state.sync.rejected[state.sync.rejected.length - 1].msg)}</span></p>` : ''}` : `<p class="note" style="margin-top:0">${t('localExplain')}</p>${cloudReady ? `<button type="button" class="gbtn sm" data-act="google">${gLogo}${t('continueGoogle')}</button>` : ''}`}</div>
      <div class="menu">
        <button type="button" data-act="lang"><span class="mi">${ic.globe}</span><span class="grow">${t('language')}<small>${L() === 'ar' ? 'العربية' : 'English'}</small></span><span class="val">${L() === 'ar' ? 'English' : 'العربية'}</span></button>
        <button type="button" data-act="theme"><span class="mi">${ic.moon}</span><span class="grow">${t('appearance')}<small>${(p.settings || {}).theme === 'dark' ? t('themeDark') : t('themeLight')}</small></span></button>
        ${orgOn() ? `<button type="button" data-act="org"><span class="mi">${ic.org}</span><span class="grow">${t('orgTitle')}<small>${depts().length ? t('orgCount', { d: depts().length, s: units().filter(u => u.kind === 'sec').length }) : t('orgEmptyHint')}</small></span></button>` : (state.mode === 'cloud' ? `<button type="button" data-act="orgold"><span class="mi">${ic.org}</span><span class="grow">${t('orgTitle')}<small class="warn-txt">${t('orgNeedsDb')}</small></span></button>` : '')}
        <button type="button" data-act="weekend"><span class="mi">${ic.cal}</span><span class="grow">${t('wkTitle')}<small>${wkLabel()}</small></span></button>
        <button type="button" data-act="askmore"><span class="mi">${ic.spark}</span><span class="grow">${t('askMoreSetting')}<small>${t(askMore() ? 'askMoreOn' : 'askMoreOff')}</small></span><span class="switch ${askMore() ? 'on' : ''}" aria-hidden="true"></span></button>
        <button type="button" data-act="asksrc"><span class="mi">${ic.users}</span><span class="grow">${t('askSrcSetting')}<small>${t(askSource() ? 'askSrcOn' : 'askSrcOff')}</small></span><span class="switch ${askSource() ? 'on' : ''}" aria-hidden="true"></span></button>
        <button type="button" data-act="tour"><span class="mi">${ic.info}</span><span class="grow">${t('tourAgain')}<small>${t('tourAgainHint')}</small></span></button>
      </div>
      ${state.mode === 'cloud' ? `<div class="menu"><button type="button" data-act="notifs"><span class="mi">${ic.bell}</span><span class="grow">${t('notifTitle')}<small>${notifSummary()}</small></span></button></div>` : ''}
      ${state.mode === 'cloud' ? `<div class="menu"><button type="button" data-act="shares"><span class="mi">${ic.link}</span><span class="grow">${t('sharesTitle')}<small>${t('sharesHint')}</small></span></button></div>` : ''}
      ${state.mode === 'cloud' ? `<div class="menu"><button type="button" data-act="aitoggle"><span class="mi">${ic.spark}</span><span class="grow">${t('aiSetting')}<small>${UI.aiBlocked === 'missing_key' ? t('aiNotReady') : UI.aiBlocked === 'quota' ? t('aiQuota') : t('aiSettingHint')}</small></span><span class="switch ${(p.settings || {}).ai !== false ? 'on' : ''}" aria-hidden="true"></span></button></div>` : ''}
      ${state.mode === 'cloud' ? `<div class="menu">${state.drive.linked === true
        ? `<a href="${rootFolderUrl()}" target="_blank" rel="noopener"><span class="mi">${ic.folder}</span><span class="grow">Google Drive<small>${t('driveOn')}</small></span>${sic('ext', 16)}</a><button type="button" data-act="unlinkdrive"><span class="mi">${ic.link}</span><span class="grow">${t('unlinkDrive')}</span></button>`
        : `<button type="button" data-act="connectdrive"><span class="mi">${ic.folder}</span><span class="grow">${t('connectBtn')}<small>${t('driveOff')}</small></span></button>`}</div>` : ''}
      <div class="sec-h"><h3>${t('backupTitle')}</h3></div>
      <div class="menu">
        <button type="button" data-act="export"><span class="mi">${ic.download}</span><span class="grow">${t('exportBackup')}<small>${t('exportHint')}</small></span></button>
        <button type="button" data-act="import"><span class="mi">${ic.upload}</span><span class="grow">${t('importBackup')}<small>${t('importHint')}</small></span></button>
      </div>
      <input type="file" id="importFile" accept="application/json,.json" hidden>
      <div class="menu">
        <button type="button" data-act="signout" class="danger"><span class="mi">${ic.logout}</span><span class="grow">${state.mode === 'cloud' ? t('signOut') : t('leaveLocal')}</span></button>
      </div>
      <p class="about">${t('about', { v: window.RAFEEQ_VERSION || '' })} · <a href="privacy.html">${t('privacy')}</a></p>`;
  },
};
function v1Data() {
  try { if (localStorage.getItem('rafeeq2.v1imported')) return null; const d = JSON.parse(localStorage.getItem('rafeeq.v1') || 'null'); return d && Array.isArray(d.tasks) && d.tasks.length ? d : null; } catch { return null; }
}
function leftoversBanner() {
  const v1 = v1Data();
  if (v1) return `<div class="banner"><span>${t('v1Found', { n: plural(v1.tasks.length, 'tasks') })}</span><button type="button" class="btn sm primary" data-act="importv1">${t('importNow')}</button><button type="button" class="btn sm ghost" data-act="skipv1">${t('notNow')}</button></div>`;
  if (state.mode !== 'cloud') return '';
  const c = localLeftovers(); if (!c) return '';
  return `<div class="banner"><span>${t('leftovers', { n: plural(c.tasks.length, 'tasks') })}</span><button type="button" class="btn sm primary" data-act="adopt">${t('uploadThem')}</button></div>`;
}

/* ================= swipe ================= */
function bindSwipes(root) {
  root.querySelectorAll('.swipe').forEach(sw => {
    const card = sw.querySelector('.tcard'); const id = sw.dataset.id;
    const rtl = $('#shell').dir === 'rtl';
    let x0 = null, y0 = 0, dx = 0, moved = false, locked = false;
    card.addEventListener('pointerdown', e => { if (e.button > 0 || e.target.closest('button')) return; x0 = e.clientX; y0 = e.clientY; dx = 0; moved = false; locked = false; card.style.transition = 'none'; });
    card.addEventListener('pointermove', e => {
      if (x0 === null) return;
      const mx = e.clientX - x0, my = e.clientY - y0;
      if (!locked) {
        if (Math.abs(mx) > 10 && Math.abs(mx) > Math.abs(my) * 1.3) { locked = true; try { card.setPointerCapture(e.pointerId); } catch { } }
        else if (Math.abs(my) > 10) { x0 = null; card.style.transition = ''; return; } else return;
      }
      dx = mx; moved = true;
      card.style.transform = `translateX(${dx}px)`;
      const fwd = dx > 0; // right = done, left = later (same in both directions, consistent with the hint)
      sw.classList.toggle('to-done', fwd && Math.abs(dx) > 40); sw.classList.toggle('to-later', !fwd && Math.abs(dx) > 40);
    });
    const end = () => {
      if (x0 === null) return;
      card.style.transition = '';
      if (dx > 90) quickDone(id); else if (dx < -90) postpone(id);
      card.style.transform = ''; sw.classList.remove('to-done', 'to-later'); x0 = null;
      if (!moved) return; setTimeout(() => { moved = false; }, 0); // a cancelled gesture must not swallow the next tap
    };
    card.addEventListener('pointerup', end); card.addEventListener('pointercancel', end);
    card.addEventListener('click', e => { if (moved) { e.stopPropagation(); e.preventDefault(); moved = false; } }, true);
  });
}

/* ================= task actions ================= */
function quickDone(id) {
  const x = taskById(id); if (!x || !isOpen(x)) return;
  const prev = structuredClone(x);
  x.status = 'done'; x.completed_on = T(); x.waiting_on = x.waiting_on || null; addLog(x, 'completed');
  const nx = spawnNext(x);
  putTask(x); UI.keepScroll = true;
  const undo = () => { putTask(prev); if (nx) { nx.deleted = true; putTask(nx); } };
  toast(t('completedToast') + (nx ? ' · ' + t('nextOn', { d: rel(nx.due) }) : ''), undo);
  askCloseFus(x, undo);
}
function askCloseFus(x, undo) {
  const n = openFus(x).length; if (!n) return;
  confirmSheet(t('closeFusQ', { n: plural(n, 'openFu') }), t('closeFus'), () => {
    const y = taskById(x.id) || x;
    openFus(y).forEach(f => { f.status = 'done'; f.closed_on = T(); f.log = f.log || []; f.log.push({ id: uuid(), at: nowISO(), date: T(), text: '', outcome: 'closed_with_task' }); });
    putTask(y); toast(t('fusClosed'), undo); // keep the undo reachable after answering the question
  }, false);
}
function postpone(id) {
  const x = taskById(id); if (!x) return;
  const prev = structuredClone(x);
  const nd = addWorkdays(x.due && x.due > T() ? x.due : T(), 1); addLog(x, 'due', '', { from: x.due, to: nd }); x.due = nd; // never pulls a future task back to tomorrow
  putTask(x); UI.keepScroll = true;
  toast(t('postponedTo', { d: rel(nd) }), () => putTask(prev));
}
function setStatus(x, s) {
  if (x.status === s) return;
  addLog(x, 'status', '', { from: x.status, to: s });
  x.status = s; if (s !== 'wait') x.waiting_on = null;
  putTask(x);
}
const fuLogText = l => l.text || (l.outcome === 'closed_with_task' ? t('closedWithTask') : l.outcome === 'person_deleted' ? t('personDeletedLog') : l.outcome === 'done' ? t('fuDone') : t('fuNotYet'));
function logText(e) {
  const d = e.data || {};
  switch (e.kind) {
    case 'created': return t('logCreated');
    case 'status': return t('logStatus', { a: t(stKey(d.from)), b: t(stKey(d.to)) });
    case 'due': return t('logDue', { d: d.to ? fmtShort(d.to) : t('noDue') });
    case 'completed': return t('logCompleted');
    case 'logged': return t('logLogged');
    case 'recurred': return t('logRecurred');
    case 'spawned': return t('logSpawned', { d: d.due ? fmtShort(d.due) : '' });
    case 'skipped': return t('logSkipped', { d: d.to ? fmtShort(d.to) : '' });
    case 'reopened': return t('logReopened');
    case 'fu_added': return t('logFuAdded', { p: esc(pname(d.person_id)) });
    case 'fu_done': return t('logFuDone', { p: esc(pname(d.person_id)) });
    case 'fu_again': return t('logFuAgain', { p: esc(pname(d.person_id)), d: d.due ? rel(d.due) : '' });
    case 'archived': return t('logArchived');
    case 'unarchived': return t('logUnarchived');
    case 'steps': return t('logSteps', { a: d.done, b: d.total });
    case 'edited': return t('logEdited');
    case 'file_added': return t('logFileAdded');
    case 'fu_reminded': return t('logReminded', { p: esc(pname(d.person_id)) });
    case 'ai': return t('logAI');
    case 'ai_undo': return t('logAIUndo');
    case 'imported': return t('logImported');
    default: return '';
  }
}
const stKey = s => ({ todo: 'stTodo', prog: 'stProg', wait: 'stWait', review: 'stReview', hold: 'stHold', done: 'stDone', cancelled: 'stCancelled' }[s] || 'stTodo');
const OPEN_ST = () => ['todo', 'prog', 'wait', ...(orgOn() ? ['review'] : []), 'hold'];

/* ================= sheets ================= */
let sheetState = null;
function openSheet(title, body, { onMount, wide, cls } = {}) {
  stopMic();
  const sh = $('#sheet');
  sh.className = 'sheet' + (wide ? ' wide' : '') + (cls ? ' ' + cls : '');
  sh.innerHTML = `<div class="grab" aria-hidden="true"></div><div class="sheet-h"><h3>${title}</h3><button type="button" class="iconbtn" data-act="closesheet" aria-label="${esc(t('close'))}">${ic.x}</button></div><div class="sheet-b">${body}</div>`;
  $('#backdrop').classList.add('on');
  requestAnimationFrame(() => sh.classList.add('on'));
  sheetState = { dirty: false, guard: cls !== 'chat' };
  sh.addEventListener('input', () => sheetState && (sheetState.dirty = true), { once: true });
  onMount && onMount(sh);
  setTimeout(() => sh.querySelector('[autofocus]')?.focus(), 250);
}
function closeSheet(force) {
  if (!force && sheetState?.dirty && sheetState.guard) { confirmSheet(t('discardQ'), t('discard'), () => closeSheet(true)); return; }
  $('#sheet').classList.remove('on'); $('#backdrop').classList.remove('on'); stopMic(); sheetState = null;
}
$('#backdrop').addEventListener('click', () => closeSheet());
function confirmSheet(msg, okLabel, onOk, danger = true) {
  const sh = $('#confirm');
  sh.innerHTML = `<div class="confirm-in"><p>${msg}</p><div class="row"><button type="button" class="btn block" id="cNo">${/^(إلغاء|cancel)/i.test(okLabel) ? t('back') : t('cancel')}</button><button type="button" class="btn block ${danger ? 'danger-btn' : 'primary'}" id="cOk">${okLabel}</button></div></div>`;
  sh.hidden = false;
  sh.querySelector('#cNo').onclick = () => { sh.hidden = true; };
  sh.querySelector('#cOk').onclick = () => { sh.hidden = true; onOk(); };
  sh.querySelector('#cOk').focus();
}

/* ================= assistant chat: capture (one or many tasks) + questions ================= */
const SR = window.SpeechRecognition || window.webkitSpeechRecognition || null;
const SELF = '@self';
const askMore = () => (state.profile.settings || {}).askMore !== false; // 2.7.9: after the required answers the chat also asks the useful ones (priority, department) — each can be skipped
const askSource = () => (state.profile.settings || {}).askSource !== false; // on by default (2.7.2): the chat asks «who assigned this?» unless the user turned it off
const srcLabel = s => !s ? '' : s === SELF ? t('srcSelf') : s;
const aiOn = () => state.mode === 'cloud' && (state.profile.settings || {}).ai !== false;
function deviceId() { try { let d = localStorage.getItem('rafeeq2.dev'); if (!d) { d = uuid(); localStorage.setItem('rafeeq2.dev', d); } return d; } catch { return 'x'; } }
function scheduleAI() { /* background refine retired in 2.4: everything is settled in the chat before saving */ }
function matchPerson(name) {
  const key = nameTokens(name).join(' '); if (!key) return null;
  const ex = people().find(p => nameTokens(p.name).join(' ') === key); if (ex) return ex.id;
  const h = findPeople(name, people()), toks = nameTokens(name);
  return h.length === 1 && h[0].ids.length === 1 && toks.length <= 3 && h[0].L === toks.length ? h[0].ids[0] : null; // «خالد محمد» must not resolve to «خالد أحمد»
}
function undoAI(id) {
  const x = taskById(id); if (!x?.ai?.before || x.ai.st !== 'done') return;
  const b = x.ai.before;
  Object.assign(x, { title: b.title, due: b.due, priority: b.priority, status: b.status, waiting_on: b.waiting_on, waiting_what: b.waiting_what });
  x.followups = (x.followups || []).filter(f => b.fuIds.includes(f.id) || (f.log || []).length);
  x.ai = Object.assign({}, x.ai, { st: 'undone' });
  addLog(x, 'ai_undo'); putTask(x);
  (b.created || []).forEach(pid => {
    const used = live().some(y => y.waiting_on === pid || (y.followups || []).some(f => f.person_id === pid)) || files().some(f => f.person_id === pid);
    const p = personById(pid); if (p && !used) { p.deleted = true; putPerson(p); }
  });
  toast(t('aiUndone'));
}

const CH = { msgs: [], hist: [], drafts: [], thinking: false, day: null, pick: new Set() };
// suggestion chips: label in the UI language, canonical question understood on-device
const QCHIPS = [['qToday', 'هات مهامي النهارده'], ['qTomorrow', 'هات مهام بكرة'], ['qLate', 'هات المتأخرات'], ['qFollow', 'مين عنده متابعات واقفة'], ['qWeekDone', 'هات اللي خلصته الأسبوع ده']];
const qChip = k => { const q = QCHIPS.find(c => c[0] === k); return q ? `<button type="button" class="chip" data-act="chatq" data-q="${esc(q[1])}" data-l="${esc(t(k))}">${t(k)}</button>` : ''; };
const workdayPlus = (from, n) => { let d = from; for (let i = 0; i < n; i++) d = addWorkdays(d, 1); return d; };
/** Quick date options that never land on Friday/Saturday. */
function dateChips() {
  const td = T(), out = [];
  if (!isWeekend(td)) out.push([t('today'), td]);
  const n1 = addWorkdays(td, 1); out.push([n1 === addDays(td, 1) ? t('tomorrow') : wd(n1), n1]);
  const n2 = addWorkdays(td, 2); out.push([wd(n2), n2]);
  const nw = addWorkdays(addDays(td, 6), 1); out.push([t('inAWeek'), nw]);
  return out;
}
const recentSources = () => { const m = new Map(); live().forEach(x => { if (x.source && x.source !== SELF) m.set(x.source, (m.get(x.source) || 0) + 1); }); return [...m].sort((a, b) => b[1] - a[1]).map(([s]) => s).slice(0, 5); };
const topPeople = n => { const c = new Map(); live().forEach(x => { [x.waiting_on, ...(x.followups || []).map(f => f.person_id)].filter(Boolean).forEach(id => c.set(id, (c.get(id) || 0) + 1)); }); return people().slice().sort((a, b) => (c.get(b.id) || 0) - (c.get(a.id) || 0)).slice(0, n); };

/* ---------- drafts ---------- */
function refOf(name, desc = '', generic = false) {
  const nm = String(name || '').replace(/\s+/g, ' ').trim(); if (!nm) return null;
  if (generic) return { kind: 'generic', label: nm };
  if (/^(م|مهندس|المهندس|د|دكتور|الدكتور|ا|استاذ|الاستاذ|eng|dr|mr)\.?$/.test(normAr(String(desc || '').trim()))) desc = '';
  const id = matchPerson(nm); return id ? { kind: 'known', id } : { kind: 'new', name: nm, desc: desc || '' };
}
const refName = r => !r ? '' : r.kind === 'known' ? pname(r.id) : r.kind === 'new' ? r.name : r.kind === 'generic' ? r.label : r.label || '?';
function fromLocal(p, raw) {
  const d = { id: uuid(), raw, title: p.title, due: p.due || null, time: p.time || null, recur: p.recur || null, noDue: !!p.noDue && !p.due, weekendOk: false, priority: p.priority || 'mid', role: 'exec', source: null, project: '', waitWhat: p.waiting?.what || null, people: [] };
  const role0 = p.waiting ? 'waiting' : p.fu ? 'followup' : 'related';
  const add = (ref, role, what = '', due = null) => { if (ref) d.people.push({ ref, role, what, due }); };
  if (p.person) {
    const r = p.person.kind === 'new' ? (p.person.generic ? { kind: 'generic', label: p.person.name } : { kind: 'new', name: p.person.name, desc: p.person.desc || '' }) : p.person;
    add(r, role0, p.fu?.what || '', p.fu?.due || null);
    (p.more || []).forEach(m => add(m.kind === 'new' ? { kind: 'new', name: m.name, desc: '' } : m, role0 === 'waiting' ? 'followup' : role0, p.fu?.what || '', p.fu?.due || null));
  }
  if (d.priority !== 'mid') d.prioAsked = true;
  if (p.deleg) d.role = 'follow';
  else if (p.fu) d.role = /^(متابعة|التواصل|سؤال)/.test(p.title) ? 'follow' : 'both';
  else if (p.waiting) d.role = 'follow';
  if (p.source) d.source = p.source.self ? SELF : (personById(matchPerson(p.source.name))?.name || p.source.name);
  if (orgOn() && depts().length) { const u = matchUnit(raw, true); if (u) { d.unit = u; d.unitAsked = true; } else unitFromPeople(d); }
  return d;
}
/** An assigner's name as stored: a known person wins, even behind a role word («المدير عدنان الصلاح» → «عدنان الصلاح»). */
function srcName(s) {
  const raw = String(s || '').replace(/\s+/g, ' ').trim().slice(0, 80); if (!raw) return null;
  const hit = personById(matchPerson(raw)); if (hit) return hit.name;
  const bare = raw.replace(/^(?:(?:ال)?(?:مدير|مديرة|مديري|مديرتي|رئيس|رئيسي|مشرف|مشرفي|مهندس|دكتور|أستاذ|استاذ)|م\.|د\.|أ\.)\s+/, '').trim();
  const hit2 = bare !== raw && bare.split(' ').length >= 2 ? personById(matchPerson(bare)) : null;
  return hit2 ? hit2.name : raw;
}
const SELF_RE = /(من نفسي|مبادره|بنفسي|ذاتي|من تنظيمي|تنظيمي|من عندي|my own|myself|self|كلف|كلفت|خلي |خليت|قول ل|قولي ل|ابلغ|بلغ |اطلب من|اطلبي من|assign|delegate|tell )/;
/** In a list, the line a task came from (best word overlap with its title); the whole text for a single task. */
function lineOf(raw, title, many) {
  if (!many) return raw;
  const toks = s => new Set(nameTokens(s).filter(w => w.length > 2));
  const segs = String(raw || '').split(/\n+|(?:^|\s)\d{1,2}\s*[-.)]\s+|؛/).filter(x => x.trim()).map(seg => [seg, toks(seg)]);
  const hit = (st, w) => st.has(w) || [...st].some(v => v.length > 3 && w.length > 3 && (v.includes(w) || w.includes(v)));
  const tt = [...toks(title)], df = w => segs.filter(([, st]) => hit(st, w)).length || 1; // rare words weigh more («مراجعة» is everywhere)
  let best = '', score = 0;
  segs.forEach(([seg, st]) => { const n = tt.reduce((a, w) => a + (hit(st, w) ? 1 / df(w) : 0), 0); if (n > score) { score = n; best = seg; } });
  return score ? best : '';
}
/* Guards on what the model returns, checked against the user's own words for that task. */
const URG_RE = /(عاجل|مستعجل|ضروري|ضرورى|مهم جدا|مهمه جدا|اولويه|بسرعه|حالا|فورا|فوري|ضغط|urgent|asap|important|critical|high priority)/;
const ALL_RE = /(كلهم|كلها|جميعها|جميعهم|الكل|all of them)/;
const NEXTWK_RE = /الاسبوع\s*(الجاي|القادم|الجاية|الجايه|القادمه|اللي جاي|الي جاي|الجديد)|next week/;
const DAYNAME_RE = /(الاحد|الاثنين|الاتنين|التلات|الثلاثاء|الثلاث|الاربع|الاربعاء|الخميس|الجمعه|السبت|sunday|monday|tuesday|wednesday|thursday|friday|saturday|اخر|نهايه|قبل|بعد|يوم\s*\d|\d{1,2}\s*[/.-]\s*\d{1,2}|اكتوبر|نوفمبر|ديسمبر|يناير|فبراير|مارس|ابريل|مايو|يونيو|يوليو|اغسطس|سبتمبر)/;
const WAIT_RE = /(مستني|مستنى|منتظر|بانتظار|في انتظار|استني|ننتظر|نستني|waiting for|awaiting)/;
const FOLLOW_VERB_RE = /(اتابع|تابع|نتابع|اكلم|كلم|اتصل|ذكر|اذكر|ابلغ|بلغ|اسال|اسأل|follow up|call|remind)/;
function fromAI(a, local, raw = '', many = false) {
  const whole = normAr(raw);
  raw = lineOf(raw, a.title || '', many);
  const ln = normAr(raw || local?.raw || '');
  if (a.priority === 'hi' && !URG_RE.test(ln) && !(ALL_RE.test(whole) && URG_RE.test(whole))) a = { ...a, priority: local?.priority === 'hi' ? 'hi' : 'mid' }; // no stated urgency → never «عالية»
  if (NEXTWK_RE.test(ln) && !DAYNAME_RE.test(ln.replace(NEXTWK_RE, ''))) a = { ...a, due: nextWeekday(T(), weekStartDay()) }; // «الأسبوع الجاي» alone = its first workday
  if (!validDate(a.due) && !a.no_due) { const mo = monthOnly(ln, T()); if (mo) a = { ...a, due: mo }; } // «في ديسمبر» → its last workday
  if (validDate(a.due) && isWeekend(a.due) && /(^|\s)قبل\s/.test(ln)) { let x = a.due; while (isWeekend(x)) x = addDays(x, -1); a = { ...a, due: x }; } // «قبل 15 نوفمبر» on a weekend → the last workday before it
  if (WAIT_RE.test(ln) && !FOLLOW_VERB_RE.test(ln) && Array.isArray(a.people)) a = { ...a, role: a.role === 'both' ? 'follow' : a.role, waiting_what: a.waiting_what || 'reply', people: a.people.map(p => p && p.role === 'followup' ? { ...p, role: 'waiting' } : p) }; // «مستني رد X» = waiting, not a follow-up
  const d = { id: uuid(), raw: local?.raw || '', title: String(a.title || local?.title || '').replace(/\s+/g, ' ').trim().slice(0, 300), due: validDate(a.due) || local?.due || null, weekendOk: false,
    priority: ['hi', 'mid', 'lo'].includes(a.priority) ? a.priority : local?.priority || 'mid', role: ['exec', 'follow', 'both'].includes(a.role) ? a.role : 'exec',
    noDue: !!a.no_due && /(بدون|من غير|دون|مفيش|ملوش|مالوش|ماله|ما له|مالها|ما لها|بلا موعد|مش محدد|مو محدد|مب محدد|غير محدد|no deadline|no due|whenever)/.test(normAr(raw || local?.raw || '')),
    source: (() => { const me = a.source && /^(انا|نفسي|me|myself|self)$/.test(normAr(a.source).trim()); if (a.source && !me) return srcName(a.source); const rw = normAr(raw || local?.raw || ''); return me || (a.source_self && SELF_RE.test(rw)) ? SELF : local?.source || null; })(), // a named assigner wins; «أنا» = self
    project: String(a.project || '').slice(0, 80), waitWhat: a.waiting_what || null, people: [],
    time: /^\d{2}:\d{2}$/.test(a.due_time || '') ? a.due_time : local?.time || null, recur: cleanRecur(a.recur) || local?.recur || null };
  if (d.recur && (!d.due || d.recur.dom === 'last' || isWeekend(d.due))) d.due = recurNext(d.recur, addDays(T(), -1)); // «آخر كل شهر» = last workday
  if (d.recur?.f === 'daily' && d.due > T() && !isWeekend(T()) && !/(بكره|بكرا|بكرة|غدا|من بكره)/.test(ln)) { const nowHM = new Date().toTimeString().slice(0, 5); if (!d.time || d.time > nowHM) d.due = T(); } // «كل يوم» starts today when today is still ahead
  (a.people || []).slice(0, 8).forEach(p => {
    if (!p || !p.name) return;
    const ref = refOf(p.name, p.desc, !!p.generic); if (!ref) return;
    if (d.people.some(q => refName(q.ref) === refName(ref))) return;
    d.people.push({ ref, role: ['followup', 'waiting', 'related'].includes(p.role) ? p.role : 'related', what: String(p.what || '').slice(0, 200), due: validDate(p.due) || null });
  });
  if (!Array.isArray(a.people) && local?.people?.length) d.people = local.people.filter(p => !(d.source && normAr(refName(p.ref)).includes(normAr(d.source).replace(/^(م|د|ا)\.\s*/, '').split(' ')[0])));
  if (d.people.some(p => p.role === 'followup') && d.role === 'exec') d.role = 'both';
  if (d.priority !== 'mid' || /(متوسط|عادي|medium|normal priority)/.test(ln)) d.prioAsked = true;
  if (orgOn() && depts().length) { const u = (a.unit && matchUnit(a.unit)) || matchUnit(ln, true); if (u) { d.unit = u; d.unitAsked = true; } else unitFromPeople(d); }
  return d;
}
const MATCH_STOP = new Set(['علي', 'بخصوص', 'بشان', 'الخاص', 'الخاصه', 'عن', 'الي', 'مع', 'من', 'في', 'and', 'the', 'for', 'with']);
/** An open task that a "done" entry most likely closes, by shared significant words. */
function matchOpenTask(title) {
  const sig = s => nameTokens(s).map(w => w.length > 4 ? w.replace(/^ال/, '') : w).filter(w => w.length > 2 && !MATCH_STOP.has(w));
  const a = sig(title); if (a.length < 2) return null;
  let best = null, bs = 0;
  live().filter(x => isOpen(x) && !x.archived).forEach(x => {
    const b = sig(x.title); if (b.length < 2) return;
    const hit = b.filter(w => a.some(v => v === w || (v.length > 3 && w.length > 3 && (v.includes(w) || w.includes(v))))).length;
    const sc = hit / Math.max(b.length, a.length > b.length + 3 ? a.length : b.length);
    if (hit >= 2 && sc > bs) { bs = sc; best = x; }
  });
  return bs >= 0.6 ? best.id : null;
}
const isLogD = d => d.kind === 'log';
function fromLog(title, date, raw, project = '') {
  const tt = String(title || '').replace(/\s+/g, ' ').trim().slice(0, 300);
  const dt = validDate(date) && date <= T() ? date : T();
  return { id: uuid(), kind: 'log', raw, title: tt, date: dt, project: String(project || '').slice(0, 80), match: matchOpenTask(tt), closeTask: undefined, people: [] };
}
/** First missing piece of a draft, or null when it can be saved. */
function missingOf(d) {
  if (isLogD(d)) return d.match && d.closeTask === undefined ? { k: 'match' } : null;
  if (!d.source && !askSource()) d.source = SELF;
  let i = d.people.findIndex(p => p.ref.kind === 'generic' && p.role !== 'related'); if (i >= 0) return { k: 'generic', i };
  i = d.people.findIndex(p => p.ref.kind === 'amb'); if (i >= 0) return { k: 'amb', i };
  if (!d.due && !d.noDue) return { k: 'due' };
  if (d.due && isWeekend(d.due) && !d.weekendOk) return { k: 'weekend' };
  if (!d.source) return { k: 'source' };
  if ((d.role === 'follow' || d.role === 'both') && !d.people.some(p => p.role !== 'related')) return { k: 'who' };
  return null;
}
/** Optional fields still empty on a complete task draft — offered with «+» chips, never required. */
const optMissing = d => isLogD(d) || missingOf(d) || softOf(d) ? [] : [d.due && !d.time ? 'time' : '', !d.project ? 'project' : ''].filter(Boolean);
const OPT_LBL = { time: 'fTime', project: 'tcProject' };
const extraOn = () => { const e = CH.extra, n = e ? CH.drafts.findIndex(d => d.id === e.id) : -1; if (n < 0 || isLogD(CH.drafts[n]) || pending() || softPending()) return (CH.extra = null); return { n, k: e.k }; };
const pending = () => { for (let n = 0; n < CH.drafts.length; n++) { const m = missingOf(CH.drafts[n]); if (m) return { n, d: CH.drafts[n], m }; } return null; };
/** The questions worth asking once the required ones are answered — never blocking: «تخطَّ» or saving skips them. */
function softOf(d) {
  if (isLogD(d) || !askMore() || d.softSkip) return null;
  if (!d.prioAsked && d.priority === 'mid') return 'prio';
  if (orgOn() && depts().length && !d.unitAsked && !unitById(d.unit)) return 'unit';
  return null;
}
/** A draft takes the department of the people it follows up (set once on each person), unless the user chose one. */
function unitFromPeople(d) {
  if (isLogD(d) || unitById(d.unit) || !orgOn()) return;
  const ids = [...new Set(d.people.filter(p => p.ref.kind === 'known' && p.role !== 'related').map(p => unitOfPerson(p.ref.id)).filter(Boolean))];
  if (ids.length === 1) d.unit = ids[0]; else if (ids.length > 1) { const dd = [...new Set(ids.map(i => deptOf(i)?.id))]; if (dd.length === 1) d.unit = dd[0]; }
}
const recentUnits = () => { const c = new Map(); live().forEach(x => { if (unitById(x.unit_id)) c.set(x.unit_id, (c.get(x.unit_id) || 0) + 1); }); return [...c].sort((a, b) => b[1] - a[1]).map(([id]) => id); };
const softPending = () => { if (pending()) return null; for (let n = 0; n < CH.drafts.length; n++) { const k = softOf(CH.drafts[n]); if (k) return { n, d: CH.drafts[n], k }; } return null; };
const softLeft = k => CH.drafts.filter(d => softOf(d) === k).length;
function softHTML({ n, d, k }) {
  const lead = CH.drafts.length > 1 ? `<b class="qn">${t('taskN', { n: n + 1 })}</b> ` : '';
  const ttl = `«${esc(d.title.length > 60 ? d.title.slice(0, 60) + '…' : d.title)}»`;
  const left = softLeft(k);
  let q = '', opts = '';
  if (k === 'unit') {
    q = t('askUnitQ', { t: ttl });
    const pick = UI.askUnitDept && unitById(UI.askUnitDept) ? UI.askUnitDept : null;
    if (pick) opts = `<button type="button" class="chip primary-chip" data-act="anssoft" data-v="${pick}" ${left > 1 ? '' : ''}>${esc(unitById(pick).name)} ${t('wholeDept')}</button>` + secsOf(pick).map(sc => `<button type="button" class="chip" data-act="anssoft" data-v="${sc.id}" dir="auto">↳ ${esc(sc.name)}</button>`).join('') + `<button type="button" class="chip ghost" data-act="unitback">${t('back')}</button>`;
    else { const ds = [...new Set([...recentUnits().map(i => deptOf(i)?.id).filter(Boolean), ...depts().map(x => x.id)])].slice(0, 10); opts = ds.map(id => `<button type="button" class="chip" data-act="${secsOf(id).length ? 'unitdept' : 'anssoft'}" data-v="${id}" dir="auto">${esc(unitById(id).name)}${secsOf(id).length ? ' ›' : ''}</button>`).join('') + (left > 1 && CH.lastUnit && unitById(CH.lastUnit) ? `<button type="button" class="chip primary-chip" data-act="anssoft" data-all="1" data-v="${CH.lastUnit}" dir="auto">${esc(t('unitForAll', { u: unitLabel(CH.lastUnit, true), n: left }))}</button>` : ''); }
    opts += `<button type="button" class="chip" data-act="anssoft" data-v="">${t('noUnit')}</button>`;
  }
  if (k === 'prio') { q = t('askPrioQ', { t: ttl }); opts = [['hi', 'prioHi'], ['mid', 'prioMid'], ['lo', 'prioLo']].map(([v, l]) => `<button type="button" class="chip" data-act="anssoft" data-v="${v}">${t(l)}</button>`).join('') + (left > 1 ? `<button type="button" class="chip primary-chip" data-act="anssoft" data-v="mid" data-all="1">${t('prioAllMid', { n: left })}</button>` : ''); }
  return `<div class="ask q soft"><p>${lead}${q}</p><div class="opts">${opts}<button type="button" class="chip ghost" data-act="skipsoft">${t('skipQs')}</button></div><small class="note">${t('softNote')}</small></div>`;
}
function setSoft(d, k, v) { if (k === 'prio') { if (['hi', 'mid', 'lo'].includes(v)) d.priority = v; d.prioAsked = true; } if (k === 'unit') { d.unit = unitById(v) ? v : null; d.unitAsked = true; if (d.unit) CH.lastUnit = d.unit; UI.askUnitDept = null; } }
/** Typed answer to a soft question. Returns false when the text is not an answer. */
function answerSoft({ d, k }, text) {
  const n = shortN(text); if (n.split(' ').length > 4) return false;
  if (k === 'unit') {
    if (/^(لا|بدون|من غير|مفيش|ولا اداره|ولا قسم|none|no)$/.test(n)) { setSoft(d, k, ''); return true; }
    const id = matchUnit(text); if (!id) return false; setSoft(d, k, id); return true;
  }
  if (k === 'prio') {
    const v = /(عاليه|عالي|عاجل|مستعجل|ضروري|مهم|high|urgent)/.test(n) ? 'hi' : /(منخفض|منخفضه|مش مستعجل|مش مهم|عادي جدا|low)/.test(n) ? 'lo' : /(متوسط|متوسطه|عادي|عاديه|medium|normal)/.test(n) ? 'mid' : null;
    if (!v) return false; setSoft(d, k, v); return true;
  }
  return false;
}

/* ---------- open / render ---------- */
function openChat(prefill = '') {
  if (CH.day !== T()) { CH.msgs = []; CH.hist = []; CH.drafts = []; CH.day = T(); CH.msgs.push({ who: 'bot', html: briefHTML() }); }
  openSheet(`<span class="chat-t">${sic('spark', 18)} ${t('assistant')}</span>`, `<div class="chat-log" id="chatLog"></div>
    <div class="chat-foot"><div class="chat-sugg" id="chatSugg"></div>
      <div class="composer"><textarea id="chatIn" rows="1" placeholder="${esc(t('chatPh'))}" aria-label="${esc(t('chatPh'))}" dir="auto"></textarea>
        <div class="tools"><span id="recState" class="rec-state"></span><button type="button" class="iconbtn sm clipbtn" data-act="chatattach" aria-label="${esc(t('attachSheet'))}" title="${esc(t('attachSheet'))}">${sic('clip', 18)}</button><button type="button" class="linkbtn sm" data-act="newtask">${sic('edit', 15)} ${t('manualEntry')}</button><button type="button" class="mic" id="micBtn" data-act="mic" aria-label="${esc(t('speak'))}">${ic.mic}</button><button type="button" class="send" id="chatSendBtn" data-act="chatsend" aria-label="${esc(t('send'))}">${ic.send}</button></div></div></div>`, {
    cls: 'chat', onMount: sh => { const ta = sh.querySelector('#chatIn'); ta.addEventListener('input', () => autoGrow(ta)); if (prefill) ta.value = prefill; }
  });
  renderChat(true);
  setTimeout(() => $('#chatIn')?.focus(), 260);
  if (aiOn() && navigator.onLine && Date.now() - (CH.warm || 0) > 240000) { CH.warm = Date.now(); invokeAssist({ v: 2, text: '' }, 8000).catch(() => { }); } // wake the assistant early
}
function renderChat(scroll = true) {
  const log = $('#chatLog'); if (!log) return;
  const pend = pending();
  const drafts = CH.drafts.length || CH.thinking ? `<div class="bubble bot drafts">${CH.drafts.length > 1 ? `<div class="dh">${CH.drafts.every(isLogD) ? t('nLogsFound', { n: plural(CH.drafts.length, 'logs') }) : CH.drafts.some(isLogD) ? t('nItemsFound', { n: plural(CH.drafts.length, 'items') }) : t('nTasksFound', { n: plural(CH.drafts.length, 'tasks') })}</div>` : ''}${CH.drafts.map((d, n) => draftCard(d, n, pend && pend.n === n)).join('')}
      ${CH.thinking ? `<div class="thinking"><span class="dots"><i></i><i></i><i></i></span>${t(CH.drafts.length ? 'improving' : 'understanding')}</div>` : pend ? questionHTML(pend) : softPending() ? softHTML(softPending()) : extraOn() ? extraHTML(extraOn()) : `${optLine()}<div class="ready">${sic('check', 16)} ${CH.drafts.some(isLogD) ? t(CH.drafts.length > 1 ? 'readyItems' : 'readyLog') : CH.drafts.length > 1 ? t('readyAll', { n: plural(CH.drafts.length, 'tasks') }) : t('readyOne')}</div>`}
      ${CH.drafts.length ? `<div class="save-row"><button type="button" class="btn primary block" data-act="chatsave" ${pend || CH.thinking ? 'disabled' : ''}>${ic.check}${CH.drafts.length > 1 ? t('saveAll', { n: CH.drafts.length }) : t('save')}</button><button type="button" class="btn" data-act="chatdiscard">${t('discardDrafts')}</button></div>` : ''}</div>` : '';
  log.innerHTML = CH.msgs.map(m => `<div class="bubble ${m.who}">${m.html}</div>`).join('') + drafts + actionsHTML();
  const sb2 = $('#chatSendBtn'); if (sb2) sb2.disabled = !!CH.thinking;
  const sg = $('#chatSugg');
  if (sg) sg.innerHTML = CH.drafts.length ? '' : `<button type="button" class="chip" data-act="chatfill">${sic('check', 14)}${t('qLog')}</button>` + QCHIPS.map(([k, q]) => `<button type="button" class="chip" data-act="chatq" data-q="${esc(q)}" data-l="${esc(t(k))}">${t(k)}</button>`).join('');
  if (scroll) requestAnimationFrame(() => { log.scrollTop = log.scrollHeight; });
}
function draftCard(d, n, active) {
  if (isLogD(d)) {
    const mt = d.closeTask ? taskById(d.closeTask) : null;
    const chips = [`<span class="chip sm logc">${sic('check', 13)}${t('logPill')}</span>`, `<span class="chip sm">${sic('cal', 13)}${d.date === T() ? t('today') : `${wd(d.date)} ${fmtShort(d.date)}`}</span>`,
      d.project ? `<span class="chip sm" dir="auto">${esc(d.project)}</span>` : '', mt ? `<span class="chip sm">${sic('reopen', 13)}<span dir="auto">${t('closesTask', { t: esc(mt.title.slice(0, 40)) })}</span></span>` : ''].filter(Boolean).join('');
    return `<div class="dcard log ${active ? 'active' : ''} ${missingOf(d) ? '' : 'ok'}">
    <div class="dtop">${CH.drafts.length > 1 ? `<span class="dn num">${n + 1}</span>` : ''}<div class="dt" dir="auto">${esc(d.title)}</div>
      <button type="button" class="iconbtn sm" data-act="draftedit" data-n="${n}" ${CH.thinking ? 'disabled' : ''} aria-label="${esc(t('edit'))}">${sic('edit', 15)}</button>
      ${n > 0 && isLogD(CH.drafts[n - 1]) ? `<button type="button" class="iconbtn sm" data-act="draftmerge" data-n="${n}" ${CH.thinking ? 'disabled' : ''} title="${esc(t('mergePrev'))}" aria-label="${esc(t('mergePrev'))}">${sic('link', 15)}</button>` : ''}
      <button type="button" class="iconbtn sm" data-act="draftdrop" data-n="${n}" ${CH.thinking ? 'disabled' : ''} aria-label="${esc(t('remove'))}">${sic('x', 15)}</button></div>
    <div class="dchips">${chips}</div></div>`;
  }
  const roleTxt = t('role_' + d.role);
  const ppl = d.people.filter(p => p.role !== 'related' || d.role === 'exec');
  const chips = [
    d.due ? `<span class="chip sm">${sic('cal', 13)}${wd(d.due)} ${fmtShort(d.due)}${isWeekend(d.due) ? ` <em class="wk">${t('weekend')}</em>` : ''}</span>` : d.noDue ? `<span class="chip sm">${t('noDue')}</span>` : `<span class="chip sm missing">${sic('cal', 13)}${t('noDueYet')}</span>`,
    d.time ? `<span class="chip sm">${sic('clock', 13)}<span class="num">${fmtTime(d.time)}</span></span>` : '',
    d.recur ? `<span class="chip sm">${sic('repeat', 13)}${recurLabel(d.recur)}</span>` : '',
    d.priority === 'hi' ? `<span class="chip sm hi">${t('prioHi')}</span>` : d.priority === 'lo' ? `<span class="chip sm">${t('prioLo')}</span>` : '',
    `<span class="chip sm">${roleTxt}</span>`,
    d.source === SELF && !askSource() ? '' : d.source ? `<span class="chip sm">${t('srcShort')}: ${esc(srcLabel(d.source))}</span>` : `<span class="chip sm missing">${t('srcShort')}: ${L() === 'ar' ? '؟' : '?'}</span>`,
    ...ppl.map(p => `<span class="chip sm ${p.ref.kind === 'generic' || p.ref.kind === 'amb' ? 'missing' : ''}">${sic(p.role === 'waiting' ? 'clock' : 'users', 13)}<span dir="auto">${esc(refName(p.ref))}</span>${p.ref.kind === 'new' && p.ref.desc ? `<small dir="auto">· ${esc(p.ref.desc)}</small>` : ''}${p.ref.kind === 'new' ? ` <em>${t('newBadge')}</em>` : ''}</span>`),
    d.project ? `<span class="chip sm" dir="auto">${esc(d.project)}</span>` : '',
    unitById(d.unit) ? `<span class="chip sm" dir="auto">${sic('org', 13)}${esc(unitLabel(d.unit, true))}</span>` : '',
    ...(CH.thinking ? [] : optMissing(d).map(k => `<button type="button" class="chip sm plus ${CH.extra && CH.extra.id === d.id && CH.extra.k === k ? 'on' : ''}" data-act="draftadd" data-n="${n}" data-k="${k}" aria-label="${esc(t('addField', { f: t(OPT_LBL[k]) }))}"><b aria-hidden="true">+</b>${t(OPT_LBL[k])}</button>`)),
  ].filter(Boolean).join('');
  return `<div class="dcard ${active ? 'active' : ''} ${missingOf(d) ? '' : 'ok'}">
    <div class="dtop">${CH.drafts.length > 1 ? `<span class="dn num">${n + 1}</span>` : ''}<div class="dt" dir="auto">${esc(d.title)}</div>
      <button type="button" class="iconbtn sm" data-act="draftedit" data-n="${n}" ${CH.thinking ? 'disabled' : ''} aria-label="${esc(t('edit'))}">${sic('edit', 15)}</button>
      ${n > 0 && !isLogD(CH.drafts[n - 1]) ? `<button type="button" class="iconbtn sm" data-act="draftmerge" data-n="${n}" ${CH.thinking ? 'disabled' : ''} title="${esc(t('mergePrev'))}" aria-label="${esc(t('mergePrev'))}">${sic('link', 15)}</button>` : ''}
      <button type="button" class="iconbtn sm" data-act="draftdrop" data-n="${n}" ${CH.thinking ? 'disabled' : ''} aria-label="${esc(t('remove'))}">${sic('x', 15)}</button></div>
    <div class="dchips">${chips}</div></div>`;
}
function questionHTML({ n, d, m }) {
  const lead = CH.drafts.length > 1 ? `<b class="qn">${t(CH.drafts.some(isLogD) ? 'itemN' : 'taskN', { n: n + 1 })}</b> ` : '';
  const ttl = `«${esc(d.title.length > 60 ? d.title.slice(0, 60) + '…' : d.title)}»`;
  let q = '', opts = '';
  if (m.k === 'match') { const mt = taskById(d.match); q = t('askMatch', { l: ttl, t: `«${esc((mt?.title || '').slice(0, 70))}»` }); opts = `<button type="button" class="chip" data-act="ansmatch" data-v="yes">${sic('check', 14)}${t('matchYes')}</button><button type="button" class="chip" data-act="ansmatch" data-v="no">${t('matchNo')}</button>`; }
  else if (m.k === 'due') { q = t('askDueQ', { t: ttl }); opts = dateChips().map(([l, v]) => `<button type="button" class="chip" data-act="ansdue" data-v="${v}">${l}</button>`).join('') + `<label class="chip datein">${sic('cal', 15)}<input type="date" id="ansDate" aria-label="${esc(t('pickDate'))}"></label><button type="button" class="chip" data-act="ansdue" data-v="">${t('noDue')}</button>`; }
  else if (m.k === 'weekend') { const nx = addWorkdays(d.due, 1), pv = addWorkdays(d.due, -1); q = t('askWeekend', { d: `${wd(d.due)} ${fmtShort(d.due)}` }); opts = (pv >= T() ? `<button type="button" class="chip" data-act="answk" data-v="${pv}">${t('moveTo', { d: `${wd(pv)} ${fmtShort(pv)}` })}</button>` : '') + `<button type="button" class="chip" data-act="answk" data-v="${nx}">${t('moveTo', { d: `${wd(nx)} ${fmtShort(nx)}` })}</button><button type="button" class="chip" data-act="answk" data-v="keep">${t('keepIt')}</button>`; }
  else if (m.k === 'source') { q = t('askSource', { t: ttl }); const srcs = [...new Set([...recentSources(), ...topPeople(4).map(p => p.name)])].slice(0, 6); const left = CH.drafts.filter(x => !isLogD(x) && !x.source).length; const all = left > 1 ? [...new Set([CH.lastSrc, SELF].filter(Boolean))].map(v => `<button type="button" class="chip primary-chip" data-act="anssrc" data-all="1" data-v="${esc(v)}" dir="auto">${esc(t('srcForAll', { s: srcLabel(v), n: left }))}</button>`).join('') : ''; opts = all + `<button type="button" class="chip on" data-act="anssrc" data-v="${SELF}">${t('srcSelf')}</button>` + srcs.map(s => `<button type="button" class="chip" data-act="anssrc" data-v="${esc(s)}" dir="auto">${esc(s)}</button>`).join(''); }
  else if (m.k === 'who') { q = t('askWho', { t: ttl }); opts = d.people.map((p, i) => p.role === 'related' ? `<button type="button" class="chip" data-act="answhorel" data-i="${i}" dir="auto">${esc(refName(p.ref))}</button>` : '').join('') + topPeople(8).filter(p => !d.people.some(q => q.ref.kind === 'known' && q.ref.id === p.id)).map(p => `<button type="button" class="chip ${CH.pick.has(p.id) ? 'on' : ''}" data-act="answho" data-v="${p.id}" dir="auto">${esc(p.name)}</button>`).join('') + `<button type="button" class="chip primary-chip" data-act="answhodone" ${CH.pick.size ? '' : 'disabled'}>${sic('check', 14)}${t('done')}</button>`; }
  else if (m.k === 'generic' || m.k === 'amb') { const p = d.people[m.i]; q = m.k === 'generic' ? t('askGeneric', { r: esc(p.ref.label) }) : t('whoExactly', { n: esc(p.ref.label) }); const ids = m.k === 'amb' ? p.ref.ids : topPeople(8).map(x => x.id); opts = ids.map(id => `<button type="button" class="chip" data-act="ansref" data-i="${m.i}" data-v="${id}" dir="auto">${esc(pname(id))}</button>`).join('') + (m.k === 'generic' ? `<button type="button" class="chip" data-act="ansref" data-i="${m.i}" data-v="keep">${t('keepAs', { r: esc(p.ref.label) })}</button>` : ''); }
  return `<div class="ask q"><p>${lead}${q}</p><div class="opts">${opts}</div>${m.k === 'match' ? '' : `<small class="note">${t(m.k === 'who' ? 'orTypeNames' : m.k === 'source' ? 'orTypeName' : m.k === 'due' ? 'orTypeDate' : 'orType')}</small>`}</div>`;
}
/** One line naming the optional fields still empty («ناقص كمان الوقت والمشروع، تحب تضيفهم؟»). */
function optLine() {
  const ks = [...new Set(CH.drafts.flatMap(optMissing))]; if (!ks.length) return '';
  const f = ks.map(k => t(OPT_LBL[k])).join(L() === 'ar' ? ' و' : ' and ');
  return `<p class="optline">${t(ks.length > 1 ? 'optMissingN' : 'optMissing1', { f })} <small>${t('optHint')}</small></p>`;
}
function extraHTML({ n, k }) {
  const d = CH.drafts[n], lead = CH.drafts.length > 1 ? `<b class="qn">${t('taskN', { n: n + 1 })}</b> ` : '';
  const ttl = `«${esc(d.title.length > 60 ? d.title.slice(0, 60) + '…' : d.title)}»`;
  const opts = k === 'time'
    ? ['09:00', '10:00', '12:00', '14:00'].map(v => `<button type="button" class="chip" data-act="ansextra" data-v="${v}"><span class="num">${fmtTime(v)}</span></button>`).join('') + `<label class="chip datein">${sic('clock', 15)}<input type="time" id="ansTime" aria-label="${esc(t('fTime'))}"></label>`
    : [...new Set(live().map(x => (x.project || '').trim()).filter(Boolean))].slice(0, 6).map(v => `<button type="button" class="chip" data-act="ansextra" data-v="${esc(v)}" dir="auto">${esc(v)}</button>`).join('');
  return `<div class="ask q"><p>${lead}${t(k === 'time' ? 'askTimeQ' : 'askProjectQ', { t: ttl })}</p><div class="opts">${opts}<button type="button" class="chip" data-act="ansextra" data-v="" data-skip="1">${t('cancel')}</button></div><small class="note">${t(k === 'time' ? 'orTypeTime' : 'orTypeProject')}</small></div>`;
}
/** Typed answer to a «+» question. Returns false when the text is not an answer. */
function answerExtra({ n, k }, text) {
  const d = CH.drafts[n]; if (!d) return false;
  if (k === 'time') {
    const tm = timeFrom(text) || (text.trim().split(/\s+/).length <= 3 ? timeFrom('الساعة ' + text) : null);
    if (tm) { d.time = tm.time; CH.extra = null; return true; }
    return text.split(/\s+/).length <= 3 ? (CH.msgs.push({ who: 'bot', html: t('timeNotUnderstood') }), true) : false;
  }
  if (text.split(/\s+/).length > 6) return false;
  d.project = text.replace(/\s+/g, ' ').trim().slice(0, 80); CH.extra = null; return true;
}

/* ---------- the chat runs the whole app: departments, people, statuses, screens ----------
   A command is understood on the device when it is plain («ضيف إدارة X وتحتها قسم Y»), otherwise by the assistant.
   Anything that changes data is shown as a card first and runs only when the user approves it. */
const ST_WORDS = [[/(قيد المراجعه|للمراجعه|تحت المراجعه|review)/, 'review'], [/(جاري|جاريه|قيد التنفيذ|شغال|بدات|بدأت|in progress)/, 'prog'], [/(متوقف|متوقفه|معلق|معلقه|اوقف|on hold)/, 'hold'], [/(لم تبدا|جديد|جديده|not started)/, 'todo'], [/(مكتمل|مكتمله|خلصت|خلصانه|منجز|منجزه|اتقفلت|done|complete)/, 'done'], [/(بانتظار|منتظر|مستني|waiting)/, 'wait']];
const stOfWords = n => { const h = ST_WORDS.find(([rx]) => rx.test(n)); return h ? h[1] : null; };
/** Open tasks that best match a description (shared significant words), best first. */
function findTasks(desc) {
  const sig = s2 => nameTokens(s2).map(w => w.length > 4 ? w.replace(/^ال/, '') : w).filter(w => w.length > 2 && !MATCH_STOP.has(w));
  const a = sig(desc); if (!a.length) return [];
  const ref = (String(desc).match(/\b\d{3,8}\b/) || [])[0];
  return live().filter(x => !isLog(x) && !x.archived).map(x => {
    if (ref && x.ext_ref === ref) return [x, 9];
    const b = sig(x.title); const hit = a.filter(w => b.some(v => v === w || (v.length > 3 && w.length > 3 && (v.includes(w) || w.includes(v))))).length;
    return [x, hit / Math.max(2, Math.min(a.length, b.length)) + (isOpen(x) ? 0.05 : 0)];
  }).filter(([, sc]) => sc >= 0.5).sort((p, q) => q[1] - p[1]).map(([x]) => x).slice(0, 4);
}
const AR_SPLIT = /\s*(?:،|,|؛|\s+و(?=\s*(?:قسم|ال|[^\s]))|\n)\s*/;
const cleanName = v => String(v || '').replace(/^[«"']|[»"'.؟?!]$/g, '').replace(/\s+/g, ' ').trim();
/** A plain command typed in the chat, or null. Returns a list of planned actions. */
function localCommand(text) {
  const raw = String(text || '').replace(/\s+/g, ' ').trim(); const n = normAr(raw);
  if (raw.length > 400) return null;
  let m;
  // screens
  if ((m = /^(افتح|اعرض|وريني|روح|ودني|خدني)\s+(?:لي\s+)?(?:صفحه\s+|شاشه\s+)?(الادارات|الاقسام|الادارات والاقسام|التقرير|الحصاد|تقرير الشهر|الجدول|جدول المهام|الاعدادات|الاشعارات|الاشخاص|المكتبه)\s*$/.exec(n))) {
    const k = /الادارات|الاقسام/.test(m[2]) ? 'org' : /تقرير|حصاد/.test(m[2]) ? 'harvest' : /جدول/.test(m[2]) ? 'table' : /اعدادات/.test(m[2]) ? 'more' : /اشعارات/.test(m[2]) ? 'notifs' : /اشخاص/.test(m[2]) ? 'people' : 'library';
    return [{ op: 'open', screen: k }];
  }
  if (/^(ارفع|استورد|حمل|ضيف|اضف)\s+(شيت|ملف|اكسل|excel)/.test(n)) return [{ op: 'open', screen: 'sheet' }];
  if (!orgOn()) return null;
  // «ضيف إدارة X (وتحتها|فيها) قسم Y، Z»
  if ((m = /^(?:ضيف|ضيفي|اضف|أضف|اضيف|انشئ|أنشئ|اعمل|سجل)\s+(?:لي\s+)?(?:اداره|إدارة|ادارة)\s+(.+?)(?:\s+(?:و\s*)?(?:تحتها|فيها|جواها|بها|وتحتها|وفيها|تتبعها|ويتبعها)\s+(?:الاقسام|الأقسام|اقسام|أقسام|قسم|قسمين)?\s*:?\s*(.+))?$/.exec(raw))) {
    const d = cleanName(m[1]).replace(/^(ال)?(اداره|إدارة)\s+/, ''); if (!d) return null;
    const dn = /^(إدارة|ادارة|اداره)\s/.test(d) ? d : 'إدارة ' + d;
    const secs = m[2] ? m[2].split(AR_SPLIT).map(x => cleanName(x)).filter(Boolean).map(x => /^قسم\s/.test(x) ? x : 'قسم ' + x) : [];
    return [{ op: 'add_unit', kind: 'dept', name: dn }, ...secs.map(sc => ({ op: 'add_unit', kind: 'sec', name: sc, parent: dn }))];
  }
  // «ضيف قسم X (في|لـ|تحت) إدارة Y»
  if ((m = /^(?:ضيف|اضف|أضف|اعمل|انشئ)\s+(?:اقسام|أقسام|قسم|قسمين)\s+(.+?)\s+(?:في|تحت|ل|لـ|على|علي|الى|إلى|جوه)\s*(?:ال)?(?:اداره|إدارة|ادارة)\s+(.+)$/.exec(raw))) {
    const pid = matchUnit('إدارة ' + cleanName(m[2])) || matchUnit(m[2]); const pn = pid ? deptOf(pid).name : 'إدارة ' + cleanName(m[2]).replace(/^(إدارة|ادارة|اداره)\s+/, '');
    return m[1].split(AR_SPLIT).map(x => cleanName(x)).filter(Boolean).map(x => ({ op: 'add_unit', kind: 'sec', name: /^قسم\s/.test(x) ? x : 'قسم ' + x, parent: pn }));
  }
  // «مأمون (تبع|في|من) قسم الجودة» / «اربط مأمون بإدارة X»
  if ((m = /^(?:اربط\s+)?(.+?)\s+(?:تبع|تابع ل|تابع|يتبع|بيتبع|في|من|ب|بـ)\s*(?:ال)?(اداره|إدارة|ادارة|قسم)\s+(.+)$/.exec(raw)) && m[1].split(' ').length <= 5) {
    const pid = matchPersonLoose(cleanName(m[1]).replace(/^(اربط|خلي)\s+/, '')); const uid = matchUnit(m[2] + ' ' + cleanName(m[3]));
    if (pid && uid) return [{ op: 'map_person', person: pid, unit: uid }];
  }
  // «خلي مهمة X قيد المراجعة» / «حوّل X لجارية»
  if ((m = /^(?:خلي|خليه|خليها|حول|حوّل|غير|غيّر|اجعل|حط)\s+(?:حاله\s+|حالة\s+)?(?:مهمه\s+|مهمة\s+|المهمه\s+|المهمة\s+)?(.+?)\s+(?:الى\s+|إلى\s+|ل|لـ|تبقى\s+|تكون\s+|على\s+)?(قيد المراجعة|قيد المراجعه|للمراجعة|جارية|جاريه|جاري|قيد التنفيذ|متوقفة|متوقفه|معلقة|معلقه|لم تبدأ|لم تبدا|جديدة|جديده|مكتملة|مكتمله|منجزة|منجزه)\s*$/.exec(raw))) {
    const st = stOfWords(normAr(m[2])); const cands = findTasks(m[1]);
    if (st && cands.length) return [{ op: 'set_status', status: st, cands: cands.map(x => x.id), task: cands.length === 1 || /\d{3,8}/.test(m[1]) ? cands[0].id : null }];
  }
  return null;
}
/** The assistant's actions → the same plan (names resolved on the device; nothing of the stored tasks is sent to it). */
function planActions(list) {
  const out = [];
  (list || []).slice(0, 30).forEach(a => {
    if (!a || !a.op) return;
    if (a.op === 'open' && a.screen) return out.push({ op: 'open', screen: a.screen });
    if (!orgOn() && a.op !== 'set_status') return;
    if (a.op === 'add_unit' && a.name) return out.push({ op: 'add_unit', kind: a.parent ? 'sec' : 'dept', name: cleanName(a.name).slice(0, 150), parent: a.parent ? cleanName(a.parent).slice(0, 150) : null });
    if (a.op === 'map_person' && a.person && a.unit) { const pid = matchPersonLoose(a.person); const uid = matchUnit(a.unit) || null; if (pid) out.push({ op: 'map_person', person: pid, unit: uid, unitName: uid ? null : cleanName(a.unit) }); return; }
    if ((a.op === 'set_status' || a.op === 'set_unit') && a.task) { const cands = findTasks(a.task); if (!cands.length) return; const st = a.op === 'set_status' ? (['todo', 'prog', 'wait', 'review', 'hold', 'done'].includes(a.status) ? a.status : stOfWords(normAr(a.status || ''))) : null; if (a.op === 'set_status' && !st) return; const uid = a.op === 'set_unit' ? matchUnit(a.unit || '') : null; if (a.op === 'set_unit' && !uid) return; out.push({ op: a.op, status: st, unit: uid, cands: cands.map(x => x.id), task: cands.length === 1 ? cands[0].id : null }); }
  });
  return out;
}
const actReady = () => (CH.actions || []).every(a => !a.cands || a.task);
function actLine(a, i) {
  if (a.op === 'add_unit') { const exists = a.kind === 'dept' ? depts().some(d => normAr(d.name) === normAr(a.name)) : false; return `${sic('org', 15)} <span dir="auto">${a.kind === 'dept' ? t('actAddDept', { n: esc(a.name) }) : t('actAddSec', { n: esc(a.name), d: esc(a.parent || '') })}</span>${exists ? ` <em class="muted">${t('actExists')}</em>` : ''}`; }
  if (a.op === 'map_person') return `${sic('users', 15)} <span dir="auto">${t('actMap', { p: esc(pname(a.person)), u: esc(a.unit ? unitLabel(a.unit) : a.unitName || '') })}</span>`;
  if (a.op === 'set_status' || a.op === 'set_unit') {
    const what = a.op === 'set_status' ? `<span class="pill ${{ prog: 'prog', wait: 'wait', review: 'review', done: 'done' }[a.status] || ''}">${t(stKey(a.status))}</span>` : `<b dir="auto">${esc(unitLabel(a.unit))}</b>`;
    if (a.task) return `${sic('check', 15)} <span dir="auto">${t(a.op === 'set_status' ? 'actStatus' : 'actUnit', { t: esc(taskById(a.task)?.title.slice(0, 70) || '') })}</span> ${what}`;
    return `${sic('check', 15)} <span>${t('actWhich')}</span> ${what}<div class="opts">${a.cands.map(id => `<button type="button" class="chip" data-act="actpick" data-i="${i}" data-v="${id}" dir="auto">${esc(taskById(id)?.title.slice(0, 60) || '')}</button>`).join('')}</div>`;
  }
  return '';
}
function actionsHTML() {
  const A = CH.actions || []; if (!A.length) return '';
  return `<div class="bubble bot drafts actcard"><div class="dh">${t('actTitle')}</div><ul class="actlist">${A.map((a, i) => `<li>${actLine(a, i)}</li>`).join('')}</ul>
    <div class="save-row"><button type="button" class="btn primary block" data-act="actrun" ${actReady() ? '' : 'disabled'}>${ic.check}${t('actRun')}</button><button type="button" class="btn" data-act="actdrop">${t('cancel')}</button></div></div>`;
}
function runActions() {
  const A = (CH.actions || []).filter(a => !a.cands || a.task); CH.actions = null;
  const done = [];
  A.forEach(a => {
    if (a.op === 'add_unit') {
      if (a.kind === 'dept') { const before = depts().length; ensureDept(a.name); if (depts().length > before) done.push(t('actDoneDept', { n: a.name })); }
      else { const pid = ensureDept(a.parent || ''); if (pid && addSecs(pid, [a.name])) done.push(t('actDoneSec', { n: a.name })); }
    } else if (a.op === 'map_person') { const p = personById(a.person); const uid = a.unit || (a.unitName ? ensureDept(a.unitName) : null); if (p && uid) { p.unit_id = uid; putPerson(p); done.push(t('actDoneMap', { p: p.name, u: unitLabel(uid) })); } }
    else if (a.op === 'set_status') {
      const x = taskById(a.task); if (!x || x.status === a.status) return;
      if (a.status === 'done') { x.status = 'done'; x.completed_on = T(); addLog(x, 'completed'); spawnNext(x); putTask(x); }
      else if (a.status === 'wait' && !x.waiting_on) { x.status = 'prog'; putTask(x); }
      else setStatus(x, a.status);
      done.push(t('actDoneStatus', { t: x.title.slice(0, 50), s: t(stKey(a.status)) }));
    } else if (a.op === 'set_unit') { const x = taskById(a.task); if (x) { x.unit_id = a.unit; addLog(x, 'edited'); putTask(x); done.push(t('actDoneUnit', { t: x.title.slice(0, 50), u: unitLabel(a.unit) })); } }
  });
  botSay(done.length ? t('actDone', { n: done.length }) + '\n' + done.join('\n') : t('actNothing'));
  renderChat();
}
function openScreen(k) {
  if (k === 'sheet') return pickSheet();
  closeSheet(true);
  if (k === 'org') return openOrg(); if (k === 'harvest') return openHarvest(); if (k === 'notifs') return openNotifSheet();
  if (k === 'table') { UI.tview = 'table'; return go('tasks'); }
  return go(k === 'more' ? 'more' : k);
}

/* ---------- sending ---------- */
const QUERY_RE = /(^|\s)(هات|اعرض|وريني|ورّيني|ايه|إيه|ما هي|ماهي|مين|من عنده|كام|عندي ايه|عندي إيه|اللي عليا|اللي علي|ملخص|show|list|what|who|which|summary)(\s|$|\?|؟)/;
const CONFIRM_RE = /^(تمام|ماشي|اوك|اوكي|ok|okay|حفظ|احفظ|احفظهم|احفظها|احفظ الكل|ضيف|ضيفهم|ضيفها|ضيفه|اضف|اضفهم|سجل|سجلهم|سجلها|اعتمد|اعتمدهم|موافق|تم|يلا|كده تمام|تمام كده|تمام ضيفهم|تمام احفظ|save|yes|done)(\s|$)/;
const CANCEL_RE = /^(الغي|الغاء|الغيهم|امسحهم|امسح الكل|انسي|انسى|سيبك|بلاش|cancel|discard)(\s|$)/;
const shortN = text => normAr(text).replace(/[!.،,؟?]/g, ' ').replace(/\s+/g, ' ').trim();
const isConfirm = text => { const n = shortN(text); return n.split(' ').length <= 4 && CONFIRM_RE.test(n) && !/(^|\s)(لا|مش|غير|عدل|خلي|بس)(\s|$)/.test(n); };
const isCancel = text => { const n = shortN(text); return n.split(' ').length <= 4 && CANCEL_RE.test(n); };
function hist(role, text) { if (!text) return; CH.hist.push({ role, text: String(text).slice(0, 600) }); if (CH.hist.length > 12) CH.hist.splice(0, CH.hist.length - 12); }
const draftBrief = () => CH.drafts.map((d, i) => isLogD(d)
  ? { n: i + 1, kind: 'done', title: d.title, date: d.date, project: d.project || null }
  : { n: i + 1, kind: 'task', title: d.title, due: d.due, no_due: !!d.noDue, time: d.time || null, recur: d.recur || null, priority: d.priority, role: d.role, source: d.source === SELF ? 'self' : d.source || null, project: d.project || null, unit: unitById(d.unit) ? unitLabel(d.unit) : null, people: d.people.map(p => ({ name: refName(p.ref), role: p.role })) });
const botSay = (text, extra = '') => { CH.msgs.push({ who: 'bot', html: `<span dir="auto">${esc(text)}</span>${extra}` }); hist('model', text); };
/** Apply one AI edit to a draft. */
function applyEdit(d, e) {
  if (!d || !e) return;
  if (e.title) d.title = String(e.title).replace(/\s+/g, ' ').trim().slice(0, 300);
  if (isLogD(d)) {
    if (validDate(e.date) && e.date <= T()) d.date = e.date;
    if (e.project != null) d.project = String(e.project || '').slice(0, 80);
    if (e.title) { d.match = matchOpenTask(d.title); d.closeTask = undefined; }
    return;
  }
  if (e.no_due === true) { d.due = null; d.noDue = true; }
  if (validDate(e.due)) { d.due = e.due; d.noDue = false; d.weekendOk = false; }
  if (e.clear_time) d.time = null; else if (/^\d{2}:\d{2}$/.test(e.due_time || '')) d.time = e.due_time;
  if (['hi', 'mid', 'lo'].includes(e.priority)) { d.priority = e.priority; d.prioAsked = true; }
  if (['exec', 'follow', 'both'].includes(e.role)) d.role = e.role;
  if (e.source_self) d.source = SELF; else if (e.source) d.source = srcName(e.source);
  if (e.project != null) d.project = String(e.project || '').slice(0, 80);
  if (e.unit != null && orgOn()) { const u = e.unit ? matchUnit(e.unit) : null; if (u || !e.unit) { d.unit = u; d.unitAsked = true; } }
  if (e.clear_recur) d.recur = null;
  else if (e.recur) { d.recur = cleanRecur(e.recur); if (d.recur && !validDate(e.due)) { d.due = recurNext(d.recur, addDays(T(), -1)); d.noDue = false; } }
  (e.remove_people || []).forEach(nm => { const k = normAr(nm); if (k) d.people = d.people.filter(p => { const r = normAr(refName(p.ref)); return !(r.includes(k) || k.includes(r)); }); });
  (e.add_people || []).forEach(p => { if (!p?.name) return; const ref = refOf(p.name, p.desc, !!p.generic); if (ref && !d.people.some(q => refName(q.ref) === refName(ref))) d.people.push({ ref, role: ['followup', 'waiting', 'related'].includes(p.role) ? p.role : 'followup', what: String(p.what || '').slice(0, 200), due: validDate(p.due) || null }); });
  if (d.people.some(p => p.role !== 'related') && d.role === 'exec') d.role = 'both';
  unitFromPeople(d);
}
/** Drafts understood on-device only (AI off, offline or failed). */
function localDrafts(text) {
  const pieces = splitTasks(text);
  const logItems = isLogText(text) ? splitLogs(text) : [];
  return logItems.length ? logItems.map(it => fromLog(it.title, logDate(text, T()), text)) : (pieces.length > 1 ? pieces : [text]).map(p => fromLocal(parseCapture(p, people(), T()), p));
}
const GREET_RE = /^(ازيك|إزيك|ازيكم|عامل ايه|السلام عليكم|سلام عليكم|مرحبا|اهلا|أهلا|هاي|هلا|صباح الخير|مساء الخير|hi|hello|hey)(\s|$|[!.،,؟?])/;
async function chatSend(text, label) {
  text = String(text || '').trim(); if (!text || CH.thinking) return;
  if (MIC) stopMic();
  if (!$('#chatLog')) openChat();
  CH.msgs.push({ who: 'me', html: `<span dir="auto">${esc(label || text)}</span>` }); hist('user', label || text);
  const ta = $('#chatIn'); if (ta) { ta.value = ''; autoGrow(ta); }
  const pend = pending();
  // a pending command card: approve / cancel by word
  if (CH.actions?.length) { if (isConfirm(text) && actReady()) return runActions(); if (isCancel(text)) { CH.actions = null; botSay(t('actCancelled')); renderChat(); return; } }
  // plain commands are understood on the device
  const lc = !pend && localCommand(text);
  if (lc && lc.length) { if (lc.length === 1 && lc[0].op === 'open') { botSay(t('actOpening')); renderChat(); return openScreen(lc[0].screen); } CH.actions = lc; botSay(t('actConfirmQ')); renderChat(); return; }
  // quick, deterministic paths
  if (CH.drafts.length && !pend && isConfirm(text)) { saveDrafts(); return; }
  if (CH.drafts.length && isCancel(text)) { CH.drafts = []; CH.pick = new Set(); botSay(t('draftsDiscarded')); renderChat(); return; }
  if (pend && pend.m.k !== 'match' && isConfirm(text)) { botSay(t('stillMissing')); renderChat(); return; }
  if (pend && !QUERY_RE.test(' ' + normAr(text) + ' ') && answerText(pend, text)) { renderChat(); return; }
  const sp = !pend && softPending();
  if (sp && !isConfirm(text) && !QUERY_RE.test(' ' + normAr(text) + ' ') && answerSoft(sp, text)) { renderChat(); return; }
  const ex = !pend && extraOn();
  if (ex && !isConfirm(text) && !QUERY_RE.test(' ' + normAr(text) + ' ') && answerExtra(ex, text)) { renderChat(); return; }
  const looksQuery = /[?؟]\s*$/.test(text) || /^(هات|اعرض|وريني|ورّيني|ايه|إيه|ما هي|ماهي|مين|من عنده|كام|عندي ايه|عندي إيه|ملخص|show|list|what|who|which|summary)(\s|$)/.test(shortN(text)) || shortN(text).split(' ').length <= 5;
  const lq = (looksQuery || !(aiOn() && navigator.onLine)) ? localQuery(text) : null; // «اعرض التقرير على المدير بكرة» is a task, not a question
  if (lq) { CH.msgs.push({ who: 'bot', html: queryHTML(lq) }); hist('model', t('answeredQuery')); renderChat(); return; }
  const fallback = note => {
    if (!aiOn() && GREET_RE.test(shortN(text))) { botSay(t('greetLocal')); return; }
    const locals = localDrafts(text); CH.drafts.push(...locals);
    if (note) CH.msgs.push({ who: 'bot', html: `<span class="muted">${note}</span>` });
    hist('model', t('nTasksFound', { n: plural(locals.length, 'tasks') }));
  };
  if (!(aiOn() && navigator.onLine)) { fallback(''); CH.pick = new Set(); renderChat(); return; }
  // show what the device understood right away; the assistant's answer replaces it a few seconds later
  let prov = [];
  if (!CH.drafts.length && !GREET_RE.test(shortN(text)) && text.split(/\s+/).length >= 3) { prov = localDrafts(text); CH.drafts.push(...prov); }
  CH.thinking = true; CH.since = Date.now(); renderChat();
  const dropProv = () => { if (prov.length) { CH.drafts = CH.drafts.filter(d => !prov.includes(d)); prov = []; } };
  try {
    const r = await invokeAssist({ v: 2, text: text.slice(0, 4000), history: CH.hist.slice(0, -1), drafts: prov.length ? [] : draftBrief(), pending: pend ? pend.m.k : sp ? 'optional ' + sp.k : null, today: T(), weekday: new Intl.DateTimeFormat('en', { weekday: 'long' }).format(pd(T())), lang: L(), weekend: weekendDays(), people: people().map(p => p.name), sources: recentSources(), units: orgOn() ? units().map(u => unitLabel(u.id)) : [] }, text.length > 450 ? 34000 : 22000);
    dropProv();
    // 1) changes to the drafts already on screen
    const n0 = CH.drafts.length;
    const asNew = (r.edits || []).filter(e => (e.n | 0) > n0 && e.title); // a model slip: new items returned as edits of drafts that do not exist
    (r.edits || []).forEach(e => { if ((e.n | 0) <= n0) applyEdit(CH.drafts[(e.n | 0) - 1], e); });
    [...new Set((r.remove || []).map(n => (n | 0) - 1))].filter(i => i >= 0 && i < n0).sort((a, b) => b - a).forEach(i => CH.drafts.splice(i, 1));
    // 2) new items
    const tk = [...(r.tasks || []), ...asNew.map(e => ({ ...e, people: e.add_people || [] }))].filter(a => a && a.title).slice(0, 20), lg = (r.logs || []).filter(a => a && a.title).slice(0, Math.max(0, 20 - tk.length));
    if (tk.length || lg.length) {
      const tl = tk.length === 1 ? localDrafts(text).filter(d => !isLogD(d)) : [];
      CH.drafts.push(...lg.map(a => fromLog(a.title, a.date, text, a.project)), ...tk.map(a => fromAI(a, tk.length === 1 && tl.length === 1 ? tl[0] : null, text, tk.length > 1)));
    }
    // 3) what to say / do
    const fresh = tk.length || lg.length; // items created by this very message have not been reviewed yet
    if (r.intent === 'save' && n0 && !fresh) { if (!pending()) { CH.thinking = false; saveDrafts(); return; } botSay(t('stillMissing')); }
    else if (r.intent === 'save') botSay(fresh ? t('reviewFirst') : (r.reply || t('nothingToSave')));
    else if (r.intent === 'cancel') { CH.drafts = []; botSay(r.reply || t('draftsDiscarded')); }
    else if (r.intent === 'action' && (r.actions || []).length) { const pl = planActions(r.actions); const op = pl.length === 1 && pl[0].op === 'open' ? pl[0] : null; if (op) { if (r.reply) botSay(r.reply); renderChat(); CH.thinking = false; return openScreen(op.screen); } if (pl.length) { CH.actions = pl; botSay(r.reply || t('actConfirmQ')); } else botSay(t('actNotUnderstood')); }
    else if (r.intent === 'query' && r.query) { if (r.reply) botSay(r.reply); CH.msgs.push({ who: 'bot', html: queryHTML(normQuery(r.query)) }); }
    else if (r.reply) botSay(r.reply);
  } catch (e) {
    console.warn('assist', e);
    const note = e?.code === 'quota' ? t('aiQuota') : e?.code === 'missing_key' ? t('aiNotReady') : e?.code === 'busy' ? t('aiBusy') : e?.code === 'timeout' ? t('aiSlow') : '';
    if (e?.code === 'quota' || e?.code === 'missing_key') UI.aiBlocked = e.code;
    if (prov.length) { prov = []; CH.msgs.push({ who: 'bot', html: `<span class="muted">${note ? note + ' · ' : ''}${t('localUnderstanding')}</span>` }); }
    else fallback(`${note ? note + ' · ' : ''}${t('localUnderstanding')}`);
  } finally { CH.thinking = false; }
  CH.pick = new Set();
  renderChat();
}
/** Free-text answer to the current question. Returns false when the text is not an answer. */
const EDIT_CUE = /(^|\s)(خلي|خليها|خليه|غير|غيّر|عدل|عدّل|شيل|ضيف|الاولي|الأولى|التانيه|الثانيه|التالته|الثالثه|الاخيره|كمان|بدل|ومصدرها|والمصدر|change|make it|instead)(\s|$)/;
function answerText({ d, m }, text) {
  const n = normAr(text);
  if (m.k !== 'who' && (text.trim().split(/\s+/).length > 5 || EDIT_CUE.test(n))) return false; // an edit, not an answer: let the assistant read it
  if (m.k === 'match') {
    if (/^(نعم|ايوه|ايوا|اه|ايه|صح|هي|هيا|بالظبط|yes|yeah|y)\b/.test(n.trim())) { d.closeTask = d.match; return true; }
    if (/^(لا|لاء|لأ|no|منفصل|مختلف)/.test(n.trim())) { d.closeTask = null; return true; }
    return false;
  }
  if (m.k === 'due' || m.k === 'weekend') {
    if (/(دون|بدون|من غير|مفيش|ملوش|no)\s*(موعد|ميعاد|تاريخ|due|deadline)/.test(n)) { d.due = null; d.noDue = true; return true; }
    const p = parseCapture(text, [], T()); if (p.due) { d.due = p.due; d.noDue = false; d.weekendOk = m.k === 'weekend' || !isWeekend(p.due) ? d.weekendOk : false; if (p.time) d.time = p.time; return true; }
    return text.split(/\s+/).length <= 4 ? (CH.msgs.push({ who: 'bot', html: t('dateNotUnderstood') }), true) : false;
  }
  if (m.k === 'source') {
    if (/^(انا|أنا|نفسي|ذاتي|me|myself|self)$/.test(n.trim()) || /(من نفسي|مبادره)/.test(n)) { d.source = SELF; CH.lastSrc = SELF; return true; }
    if (text.split(/\s+/).length > 4) return false;
    d.source = srcName(text); CH.lastSrc = d.source; return true;
  }
  if (m.k === 'who') {
    if (text.split(/\s+/).length > 8 || EDIT_CUE.test(n)) return false; // «خلي السابعة تنفيذ شخصي» is an edit, not a name
    if (text.split(/[،,+؛;]|\s+و\s*/).some(x => x.trim().split(/\s+/).length > 3)) return false;
    const names = text.split(/[،,+؛;]|\s+و\s*/).map(s => s.trim()).filter(Boolean);
    names.forEach(nm => { const ref = refOf(nm); if (!ref) return; const ex = d.people.find(p => refName(p.ref) === refName(ref)); if (ex) { if (ex.role === 'related') ex.role = 'followup'; } else d.people.push({ ref, role: 'followup', what: '', due: null }); }); // a name already on the task (as «related») becomes the one followed up
    return names.length > 0;
  }
  if (m.k === 'generic' || m.k === 'amb') {
    if (text.split(/\s+/).length > 4) return false;
    const p = d.people[m.i]; const label = p.ref.label; const ref = refOf(text, m.k === 'generic' ? label : '');
    if (ref) p.ref = ref; return !!ref;
  }
  return false;
}
function saveDrafts() {
  if (pending() || !CH.drafts.length) return;
  const made = [], closed = [];
  for (const d of CH.drafts) {
    if (isLogD(d)) {
      const y = d.closeTask ? taskById(d.closeTask) : null;
      if (y && isOpen(y)) { y.status = 'done'; y.completed_on = d.date; addLog(y, 'completed', normAr(d.title) !== normAr(y.title) ? d.title : ''); spawnNext(y); putTask(y); made.push(y); if (openFus(y).length) closed.push(y); continue; }
      const lx = newTask({ kind: 'log', title: d.title.slice(0, 300) || t('untitled'), status: 'done', completed_on: d.date, due: null, project: d.project || '', source: '' });
      addLog(lx, 'logged'); putTask(lx); made.push(lx); continue;
    }
    const ex = d.extra || {};
    const x = newTask({ title: d.title.slice(0, 300) || t('untitled'), due: d.noDue && !d.recur ? null : d.due, due_time: d.time || null, recur: d.recur || null, priority: d.priority, role: d.role, source: d.source, project: d.project || '', details: ex.details || '', notes: ex.notes || '', steps_total: ex.steps_total || 0, remind_min: ex.remind_min ?? null });
    if (unitById(d.unit)) x.unit_id = d.unit;
    if (ex.status && ex.status !== 'wait') x.status = ex.status;
    if (ex.status === 'wait' && ex.waiting_on) { const wid = getOrCreatePerson(ex.waiting_on); if (wid) { x.status = 'wait'; x.waiting_on = wid; x.waiting_what = ex.waiting_what || 'reply'; } }
    addLog(x, 'created');
    const idOf = r => { if (r.kind === 'known') return r.id; const id = getOrCreatePerson(r.kind === 'new' ? r.name : r.label); const p = personById(id); if (p && r.desc && !p.org) { p.org = r.desc; putPerson(p); } return id; };
    for (const p of d.people) {
      if (p.role === 'related' && d.role === 'exec') continue;
      const pid = idOf(p.ref); if (!pid) continue;
      if (p.role === 'waiting' && x.status !== 'wait') { x.status = 'wait'; x.waiting_on = pid; x.waiting_what = d.waitWhat || 'reply'; continue; }
      if (x.followups.some(f => f.person_id === pid)) continue;
      const due = validDate(p.due) || (x.due && x.due > T() ? (diffDays(x.due, T()) > 1 ? addWorkdays(x.due, -1) : x.due) : addWorkdays(T(), 1));
      x.followups.push({ id: uuid(), person_id: pid, what: p.what || '', due, status: 'open', created_at: nowISO(), closed_on: null, log: [] });
      addLog(x, 'fu_added', '', { person_id: pid });
    }
    putTask(x); made.push(x);
  }
  CH.drafts = []; CH.pick = new Set(); CH.lastSrc = null;
  const nDone = made.filter(x => x.status === 'done').length;
  const head = !nDone ? (made.length > 1 ? t('savedN', { n: plural(made.length, 'tasks') }) : t('savedTask')) : nDone === made.length ? (made.length > 1 ? t('savedLogs', { n: plural(made.length, 'logs') }) : t('savedLog')) : t('savedMixed', { n: plural(made.length, 'items') });
  hist('model', head + ' ' + made.map(x => x.title).join(' / '));
  const canAttach = state.mode === 'cloud';
  CH.msgs.push({ who: 'bot', html: `${sic('check', 16)} ${head}<div class="qlist">${made.map(x => `<div class="qrow"><span class="body" data-open="${x.id}" role="button" tabindex="0"><span class="t" dir="auto">${esc(x.title)}</span><span class="meta">${x.status === 'done' ? `${statusPill(x)}<span>${x.completed_on === T() ? t('today') : fmtShort(x.completed_on)}</span>` : `<span>${x.due ? rel(x.due) : t('noDue')}${x.due_time ? ' · ' + fmtTime(x.due_time) : ''}</span>${x.recur ? `<span>${sic('repeat', 12)}</span>` : ''}`}</span></span>${canAttach ? `<button type="button" class="iconbtn sm" data-act="upload" data-task="${x.id}" aria-label="${esc(t('attach'))}" title="${esc(t('attach'))}">${sic('clip', 16)}</button>` : ''}</div>`).join('')}</div>${closed.map(y => `<div class="ask"><p>${t('closeFusQ', { n: plural(openFus(y).length, 'openFu') })} <span class="muted" dir="auto">«${esc(y.title.slice(0, 50))}»</span></p><button type="button" class="chip" data-act="closefus" data-id="${y.id}">${sic('check', 14)}${t('closeFus')}</button></div>`).join('')}` });
  renderChat();
  if (UI.tab !== 'today' && UI.tab !== 'tasks') { UI.tab = 'today'; renderNav(); renderMain(); }
}

/* ---------- questions about my work (answered on-device) ---------- */
function weekRange(td = T()) { const back = (pd(td).getDay() - weekStartDay() + 7) % 7; const s = addDays(td, -back); return [s, addDays(s, 6)]; }
function localQuery(text) {
  const n = ' ' + normAr(text) + ' ';
  if (orgOn() && depts().length && /(^|\s)(مهام|المهام|شغل|تاسكات|وضع|حاله|tasks|status)(\s|$)/.test(n) && /(اداره|قسم|قطاع|department|section)/.test(n) && n.trim().split(' ').length <= 8) { const u = matchUnit(text); if (u) return { type: 'unit', unit: u }; }
  if (!QUERY_RE.test(n) && !/(مهامي|مهام اليوم|مهام بكره|متاخرات|متابعاتي|my tasks)/.test(n)) return null;
  const td = T();
  if (/(متاخر|(^|\s)(overdue|late)(\s|$))/.test(n)) return { type: 'late' };
  if (/(واقف|متعطل|معلق|مين عنده|من عنده|who owes|pending follow)/.test(n)) return { type: 'followups' };
  if (/(مستني|منتظر|بانتظار|waiting)/.test(n)) return { type: 'waiting' };
  if (/(من غير موعد|بدون موعد|no due)/.test(n)) return { type: 'nodue' };
  if (/(خلصت|انجزت|منجز|اكتمل|خلصته|عملت|عملته|انجازات|(^|\s)(خلص|خلصو|خلصوا|اتعمل|اتنجز|انتهى|انتهت|done|completed|finished|did)(\s|$))/.test(n)) { if (/(النهارده|اليوم|today)/.test(n)) return { type: 'done', from: td, to: td }; if (/(امبارح|امس|yesterday)/.test(n)) return { type: 'done', from: addDays(td, -1), to: addDays(td, -1) }; if (/(الشهر|month)/.test(n)) return { type: 'done', from: td.slice(0, 8) + '01', to: td }; const [s] = weekRange(td); return { type: 'done', from: s, to: td }; }
  if (orgOn() && /(اداره|قسم|قطاع|department|section)/.test(n)) { const u = matchUnit(text); if (u) return { type: 'unit', unit: u }; }
  const hit = findPeople(text, people()); if (hit.length && hit[0].ids.length === 1) return { type: 'person', person: hit[0].ids[0] };
  const pm = text.match(/مشروع\s+(.+?)(?:\s*[?؟]|$)/); if (pm) return { type: 'project', project: pm[1].trim() };
  if (/(الاسبوع الجاي|الاسبوع القادم|next week)/.test(n)) { const s = nextWeekday(td, weekStartDay()); return { type: 'due', from: s, to: addDays(s, 6) }; }
  if (/(الاسبوع|this week)/.test(n)) { return { type: 'due', from: td, to: weekRange(td)[1] }; }
  const aboutWork = /(مهام|مهمه|متابعات|عندي|ايه|what|tasks|follow|schedule|جدول|اجنده|ملخص|summary)/.test(n); // «اعرض التقرير على المدير بكرة» is work to do, not a question
  const p = parseCapture(text, [], td); if (p.due) return aboutWork ? { type: 'due', from: p.due, to: p.due } : null;
  if (/(النهارده|اليوم|today|ملخص|summary)/.test(n)) return aboutWork ? { type: 'due', from: td, to: td } : null;
  return /(مهام|متابعات|tasks|follow)/.test(n) ? { type: 'summary' } : null;
}
function normQuery(q) {
  const out = { type: q.type || 'summary', from: validDate(q.from), to: validDate(q.to), project: q.project || '' };
  if (q.person) out.person = matchPerson(q.person) || null;
  if (q.unit) out.unit = unitById(q.unit) ? q.unit : orgOn() ? matchUnit(q.unit) : null;
  if (out.type === 'unit' && !out.unit) out.type = 'summary';
  if (out.type === 'due' && !out.from) out.from = out.to = T();
  if (out.type === 'due' && !out.to) out.to = out.from;
  if (out.type === 'done') { if (!out.from) out.from = weekRange()[0]; if (!out.to) out.to = T(); }
  if (out.type === 'person' && !out.person) out.type = 'summary';
  return out;
}
const qRow = x => `<button type="button" class="qrow ${isLate(x) ? 'late' : ''}" data-open="${x.id}"><span class="t" dir="auto">${esc(x.title)}</span><span class="meta">${statusPill(x)}${x.status === 'done' ? (x.completed_on ? `<span>${x.completed_on === T() ? t('today') : x.completed_on === addDays(T(), -1) ? t('yesterday') : `${wd(x.completed_on)} ${fmtShort(x.completed_on)}`}</span>` : '') : x.due ? `<span>${isLate(x) ? lateTxt(x.due) : rel(x.due)}</span>` : ''}${x.waiting_on ? `<span dir="auto">${sic('clock', 12)} ${esc(pname(x.waiting_on))}</span>` : ''}</span></button>`;
const qFu = (x, f) => `<div class="qrow fu ${fuLate(f) ? 'late' : ''}"><span class="av sm">${esc(initials(pname(f.person_id)))}</span><span class="body" data-open="${x.id}" role="button" tabindex="0"><span class="t" dir="auto">${esc(pname(f.person_id))}</span><span class="meta"><span dir="auto">${esc(f.what || x.title)}</span><span>${fuLate(f) ? lateTxt(f.due) : rel(f.due)}</span></span></span><button type="button" class="btn sm" data-act="furesult" data-id="${x.id}" data-fu="${esc(f.id)}">${t('logResult')}</button></div>`;
function queryHTML(q) {
  const open = live().filter(x => isOpen(x) && !x.archived);
  const fus = []; open.forEach(x => openFus(x).forEach(f => fus.push([x, f])));
  const list = (title, rows, empty) => `<div class="qhead">${title}</div>${rows.length ? `<div class="qlist">${rows.join('')}</div>` : `<div class="muted">${empty}</div>`}`;
  const rangeLabel = (a, b) => a === b ? (a === T() ? t('today') : a === addDays(T(), 1) ? t('tomorrow') : `${wd(a)} ${fmtShort(a)}`) : `${fmtShort(a)} – ${fmtShort(b)}`;
  switch (q.type) {
    case 'late': { const ts = sortTasks(open.filter(isLate)); const fs = fus.filter(([, f]) => fuLate(f)); return list(ts.length && fs.length ? t('qaLate', { n: plural(ts.length, 'tasks'), m: plural(fs.length, 'fus') }) : t('qaLate1', { n: plural(ts.length || fs.length, ts.length ? 'tasks' : 'fus') }), [...ts.map(qRow), ...fs.map(([x, f]) => qFu(x, f))], t('qaNothingLate')); }
    case 'waiting': { const ts = sortTasks(open.filter(isWaitLike)); return list(t('qaWaiting', { n: ts.length }), ts.map(qRow), t('qaNone')); }
    case 'nodue': { const ts = sortTasks(open.filter(x => !x.due)); return list(t('qaNoDue', { n: ts.length }), ts.map(qRow), t('qaNone')); }
    case 'followups': {
      const by = new Map(); fus.forEach(([x, f]) => { if (!by.has(f.person_id)) by.set(f.person_id, []); by.get(f.person_id).push([x, f]); });
      open.filter(x => x.status === 'wait' && x.waiting_on).forEach(x => { if (!by.has(x.waiting_on)) by.set(x.waiting_on, []); by.get(x.waiting_on).push([x, null]); });
      const groups = [...by].sort((a, b) => b[1].length - a[1].length);
      return `<div class="qhead">${t('qaFollow', { n: plural(groups.length, 'people') })}</div>${groups.length ? `<div class="qlist">${groups.map(([pid, items]) => { const late = items.filter(([, f]) => f && fuLate(f)).length; return `<button type="button" class="qrow" data-person="${pid}"><span class="av sm">${esc(initials(pname(pid)))}</span><span class="body"><span class="t" dir="auto">${esc(pname(pid))}</span><span class="meta"><span>${plural(items.length, 'openFu')}</span>${late ? `<span class="pill late">${t('nLate', { n: late })}</span>` : ''}</span><span class="sub" dir="auto">${esc(items.slice(0, 2).map(([x, f]) => f?.what || x.title).join(' · '))}</span></span></button>`; }).join('')}</div>` : `<div class="muted">${t('qaNone')}</div>`}`;
    }
    case 'person': { const pid = q.person; const ts = sortTasks(open.filter(x => x.waiting_on === pid)); const fs = fus.filter(([, f]) => f.person_id === pid); return list(t('qaPerson', { p: esc(pname(pid)), n: plural(ts.length + fs.length, 'items') }), [...ts.map(qRow), ...fs.map(([x, f]) => qFu(x, f))], t('qaNone')) + `<button type="button" class="linkbtn" data-person="${pid}">${t('openPerson')}</button>`; }
    case 'unit': { const ts = sortTasks(open.filter(x => inUnit(x, q.unit))); const late = ts.filter(isLate).length; const byS = {}; ts.forEach(x => { const k = x.status; byS[k] = (byS[k] || 0) + 1; });
      return list(t('qaUnit', { u: esc(unitLabel(q.unit)), n: plural(ts.length, 'tasks') }) + (ts.length ? `<div class="qsub">${Object.entries(byS).map(([k, n]) => `${t(stKey(k))}: <b class="num">${n}</b>`).join(' · ')}${late ? ` · <span class="late-txt">${t('nLate', { n: late })}</span>` : ''}</div>` : ''), ts.slice(0, 30).map(qRow), t('qaNone')) + `<button type="button" class="linkbtn" data-unitf="${q.unit}">${t('openUnitTasks')}</button>`; }
    case 'project': { const key = normAr(q.project); const ts = sortTasks(live().filter(x => normAr(x.project || '').includes(key) && isOpen(x))); return list(t('qaProject', { p: esc(q.project), n: plural(ts.length, 'tasks') }), ts.map(qRow), t('qaNone')); }
    case 'done': { const ts = live().filter(x => x.status === 'done' && x.completed_on >= q.from && x.completed_on <= q.to).sort((a, b) => a.completed_on < b.completed_on ? 1 : -1); const nl = ts.filter(isLog).length; return list(t('qaDone1', { r: rangeLabel(q.from, q.to), n: [ts.length - nl ? plural(ts.length - nl, 'doneN') : '', nl ? plural(nl, 'logs') : ''].filter(Boolean).join(' · ') || plural(0, 'doneN') }), ts.map(qRow), t('qaNoneDone')) + `<button type="button" class="linkbtn" data-act="chatfill">${sic('plus', 14)} ${t('qLog')}</button>`; }
    case 'due': { const ts = sortTasks(open.filter(x => x.due && x.due >= q.from && x.due <= q.to)); const fs = fus.filter(([, f]) => f.due >= q.from && f.due <= q.to); const late = q.from <= T() && q.to >= T() ? sortTasks(open.filter(isLate)) : [];
      return list(fs.length ? t('qaDue', { r: rangeLabel(q.from, q.to), n: plural(ts.length, 'tasks'), m: plural(fs.length, 'fus') }) : t('qaDue1', { r: rangeLabel(q.from, q.to), n: plural(ts.length, 'tasks') }), [...ts.map(qRow), ...fs.map(([x, f]) => qFu(x, f))], t('qaFree')) + (late.length ? `<div class="qhead sm">${t('plusLate', { n: late.length })}</div><div class="qlist">${late.map(qRow).join('')}</div>` : ''); }
    default: return briefHTML(true);
  }
}
function briefHTML(plain) {
  const td = T(), open = live().filter(x => isOpen(x) && !x.archived);
  const late = open.filter(isLate).length, today = open.filter(x => x.due === td).length;
  let fusToday = 0; open.forEach(x => openFus(x).forEach(f => { if (f.due && f.due <= td) fusToday++; }));
  const hr = new Date().getHours(); const name = (state.profile.display_name || '').split(' ')[0];
  const hello = plain ? '' : `<b>${t(hr < 12 ? 'goodMorning' : 'hello2', { n: esc(name) })}</b><br>`;
  if (!open.length) return hello + t('briefEmpty');
  if (!today && !fusToday && !late) return hello + t('briefClear') + `<div class="chips"><button type="button" class="chip" data-act="chatfill">${sic('check', 14)}${t('qLog')}</button>${qChip('qTomorrow')}</div>`;
  const parts = [today ? (today === 2 && L() === 'ar' ? t('briefDue2') : t('briefDue', { n: plural(today, 'tasks') })) : '', fusToday ? plural(fusToday, 'fus') : '', late ? (late === 2 && L() === 'ar' ? t('briefLate2') : t('briefLate', { n: plural(late, 'tasks') })) : ''].filter(Boolean);
  return hello + t('briefHead') + parts.join(L() === 'ar' ? '، و' : ', ') + '.' + `<div class="chips">${[['qToday', 1], ['qLate', late], ['qFollow', 1]].filter(([, v]) => v).map(([k]) => qChip(k)).join('')}</div>`;
}

/* ---------- voice ----------
   Listening goes on until the user taps the mic again: a pause to think never ends it. The browser's own recognizer
   stops after a silence (always on phones), so it is restarted at once and the words keep adding up in the box.
   Nothing is sent by itself — the user reviews the text, edits it if needed, then taps send. */
let rec = null, MIC = null;
const IS_IOS = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
const IS_PHONE = IS_IOS || /android/i.test(navigator.userAgent);
const MIC_IDLE = 60000; // a whole minute of silence ends listening (battery)
const kbdKey = 'rafeeq2.micKbd';
const micKbd = () => { try { return localStorage.getItem(kbdKey) === (window.RAFEEQ_VERSION || '1'); } catch { return false; } }; // retried once on each new release
const setMicKbd = () => { try { localStorage.setItem(kbdKey, window.RAFEEQ_VERSION || '1'); } catch { } };
function kbdHint(ta) { ta.focus(); toast(t(IS_IOS ? 'micUseKeyboard' : 'micUnsupported'), null, undefined, 4500); }
function toggleMic() {
  const ta = $('#chatIn'); if (!ta) return;
  if (MIC) return stopMic(true);
  if (!SR || micKbd()) return kbdHint(ta);
  if (!navigator.onLine) { toast(t('micOffline'), null, 'err'); return; }
  MIC = { ta, base: ta.value ? ta.value.trimEnd() + ' ' : '', fin: '', interim: '', ok: false, quiet: Date.now(), fast: 0 };
  const S = MIC; ta.addEventListener('input', () => { if (MIC === S || S.closing) { S.base = ta.value.trimEnd() + ' '; S.fin = ''; S.interim = ''; } }); // the user's own edit while listening is kept
  micStart(S); micUI();
}
function micShow(S) { if (!S.ta.isConnected) return; S.ta.value = (S.base + S.fin + S.interim).replace(/^\s+/, ''); autoGrow(S.ta); }
function micStart(S) {
  let r; try { r = new SR(); } catch { micEnd(S); toast(t('micUnsupported'), null, 'err'); return; }
  rec = r; S.t0 = Date.now(); S.got = false;
  r.lang = L() === 'ar' ? 'ar-SA' : 'en-US'; r.interimResults = true; r.maxAlternatives = 1;
  r.continuous = !IS_PHONE; // phones misbehave in continuous mode (repeated words): short sessions, restarted
  r.onresult = e => {
    if (MIC !== S && S.closing !== r) return;
    S.ok = S.got = true; S.quiet = Date.now(); let interim = '';
    for (let k = e.resultIndex; k < e.results.length; k++) { const x = e.results[k], w = (x[0]?.transcript || '').trim(); if (!w) continue; if (x.isFinal) S.fin += w + ' '; else interim += w + ' '; }
    S.interim = interim; micShow(S);
  };
  r.onerror = ev => {
    if (rec !== r) return;
    const code = ev.error || '';
    if (code === 'no-speech' || code === 'aborted') return; // a silence: onend restarts
    if ((code === 'not-allowed' || code === 'service-not-allowed') && !S.ok && IS_IOS) { setMicKbd(); micEnd(S); kbdHint(S.ta); return; } // this iPhone mode has no web dictation: the keyboard's mic from now on
    micEnd(S);
    toast(t(code === 'not-allowed' || code === 'service-not-allowed' ? 'micDenied' : code === 'network' ? 'micNetwork' : 'micError'), null, 'err');
  };
  r.onend = () => {
    if (S.closing === r) { S.closing = null; if (S.interim) { S.fin += S.interim; S.interim = ''; micShow(S); } return; }
    if (rec !== r || MIC !== S) return;
    if (S.interim) { S.fin += S.interim; S.interim = ''; micShow(S); } // some phones never mark the last words final
    if (Date.now() - S.quiet > MIC_IDLE) { micEnd(S); toast(t(S.ok ? 'micIdle' : 'micNoSpeech'), null, S.ok ? undefined : 'err'); return; }
    S.fast = !S.got && Date.now() - S.t0 < 700 ? S.fast + 1 : 0; // a recognizer that dies at once, again and again, is not coming back
    if (S.fast >= 6) { micEnd(S); toast(t('micError'), null, 'err'); return; }
    setTimeout(() => { if (MIC === S) micStart(S); }, S.fast ? 400 : 60);
  };
  try { r.start(); } catch { micEnd(S); toast(t('micError'), null, 'err'); }
}
function micEnd(S) { if (MIC === S) MIC = null; rec = null; micUI(); }
/** keep = the user tapped the mic to finish: let the last words arrive. Otherwise (sending, closing) drop them. */
function stopMic(keep = false) {
  const S = MIC; if (!S) return;
  const r = rec; MIC = null; rec = null;
  if (r) { if (keep) { S.closing = r; try { r.stop(); } catch { } setTimeout(() => { if (S.closing === r) { S.closing = null; try { r.abort(); } catch { } } }, 1500); } else { try { r.abort(); } catch { } } }
  if (keep && S.interim) { S.fin += S.interim; S.interim = ''; micShow(S); }
  micUI();
  if (keep && S.ta.isConnected) { S.ta.focus(); if ((S.fin || '').trim()) toast(t('micReview'), null, undefined, 3500); }
}
function micUI() {
  const b = $('#micBtn'); if (b) { const on = !!MIC; if (b.classList.contains('rec') !== on) b.innerHTML = on ? I('<rect x="7" y="7" width="10" height="10" rx="2" fill="currentColor" stroke="none"/>') : ic.mic; b.classList.toggle('rec', on); b.setAttribute('aria-label', t(on ? 'micStop' : 'speak')); }
  const s = $('#recState'); if (s) s.innerHTML = MIC ? `<span class="wave">${'<i></i>'.repeat(7)}</span>${t('listeningLong')}` : '';
}
function autoGrow(ta) { ta.style.height = 'auto'; ta.style.height = Math.min(ta.scrollHeight, 160) + 'px'; }

/* ---------- reminder message for a follow-up ---------- */
function remindText(x, f) {
  const p = personById(f.person_id); const nm = p ? p.name : '';
  const honor = /^(?:م|د|أ|ا)\.|^(?:مهندس|دكتور|أستاذ|استاذ|Eng|Dr|Mr|Ms)/.test(nm) ? nm : (L() === 'ar' ? `أ. ${nm}` : nm);
  const me = state.profile.display_name || '';
  if (L() === 'en') return `Hello ${honor},\nKindly ${f.what ? f.what.toLowerCase() : 'update me'} regarding “${x.title}”${f.due ? ` by ${fmt(f.due)}` : ''}.\nThank you,\n${me}`;
  return `السلام عليكم ${honor}،\nأرجو التكرّم ${f.what ? 'بـ' + f.what : 'بالإفادة'} بخصوص «${x.title}»${x.due ? `، علمًا بأن موعد التسليم ${wd(x.due)} ${fmt(x.due)}` : ''}.\nشاكرًا تعاونكم،\n${me}`;
}
const waNumber = c => { const d = String(c || '').replace(/[^\d+]/g, '').replace(/^\+/, '').replace(/^00/, '').replace(/^0(5\d{8})$/, '966$1').replace(/^0(1\d{9})$/, '20$1'); return /^\d{8,15}$/.test(d) ? d : ''; };
function openRemind(x, f) {
  const p = personById(f.person_id); const phone = waNumber(p?.contact);
  openSheet(t('remindTitle', { p: esc(pname(f.person_id)) }), `<div class="form"><label class="fld"><span>${t('remindMsg')}</span><textarea id="remTxt" rows="6" dir="auto">${esc(remindText(x, f))}</textarea></label>
    <div class="save-row"><button type="button" class="btn primary block" data-act="remwa" data-id="${x.id}" data-fu="${esc(f.id)}" data-phone="${esc(/^\d{8,15}$/.test(phone) ? phone : '')}">${t('sendWhatsApp')}</button><button type="button" class="btn block" data-act="remcopy" data-id="${x.id}" data-fu="${esc(f.id)}">${t('copyMsg')}</button></div>
    ${/^\d{8,15}$/.test(phone) ? '' : `<p class="note">${t('remindNoPhone')}</p>`}</div>`);
}
function logReminder(x, f) { f.log = f.log || []; f.log.push({ id: uuid(), at: nowISO(), date: T(), text: t('reminderSent'), outcome: 'reminded' }); addLog(x, 'fu_reminded', '', { person_id: f.person_id }); putTask(x); }

/* ---------- full task form ---------- */
/** A department/section picker: each department, then its sections under it. */
function unitSelect(name, cur, emptyLbl) {
  return `<select name="${name}"><option value="">${esc(emptyLbl || t('noUnit'))}</option>${depts().map(d => `<option value="${d.id}" ${cur === d.id ? 'selected' : ''}>${esc(d.name)}</option>${secsOf(d.id).map(sc => `<option value="${sc.id}" ${cur === sc.id ? 'selected' : ''}>&nbsp;&nbsp;&nbsp;↳ ${esc(sc.name)}</option>`).join('')}`).join('')}</select>`;
}
function openTaskForm(x, pre = null) {
  const isNew = !x; const v = x ? (pre?.status ? { ...x, status: pre.status } : x) : newTask(pre ? { title: pre.title || '', due: pre.due || null, due_time: pre.time || null, recur: pre.recur || null, priority: pre.priority || 'mid', role: pre.role || 'exec', source: pre.source || (askSource() ? '' : SELF), project: pre.project || '', ...Object.fromEntries(Object.entries(pre.extra || {}).filter(([k, val]) => val !== undefined && k !== 'waiting_on')) } : { source: askSource() ? '' : SELF });
  const isDraft = !!pre?.draftId;
  const rc = v.recur || null, rcDom = rc && rc.dom !== 'last' && rc.dom ? rc.dom : pd(v.due || T()).getDate(), rcDays = rc?.days?.length ? rc.days : [pd(v.due || T()).getDay()];
  const projects = [...new Set(live().map(y => (y.project || '').trim()).filter(Boolean))];
  const srcOpts = [...new Set([...recentSources(), ...people().map(p => p.name)])];
  const stOpts = OPEN_ST().concat(isNew ? [] : ['cancelled']);
  const whoNames = pre?.who || (x ? [...new Set(openFus(x).map(f => pname(f.person_id)))] : []);
  const noDue0 = pre ? !!pre.noDue : (!isNew && !v.due);
  openSheet(isNew ? t('newTask') : t('editTask'), `<form id="taskForm" class="form" autocomplete="off" novalidate>
      <label class="fld"><span>${t('fTitle')} <b class="req">*</b></span><textarea name="title" rows="2" maxlength="500" dir="auto" autofocus>${esc(v.title)}</textarea></label>
      <label class="fld"><span>${t('fDetails')} <em>${t('optional')}</em></span><textarea name="details" rows="2" dir="auto">${esc(v.details)}</textarea></label>
      <div class="grid2">
        <div class="fld"><span>${t('fDue')} <b class="req">*</b></span><input type="date" name="due" value="${esc(v.due || '')}" ${noDue0 ? 'disabled' : ''}><label class="chk-line"><input type="checkbox" name="noDue" ${noDue0 ? 'checked' : ''}> ${t('noDue')}</label></div>
        <label class="fld"><span>${t('fPriority')}</span><select name="priority">${['hi', 'mid', 'lo'].map(p => `<option value="${p}" ${v.priority === p ? 'selected' : ''}>${t(p === 'hi' ? 'prioHi' : p === 'lo' ? 'prioLo' : 'prioMid')}</option>`).join('')}</select></label>
      </div>
      <div class="grid2">
        <label class="fld"><span>${t('fTime')} <em>${t('optional')}</em></span><input type="time" name="tm" value="${esc(v.due_time || '')}"></label>
        <label class="fld" id="rmBox" ${v.due_time ? '' : 'hidden'}><span>${t('fRemind')}</span><select name="rm">${[['', 'rmDefault'], ['0'], ['15'], ['30'], ['60'], ['120'], ['-1']].map(([val]) => `<option value="${val}" ${String(v.remind_min ?? '') === val ? 'selected' : ''}>${remindLabel(val === '' ? null : +val)}</option>`).join('')}</select></label>
      </div>
      <div class="fld"><span>${sic('repeat', 14)} ${t('fRecur')}</span><select name="rf" id="rfSel">${[['', 'rcNone'], ['daily', 'rcDaily'], ['weekly', 'rcWeeklyOpt'], ['monthly', 'rcMonthlyOpt'], ['quarterly', 'rcQuarterlyOpt']].map(([val, k]) => `<option value="${val}" ${(rc?.f || '') === val ? 'selected' : ''}>${t(k, { d: '', n: '' })}</option>`).join('')}</select>
        <div class="opts" id="rfDays" ${rc?.f === 'weekly' ? '' : 'hidden'}>${[0, 1, 2, 3, 4, 5, 6].map(d => `<button type="button" class="chip ${rcDays.includes(d) ? 'on' : ''}" data-rd="${d}">${wdName(d)}</button>`).join('')}</div>
        <div class="grid2" id="rfDom" ${rc?.f === 'monthly' || rc?.f === 'quarterly' ? '' : 'hidden'}><label class="fld"><span>${t('rcDay')}</span><input type="number" name="rdom" min="1" max="31" inputmode="numeric" value="${rcDom}" ${rc?.dom === 'last' ? 'disabled' : ''}></label><label class="chk-line"><input type="checkbox" name="rlast" ${rc?.dom === 'last' ? 'checked' : ''}> ${t('rcLastWork')}</label></div>
        <small class="note" id="rfNote" ${rc ? '' : 'hidden'}>${t('rcNote')}</small></div>
      <div class="fld"><span>${t('fSource')} ${askSource() ? '<b class="req">*</b>' : `<em>${t('optional')}</em>`}</span><div class="src-row"><input name="source" list="srcList" value="${esc(srcLabel(v.source))}" placeholder="${esc(t('fSourcePh'))}" dir="auto"><button type="button" class="chip ${v.source === SELF ? 'on' : ''}" data-srcself>${t('srcSelf')}</button></div></div>
      <div class="fld"><span>${t('fRole')} <b class="req">*</b></span><div class="seg2" id="roleSeg">${['exec', 'follow', 'both'].map(r => `<button type="button" class="${v.role === r ? 'on' : ''}" data-role="${r}">${t('role_' + r)}</button>`).join('')}</div><input type="hidden" name="role" value="${v.role}"></div>
      <label class="fld" id="whoBox" ${v.role === 'exec' ? 'hidden' : ''}><span>${t('fWho')} <b class="req">*</b></span><input name="who" list="pplList" value="${esc(whoNames.join(L() === 'ar' ? '، ' : ', '))}" placeholder="${esc(t('personPhMulti'))}" dir="auto"></label>
      ${v.status === 'done' ? '' : `<label class="fld"><span>${t('fStatus')}</span><select name="status" id="fStatus">${stOpts.map(s => `<option value="${s}" ${v.status === s ? 'selected' : ''}>${t(stKey(s))}</option>`).join('')}</select></label>`}
      <div class="grid2" id="waitBox" ${v.status === 'wait' ? '' : 'hidden'}>
        <label class="fld"><span>${t('fWaitingOn')}</span><input name="waiting_on" list="pplList" value="${esc(pre?.extra?.waiting_on || (v.waiting_on ? pname(v.waiting_on) : ''))}" dir="auto"></label>
        <label class="fld"><span>${t('fWaitingWhat')}</span><select name="waiting_what">${['reply', 'decision', 'approval'].map(w => `<option value="${w}" ${v.waiting_what === w ? 'selected' : ''}>${t('w_' + w)}</option>`).join('')}</select></label>
      </div>
      ${orgOn() && depts().length ? `<label class="fld"><span>${sic('org', 14)} ${t('fUnit')} <em>${t('optional')}</em></span>${unitSelect('unit', v.unit_id || pre?.unit || '')}</label>` : ''}
      <div class="grid2">
        <label class="fld"><span>${t('fProject')} <em>${t('optional')}</em></span><input name="project" list="projList" value="${esc(v.project)}" dir="auto"></label>
        <label class="fld"><span>${t('fSteps')} <em>${t('fStepsHint')}</em></span><input type="number" name="steps_total" min="0" max="50" inputmode="numeric" value="${v.steps_total || ''}"></label>
      </div>
      <label class="fld"><span>${t('fNotes')} <em>${t('fNotesHint')}</em></span><textarea name="notes" rows="2" dir="auto">${esc(v.notes)}</textarea></label>
      <datalist id="pplList">${people().map(p => `<option value="${esc(p.name)}"></option>`).join('')}</datalist>
      <datalist id="projList">${projects.map(p => `<option value="${esc(p)}"></option>`).join('')}</datalist>
      <datalist id="srcList">${srcOpts.map(p => `<option value="${esc(p)}"></option>`).join('')}</datalist>
      <p class="note err-txt" id="formErr" hidden></p>
      <div class="save-row sticky"><button type="submit" class="btn primary block">${ic.check}${isDraft ? t('draftApply') : t('save')}</button></div>
      ${isDraft ? `<p class="note">${t('draftApplyNote')}</p>` : ''}
    </form>`, {
    onMount: sh => {
      const f = sh.querySelector('#taskForm');
      const st = sh.querySelector('#fStatus'); if (st) st.onchange = () => sh.querySelector('#waitBox').hidden = st.value !== 'wait';
      sh.querySelectorAll('[data-role]').forEach(b => b.onclick = () => { sh.querySelectorAll('[data-role]').forEach(y => y.classList.remove('on')); b.classList.add('on'); f.elements.role.value = b.dataset.role; sh.querySelector('#whoBox').hidden = b.dataset.role === 'exec'; });
      f.elements.noDue.onchange = () => { f.elements.due.disabled = f.elements.noDue.checked; if (f.elements.noDue.checked) f.elements.due.value = ''; };
      f.elements.tm.oninput = () => { sh.querySelector('#rmBox').hidden = !f.elements.tm.value; };
      const rfSync = () => { const v2 = f.elements.rf.value; sh.querySelector('#rfDays').hidden = v2 !== 'weekly'; sh.querySelector('#rfDom').hidden = !(v2 === 'monthly' || v2 === 'quarterly'); sh.querySelector('#rfNote').hidden = !v2; };
      f.elements.rf.onchange = rfSync;
      sh.querySelectorAll('[data-rd]').forEach(b => b.onclick = () => b.classList.toggle('on'));
      f.elements.rlast.onchange = () => { f.elements.rdom.disabled = f.elements.rlast.checked; };
      const selfBtn = sh.querySelector('[data-srcself]'); selfBtn.onclick = () => { f.elements.source.value = t('srcSelf'); selfBtn.classList.add('on'); };
      f.elements.source.oninput = () => selfBtn.classList.toggle('on', f.elements.source.value.trim() === t('srcSelf'));
      f.onsubmit = e => {
        e.preventDefault();
        const fd = new FormData(f); const err = msg => { const el = sh.querySelector('#formErr'); el.textContent = msg; el.hidden = false; };
        const title = (fd.get('title') || '').trim(); if (!title) { f.elements.title.focus(); return err(t('reqTitle')); }
        let dueV = validDate(f.elements.due.value); const noDue = f.elements.noDue.checked;
        const rfv = fd.get('rf'); let recur = null;
        if (rfv) { const days = [...sh.querySelectorAll('[data-rd].on')].map(b => +b.dataset.rd); recur = cleanRecur({ f: rfv, days: days.length ? days : [pd(dueV || T()).getDay()], dom: f.elements.rlast.checked ? 'last' : (parseInt(f.elements.rdom.value, 10) || pd(dueV || T()).getDate()) }); }
        if (recur && !dueV) dueV = recurNext(recur, addDays(T(), -1));
        if (!dueV && !noDue) { f.elements.due.focus(); return err(t('reqDue')); }
        const srcRaw = (fd.get('source') || '').trim();
        if (!srcRaw && askSource()) { f.elements.source.focus(); return err(t('reqSource')); }
        const role = fd.get('role') || 'exec';
        const whoList = role === 'exec' ? [] : String(fd.get('who') || '').split(/[،,+؛;]/).map(s => s.trim()).filter(Boolean);
        if (role !== 'exec' && !whoList.length) { f.elements.who.focus(); return err(t('reqWho')); }
        if (fd.get('status') === 'wait' && !String(fd.get('waiting_on') || '').trim()) { f.elements.waiting_on.focus(); return err(t('reqWaitingOn')); }
        if (isDraft) { // editing a chat draft changes the draft only: nothing is saved until the user approves in the chat
          const d = CH.drafts.find(y => y.id === pre.draftId); if (!d) { closeSheet(true); return openChat(); }
          d.title = title.slice(0, 300); d.noDue = noDue && !recur; d.due = d.noDue ? null : dueV; d.weekendOk = !!(d.due && isWeekend(d.due)); d.recur = recur;
          const tmv = fd.get('tm') || ''; d.time = /^\d{2}:\d{2}$/.test(tmv) ? tmv : null;
          d.priority = fd.get('priority') || 'mid'; d.prioAsked = true; d.role = role; d.project = (fd.get('project') || '').trim();
          if (f.elements.unit) { d.unit = unitById(f.elements.unit.value) ? f.elements.unit.value : null; d.unitAsked = true; }
          d.source = !srcRaw || /^(أنا|انا|me|self)$/i.test(srcRaw) || srcRaw === t('srcSelf') ? SELF : srcName(srcRaw);
          const keep = d.people.filter(p => p.role === 'waiting' || (p.role === 'related' && role === 'exec'));
          const fu = whoList.map(nm => { const old = d.people.find(p => p.role !== 'waiting' && normAr(refName(p.ref)) === normAr(nm)); return old ? { ...old, role: 'followup' } : { ref: refOf(nm), role: 'followup', what: '', due: null }; }).filter(p => p.ref);
          d.people = [...keep, ...fu];
          const st = fd.get('status'); const rmv = fd.get('rm');
          d.extra = { details: (fd.get('details') || '').trim(), notes: (fd.get('notes') || '').trim(), steps_total: Math.max(0, Math.min(50, parseInt(fd.get('steps_total') || '0', 10) || 0)), remind_min: d.time && rmv !== '' && rmv != null ? +rmv : null, status: st && st !== 'todo' ? st : undefined, waiting_on: st === 'wait' ? String(fd.get('waiting_on') || '').trim() : undefined, waiting_what: st === 'wait' ? fd.get('waiting_what') || 'reply' : undefined };
          closeSheet(true); openChat(); toast(t('draftUpdated')); return;
        }
        const tgt = (x && taskById(x.id)) || v;
        const prevDue = tgt.due || null, prevSt = tgt.status;
        tgt.title = title; tgt.details = (fd.get('details') || '').trim();
        tgt.due = noDue && !recur ? null : dueV; tgt.priority = fd.get('priority');
        const tmv = fd.get('tm') || ''; tgt.due_time = /^\d{2}:\d{2}$/.test(tmv) ? tmv : null;
        const rmv = fd.get('rm'); tgt.remind_min = tgt.due_time && rmv !== '' && rmv != null ? +rmv : null;
        tgt.recur = recur;
        if (fd.get('status')) tgt.status = fd.get('status');
        tgt.role = role; tgt.project = (fd.get('project') || '').trim();
        if (f.elements.unit) tgt.unit_id = unitById(f.elements.unit.value) ? f.elements.unit.value : null;
        tgt.source = !srcRaw || /^(أنا|انا|me|self)$/i.test(srcRaw) || srcRaw === t('srcSelf') ? SELF : srcRaw;
        tgt.steps_total = Math.max(0, Math.min(50, parseInt(fd.get('steps_total') || '0', 10) || 0)); tgt.steps_done = Math.min(tgt.steps_done || 0, tgt.steps_total);
        tgt.notes = (fd.get('notes') || '').trim();
        if (tgt.status === 'wait') { tgt.waiting_on = getOrCreatePerson(fd.get('waiting_on')); if (!tgt.waiting_on) tgt.status = prevSt === 'wait' ? 'prog' : prevSt; tgt.waiting_what = fd.get('waiting_what') || 'reply'; } else tgt.waiting_on = null;
        whoList.forEach(nm => {
          const pid = getOrCreatePerson(nm); if (!pid || (tgt.followups || []).some(fu => fu.person_id === pid && fu.status === 'open')) return;
          const due = tgt.due && tgt.due > T() ? (diffDays(tgt.due, T()) > 1 ? addWorkdays(tgt.due, -1) : tgt.due) : addWorkdays(T(), 1);
          tgt.followups.push({ id: uuid(), person_id: pid, what: '', due, status: 'open', created_at: nowISO(), closed_on: null, log: [] }); addLog(tgt, 'fu_added', '', { person_id: pid });
        });
        if (isNew) addLog(tgt, 'created'); else { if (prevSt !== tgt.status) addLog(tgt, 'status', '', { from: prevSt, to: tgt.status }); if (prevDue !== tgt.due) addLog(tgt, 'due', '', { from: prevDue, to: tgt.due }); }
        putTask(tgt); closeSheet(true); toast(t('saved'));
        if (isNew) openTask(tgt.id);
      };
    }
  });
  sheetState.guard = true;
}

/* ---------- work log entry (something already done, not a task) ---------- */
function openLogForm(x, pre = null) {
  const isNew = !x;
  const v = x || { title: pre?.title || '', completed_on: pre?.date || T(), project: pre?.project || '', details: '' };
  const projects = [...new Set(live().map(y => (y.project || '').trim()).filter(Boolean))];
  openSheet(isNew ? t('newLog') : t('editLog'), `<form id="logForm" class="form" autocomplete="off" novalidate>
      <p class="note">${t('logHint')}</p>
      <label class="fld"><span>${t('fLogWhat')} <b class="req">*</b></span><textarea name="title" rows="2" maxlength="500" dir="auto" placeholder="${esc(t('fLogPh'))}">${esc(v.title)}</textarea></label>
      <div class="grid2">
        <label class="fld"><span>${t('fLogDate')} <b class="req">*</b></span><input type="date" name="d" value="${esc(v.completed_on || T())}" max="${T()}"></label>
        <label class="fld"><span>${t('fProject')} <em>${t('optional')}</em></span><input name="project" list="projList" value="${esc(v.project || '')}" dir="auto"></label>
      </div>
      ${pre?.draftId ? '' : `<label class="fld"><span>${t('fDetails')} <em>${t('optional')}</em></span><textarea name="details" rows="2" dir="auto">${esc(v.details || '')}</textarea></label>`}
      ${pre?.draftId || state.mode !== 'cloud' ? '' : `<div class="fld"><span>${t('attachFiles')} <em>${t('optional')}</em></span>${state.drive.linked === true
        ? `<label class="btn sm filepick">${sic('clip', 16)}${t('chooseFiles')}<input type="file" name="files" multiple hidden></label><div class="picked" id="pickedFiles"></div>`
        : `<button type="button" class="btn sm" data-act="connectdrive">${sic('folder', 16)}${t('driveForFiles')}</button>`}</div>`}
      <datalist id="projList">${projects.map(p => `<option value="${esc(p)}"></option>`).join('')}</datalist>
      <p class="note err-txt" id="formErr" hidden></p>
      <div class="save-row sticky"><button type="submit" class="btn primary block">${ic.check}${isDraft ? t('draftApply') : t('save')}</button></div>
      ${isDraft ? `<p class="note">${t('draftApplyNote')}</p>` : ''}
    </form>`, {
    onMount: sh => {
      const f = sh.querySelector('#logForm');
      setTimeout(() => f.elements.title.focus(), 200);
      const fi = f.elements.files; if (fi) fi.onchange = () => { sh.querySelector('#pickedFiles').innerHTML = [...fi.files].map(x => `<span class="chip sm">${sic('clip', 12)}<span dir="auto">${esc(x.name)}</span></span>`).join(''); };
      f.onsubmit = e => {
        e.preventDefault();
        const err = msg => { const el = sh.querySelector('#formErr'); el.textContent = msg; el.hidden = false; };
        const title = f.elements.title.value.replace(/\s+/g, ' ').trim(); if (!title) { f.elements.title.focus(); return err(t('reqLogTitle')); }
        const dt = validDate(f.elements.d.value); if (!dt || dt > T()) { f.elements.d.focus(); return err(t('reqLogDate')); }
        const proj = (f.elements.project.value || '').trim().slice(0, 80);
        if (pre?.draftId) {
          const d = CH.drafts.find(y => y.id === pre.draftId); if (d) { d.title = title.slice(0, 300); d.date = dt; d.project = proj; d.match = matchOpenTask(d.title); d.closeTask = undefined; }
          closeSheet(true); return openChat();
        }
        const tgt = x ? taskById(x.id) : newTask({ kind: 'log', status: 'done', due: null, source: '' });
        tgt.title = title.slice(0, 300); tgt.completed_on = dt; tgt.project = proj; tgt.details = (f.elements.details?.value || '').trim();
        if (isNew) addLog(tgt, 'logged');
        const picked = fi ? [...fi.files] : [];
        putTask(tgt); closeSheet(true); toast(t(isNew ? 'loggedToast' : 'saved'));
        if (picked.length) queueFiles(picked, { task: tgt.id });
      };
    }
  });
  sheetState.guard = true;
}

/* ---------- complete ---------- */
const fresh = x => taskById(x.id) || x;
const freshFu = (x, f) => (fresh(x).followups || []).find(y => y.id === f.id) || f;
function openComplete(x0) {
  const x = x0;
  openSheet(t('completeTask'), `<form id="cForm" class="form"><p class="ctx" dir="auto">${esc(x.title)}</p>
      <label class="fld"><span>${t('completedOn')}</span><input type="date" name="d" value="${T()}" max="${T()}" required></label>
      <label class="fld"><span>${t('result')} <em>${t('optional')}</em></span><textarea name="r" rows="3" dir="auto" placeholder="${esc(t('resultPh'))}" autofocus></textarea></label>
      <div class="save-row"><button class="btn primary block" type="submit">${ic.check}${t('complete')}</button></div></form>`, {
    onMount: sh => sh.querySelector('#cForm').onsubmit = e => {
      e.preventDefault(); const fd = new FormData(e.target); const x = fresh(x0); if (!isOpen(x)) { closeSheet(true); return; }
      const prev = structuredClone(x);
      x.status = 'done'; x.completed_on = validDate(fd.get('d')) || T(); x.result = (fd.get('r') || '').trim(); addLog(x, 'completed', x.result);
      const nx = spawnNext(x);
      const undo = () => { putTask(prev); if (nx) { nx.deleted = true; putTask(nx); } };
      putTask(x); closeSheet(true); toast(t('completedToast') + (nx ? ' · ' + t('nextOn', { d: rel(nx.due) }) : ''), undo); askCloseFus(x, undo);
    }
  });
}

/* ---------- follow-ups ---------- */
function openFuAdd(x0, f0) {
  const isNew = !f0; const x = x0, f = f0;
  openSheet(isNew ? t('addFu') : t('editFu'), `<form id="fuForm" class="form" autocomplete="off"><p class="ctx" dir="auto">${esc(x.title)}</p>
      <label class="fld"><span>${isNew ? t('personMulti') : t('person')}</span><input name="p" list="pplList2" required dir="auto" value="${esc(f ? pname(f.person_id) : '')}" placeholder="${esc(isNew ? t('personPhMulti') : t('personPh'))}" autofocus></label>
      <label class="fld"><span>${t('fuWhat')}</span><input name="w" dir="auto" value="${esc(f?.what || '')}" placeholder="${esc(t('fuWhatPh'))}"></label>
      <div class="fld"><span>${t('fuDue')}</span><div class="opts">${dateChips().map(([l, v]) => `<button type="button" class="chip" data-quick="${v}">${l}</button>`).join('')}</div><input type="date" name="d" value="${esc(f?.due || addWorkdays(T(), 1))}" required></div>
      <label class="fld"><span>${t('fTime')} <em>${t('optional')}</em></span><input type="time" name="tm" value="${esc(f?.time || '')}"></label>
      <datalist id="pplList2">${people().map(p => `<option value="${esc(p.name)}"></option>`).join('')}</datalist>
      <div class="save-row"><button class="btn primary block" type="submit">${ic.check}${t('save')}</button></div></form>`, {
    onMount: sh => {
      const fm = sh.querySelector('#fuForm');
      sh.querySelectorAll('[data-quick]').forEach(b => b.onclick = () => { fm.elements.d.value = b.dataset.quick; sh.querySelectorAll('[data-quick]').forEach(y => y.classList.remove('on')); b.classList.add('on'); });
      fm.onsubmit = e => {
        e.preventDefault(); const fd = new FormData(fm); const x = fresh(x0), f = f0 && freshFu(x0, f0);
        if (isNew) {
          const ids = [...new Set(String(fd.get('p') || '').split(/[،,+؛;]/).map(n => getOrCreatePerson(n)).filter(Boolean))]; if (!ids.length) return;
          ids.forEach(pid => { x.followups.push({ id: uuid(), person_id: pid, what: (fd.get('w') || '').trim(), due: validDate(fd.get('d')) || nextWorkday(), time: /^\d{2}:\d{2}$/.test(fd.get('tm') || '') ? fd.get('tm') : null, status: 'open', created_at: nowISO(), closed_on: null, log: [] }); addLog(x, 'fu_added', '', { person_id: pid }); });
        }
        else { const pid = getOrCreatePerson(fd.get('p')); if (!pid) return; const d0 = validDate(fd.get('d')); if (d0 && d0 < T() && f.status === 'open') { toast(t('pastDate')); return; } f.person_id = pid; f.what = (fd.get('w') || '').trim(); f.due = validDate(fd.get('d')) || f.due; f.time = /^\d{2}:\d{2}$/.test(fd.get('tm') || '') ? fd.get('tm') : null; }
        if (isNew && x.role === 'exec') x.role = 'both';
        putTask(x); closeSheet(true); toast(t('saved'));
      };
    }
  });
  sheetState.guard = true;
}
function openFuResult(x, f) {
  openSheet(t('fuResult'), `<div class="fu-head"><span class="av">${esc(initials(pname(f.person_id)))}</span><div><b dir="auto">${esc(pname(f.person_id))}</b><div class="sub" dir="auto">${esc(f.what || x.title)}</div></div></div>
      <label class="fld"><span>${t('whatHappened')} <em>${t('optional')}</em></span><textarea id="fuNote" rows="2" dir="auto" placeholder="${esc(t('whatHappenedPh'))}"></textarea></label>
      <div class="save-row"><button type="button" class="btn ok block" data-act="fudone">${ic.check}${t('fuDone')}</button><button type="button" class="btn block" data-act="funotyet">${t('fuNotYet')}</button></div>
      <div id="fuNext" hidden><div class="lab">${t('nextFu')}</div><div class="opts">${dateChips().filter(([, v]) => v > T()).map(([l, v]) => `<button type="button" class="chip" data-act="funext" data-v="${v}">${l}</button>`).join('')}<label class="chip datein">${sic('cal', 15)}<input type="date" id="fuNextIn" min="${addDays(T(), 1)}" aria-label="${esc(t('pickDate'))}"></label></div></div>`, {
    onMount: sh => { const di = sh.querySelector('#fuNextIn'); di.onchange = () => validDate(di.value) && fuAgain(x, f, di.value); }
  });
  UI.fuCtx = { x, f };
}
function fuDone(x0, f0) {
  const x = fresh(x0), f = freshFu(x0, f0); if (f.status !== 'open') { closeSheet(true); return; }
  const note = ($('#fuNote')?.value || '').trim();
  f.log = f.log || []; f.log.push({ id: uuid(), at: nowISO(), date: T(), text: note, outcome: 'done' });
  f.status = 'done'; f.closed_on = T(); addLog(x, 'fu_done', note, { person_id: f.person_id });
  putTask(x); closeSheet(true); toast(t('fuLogged'));
}
function fuAgain(x0, f0, d) {
  const x = fresh(x0), f = freshFu(x0, f0); if (!validDate(d) || d <= T()) { toast(t('pastDate')); return; }
  const note = ($('#fuNote')?.value || '').trim();
  f.log = f.log || []; f.log.push({ id: uuid(), at: nowISO(), date: T(), text: note, outcome: 'again', next: d });
  f.due = d; addLog(x, 'fu_again', note, { person_id: f.person_id, due: d });
  putTask(x); closeSheet(true); toast(t('remindOn', { d: rel(d) }));
}

/* ---------- first-use tour: three short screens, shown once ---------- */
const TOUR = [
  { ic: 'spark', h: 'tour1H', p: 'tour1P', ex: 'tour1Ex' },
  { ic: 'users', h: 'tour2H', p: 'tour2P', roles: true },
  { ic: 'cal', h: 'tour3H', p: 'tour3P' },
];
function tourSeen() { try { return !!(state.profile.settings || {}).tourDone || localStorage.getItem('rafeeq2.tour.' + (state.user?.id || 'local')) === '1'; } catch { return !!(state.profile.settings || {}).tourDone; } }
function maybeTour() {
  if (!state.mode || UI.booting || $('#onbForm') || $('#sheet').classList.contains('on') || UI.layers.length || tourSeen()) return;
  openTour(0);
}
function openTour(i) {
  const st = TOUR[i]; if (!st) return;
  if (i > 0) { try { localStorage.setItem('rafeeq2.tour.' + (state.user?.id || 'local'), '1'); } catch { } }
  const last = i === TOUR.length - 1;
  openSheet(t('tourTitle'), `<div class="tour">
      <div class="tour-ic">${ic[st.ic]}</div>
      <h4>${t(st.h)}</h4><p>${t(st.p)}</p>
      ${st.ex ? `<div class="tour-ex" dir="auto">${t(st.ex)}</div>` : ''}
      ${st.roles ? `<ul class="tour-roles"><li><b>${t('role_exec')}</b> — ${t('tourRoleExec')}</li><li><b>${t('role_follow')}</b> — ${t('tourRoleFollow')}</li><li><b>${t('role_both')}</b> — ${t('tourRoleBoth')}</li></ul>` : ''}
      <div class="tour-dots" aria-hidden="true">${TOUR.map((_, k) => `<i class="${k === i ? 'on' : ''}"></i>`).join('')}</div>
      <div class="save-row">${last
        ? `<button type="button" class="btn primary block" data-act="tourdone" data-chat="1">${ic.spark}${t('tourStart')}</button>`
        : `<button type="button" class="btn primary block" data-act="tournext" data-v="${i + 1}">${t('tourNext')}</button>`}
        <button type="button" class="btn block" data-act="tourdone">${last ? t('close') : t('tourSkip')}</button></div>
    </div>`);
}

/* ---------- notifications & weekend ---------- */
const WK_OPTS = [[[5, 6], 'wkFriSat'], [[0, 6], 'wkSatSun'], [[5], 'wkFri'], [[6], 'wkSat'], [[0], 'wkSun']];
function wkLabel() { const cur = ((state.profile.settings || {}).wk || [5, 6]).slice().sort().join(); const o = WK_OPTS.find(([d]) => d.slice().sort().join() === cur); return o ? t(o[1]) : cur; }
function openWeekendSheet() {
  const cur = wkLabel();
  openSheet(t('wkTitle'), `<p class="note" style="margin-top:0">${t('wkHint')}</p><div class="menu">${WK_OPTS.map(([d, k]) => `<button type="button" data-act="wkset" data-v="${d.join(',')}"><span class="grow">${t(k)}</span>${cur === t(k) ? sic('check', 18) : ''}</button>`).join('')}</div>`);
}
function notifSummary() {
  const nt = (state.profile.settings || {}).notif || {};
  if (UI.pushOn === true) return t('notifOnSum', { b: nt.brief === false ? t('off') : fmtTime(nt.brief || '07:30') });
  return UI.pushOn === false ? t('notifOffSum') : '…';
}
function saveNotifPrefs(f) {
  const st = Object.assign({}, state.profile.settings); const nt = Object.assign({ brief: '07:30', weekend: false, remind: 0 }, st.notif || {});
  if (f) {
    nt.brief = f.elements.briefOn.checked ? (/^\d{2}:\d{2}$/.test(f.elements.brief.value) ? f.elements.brief.value : '07:30') : false;
    nt.weekend = f.elements.wkOn.checked;
    const r = parseInt(f.elements.remind.value, 10); nt.remind = Number.isInteger(r) ? r : 0;
  }
  nt.tz = Intl.DateTimeFormat().resolvedOptions().timeZone || 'Asia/Riyadh'; nt.on = true;
  st.notif = nt; saveProfile({ settings: st }, { quiet: true });
}
async function openNotifSheet() {
  const nt = Object.assign({ brief: '07:30', weekend: false, remind: 0 }, (state.profile.settings || {}).notif || {});
  const sub = await pushCurrent().catch(() => null); UI.pushOn = !!sub;
  const ios = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const standalone = matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
  const perm = 'Notification' in window ? Notification.permission : 'unsupported';
  let top;
  if (!pushSupported() || (ios && !standalone)) top = ios && !standalone ? `<div class="notice"><b>${t('notifIosTitle')}</b><ol class="steps">${t('notifIosSteps')}</ol></div>` : `<p class="note">${t('notifUnsupported')}</p>`;
  else if (sub) top = `<div class="notice ok-notice">${sic('check', 18)} <b>${t('notifOnHere')}</b></div><div class="save-row"><button type="button" class="btn block" data-act="notiftest">${sic('bell', 17)}${t('notifTest')}</button><button type="button" class="btn ghost block" data-act="notifoff">${t('notifOff')}</button></div>`;
  else if (perm === 'denied') top = `<p class="note err-txt">${t('notifBlocked')}</p>`;
  else top = `<p class="note" style="margin-top:0">${t('notifIntro')}</p><button type="button" class="btn primary block" data-act="notifon">${ic.bell}${t('notifEnable')}</button>`;
  openSheet(t('notifTitle'), `<form id="nForm" class="form">${top}
      <div class="sec-h"><h3>${t('notifPrefs')}</h3></div>
      <label class="chk-line"><input type="checkbox" name="briefOn" ${nt.brief !== false ? 'checked' : ''}> ${t('briefOn')}</label>
      <label class="fld"><span>${t('briefTime')}</span><input type="time" name="brief" value="${esc(nt.brief || '07:30')}"></label>
      <label class="chk-line"><input type="checkbox" name="wkOn" ${nt.weekend ? 'checked' : ''}> ${t('notifWeekend')}</label>
      <label class="fld"><span>${t('remindDefault')}</span><select name="remind">${[0, 15, 30, 60, 120, -1].map(m => `<option value="${m}" ${nt.remind === m ? 'selected' : ''}>${remindLabel(m)}</option>`).join('')}</select></label>
      <p class="note">${t('notifPrivacy')}</p>
      <div class="save-row sticky"><button type="submit" class="btn primary block">${ic.check}${t('save')}</button></div></form>`, {
    onMount: sh => { const f = sh.querySelector('#nForm'); f.onsubmit = e => { e.preventDefault(); saveNotifPrefs(f); closeSheet(true); toast(t('saved')); renderMain(); }; }
  });
}
/** Open the task a notification points to (#t=<id>). */
function routeHash() {
  const m = /^#t=([0-9a-f-]{36})$/i.exec(location.hash || ''); if (!m || !state.mode) return;
  if (!taskById(m[1])) return;
  history.replaceState(null, '', location.pathname + location.search);
  closeAllLayers(); openTask(m[1]);
}
window.addEventListener('hashchange', routeHash);
if ('serviceWorker' in navigator) navigator.serviceWorker.addEventListener('message', e => { const id = e.data && e.data.openTask; if (!id || !/^[0-9a-f-]{36}$/i.test(id)) return; if (taskById(id)) { closeAllLayers(); openTask(id); } else history.replaceState(null, '', location.pathname + location.search + '#t=' + id); });

/* ---------- people ---------- */
/* ---------- departments & sections ---------- */
function openOrg() {
  pushLayer(() => {
    const ds = depts(), open = live().filter(x => isOpen(x) && !x.archived);
    const cnt = ids => { const set = new Set(ids); const ts = open.filter(x => set.has(x.unit_id)); return { n: ts.length, late: ts.filter(isLate).length }; };
    const ppl = people();
    return topBar(t('orgTitle'), `<button type="button" class="btn sm primary" data-act="unitadd">${ic.plus}${t('addDept')}</button>`) + `<div class="layer-b">
      <p class="note" style="margin-top:0">${t('orgIntro')}</p>
      <div class="org-acts"><button type="button" class="btn sm" data-act="orgimport">${sic('clip', 15)}${t('orgImportBtn')}</button><button type="button" class="btn sm ghost" data-act="chat">${sic('spark', 15)}${t('orgByChat')}</button></div>
      ${ds.length ? ds.map(d => { const secs = secsOf(d.id), c = cnt([d.id, ...secs.map(sc => sc.id)]); return `<section class="org-card">
        <div class="org-h"><button type="button" class="org-name" data-unitf="${d.id}"><span class="mi">${ic.org}</span><b dir="auto">${esc(d.name)}</b></button>${c.n ? `<span class="pill ${c.late ? 'late' : 'prog'}">${plural(c.n, 'openTasks')}${c.late ? ` · ${t('nLate', { n: c.late })}` : ''}</span>` : ''}<button type="button" class="iconbtn sm" data-act="unitedit" data-u="${d.id}" aria-label="${esc(t('edit'))}">${sic('edit', 16)}</button></div>
        <div class="org-secs">${secs.map(sc => { const c2 = cnt([sc.id]); return `<div class="org-sec"><button type="button" class="grow linkbtn inl" data-unitf="${sc.id}" dir="auto">↳ ${esc(sc.name)}</button>${c2.n ? `<span class="num muted">${c2.n}</span>` : ''}<button type="button" class="iconbtn sm" data-act="unitedit" data-u="${sc.id}" aria-label="${esc(t('edit'))}">${sic('edit', 15)}</button></div>`; }).join('')}
          <button type="button" class="linkbtn inl" data-act="unitadd" data-parent="${d.id}">${sic('plus', 14)} ${t('addSec')}</button></div></section>`; }).join('')
      : emptyBox(t('orgEmptyTitle'), t('orgEmptyText'), `<button type="button" class="btn primary" data-act="unitadd">${ic.plus}${t('addDept')}</button>`)}
      ${ds.length && ppl.length ? `<div class="sec-h"><h3>${t('orgPeople')}</h3><span class="cnt">${ppl.filter(p => unitById(p.unit_id)).length}/${ppl.length}</span></div><p class="note" style="margin-top:0">${t('orgPeopleHint')}</p>
        <div class="box org-ppl">${ppl.map(p => `<label class="org-p"><span class="av sm">${esc(initials(p.name))}</span><span class="grow" dir="auto">${esc(p.name)}</span>${unitSelect('pu_' + p.id, unitById(p.unit_id) ? p.unit_id : '', t('noUnitShort')).replace('<select ', `<select data-pu="${p.id}" `)}</label>`).join('')}</div>` : ''}
    </div>`;
  });
}
function openUnitForm(u, parentId) {
  const isNew = !u, kind = u ? u.kind : parentId ? 'sec' : 'dept';
  const ttl = isNew ? t(kind === 'sec' ? 'addSec' : 'addDept') : t(kind === 'sec' ? 'editSec' : 'editDept');
  openSheet(ttl, `<form id="uForm" class="form" autocomplete="off">
      <label class="fld"><span>${t(kind === 'sec' ? 'secName' : 'deptName')}</span><input name="name" required maxlength="150" dir="auto" value="${esc(u?.name || '')}" placeholder="${esc(t(kind === 'sec' ? 'secPh' : 'deptPh'))}" autofocus></label>
      ${kind === 'sec' ? `<label class="fld"><span>${t('secOf')}</span><select name="parent">${depts().map(d => `<option value="${d.id}" ${(u?.parent_id || parentId) === d.id ? 'selected' : ''}>${esc(d.name)}</option>`).join('')}</select></label>` : ''}
      ${isNew && kind === 'dept' ? `<label class="fld"><span>${t('deptSecs')} <em>${t('optional')}</em></span><textarea name="secs" rows="3" dir="auto" placeholder="${esc(t('deptSecsPh'))}"></textarea></label>` : ''}
      <p class="note warn-txt" id="uErr" hidden>${t('unitExists')}</p>
      <div class="save-row"><button class="btn primary block" type="submit">${ic.check}${t('save')}</button></div>
      ${isNew ? '' : `<div class="danger-row"><button type="button" class="btn sm ghost danger" id="uDel">${sic('archive', 16)}${t('unitRemove')}</button></div>`}</form>`, {
    onMount: sh => {
      const f = sh.querySelector('#uForm');
      f.onsubmit = e => {
        e.preventDefault();
        const name = String(f.elements.name.value || '').replace(/\s+/g, ' ').trim(); if (!name) return;
        const parent = kind === 'sec' ? f.elements.parent.value : null;
        if (units().some(x => x.id !== u?.id && x.kind === kind && (kind === 'dept' || x.parent_id === parent) && normAr(x.name) === normAr(name))) { sh.querySelector('#uErr').hidden = false; return; }
        const tg = u ? unitById(u.id) : newUnit({ kind, parent_id: parent, sort: units().length });
        tg.name = name.slice(0, 150); if (kind === 'sec') tg.parent_id = parent; putUnit(tg);
        if (isNew && kind === 'dept') addSecs(tg.id, String(f.elements.secs?.value || '').split(/[\n،,؛;]+/));
        closeSheet(true); toast(t('saved'));
      };
      sh.querySelector('#uDel')?.addEventListener('click', () => confirmSheet(t(kind === 'dept' ? 'deptRemoveQ' : 'secRemoveQ', { n: esc(u.name) }), t('unitRemove'), () => {
        const tg = unitById(u.id); if (!tg) return; tg.deleted = true; putUnit(tg);
        if (kind === 'dept') secsOf(tg.id).forEach(sc => { const y = unitById(sc.id); y.deleted = true; putUnit(y); });
        if (UI.unitF === tg.id) UI.unitF = null; closeSheet(true); toast(t('unitRemoved'));
      }));
    }
  });
  sheetState.guard = true;
}
/** Adds sections under a department, skipping blanks and names it already has. Returns how many were added. */
function addSecs(deptId, names) {
  let n = 0;
  names.map(x => String(x || '').replace(/\s+/g, ' ').trim()).filter(Boolean).forEach(nm => {
    if (secsOf(deptId).some(sc => normAr(sc.name) === normAr(nm))) return;
    putUnit(newUnit({ kind: 'sec', parent_id: deptId, name: nm.slice(0, 150), sort: units().length })); n++;
  });
  return n;
}
/** A department by name (created when missing). */
function ensureDept(name) {
  const nm = String(name || '').replace(/\s+/g, ' ').trim(); if (!nm) return null;
  const hit = depts().find(d => normAr(d.name) === normAr(nm)); if (hit) return hit.id;
  const u = newUnit({ kind: 'dept', name: nm.slice(0, 150), sort: units().length }); putUnit(u); return u.id;
}
/* ---------- Excel import: a tasks sheet (e.g. exported from the company portal) or a departments sheet ----------
   Read on the device; nothing is saved before the review screen is approved. Re-importing the same portal sheet updates
   the tasks it created (by the portal's task number) and never overwrites a field the user changed in Rafeeq since. */
const HDR = [
  ['ref', /^(رقم المهمه|رقم|الرقم|م|مسلسل|id|#|no|task id|task no)$/],
  ['due', /(تاريخ الاستحقاق|الاستحقاق|تاريخ التسليم|موعد التسليم|الموعد النهائي|^الموعد$|due|deadline)/],
  ['start', /(تاريخ البدء|تاريخ البدايه|start)/],
  ['created', /(تاريخ الانشاء|created)/],
  ['closed', /(تاريخ الاغلاق|تاريخ الانجاز|closed|completed on|completion)/],
  ['pct', /(نسبه الانجاز|نسبه|progress|percent|%)/],
  ['title', /(اسم المهمه|^المهمه$|^المهام$|عنوان|الموضوع|^task$|task name|^title$|subject|^المطلوب$)/],
  ['desc', /(الوصف|التفاصيل|description|details)/],
  ['status', /(الحاله|status)/],
  ['prio', /(الاولويه|priority)/],
  ['assignees', /(المسند لهم|المسند اليه|المسند|المكلف|المسؤول|assignee|assigned)/],
  ['owner', /(المالك|المنشئ|صاحب المهمه|owner|created by|reporter)/],
  ['project', /(المشروع|project)/],
  ['channel', /^(المصدر|مصدر|source|channel)$/],
  ['tags', /(الوسوم|الوسم|tags|labels)/],
  ['dept', /^(الاداره|اداره|الادارات|department|dept)$/],
  ['sec', /^(القسم|قسم|الاقسام|section)$/],
  ['person', /^(الموظف|الموظفين|الاسم|اسم الموظف|الشخص|employee|name|person)$/],
];
function sheetCols(rows) {
  for (let r = 0; r < Math.min(rows.length, 6); r++) {
    const cols = {}; let hits = 0;
    rows[r].forEach((h, i) => { const n = normAr(String(h || '')).replace(/[:*]/g, '').replace(/\s+/g, ' ').trim(); if (!n) return; const k = HDR.find(([key, rx]) => cols[key] == null && rx.test(n)); if (k) { cols[k[0]] = i; hits++; } });
    if (hits >= 2 || cols.title != null || cols.dept != null) return { cols, head: r };
  }
  return null;
}
const cleanHTML = v => String(v ?? '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|li)>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\*\*/g, '').replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n').trim();
function cellDate(v) {
  if (typeof v === 'number') return serialDate(v);
  const s2 = String(v || '').trim(); let m;
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s2))) return validDate(`${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`);
  if ((m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/.exec(s2))) return validDate(`${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`);
  return null;
}
const splitNames = v => String(v || '').split(/[,،;؛\n]+/).map(x => x.replace(/\s+/g, ' ').trim()).filter(Boolean);
const PORTAL_ST = [[/(مكتمل|منجز|منتهي|مغلق|اغلق|تم|done|complete|closed)/, 'done'], [/(ملغ|cancel)/, 'cancelled'], [/(مراجع|review)/, 'review'], [/(قيد التنفيذ|جار|جاري|in progress|progress|doing)/, 'prog'], [/(معلق|متوقف|hold|paused)/, 'hold'], [/(انتظار|بانتظار|waiting)/, 'prog'], [/(جديد|لم تبدا|لم يبدا|new|todo|open|not started)/, 'todo']];
function portalStatus(v, pct) { const n = normAr(String(v || '')); if (/(متاخر|overdue|late)/.test(n)) return pct > 0 ? 'prog' : 'todo'; const hit = PORTAL_ST.find(([rx]) => rx.test(n)); return hit ? hit[1] : pct >= 100 ? 'review' : pct > 0 ? 'prog' : 'todo'; }
function portalPrio(v) { const n = normAr(String(v || '')); return /(عاجل|حرج|urgent|critical)/.test(n) ? ['hi', true] : /(عالي|مرتفع|high)/.test(n) ? ['hi', false] : /(منخفض|low)/.test(n) ? ['lo', false] : ['mid', false]; }
/** A person from a sheet: the exact contact, or one whose name is contained in it («مأمون الصقور» ⊂ «مامون محمود هزاع الصقور»). */
function matchPersonLoose(name) {
  const id = matchPerson(name); if (id) return id;
  const tk = nameTokens(name); if (tk.length < 2) return null;
  const hits = people().filter(p => { const pt = nameTokens(p.name); return pt.length >= 2 && pt[0] === tk[0] && pt.every(w => tk.includes(w)); });
  return hits.length === 1 ? hits[0].id : null;
}
/** Department names written in a portal tag («إدارة تقنية المعلومات CEO» → «إدارة تقنية المعلومات»). */
const tagUnits = v => splitNames(v).map(x => x.replace(/\s+(ceo|coo|cfo|[A-Z]{2,5})$/i, '').trim()).filter(x => /^(اداره|قسم|قطاع|department|section)\s/.test(normAr(x))); // «قرار الإدارة التنفيذية» is a tag, not a department
function sheetKind(rows) { const sc = sheetCols(rows); if (!sc) return null; const c = sc.cols; return c.title != null ? 'tasks' : c.dept != null || c.sec != null ? 'org' : null; }

async function importSheetFile(file, from = 'chat') {
  let sh; try { sh = await readSheet(file); } catch (e) { const m = String(e?.message || ''); toast(t(m === 'unsupported' ? 'impUnsupported' : m === 'old_browser' ? 'impOldBrowser' : m === 'empty' ? 'impEmpty' : 'impBad'), null, 'err', 6000); return; }
  const kind = sheetKind(sh.rows);
  if (!kind) { const heads = (sh.rows[0] || []).filter(Boolean).slice(0, 8).join('، '); if ($('#chatLog')) { CH.msgs.push({ who: 'bot', html: `<span>${t('impUnknown', { h: esc(heads) })}</span>` }); renderChat(); } else toast(t('impUnknown', { h: heads }), null, 'err', 8000); return; }
  if (kind === 'org' && !orgOn()) { toast(t('orgNeedsDb'), null, 'err', 6000); return; }
  if ($('#chatLog')) { CH.msgs.push({ who: 'me', html: `<span class="filemsg">${sic('clip', 14)} <span dir="auto">${esc(file.name)}</span></span>` }); CH.msgs.push({ who: 'bot', html: `<span>${t(kind === 'tasks' ? 'impReadTasks' : 'impReadOrg', { n: sh.rows.length - 1 })}</span>` }); renderChat(); }
  const fromChat = !!$('#chatLog'); if (fromChat) closeSheet(true); // the review screen opens full-size; the chat comes back with the result
  const lr = kind === 'tasks' ? openTaskImport(sh, file.name) : openOrgImport(sh, file.name);
  if (lr && fromChat) lr.onPop = () => openChat();
}

/* --- tasks --- */
function parseTaskRows(sh) {
  const { cols: c, head } = sheetCols(sh.rows); const get = (r, k) => c[k] != null ? r[c[k]] : '';
  return sh.rows.slice(head + 1).map((r, i) => {
    const title = cleanHTML(get(r, 'title')).replace(/\s+/g, ' ').slice(0, 300); if (!title) return null;
    const pctRaw = get(r, 'pct'); const pct = pctRaw === '' ? null : Math.max(0, Math.min(100, Math.round(typeof pctRaw === 'number' && pctRaw <= 1 && String(pctRaw).includes('.') ? pctRaw * 100 : parseFloat(String(pctRaw).replace('%', '')) || 0)));
    const desc = cleanHTML(get(r, 'desc')); const [prio, urgent] = portalPrio(get(r, 'prio'));
    const project = String(get(r, 'project') || '').trim();
    return { row: i, ref: String(get(r, 'ref') ?? '').trim().slice(0, 100), title, details: desc && normAr(desc) !== normAr(title) ? desc.slice(0, 5000) : '', status: portalStatus(get(r, 'status'), pct || 0), statusTxt: String(get(r, 'status') || '').trim(), priority: prio, urgent,
      due: cellDate(get(r, 'due')), closed: cellDate(get(r, 'closed')), progress: c.pct != null ? pct : null, project: /^(غير محدد|بدون|none|-|n\/a)?$/i.test(normAr(project)) ? '' : project.slice(0, 80),
      channel: String(get(r, 'channel') || '').trim().slice(0, 60), units: [...tagUnits(get(r, 'tags')), ...[get(r, 'dept'), get(r, 'sec')].map(x => String(x || '').trim()).filter(Boolean)], owner: String(get(r, 'owner') || '').replace(/\s+/g, ' ').trim(), assignees: splitNames(get(r, 'assignees')) };
  }).filter(Boolean);
}
/** Who the user is in the sheet: a saved answer, else the name that matches the profile name, else the most frequent name. */
function guessMe(items) {
  const freq = new Map(); items.forEach(it => [it.owner, ...it.assignees].filter(Boolean).forEach(n => freq.set(n, (freq.get(n) || 0) + 1)));
  const names = [...freq.keys()]; const saved = (state.profile.settings || {}).portalName;
  if (saved && names.some(n => normAr(n) === normAr(saved))) return { me: names.find(n => normAr(n) === normAr(saved)), names, freq };
  const mine = nameTokens(state.profile.display_name || '');
  const scored = names.map(n => { const tk = nameTokens(n); const hit = mine.filter(w => tk.includes(w)).length; return [n, hit && tk[0] === mine[0] ? hit + 1 : hit, freq.get(n)]; }).sort((a, b) => b[1] - a[1] || b[2] - a[2]);
  const top = scored[0];
  const me = top && top[1] >= 2 ? top[0] : names.sort((a, b) => freq.get(b) - freq.get(a)).find(n => freq.get(n) >= Math.max(2, items.length * 0.5)) || '';
  return { me, names: names.sort((a, b) => freq.get(b) - freq.get(a)), freq };
}
/** What one sheet row becomes, given who the user is. */
function planTask(it, me) {
  const isMe = n => !!me && normAr(n) === normAr(me);
  const mine = it.assignees.some(isMe), byMe = isMe(it.owner) || (!it.owner && mine);
  const others = it.assignees.filter(n => !isMe(n));
  let role = byMe ? (mine ? (others.length ? 'both' : 'exec') : 'follow') : mine ? 'exec' : 'follow';
  const fus = role === 'exec' ? [] : others.slice(0, 12);
  if (!fus.length && role !== 'exec') role = 'exec'; // nobody to follow up with
  const source = byMe ? SELF : it.owner ? srcName(it.owner) : null;
  let unit = null, newUnitName = '';
  for (const u of it.units) { const id = matchUnit(u, true) || depts().find(d => normAr(d.name) === normAr(u))?.id; if (id) { unit = id; break; } if (!newUnitName) newUnitName = u; }
  if (unit) newUnitName = '';
  if (!unit && !newUnitName) { const ids = [...new Set(fus.map(n => unitOfPerson(matchPersonLoose(n))).filter(Boolean))]; if (ids.length === 1) unit = ids[0]; }
  if (!unit && !newUnitName) unit = matchUnit(it.title + ' ' + it.details.slice(0, 300), true);
  const cos = role === 'exec' ? others : [];
  return { ...it, role, fus, source, unit, newUnitName, cos };
}
const TRACK = ['title', 'status', 'due', 'priority', 'progress'];
function importDiff(x, p) {
  const was = x.ai?.portal || {}; const ch = []; const kept = [];
  TRACK.forEach(k => { const nv = k === 'status' ? (p.status === 'done' ? 'done' : p.status) : p[k] ?? null; const cur = x[k] ?? null; if (nv === (was[k] ?? null) || nv === cur) return; if ((was[k] ?? null) !== cur && was[k] !== undefined) kept.push(k); else ch.push(k); });
  return { ch, kept };
}
function openTaskImport(sh, fname) {
  const items = parseTaskRows(sh); if (!items.length) { toast(t('impEmpty'), null, 'err'); return null; }
  const g = guessMe(items);
  const S = { me: g.me, names: g.names, sel: new Set(), newUnits: new Set(), grp: 'unit', fname };
  const byRef = new Map(live().filter(x => x.ext_ref).map(x => [x.ext_ref, x]));
  const dupOf = p => p.ref ? byRef.get(p.ref) : live().find(x => !x.ext_ref && normAr(x.title) === normAr(p.title) && (x.due || null) === (p.due || null));
  const plan = () => items.map(it => { const p = planTask(it, S.me); const ex = dupOf(p); const d = ex ? importDiff(ex, p) : null; return { ...p, ex, d, kind: !ex ? 'new' : d.ch.length ? 'upd' : 'same' }; });
  let P = plan(); P.forEach((p, i) => { if (p.kind !== 'same') S.sel.add(i); if (p.newUnitName) S.newUnits.add(p.newUnitName); });
  S.layer = pushLayer(() => {
    const cN = P.filter(p => p.kind === 'new').length, cU = P.filter(p => p.kind === 'upd').length, cS = P.filter(p => p.kind === 'same').length;
    const newU = [...new Set(P.map(p => p.newUnitName).filter(Boolean))];
    const gk = p => S.grp === 'status' ? p.status : (p.unit ? deptOf(p.unit)?.id : p.newUnitName && S.newUnits.has(p.newUnitName) ? 'new:' + p.newUnitName : '-');
    const gl = k => S.grp === 'status' ? t(stKey(k)) : k === '-' ? t('noUnit') : k.startsWith('new:') ? `${k.slice(4)} <em class="newb">${t('newBadge')}</em>` : esc(unitById(k)?.name || '');
    const groups = new Map(); P.forEach((p, i) => { const k = gk(p); if (!groups.has(k)) groups.set(k, []); groups.get(k).push([p, i]); });
    const row = ([p, i]) => `<label class="imp-row ${p.kind}"><input type="checkbox" data-imp="${i}" ${S.sel.has(i) ? 'checked' : ''}><span class="body"><span class="t" dir="auto">${esc(p.title)}</span><span class="meta"><span class="pill ${{ prog: 'prog', wait: 'wait', review: 'review', done: 'done' }[p.status] || ''}">${t(stKey(p.status))}</span>${p.due ? `<span class="${p.status !== 'done' && p.status !== 'review' && p.due < T() ? 'late-txt' : ''}">${fmtShort(p.due)}</span>` : ''}${p.priority === 'hi' ? `<span class="pill hi">${t(p.urgent ? 'urgentPill' : 'prioHi')}</span>` : ''}${p.progress && p.status !== 'done' ? `<span class="num">${p.progress}%</span>` : ''}<span>${t('role_' + p.role)}</span>${p.fus.length ? `<span dir="auto">${sic('users', 12)} ${esc(p.fus.slice(0, 2).map(n => n.split(' ')[0]).join('، '))}${p.fus.length > 2 ? ` +${p.fus.length - 2}` : ''}</span>` : ''}${p.source && p.source !== SELF ? `<span dir="auto">${t('srcShort')}: ${esc(p.source)}</span>` : ''}${p.unit && S.grp === 'status' ? `<span>${esc(unitLabel(p.unit, true))}</span>` : p.unit && deptOf(p.unit)?.id !== p.unit ? `<span>↳ ${esc(unitLabel(p.unit, true))}</span>` : ''}</span>
      ${p.kind === 'upd' ? `<span class="imp-upd">${t('impWillUpdate', { f: p.d.ch.map(k => t('impF_' + k)).join('، ') })}${p.d.kept.length ? ` · ${t('impKept', { f: p.d.kept.map(k => t('impF_' + k)).join('، ') })}` : ''}</span>` : p.kind === 'same' ? `<span class="imp-upd muted">${t('impSame')}</span>` : ''}</span><span class="badge ${p.kind}">${t('impK_' + p.kind)}</span></label>`;
    return topBar(t('impTitle')) + `<div class="layer-b imp">
      <p class="note" style="margin-top:0" dir="auto">${sic('clip', 14)} ${esc(S.fname)} · ${t('impCounts', { a: cN, b: cU, c: cS })}</p>
      <div class="box imp-me"><label class="fld"><span>${t('impWhoMe')}</span><select id="impMe"><option value="">${t('impMeNone')}</option>${S.names.slice(0, 40).map(n => `<option ${normAr(n) === normAr(S.me) ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></label><small class="note">${t('impWhoMeHint')}</small></div>
      ${newU.length && orgOn() ? `<div class="box"><p class="impq">${t('impNewUnits')}</p>${newU.map(n => `<label class="imp-nu"><input type="checkbox" data-nu="${esc(n)}" ${S.newUnits.has(n) ? 'checked' : ''}> <span dir="auto">${esc(n)}</span></label>`).join('')}</div>` : ''}
      <div class="imp-bar"><div class="seg two" role="tablist">${[['unit', 'impByUnit'], ['status', 'impByStatus']].map(([k, l]) => `<button type="button" class="${S.grp === k ? 'on' : ''}" data-impgrp="${k}">${t(l)}</button>`).join('')}</div><button type="button" class="linkbtn" data-impall="1">${S.sel.size === P.length ? t('impNone') : t('impAll')}</button></div>
      ${[...groups].map(([k, rs]) => `<section class="imp-g"><div class="sec-h"><h3>${gl(k)}</h3><span class="cnt">${rs.length}</span></div><div class="box imp-list">${rs.map(row).join('')}</div></section>`).join('')}
      <div class="save-row sticky"><button type="button" class="btn primary block" data-act="impgo" ${S.sel.size ? '' : 'disabled'}>${ic.check}${t('impGo', { n: S.sel.size })}</button></div>
    </div>`;
  });
  const lr = S.layer;
  lr.el.addEventListener('change', e => {
    const tg = e.target;
    if (tg.dataset.imp != null) { const i = +tg.dataset.imp; if (tg.checked) S.sel.add(i); else S.sel.delete(i); return lr.rerender(); }
    if (tg.dataset.nu != null) { if (tg.checked) S.newUnits.add(tg.dataset.nu); else S.newUnits.delete(tg.dataset.nu); return lr.rerender(); }
    if (tg.id === 'impMe') { S.me = tg.value; P = plan(); return lr.rerender(); }
  });
  lr.el.addEventListener('click', e => {
    const b = e.target.closest('[data-impgrp],[data-impall],[data-act=impgo]'); if (!b) return; e.stopPropagation();
    if (b.dataset.impgrp) { S.grp = b.dataset.impgrp; return lr.rerender(); }
    if (b.dataset.impall) { if (S.sel.size === P.length) S.sel.clear(); else P.forEach((_, i) => S.sel.add(i)); return lr.rerender(); }
    if (b.dataset.act === 'impgo') return runTaskImport(P, S);
  });
  return lr;
}
function runTaskImport(P, S) {
  if (S.me) { const st = Object.assign({}, state.profile.settings); if (st.portalName !== S.me) { st.portalName = S.me; saveProfile({ settings: st }); } }
  const made = { n: 0, u: 0, units: 0 }; const unitIds = new Map();
  if (orgOn()) [...S.newUnits].forEach(n => { const before = depts().length; const id = ensureDept(n); if (id) { unitIds.set(n, id); if (depts().length > before) made.units++; } });
  [...S.sel].sort((a, b) => a - b).forEach(i => {
    const p = P[i]; if (!p) return;
    const unit = orgOn() ? (p.unit || unitIds.get(p.newUnitName) || null) : null;
    const snap = { title: p.title, status: p.status, due: p.due, priority: p.priority, progress: p.progress };
    if (p.ex) {
      const x = taskById(p.ex.id); if (!x) return;
      p.d.ch.forEach(k => { if (k === 'status') { if (x.status !== p.status) { addLog(x, 'status', '', { from: x.status, to: p.status }); x.status = p.status; if (p.status === 'done') x.completed_on = p.closed || T(); else x.completed_on = null; } } else if (k === 'due') { addLog(x, 'due', '', { from: x.due, to: p.due }); x.due = p.due; } else x[k] = p[k]; });
      if (unit && !x.unit_id) x.unit_id = unit;
      x.ai = { ...(x.ai || {}), portal: snap }; addLog(x, 'imported', S.fname.slice(0, 80)); putTask(x); made.u++; return;
    }
    const notes = [p.urgent ? t('impUrgentNote') : '', p.channel ? t('impChannelNote', { c: p.channel }) : '', p.cos.length ? t('impCoNote', { n: p.cos.join('، ') }) : ''].filter(Boolean).join('\n');
    const x = newTask({ title: p.title, details: p.details, status: p.status, priority: p.priority, role: p.role, due: p.due, source: p.source || '', project: p.project, notes, progress: p.progress, ext_ref: p.ref, unit_id: unit, completed_on: p.status === 'done' ? (p.closed || T()) : null, ai: { portal: snap } });
    if (x.status === 'wait') x.status = 'prog'; // the sheet does not say on whom
    addLog(x, 'imported', S.fname.slice(0, 80));
    const due = x.due && x.due > T() ? (diffDays(x.due, T()) > 1 ? addWorkdays(x.due, -1) : x.due) : addWorkdays(T(), 1);
    p.fus.forEach(n => { const pid = matchPersonLoose(n) || getOrCreatePerson(n); if (!pid || x.followups.some(f => f.person_id === pid)) return; x.followups.push({ id: uuid(), person_id: pid, what: '', due, status: isOpen(x) ? 'open' : 'done', created_at: nowISO(), closed_on: isOpen(x) ? null : T(), log: [] }); });
    putTask(x); made.n++;
  });
  popLayer();
  const msg = t('impDone', { a: made.n, b: made.u }) + (made.units ? ' ' + t('impDoneUnits', { n: made.units }) : '');
  toast(msg);
  { CH.msgs.push({ who: 'bot', html: `${sic('check', 16)} <span>${msg}</span><div class="chips"><button type="button" class="chip" data-tab="tasks">${t('navTasks')}</button>${orgOn() && depts().length ? `<button type="button" class="chip" data-act="org">${t('orgTitle')}</button>` : ''}</div>` }); renderChat(); }
}

/* --- departments sheet: columns Department / Section / (optional) Person --- */
function openOrgImport(sh, fname) {
  const { cols: c, head } = sheetCols(sh.rows); const get = (r, k) => c[k] != null ? String(r[c[k]] ?? '').replace(/\s+/g, ' ').trim() : '';
  let lastD = '';
  const rows = sh.rows.slice(head + 1).map(r => { let d = get(r, 'dept'); const sc = get(r, 'sec'), pn = get(r, 'person'); if (!d && sc) d = lastD; if (d) lastD = d; return d ? { d, sc, pn } : null; }).filter(Boolean);
  if (!rows.length) { toast(t('impEmpty'), null, 'err'); return null; }
  const tree = new Map(); rows.forEach(r => { if (!tree.has(r.d)) tree.set(r.d, { secs: new Set(), ppl: [] }); const n = tree.get(r.d); if (r.sc) n.secs.add(r.sc); if (r.pn) splitNames(r.pn).forEach(x => n.ppl.push([x, r.sc])); });
  const exD = d => depts().find(x => normAr(x.name) === normAr(d));
  pushLayer(() => topBar(t('impOrgTitle')) + `<div class="layer-b imp">
    <p class="note" style="margin-top:0" dir="auto">${sic('clip', 14)} ${esc(fname)}</p>
    ${[...tree].map(([d, n]) => { const ed = exD(d); return `<section class="org-card"><div class="org-h"><span class="mi">${ic.org}</span><b class="grow" dir="auto">${esc(d)}</b><span class="badge ${ed ? 'same' : 'new'}">${t(ed ? 'impK_same' : 'impK_new')}</span></div>
      ${n.secs.size ? `<div class="org-secs">${[...n.secs].map(sc => `<div class="org-sec"><span class="grow" dir="auto">↳ ${esc(sc)}</span>${ed && secsOf(ed.id).some(x => normAr(x.name) === normAr(sc)) ? '' : `<span class="badge new">${t('impK_new')}</span>`}</div>`).join('')}</div>` : ''}
      ${n.ppl.length ? `<p class="note" dir="auto">${sic('users', 13)} ${esc(n.ppl.map(([pn]) => pn).slice(0, 12).join('، '))}${n.ppl.length > 12 ? ` +${n.ppl.length - 12}` : ''}</p>` : ''}</section>`; }).join('')}
    <div class="save-row sticky"><button type="button" class="btn primary block" id="orgGo">${ic.check}${t('impOrgGo', { d: tree.size })}</button></div></div>`);
  const lr = UI.layers[UI.layers.length - 1];
  lr.el.addEventListener('click', e => {
    if (!e.target.closest('#orgGo')) return; e.stopPropagation();
    let nd = 0, ns = 0, np = 0;
    tree.forEach((n, d) => {
      const before = depts().length; const did = ensureDept(d); if (depts().length > before) nd++;
      ns += addSecs(did, [...n.secs]);
      n.ppl.forEach(([pn, sc]) => { const pid = matchPersonLoose(pn) || getOrCreatePerson(pn); const p = personById(pid); if (!p) return; const sid = sc ? secsOf(did).find(x => normAr(x.name) === normAr(sc))?.id : null; p.unit_id = sid || did; putPerson(p); np++; });
    });
    popLayer(); const msg = t('impOrgDone', { d: nd, s: ns, p: np }); toast(msg);
    CH.msgs.push({ who: 'bot', html: `${sic('check', 16)} <span>${msg}</span>` }); if ($('#chatLog')) renderChat();
  });
  return lr;
}
function pickSheet() {
  let inp = $('#sheetFile'); if (!inp) { inp = document.createElement('input'); inp.type = 'file'; inp.id = 'sheetFile'; inp.hidden = true; inp.accept = '.xlsx,.xlsm,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv'; document.body.appendChild(inp); inp.addEventListener('change', () => { const f = inp.files?.[0]; inp.value = ''; if (f) importSheetFile(f); }); }
  inp.click();
}

function openPersonForm(p) {
  const isNew = !p;
  openSheet(isNew ? t('addPerson') : t('editPerson'), `<form id="pForm" class="form" autocomplete="off">
      <label class="fld"><span>${t('pName')}</span><input name="name" required maxlength="120" dir="auto" value="${esc(p?.name || '')}" autofocus></label>
      <label class="fld"><span>${t('pOrg')} <em>${t('optional')}</em></span><input name="org" dir="auto" value="${esc(p?.org || '')}" placeholder="${esc(t('pOrgPh'))}"></label>
      ${orgOn() && depts().length ? `<label class="fld"><span>${sic('org', 14)} ${t('fUnit')} <em>${t('optional')}</em></span>${unitSelect('unit', p?.unit_id || '')}<small class="note">${t('pUnitHint')}</small></label>` : ''}
      <label class="fld"><span>${t('pContact')} <em>${t('optional')}</em></span><input name="contact" dir="auto" value="${esc(p?.contact || '')}" placeholder="${esc(t('pContactPh'))}"></label>
      <label class="fld"><span>${t('pNotes')} <em>${t('optional')}</em></span><textarea name="notes" rows="2" dir="auto">${esc(p?.notes || '')}</textarea></label>
      <p class="note warn-txt" id="pErr" hidden>${t('nameExists')}</p>
      <div class="save-row"><button class="btn primary block" type="submit">${ic.check}${t('save')}</button></div></form>`, {
    onMount: sh => sh.querySelector('#pForm').onsubmit = e => {
      e.preventDefault(); const fd = new FormData(e.target); const name = (fd.get('name') || '').trim().replace(/\s+/g, ' '); if (!name) return;
      const key = nameTokens(name).join(' ');
      if (people().some(q => q.id !== p?.id && nameTokens(q.name).join(' ') === key)) { sh.querySelector('#pErr').hidden = false; return; }
      const tg = (p && (personById(p.id) || p)) || newPerson(name);
      tg.name = name; tg.org = (fd.get('org') || '').trim(); tg.contact = (fd.get('contact') || '').trim(); tg.notes = (fd.get('notes') || '').trim();
      if (e.target.elements.unit) tg.unit_id = unitById(e.target.elements.unit.value) ? e.target.elements.unit.value : null;
      putPerson(tg); closeSheet(true); toast(t('saved')); if (isNew) openPerson(tg.id);
    }
  });
  sheetState.guard = true;
}
function openProfile() {
  const p = state.profile;
  openSheet(t('profile'), `<form id="prForm" class="form" autocomplete="off">
      <label class="fld"><span>${t('yourName')}</span><input name="n" required maxlength="80" dir="auto" value="${esc(p.display_name)}" autofocus></label>
      <label class="fld"><span>${t('jobTitle')} <em>${t('optional')}</em></span><input name="j" maxlength="80" dir="auto" value="${esc(p.job_title)}" placeholder="${esc(t('jobTitlePh'))}"></label>
      <div class="save-row"><button class="btn primary block" type="submit">${ic.check}${t('save')}</button></div></form>`, {
    onMount: sh => sh.querySelector('#prForm').onsubmit = e => { e.preventDefault(); const fd = new FormData(e.target); saveProfile({ display_name: (fd.get('n') || '').trim(), job_title: (fd.get('j') || '').trim() }); closeSheet(true); toast(t('saved')); }
  });
}

/* ================= layers ================= */
function pushLayer(build) {
  const el = document.createElement('div');
  el.className = 'layer'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-modal', 'true');
  const Lr = { el, rerender: () => {
    const sc = el.querySelector('.layer-b')?.scrollTop || 0;
    const kept = [...el.querySelectorAll('[data-keep]')];
    const typed = [...el.querySelectorAll('input[name], textarea[name]')].filter(i => i.value && i.type !== 'hidden').map(i => ({ name: i.name, value: i.value }));
    const act = document.activeElement, focusName = act && el.contains(act) && act.name ? act.name : null, pos = focusName ? act.selectionStart : null;
    el.innerHTML = build();
    kept.forEach(k => { const n = el.querySelector(`[data-keep="${k.dataset.keep}"]`); if (n) n.replaceWith(k); });
    typed.forEach(x => { const n = el.querySelector(`[name="${x.name}"]`); if (n && !n.value) n.value = x.value; });
    if (focusName) { const n = el.querySelector(`[name="${focusName}"]`); if (n) { n.focus({ preventScroll: true }); try { if (pos != null) n.setSelectionRange(pos, pos); } catch { } } }
    const b = el.querySelector('.layer-b'); if (b) b.scrollTop = sc;
  } };
  Lr.rerender();
  Lr.opener = document.activeElement; // restored on close for keyboard / screen-reader users
  $('#layers').appendChild(el); UI.layers.push(Lr);
  $('#layerBackdrop').classList.add('on');
  el.tabIndex = -1; el.focus({ preventScroll: true });
  requestAnimationFrame(() => requestAnimationFrame(() => el.classList.add('on')));
  return Lr;
}
function popLayer() {
  const Lr = UI.layers.pop(); if (!Lr) return;
  if (Lr.onPop) setTimeout(Lr.onPop, 0);
  Lr.el.classList.remove('on'); setTimeout(() => Lr.el.remove(), 280);
  if (!UI.layers.length) $('#layerBackdrop').classList.remove('on');
  try { if (Lr.opener && Lr.opener.isConnected) Lr.opener.focus({ preventScroll: true }); } catch { }
}
function closeAllLayers() { while (UI.layers.length) { const Lr = UI.layers.pop(); Lr.el.remove(); } $('#layerBackdrop')?.classList.remove('on'); }
$('#layerBackdrop').addEventListener('click', () => popLayer());
const topBar = (title, right = '') => `<div class="layer-top"><button type="button" class="iconbtn flipx" data-act="pop" aria-label="${esc(t('back'))}">${ic.back}</button><div class="t">${title}</div>${right}</div>`;

function openTask(id) {
  pushLayer(() => {
    const x = taskById(id);
    if (!x || x.deleted) return topBar('') + `<div class="layer-b"><div class="empty small"><p>${t('notFound')}</p></div></div>`;
    if (isLog(x)) return logDetail(x);
    const open = isOpen(x);
    const pct = x.steps_total ? Math.round((x.steps_done / x.steps_total) * 100) : 0;
    const fus = (x.followups || []).slice().sort((a, b) => a.status === b.status ? ((a.due || '') < (b.due || '') ? -1 : 1) : a.status === 'open' ? -1 : 1);
    const logs = (x.log || []).slice().reverse();
    return topBar(t('taskDetails'), `<button type="button" class="iconbtn" data-act="edittask" data-id="${x.id}" aria-label="${esc(t('edit'))}">${ic.edit}</button>`) + `<div class="layer-b">
      <div class="chips top">${statusPill(x)}${x.priority === 'hi' ? `<span class="pill hi">${t('prioHi')}</span>` : ''}${x.project ? `<button type="button" class="pill projlink" data-proj="${esc(x.project)}" dir="auto">${sic('folder', 12)} ${esc(x.project)}</button>` : ''}${x.archived ? `<span class="pill">${t('archived')}</span>` : ''}${x.ai?.st === 'done' ? `<button type="button" class="pill ai" data-act="aiundo" data-id="${x.id}">${sic('spark', 13)}${t('aiPill')} · ${t('undo')}</button>` : ''}</div>
      <h2 class="d-title" dir="auto">${esc(x.title)}</h2>
      ${x.details ? `<p class="d-text" dir="auto">${esc(x.details)}</p>` : ''}
      ${x.status === 'wait' && x.waiting_on ? `<div class="waitbox">${sic('clock', 16)}<span>${t('waitingFrom', { w: t('w_' + (x.waiting_what || 'reply')), p: `<b dir="auto">${esc(pname(x.waiting_on))}</b>` })}</span></div>` : ''}
      ${x.status === 'done' ? `<div class="donebox"><b>${t('doneOn', { d: fmt(x.completed_on || T()) })}</b>${x.result ? `<p dir="auto">${esc(x.result)}</p>` : ''}</div>` : ''}
      <div class="kv num">
        <div><small>${t('fDue')}</small><b>${x.due ? `${wd(x.due)} ${fmtShort(x.due)}${x.due_time ? ` · <span class="num">${fmtTime(x.due_time)}</span>` : ''}` : t('noDue')}</b>${isLate(x) ? `<span class="pill late">${lateTxt(x.due)}</span>` : ''}</div>
        <div><small>${t('fRole')}</small><b>${t('role_' + x.role)}</b></div>
        ${orgOn() && (unitOf(x) || depts().length) ? `<div class="wide"><small>${t('fUnit')}</small>${unitOf(x) ? `<button type="button" class="linkbtn inl" data-unitf="${unitOf(x).id}" dir="auto"><b>${sic('org', 14)} ${esc(unitLabel(x.unit_id))}</b></button>` : `<button type="button" class="linkbtn inl" data-act="edittask" data-id="${x.id}">${t('addUnit')}</button>`}</div>` : ''}
        ${x.progress != null && isOpen(x) ? `<div class="wide"><small>${t('fProgress')}</small><b class="num">${x.progress}%</b><div class="progress"><i style="width:${x.progress}%"></i></div></div>` : ''}
        ${x.ext_ref ? `<div><small>${t('fExtRef')}</small><b class="num" dir="ltr">${esc(x.ext_ref)}</b></div>` : ''}
        ${x.steps_total ? `<div class="wide"><small>${t('stepsDone')}</small><div class="stepper"><button type="button" class="iconbtn sm" data-act="stepminus" data-id="${x.id}" aria-label="−">${ic.minus}</button><b>${t('stepsOf', { a: x.steps_done, b: x.steps_total })}</b><button type="button" class="iconbtn sm" data-act="stepplus" data-id="${x.id}" aria-label="+">${ic.plus}</button></div><div class="progress"><i style="width:${pct}%"></i></div></div>` : ''}
        ${x.recur ? `<div class="wide"><small>${t('fRecur')}</small><b>${sic('repeat', 14)} ${recurLabel(x.recur)}</b>${open ? ` <button type="button" class="linkbtn inl" data-act="skiprec" data-id="${x.id}">${t('skipOnce')}</button>` : ''}</div>` : ''}
        ${x.due_time && open ? `<div><small>${t('fRemind')}</small><b>${remindLabel(x.remind_min)}</b></div>` : ''}
        ${x.source ? `<div><small>${t('fSource')}</small><b dir="auto">${esc(srcLabel(x.source))}</b></div>` : `<div><small>${t('fSource')}</small><button type="button" class="linkbtn inl" data-act="edittask" data-id="${x.id}">${t('addSource')}</button></div>`}
      </div>
      ${open ? `<div class="save-row"><button type="button" class="btn primary block" data-act="complete" data-id="${x.id}">${ic.check}${t('complete')}</button><button type="button" class="btn block" data-act="postpone" data-id="${x.id}">${sic('later', 18)}${t('postpone')}</button></div>
        <div class="st-row">${OPEN_ST().map(s => `<button type="button" class="chip ${x.status === s ? 'on' : ''}" data-act="setst" data-id="${x.id}" data-v="${s}">${t(stKey(s))}</button>`).join('')}</div>` : `<div class="save-row"><button type="button" class="btn block" data-act="reopen" data-id="${x.id}">${sic('reopen', 18)}${t('reopen')}</button></div>`}
      <div class="sec-h"><h3>${t('followups')}</h3><span class="cnt">${openFus(x).length}</span><span class="grow"></span><button type="button" class="btn sm ghost" data-act="addfu" data-id="${x.id}">${sic('plus', 16)}${t('add')}</button></div>
      <div class="box">${fus.length ? fus.map(f => `<div class="fu ${f.status !== 'open' ? 'closed' : ''}"><div class="fu-top"><span class="av">${esc(initials(pname(f.person_id)))}</span><div class="body"><div class="t" dir="auto">${esc(pname(f.person_id))}</div><div class="meta">${f.what ? `<span dir="auto">${esc(f.what)}</span>` : ''}${f.status === 'open' ? `<span class="pill ${fuLate(f) ? 'late' : f.due <= T() ? 'wait' : ''}">${fuLate(f) ? lateTxt(f.due) : rel(f.due)}${f.time ? ' · ' + fmtTime(f.time) : ''}</span>` : `<span class="pill done">${t('fuClosed', { d: fmtShort(f.closed_on || T()) })}</span>`}</div></div>
          ${f.status === 'open' && open ? `<button type="button" class="iconbtn sm" data-act="remind" data-id="${x.id}" data-fu="${esc(f.id)}" aria-label="${esc(t('remind'))}" title="${esc(t('remind'))}">${sic('send', 17)}</button><button type="button" class="btn sm" data-act="furesult" data-id="${x.id}" data-fu="${esc(f.id)}">${t('logResult')}</button>` : ''}<button type="button" class="iconbtn sm" data-act="fumenu" data-id="${x.id}" data-fu="${esc(f.id)}" aria-label="${esc(t('edit'))}">${sic('edit', 16)}</button></div>
          ${(f.log || []).length ? `<div class="fu-log">${f.log.map(l => `<div><span class="when num">${fmtShort(l.date)}</span> <span dir="auto">${esc(fuLogText(l))}</span></div>`).join('')}</div>` : ''}</div>`).join('') : `<div class="fu muted">${t('noFollowups')}</div>`}</div>
      ${state.mode === 'cloud' ? (() => { const fs = filesOfTask(x.id); return `<div class="sec-h"><h3>${t('filesTitle')}</h3>${fs.length ? `<span class="cnt">${fs.length}</span>` : ''}<span class="grow"></span><button type="button" class="btn sm ghost" data-act="upload" data-task="${x.id}">${sic('clip', 16)}${t('attach')}</button></div>
      <div class="stack fstack">${uploadsBox({ task: x.id })}${fs.length ? fs.map(f => fileRow(f, { ctx: 'task' })).join('') : `<div class="box"><div class="fu muted">${t('noFilesTask')}</div></div>`}</div>`; })() : ''}
      <div class="sec-h"><h3>${t('progressLog')}</h3></div>
      <div class="box"><form class="upd" data-form="upd" data-id="${x.id}"><input name="u" placeholder="${esc(t('updatePh'))}" aria-label="${esc(t('addUpdate'))}" dir="auto"><button type="submit" class="btn primary sm">${t('add')}</button></form>
        <ul class="tl">${logs.map(l => `<li class="${l.kind === 'note' || l.text ? 'k' : ''}"><small class="num">${fmtShort(l.at.slice(0, 10))} · ${new Intl.DateTimeFormat(LOC(), { hour: 'numeric', minute: '2-digit' }).format(new Date(l.at))}</small>${l.kind === 'note' ? `<span dir="auto">${esc(l.text)}</span>` : `<span>${logText(l)}</span>${l.text ? `<span class="lt" dir="auto">${esc(l.text)}</span>` : ''}`}</li>`).join('') || `<li class="muted">${t('noLog')}</li>`}</ul></div>
      ${x.notes ? `<div class="sec-h"><h3>${t('fNotes')}</h3></div><div class="box"><p class="d-text" dir="auto" style="padding:12px 0;margin:0">${esc(x.notes)}</p></div>` : ''}
      <div class="danger-row"><button type="button" class="btn sm ghost" data-act="archive" data-id="${x.id}">${sic('archive', 16)}${x.archived ? t('unarchive') : t('archive')}</button><button type="button" class="btn sm ghost danger" data-act="deltask" data-id="${x.id}">${sic('trash', 16)}${t('delete')}</button></div>
    </div>`;
  });
}
function logDetail(x) {
  const fs = state.mode === 'cloud' ? filesOfTask(x.id) : [];
  return topBar(t('logDetails'), `<button type="button" class="iconbtn" data-act="edittask" data-id="${x.id}" aria-label="${esc(t('edit'))}">${ic.edit}</button>`) + `<div class="layer-b">
      <div class="chips top">${statusPill(x)}${x.project ? `<button type="button" class="pill projlink" data-proj="${esc(x.project)}" dir="auto">${sic('folder', 12)} ${esc(x.project)}</button>` : ''}</div>
      <h2 class="d-title" dir="auto">${esc(x.title)}</h2>
      ${x.details ? `<p class="d-text" dir="auto">${esc(x.details)}</p>` : ''}
      <div class="donebox"><b>${t('loggedOn', { d: `${wd(x.completed_on || T())} ${fmt(x.completed_on || T())}` })}</b></div>
      ${state.mode === 'cloud' ? `<div class="sec-h"><h3>${t('filesTitle')}</h3>${fs.length ? `<span class="cnt">${fs.length}</span>` : ''}<span class="grow"></span><button type="button" class="btn sm ghost" data-act="upload" data-task="${x.id}">${sic('clip', 16)}${t('attach')}</button></div>
      <div class="stack fstack">${uploadsBox({ task: x.id })}${fs.length ? fs.map(f => fileRow(f, { ctx: 'task' })).join('') : `<div class="box"><div class="fu muted">${t('noFilesLog')}</div></div>`}</div>` : ''}
      <div class="danger-row"><button type="button" class="btn sm ghost danger" data-act="deltask" data-id="${x.id}">${sic('trash', 16)}${t('delete')}</button></div>
    </div>`;
}
function openPerson(id) {
  pushLayer(() => {
    const p = personById(id);
    if (!p || p.deleted) return topBar('') + `<div class="layer-b"><div class="empty small"><p>${t('notFound')}</p></div></div>`;
    const openF = [], waits = [], hist = [], related = new Set();
    live().forEach(x => {
      (x.followups || []).forEach(f => { if (f.person_id !== id) return; related.add(x.id); if (f.status === 'open' && isOpen(x)) openF.push([x, f]); (f.log || []).forEach(l => hist.push({ x, f, l })); });
      if (x.waiting_on === id) { related.add(x.id); if (x.status === 'wait') waits.push(x); }
    });
    hist.sort((a, b) => (b.l.at || '') > (a.l.at || '') ? 1 : -1);
    const rel2 = sortTasks([...related].map(taskById).filter(x => x && isOpen(x)));
    return topBar(t('person'), `<button type="button" class="iconbtn" data-act="editperson" data-id="${p.id}" aria-label="${esc(t('edit'))}">${ic.edit}</button>`) + `<div class="layer-b">
      <div class="p-head"><span class="av xl">${esc(initials(p.name))}</span><div><h2 class="d-title" dir="auto" style="margin:0">${esc(p.name)}</h2>${p.org ? `<div class="sub" dir="auto">${esc(p.org)}</div>` : ''}${orgOn() && unitById(p.unit_id) ? `<button type="button" class="linkbtn inl sub" data-unitf="${p.unit_id}" dir="auto">${sic('org', 13)} ${esc(unitLabel(p.unit_id))}</button>` : ''}${p.contact ? `<div class="sub" dir="auto">${esc(p.contact)}</div>` : ''}</div></div>
      ${p.contact && /\d{7,}/.test(p.contact) ? `<div class="row" style="margin:-4px 0 12px"><a class="btn sm" href="tel:${esc(p.contact.replace(/[^\d+]/g, ''))}">${sic('phone', 16)}${t('call')}</a><a class="btn sm" href="https://wa.me/${esc(waNumber(p.contact))}" target="_blank" rel="noopener">${sic('send', 16)}WhatsApp</a></div>` : ''}
      ${p.notes ? `<p class="d-text" dir="auto">${esc(p.notes)}</p>` : ''}
      ${openF.length ? sec(t('openFollowups'), openF.length, openF.map(([x, f]) => fuRow(x, f)).join('')) : ''}
      ${waits.length ? sec(t('waitingOnThem'), waits.length, waits.map(x => taskCard(x, { noSwipe: true })).join('')) : ''}
      ${rel2.filter(x => !waits.includes(x)).length ? sec(t('relatedWork'), null, rel2.filter(x => !waits.includes(x)).map(x => taskCard(x, { noSwipe: true })).join('')) : ''}
      ${state.mode === 'cloud' ? (() => { const fs = filesOfPerson(p.id); return `<div class="sec-h"><h3>${t('filesTitle')}</h3>${fs.length ? `<span class="cnt">${fs.length}</span>` : ''}<span class="grow"></span><button type="button" class="btn sm ghost" data-act="upload" data-person="${p.id}">${sic('clip', 16)}${t('attach')}</button></div>
      <div class="stack fstack">${uploadsBox({ person: p.id })}${fs.map(f => fileRow(f, { ctx: 'person' })).join('')}</div>`; })() : ''}
      <div class="sec-h"><h3>${t('contactHistory')}</h3></div>
      <div class="box">${hist.length ? `<ul class="tl">${hist.map(({ x, l, f }) => `<li class="k"><small class="num">${fmtShort(l.date)}</small><span dir="auto">${esc(fuLogText(l))}</span><button type="button" class="linkbtn" data-open="${x.id}" dir="auto">${esc(x.title)}</button></li>`).join('')}</ul>` : `<div class="fu muted">${t('noContact')}</div>`}</div>
      <div class="danger-row"><button type="button" class="btn sm ghost danger" data-act="delperson" data-id="${p.id}">${sic('trash', 16)}${t('delete')}</button></div>
    </div>`;
  });
}

/* ================= harvest · presentation · share · PDF ================= */
const thisYM = () => T().slice(0, 7);
const shiftYM = (ym, n) => { const d = pd(ym + '-01'); d.setMonth(d.getMonth() + n); return ds(d).slice(0, 7); };
function harvest(ym) {
  const all = live();
  const byDate = (a, b) => (b.completed_on || '') > (a.completed_on || '') ? 1 : -1;
  const done = all.filter(x => !isLog(x) && x.status === 'done' && (x.completed_on || '').startsWith(ym)).sort(byDate);
  const logs = all.filter(x => isLog(x) && (x.completed_on || '').startsWith(ym)).sort(byDate);
  const withDue = done.filter(x => x.due), onTime = withDue.filter(x => x.completed_on <= x.due);
  const endOfYm = lastOfMonth(ym + '-01'), past = endOfYm < T();
  const open = sortTasks(all.filter(x => !x.archived && !isLog(x) && (past ? (x.created_at || '').slice(0, 10) <= endOfYm && (isOpen(x) || (x.status === 'done' && (x.completed_on || '') > endOfYm)) : isOpen(x))));
  let fuClosed = 0; all.forEach(x => (x.followups || []).forEach(f => { if (f.status === 'done' && (f.closed_on || '').startsWith(ym)) fuClosed++; }));
  const proj = {}; [...done, ...logs].forEach(x => { const p = (x.project || '').trim(); proj[p] = (proj[p] || 0) + 1; });
  const byProject = Object.entries(proj).map(([name, n]) => ({ name, n })).sort((a, b) => b.n - a.n).slice(0, 8);
  return { done, logs, open, kpis: { done: done.length, onTimeRate: withDue.length ? onTime.length / withDue.length : null, fuClosed, open: open.length, logs: logs.length }, byProject };
}
function selOf(ym) {
  if (!UI.hsel || UI.hsel.ym !== ym) {
    const h = harvest(ym);
    UI.hsel = { ym, ids: new Set([...h.done.map(x => x.id), ...h.logs.map(x => x.id), ...h.open.filter(x => x.status === 'prog' || x.priority === 'hi').slice(0, 12).map(x => x.id)]) };
  }
  return UI.hsel.ids;
}
const lastNote = x => { const n = (x.log || []).filter(l => l.kind === 'note' && l.text).pop(); return n ? n.text.slice(0, 300) : ''; };
function snapshot(ym) {
  const h = harvest(ym), sel = selOf(ym);
  const item = (x, kind) => ({
    kind, title: x.title, project: x.project || '', priority: x.priority, status: kind === 'open' ? stKey(x.status) : '', due: x.due, completed_on: x.completed_on,
    onTime: !!(x.due && x.completed_on && x.completed_on <= x.due), result: (x.result || '').slice(0, 600), details: (x.details || '').slice(0, 600), update: kind === 'open' ? lastNote(x) : '',
    steps_done: x.steps_done || 0, steps_total: x.steps_total || 0,
    people: [...new Set([x.waiting_on, ...(x.followups || []).map(f => f.person_id)].filter(Boolean).map(pname))].slice(0, 6),
    files: filesOfTask(x.id).map(f => ({ name: f.name, kind: kindOf(f) })),
  });
  return { v: 1, lang: L(), period: ym, generated_at: nowISO(), owner: { name: state.profile.display_name || '', title: state.profile.job_title || '' }, kpis: h.kpis, byProject: h.byProject,
    generated_on: T(), items: [...h.done.filter(x => sel.has(x.id)).map(x => item(x, 'done')), ...h.logs.filter(x => sel.has(x.id)).map(x => item(x, 'log')), ...h.open.filter(x => sel.has(x.id)).map(x => item(x, 'open'))].slice(0, 300) };
}
function openHarvest(ym = thisYM()) {
  UI.hym = ym; UI.hsel = null; // recompute the default selection: items completed since the last visit must be included
  pushLayer(() => {
    const m = UI.hym, h = harvest(m), sel = selOf(m);
    const rate = h.kpis.onTimeRate == null ? '—' : Math.round(h.kpis.onTimeRate * 100) + '%';
    const maxP = Math.max(1, ...h.byProject.map(p => p.n));
    const row = (x, kind) => `<label class="hrow"><input type="checkbox" data-hsel="${x.id}" ${sel.has(x.id) ? 'checked' : ''}><span class="body"><span class="t" dir="auto">${esc(x.title)}</span><span class="meta">${kind !== 'open' ? `<span>${t('doneOn', { d: fmtShort(x.completed_on) })}</span>${x.due && x.completed_on <= x.due ? `<span class="pill done">${t('rpOnTime')}</span>` : ''}` : `${statusPill(x)}${x.due ? `<span>${rel(x.due)}</span>` : ''}`}${x.project ? `<span dir="auto">${esc(x.project)}</span>` : ''}${filesOfTask(x.id).length ? `<span>${sic('clip', 13)} ${filesOfTask(x.id).length}</span>` : ''}</span></span></label>`;
    return topBar(t('harvestTitle')) + `<div class="layer-b harvest">
      <div class="mswitch"><button type="button" class="iconbtn flipx" data-act="hm" data-d="-1" aria-label="${esc(t('prevMonth'))}">${ic.back}</button><b>${monthName(L(), m)}</b><button type="button" class="iconbtn flipx" data-act="hm" data-d="1" ${m >= thisYM() ? 'disabled' : ''} aria-label="${esc(t('nextMonth'))}">${ic.back.replace('<svg', '<svg style="transform:scaleX(-1)"')}</button></div>
      <div class="hk num ${h.kpis.logs ? 'n5' : ''}"><div><b>${h.kpis.done}</b><span>${t('rpKDone')}</span></div><div class="acc"><b>${rate}</b><span>${t('rpKRate')}</span></div><div><b>${h.kpis.fuClosed}</b><span>${t('rpKFu')}</span></div><div><b>${h.kpis.open}</b><span>${t('rpKOpen')}</span></div>${h.kpis.logs ? `<div><b>${h.kpis.logs}</b><span>${t('rpKLogs')}</span></div>` : ''}</div>
      ${h.byProject.length > 1 ? `<div class="sec-h"><h3>${t('rpByProject')}</h3></div><div class="box hbars">${h.byProject.map(p => `<div class="hbar"><span dir="auto">${esc(p.name || t('rpNoProject'))}</span><i><u style="width:${Math.max(4, Math.round(100 * p.n / maxP))}%"></u></i><b class="num">${p.n}</b></div>`).join('')}</div>` : ''}
      <p class="note">${t('selectHint')}</p>
      <div class="sec-h"><h3>${t('rpAchievements')}</h3><span class="cnt">${h.done.length}</span></div>
      <div class="stack">${h.done.length ? h.done.map(x => row(x, 'done')).join('') : `<div class="box"><div class="fu muted">${t('noDoneMonth')}</div></div>`}</div>
      ${h.logs.length ? `<div class="sec-h" style="margin-top:18px"><h3>${t('rpOther')}</h3><span class="cnt">${h.logs.length}</span></div>
      <div class="stack">${h.logs.map(x => row(x, 'log')).join('')}</div>` : ''}
      <div class="sec-h" style="margin-top:18px"><h3>${t('rpInProgress')}</h3><span class="cnt">${h.open.length}</span></div>
      <div class="stack">${h.open.length ? h.open.slice(0, 40).map(x => row(x, 'open')).join('') : `<div class="box"><div class="fu muted">${t('noOpenNow')}</div></div>`}</div>
      <div class="hact"><button type="button" class="btn primary" data-act="present">${sic('spark', 18)}${t('present')}</button><button type="button" class="btn" data-act="sharelink">${sic('link', 18)}${t('shareLink')}</button><button type="button" class="btn" data-act="pdf">${sic('download', 18)}PDF</button></div>
    </div>`;
  });
}
function openPresent(snap0) {
  const snap = snap0 || snapshot(UI.hym || thisYM()); UI.presentSnap = snap0 || null;
  if (!snap.items.length) return toast(t('nothingSelected'), null, 'err');
  ensureReportCSS();
  let el = $('#present'); if (!el) { el = document.createElement('div'); el.id = 'present'; document.body.appendChild(el); }
  el.innerHTML = `<div class="pr-bar"><button type="button" class="btn sm" data-act="closepresent">${sic('x', 16)}${t('closePresent')}</button><span class="grow"></span>${document.fullscreenEnabled ? `<button type="button" class="btn sm" data-act="fullscreen">${t('fullscreen')}</button>` : ''}<button type="button" class="btn sm primary" data-act="pdf">${sic('download', 16)}PDF</button></div><div class="pr-body">${reportHTML(snap)}</div>`;
  el.hidden = false; document.body.classList.add('presenting'); el.scrollTop = 0;
}
function closePresent() { const el = $('#present'); if (el) el.hidden = true; document.body.classList.remove('presenting'); if (document.fullscreenElement) document.exitFullscreen().catch(() => { }); }
function printReport(snap0) {
  const snap = snap0 || snapshot(UI.hym || thisYM());
  if (!snap.items.length) return toast(t('nothingSelected'), null, 'err');
  ensureReportCSS();
  let pa = $('#printArea'); if (!pa) { pa = document.createElement('div'); pa.id = 'printArea'; document.body.appendChild(pa); }
  pa.innerHTML = reportHTML(snap, { mode: 'print' });
  const old = document.title; document.title = `${t('rpTitle', { m: monthName(L(), snap.period) })} — ${snap.owner.name}`;
  document.body.classList.add('printing');
  const done = () => { document.body.classList.remove('printing'); document.title = old; window.removeEventListener('afterprint', done); };
  window.addEventListener('afterprint', done);
  setTimeout(() => { window.print(); setTimeout(() => { if (!matchMedia('print').matches) done(); }, 1500); }, 60);
}
const shareUrl = tok => location.origin + location.pathname.replace(/[^/]*$/, '') + 'share.html#' + tok;
function newToken() { const b = new Uint8Array(24); crypto.getRandomValues(b); return btoa(String.fromCharCode(...b)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
function openShareSheet() {
  if (state.mode !== 'cloud' || !sb) return toast(t('shareNeedsAccount'), null, 'err');
  const snap = snapshot(UI.hym || thisYM());
  if (!snap.items.length) return toast(t('nothingSelected'), null, 'err');
  UI.shareExp = 30;
  openSheet(t('shareTitle'), `<div class="form" id="shareBox"><p class="ctx" style="font-weight:500">${t('shareExplain', { n: plural(snap.items.length, 'items') })}</p>
    <div class="fld"><span>${t('expiry')}</span><div class="seg2">${[[7, 'exp7'], [30, 'exp30'], [0, 'expNone']].map(([v, l]) => `<button type="button" class="${v === 30 ? 'on' : ''}" data-exp="${v}">${t(l)}</button>`).join('')}</div></div>
    <div class="save-row"><button type="button" class="btn primary block" data-act="createshare">${sic('link', 18)}${t('createLink')}</button></div></div>`, {
    onMount: sh => sh.querySelectorAll('[data-exp]').forEach(b => b.onclick = () => { UI.shareExp = +b.dataset.exp; sh.querySelectorAll('[data-exp]').forEach(y => y.classList.toggle('on', y === b)); })
  });
}
async function createShare(btn) {
  const snap = snapshot(UI.hym || thisYM());
  if (!navigator.onLine) return toast(t('needOnline'), null, 'err');
  btn.disabled = true;
  const token = newToken();
  const exp = UI.shareExp ? new Date(Date.now() + UI.shareExp * 864e5).toISOString() : null;
  const { error } = await sb.from('shares').insert({ token, title: `${t('rpTitle', { m: monthName(L(), snap.period) })} — ${snap.owner.name}`.slice(0, 200), period: snap.period, payload: snap, expires_at: exp });
  if (error) { btn.disabled = false; console.warn(error); return toast(error.code === '23514' ? t('shareTooBig') : t('shareFail'), null, 'err'); }
  const url = shareUrl(token);
  const box = $('#shareBox'); if (!box) return;
  box.innerHTML = `<p class="ctx">${sic('check', 16)} ${t('linkReady')}</p><div class="linkbox" dir="ltr">${esc(url)}</div>
    <div class="save-row"><button type="button" class="btn primary block" data-act="copylink" data-url="${esc(url)}">${t('copy')}</button>${navigator.share ? `<button type="button" class="btn block" data-act="sharevia" data-url="${esc(url)}">${t('shareVia')}</button>` : ''}</div>
    <a class="linkbtn center" href="${esc(url)}" target="_blank" rel="noopener">${t('previewLink')}</a>`;
}
async function copyText(s) { try { await navigator.clipboard.writeText(s); } catch { const ta = document.createElement('textarea'); ta.value = s; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); } catch { } ta.remove(); } toast(t('copied')); }
function openShares() {
  UI.sharesList = null;
  const Lr = pushLayer(() => {
    const list = UI.sharesList;
    const body = list == null ? `<div class="pv-wait"><span class="spin"></span></div>` : !list.length ? `<div class="empty small"><p>${t('sharesEmpty')}</p></div>` : `<div class="stack">${list.map(sh => {
      const expired = sh.expires_at && sh.expires_at < nowISO();
      const st = sh.revoked ? `<span class="pill">${t('revokedSt')}</span>` : expired ? `<span class="pill">${t('expired')}</span>` : `<span class="pill done">${t('active')}</span>`;
      return `<div class="box share-row"><div class="t" dir="auto">${esc(sh.title)}</div><div class="meta">${st}<span class="num">${fmtShort(localDay(sh.created_at))}</span><span>${t('views', { n: plural(sh.views | 0, 'views') })}</span><span>${sh.expires_at ? t('expiresOn', { d: fmtShort(localDay(sh.expires_at)) }) : t('expNone')}</span></div>
        ${!sh.revoked && !expired ? `<div class="row"><button type="button" class="btn sm" data-act="copylink" data-url="${esc(shareUrl(sh.token))}">${t('copy')}</button><a class="btn sm ghost" href="${esc(shareUrl(sh.token))}" target="_blank" rel="noopener">${t('previewLink')}</a><span class="grow"></span><button type="button" class="btn sm ghost danger" data-act="revokeshare" data-sid="${sh.id}">${t('revoke')}</button></div>` : ''}</div>`;
    }).join('')}</div>`;
    return topBar(t('sharesTitle')) + `<div class="layer-b">${body}</div>`;
  });
  loadShares(Lr);
}
async function loadShares(Lr) {
  const { data, error } = await sb.from('shares').select('id,token,title,period,created_at,expires_at,revoked,views').order('created_at', { ascending: false }).limit(50);
  UI.sharesList = error ? [] : data; Lr.rerender();
  if (error) toast(t('shareFail'), null, 'err');
}

/* ================= backup ================= */
function exportBackup() {
  const data = { app: 'rafeeq', version: 2, exported_at: nowISO(), profile: state.profile, people: state.people.filter(p => !p.deleted), tasks: state.tasks.filter(x => !x.deleted), files: state.files.filter(f => !f.deleted), units: (state.units || []).filter(u => !u.deleted) };
  const blob = new Blob([JSON.stringify(data, null, 1)], { type: 'application/json' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `rafeeq-backup-${T()}.json`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 3000);
  toast(t('exported'));
}
function convertV1(d) {
  const pid = new Map(), tid = new Map();
  const ppl = (d.people || []).map(p => { const id = uuid(); pid.set(p.id, id); return { id, name: p.name || '?', org: p.org || '', contact: p.contact || '', notes: p.notes || '', deleted: false, client_ts: nowISO(), created_at: p.createdAt || nowISO() }; });
  const stMap = { not_started: 'todo', in_progress: 'prog', waiting: 'wait', paused: 'hold', done: 'done', cancelled: 'cancelled' };
  const prMap = { high: 'hi', medium: 'mid', low: 'lo' };
  const tasks = (d.tasks || []).map(x => {
    const id = uuid(); tid.set(x.id, id);
    return newTask({ id, title: x.title || '?', details: [x.desc, x.links].filter(Boolean).join('\n\n'), status: stMap[x.status] || 'todo', priority: prMap[x.priority] || 'mid', role: ['exec', 'follow', 'both'].includes(x.role) ? x.role : 'exec', due: validDate(x.due) || null, project: x.project || '', source: x.source || '', waiting_on: x.waitingOn ? pid.get(x.waitingOn) || null : null, waiting_what: x.waitingWhat || 'reply', notes: x.notes || '', completed_on: validDate(x.completedAt) || null, result: x.result || '', archived: !!x.archived, created_at: x.createdAt || nowISO(),
      followups: (x.followups || []).map(f => ({ id: uuid(), person_id: pid.get(f.personId) || null, what: f.what || '', due: validDate(f.due) || T(), status: f.status === 'closed' ? 'done' : 'open', created_at: f.createdAt || nowISO(), closed_on: validDate(f.closedOn) || null, log: (f.log || []).map(l => ({ id: uuid(), at: l.at || nowISO(), date: validDate(l.date) || T(), text: l.text || '' })) })).filter(f => f.person_id),
      log: (x.log || []).map(l => ({ id: uuid(), at: l.at || nowISO(), kind: l.kind === 'note' ? 'note' : l.kind === 'completed' ? 'completed' : l.kind === 'created' ? 'created' : 'note', text: l.kind === 'note' ? (l.text || '') : (l.text || ''), data: {} })).filter(l => l.kind !== 'note' || l.text) });
  });
  return { people: ppl, tasks };
}
async function importBackup(file) {
  let d; try { d = JSON.parse(await file.text()); } catch { toast(t('importBad'), null, 'err'); return; }
  if (!d || d.app !== 'rafeeq' || !Array.isArray(d.tasks)) { toast(t('importBad'), null, 'err'); return; }
  const conv = d.version === 2 ? { people: d.people || [], tasks: d.tasks || [], files: state.mode === 'cloud' ? (d.files || []) : [] } : convertV1(d);
  const okId = v => typeof v === 'string' && /^[0-9a-f-]{36}$/i.test(v);
  const str = (v, n) => typeof v === 'string' ? v.slice(0, n) : '';
  const iso = v => typeof v === 'string' && !isNaN(Date.parse(v)) ? v.slice(0, 40) : nowISO();
  const hhmm = v => /^([01]\d|2[0-3]):[0-5]\d$/.test(v || '') ? v : null;
  const int = (v, lo, hi, dflt) => { const n = parseInt(v, 10); return Number.isFinite(n) ? Math.max(lo, Math.min(hi, n)) : dflt; };
  const fuLog = l => ({ id: okId(l?.id) ? l.id : uuid(), at: iso(l?.at), date: validDate(l?.date) || T(), text: str(l?.text, 1000), outcome: ['done', 'again', 'closed_with_task', 'person_deleted'].includes(l?.outcome) ? l.outcome : 'done', ...(validDate(l?.next) ? { next: l.next } : {}) });
  const fu = f => ({ id: okId(f.id) ? f.id : uuid(), person_id: f.person_id, what: str(f.what, 300), due: validDate(f.due), time: hhmm(f.time), status: f.status === 'done' ? 'done' : 'open', created_at: iso(f.created_at), closed_on: validDate(f.closed_on), log: (Array.isArray(f.log) ? f.log : []).slice(0, 200).map(fuLog) });
  const tlog = e => ({ id: okId(e?.id) ? e.id : uuid(), at: iso(e?.at), kind: str(e?.kind, 20) || 'edited', text: str(e?.text, 2000), data: e?.data && typeof e.data === 'object' && !Array.isArray(e.data) ? e.data : {} });
  conv.tasks = conv.tasks.filter(x => x && okId(x.id) && typeof x.title === 'string').map(x => newTask({
    ...x, kind: x.kind === 'log' ? 'log' : 'task', status: ['todo', 'prog', 'wait', 'review', 'hold', 'done', 'cancelled'].includes(x.status) ? x.status : 'todo', priority: ['hi', 'mid', 'lo'].includes(x.priority) ? x.priority : 'mid',
    role: ['exec', 'follow', 'both'].includes(x.role) ? x.role : 'exec', waiting_what: ['reply', 'decision', 'approval'].includes(x.waiting_what) ? x.waiting_what : 'reply', waiting_on: okId(x.waiting_on) ? x.waiting_on : null,
    due: validDate(x.due), due_time: hhmm(x.due_time), remind_min: x.remind_min == null ? null : int(x.remind_min, -1, 1440, null), completed_on: validDate(x.completed_on), created_at: iso(x.created_at), client_ts: nowISO(),
    title: String(x.title).slice(0, 500), details: str(x.details, 20000), notes: str(x.notes, 20000), result: str(x.result, 5000), project: str(x.project, 200), source: str(x.source, 200),
    steps_done: int(x.steps_done, 0, 50, 0), steps_total: int(x.steps_total, 0, 50, 0), archived: !!x.archived, deleted: !!x.deleted, ai: {},
    unit_id: okId(x.unit_id) ? x.unit_id : null, ext_ref: str(x.ext_ref, 100), progress: x.progress == null ? null : int(x.progress, 0, 100, null),
    followups: (Array.isArray(x.followups) ? x.followups : []).filter(f => f && okId(f.person_id)).slice(0, 100).map(fu), log: (Array.isArray(x.log) ? x.log : []).slice(0, 500).map(tlog), recur: cleanRecur(x.recur),
  }));
  conv.people = conv.people.filter(p => p && okId(p.id) && typeof p.name === 'string' && p.name.trim()).map(p => ({ ...newPerson(p.name.slice(0, 200)), id: p.id, org: str(p.org, 200), contact: str(p.contact, 200), notes: str(p.notes, 5000), deleted: !!p.deleted, created_at: iso(p.created_at) }));
  conv.files = (conv.files || []).filter(f => f && okId(f.id) && typeof f.name === 'string' && /^[\w-]{10,200}$/.test(f.drive_id || '')).map(f => newFile({ id: f.id, drive_id: f.drive_id, name: String(f.name).slice(0, 300), mime: str(f.mime, 100), size: int(f.size, 0, 1e12, 0), task_id: okId(f.task_id) ? f.task_id : null, person_id: okId(f.person_id) ? f.person_id : null, note: str(f.note, 1000), deleted: !!f.deleted, created_at: iso(f.created_at) }));
  const unitsIn = d.version === 2 && Array.isArray(d.units) ? d.units.filter(u => u && okId(u.id) && typeof u.name === 'string' && u.name.trim()).slice(0, 500).map(u => newUnit({ id: u.id, name: u.name.slice(0, 150), kind: u.kind === 'sec' ? 'sec' : 'dept', parent_id: okId(u.parent_id) ? u.parent_id : null, sort: int(u.sort, 0, 100000, 0) })) : [];
  conv.people.forEach(p => { const src = (d.people || []).find(q => q && q.id === p.id); p.unit_id = okId(src?.unit_id) ? src.unit_id : null; });
  confirmSheet(t('importQ', { t: plural(conv.tasks.length, 'tasks'), p: plural(conv.people.length, 'people') }), t('importAdd'), () => {
    const haveU = new Set((state.units || []).map(u => u.id)); unitsIn.filter(u => !haveU.has(u.id)).forEach(u => putUnit(u));
    const haveT = new Set(state.tasks.map(x => x.id)), haveP = new Set(state.people.map(p => p.id));
    const tasks = state.tasks.concat(conv.tasks.filter(x => !haveT.has(x.id)));
    const ppl = state.people.concat(conv.people.filter(p => !haveP.has(p.id)));
    const haveF = new Set(state.files.map(f => f.id));
    const fls = (conv.files || []).length ? state.files.concat(conv.files.filter(f => !haveF.has(f.id))) : undefined;
    bulkReplace({ tasks, people: ppl, files: fls, added: { tasks: conv.tasks.filter(x => !haveT.has(x.id)).map(x => x.id), people: conv.people.filter(p => !haveP.has(p.id)).map(p => p.id), files: (conv.files || []).filter(f => !haveF.has(f.id)).map(f => f.id) } }); toast(t('imported'));
  }, false);
}

/* ================= events ================= */
document.addEventListener('click', async e => {
  const q = s => e.target.closest(s); let el;
  if ((el = q('[data-tab]'))) { e.preventDefault(); return go(el.dataset.tab); }
  if ((el = q('[data-jump]'))) { document.getElementById(el.dataset.jump)?.scrollIntoView({ behavior: 'smooth', block: 'start' }); return; }
  if ((el = q('[data-seg]'))) { UI.seg = el.dataset.seg; return renderMain(); }
  if ((el = q('[data-tview]'))) { UI.tview = el.dataset.tview === 'table' ? 'table' : 'cards'; return renderMain(); }
  if ((el = q('[data-tsort]'))) { const k = el.dataset.tsort; UI.tsort = { k, d: UI.tsort?.k === k ? -(UI.tsort.d || 1) : 1 }; UI.keepScroll = true; return renderMain(); }
  if ((el = q('[data-wk]'))) { UI.wk = +el.dataset.wk || 0; return renderMain(); }
  if ((el = q('[data-proj]'))) { e.stopPropagation(); return openProject(el.dataset.proj); }
  if ((el = q('[data-pview]'))) { UI.pview = el.dataset.pview; return renderMain(); }
  if ((el = q('[data-pf]'))) { UI.personF = UI.personF === el.dataset.pf ? null : el.dataset.pf; return renderMain(); }
  if ((el = q('[data-unitf]'))) { const v = el.dataset.unitf; UI.unitF = !el.dataset.sec && (UI.unitF === v || (el.closest('.unit-row') && deptOf(UI.unitF)?.id === v)) ? null : v; if (UI.tab !== 'tasks' || UI.layers.length) { UI.tview = 'cards'; UI.seg = 'active'; go('tasks'); } else renderMain(); return; }
  if ((el = q('[data-ltype]'))) { UI.ltype = el.dataset.ltype; return renderMain(); }
  if ((el = q('[data-file]'))) return openFile(el.dataset.file);
  if ((el = q('[data-person]')) && !q('[data-act]')) return openPerson(el.dataset.person);
  el = q('[data-act]');
  if (el) {
    const a = el.dataset.act, id = el.dataset.id, x = id ? taskById(id) : null;
    const f = x && el.dataset.fu ? x.followups.find(y => y.id === el.dataset.fu) : null;
    e.stopPropagation();
    switch (a) {
      case 'google': { const r = await signInGoogle(); if (r?.error) toast(t('loginFailed'), null, 'err'); return; }
      case 'local': return enterLocal();
      case 'lang': { const nl = L() === 'ar' ? 'en' : 'ar'; state.profile.lang = nl; if (state.mode) saveProfile({ lang: nl }); else applyLangTheme(); return renderAll(); }
      case 'setlang': { const nm = $('#onbName')?.value, jb = $('#onbJob')?.value; state.profile.lang = el.dataset.v; renderOnboarding(); applyLangTheme(); if (nm != null) $('#onbName').value = nm; if (jb != null) $('#onbJob').value = jb; return; }
      case 'theme': { const s = Object.assign({}, state.profile.settings); s.theme = s.theme === 'dark' ? 'light' : 'dark'; saveProfile({ settings: s }); return renderAll(); }
      case 'capture': case 'chat': return openChat();
      case 'chatsend': return chatSend($('#chatIn')?.value);
      case 'chatq': return chatSend(el.dataset.q, el.dataset.l);
      case 'chatsave': return saveDrafts();
      case 'draftadd': { const d = CH.drafts[+el.dataset.n], k = el.dataset.k; if (!d) return; CH.extra = CH.extra && CH.extra.id === d.id && CH.extra.k === k ? null : { id: d.id, k }; renderChat(); if (CH.extra) setTimeout(() => $('#chatIn')?.focus(), 50); return; }
      case 'anssoft': { const sp = softPending(); if (!sp) return renderChat(); if (el.dataset.all) CH.drafts.filter(d => softOf(d) === sp.k).forEach(d => setSoft(d, sp.k, el.dataset.v)); else setSoft(sp.d, sp.k, el.dataset.v); return renderChat(); }
      case 'actrun': return runActions();
      case 'actdrop': CH.actions = null; botSay(t('actCancelled')); return renderChat();
      case 'actpick': { const a = (CH.actions || [])[+el.dataset.i]; if (a) a.task = el.dataset.v; return renderChat(); }
      case 'unitdept': UI.askUnitDept = el.dataset.v; return renderChat();
      case 'unitback': UI.askUnitDept = null; return renderChat();
      case 'skipsoft': { CH.drafts.forEach(d => { d.softSkip = true; }); return renderChat(); }
      case 'ansextra': { const e = extraOn(); if (e && !el.dataset.skip) { const d = CH.drafts[e.n]; if (e.k === 'time') d.time = el.dataset.v; else d.project = el.dataset.v; } CH.extra = null; return renderChat(); }
      case 'chatdiscard': CH.drafts = []; CH.pick = new Set(); return renderChat();
      case 'draftdrop': CH.drafts.splice(+el.dataset.n, 1); return renderChat();
      case 'draftmerge': { const n = +el.dataset.n; const a = CH.drafts[n - 1], b = CH.drafts[n]; if (!a || !b) return; if (isLogD(a) || isLogD(b)) { if (!isLogD(a) || !isLogD(b)) return; a.title = `${a.title} ${t('andJoin2')} ${b.title}`.slice(0, 300); a.project = a.project || b.project; a.match = matchOpenTask(a.title); a.closeTask = undefined; CH.drafts.splice(n, 1); return renderChat(); } a.title = `${a.title} ${t('andJoin')} ${b.title}`.slice(0, 300); b.people.forEach(p => { if (!a.people.some(q => refName(q.ref) === refName(p.ref))) a.people.push(p); }); if (b.due && (!a.due || b.due < a.due)) a.due = b.due; if (b.priority === 'hi') a.priority = 'hi'; a.source = a.source || b.source; a.time = a.time || b.time; a.recur = a.recur || b.recur; if (a.role !== b.role) a.role = 'both'; CH.drafts.splice(n, 1); return renderChat(); }
      case 'draftedit': { const d = CH.drafts[+el.dataset.n]; if (!d) return; if (isLogD(d)) return openLogForm(null, { draftId: d.id, title: d.title, date: d.date, project: d.project }); return openTaskForm(null, { draftId: d.id, title: d.title, due: d.due, time: d.time, recur: d.recur, noDue: d.noDue, priority: d.priority, role: d.role, source: d.source, project: d.project, unit: d.unit || '', who: d.people.filter(p => p.role === 'followup').map(p => refName(p.ref)), extra: d.extra || null }); }
      case 'ansdue': { const pp = pending(); if (pp) { pp.d.due = el.dataset.v || null; pp.d.noDue = !el.dataset.v; pp.d.weekendOk = false; } return renderChat(); }
      case 'answk': { const pp = pending(); if (pp) { if (el.dataset.v === 'keep') pp.d.weekendOk = true; else pp.d.due = el.dataset.v; } return renderChat(); }
      case 'anssrc': { const pp = pending(), v = el.dataset.v; if (el.dataset.all) CH.drafts.forEach(x => { if (!isLogD(x) && !x.source) x.source = v; }); else if (pp) pp.d.source = v; CH.lastSrc = v; return renderChat(); }
      case 'answho': { const v = el.dataset.v; if (CH.pick.has(v)) CH.pick.delete(v); else CH.pick.add(v); return renderChat(false); }
      case 'answhodone': { const pp = pending(); if (pp) CH.pick.forEach(pid => { const ex = pp.d.people.find(p => p.ref.kind === 'known' && p.ref.id === pid); if (ex) { if (ex.role === 'related') ex.role = 'followup'; } else pp.d.people.push({ ref: { kind: 'known', id: pid }, role: 'followup', what: '', due: null }); }); CH.pick = new Set(); return renderChat(); }
      case 'answhorel': { const pp = pending(); const pr = pp && pp.d.people[+el.dataset.i]; if (pr) pr.role = 'followup'; return renderChat(); }
      case 'ansref': { const pp = pending(); if (pp && pp.d.people[+el.dataset.i]) { const pr = pp.d.people[+el.dataset.i]; pr.ref = el.dataset.v === 'keep' ? { kind: 'new', name: pr.ref.label, desc: '' } : { kind: 'known', id: el.dataset.v }; } return renderChat(); }
      case 'remind': return x && f && openRemind(x, f);
      case 'remwa': { const msg = $('#remTxt')?.value || ''; const ph = el.dataset.phone; window.open(`https://wa.me/${ph || ''}?text=${encodeURIComponent(msg)}`, '_blank', 'noopener'); if (x && f) logReminder(x, f); closeSheet(true); return toast(t('reminderLogged')); }
      case 'remcopy': { await copyText($('#remTxt')?.value || ''); if (x && f) logReminder(x, f); return closeSheet(true); }
      case 'newtask': return openTaskForm();
      case 'newlog': return openLogForm();
      case 'morelist': UI.listMax = (UI.listMax || 80) + 100; UI.keepScroll = true; return renderMain();
      case 'weekend': return openWeekendSheet();
      case 'wkset': { const st = Object.assign({}, state.profile.settings); st.wk = el.dataset.v.split(',').map(Number); saveProfile({ settings: st }); setWeekend(st.wk); closeSheet(true); toast(t('saved')); return renderAll(); }
      case 'notifs': return openNotifSheet();
      case 'notifon': { el.disabled = true; const f = $('#nForm'); return pushEnable().then(res => { el.disabled = false; if (res === 'ok') { saveNotifPrefs(f); UI.pushOn = true; toast(t('notifEnabled')); closeSheet(true); openNotifSheet(); renderMain(); } else toast(t(res === 'denied' ? 'notifDenied' : res === 'unsupported' ? 'notifUnsupported' : 'notifError'), null, 'err'); }); }
      case 'notiftest': { el.disabled = true; const h0 = el.innerHTML; el.innerHTML = `<span class="spin" aria-hidden="true"></span>${t('sendingTest')}`; return pushTest().catch(() => null).then(r => { el.disabled = false; el.innerHTML = h0; toast(r && r.devices ? t('testSent') : t('notifError'), null, r && r.devices ? undefined : 'err'); }); }
      case 'notifoff': return pushDisable().then(() => { UI.pushOn = false; closeSheet(true); toast(t('notifOffDone')); renderMain(); });
      case 'skiprec': { if (!x?.recur) return; let nd = recurNext(x.recur, x.due || T()), g = 0; while (nd && nd <= T() && g++ < 400) nd = recurNext(x.recur, nd); if (!nd) return; addLog(x, 'skipped', '', { from: x.due, to: nd }); x.due = nd; openFus(x).forEach(f => { f.due = diffDays(nd, T()) > 1 ? addWorkdays(nd, -1) : nd; }); putTask(x); return toast(t('skippedTo', { d: rel(nd) })); }
      case 'chatfill': { if (!$('#chatIn')) openChat(); const ta = $('#chatIn'); if (ta) { ta.value = el.dataset.q || t('logStarter'); autoGrow(ta); ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length); } return; }
      case 'ansmatch': { const pp = pending(); if (pp && pp.m.k === 'match') pp.d.closeTask = el.dataset.v === 'yes' ? pp.d.match : null; return renderChat(); }
      case 'closefus': { if (x && openFus(x).length) { openFus(x).forEach(f => { f.status = 'done'; f.closed_on = T(); f.log = f.log || []; f.log.push({ id: uuid(), at: nowISO(), date: T(), text: '', outcome: 'closed_with_task' }); }); putTask(x); toast(t('fusClosed')); } el.disabled = true; el.closest('.ask')?.remove(); return; }
      case 'newperson': return openPersonForm();
      case 'profile': return openProfile();
      case 'closesheet': return closeSheet();
      case 'pop': return popLayer();
      case 'quickdone': return quickDone(id);
      case 'postpone': return postpone(id);
      case 'complete': return x && openComplete(x);
      case 'reopen': if (x) { addLog(x, 'reopened', '', { prev: x.completed_on }); x.status = 'prog'; x.completed_on = null; putTask(x); toast(t('reopened')); } return;
      case 'setst': if (x) { if (el.dataset.v === 'wait' && !x.waiting_on) openTaskForm(x, { status: 'wait' }); else setStatus(x, el.dataset.v); } return;
      case 'edittask': return x && (isLog(x) ? openLogForm(x) : openTaskForm(x));
      case 'archive': if (x) { x.archived = !x.archived; addLog(x, x.archived ? 'archived' : 'unarchived'); putTask(x); toast(x.archived ? t('archivedToast') : t('unarchivedToast')); } return;
      case 'deltask': return x && confirmSheet(t(isLog(x) ? 'delLogQ' : 'delTaskQ'), t('delete'), () => { const prev = structuredClone(x); x.deleted = true; putTask(x); popLayer(); toast(t('deleted'), () => putTask(prev)); });
      case 'stepplus': if (x && x.steps_done < x.steps_total) { x.steps_done++; addLog(x, 'steps', '', { done: x.steps_done, total: x.steps_total }); if (x.status === 'todo') x.status = 'prog'; putTask(x); } return;
      case 'stepminus': if (x && x.steps_done > 0) { x.steps_done--; putTask(x); } return;
      case 'addfu': return x && openFuAdd(x);
      case 'fumenu': return x && f && openFuAdd(x, f);
      case 'furesult': return x && f && openFuResult(x, f);
      case 'fudone': return UI.fuCtx && fuDone(UI.fuCtx.x, UI.fuCtx.f);
      case 'funotyet': { const b = $('#fuNext'); if (b) b.hidden = false; return; }
      case 'funext': return UI.fuCtx && fuAgain(UI.fuCtx.x, UI.fuCtx.f, el.dataset.v);
      case 'editperson': { const p = personById(id); return p && openPersonForm(p); }
      case 'delperson': {
        const p = personById(id); if (!p) return;
        const refs = live().filter(x => isOpen(x) && (x.waiting_on === p.id || openFus(x).some(f => f.person_id === p.id)));
        return confirmSheet(refs.length ? t('delPersonRefsQ', { n: refs.length }) : t('delPersonQ'), t('delete'), () => {
          const snap = { p: structuredClone(p), tasks: refs.map(x => structuredClone(x)) };
          refs.forEach(x => { if (x.waiting_on === p.id) { x.waiting_on = null; if (x.status === 'wait') x.status = 'prog'; } openFus(x).filter(f => f.person_id === p.id).forEach(f => { f.status = 'done'; f.closed_on = T(); f.log = f.log || []; f.log.push({ id: uuid(), at: nowISO(), date: T(), text: '', outcome: 'person_deleted' }); }); putTask(x); });
          p.deleted = true; putPerson(p); popLayer();
          toast(t('deleted'), () => { putPerson(snap.p); snap.tasks.forEach(putTask); });
        });
      }
      case 'mic': return toggleMic();
      case 'aiundo': return undoAI(id);
      case 'harvest': return openHarvest();
      case 'org': return openOrg();
      case 'orgimport': case 'chatattach': return pickSheet();
      case 'orgold': return toast(t('orgNeedsDb'), null, 'err', 6000);
      case 'unitadd': return openUnitForm(null, el.dataset.parent || null);
      case 'unitedit': { const u = unitById(el.dataset.u); return u && openUnitForm(u); }
      case 'stmenu': return x && openStatusMenu(x);
      case 'stpick': { if (!x) return; const v = el.dataset.v; closeSheet(true); if (v === 'done') return isOpen(x) && openComplete(x); if (v === 'wait' && !x.waiting_on) return openTaskForm(x, { status: 'wait' }); if (x.status !== v) { setStatus(x, v); toast(t('statusSet', { s: t(stKey(v)) })); } return; }
      case 'hm': { const n = shiftYM(UI.hym, +el.dataset.d); if (n > thisYM()) return; UI.hym = n; UI.layers.forEach(Lr => Lr.rerender()); return; }
      case 'present': return openPresent();
      case 'projreport': return openPresent(projectSnapshot(el.dataset.v));
      case 'closepresent': return closePresent();
      case 'fullscreen': { const pe = $('#present'); if (pe && !document.fullscreenElement) pe.requestFullscreen?.().catch(() => { }); else document.exitFullscreen?.(); return; }
      case 'pdf': return printReport($('#present') && !$('#present').hidden ? UI.presentSnap : null);
      case 'sharelink': return openShareSheet();
      case 'createshare': return createShare(el);
      case 'copylink': return copyText(el.dataset.url);
      case 'sharevia': { try { await navigator.share({ title: document.title, url: el.dataset.url }); } catch { } return; }
      case 'shares': return openShares();
      case 'revokeshare': return confirmSheet(t('revokeQ'), t('revoke'), async () => { const { error } = await sb.from('shares').update({ revoked: true }).eq('id', el.dataset.sid); if (error) return toast(t('shareFail'), null, 'err'); const it = (UI.sharesList || []).find(x => x.id === el.dataset.sid); if (it) it.revoked = true; UI.layers.forEach(Lr => Lr.rerender()); toast(t('revokedSt')); });
      case 'askmore': { const st = Object.assign({}, state.profile.settings); st.askMore = !askMore(); saveProfile({ settings: st }); return renderMain(); }
      case 'asksrc': { const st = Object.assign({}, state.profile.settings); st.askSource = !askSource(); saveProfile({ settings: st }); return renderMain(); }
      case 'tour': return openTour(0);
      case 'xlsx': return exportXlsx();
      case 'tcols': UI.tcols = UI.tcols === 'all' ? 'key' : 'all'; UI.keepScroll = true; return renderMain();
      case 'tfreset': UI.tf = {}; return renderMain();
      case 'tournext': return openTour(+el.dataset.v);
      case 'tourdone': { try { localStorage.setItem('rafeeq2.tour.' + (state.user?.id || 'local'), '1'); } catch { } closeSheet(true); if (!(state.profile.settings || {}).tourDone) { const st = Object.assign({}, state.profile.settings, { tourDone: 1 }); saveProfile({ settings: st }, { quiet: true }); } if (el.dataset.chat) openChat(); return; }
      case 'aitoggle': { const st = Object.assign({}, state.profile.settings); st.ai = st.ai === false; saveProfile({ settings: st }); if (st.ai) { UI.aiBlocked = null; scheduleAI(500); } return renderMain(); }
      case 'filetask': {
        const fl = fileById(el.dataset.fid); if (!fl) return;
        const nm = fl.name.replace(/\.[^.]{1,5}$/, '').replace(/[_\-]+/g, ' ').replace(/\s+/g, ' ').trim() || fl.name;
        const nx = newTask({ title: nm.slice(0, 300) }); addLog(nx, 'created'); addLog(nx, 'file_added', fl.name); putTask(nx);
        fl.task_id = nx.id; putFile(fl); closeSheet(true); openTask(nx.id); toast(t('taskFromFile')); return;
      }
      case 'upload': return pickFiles({ task: el.dataset.task || null, person: el.dataset.person || null });
      case 'connectdrive': return openConnectSheet();
      case 'connectgo': { if (!navigator.onLine) return toast(t('needOnline'), null, 'err'); const r = await connectDrive(); if (r?.error) toast(t('loginFailed'), null, 'err'); return; }
      case 'unlinkdrive': return confirmSheet(t('unlinkQ'), t('unlinkDrive'), async () => { await unlinkDrive(); toast(t('unlinked')); });
      case 'retryup': { const u = UPQ.find(y => y.id === el.dataset.up); if (u && u.file) { u.st = 'wait'; u.err = ''; drawUps(); runUploads(); } return; }
      case 'dropup': { const i = UPQ.findIndex(y => y.id === el.dataset.up); if (i >= 0) UPQ.splice(i, 1); return drawUps(); }
      case 'editfile': { const fl = fileById(el.dataset.fid); return fl && openFileForm(fl); }
      case 'dlfile': { const fl = fileById(el.dataset.fid); return fl && downloadFile(fl); }
      case 'delfile': { const fl = fileById(el.dataset.fid); return fl && deleteFile(fl); }
      case 'sync': return syncNow();
      case 'importv1': { const d = v1Data(); if (!d) return; const conv = convertV1(d); bulkReplace({ tasks: state.tasks.concat(conv.tasks), people: state.people.concat(conv.people) }); try { localStorage.setItem('rafeeq2.v1imported', '1'); } catch { } toast(t('imported')); return renderMain(); }
      case 'skipv1': { try { localStorage.setItem('rafeeq2.v1imported', 'skip'); } catch { } return renderMain(); }
      case 'adopt': { const n = adoptLocalLeftovers(); toast(t('adopted', { n: plural(n, 'tasks') })); return renderMain(); }
      case 'export': return exportBackup();
      case 'import': return $('#importFile').click();
      case 'signout': {
        const pend = state.sync.pending;
        const msg = state.mode === 'cloud' ? (pend ? t('signOutPendingQ', { n: plural(pend, 'edits') }) : t('signOutQ')) : t('leaveLocalQ');
        return confirmSheet(msg, state.mode === 'cloud' ? t('signOut') : t('leaveLocal'), () => signOut({ wipe: state.mode === 'cloud' }));
      }
    }
    return;
  }
  if ((el = q('[data-open]'))) return openTask(el.dataset.open);
});
document.addEventListener('submit', e => {
  const f = e.target;
  if (f.id === 'onbForm') {
    e.preventDefault();
    const n = $('#onbName').value.trim(); if (!n) return $('#onbName').focus();
    UI.onboarded = true; saveProfile({ display_name: n, job_title: $('#onbJob').value.trim(), lang: L() });
    renderAll(); return setTimeout(() => openTour(0), 350);
  }
  if (f.dataset.form === 'upd') {
    e.preventDefault(); const x = taskById(f.dataset.id); const v = (f.elements.u.value || '').trim(); if (!x || !v) return;
    addLog(x, 'note', v); if (x.status === 'todo') { addLog(x, 'status', '', { from: 'todo', to: 'prog' }); x.status = 'prog'; }
    putTask(x); toast(t('added'));
  }
});
document.addEventListener('input', e => {
  if (e.target.id === 'libQ') { UI.lq = e.target.value; clearTimeout(UI.lqt); UI.lqt = setTimeout(() => renderListOnly('library', '#libList'), 120); return; }
  if (e.target.id === 'taskQ') { UI.q = e.target.value; clearTimeout(UI.qt); UI.qt = setTimeout(() => renderListOnly('tasks', '#taskList'), 120); }
});
document.addEventListener('change', e => {
  if (e.target.dataset?.pu) { const p = personById(e.target.dataset.pu); if (p) { p.unit_id = unitById(e.target.value) ? e.target.value : null; putPerson(p); toast(t('saved')); } return; } if (/^tf(St|Person|Project|Unit|From|To)$/.test(e.target.id || '')) { const k = { tfSt: 'st', tfPerson: 'person', tfProject: 'project', tfUnit: 'unit', tfFrom: 'from', tfTo: 'to' }[e.target.id]; UI.tf = Object.assign({}, UI.tf, { [k]: e.target.type === 'date' ? (validDate(e.target.value) || '') : e.target.value }); UI.keepScroll = true; return renderMain(); } if (e.target.id === 'ansTime' && /^\d{2}:\d{2}$/.test(e.target.value)) { const ex = extraOn(); if (ex) { CH.drafts[ex.n].time = e.target.value; CH.extra = null; } return renderChat(); } if (e.target.id === 'ansDate' && validDate(e.target.value)) { const pp = pending(); if (pp) { pp.d.due = e.target.value; pp.d.noDue = false; pp.d.weekendOk = isWeekend(e.target.value); } return renderChat(); } if (e.target.dataset?.hsel) { const ids = selOf(UI.hym); if (e.target.checked) ids.add(e.target.dataset.hsel); else ids.delete(e.target.dataset.hsel); return; } if (e.target.id === 'libMonth') { UI.lmonth = e.target.value; return renderMain(); } if (e.target.id === 'importFile' && e.target.files[0]) { importBackup(e.target.files[0]); e.target.value = ''; } });
document.addEventListener('keydown', e => {
  const tgt = e.target;
  if ((e.key === 'Enter' || e.key === ' ') && tgt.matches?.('[role="button"][data-open]')) { e.preventDefault(); openTask(tgt.dataset.open); return; }
  if (e.key === 'Enter' && tgt.id === 'chatIn' && !e.shiftKey && !e.isComposing) { e.preventDefault(); chatSend(tgt.value); return; }
  if (e.key === 'Escape') { if ($('#present') && !$('#present').hidden) { closePresent(); return; } if (!$('#confirm').hidden) { $('#confirm').hidden = true; return; } if ($('#sheet').classList.contains('on')) { closeSheet(); return; } if (UI.layers.length) { popLayer(); return; } }
  if (state.mode && !tgt.closest?.('input,textarea,select,[contenteditable]') && !$('#sheet').classList.contains('on') && $('#confirm').hidden && ($('#present')?.hidden ?? true) && !e.ctrlKey && !e.metaKey && !e.altKey) {
    if (e.key === '/' || e.key === 'n' || e.key === 'N' || e.key === 'ى') { e.preventDefault(); openChat(); }
  }
});

/* ================= live updates ================= */
onChange(why => {
  if (why === 'auth') { if (!state.mode) { closeAllLayers(); closeSheet(true); $('#confirm').hidden = true; if ($('#present')) $('#present').hidden = true; } return renderAll(); }
  if (why === 'profile') { if ($('#onbForm') || $('#sheet')?.classList.contains('on')) return; UI.keepScroll = true; return renderAll(); }
  if (why === 'drive') {
    if (state.drive.wrongAccount) { state.drive.wrongAccount = false; toast(t('driveWrongAccount'), null, 'err', 9000); signOut({ wipe: false }); return; }
    if (state.drive.justLinked) { state.drive.justLinked = false; toast(t('driveLinked')); if (state.mode) return go('library'); }
    if (state.mode) { renderMain(); UI.layers.forEach(Lr => Lr.rerender()); }
    return;
  }
  if (!state.mode) return;
  if (why === 'sync') { document.querySelectorAll('.nav-foot, .hdr-side').forEach(n => n.innerHTML = syncChip()); const sc = document.querySelector('.sync-card'); if (sc && UI.tab === 'more') renderMain(); return; }
  if (why === 'data') {
    if ($('#onbForm')) return; renderNav(); UI.keepScroll = true;
    const a = document.activeElement?.id; // keep the search box (and the phone keyboard) mounted while typing
    if (a === 'taskQ' && UI.tab === 'tasks') renderListOnly('tasks', '#taskList'); else if (a === 'libQ' && UI.tab === 'library') renderListOnly('library', '#libList'); else renderMain();
    UI.layers.forEach(Lr => Lr.rerender());
    if (location.hash.startsWith('#t=')) routeHash(); // a notification tap that arrived before the task was pulled
  }
});


/* ================= boot ================= */
renderAll();
/* Native date fields follow the BROWSER's locale (an English Chrome shows mm/dd/yyyy inside the Arabic UI).
   When it differs from the app language, the field's own text is hidden and the date is written over it in the app's language. */
const DATE_VAL = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
function dressDate(inp) {
  if (inp.dataset.dd) return; inp.dataset.dd = '1';
  const host = inp.parentElement; if (!host) return;
  host.classList.add('dd-host'); inp.classList.add('dd');
  const s = document.createElement('span'); s.className = 'dd-show'; s.setAttribute('aria-hidden', 'true'); inp.after(s);
  const upd = () => {
    const v = DATE_VAL.get.call(inp);
    s.textContent = inp.type === 'time' ? (v ? fmtTime(v) : t('pickTime')) : v ? fmt(v, inp.offsetWidth >= 210 ? { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' } : { day: 'numeric', month: 'short', year: 'numeric' }) : t('pickDate');
    s.classList.toggle('ph', !v); s.classList.toggle('off', inp.disabled);
    Object.assign(s.style, { left: inp.offsetLeft + 'px', top: inp.offsetTop + 'px', width: inp.offsetWidth + 'px', height: inp.offsetHeight + 'px' });
  };
  Object.defineProperty(inp, 'value', { configurable: true, get() { return DATE_VAL.get.call(this); }, set(v) { DATE_VAL.set.call(this, v); upd(); } });
  ['input', 'change', 'focus', 'blur'].forEach(e => inp.addEventListener(e, upd));
  inp.addEventListener('click', () => { try { inp.showPicker(); } catch { } });
  new MutationObserver(upd).observe(inp, { attributes: true, attributeFilter: ['disabled'] });
  requestAnimationFrame(upd); setTimeout(upd, 350);
}
const dateLocaleOff = () => (navigator.language || 'en').slice(0, 2).toLowerCase() !== L();
const dressAll = root => { if (dateLocaleOff()) (root.querySelectorAll ? root : document).querySelectorAll('input[type=date],input[type=time]').forEach(dressDate); };
new MutationObserver(ms => { if (!dateLocaleOff()) return; for (const m of ms) for (const n of m.addedNodes) if (n.nodeType === 1) { if (n.matches('input[type=date],input[type=time]')) dressDate(n); else dressAll(n); } }).observe(document.body, { childList: true, subtree: true });
addEventListener('resize', () => document.querySelectorAll('input.dd').forEach(i => i.dispatchEvent(new Event('blur'))));

boot().catch(err => console.error(err)).finally(() => { UI.booting = false; renderAll(); setTimeout(maybeTour, 2500); });
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  const bootAt = Date.now();
  let reloaded = false;
  const WANT = 'rafeeq-' + (window.RAFEEQ_VERSION || '');
  // a newer release is active (or just took over): reload at once on a fresh open, or offer a reload while the user is working
  const newRelease = () => {
    if (reloaded) return; reloaded = true;
    if (/[?&](code|error)=/.test(location.search)) return;
    let n = 0; try { n = +sessionStorage.getItem('rafeeq2.rl') || 0; sessionStorage.setItem('rafeeq2.rl', String(n + 1)); } catch { }
    if (n >= 2) return; // never loop
    if (Date.now() - bootAt < 8000 && !$('#sheet')?.classList.contains('on') && !$('#onbForm') && !UI.layers.length) { location.reload(); return; }
    toast(t('updateReady'), () => location.reload(), undefined, 15000); const b = $('#undoBtn'); if (b) b.textContent = t('updateNow');
  };
  navigator.serviceWorker.addEventListener('controllerchange', newRelease);
  navigator.serviceWorker.addEventListener('message', e => { if (e.data && e.data.version && e.data.version !== WANT) newRelease(); else if (e.data && e.data.version === WANT) { try { sessionStorage.removeItem('rafeeq2.rl'); } catch { } } });
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').then(reg => {
    const ask = () => navigator.serviceWorker.controller?.postMessage('version');
    navigator.serviceWorker.ready.then(ask); setTimeout(ask, 3000);
    if (navigator.onLine) reg.update().catch(() => { });
    setInterval(() => reg.update().catch(() => { }), 6 * 3600 * 1000);
  }).catch(() => { }));
}
