// UI wiring. All parsing happens in the browser; files are cached in IndexedDB, grades in localStorage.
// `G` (grading logic) comes from grading.js.
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const KEY = 'labgrader:v1';
const TIER_CSS = ['t0', 't1', 't2', 't3'];

let S = loadState();
let zip = null, zipPaths = [], roster = null, keyChunks = null, instrLabels = [], students = [];
let cur = { s: 0, q: 0 };
let objectUrls = [];
let formulas = [], subKinds = new Set();  // Excel formulas found in the current submission, and its file types

function loadState() {
  // Defaults mirror the Lab 2 gradebook: 10 completeness + 5 accuracy points, Excel submissions.
  const base = { name: '', questions: [], cfg: { cTotal: 10, aTotal: 5 }, expectedExt: 'xlsx', canvasCol: null, grades: {}, fileNames: {} };
  try {
    const s = JSON.parse(localStorage.getItem(KEY));
    if (s) return { ...base, ...s, cfg: { ...base.cfg, ...s.cfg } };  // older saves had per-question cPts/aPts
  } catch {}
  return base;
}
const save = () => localStorage.setItem(KEY, JSON.stringify(S));

function toast(msg, ms = 3500) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove('show'), ms);
}

// ---------- IndexedDB file cache (so a reload resumes without re-dropping files) ----------
function idb(mode, fn) {
  return new Promise((resolve, reject) => {
    const open = indexedDB.open('labgrader', 1);
    open.onupgradeneeded = () => open.result.createObjectStore('files');
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const tx = open.result.transaction('files', mode), req = fn(tx.objectStore('files'));
      tx.oncomplete = () => resolve(req && req.result);
      tx.onerror = () => reject(tx.error);
    };
  });
}

// ---------- Loading files ----------
async function docxParagraphs(blob) {
  const r = await mammoth.extractRawText({ arrayBuffer: await blob.arrayBuffer() });
  return r.value.split('\n');
}

async function ingest(kind, blob) {
  if (kind === 'zip') {
    zip = await JSZip.loadAsync(blob);
    zipPaths = Object.keys(zip.files).filter(p => !zip.files[p].dir && !p.startsWith('__MACOSX/') && !/(^|\/)\.DS_Store$/.test(p));
  } else if (kind === 'roster') {
    roster = G.parseCanvasRoster(await blob.text());
  } else if (kind === 'instr') {
    instrLabels = G.parseQuestionLabels(await docxParagraphs(blob));
  } else if (kind === 'key') {
    keyChunks = G.parseKeyChunks(await docxParagraphs(blob));
  }
}

function resetQuestions() {
  const src = instrLabels.length ? instrLabels : (keyChunks ? keyChunks.labels.map(label => ({ label, prompt: '' })) : []);
  S.questions = src.map(({ label, prompt }) => ({ label, prompt, void: false, fullCredit: false }));
}

function refresh() {
  if (!S.questions.length) resetQuestions();
  students = G.buildStudents(roster, zipPaths);
  cur.s = Math.min(cur.s, Math.max(0, students.length - 1));
  cur.q = Math.min(cur.q, Math.max(0, S.questions.length - 1));
  save();
  renderSetup();
  renderHeader();
  if (currentView() === 'grade') renderGrade(true);
  if (currentView() === 'overview') renderOverview();
}

async function onFile(kind, file) {
  try {
    await ingest(kind, file);
    await idb('readwrite', st => st.put(file, kind));
    S.fileNames[kind] = file.name;
    if (kind === 'instr' && !S.name) S.name = file.name.replace(/\.docx$/i, '');
    if (kind === 'instr' || (kind === 'key' && !instrLabels.length)) resetQuestions();
    refresh();
    toast(`Loaded ${file.name}`);
  } catch (e) {
    console.error(e); toast(`Could not read ${file.name}: ${e.message}`, 6000);
  }
}

async function restore() {
  for (const kind of ['zip', 'roster', 'instr', 'key']) {
    try { const blob = await idb('readonly', st => st.get(kind)); if (blob) await ingest(kind, blob); }
    catch (e) { console.error(kind, e); }
  }
  refresh();
  showView(zipPaths.length ? 'grade' : 'setup');
}

// ---------- Views ----------
const currentView = () => document.body.dataset.view;
function showView(v) {
  document.body.dataset.view = v;
  document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.view === v));
  if (v === 'grade') renderGrade(true);
  if (v === 'overview') renderOverview();
}

const perQ = () => G.perQuestion(S.cfg, S.questions);
const maxPts = () => (perQ().n ? S.cfg.cTotal + S.cfg.aTotal : 0);
const issues = s => G.formatIssues(s.files, S.expectedExt);
const warnIcon = s => { const w = issues(s); return w.length ? ` <span class="warnicon" title="${esc(w.join('; '))}">⚠</span>` : ''; };
const score = s => G.scoreStudent(S.grades[s.id], S.questions, S.cfg, s.files.length > 0);
const fmt = n => String(+n.toFixed(4));

function renderHeader() {
  $('#labName').textContent = S.name ? '· ' + S.name : '';
  const subs = students.filter(s => s.files.length);
  const done = subs.filter(s => score(s).ungraded === 0).length;
  $('#progress').textContent = students.length ? `${done}/${subs.length} graded · ${students.length - subs.length} no submission` : '';
}

function renderSetup() {
  document.querySelectorAll('.drop').forEach(d => {
    const n = S.fileNames[d.dataset.kind];
    d.querySelector('em').textContent = n ? '✓ ' + n : '';
    d.classList.toggle('ok', !!n);
  });
  $('#name').value = S.name;
  $('#cTotal').value = S.cfg.cTotal;
  $('#aTotal').value = S.cfg.aTotal;
  $('#expectedExt').value = S.expectedExt;
  const p = perQ();
  $('#maxPts').textContent = `· ${p.n} graded · max ${fmt(maxPts())} pts`;
  // Same derivation as the gradebook legend: full = total ÷ graded questions, partial = ½, low = ¼, missing = 0.
  $('#ptsInfo').innerHTML = p.n ? `
    <p class="muted small">Each graded question is worth <b>${fmt(p.c)}</b> completeness (${fmt(S.cfg.cTotal)} ÷ ${p.n}) and
      <b>${fmt(p.a)}</b> accuracy (${fmt(S.cfg.aTotal)} ÷ ${p.n}). Voiding a question spreads its points over the rest.</p>
    <table class="legend"><tr><th></th>${G.TIERS.map((t, i) => `<th class="${TIER_CSS[i]}">${t.name}</th>`).join('')}</tr>
      <tr><td>Completeness</td>${G.TIERS.map(t => `<td>${fmt(p.c * t.mult)}</td>`).join('')}</tr>
      <tr><td>Accuracy</td>${G.TIERS.map(t => `<td>${fmt(p.a * t.mult)}</td>`).join('')}</tr></table>` : '';
  $('#canvasCol').innerHTML = '<option value="">— load roster, then pick the assignment —</option>' +
    (roster ? roster.assignmentCols.map(c => `<option value="${c.index}" ${S.canvasCol === c.index ? 'selected' : ''}>${esc(c.name)}</option>`).join('') : '');
  $('#qTable tbody').innerHTML = S.questions.map((q, i) => `
    <tr data-i="${i}">
      <td><input data-f="label" value="${esc(q.label)}" size="5"></td>
      <td class="prompt" title="${esc(q.prompt)}">${esc(q.prompt)}</td>
      <td><input type="checkbox" data-f="void" ${q.void ? 'checked' : ''}></td>
      <td><input type="checkbox" data-f="fullCredit" ${q.fullCredit ? 'checked' : ''}></td>
      <td><button data-del="${i}" title="Remove">✕</button></td>
    </tr>`).join('');
}

const tierOf = (s, q, kind) => G.cellTier(S.grades[s.id], q, kind, s.files.length > 0);

function chip(s, q, i, extra = '') {
  const c = tierOf(s, q, 'c'), a = tierOf(s, q, 'a');
  return `<span class="chip ${q.void ? 'void' : ''} ${extra}" data-q="${i}" title="Q${esc(q.label)}"><i class="${TIER_CSS[c] || ''}"></i><i class="${TIER_CSS[a] || ''}"></i></span>`;
}

function renderStudents() {
  $('#students').innerHTML = students.map((s, i) => {
    const sc = score(s);
    return `<div class="stu ${i === cur.s ? 'on' : ''} ${s.files.length ? '' : 'none'}" data-s="${i}">
      <div class="nm">${esc(s.name)}${warnIcon(s)} ${s.late ? '<span class="badge late">LATE</span>' : ''}${s.unmatched ? '<span class="badge">not in roster</span>' : ''}</div>
      <div class="meta">${s.files.length ? `${fmt(sc.total)} / ${fmt(maxPts())}${sc.ungraded ? ` · ${sc.ungraded} left` : ' ✓'}` : 'no submission'}</div>
      <div class="strip">${S.questions.map((q, j) => chip(s, q, j)).join('')}</div>
    </div>`;
  }).join('') || '<p class="muted pad">No students yet.</p>';
}

function renderPanel() {
  const s = students[cur.s], q = S.questions[cur.q];
  if (!s || !q) { $('#panel').innerHTML = '<p class="muted pad">Set up questions in Setup.</p>'; return; }
  const g = S.grades[s.id] || {}, sc = score(s);
  const btns = kind => G.TIERS.map((t, i) => `<button class="tier ${TIER_CSS[i]} ${tierOf(s, q, kind) === i ? 'sel' : ''}" data-kind="${kind}" data-t="${i}">${i + 1} · ${t.name}</button>`).join('');
  const locked = q.void ? 'Voided — not graded.' : q.fullCredit ? 'Full credit for everyone.' : !s.files.length ? 'No submission — counted as missing.' : '';
  $('#panel').innerHTML = `
    <div class="pad">
      <div class="who">${esc(s.name)} ${s.late ? '<span class="badge late">LATE</span>' : ''}</div>
      ${issues(s).map(w => `<div class="fmtwarn">⚠ ${esc(w)}</div>`).join('')}
      <div class="muted">${fmt(sc.total)} / ${fmt(maxPts())} · completeness ${fmt(sc.comp)} · accuracy ${fmt(sc.acc)}</div>
      <div class="strip big">${S.questions.map((qq, j) => chip(s, qq, j, j === cur.q ? 'cur' : '')).join('')}</div>
      <h2>Q${esc(q.label)}</h2>
      ${q.prompt ? `<p class="prompt">${esc(q.prompt)}</p>` : ''}
      ${locked ? `<p class="locked">${locked}</p>` : `
        <div class="lbl">Completeness <span class="muted">— key <kbd>1</kbd>–<kbd>4</kbd> sets both</span></div><div class="tiers">${btns('c')}</div>
        <div class="lbl">Accuracy <span class="muted">— <kbd>⇧</kbd>+<kbd>1</kbd>–<kbd>4</kbd></span></div><div class="tiers">${btns('a')}</div>`}
      ${formulaSection(q)}
      <details open><summary>Answer key</summary><pre class="key">${esc(keyChunks ? keyChunks.get(q.label) || '(no key text for this label)' : 'Load the key .docx in Setup.')}</pre></details>
      <div class="lbl">Notes for this student</div>
      <textarea id="note" rows="3" placeholder="e.g. 1e(ii) incomplete; collab: …">${esc(g.note || '')}</textarea>
      <div class="row"><button id="allFull">Fill remaining with full marks (F)</button></div>
      <p class="muted small"><kbd>↑</kbd><kbd>↓</kbd> question · <kbd>←</kbd><kbd>→</kbd> student · <kbd>F</kbd> fill remaining full</p>
    </div>`;
}

// Student's Excel formulas for this question, checked against the functions and numbers in the key.
function formulaSection(q) {
  if (!subKinds.has('xlsx') && !subKinds.has('xlsm')) {
    return subKinds.size ? '<div class="lbl">Excel formulas</div><p class="muted small">Not an Excel file — no formulas to check.</p>' : '';
  }
  const lab = q.label.toLowerCase(), n = (lab.match(/^\d+/) || [''])[0];
  const keyText = keyChunks ? keyChunks.get(q.label) : '';
  const keyFns = G.keyFunctions(keyText);
  const item = (f, showQ) => {
    // In the all-formulas list, check each formula against the key for the question it sits under.
    const kt = showQ && f.label && keyChunks ? keyChunks.get(f.label) : keyText;
    const c = G.checkFormula(f.formula, f.result, kt);
    const tags = [
      ...c.fns.map(fn => `<span class="tag ${c.sameFns.includes(fn) ? 'ok' : ''}">${c.sameFns.includes(fn) ? '✓ ' : ''}${esc(fn)}</span>`),
      c.constant ? '<span class="tag warn">typed value, no work shown</span>' : '',
      c.match != null ? `<span class="tag ok">≈ key ${esc(c.match)}</span>` : '',
    ].join('');
    return `<li data-addr="${esc(f.addr)}" title="Show in sheet">${showQ ? `<span class="muted">Q${esc(f.label || f.sec || '?')}</span> ` : ''}<b>${esc(f.addr.split('!').pop())}</b> <code>=${esc(f.formula)}</code> → ${esc(f.text)}<div>${tags}</div></li>`;
  };
  const mine = formulas.filter(f => (f.label ? f.label === lab : f.sec === n));
  return `<div class="lbl">Excel formulas for Q${esc(q.label)}${keyFns.length ? ` <span class="muted">— key uses ${keyFns.map(esc).join(', ')}</span>` : ''}</div>
    ${!formulas.length ? '<p class="warn small">No formulas anywhere in this workbook — values were typed in.</p>'
      : mine.length ? `<ul class="formulas">${mine.map(f => item(f, false)).join('')}</ul>`
      : '<p class="muted small">None under this question’s label (check “All formulas” below).</p>'}
    ${formulas.length ? `<details><summary>All formulas in this workbook (${formulas.length})</summary><ul class="formulas">${formulas.map(f => item(f, true)).join('')}</ul></details>` : ''}`;
}

// ---------- Submission rendering ----------
function revokeUrls() { objectUrls.forEach(URL.revokeObjectURL); objectUrls = []; }
function blobUrl(data, type) { const u = URL.createObjectURL(new Blob([data], { type })); objectUrls.push(u); return u; }

// A label at the start of text, kept only if it names a question in this assignment ("2x" in "2 x 3" is not one).
function labelIn(text) {
  const m = G.matchLabel(text);
  if (!m) return null;
  const ok = m.label ? S.questions.some(q => q.label.toLowerCase() === m.label)
    : S.questions.some(q => (q.label.match(/^\d+/) || [])[0] === m.sec);
  return ok ? m : null;
}

// Tag elements that start with a question label ("1a.", "Q1.") so the grader can jump to them.
function tagLabel(el, text) {
  const m = labelIn(text);
  if (!m) return false;
  if (m.label) el.dataset.label = m.label; else el.dataset.sec = m.sec;
  return true;
}

async function renderXlsx(buf, host) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  for (const ws of wb.worksheets) {
    const rows = new Map(); let maxCol = 1;
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      const cells = [];
      row.eachCell({ includeEmpty: false }, (c, col) => {
        if (c.isMerged && c.master !== c) return;
        if (c.text === '' && !c.formula) return;
        cells[col] = c; maxCol = Math.max(maxCol, col);
      });
      if (cells.length) rows.set(r, cells);
    });
    const imgs = new Map();
    for (const im of ws.getImages()) {
      const r = Math.floor(im.range.tl.nativeRow ?? im.range.tl.row) + 1;
      if (!imgs.has(r)) imgs.set(r, []);
      imgs.get(r).push(wb.getImage(im.imageId));
    }
    if (!rows.size && !imgs.size) continue;
    const order = [...new Set([...rows.keys(), ...imgs.keys()])].sort((a, b) => a - b);
    let html = `<thead><tr><th></th>${Array.from({ length: maxCol }, (_, i) => `<th>${G.colLetter(i + 1)}</th>`).join('')}</tr></thead><tbody>`;
    let prev = 0, curLabel = null, curSec = null;  // the question label a formula sits under
    for (const r of order) {
      if (r - prev > 2) html += `<tr class="gap"><td colspan="${maxCol + 1}">⋯ ${r - prev - 1} empty rows</td></tr>`;
      prev = r;
      const cells = rows.get(r);
      if (cells) {
        const first = cells.find(c => c && c.text.trim());
        const m = first && labelIn(first.text);
        if (m) { curSec = m.sec; curLabel = m.label; }
        html += `<tr data-row="${r}"><th>${r}</th>`;
        for (let c = 1; c <= maxCol; c++) {
          const cell = cells[c];
          if (!cell) { html += '<td></td>'; continue; }
          const f = cell.formula ? cell.formula.replace(/_xlfn\./g, '') : '';
          const addr = `${ws.name}!${cell.address}`;
          if (f) formulas.push({ addr, formula: f, result: cell.result, text: cell.text, label: curLabel, sec: curSec });
          html += `<td${f ? ` class="f" data-addr="${esc(addr)}"` : ''}><span>${esc(cell.text)}</span>${f ? `<code>=${esc(f)}</code>` : ''}</td>`;
        }
        html += '</tr>';
      }
      for (const im of imgs.get(r) || []) {
        const ext = String(im.extension).toLowerCase().replace('jpg', 'jpeg');
        html += /^(png|jpeg|gif|webp|bmp)$/.test(ext)
          ? `<tr class="img"><th>${r}</th><td colspan="${maxCol}"><img src="${blobUrl(im.buffer, 'image/' + ext)}" alt="Embedded image at row ${r}"></td></tr>`
          : `<tr class="img"><th>${r}</th><td colspan="${maxCol}" class="muted">[${esc(ext)} image — open the original file to view]</td></tr>`;
      }
    }
    const t = document.createElement('table');
    t.className = 'sheet';
    t.innerHTML = html + '</tbody>';
    t.querySelectorAll('tr[data-row]').forEach(tr => {
      for (const sp of tr.querySelectorAll('td > span')) if (sp.textContent.trim()) { tagLabel(tr, sp.textContent); break; }
    });
    if (wb.worksheets.length > 1) host.insertAdjacentHTML('beforeend', `<div class="sheetname">Sheet: ${esc(ws.name)}</div>`);
    host.appendChild(t);
  }
}

async function renderDocx(buf, host) {
  const r = await mammoth.convertToHtml({ arrayBuffer: buf });
  const div = document.createElement('div');
  div.className = 'docx';
  div.innerHTML = r.value;
  div.querySelectorAll('a[href]').forEach(a => { if (!/^(https?:|mailto:|#)/i.test(a.getAttribute('href'))) a.removeAttribute('href'); });
  div.querySelectorAll('p,li,h1,h2,h3,h4').forEach(el => tagLabel(el, el.textContent));
  host.appendChild(div);
}

let renderToken = 0;
async function renderSubmission() {
  const s = students[cur.s], host = $('#submission');
  const token = ++renderToken;
  revokeUrls();
  host.innerHTML = '';
  formulas = [];
  subKinds = new Set(s ? s.files.map(p => G.parseCanvasFilename(p).ext) : []);
  if (!s) return;
  if (!s.files.length) { host.innerHTML = '<p class="muted pad">No submission — every question counts as missing (black).</p>'; return; }
  host.innerHTML = '<p class="muted pad">Loading…</p>';
  const frag = document.createElement('div');
  for (const path of s.files) {
    const f = G.parseCanvasFilename(path);
    const sec = document.createElement('div');
    const bad = G.formatIssues([path], S.expectedExt)[0];
    sec.innerHTML = `<div class="file${bad ? ' bad' : ''}">${bad ? `⚠ ${esc(f.original)} — ${esc(bad)}; showing its content anyway` : esc(f.original)}</div>`;
    frag.appendChild(sec);
    try {
      const buf = await zip.file(path).async('arraybuffer');
      if (f.ext === 'xlsx' || f.ext === 'xlsm') await renderXlsx(buf, sec);
      else if (f.ext === 'docx') await renderDocx(buf, sec);
      else if (f.ext === 'pdf') sec.insertAdjacentHTML('beforeend', `<iframe class="pdf" src="${blobUrl(buf, 'application/pdf')}"></iframe>`);
      else if (/^(png|jpe?g|gif|webp)$/.test(f.ext)) sec.insertAdjacentHTML('beforeend', `<img src="${blobUrl(buf, 'image/' + f.ext.replace('jpg', 'jpeg'))}" alt="">`);
      else sec.insertAdjacentHTML('beforeend', `<p class="pad">Can't preview .${esc(f.ext)} — <a download="${esc(f.original)}" href="${blobUrl(buf, 'application/octet-stream')}">download it</a>.</p>`);
    } catch (e) {
      console.error(e);
      sec.insertAdjacentHTML('beforeend', `<p class="pad error">Could not render this file (${esc(e.message)}).</p>`);
    }
    if (token !== renderToken) return;  // user moved on while this was loading
  }
  host.innerHTML = '';
  host.appendChild(frag);
  host.scrollTop = 0;
  renderPanel();  // formulas are known now
  focusQuestion();
}

function focusQuestion() {
  const q = S.questions[cur.q], host = $('#submission');
  host.querySelectorAll('.hit').forEach(e => e.classList.remove('hit'));
  if (!q) return;
  const lab = q.label.toLowerCase(), n = (lab.match(/^\d+/) || [''])[0];
  const el = host.querySelector(`[data-label="${CSS.escape(lab)}"]`) || (n && host.querySelector(`[data-sec="${CSS.escape(n)}"]`));
  if (el) { el.classList.add('hit'); el.scrollIntoView({ block: 'start', behavior: 'smooth' }); }
}

function renderGrade(full) {
  renderStudents();
  renderPanel();
  renderHeader();
  if (full) { renderSubmission(); $('#students .stu.on')?.scrollIntoView({ block: 'nearest' }); }
}

function renderOverview() {
  const qs = S.questions;
  const head = `<tr><th>Student</th>${qs.map(q => `<th class="${q.void ? 'void' : ''}">${esc(q.label)}</th>`).join('')}<th></th><th>Score</th><th>Notes</th></tr>`;
  const body = students.map((s, i) => ['c', 'a'].map((kind, k) => `<tr class="${k ? 'acc' : ''}">
      ${k ? '' : `<td rowspan="2" class="nm"><a href="#" data-s="${i}">${esc(s.name)}</a>${warnIcon(s)}${s.late ? ' <span class="badge late">LATE</span>' : ''}</td>`}
      ${qs.map((q, j) => { const t = tierOf(s, q, kind); return `<td class="cell ${TIER_CSS[t] || ''} ${q.void ? 'void' : ''}" data-s="${i}" data-q="${j}">${t == null ? '' : fmt((kind === 'c' ? perQ().c : perQ().a) * G.TIERS[t].mult)}</td>`; }).join('')}
      <td class="muted">${kind === 'c' ? 'Completeness' : 'Accuracy'}</td>
      ${k ? '' : `<td rowspan="2"><b>${fmt(score(s).total)}</b></td><td rowspan="2" class="muted">${esc((S.grades[s.id] || {}).note || '')}</td>`}
    </tr>`).join('')).join('');
  $('#overview').innerHTML = students.length ? `<table class="ov">${head}${body}</table>` : '<p class="muted">No students yet.</p>';
}

// ---------- Grading actions ----------
const gradeOf = s => (S.grades[s.id] = S.grades[s.id] || { q: {} });

function setTier(kinds, t) {
  const s = students[cur.s], q = S.questions[cur.q];
  if (!s || !q || q.void || q.fullCredit || !s.files.length) return;
  const g = gradeOf(s);
  g.q[q.label] = g.q[q.label] || { c: null, a: null };
  for (const k of kinds) g.q[q.label][k] = t;
  save();
  renderGrade(false);
}

function fillFull() {
  const s = students[cur.s];
  if (!s || !s.files.length) return;
  const g = gradeOf(s);
  for (const q of S.questions) {
    const e = g.q[q.label] = g.q[q.label] || { c: null, a: null };
    if (e.c == null) e.c = 0;
    if (e.a == null) e.a = 0;
  }
  save();
  renderGrade(false);
}

function moveStudent(d) {
  // Arrow keys skip non-submitters (they need no grading); clicking still selects them.
  let i = cur.s;
  do { i += d; } while (students[i] && !students[i].files.length);
  if (students[i]) { cur.s = i; renderGrade(true); }
}
function moveQuestion(d) {
  const i = cur.q + d;
  if (S.questions[i]) { cur.q = i; renderPanel(); focusQuestion(); }
}

function download(name, data, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([data], { type }));
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

const safeName = () => (S.name || 'Lab').replace(/[^\w #.-]+/g, '_');

async function exportXlsx() {
  const left = students.filter(s => s.files.length && score(s).ungraded).length;
  if (left && !confirm(`${left} student(s) still have ungraded questions (left blank, counted as 0). Export anyway?`)) return;
  const wb = G.buildGradebook(ExcelJS, { name: S.name, questions: S.questions, cfg: S.cfg, students, grades: S.grades });
  download(`${safeName()} gradebook.xlsx`, await wb.xlsx.writeBuffer(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

function exportCsv() {
  if (!roster) return toast('Load the Canvas roster .csv in Setup first.');
  if (S.canvasCol == null) { showView('setup'); return toast('Pick the Canvas column for this assignment in Setup.'); }
  // Only fully graded students get a number; blank cells leave the Canvas grade untouched on import.
  const totals = {};
  for (const s of students) if (!s.unmatched && score(s).ungraded === 0) totals[s.id] = score(s).total;
  const skipped = students.filter(s => s.files.length && score(s).ungraded).length;
  if (skipped) toast(`${skipped} partially graded student(s) left blank in the CSV.`, 5000);
  download(`${safeName()} canvas import.csv`, G.buildCanvasCsv(roster, S.canvasCol, totals), 'text/csv');
}

// ---------- Events ----------
document.querySelectorAll('nav button').forEach(b => b.onclick = () => showView(b.dataset.view));
document.querySelectorAll('.drop input').forEach(inp => inp.onchange = () => {
  if (inp.files[0]) onFile(inp.closest('.drop').dataset.kind, inp.files[0]);
  inp.value = '';
});
$('#name').oninput = e => { S.name = e.target.value; save(); renderHeader(); };
$('#cTotal').oninput = e => { S.cfg.cTotal = +e.target.value || 0; save(); };
$('#aTotal').oninput = e => { S.cfg.aTotal = +e.target.value || 0; save(); };
$('#expectedExt').oninput = e => { S.expectedExt = e.target.value.trim(); save(); };
$('#cTotal').onchange = $('#aTotal').onchange = $('#expectedExt').onchange = renderSetup;
$('#canvasCol').onchange = e => { S.canvasCol = e.target.value === '' ? null : +e.target.value; save(); };
$('#addQ').onclick = () => { S.questions.push({ label: '', prompt: '', void: false, fullCredit: false }); save(); renderSetup(); };
$('#resetQ').onclick = () => { if (confirm('Replace the question list with the one read from the files?')) { resetQuestions(); save(); renderSetup(); } };
$('#qTable').addEventListener('input', e => {
  const tr = e.target.closest('tr[data-i]'), f = e.target.dataset.f;
  if (!tr || !f) return;
  S.questions[+tr.dataset.i][f] = e.target.type === 'checkbox' ? e.target.checked : e.target.value.trim();
  save();
  if (e.target.type === 'checkbox') renderSetup();
});
$('#qTable').addEventListener('click', e => {
  const d = e.target.dataset.del;
  if (d != null) { S.questions.splice(+d, 1); save(); renderSetup(); }
});
$('#clearAll').onclick = async () => {
  if (!confirm('Clear all loaded files AND all grades for this lab? Export the gradebook first if you need it.')) return;
  localStorage.removeItem(KEY);
  await idb('readwrite', st => st.clear());
  location.reload();
};
$('#exportXlsx').onclick = () => exportXlsx().catch(e => { console.error(e); toast('Export failed: ' + e.message, 6000); });
$('#exportCsv').onclick = exportCsv;

$('#students').addEventListener('click', e => {
  const el = e.target.closest('[data-s]');
  if (!el) return;
  cur.s = +el.dataset.s;
  const c = e.target.closest('[data-q]');
  if (c) cur.q = +c.dataset.q;
  renderGrade(true);
});
$('#panel').addEventListener('click', e => {
  const b = e.target.closest('.tier');
  if (b) return setTier([b.dataset.kind], +b.dataset.t);
  const c = e.target.closest('[data-q]');
  if (c) { cur.q = +c.dataset.q; renderPanel(); focusQuestion(); }
  if (e.target.id === 'allFull') fillFull();
  const fa = e.target.closest('li[data-addr]');
  const td = fa && $('#submission').querySelector(`td[data-addr="${CSS.escape(fa.dataset.addr)}"]`);
  if (td) {
    $('#submission').querySelectorAll('.hit').forEach(x => x.classList.remove('hit'));
    td.classList.add('hit');
    td.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }
});
$('#panel').addEventListener('input', e => {
  if (e.target.id !== 'note') return;
  gradeOf(students[cur.s]).note = e.target.value;
  save();
});
$('#overview').addEventListener('click', e => {
  const el = e.target.closest('[data-s]');
  if (!el) return;
  e.preventDefault();
  cur.s = +el.dataset.s;
  if (el.dataset.q) cur.q = +el.dataset.q;
  showView('grade');
});

document.addEventListener('keydown', e => {
  if (currentView() !== 'grade' || e.target.matches('input, textarea, select') || e.metaKey || e.ctrlKey || e.altKey) return;
  const digit = /^Digit([1-4])$/.exec(e.code);
  if (digit) { e.preventDefault(); return setTier(e.shiftKey ? ['a'] : ['c', 'a'], +digit[1] - 1); }
  const k = e.key;
  if (k === 'ArrowDown') { e.preventDefault(); moveQuestion(1); }
  else if (k === 'ArrowUp') { e.preventDefault(); moveQuestion(-1); }
  else if (k === 'ArrowRight' || k === 'j') { e.preventDefault(); moveStudent(1); }
  else if (k === 'ArrowLeft' || k === 'k') { e.preventDefault(); moveStudent(-1); }
  else if (k === 'f' || k === 'F') { e.preventDefault(); fillFull(); }
});

restore().catch(e => { console.error(e); toast('Could not restore saved files: ' + e.message, 6000); showView('setup'); });
