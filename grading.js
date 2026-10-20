// Pure grading logic (no DOM). Classic script so index.html works from file://; also require()-able from Node for tests.

// Tier index -> multiplier of the question's points. Matches the professor's legend.
const TIERS = [
  { name: 'Full marks', mult: 1, argb: 'FF196B24' },
  { name: 'Partial / effort', mult: 0.5, argb: 'FFFFFF00' },
  { name: 'Low / no effort', mult: 0.25, argb: 'FFC00000' },
  { name: 'Missing', mult: 0, argb: 'FF000000' },
];

// Minimal RFC 4180 CSV parser (quoted fields, embedded commas/newlines, "" escapes).
function parseCsv(text) {
  const rows = []; let row = [], field = '', q = false;
  text = text.replace(/^﻿/, '');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') q = false;
      else field += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

const csvEscape = v => /[",\n\r]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);

// Canvas gradebook export: identity columns, then assignment columns named "Title (12345)".
function parseCanvasRoster(text) {
  const rows = parseCsv(text).filter(r => r.some(c => c.trim()));
  const header = rows[0];
  const idIdx = header.indexOf('ID');
  if (header[0] !== 'Student' || idIdx < 0) throw new Error('Not a Canvas gradebook export (expected "Student" and "ID" columns).');
  const assignmentCols = header.map((name, index) => ({ name, index })).filter(c => /\(\d+\)\s*$/.test(c.name));
  const identityCount = assignmentCols.length ? assignmentCols[0].index : header.length;
  const meta = rows.slice(1).filter(r => !r[idIdx]); // "Points Possible" and similar rows
  const students = rows.slice(1)
    .filter(r => r[idIdx] && r[0] !== 'Student, Test')
    .map(r => ({ id: r[idIdx].trim(), name: r[0].trim(), row: r }));
  return { header, identityCount, assignmentCols, meta, students };
}

// Canvas zip entry: "lastfirst[_LATE]_<userId>_<submissionId>_<original name>"
function parseCanvasFilename(path) {
  const base = path.split('/').pop();
  const m = base.match(/^(.+?)(_LATE)?_(\d+)_(\d+)_(.+)$/) || base.match(/^(.+?)(_LATE)?_(\d+)_(.+)$/);
  if (!m) return null;
  const original = m[5] || m[4];
  return { slug: m[1], late: !!m[2], userId: m[3], original, ext: (original.split('.').pop() || '').toLowerCase() };
}

// Merge roster (all students) with zip entries (submitters). Roster order first, unmatched submitters appended.
function buildStudents(roster, zipPaths) {
  const byId = new Map();
  for (const s of roster ? roster.students : []) byId.set(s.id, { id: s.id, name: s.name, late: false, files: [] });
  for (const p of zipPaths) {
    const f = parseCanvasFilename(p);
    if (!f) continue;
    if (!byId.has(f.userId)) byId.set(f.userId, { id: f.userId, name: f.slug, late: false, files: [], unmatched: !!roster });
    const s = byId.get(f.userId);
    s.files.push(p);
    s.late = s.late || f.late;
  }
  return [...byId.values()];
}

// Question labels ("1a.") from the assignment instructions' paragraphs.
function parseQuestionLabels(paragraphs) {
  const seen = new Set(), out = [];
  for (const p of paragraphs) {
    const m = p.trim().match(/^(\d+[a-z])\b/);
    if (m && !seen.has(m[1])) { seen.add(m[1]); out.push({ label: m[1], prompt: p.trim() }); }
  }
  return out;
}

// Question label at the start of a cell/paragraph: "1a.", "Q1a", "1E", "Question 2b" -> {label:'1a', sec:'1'}.
// Section-only when unambiguous ("Q1.", "Question 3", "4.") -> {label:null, sec}. Work like "19.8=x-45" is not a label.
function matchLabel(text) {
  const t = String(text);
  let m = t.match(/^\s*(?:Q(?:uestion)?\s*#?\s*)?(\d+)\s*([a-z])(?![a-z0-9])/i);
  if (m) return { label: (m[1] + m[2]).toLowerCase(), sec: m[1] };
  m = t.match(/^\s*Q(?:uestion)?\s*#?\s*(\d+)(?!\d|\.\d)/i) || t.match(/^\s*(\d+)\s*[.)]\s*$/);
  return m ? { label: null, sec: m[1] } : null;
}

// Key text per label. "1e(i)."/"1e(ii)." fold into "1e"; text under a "Question N:" heading before the
// first label (e.g. Q6's verification lines) is kept as that section's preamble for every Nx label.
function parseKeyChunks(paragraphs) {
  const chunks = {}, preamble = {};
  let cur = null, sec = null;
  for (const raw of paragraphs) {
    const p = raw.trim();
    if (!p) continue;
    const h = p.match(/^Question\s+(\d+)\b/i);
    if (h) { sec = h[1]; cur = null; preamble[sec] = [p]; continue; }
    const m = p.match(/^(\d+[a-z])(\([ivx]+\))?\.?(\s|$)/);
    if (m) { cur = m[1]; sec = cur.match(/^\d+/)[0]; }
    if (cur) (chunks[cur] = chunks[cur] || []).push(p);
    else if (sec) preamble[sec].push(p);
  }
  const get = label => {
    const n = (label.match(/^\d+/) || [])[0];
    return [...(preamble[n] || []), ...(chunks[label] || [])].join('\n');
  };
  return { labels: Object.keys(chunks), get };
}

// grade = { q: { [label]: { c: tierIdx|null, a: tierIdx|null } }, note }
function cellTier(grade, question, kind, hasSubmission) {
  if (question.void) return null;
  if (question.fullCredit) return 0;
  if (!hasSubmission) return 3;
  const g = grade && grade.q && grade.q[question.label];
  return g && g[kind] != null ? g[kind] : null;
}

function scoreStudent(grade, questions, cfg, hasSubmission) {
  let comp = 0, acc = 0, ungraded = 0;
  for (const q of questions) {
    if (q.void) continue;
    const c = cellTier(grade, q, 'c', hasSubmission), a = cellTier(grade, q, 'a', hasSubmission);
    if (c == null || a == null) ungraded++;
    comp += c == null ? 0 : cfg.cPts * TIERS[c].mult;
    acc += a == null ? 0 : cfg.aPts * TIERS[a].mult;
  }
  return { comp, acc, total: comp + acc, ungraded };
}

// ---------- Excel formula checks ----------
// Legacy/compatibility names -> modern names, so a student's NORMSINV counts as the key's NORM.S.INV.
const FN_ALIASES = {
  NORMSINV: 'NORM.S.INV', NORMSDIST: 'NORM.S.DIST', NORMINV: 'NORM.INV', NORMDIST: 'NORM.DIST',
  STDEV: 'STDEV.S', STDEVP: 'STDEV.P', VAR: 'VAR.S', VARP: 'VAR.P', TINV: 'T.INV.2T', TDIST: 'T.DIST.2T',
  PERCENTILE: 'PERCENTILE.INC', QUARTILE: 'QUARTILE.INC', MODE: 'MODE.SNGL',
};
const normFn = name => { const u = name.toUpperCase().replace(/^_XL(FN|WS)\./, ''); return FN_ALIASES[u] || u; };
const FN_CALL = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;

const formulaFunctions = formula => [...new Set([...String(formula).matchAll(FN_CALL)].map(m => normFn(m[1])))];

// Functions the key mentions, e.g. “=norm.s.inv(.2)” or “=Norm.S.Dist”.
const keyFunctions = text => [...new Set([...String(text).matchAll(/=\s*([A-Za-z][A-Za-z0-9]*(?:\.[A-Za-z0-9]+)+|[A-Za-z][A-Za-z0-9]*(?=\s*\())/g)].map(m => normFn(m[1])))];

const keyNumbers = text => [...String(text).replace(/[−–]/g, '-').matchAll(/-?\d*\.?\d+/g)].map(m => m[0]);

// What one student formula tells the grader, relative to the key text for that question.
function checkFormula(formula, result, keyText) {
  const fns = formulaFunctions(formula), keyFns = keyFunctions(keyText);
  const refs = /\$?[A-Z]{1,3}\$?\d+/.test(String(formula).replace(FN_CALL, ''));
  const constant = !fns.length && !refs && /^\s*-?[\d.]+\s*$/.test(formula);
  let match = null;
  if (typeof result === 'number') {
    for (const k of keyNumbers(keyText)) {
      // Compare at the key's precision (-0.8416 ≈ -0.84); whole numbers must match exactly (0.8849 is not "1").
      const dp = (k.split('.')[1] || '').length;
      if (Math.abs(result - parseFloat(k)) <= (dp ? 0.5 * 10 ** -dp : 0) + 1e-9) { match = k; break; }
    }
  }
  return { fns, keyFns, sameFns: fns.filter(f => keyFns.includes(f)), constant, refs, match };
}

const colLetter = n => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };

// Same layout as the professor's gradebook: 2 rows per student (Completeness / Accuracy),
// question columns B.., label, SUM per row, merged total, notes, and a legend block.
function buildGradebook(ExcelJS, { name, questions, cfg, students, grades }) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet((name || 'Grades').replace(/[\\/?*[\]:]/g, ' ').slice(0, 31));
  const nQ = questions.length;
  const C = { first: 2, last: nQ + 1, label: nQ + 2, sum: nQ + 3, total: nQ + 4, notes: nQ + 5, legend: nQ + 6 };
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const range = r => `${colLetter(C.first)}${r}:${colLetter(C.last)}${r}`;

  ws.getRow(1).values = ['Student ID', ...questions.map(q => 'Q' + q.label)];
  ws.getCell(1, C.total).value = 'Score';
  ws.getCell(1, C.notes).value = 'Notes';
  ws.getColumn(1).width = 26;
  for (let c = C.first; c <= C.last; c++) ws.getColumn(c).width = 6;
  ws.getColumn(C.label).width = 13;
  ws.getColumn(C.notes).width = 40;
  questions.forEach((q, i) => { if (q.void) ws.getCell(1, C.first + i).note = 'Voided'; });

  students.forEach((s, i) => {
    const r = 2 + i * 2, g = grades[s.id] || {}, has = s.files.length > 0;
    ws.getCell(r, 1).value = s.name;
    [['c', 'Completeness', cfg.cPts], ['a', 'Accuracy', cfg.aPts]].forEach(([kind, label, pts], k) => {
      const row = r + k;
      questions.forEach((q, j) => {
        const t = cellTier(g, q, kind, has);
        if (t == null) return;
        const cell = ws.getCell(row, C.first + j);
        cell.value = pts * TIERS[t].mult;
        cell.fill = fill(TIERS[t].argb);
      });
      ws.getCell(row, C.label).value = label;
      ws.getCell(row, C.sum).value = { formula: `SUM(${range(row)})`, result: scoreStudent(g, questions, cfg, has)[kind === 'c' ? 'comp' : 'acc'] };
    });
    ws.mergeCells(r, C.total, r + 1, C.total);
    const tot = ws.getCell(r, C.total);
    tot.value = { formula: `SUM(${colLetter(C.sum)}${r}:${colLetter(C.sum)}${r + 1})`, result: scoreStudent(g, questions, cfg, has).total };
    tot.alignment = { vertical: 'middle' };
    const note = [s.late && !/late/i.test(g.note || '') ? 'late' : '', has ? '' : 'no submission', g.note || ''].filter(Boolean).join('; ');
    if (note) ws.getCell(r, C.notes).value = note;
  });

  // Legend, like the original sheet's right-hand block.
  ws.getCell(3, C.legend + 1).value = 'completeness';
  ws.getCell(3, C.legend + 2).value = 'accuracy';
  TIERS.forEach((t, i) => {
    const r = 4 + i;
    [t.name, cfg.cPts * t.mult, cfg.aPts * t.mult].forEach((v, k) => {
      const cell = ws.getCell(r, C.legend + k);
      cell.value = v;
      cell.fill = fill(t.argb);
      if (i >= 2) cell.font = { color: { argb: 'FFFFFFFF' } };
    });
  });
  return wb;
}

// Canvas import CSV: identity columns + the chosen assignment column, total per student.
function buildCanvasCsv(roster, colIndex, totals) {
  const pick = r => [...r.slice(0, roster.identityCount), r[colIndex] ?? ''];
  const lines = [pick(roster.header), ...roster.meta.map(pick)];
  for (const s of roster.students) {
    const r = pick(s.row);
    if (s.id in totals) r[r.length - 1] = String(+totals[s.id].toFixed(4));
    lines.push(r);
  }
  return lines.map(r => r.map(csvEscape).join(',')).join('\r\n') + '\r\n';
}

// One global `G` in the browser (shared by app.js); module.exports in Node.
const G = {
  TIERS, parseCsv, parseCanvasRoster, parseCanvasFilename, buildStudents, parseQuestionLabels,
  parseKeyChunks, cellTier, scoreStudent, colLetter, buildGradebook, buildCanvasCsv,
  formulaFunctions, keyFunctions, checkFormula, matchLabel,
};
if (typeof module !== 'undefined') module.exports = G;
