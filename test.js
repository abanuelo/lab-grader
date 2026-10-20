// node test.js — synthetic data only (this repo is public; never commit real student files).
const assert = require('assert');
const ExcelJS = require('./vendor/exceljs.min.js');
const G = require('./grading.js');

// Filenames in the shapes Canvas produces.
const f = G.parseCanvasFilename('submissions/doejane_LATE_1001_5001_Doe, Jane - Lab 2.xlsx');
assert.deepStrictEqual(f, { slug: 'doejane', late: true, userId: '1001', original: 'Doe, Jane - Lab 2.xlsx', ext: 'xlsx' });
assert.strictEqual(G.parseCanvasFilename('roesam_1002_5002_320_Lab#2.docx').late, false);
assert.strictEqual(G.parseCanvasFilename('.DS_Store'), null);

// Canvas roster: quoted names, points-possible row, test student, assignment columns.
const csv = 'Student,ID,SIS User ID,SIS Login ID,Section,Lab 1 (900),Lab 2 (901),Current Score\r\n' +
  '    Points Possible,,,,,15,15,\r\n"Doe, Jane",1001,s1,jdoe,PSY 320,14,,\r\n"Roe, Sam",1002,s2,sroe,PSY 320,12,,\r\n' +
  '"Poe, Al",1003,s3,apoe,PSY 320,15,,\r\n"Student, Test",9999,,,PSY 320,,,\r\n';
const roster = G.parseCanvasRoster(csv);
assert.deepStrictEqual(roster.students.map(s => s.id), ['1001', '1002', '1003']);
assert.strictEqual(roster.identityCount, 5);
assert.deepStrictEqual(roster.assignmentCols.map(c => c.index), [5, 6]);

const students = G.buildStudents(roster, ['x/doejane_LATE_1001_5001_a.xlsx', 'x/roesam_1002_5002_b.docx', 'x/newkid_1004_5004_c.xlsx']);
assert.deepStrictEqual(students.map(s => [s.id, s.files.length, s.late]), [['1001', 1, true], ['1002', 1, false], ['1003', 0, false], ['1004', 1, false]]);
assert.ok(students[3].unmatched);

// Labels from instructions; key chunks fold 1e(i)/(ii) and carry a section preamble.
const qs = G.parseQuestionLabels(['Intro', '1a. Compute z.', '1e. (i) Raw score', '(ii) cutoff', '2a. Compare', '1a. dup']);
assert.deepStrictEqual(qs.map(q => q.label), ['1a', '1e', '2a']);
const key = G.parseKeyChunks(['Question 1: z', '1a. z = 2', '1e(i). 64.8', '1e(ii). 34.92', 'Question 6: SPSS', 'Verification: M = 7.45', '6b. z = 1.86']);
assert.strictEqual(key.get('1e'), 'Question 1: z\n1e(i). 64.8\n1e(ii). 34.92');
assert.strictEqual(key.get('6a'), 'Question 6: SPSS\nVerification: M = 7.45');

// Scoring reproduces the gradebook legend: 20 graded questions, max 15, 2 voided.
const cfg = { cTotal: 10, aTotal: 5 };  // per question = total / graded questions (10/20, 5/20)
const labels = '1a 1b 1c 1d 1e 2a 2b 2c 3a 3b 3c 4a 4b 4c 4d 5a 5b 6a 6b 6c 7a 7b'.split(' ');
const questions = labels.map(label => ({ label, void: label.startsWith('7') }));
const allFull = { q: Object.fromEntries(labels.map(l => [l, { c: 0, a: 0 }])) };
assert.strictEqual(G.scoreStudent(allFull, questions, cfg, true).total, 15);
const oneYellow = structuredClone(allFull); oneYellow.q['4c'].a = 1;          // accuracy partial on one question
assert.strictEqual(G.scoreStudent(oneYellow, questions, cfg, true).total, 14.875);
const twoBlack = structuredClone(allFull); twoBlack.q['1b'] = { c: 3, a: 3 }; twoBlack.q['5b'] = { c: 3, a: 3 };
assert.strictEqual(G.scoreStudent(twoBlack, questions, cfg, true).total, 13.5);
assert.strictEqual(G.scoreStudent({}, questions, cfg, false).total, 0);       // no submission = all missing
assert.strictEqual(G.scoreStudent({}, questions, cfg, true).ungraded, 20);
const fc = questions.map(q => q.label === '1a' ? { ...q, fullCredit: true } : q);
assert.strictEqual(G.scoreStudent({}, fc, cfg, true).total, 0.75);           // everyone gets 1a

// Voiding spreads points over the remaining questions: max stays 15 (professor's =10/20 legend).
const moreVoid = questions.map(q => ['6c', '5b'].includes(q.label) ? { ...q, void: true } : q);
assert.deepStrictEqual(G.perQuestion(cfg, moreVoid), { n: 18, c: 10 / 18, a: 5 / 18 });
assert.ok(Math.abs(G.scoreStudent(allFull, moreVoid, cfg, true).total - 15) < 1e-9);

// File-format warnings (content is still shown).
assert.deepStrictEqual(G.formatIssues(['x/doejane_1001_5001_Lab2.xlsx'], 'xlsx'), []);
assert.deepStrictEqual(G.formatIssues(['x/doejane_1001_5001_Lab2.docx'], 'xlsx'), ['.docx submitted, expected .xlsx']);
assert.deepStrictEqual(G.formatIssues(['x/a_1_2_A.xlsx', 'x/a_1_3_B.pdf'], '.xlsx, xlsm'), ['2 files submitted (expected one)', '.pdf submitted, expected .xlsx or .xlsm']);

// Labels in student work: real labels match, work that merely starts with a number does not.
const L = t => JSON.stringify(G.matchLabel(t));
assert.strictEqual(L('1a. Identify'), '{"label":"1a","sec":"1"}');
assert.strictEqual(L('1E'), '{"label":"1e","sec":"1"}');
assert.strictEqual(L('Q2b'), '{"label":"2b","sec":"2"}');
assert.strictEqual(L('1e(ii)'), '{"label":"1e","sec":"1"}');
assert.strictEqual(L('Q1.'), '{"label":null,"sec":"1"}');
assert.strictEqual(L('Question 3'), '{"label":null,"sec":"3"}');
assert.strictEqual(L('4.'), '{"label":null,"sec":"4"}');
for (const t of ['19.8=x-45-', '1.65=x-45/12', '12 SD', '2 standard deviations', '64.8', '-10.08', '(ii)', '4x4']) assert.strictEqual(G.matchLabel(t), null, t);

// Formula checks: legacy names alias to modern ones, constants flagged, results matched at key precision.
assert.deepStrictEqual(G.formulaFunctions('_xlfn.NORM.S.INV(0.2)'), ['NORM.S.INV']);
assert.deepStrictEqual(G.formulaFunctions('NORMSDIST(0-0.84)+ABS(B2)'), ['NORM.S.DIST', 'ABS']);
assert.deepStrictEqual(G.keyFunctions('Below: “=Norm.S.Dist” = .9772; cutoff (“=norm.s.inv(.2)”) X = 45 + z(σ); =SUM(A1:A3)'), ['NORM.S.DIST', 'NORM.S.INV', 'SUM']);
const keyText = 'Cutoff (“=norm.s.inv(.2)”) → z ≈ −0.84. X = 45 + (−0.84)(12) = 34.92';
let chk = G.checkFormula('NORMSINV(0.2)', -0.8416212335729145, keyText);
assert.deepStrictEqual([chk.sameFns, chk.constant, chk.match], [['NORM.S.INV'], false, '-0.84']);
chk = G.checkFormula('-10.08+45', 34.92, keyText);
assert.deepStrictEqual([chk.fns, chk.constant, chk.match], [[], false, '34.92']);
chk = G.checkFormula('64.8', 64.8, keyText);
assert.deepStrictEqual([chk.constant, chk.match], [true, null]);
assert.strictEqual(G.checkFormula('(B6-B7)/B8', 2, 'z = 2').refs, true);
assert.strictEqual(G.checkFormula('0.8849', 0.8849, 'z = 1.20: “=norm.s.dist(1.20)” = .8849').match, '.8849');  // not "1"
assert.strictEqual(G.checkFormula('STANDARDIZE(57,45,12)', 1, 'z = 12 / 12 = 1').match, '1');

// Gradebook export: layout, fills, formulas, merged total, notes.
(async () => {
  const grades = { 1001: oneYellow, 1002: { ...twoBlack, note: 'wow' } };
  const wb = G.buildGradebook(ExcelJS, { name: 'Lab 2', questions, cfg, students: students.slice(0, 3), grades });
  const back = new ExcelJS.Workbook(); await back.xlsx.load(await wb.xlsx.writeBuffer());
  const ws = back.worksheets[0];
  assert.strictEqual(ws.getCell('B1').value, 'Q1a');
  assert.strictEqual(ws.getCell('W1').value, 'Q7b');
  assert.strictEqual(ws.getCell('X2').value, 'Completeness');
  assert.strictEqual(ws.getCell('X3').value, 'Accuracy');
  assert.strictEqual(ws.getCell('Y2').formula, 'SUM(B2:W2)');
  assert.strictEqual(ws.getCell('Z2').formula, 'SUM(Y2:Y3)');
  assert.strictEqual(ws.getCell('Z2').result, 14.875);
  assert.ok(ws.getCell('Z3').isMerged);
  assert.strictEqual(ws.getCell('O3').value, 0.125);
  assert.strictEqual(ws.getCell('O3').fill.fgColor.argb, 'FFFFFF00');
  assert.strictEqual(ws.getCell('B2').fill.fgColor.argb, 'FF196B24');
  assert.strictEqual(ws.getCell('V2').value, null);                          // voided Q7 left blank
  assert.strictEqual(ws.getCell('AA2').value, 'late');
  assert.strictEqual(ws.getCell('AA4').value, 'wow');
  assert.strictEqual(ws.getCell('B6').fill.fgColor.argb, 'FF000000');        // non-submitter is black
  assert.strictEqual(ws.getCell('AA6').value, 'no submission');
  assert.strictEqual(ws.getCell('AC4').formula, '10/20');                      // legend derivation like the original
  assert.strictEqual(ws.getCell('AD4').formula, '5/20');
  assert.strictEqual(ws.getCell('AC5').formula, 'AC4/2');
  assert.strictEqual(ws.getCell('AD6').formula, 'AD4/4');
  assert.strictEqual(ws.getCell('AC7').value, 0);

  // Canvas import CSV keeps identity cols + chosen assignment, fills totals, preserves points row.
  const out = G.buildCanvasCsv(roster, 6, { 1001: 14.875, 1002: 13.5, 1003: 0 });
  assert.strictEqual(out.split('\r\n')[0], 'Student,ID,SIS User ID,SIS Login ID,Section,Lab 2 (901)');
  assert.strictEqual(out.split('\r\n')[1], '    Points Possible,,,,,15');
  assert.strictEqual(out.split('\r\n')[2], '"Doe, Jane",1001,s1,jdoe,PSY 320,14.875');
  assert.strictEqual(G.parseCsv(out).length, 5);
  console.log('all checks passed');
})().catch(e => { console.error(e); process.exit(1); });
