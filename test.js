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
const cfg = { cPts: 0.5, aPts: 0.25 };
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

  // Canvas import CSV keeps identity cols + chosen assignment, fills totals, preserves points row.
  const out = G.buildCanvasCsv(roster, 6, { 1001: 14.875, 1002: 13.5, 1003: 0 });
  assert.strictEqual(out.split('\r\n')[0], 'Student,ID,SIS User ID,SIS Login ID,Section,Lab 2 (901)');
  assert.strictEqual(out.split('\r\n')[1], '    Points Possible,,,,,15');
  assert.strictEqual(out.split('\r\n')[2], '"Doe, Jane",1001,s1,jdoe,PSY 320,14.875');
  assert.strictEqual(G.parseCsv(out).length, 5);
  console.log('all checks passed');
})().catch(e => { console.error(e); process.exit(1); });
