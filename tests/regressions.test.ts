import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import { parseExcelWorkbook } from '../src/utils/parser';
import { calculateLessonBase, detectFrequency, detectFrequencyFromLessons, normalizeDateString } from '../src/utils/lessonRules';
import { calculatePayroll, collectTeacherNames, resolveLessons } from '../src/utils/payroll';
import { applyLessonCorrection } from '../src/utils/corrections';
import { migrateClassBlocks, validateBackup, writeStorageBatch } from '../src/utils/backup';
import type { ClassBlock, Lesson, TeacherBaseRate, SubstitutionRecord, BonusRule, MakeupRecord } from '../src/types';

const lesson = (overrides: Partial<Lesson> = {}): Lesson => ({
  id: 'lesson-1', index: 1, type: '中教', dateStr: '2026-10-01', rawDate: '2026-10-01',
  studentStatus: { 张三: '是' }, attendedCount: 5, totalStudents: 5, ...overrides
});
const block = (overrides: Partial<ClassBlock> = {}): ClassBlock => ({
  id: 'block-1', className: 'E1.260905', classCode: '260905', teacher: 'Ann',
  schedule: '周六 09:00-11:00', frequency: 'once', frequencySource: 'schedule', students: [], lessons: [lesson()], ...overrides
});
function payroll({ blocks = [block()], rates = [], substitutions = [], bonuses = [], makeups = [], startDate = '', endDate = '', commissionRate = .07 }: {
  blocks?: ClassBlock[]; rates?: TeacherBaseRate[]; substitutions?: SubstitutionRecord[];
  bonuses?: BonusRule[]; makeups?: MakeupRecord[]; startDate?: string; endDate?: string; commissionRate?: number;
} = {}) {
  const input = { classBlocks: blocks, teacherBaseRates: rates, substitutionRecords: substitutions, makeupRecords: makeups };
  return calculatePayroll({ resolvedLessons: resolveLessons({ ...input, startDate, endDate }), teacherBaseRates: rates,
    uniqueTeachers: collectTeacherNames(input), bonusRules: bonuses, makeupRecords: makeups, startDate, endDate, commissionRate });
}
const sub: SubstitutionRecord = { id: 'sub', blockId: 'block-1', lessonId: 'lesson-1', dateStr: '2026-10-01',
  classCode: '260905', className: 'E1.260905', originalTeacher: 'Ann', substituteTeacher: 'Doris' };

test('manual teacher override transfers consumption and pay, including a new teacher', () => {
  const rows = payroll({ blocks: [block({ lessons: [lesson({ teacherOverride: 'NewTeacher' })] })] });
  const original = rows.find(r => r.teacherName === 'Ann')!;
  const actual = rows.find(r => r.teacherName === 'NewTeacher')!;
  assert.equal(original.settlementHours, 0);
  assert.equal(original.sessionsCount, 0);
  assert.equal(actual.settlementHours, 15);
  assert.equal(actual.sessionsCount, 1);
  assert.ok(Math.abs(actual.totalSalary - 105) < 1e-10);
  assert.equal(rows.reduce((sum, r) => sum + r.settlementHours, 0), 15);
});

test('substitution credits the actual teacher and applies only their bonus', () => {
  const rows = payroll({ substitutions: [sub], bonuses: [
    { id: 'a', teacherName: 'Ann', classCode: '260905', bonusRate: 20 },
    { id: 'd', teacherName: 'Doris', classCode: '260905', bonusRate: 5 }
  ] });
  assert.equal(rows.find(r => r.teacherName === 'Ann')!.totalSalary, 0);
  assert.equal(rows.find(r => r.teacherName === 'Doris')!.bonusAmount, 15);
  assert.ok(Math.abs(rows.find(r => r.teacherName === 'Doris')!.totalSalary - 120) < 1e-10);
});

test('manual override wins over substitution without counting hours twice', () => {
  const rows = payroll({ substitutions: [sub], blocks: [block({ lessons: [lesson({ teacherOverride: 'Emma' })] })] });
  assert.equal(rows.find(r => r.teacherName === 'Doris')!.settlementHours, 0);
  assert.equal(rows.find(r => r.teacherName === 'Emma')!.settlementHours, 15);
  assert.equal(rows.reduce((sum, r) => sum + r.settlementHours, 0), 15);
});

test('teacher capitalization and spacing do not split payroll entries', () => {
  const rows = payroll({ blocks: [block({ teacher: ' Ann ', lessons: [lesson({ teacherOverride: 'ANN' })] })],
    rates: [{ teacherName: 'ann', baseRate: 80, commissionRate: .06 }] });
  assert.equal(rows.length, 1);
  assert.equal(rows[0].substitutedOutHours, 0);
  assert.equal(rows[0].settlementHours, 15);
  assert.equal(rows[0].totalSalary, 72);
});

test('explicit Chinese teacher setting beats the name heuristic', () => {
  const input = { classBlocks: [block({ teacher: 'Alex', lessons: [lesson({ type: '' })] })],
    teacherBaseRates: [{ teacherName: 'Alex', baseRate: 100, teacherType: '中教' as const }], substitutionRecords: [], startDate: '', endDate: '' };
  assert.equal(resolveLessons(input)[0].baseHours, 3);
});

test('lesson type remains authoritative when a substitute has another teacher type', () => {
  const resolved = resolveLessons({ classBlocks: [block()], substitutionRecords: [sub],
    teacherBaseRates: [{ teacherName: 'Doris', baseRate: 100, teacherType: '外教' }], startDate: '', endDate: '' });
  assert.equal(resolved[0].resolvedTeacherType, '中教');
  assert.equal(resolved[0].baseHours, 3);
});

test('lesson-level substitution does not alter another session or another block', () => {
  const blocks = [block({ lessons: [lesson(), lesson({ id: 'lesson-2', index: 1 })] }), block({ id: 'block-2' })];
  const rows = payroll({ blocks, substitutions: [sub] });
  assert.equal(rows.find(r => r.teacherName === 'Ann')!.settlementHours, 30);
  assert.equal(rows.find(r => r.teacherName === 'Doris')!.settlementHours, 15);
});

test('legacy substitution matches original teacher as well as code and date', () => {
  const { blockId, lessonId, ...legacy } = sub;
  const rows = payroll({ blocks: [block(), block({ id: 'b2', teacher: 'Emma' })], substitutions: [legacy] });
  assert.equal(rows.find(r => r.teacherName === 'Emma')!.settlementHours, 15);
  assert.equal(rows.find(r => r.teacherName === 'Doris')!.settlementHours, 15);
});

test('makeups, personal commission and date/bonus boundaries use the filtered period', () => {
  const rows = payroll({ startDate: '2026-10-01', endDate: '2026-10-01', rates: [{ teacherName: 'Ann', baseRate: 80, commissionRate: .06 }],
    blocks: [block({ lessons: [lesson(), lesson({ id: 'other', dateStr: '2026-10-02' })] })],
    bonuses: [{ id: 'b', teacherName: 'Ann', classCode: '260905', bonusRate: 5, startDate: '2026-10-02' }],
    makeups: [ { id: 'm1', dateStr: '2026-10-01', teacherName: 'Ann', hours: 2 },
      { id: 'm2', dateStr: '2026-09-30', teacherName: 'Ann', hours: 100 } ] });
  assert.equal(rows[0].settlementHours, 17);
  assert.equal(rows[0].bonusAmount, 0);
  assert.equal(rows[0].commissionRate, .06);
  assert.equal(rows[0].totalSalary, 81.6);
});

test('bonus applies on its start date and uses pure class hours', () => {
  const rows = payroll({ bonuses: [{ id: 'b', teacherName: 'Ann', classCode: '260905', bonusRate: 5, startDate: '2026-10-01' }] });
  assert.equal(rows[0].bonusHours, 3);
  assert.equal(rows[0].bonusAmount, 15);
});

test('all frequency and lesson-type combinations have their expected base hours', () => {
  assert.deepEqual([calculateLessonBase('once', '中教'), calculateLessonBase('once', '外教'), calculateLessonBase('twice', '中教'), calculateLessonBase('twice', '外教')], [3, 2, 2, 1]);
});

test('reset restores original values, including attendance and lesson type', () => {
  const original = lesson({ type: '外教' });
  const modified = applyLessonCorrection(original, { type: '中教', attendedCount: 8, teacherOverride: 'Doris', baseHoursOverride: 4, hoursOverride: 25 });
  assert.equal(original.type, '外教');
  assert.equal(original.attendedCount, 5);
  assert.deepEqual(applyLessonCorrection(modified, { reset: true }), original);
});

test('invalid correction values are rejected before updating data', () => {
  for (const value of [NaN, Infinity, -1]) assert.throws(() => applyLessonCorrection(lesson(), { hoursOverride: value }));
  for (const value of [NaN, Infinity, -1, 1.5]) assert.throws(() => applyLessonCorrection(lesson(), { attendedCount: value }));
  assert.equal(applyLessonCorrection(lesson(), { hoursOverride: 0 }).hoursOverride, 0);
});

function workbook(schedule = '周六 09:00-11:00', dates: (string | number)[] = ['2026-10-01', '2026-10-02']) {
  const rows: any[][] = [[`班级：E1.260905`, '老师：Ann', `上课时间：${schedule}`],
    ['姓名', null, null, '张三'], ['英文名', null, null, 'Tom'], ['上课次数', '中/外', '日期'],
    ...dates.map(date => [1, '中', date, '是'])];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'A');
  return { wb, rows };
}

test('metadata retains full time and a once-weekly schedule despite extra lessons', () => {
  const parsed = parseExcelWorkbook(workbook().wb);
  assert.equal(parsed[0].schedule, '周六 09:00-11:00');
  assert.equal(parsed[0].frequency, 'once');
});

test('explicit twice-weekly schedule survives sparse history', () => {
  const parsed = parseExcelWorkbook(workbook('周三/周六 09:00-11:00', ['2026-10-01']).wb);
  assert.equal(parsed[0].frequency, 'twice');
  assert.equal(detectFrequencyFromLessons([lesson()], 'twice'), 'twice');
});

test('time punctuation alone does not imply multiple weekly sessions', () => {
  assert.equal(detectFrequency('周六\n09:00/11:00'), 'once');
  assert.equal(detectFrequency('星期三、星期六'), 'twice');
});

test('same metadata across worksheets and repeated lesson indices get distinct identities', () => {
  const { wb, rows } = workbook();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'B');
  const parsed = parseExcelWorkbook(wb);
  assert.notEqual(parsed[0].id, parsed[1].id);
  assert.notEqual(parsed[0].lessons[0].id, parsed[0].lessons[1].id);
  const corrected = parsed.map(b => b.id === parsed[0].id ? { ...b, lessons: b.lessons.map(l => l.id === b.lessons[0].id ? applyLessonCorrection(l, { hoursOverride: 99 }) : l) } : b);
  assert.equal(corrected[0].lessons[0].hoursOverride, 99);
  assert.equal(corrected[0].lessons[1].hoursOverride, undefined);
  assert.equal(corrected[1].lessons[0].hoursOverride, undefined);
});

test('dates are normalized and invalid calendar dates fail import', () => {
  assert.equal(normalizeDateString('26/10/1'), '2026-10-01');
  assert.equal(normalizeDateString('2026-13-40'), '');
  assert.equal(normalizeDateString('2026-02-29'), '');
  assert.throws(() => parseExcelWorkbook(workbook('周六', ['2026-13-40']).wb), /日期无效/);
});

test('Excel 1904 date system is honored', () => {
  const { wb } = workbook('周六', [44834]);
  const standard = parseExcelWorkbook(wb)[0].lessons[0].dateStr;
  wb.Workbook = { WBProps: { date1904: true } };
  const mac = parseExcelWorkbook(wb)[0].lessons[0].dateStr;
  assert.equal((Date.parse(mac) - Date.parse(standard)) / 86400000, 1462);
});

test('legacy class data receives unique lesson IDs and preserves manually set frequency', () => {
  const legacy = block({ frequencySource: undefined, lessons: [lesson({ id: undefined }), lesson({ id: undefined })] });
  const migrated = migrateClassBlocks([legacy, legacy]);
  assert.notEqual(migrated[0].id, migrated[1].id);
  assert.notEqual(migrated[0].lessons[0].id, migrated[0].lessons[1].id);
  assert.equal(migrated[0].frequencySource, 'manual');
});

test('legacy configuration backups remain valid and absent sections stay absent', () => {
  const backup = validateBackup({ bonusRules: [], teacherBaseRates: [], substitutionRecords: [] });
  assert.equal(backup.makeupRecords, undefined);
  assert.equal(backup.classBlocks, undefined);
});

test('full backup roundtrip retains corrections, rates and imported records', () => {
  const corrected = block({ lessons: [applyLessonCorrection(lesson(), { teacherOverride: 'Doris', attendedCount: 4 })] });
  const restored = validateBackup(JSON.parse(JSON.stringify({ version: 2, classBlocks: [corrected], fileName: 'October.xlsx', commissionRate: .06,
    bonusRules: [], teacherBaseRates: [], substitutionRecords: [], makeupRecords: [] })));
  assert.equal(restored.fileName, 'October.xlsx');
  assert.equal(restored.commissionRate, .06);
  assert.equal(restored.classBlocks![0].lessons[0].attendedCountOverride, 4);
  assert.equal(restored.classBlocks![0].lessons[0].teacherOverride, 'Doris');
});

test('malformed backups cannot introduce invalid types or numeric values', () => {
  for (const bad of [null, {}, { teacherBaseRates: {} }, { teacherBaseRates: [{ teacherName: 'Ann', baseRate: '100' }] },
    { makeupRecords: [{ id: 'x', dateStr: '2026-10-01', teacherName: 'Ann', hours: -1 }] },
    { bonusRules: [], commissionRate: Infinity }, { version: 3, bonusRules: [] },
    { classBlocks: [block({ lessons: [lesson({ hoursOverride: NaN })] })] }]) assert.throws(() => validateBackup(bad));
});

test('failed multi-section save restores earlier keys', () => {
  const values = new Map([['first', 'original']]);
  const storage = { getItem: (k: string) => values.get(k) ?? null, removeItem: (k: string) => { values.delete(k); },
    setItem: (k: string, v: string) => { if (k === 'last') throw new Error('quota exceeded'); values.set(k, v); } };
  assert.throws(() => writeStorageBatch(storage, { first: 'new', second: 'new', last: 'new' }), /quota/);
  assert.equal(values.get('first'), 'original');
  assert.equal(values.has('second'), false);
});

test('row insertion on reimport preserves block and lesson identity', () => {
  const { wb, rows } = workbook();
  const before = parseExcelWorkbook(wb)[0];
  const inserted = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(inserted, XLSX.utils.aoa_to_sheet([['标题'], [], ...rows]), 'A');
  const after = parseExcelWorkbook(inserted)[0];
  assert.equal(before.id, after.id);
  assert.equal(before.lessons[0].id, after.lessons[0].id);
});

test('attendance and final hours overrides agree with payroll and reset', () => {
  const edited = applyLessonCorrection(lesson(), { attendedCount: 4, baseHoursOverride: 2 });
  assert.equal(payroll({ blocks: [block({ lessons: [edited] })] })[0].settlementHours, 8);
  const final = applyLessonCorrection(edited, { hoursOverride: 11 });
  assert.equal(payroll({ blocks: [block({ lessons: [final] })] })[0].settlementHours, 11);
  const reset = applyLessonCorrection(final, { reset: true });
  assert.equal(payroll({ blocks: [block({ lessons: [reset] })] })[0].settlementHours, 15);
});

test('payroll component renders custom commission rates and calculated pay', async () => {
  const { createElement } = await import('react');
  const { renderToStaticMarkup } = await import('react-dom/server');
  const { PayrollReport } = await import('../src/components/PayrollReport');
  const rows = payroll({ commissionRate: .125 });
  const markup = renderToStaticMarkup(createElement(PayrollReport, { teacherReportData: rows, commissionRate: .125,
    onCommissionRateChange: () => {}, onTeacherCommissionRateChange: () => {}, onExport: () => {} }));
  assert.match(markup, /Ann/);
  assert.match(markup, /12\.5%/);
  assert.match(markup, /187\.5/);
});
