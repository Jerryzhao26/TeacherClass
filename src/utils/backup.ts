import type { AppStateBackup, ClassBlock } from '../types';
import { isValidDateString } from './lessonRules';

type RecordValue = Record<string, any>;
const object = (v: unknown): v is RecordValue => !!v && typeof v === 'object' && !Array.isArray(v);
const text = (v: unknown): v is string => typeof v === 'string';
const name = (v: unknown) => text(v) && !!v.trim();
const nonnegative = (v: unknown) => typeof v === 'number' && Number.isFinite(v) && v >= 0;
const optional = (v: unknown, check: (v: any) => boolean) => v === undefined || check(v);
const date = (v: unknown) => text(v) && isValidDateString(v);
const array = (v: unknown, check: (v: RecordValue) => boolean) => Array.isArray(v) && v.every(x => object(x) && check(x));

export function migrateClassBlocks(value: unknown): ClassBlock[] {
  if (!array(value, b => name(b.id) && name(b.className) && text(b.classCode) && text(b.teacher) && text(b.schedule) &&
    ['once', 'twice'].includes(b.frequency) && optional(b.frequencySource, x => ['manual', 'schedule', 'history'].includes(x)) &&
    array(b.students, s => text(s.chineseName) && text(s.englishName) && Number.isInteger(s.colIndex) && s.colIndex >= 0) &&
    array(b.lessons, l => Number.isInteger(l.index) && date(l.dateStr) && text(l.rawDate) && text(l.type) &&
      object(l.studentStatus) && Object.values(l.studentStatus).every(text) && Number.isInteger(l.attendedCount) && l.attendedCount >= 0 &&
      Number.isInteger(l.totalStudents) && l.totalStudents >= 0 && optional(l.id, name) && optional(l.teacherOverride, text) &&
      optional(l.typeOverride, x => ['中教', '外教', ''].includes(x)) &&
      optional(l.attendedCountOverride, x => Number.isInteger(x) && x >= 0) &&
      optional(l.hoursOverride, nonnegative) && optional(l.baseHoursOverride, nonnegative)))) {
    throw new Error('课表数据格式无效，请重新导入 Excel');
  }
  const ids = new Set<string>();
  return (value as ClassBlock[]).map((b, bi) => {
    let id = b.id;
    while (ids.has(id)) id = `${id}:legacy:${bi}`;
    ids.add(id);
    const lessonIds = new Set<string>();
    return { ...b, id,
      // Legacy records may already contain manual frequency changes; preserve them.
      frequencySource: b.frequencySource ?? 'manual',
      lessons: b.lessons.map((l, li) => {
        let lessonId = l.id || `${id}:lesson:${li}`;
        while (lessonIds.has(lessonId)) lessonId = `${lessonId}:${li}`;
        lessonIds.add(lessonId);
        return { ...l, id: lessonId };
      })
    };
  });
}

export function validateBackup(value: unknown): Partial<AppStateBackup> {
  if (!object(value)) throw new Error('备份必须为 JSON 对象');
  if (!['bonusRules', 'teacherBaseRates', 'substitutionRecords', 'makeupRecords', 'classBlocks'].some(k => k in value)) {
    throw new Error('备份未包含支持的数据');
  }
  if (!optional(value.version, v => v === 1 || v === 2)) throw new Error('不支持该备份版本');
  const checks: Record<string, (v: RecordValue) => boolean> = {
    bonusRules: r => name(r.id) && name(r.teacherName) && name(r.classCode) && nonnegative(r.bonusRate) && optional(r.startDate, date) && optional(r.notes, text),
    teacherBaseRates: r => name(r.teacherName) && nonnegative(r.baseRate) && optional(r.teacherType, v => ['中教', '外教'].includes(v)) && optional(r.commissionRate, v => nonnegative(v) && v <= 1),
    substitutionRecords: r => name(r.id) && date(r.dateStr) && name(r.classCode) && text(r.className) && name(r.originalTeacher) && name(r.substituteTeacher) && r.originalTeacher.trim().toLowerCase() !== r.substituteTeacher.trim().toLowerCase() && optional(r.blockId, name) && optional(r.lessonId, name) && optional(r.notes, text),
    makeupRecords: r => name(r.id) && date(r.dateStr) && name(r.teacherName) && nonnegative(r.hours) && r.hours > 0 && optional(r.notes, text)
  };
  const result: Partial<AppStateBackup> = {};
  for (const [key, check] of Object.entries(checks)) {
    if (!(key in value)) continue;
    if (!array(value[key], check)) throw new Error(`备份中的 ${key} 格式无效`);
    (result as RecordValue)[key] = value[key];
  }
  if ('classBlocks' in value) result.classBlocks = migrateClassBlocks(value.classBlocks);
  if ('fileName' in value) {
    if (!text(value.fileName)) throw new Error('文件名格式无效');
    result.fileName = value.fileName;
  }
  if ('commissionRate' in value) {
    if (!nonnegative(value.commissionRate) || value.commissionRate > 1) throw new Error('提成比例无效');
    result.commissionRate = value.commissionRate;
  }
  return result;
}

// Roll back earlier keys if a quota/security error interrupts a multi-key save.
export function writeStorageBatch(storage: Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>, entries: Record<string, string>): void {
  const previous = Object.fromEntries(Object.keys(entries).map(k => [k, storage.getItem(k)]));
  const written: string[] = [];
  try {
    for (const [key, value] of Object.entries(entries)) {
      storage.setItem(key, value);
      written.push(key);
    }
  } catch (error) {
    for (const key of written.reverse()) {
      try {
        if (previous[key] === null) storage.removeItem(key);
        else storage.setItem(key, previous[key]);
      } catch { /* Preserve the original save error. */ }
    }
    throw error;
  }
}
