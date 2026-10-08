import type { Lesson } from '../types';

export interface LessonCorrection {
  type?: string;
  attendedCount?: number;
  teacherOverride?: string;
  hoursOverride?: number | null;
  baseHoursOverride?: number | null;
  reset?: boolean;
}

export function applyLessonCorrection(lesson: Lesson, fields: LessonCorrection): Lesson {
  const updated = { ...lesson };
  if (fields.reset) {
    delete updated.typeOverride;
    delete updated.attendedCountOverride;
    delete updated.teacherOverride;
    delete updated.hoursOverride;
    delete updated.baseHoursOverride;
    return updated;
  }
  if (fields.type !== undefined) {
    if (fields.type && !['中教', '外教'].includes(fields.type)) throw new Error('请选择有效课型');
    updated.typeOverride = fields.type || undefined;
  }
  if (fields.attendedCount !== undefined) {
    if (!Number.isInteger(fields.attendedCount) || fields.attendedCount < 0) throw new Error('参课人数必须为非负整数');
    updated.attendedCountOverride = fields.attendedCount;
  }
  if (fields.teacherOverride !== undefined) updated.teacherOverride = fields.teacherOverride.trim() || undefined;
  for (const key of ['hoursOverride', 'baseHoursOverride'] as const) {
    const value = fields[key];
    if (value === undefined) continue;
    if (value !== null && (!Number.isFinite(value) || value < 0)) throw new Error('课时必须为有效的非负数');
    updated[key] = value === null ? undefined : value;
  }
  return updated;
}
