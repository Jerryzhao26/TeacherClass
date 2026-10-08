import type { ClassBlock, TeacherBaseRate, SubstitutionRecord, MakeupRecord, BonusRule, Lesson } from '../types';
import { calculateLessonBase, isLikelyForeignTeacher } from './lessonRules';

export interface CalculationInput {
  classBlocks: ClassBlock[];
  substitutionRecords: SubstitutionRecord[];
  teacherBaseRates: TeacherBaseRate[];
  startDate: string;
  endDate: string;
}
export function resolveLessons({ classBlocks, substitutionRecords, teacherBaseRates, startDate, endDate }: CalculationInput) {

  const list: Array<{
    block: ClassBlock;
    lesson: Lesson;
    actualTeacher: string;
    isSubstituted: boolean;
    subRecord?: SubstitutionRecord;
    classCode: string;
    hours: number;
    baseHours: number;
    resolvedTeacherType: string;
  }> = [];

  classBlocks.forEach(block => {
    const classCode = block.classCode;
    
    block.lessons.forEach(lesson => {
      // Date filter
      if (startDate && lesson.dateStr < startDate) return;
      if (endDate && lesson.dateStr > endDate) return;

      // Check if there is a substitution record for this class code and date (by classCode and dateStr, case-insensitive)
      const sub = substitutionRecords.find(s => 
        (!s.blockId || s.blockId === block.id) &&
        s.originalTeacher.trim().toLowerCase() === block.teacher.trim().toLowerCase() &&
        (!s.lessonId || s.lessonId === lesson.id) &&
        s.dateStr === lesson.dateStr && 
        s.classCode.trim().toLowerCase() === classCode.trim().toLowerCase()
      );

      const actualTeacher = lesson.teacherOverride || (sub ? sub.substituteTeacher : block.teacher);
      const isSubstituted = actualTeacher.trim().toLowerCase() !== block.teacher.trim().toLowerCase();

      // Determine whether this lesson's type is Chinese or Foreign
      const teacherConfig = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === actualTeacher.trim().toLowerCase());
      const isTeacherForeign = teacherConfig?.teacherType ? teacherConfig.teacherType === '外教' : isLikelyForeignTeacher(actualTeacher);
      const resolvedTeacherType = lesson.typeOverride || 
                                   lesson.type || 
                                   (isTeacherForeign ? '外教' : '中教');

      const baseHours = lesson.baseHoursOverride !== undefined && lesson.baseHoursOverride !== null
        ? lesson.baseHoursOverride
        : calculateLessonBase(block.frequency, resolvedTeacherType);

      const hours = lesson.hoursOverride !== undefined && lesson.hoursOverride !== null
        ? lesson.hoursOverride
        : (lesson.attendedCountOverride ?? lesson.attendedCount) * baseHours;

      list.push({
        block,
        lesson,
        actualTeacher,
        isSubstituted,
        subRecord: sub,
        classCode,
        hours,
        baseHours,
        resolvedTeacherType
      });
    });
  });

  // Sort by date descending
  return list.sort((a, b) => b.lesson.dateStr.localeCompare(a.lesson.dateStr));
}

export type ResolvedLesson = ReturnType<typeof resolveLessons>[number];

export function calculatePayroll({ resolvedLessons, teacherBaseRates, uniqueTeachers, bonusRules, makeupRecords, startDate, endDate, commissionRate }: {
 resolvedLessons: ResolvedLesson[]; teacherBaseRates: TeacherBaseRate[]; uniqueTeachers: string[]; bonusRules: BonusRule[]; makeupRecords: MakeupRecord[]; startDate: string; endDate: string; commissionRate: number;
}) {

  const report: { [teacherName: string]: {
    teacherName: string;
    baseRate: number;
    commissionRate: number;
    sessionsCount: number;
    baseHours: number; // Taught normal class hours
    baseClassHours: number; // Taught normal class hours (without student count)
    substitutedOutHours: number; // Substituted by others (deducted)
    substitutedOutClassHours: number; // Substituted by others (deducted, without student count)
    substitutedInHours: number; // Substituting for others (added)
    substitutedInClassHours: number; // Substituting for others (added, without student count)
    makeupHours: number; // Makeup lesson hours added
    settlementHours: number; // Final credited hours: base - out + in + makeup
    settlementClassHours: number; // Final credited hours: base - out + in + makeup (without student count)
    bonusHours: number; // Hours eligible for bonus
    bonusAmount: number; // Total bonus money
    baseSalary: number; // settlementHours * baseRate (Total Course Deduction Value)
    commissionAmount: number; // baseSalary * commissionRate (Teacher's commission from course deduction)
    totalSalary: number; // commissionAmount + bonusAmount (Actual take-home lesson salary)
    substitutionDetails: string[];
    makeupDetails: string[];
  }} = Object.create(null);

  // Initialize report structure for all unique teachers
  uniqueTeachers.forEach(tName => {
    const baseRateObj = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === tName.trim().toLowerCase());
    report[tName] = {
      teacherName: tName,
      commissionRate: baseRateObj?.commissionRate ?? commissionRate,
      baseRate: baseRateObj ? baseRateObj.baseRate : 100, // Default 100
      sessionsCount: 0,
      baseHours: 0,
      baseClassHours: 0,
      substitutedOutHours: 0,
      substitutedOutClassHours: 0,
      substitutedInHours: 0,
      substitutedInClassHours: 0,
      makeupHours: 0,
      settlementHours: 0,
      settlementClassHours: 0,
      bonusHours: 0,
      bonusAmount: 0,
      baseSalary: 0,
      commissionAmount: 0,
      totalSalary: 0,
      substitutionDetails: [],
      makeupDetails: []
    };
  });

  const reportKeys = new Map(Object.keys(report).map(k => [k.trim().toLowerCase(), k]));

  // Helper to find report entry case-insensitively and with trimming
  const findReportEntry = (name: string) => {
    if (!name) return null;
    const cleanName = name.trim().toLowerCase();
    const matchingKey = reportKeys.get(cleanName);
    return matchingKey ? report[matchingKey] : null;
  };

  // Go through all resolved lessons in the filtered date range
  resolvedLessons.forEach(item => {
    const origTeacher = item.block.teacher;
    const actTeacher = item.actualTeacher;
    const hours = item.hours;
    const baseHours = item.baseHours;
    const classCode = item.classCode;
    const date = item.lesson.dateStr;

    const origReport = findReportEntry(origTeacher);
    const actReport = findReportEntry(actTeacher);

    // 1. Taught sessions & Base Hours tracking
    if (origReport) {
      origReport.baseHours += hours;
      origReport.baseClassHours += baseHours;
    }

    if (item.isSubstituted) {
      // Original teacher is substituted OUT
      if (origReport) {
        origReport.substitutedOutHours += hours;
        origReport.substitutedOutClassHours += baseHours;
        origReport.substitutionDetails.push(
          `[-] ${date} 由 [${actTeacher}] 代课 ${item.block.className} (${hours} 课时 / 纯课时:${baseHours})`
        );
      }
      // Substitute teacher is substituted IN
      if (actReport) {
        actReport.substitutedInHours += hours;
        actReport.substitutedInClassHours += baseHours;
        actReport.sessionsCount += 1;
        actReport.substitutionDetails.push(
          `[+] ${date} 代替 [${origTeacher}] 授课 ${item.block.className} (${hours} 课时 / 纯课时:${baseHours})`
        );
      }
    } else {
      // Taught by original teacher
      if (actReport) {
        actReport.sessionsCount += 1;
      }
    }

    // 2. Bonus calculation
    // Bonus goes to the actual teacher who taught, provided they have a bonus rule for this class code
    const rule = bonusRules.find(r => 
      r.teacherName.trim().toLowerCase() === actTeacher.trim().toLowerCase() && 
      r.classCode.trim().toLowerCase() === classCode.trim().toLowerCase()
    );

    if (rule && actReport) {
      // If the rule specifies a start date, only apply the bonus if the lesson date is >= the start date
      const isEligible = !rule.startDate || date >= rule.startDate;
      if (isEligible) {
        const bonusPay = baseHours * rule.bonusRate;
        actReport.bonusHours += baseHours;
        actReport.bonusAmount += bonusPay;
      }
    }
  });

  // 3. Sum up makeups for each teacher
  const activeMakeups = makeupRecords.filter(m => {
    if (startDate && m.dateStr < startDate) return false;
    if (endDate && m.dateStr > endDate) return false;
    return true;
  });

  activeMakeups.forEach(m => {
    const tName = m.teacherName;
    const mReport = findReportEntry(tName);
    if (mReport) {
      mReport.makeupHours += m.hours;
      mReport.sessionsCount += 1;
      mReport.makeupDetails.push(
        `${m.dateStr} 补课: +${m.hours} 课时${m.notes ? ` (${m.notes})` : ''}`
      );
    }
  });

  // Final mathematical summaries for each teacher
  Object.keys(report).forEach(tName => {
    const data = report[tName];
    const baseRateObj = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === tName.trim().toLowerCase());
    const rateOfCommission = (baseRateObj && baseRateObj.commissionRate !== undefined) ? baseRateObj.commissionRate : commissionRate;

    data.settlementHours = data.baseHours - data.substitutedOutHours + data.substitutedInHours + data.makeupHours;
    data.settlementClassHours = data.baseClassHours - data.substitutedOutClassHours + data.substitutedInClassHours + data.makeupHours;
    data.baseSalary = data.settlementHours * data.baseRate;
    data.commissionAmount = data.baseSalary * rateOfCommission;
    data.totalSalary = data.commissionAmount + data.bonusAmount;
  });

  return Object.values(report).sort((a, b) => b.totalSalary - a.totalSalary);
}

export type PayrollRow = ReturnType<typeof calculatePayroll>[number];

export function collectTeacherNames({ classBlocks, teacherBaseRates, substitutionRecords, makeupRecords }: { classBlocks: ClassBlock[]; teacherBaseRates: TeacherBaseRate[]; substitutionRecords: SubstitutionRecord[]; makeupRecords: MakeupRecord[] }) {

  const namesMap = new Map<string, string>(); // lowercase -> original trimmed representation
  
  classBlocks.forEach(b => {
    if (b.teacher) {
      const trimmed = b.teacher.trim();
      if (trimmed) {
        namesMap.set(trimmed.toLowerCase(), trimmed);
      }
    }
  });
  
  classBlocks.forEach(b => b.lessons.forEach(l => {
    const name = l.teacherOverride?.trim();
    if (name) namesMap.set(name.toLowerCase(), name);
  }));

  teacherBaseRates.forEach(r => {
    if (r.teacherName) {
      const trimmed = r.teacherName.trim();
      if (trimmed) {
        namesMap.set(trimmed.toLowerCase(), trimmed);
      }
    }
  });
  
  substitutionRecords.forEach(s => {
    if (s.originalTeacher) {
      const trimmed = s.originalTeacher.trim();
      if (trimmed) {
        namesMap.set(trimmed.toLowerCase(), trimmed);
      }
    }
    if (s.substituteTeacher) {
      const trimmed = s.substituteTeacher.trim();
      if (trimmed) {
        namesMap.set(trimmed.toLowerCase(), trimmed);
      }
    }
  });
  
  makeupRecords.forEach(m => {
    if (m.teacherName) {
      const trimmed = m.teacherName.trim();
      if (trimmed) {
        namesMap.set(trimmed.toLowerCase(), trimmed);
      }
    }
  });
  
  return Array.from(namesMap.values());
}
