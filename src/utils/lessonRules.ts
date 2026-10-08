import type { Lesson } from '../types';

export function formatExcelDate(serial: number): string {
  // Excel base date is 1899-12-30 due to 1900 leap year bug
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);

  const year = date_info.getUTCFullYear();
  const month = String(date_info.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date_info.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

export function normalizeDateString(str: string): string {
  if (!str) return '';
  // Remove non-numeric spacing, replace dots, slashes, or backslashes with dashes
  let normalized = str.trim().replace(/[\.\/\\]/g, '-');
  
  // Clean up any double spaces or odd characters
  normalized = normalized.replace(/\s+/g, '');

  const parts = normalized.split('-');
  if (parts.length === 3) {
    let y = parts[0];
    let m = parts[1];
    let d = parts[2];
    
    // If year is 2 digits (e.g. "26" -> "2026")
    if (y.length === 2) {
      y = "20" + y;
    }
    m = m.padStart(2, '0');
    d = d.padStart(2, '0');
    
    // Check if valid numbers
    if (!isNaN(Number(y)) && !isNaN(Number(m)) && !isNaN(Number(d))) {
      const result = `${y}-${m}-${d}`;
      return isValidDateString(result) ? result : '';
    }
  }
  return isValidDateString(normalized) ? normalized : '';
}

export function extractClassCode(className: string): string {
  if (!className) return '';
  // Match trailing digits. For example "E3.200905" -> "200905"
  const match = className.match(/\d+$/);
  if (match) {
    return match[0];
  }
  // Fallback: match any digit group in the string
  const matchGroup = className.match(/\d+/g);
  if (matchGroup && matchGroup.length > 0) {
    return matchGroup[matchGroup.length - 1]; // Return the last number sequence
  }
  return className;
}

export function detectFrequency(schedule: string): 'once' | 'twice' {
  if (!schedule) return 'once';
  
  const scheduleClean = schedule.replace(/\s+/g, '').toLowerCase();
  
  if (
    scheduleClean.includes('两次') || 
    scheduleClean.includes('2次') || 
    scheduleClean.includes('双次') || 
    scheduleClean.includes('一周两次') || 
    scheduleClean.includes('一周2次') || 
    scheduleClean.includes('1周2次') || 
    scheduleClean.includes('1周两次')
  ) {
    return 'twice';
  }
  if (
    scheduleClean.includes('一次') || 
    scheduleClean.includes('1次') || 
    scheduleClean.includes('单次') || 
    scheduleClean.includes('一周一次') || 
    scheduleClean.includes('一周1次') || 
    scheduleClean.includes('1周1次') || 
    scheduleClean.includes('1周一次')
  ) {
    return 'once';
  }

  // Count actual weekdays; punctuation and line breaks alone are not evidence of another session.
  const matches = scheduleClean.match(/(?:周|星期|礼拜)([一二三四五六日天])/g) || [];
  const days = new Set(matches.map(day => day.slice(-1).replace('天', '日')));
  return days.size >= 2 ? 'twice' : 'once';
}

export function detectFrequencyFromLessons(lessons: Lesson[], scheduleFrequency: 'once' | 'twice'): 'once' | 'twice' {
  if (scheduleFrequency === 'twice' || !lessons || lessons.length === 0) return scheduleFrequency;

  // Group by week using YYYY-MM-DD representing the Monday of that week
  const weekMap = new Map<string, number>();
  lessons.forEach(l => {
    if (!l.dateStr) return;
    const parts = l.dateStr.split('-');
    if (parts.length !== 3) return;
    const y = parseInt(parts[0]);
    const m = parseInt(parts[1]) - 1;
    const d = parseInt(parts[2]);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return;
    
    // Create Date safely
    const dateObj = new Date(y, m, d);
    const day = dateObj.getDay();
    const dayDiff = day === 0 ? -6 : 1 - day; // Align to Monday
    const monday = new Date(dateObj.getTime() + dayDiff * 24 * 60 * 60 * 1000);
    const mondayStr = `${monday.getFullYear()}-${String(monday.getMonth() + 1).padStart(2, '0')}-${String(monday.getDate()).padStart(2, '0')}`;
    weekMap.set(mondayStr, (weekMap.get(mondayStr) || 0) + 1);
  });

  if (weekMap.size === 0) return scheduleFrequency;

  // Count active weeks and total lessons
  const activeWeeks = weekMap.size;
  let totalLessons = 0;
  let weeksWithTwoOrMore = 0;
  weekMap.forEach((count) => {
    totalLessons += count;
    if (count >= 2) {
      weeksWithTwoOrMore++;
    }
  });

  const avgLessonsPerWeek = totalLessons / activeWeeks;

  // If the average number of lessons per active week is >= 1.4,
  // or at least 40% of active weeks have 2 or more lessons, then it is "twice".
  // Otherwise, it is "once".
  if (avgLessonsPerWeek >= 1.4 || (weeksWithTwoOrMore / activeWeeks) >= 0.4) {
    return 'twice';
  }

  return 'once';
}

export function isLikelyForeignTeacher(name: string): boolean {
  if (!name) return false;
  const cleanName = name.trim().toLowerCase();
  
  // Explicit foreign indicators
  if (cleanName.includes('外教') || cleanName.includes('ft') || cleanName.includes('foreign') || cleanName.includes('teacher') || cleanName.includes('外')) {
    return true;
  }
  
  // Explicit Chinese indicators
  if (cleanName.includes('中教') || cleanName.includes('ct') || cleanName.includes('chinese') || cleanName.includes('local') || cleanName.includes('中')) {
    return false;
  }

  const hasChinese = /[\u4e00-\u9fa5]/.test(name);
  if (hasChinese) {
    // If it has Chinese characters, it is highly likely to be a Chinese teacher unless explicitly marked as foreign above.
    // E.g., "Amy老师", "Sunny老师", "张老师"
    return false;
  }

  // Purely English names (e.g., "Amy", "Sunny", "Alex")
  // Common English names used by Chinese teachers (definitely Chinese/中教 by default)
  const chineseEnglishNames = [
    'amy', 'sunny', 'eva', 'coco', 'tony', 'leo', 'jerry', 'kevin', 'fiona', 
    'vivian', 'cherry', 'vicky', 'cindy', 'helen', 'iris', 'jason', 'eric', 
    'jack', 'tom', 'bob', 'sam', 'angel', 'apple', 'lily', 'lucy', 'mary', 
    'sherry', 'vivi', 'zoe', 'joy', 'daisy', 'grace', 'clover', 'elsa', 
    'ivy', 'may', 'june', 'april', 'alice', 'annie', 'betty', 'candy', 
    'ella', 'gina', 'judy', 'karen', 'lisa', 'mandy', 'nancy', 'penny', 
    'rose', 'sally', 'tina', 'wendy', 'abby', 'bella', 'doris', 'gloria', 
    'irene', 'jane', 'kate', 'linda', 'mimi', 'olivia', 'rita', 'sharon', 
    'tracy', 'yoyo', 'selina', 'lucia', 'sophia', 'flora', 'maggie', 'sandy', 
    'stella', 'anna', 'clara', 'judith', 'kathy', 'paula', 'shirley',
    'winni', 'winnie', 'yuki', 'zoey'
  ];
  
  if (chineseEnglishNames.includes(cleanName)) {
    return false;
  }

  // Common foreign teacher names (or full Western names with a space)
  const foreignNames = [
    'alex', 'steve', 'jonathan', 'david', 'michael', 'sarah', 'emily', 'chris',
    'james', 'robert', 'john', 'william', 'brian', 'mark', 'richard',
    'thomas', 'charles', 'joseph', 'matthew', 'daniel', 'paul', 'andrew',
    'joshua', 'kenneth', 'steven', 'george', 'edward', 'ronald', 'timothy',
    'ryan', 'jeffrey', 'gary', 'nicholas', 'stephen', 'larry', 'gregory'
  ];
  
  if (foreignNames.includes(cleanName) || cleanName.includes(' ')) {
    return true;
  }

  // In a Chinese educational setting, the default teacher is Chinese.
  // So if we cannot be 100% sure, default to "中教" (Chinese)!
  return false;
}

export function isForeignType(typeStr: string): boolean {
  if (!typeStr) return false;
  const t = typeStr.trim().toLowerCase();
  
  // If it's explicitly a foreign teacher type
  if (t.includes('外教') || t.includes('foreign') || t === 'f' || t === 'ft' || t === 'w' || t === 'wai' || t === 'waijiao' || t === 'foreign teacher' || t === '外') {
    return true;
  }
  
  // If it matches a likely foreign teacher name directly in the class type column
  if (isLikelyForeignTeacher(typeStr)) {
    return true;
  }
  
  return false;
}

export function isValidStudentName(name: string): boolean {
  if (!name) return false;
  const clean = name.trim();
  if (clean.length === 0) return false;
  
  // Skip if purely numeric
  if (!isNaN(Number(clean))) return false;
  
  // Skip if it is a date format or has symbols like '-' or '/' exclusively
  if (/^\d+[\-\/\.]\d+/.test(clean)) return false;
  
  // Blacklist of common non-student column words in class records
  const blacklist = [
    '正常', '事假', '病假', '合计', '不计', '沙龙', '补课', '转班', '请假', '旷课',
    '出勤', '课时', '备注', '学费', '体验', '试听', '新生', '常规', '退费', '次数',
    '总计', '总数', '人数', '班级', '老师', '日期', '中教', '外教', '中/外', '中外',
    '上课', '消课', '扣课', '结余', '剩余', '退课', '签到', '是否', '状态', '级别',
    '阶段', '学号', '电话', '联系', '家长', '微信', '缴费', '金额', '单价', '总价',
    '提成', '结算', '提报', '确认', '核对', '统计', '说明', '原因', '情况'
  ];
  
  const shouldSkip = blacklist.some(keyword => clean.includes(keyword));
  if (shouldSkip) return false;
  
  return true;
}

export function calculateLessonBase(frequency: 'once' | 'twice', type: string): number {
  const isForeign = type.includes('外');
  if (frequency === 'twice') {
    return isForeign ? 1 : 2;
  } else {
    return isForeign ? 2 : 3;
  }
}

export function isValidDateString(value: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(value + 'T00:00:00Z');
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}
