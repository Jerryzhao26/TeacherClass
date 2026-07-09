import * as XLSX from 'xlsx';
import { ClassBlock, Student, Lesson } from '../types';

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
      return `${y}-${m}-${d}`;
    }
  }
  return normalized;
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

  const days = ['周一', '周二', '周三', '周四', '周五', '周六', '周日', '星期'];
  let count = 0;
  
  for (const day of days) {
    if (scheduleClean.includes(day)) {
      const regex = new RegExp(day, 'g');
      const matches = scheduleClean.match(regex);
      count += matches ? matches.length : 0;
    }
  }

  // Check if there are newline or multiple lines or common delimiters indicating multiple sessions
  if (
    schedule.includes('\n') || 
    schedule.includes('\r') || 
    schedule.includes('/') || 
    schedule.includes('&') || 
    schedule.includes('、') ||
    schedule.includes(',') ||
    schedule.includes('，') ||
    schedule.includes('和')
  ) {
    return 'twice';
  }
  return count >= 2 ? 'twice' : 'once';
}

export function detectFrequencyFromLessons(lessons: Lesson[], scheduleFrequency: 'once' | 'twice'): 'once' | 'twice' {
  if (!lessons || lessons.length === 0) return scheduleFrequency;

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

export function parseExcelWorkbook(workbook: XLSX.WorkBook): ClassBlock[] {
  const classBlocks: ClassBlock[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet['!ref']) continue;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const maxRow = range.e.r;
    const maxCol = range.e.c;

    // Create 2D grid
    const grid: any[][] = [];
    for (let r = 0; r <= maxRow; r++) {
      const rowData: any[] = [];
      for (let c = 0; c <= maxCol; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellAddress];
        rowData.push(cell ? cell.v : null);
      }
      grid.push(rowData);
    }

    let r = 0;
    while (r < grid.length) {
      const row = grid[r];
      let foundClassCell = false;
      let classColIndex = -1;

      // Look for a cell containing "班级" in this row
      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (val && typeof val === 'string' && (val.includes('班级') || val.trim() === '班级')) {
          foundClassCell = true;
          classColIndex = c;
          break;
        }
      }

      if (foundClassCell) {
        // Parse class metadata
        let className = '';
        let teacher = '';
        let schedule = '';

        const classCellVal = String(row[classColIndex]);
        if (classCellVal.includes('：') || classCellVal.includes(':')) {
          className = classCellVal.split(/：|:/)[1]?.trim() || '';
        } else if (classColIndex + 1 < row.length && row[classColIndex + 1]) {
          className = String(row[classColIndex + 1]).trim();
        }

        // Look for teacher (老师)
        for (let c = 0; c < row.length; c++) {
          const val = row[c];
          if (val && typeof val === 'string' && (val.includes('老师') || val.trim() === '老师')) {
            if (val.includes('：') || val.includes(':')) {
              teacher = val.split(/：|:/)[1]?.trim() || '';
            } else if (c + 1 < row.length && row[c + 1]) {
              teacher = String(row[c + 1]).trim();
            }
            break;
          }
        }

        // Look for schedule (时间段 / 上课时间)
        for (let c = 0; c < row.length; c++) {
          const val = row[c];
          if (val && typeof val === 'string' && (val.includes('时间段') || val.includes('上课时间') || val.includes('时间'))) {
            if (val.includes('：') || val.includes(':')) {
              schedule = val.split(/：|:/)[1]?.trim() || '';
            } else if (c + 1 < row.length && row[c + 1]) {
              schedule = String(row[c + 1]).trim();
            }
            break;
          }
        }

        // Search for student Names Row "姓名" / "英文名" in the next 5 rows
        let studentNamesRowIndex = -1;
        let englishNamesRowIndex = -1;
        
        for (let offset = 1; offset <= 5; offset++) {
          const nextR = r + offset;
          if (nextR >= grid.length) break;
          const nextRow = grid[nextR];
          for (let c = 0; c < nextRow.length; c++) {
            const val = nextRow[c];
            if (val && typeof val === 'string') {
              const cleanVal = val.trim();
              if (cleanVal === '姓名' || cleanVal === '学生姓名' || cleanVal === '学生') {
                studentNamesRowIndex = nextR;
              }
              if (cleanVal === '英文名') {
                englishNamesRowIndex = nextR;
              }
            }
          }
        }

        const students: Student[] = [];
        if (studentNamesRowIndex !== -1) {
          const sRow = grid[studentNamesRowIndex];
          const eRow = englishNamesRowIndex !== -1 ? grid[englishNamesRowIndex] : null;

          const nameColIndex = sRow.findIndex(val => {
            if (!val || typeof val !== 'string') return false;
            const clean = val.trim();
            return clean === '姓名' || clean === '学生姓名' || clean === '学生';
          });
          if (nameColIndex !== -1) {
            for (let c = nameColIndex + 1; c < sRow.length; c++) {
              const nameVal = sRow[c];
              if (nameVal && typeof nameVal === 'string' && nameVal.trim() !== '') {
                const cleanName = nameVal.trim();
                if (!isValidStudentName(cleanName)) {
                  continue;
                }
                const englishName = eRow && eRow[c] ? String(eRow[c]).trim() : '';
                students.push({
                  englishName,
                  chineseName: cleanName,
                  colIndex: c
                });
              }
            }
          }
        }

        // Look for the table headers with "日期" and "中/外"
        let dateHeaderRowIndex = -1;
        for (let offset = 3; offset <= 15; offset++) {
          const nextR = r + offset;
          if (nextR >= grid.length) break;
          const nextRow = grid[nextR];
          const hasDate = nextRow.some(val => val && typeof val === 'string' && val.trim().includes('日期'));
          if (hasDate) {
            dateHeaderRowIndex = nextR;
            break;
          }
        }

        if (dateHeaderRowIndex !== -1) {
          const headerRow = grid[dateHeaderRowIndex];
          const dateColIndex = headerRow.findIndex(val => val && typeof val === 'string' && val.trim().includes('日期'));
          const typeColIndex = headerRow.findIndex(val => {
            if (!val || typeof val !== 'string') return false;
            const clean = val.trim().toLowerCase();
            return clean.includes('中/外') || 
                   clean.includes('中外') || 
                   clean.includes('外/中') ||
                   clean.includes('外中') ||
                   clean.includes('教类') ||
                   clean.includes('中/外教') || 
                   clean.includes('中外教') || 
                   clean.includes('外教/中教') || 
                   clean.includes('课型') || 
                   clean.includes('类型') || 
                   clean.includes('类别') || 
                   clean.includes('属性') ||
                   clean.includes('师资') ||
                   clean === 'type' ||
                   clean === 'class type' ||
                   clean === 'teacher type' ||
                   clean === 'category';
          });
          const indexColIndex = headerRow.findIndex(val => val && typeof val === 'string' && (val.trim().includes('上课次数') || val.trim().includes('次数')));

          const lessons: Lesson[] = [];
          let lr = dateHeaderRowIndex + 1;

          while (lr < grid.length) {
            const lessonRow = grid[lr];
            if (!lessonRow) break;

            // Stop if we hit a row indicating a new class block
            const hasClassWord = lessonRow.some(val => val && typeof val === 'string' && val.includes('班级'));
            if (hasClassWord) {
              break;
            }

            const rawDateVal = dateColIndex !== -1 ? lessonRow[dateColIndex] : null;
            if (rawDateVal === null || rawDateVal === undefined || String(rawDateVal).trim() === '') {
              // End of the lesson table or empty row
              break;
            }

            // Parse Date
            let dateStr = '';
            if (typeof rawDateVal === 'number') {
              dateStr = formatExcelDate(rawDateVal);
            } else if (rawDateVal instanceof Date) {
              dateStr = rawDateVal.toISOString().split('T')[0];
            } else {
              dateStr = normalizeDateString(String(rawDateVal));
            }

            // Check if date looks semi-valid (should contain numbers)
            if (!dateStr || !/\d+/.test(dateStr)) {
              break;
            }

            // 1. First look at the cell directly preceding the date cell (dateColIndex - 1)
            let typeStr = '';
            if (dateColIndex > 0) {
              const precVal = lessonRow[dateColIndex - 1] !== null && lessonRow[dateColIndex - 1] !== undefined 
                ? String(lessonRow[dateColIndex - 1]).trim() 
                : '';
              if (precVal.includes('外') || precVal.toLowerCase() === 'f' || precVal.toLowerCase() === 'ft' || precVal.toLowerCase() === 'w') {
                typeStr = '外教';
              } else if (precVal.includes('中') || precVal.toLowerCase() === 'c' || precVal.toLowerCase() === 'ct' || precVal.toLowerCase() === 'z') {
                typeStr = '中教';
              }
            }

            // 2. Fall back to B列 (column index 1) of the lessonRow
            if (!typeStr) {
              const colBVal = lessonRow[1] !== null && lessonRow[1] !== undefined ? String(lessonRow[1]).trim() : '';
              if (colBVal.includes('外') || colBVal.toLowerCase() === 'f' || colBVal.toLowerCase() === 'ft' || colBVal.toLowerCase() === 'w') {
                typeStr = '外教';
              } else if (colBVal.includes('中') || colBVal.toLowerCase() === 'c' || colBVal.toLowerCase() === 'ct' || colBVal.toLowerCase() === 'z') {
                typeStr = '中教';
              }
            }

            // 3. Fall back to other potential type columns if not found
            if (!typeStr) {
              const rawTypeVal = (typeColIndex !== -1 && typeColIndex !== 1) ? lessonRow[typeColIndex] : null;
              const tempTypeStr = rawTypeVal !== null && rawTypeVal !== undefined ? String(rawTypeVal).trim() : '';
              if (tempTypeStr) {
                if (tempTypeStr.includes('外') || tempTypeStr.toLowerCase() === 'f' || tempTypeStr.toLowerCase() === 'ft' || tempTypeStr.toLowerCase() === 'w') {
                  typeStr = '外教';
                } else if (tempTypeStr.includes('中') || tempTypeStr.toLowerCase() === 'c' || tempTypeStr.toLowerCase() === 'ct' || tempTypeStr.toLowerCase() === 'z') {
                  typeStr = '中教';
                }
              }
            }
            
            // 3. Keep it empty if not explicitly found, so that we can fall back to the actual teacher type dynamically in the application
            if (!typeStr) {
              typeStr = '';
            }

            const rawIndexVal = indexColIndex !== -1 ? lessonRow[indexColIndex] : (lessons.length + 1);
            const indexNum = typeof rawIndexVal === 'number' ? rawIndexVal : parseInt(String(rawIndexVal)) || (lessons.length + 1);

            // Attendance parsing
            const studentStatus: { [studentName: string]: string } = {};
            let attendedCount = 0;

            for (const student of students) {
              const statusVal = lessonRow[student.colIndex];
              const statusStr = statusVal ? String(statusVal).trim() : '-';
              studentStatus[student.chineseName] = statusStr;
              if (statusStr === '是') {
                attendedCount++;
              }
            }

            lessons.push({
              index: indexNum,
              type: typeStr,
              dateStr,
              rawDate: String(rawDateVal),
              studentStatus,
              attendedCount,
              totalStudents: students.length
            });

            lr++;
          }

          const classCode = extractClassCode(className);
          const scheduleFrequency = detectFrequency(schedule);
          const finalFrequency = detectFrequencyFromLessons(lessons, scheduleFrequency);

          if (className) {
            classBlocks.push({
              id: `${className}_${teacher}_${schedule || ''}`.replace(/\s+/g, '_'),
              className,
              classCode,
              teacher,
              schedule,
              frequency: finalFrequency,
              students,
              lessons
            });
          }

          // Advance row scan index r to the end of lessons block
          r = lr - 1;
        }

        r++;
      } else {
        r++;
      }
    }
  }

  return classBlocks;
}

export function calculateLessonBase(frequency: 'once' | 'twice', type: string): number {
  const isForeign = type.includes('外');
  if (frequency === 'twice') {
    return isForeign ? 1 : 2;
  } else {
    return isForeign ? 2 : 3;
  }
}
