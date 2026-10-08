import * as XLSX from 'xlsx';
import type { ClassBlock, Student, Lesson } from '../types';
import { formatExcelDate, normalizeDateString, extractClassCode, detectFrequency, detectFrequencyFromLessons, isValidStudentName, isValidDateString } from './lessonRules';
export * from './lessonRules';

export function parseExcelWorkbook(workbook: XLSX.WorkBook): ClassBlock[] {
  const classBlocks: ClassBlock[] = [];
  const blockOccurrences = new Map<string, number>();

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
          className = classCellVal.slice(classCellVal.search(/：|:/) + 1).trim();
        } else if (classColIndex + 1 < row.length && row[classColIndex + 1]) {
          className = String(row[classColIndex + 1]).trim();
        }

        // Look for teacher (老师)
        for (let c = 0; c < row.length; c++) {
          const val = row[c];
          if (val && typeof val === 'string' && (val.includes('老师') || val.trim() === '老师')) {
            if (val.includes('：') || val.includes(':')) {
              teacher = val.slice(val.search(/：|:/) + 1).trim();
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
              schedule = val.slice(val.search(/：|:/) + 1).trim();
            } else if (c + 1 < row.length && row[c + 1]) {
              schedule = String(row[c + 1]).trim();
            }
            break;
          }
        }

        const blockKey = JSON.stringify([sheetName, className, teacher, schedule]);
        const occurrence = blockOccurrences.get(blockKey) || 0;
        blockOccurrences.set(blockKey, occurrence + 1);
        const blockId = `${blockKey}:${occurrence}`;
        const lessonOccurrences = new Map<string, number>();

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
              dateStr = formatExcelDate(rawDateVal + (workbook.Workbook?.WBProps?.date1904 ? 1462 : 0));
            } else if (rawDateVal instanceof Date) {
              dateStr = rawDateVal.toISOString().split('T')[0];
            } else {
              dateStr = normalizeDateString(String(rawDateVal));
            }

            // Check if date looks semi-valid (should contain numbers)
            if (!isValidDateString(dateStr)) {
              throw new Error(`工作表 ${sheetName} 第 ${lr + 1} 行日期无效：${rawDateVal}`);
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

            const lessonOccurrence = lessonOccurrences.get(dateStr) || 0;
            lessonOccurrences.set(dateStr, lessonOccurrence + 1);
            lessons.push({
              id: `${blockId}:date:${dateStr}:${lessonOccurrence}`,
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
          const finalFrequency = schedule.trim() ? scheduleFrequency : detectFrequencyFromLessons(lessons, scheduleFrequency);

          if (className) {
            classBlocks.push({
              id: blockId,
              frequencySource: schedule.trim() ? 'schedule' : 'history',
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
