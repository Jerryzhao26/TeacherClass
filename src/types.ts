export interface Student {
  englishName: string;
  chineseName: string;
  colIndex: number;
}

export interface Lesson {
  index: number; // 上课次数
  type: string; // "中" or "外"
  dateStr: string; // normalized date (YYYY-MM-DD)
  rawDate: string; // raw string from excel
  studentStatus: { [studentName: string]: string }; // e.g. { "张三": "是", "李四": "事假" }
  attendedCount: number; // number of students with "是"
  totalStudents: number; // total students in class
  teacherOverride?: string;
  hoursOverride?: number;
  baseHoursOverride?: number;
  typeOverride?: string; // "中教" or "外教" manually overridden
}

export interface ClassBlock {
  id: string; // unique identifier
  className: string; // original class name, e.g. "E3.200905"
  classCode: string; // extracted numeric part, e.g. "200905"
  teacher: string; // teacher name
  schedule: string; // schedule text
  frequency: 'once' | 'twice'; // once a week or twice a week
  students: Student[];
  lessons: Lesson[];
}

export interface BonusRule {
  id: string;
  teacherName: string;
  classCode: string;
  bonusRate: number; // extra pay per hour (元/课时)
  startDate?: string; // Optional start date for the bonus (YYYY-MM-DD)
  notes?: string;
}

export interface TeacherBaseRate {
  teacherName: string;
  baseRate: number; // base hourly rate (元/课时)
  teacherType?: '中教' | '外教';
  commissionRate?: number; // teacher's specific commission rate (e.g. 0.07 or 0.06)
}

export interface SubstitutionRecord {
  id: string;
  dateStr: string; // YYYY-MM-DD
  classCode: string; // class code
  className: string; // full class name (for display)
  originalTeacher: string; // original teacher
  substituteTeacher: string; // substitute teacher
  notes?: string;
}

export interface MakeupRecord {
  id: string;
  dateStr: string; // YYYY-MM-DD
  teacherName: string;
  hours: number; // 增加的课销数量
  notes?: string;
}

export interface AppStateBackup {
  bonusRules: BonusRule[];
  teacherBaseRates: TeacherBaseRate[];
  substitutionRecords: SubstitutionRecord[];
  makeupRecords?: MakeupRecord[];
}
