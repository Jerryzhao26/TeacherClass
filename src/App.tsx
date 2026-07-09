import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  Calendar, 
  DollarSign, 
  Users, 
  CheckCircle2, 
  XCircle, 
  AlertCircle, 
  Settings, 
  Plus, 
  Trash2, 
  Download, 
  RefreshCw, 
  FileSpreadsheet, 
  UserCheck, 
  Search, 
  FileDown, 
  ChevronRight, 
  Info, 
  Layers, 
  BookOpen,
  Check,
  Edit,
  UserX,
  Sparkles,
  CheckSquare
} from 'lucide-react';
import { ClassBlock, Student, Lesson, BonusRule, TeacherBaseRate, SubstitutionRecord, MakeupRecord, AppStateBackup } from './types';
import { parseExcelWorkbook, calculateLessonBase, extractClassCode, detectFrequency, detectFrequencyFromLessons, isLikelyForeignTeacher } from './utils/parser';

export default function App() {
  // --- STATE ---
  const [classBlocks, setClassBlocks] = useState<ClassBlock[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [activeTab, setActiveTab] = useState<'statistics' | 'settings' | 'substitutions' | 'makeups' | 'report'>('statistics');

  // Filter dates
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');

  // Filters for course consumption details table
  const [filterTeacher, setFilterTeacher] = useState<string>('');
  const [filterClassCode, setFilterClassCode] = useState<string>('');

  // Persistent user configurations (loaded from LocalStorage)
  const [bonusRules, setBonusRules] = useState<BonusRule[]>([]);
  const [teacherBaseRates, setTeacherBaseRates] = useState<TeacherBaseRate[]>([]);
  const [substitutionRecords, setSubstitutionRecords] = useState<SubstitutionRecord[]>([]);
  const [makeupRecords, setMakeupRecords] = useState<MakeupRecord[]>([]);

  // Selection states for details viewing
  const [selectedBlockId, setSelectedBlockId] = useState<string>('');
  const [selectedLessonIndex, setSelectedLessonIndex] = useState<number | null>(null);

  // Manual correction form states
  const [editType, setEditType] = useState<string>('');
  const [editAttendedCount, setEditAttendedCount] = useState<number>(0);
  const [editTeacher, setEditTeacher] = useState<string>('');
  const [editBaseHours, setEditBaseHours] = useState<string>('');
  const [editHours, setEditHours] = useState<string>('');

  // Form states for adding rules
  const [newRuleTeacher, setNewRuleTeacher] = useState<string>('');
  const [newRuleClassCode, setNewRuleClassCode] = useState<string>('');
  const [newRuleBonusRate, setNewRuleBonusRate] = useState<string>('');
  const [newRuleNotes, setNewRuleNotes] = useState<string>('');

  // Form states for teacher base rates
  const [newTeacherName, setNewTeacherName] = useState<string>('');
  const [newTeacherBaseRate, setNewTeacherBaseRate] = useState<string>('');
  const [newTeacherType, setNewTeacherType] = useState<'中教' | '外教'>('中教');
  const [newTeacherCommissionRate, setNewTeacherCommissionRate] = useState<number>(0.07);

  // Form states for substitutions
  const [subDate, setSubDate] = useState<string>('');
  const [subClassCode, setSubClassCode] = useState<string>('');
  const [subOriginalTeacher, setSubOriginalTeacher] = useState<string>('');
  const [subSubstituteTeacher, setSubSubstituteTeacher] = useState<string>('');
  const [subNotes, setSubNotes] = useState<string>('');

  // Form states for makeups
  const [makeupDate, setMakeupDate] = useState<string>('');
  const [makeupTeacher, setMakeupTeacher] = useState<string>('');
  const [makeupHours, setMakeupHours] = useState<string>('');
  const [makeupNotes, setMakeupNotes] = useState<string>('');

  // Selected commission rate (default 7% or 6%)
  const [commissionRate, setCommissionRate] = useState<number>(0.07);

  // Feedback notifications
  const [notification, setNotification] = useState<{ message: string; type: 'success' | 'error' } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- LOCAL STORAGE PERSISTENCE ---
  useEffect(() => {
    const savedBonusRules = localStorage.getItem('course_deduction_bonus_rules');
    const savedBaseRates = localStorage.getItem('course_deduction_base_rates');
    const savedSubstitutions = localStorage.getItem('course_deduction_substitutions');
    const savedMakeups = localStorage.getItem('course_deduction_makeups');
    const savedClassBlocks = localStorage.getItem('course_deduction_class_blocks');
    const savedFileName = localStorage.getItem('course_deduction_file_name');
    const savedCommissionRate = localStorage.getItem('course_deduction_commission_rate');

    if (savedBonusRules) setBonusRules(JSON.parse(savedBonusRules));
    if (savedBaseRates) setTeacherBaseRates(JSON.parse(savedBaseRates));
    if (savedSubstitutions) setSubstitutionRecords(JSON.parse(savedSubstitutions));
    if (savedMakeups) setMakeupRecords(JSON.parse(savedMakeups));
    if (savedClassBlocks) setClassBlocks(JSON.parse(savedClassBlocks));
    if (savedFileName) setFileName(savedFileName);
    if (savedCommissionRate) setCommissionRate(parseFloat(savedCommissionRate));
  }, []);

  const saveToLocalStorage = (key: string, data: any) => {
    localStorage.setItem(key, JSON.stringify(data));
  };

  // Trigger temporary alerts
  const showAlert = (message: string, type: 'success' | 'error' = 'success') => {
    setNotification({ message, type });
    setTimeout(() => {
      setNotification(null);
    }, 4000);
  };

  // --- EXTRACT DATE RANGE FROM EXCEL DATA ---
  useEffect(() => {
    if (classBlocks.length > 0) {
      let earliest = '';
      let latest = '';
      classBlocks.forEach(block => {
        block.lessons.forEach(lesson => {
          if (!earliest || lesson.dateStr < earliest) earliest = lesson.dateStr;
          if (!latest || lesson.dateStr > latest) latest = lesson.dateStr;
        });
      });
      if (earliest && !startDate) setStartDate(earliest);
      if (latest && !endDate) setEndDate(latest);
    }
  }, [classBlocks]);

  // --- FILE HANDLING ---
  const handleFile = (file: File) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const parsedBlocks = parseExcelWorkbook(workbook);

        if (parsedBlocks.length === 0) {
          showAlert('没有在Excel中解析到符合格式的班级课时块，请检查文件。', 'error');
          return;
        }

        setClassBlocks(parsedBlocks);
        setFileName(file.name);
        saveToLocalStorage('course_deduction_class_blocks', parsedBlocks);
        saveToLocalStorage('course_deduction_file_name', file.name);

        // Auto initialize teacher base rates if they don't exist
        const uniqueTeachers = Array.from(new Set(parsedBlocks.map(b => b.teacher).filter(Boolean)));
        const updatedBaseRates = [...teacherBaseRates];
        let addedCount = 0;
        uniqueTeachers.forEach(t => {
          if (!updatedBaseRates.some(r => r.teacherName === t)) {
            const detectedType = isLikelyForeignTeacher(t) ? '外教' : '中教';
            updatedBaseRates.push({ teacherName: t, baseRate: 100, teacherType: detectedType }); // Default rate 100 with auto detected type
            addedCount++;
          }
        });
        if (addedCount > 0) {
          setTeacherBaseRates(updatedBaseRates);
          saveToLocalStorage('course_deduction_base_rates', updatedBaseRates);
        }

        showAlert(`成功导入 ${parsedBlocks.length} 个班级，共 ${parsedBlocks.reduce((acc, b) => acc + b.lessons.length, 0)} 节上课记录！`);
      } catch (err: any) {
        console.error(err);
        showAlert(`解析Excel失败：${err.message || '未知错误'}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const onDragLeave = () => {
    setIsDragging(false);
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFile(files[0]);
    }
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  const clearImportedData = () => {
    if (window.confirm('确定要清除所有导入的课表数据吗？您的加成和代课配置仍会保留。')) {
      setClassBlocks([]);
      setFileName('');
      localStorage.removeItem('course_deduction_class_blocks');
      localStorage.removeItem('course_deduction_file_name');
      setSelectedBlockId('');
      setSelectedLessonIndex(null);
      showAlert('已清空导入的课表数据');
    }
  };

  // --- DYNAMIC DATA PROCESSING & CALCULATIONS ---

  // Update frequency overrides or details on Class Blocks
  const updateClassFrequency = (blockId: string, frequency: 'once' | 'twice') => {
    const updated = classBlocks.map(b => {
      if (b.id === blockId) {
        return { ...b, frequency };
      }
      return b;
    });
    setClassBlocks(updated);
    saveToLocalStorage('course_deduction_class_blocks', updated);
    showAlert('班级上课频次已手动更新');
  };

  // Automatically detect and update the frequency for all class blocks based on their lesson history
  const autoDetectAllFrequencies = () => {
    if (classBlocks.length === 0) {
      showAlert('暂无导入的排课数据！', 'error');
      return;
    }
    const updated = classBlocks.map(block => {
      const scheduleFrequency = detectFrequency(block.schedule);
      const finalFrequency = detectFrequencyFromLessons(block.lessons, scheduleFrequency);
      return { ...block, frequency: finalFrequency };
    });
    setClassBlocks(updated);
    saveToLocalStorage('course_deduction_class_blocks', updated);
    showAlert('已根据上课记录日期自动判定所有班级的上课频次！');
  };

  // Resolve actual teacher and lesson details taking substitution into account
  const resolvedLessons = useMemo(() => {
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

        // Check if there is a substitution record for this class code and date
        const sub = substitutionRecords.find(s => 
          s.dateStr === lesson.dateStr && 
          s.classCode === classCode && 
          s.originalTeacher.trim().toLowerCase() === block.teacher.trim().toLowerCase()
        );

        const actualTeacher = lesson.teacherOverride || (sub ? sub.substituteTeacher : block.teacher);
        const isSubstituted = !!sub && !lesson.teacherOverride;

        // Determine whether this lesson's type is Chinese or Foreign
        const teacherConfig = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === actualTeacher.trim().toLowerCase());
        const isTeacherForeign = teacherConfig?.teacherType === '外教' || isLikelyForeignTeacher(actualTeacher);
        const resolvedTeacherType = lesson.typeOverride || 
                                     lesson.type || 
                                     (isTeacherForeign ? '外教' : '中教');

        const baseHours = lesson.baseHoursOverride !== undefined && lesson.baseHoursOverride !== null
          ? lesson.baseHoursOverride
          : calculateLessonBase(block.frequency, resolvedTeacherType);

        const hours = lesson.hoursOverride !== undefined && lesson.hoursOverride !== null
          ? lesson.hoursOverride
          : lesson.attendedCount * baseHours;

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
  }, [classBlocks, substitutionRecords, startDate, endDate, teacherBaseRates]);

  // Unique list of class codes currently imported
  const uniqueClassCodes = useMemo(() => {
    const codes = new Set<string>();
    classBlocks.forEach(b => {
      if (b.classCode) codes.add(b.classCode);
    });
    return Array.from(codes);
  }, [classBlocks]);

  // Unique list of teacher names currently imported or configured
  const uniqueTeachers = useMemo(() => {
    const names = new Set<string>();
    classBlocks.forEach(b => {
      if (b.teacher) names.add(b.teacher);
    });
    teacherBaseRates.forEach(r => names.add(r.teacherName));
    substitutionRecords.forEach(s => {
      names.add(s.originalTeacher);
      names.add(s.substituteTeacher);
    });
    return Array.from(names).filter(Boolean);
  }, [classBlocks, teacherBaseRates, substitutionRecords]);

  // Unique list of classes (code and name) currently imported
  const uniqueClassesList = useMemo(() => {
    const list: Array<{ classCode: string; className: string }> = [];
    const codes = new Set<string>();
    classBlocks.forEach(b => {
      if (b.classCode && !codes.has(b.classCode)) {
        codes.add(b.classCode);
        list.push({ classCode: b.classCode, className: b.className });
      }
    });
    return list.sort((a, b) => a.className.localeCompare(b.className));
  }, [classBlocks]);

  // Filtered list of resolved lessons for display in the table
  const filteredResolvedLessons = useMemo(() => {
    return resolvedLessons.filter(item => {
      if (filterTeacher && item.actualTeacher.trim().toLowerCase() !== filterTeacher.trim().toLowerCase()) {
        return false;
      }
      if (filterClassCode && item.classCode !== filterClassCode) {
        return false;
      }
      return true;
    });
  }, [resolvedLessons, filterTeacher, filterClassCode]);

  // Selected class block for viewing detail
  const selectedBlock = useMemo(() => {
    return classBlocks.find(b => b.id === selectedBlockId);
  }, [classBlocks, selectedBlockId]);

  // Synchronize manual correction form states when selected lesson changes
  useEffect(() => {
    if (selectedBlock && selectedLessonIndex !== null) {
      const matchedLesson = selectedBlock.lessons.find(l => l.index === selectedLessonIndex);
      if (matchedLesson) {
        setEditType(matchedLesson.typeOverride || matchedLesson.type);
        setEditAttendedCount(matchedLesson.attendedCount);
        setEditTeacher(matchedLesson.teacherOverride || '');
        setEditBaseHours(matchedLesson.baseHoursOverride !== undefined && matchedLesson.baseHoursOverride !== null ? String(matchedLesson.baseHoursOverride) : '');
        setEditHours(matchedLesson.hoursOverride !== undefined && matchedLesson.hoursOverride !== null ? String(matchedLesson.hoursOverride) : '');
      }
    }
  }, [selectedBlockId, selectedLessonIndex, selectedBlock]);

  // Update a single lesson's manual overrides
  const handleUpdateLessonOverride = (
    blockId: string, 
    lessonIndex: number, 
    fields: {
      type?: string;
      attendedCount?: number;
      teacherOverride?: string;
      hoursOverride?: number | null;
      baseHoursOverride?: number | null;
    }
  ) => {
    const updatedBlocks = classBlocks.map(b => {
      if (b.id !== blockId) return b;
      
      const updatedLessons = b.lessons.map(l => {
        if (l.index !== lessonIndex) return l;
        
        const newLesson = { ...l };
        if (fields.type !== undefined) {
          newLesson.type = fields.type;
          newLesson.typeOverride = fields.type;
        }
        if (fields.attendedCount !== undefined) newLesson.attendedCount = fields.attendedCount;
        
        if (fields.teacherOverride !== undefined) {
          newLesson.teacherOverride = fields.teacherOverride || undefined;
        }
        if (fields.hoursOverride !== undefined) {
          newLesson.hoursOverride = fields.hoursOverride === null ? undefined : fields.hoursOverride;
        }
        if (fields.baseHoursOverride !== undefined) {
          newLesson.baseHoursOverride = fields.baseHoursOverride === null ? undefined : fields.baseHoursOverride;
        }
        
        return newLesson;
      });
      
      return { ...b, lessons: updatedLessons };
    });
    
    setClassBlocks(updatedBlocks);
    saveToLocalStorage('course_deduction_class_blocks', updatedBlocks);
    showAlert('该节课的排课记录已成功更新！');
  };

  // --- CONFIG MANAGEMENTS ---

  // Add Bonus Rule
  const handleAddBonusRule = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newRuleTeacher || !newRuleClassCode || !newRuleBonusRate) {
      showAlert('请填完整加成规则信息！', 'error');
      return;
    }

    const rate = parseFloat(newRuleBonusRate);
    if (isNaN(rate) || rate < 0) {
      showAlert('加成单价必须为正数！', 'error');
      return;
    }

    // Check if duplicate exists (teacher + class code)
    const exists = bonusRules.some(r => 
      r.teacherName.trim().toLowerCase() === newRuleTeacher.trim().toLowerCase() && 
      r.classCode.trim() === newRuleClassCode.trim()
    );

    if (exists) {
      showAlert(`已经存在该老师（${newRuleTeacher}）和班级编号（${newRuleClassCode}）的加成配置，请先删除旧的再添加。`, 'error');
      return;
    }

    const newRule: BonusRule = {
      id: `rule_${Date.now()}`,
      teacherName: newRuleTeacher.trim(),
      classCode: newRuleClassCode.trim(),
      bonusRate: rate,
      notes: newRuleNotes.trim()
    };

    const updated = [...bonusRules, newRule];
    setBonusRules(updated);
    saveToLocalStorage('course_deduction_bonus_rules', updated);

    // Clear form
    setNewRuleBonusRate('');
    setNewRuleNotes('');
    showAlert('成功添加课时费加成规则！');
  };

  const handleDeleteBonusRule = (id: string) => {
    const updated = bonusRules.filter(r => r.id !== id);
    setBonusRules(updated);
    saveToLocalStorage('course_deduction_bonus_rules', updated);
    showAlert('加成规则已删除');
  };

  // Update/Add Teacher Base Rate
  const handleAddTeacherRate = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTeacherName || !newTeacherBaseRate) {
      showAlert('请填完整教师单价信息！', 'error');
      return;
    }

    const rate = parseFloat(newTeacherBaseRate);
    if (isNaN(rate) || rate < 0) {
      showAlert('基础单价必须为正数！', 'error');
      return;
    }

    const updated = [...teacherBaseRates];
    const index = updated.findIndex(r => r.teacherName.trim().toLowerCase() === newTeacherName.trim().toLowerCase());

    if (index !== -1) {
      updated[index].baseRate = rate;
      updated[index].teacherType = newTeacherType;
      updated[index].commissionRate = newTeacherCommissionRate;
    } else {
      updated.push({ 
        teacherName: newTeacherName.trim(), 
        baseRate: rate, 
        teacherType: newTeacherType,
        commissionRate: newTeacherCommissionRate
      });
    }

    setTeacherBaseRates(updated);
    saveToLocalStorage('course_deduction_base_rates', updated);

    setNewTeacherName('');
    setNewTeacherBaseRate('');
    setNewTeacherType('colleagues' as any === 'colleagues' ? '中教' : '外教');
    setNewTeacherCommissionRate(0.07);
    showAlert('成功更新教师基础课时与提成配置！');
  };

  const toggleTeacherType = (teacherName: string) => {
    const updated = teacherBaseRates.map(r => {
      if (r.teacherName.trim().toLowerCase() === teacherName.trim().toLowerCase()) {
        const nextType = r.teacherType === '外教' ? '中教' : '外教';
        return { ...r, teacherType: nextType };
      }
      return r;
    });
    setTeacherBaseRates(updated);
    saveToLocalStorage('course_deduction_base_rates', updated);
    showAlert(`已成功将教师 ${teacherName} 的属性切换为 [${updated.find(r => r.teacherName.trim().toLowerCase() === teacherName.trim().toLowerCase())?.teacherType || '中教'}]！`);
  };

  const updateTeacherCommissionRate = (teacherName: string, rate: number) => {
    const updated = teacherBaseRates.map(r => {
      if (r.teacherName.trim().toLowerCase() === teacherName.trim().toLowerCase()) {
        return { ...r, commissionRate: rate };
      }
      return r;
    });
    setTeacherBaseRates(updated);
    saveToLocalStorage('course_deduction_base_rates', updated);
    showAlert(`已成功将教师 ${teacherName} 的专属提成比例修改为 ${(rate * 100).toFixed(0)}%！`);
  };

  const handleDeleteTeacherRate = (teacherName: string) => {
    const updated = teacherBaseRates.filter(r => r.teacherName !== teacherName);
    setTeacherBaseRates(updated);
    saveToLocalStorage('course_deduction_base_rates', updated);
    showAlert('教师课时单价配置已清除');
  };

  // Add Substitution Record
  const handleAddSubstitution = (e: React.FormEvent) => {
    e.preventDefault();
    if (!subDate || !subClassCode || !subOriginalTeacher || !subSubstituteTeacher) {
      showAlert('请填写完整的代课记录信息！', 'error');
      return;
    }

    if (subOriginalTeacher.trim().toLowerCase() === subSubstituteTeacher.trim().toLowerCase()) {
      showAlert('代课老师和原任课老师不能是同一个人！', 'error');
      return;
    }

    // Find full name of class if available for display
    const matchedClass = classBlocks.find(b => b.classCode === subClassCode);
    const className = matchedClass ? matchedClass.className : `班级(${subClassCode})`;

    const record: SubstitutionRecord = {
      id: `sub_${Date.now()}`,
      dateStr: subDate,
      classCode: subClassCode,
      className,
      originalTeacher: subOriginalTeacher.trim(),
      substituteTeacher: subSubstituteTeacher.trim(),
      notes: subNotes.trim()
    };

    const updated = [...substitutionRecords, record];
    setSubstitutionRecords(updated);
    saveToLocalStorage('course_deduction_substitutions', updated);

    // Clear form
    setSubDate('');
    setSubClassCode('');
    setSubOriginalTeacher('');
    setSubSubstituteTeacher('');
    setSubNotes('');
    showAlert('成功录入代课记录，相关教师的薪资已被自动重算！');
  };

  const handleDeleteSubstitution = (id: string) => {
    const updated = substitutionRecords.filter(s => s.id !== id);
    setSubstitutionRecords(updated);
    saveToLocalStorage('course_deduction_substitutions', updated);
    showAlert('代课记录已删除，薪资自动恢复');
  };

  // Add Makeup Record
  const handleAddMakeup = (e: React.FormEvent) => {
    e.preventDefault();
    if (!makeupDate || !makeupTeacher || !makeupHours) {
      showAlert('请填写完整的补课记录信息！', 'error');
      return;
    }

    const hoursNum = parseFloat(makeupHours);
    if (isNaN(hoursNum) || hoursNum <= 0) {
      showAlert('请输入有效的补课课时数量！', 'error');
      return;
    }

    const record: MakeupRecord = {
      id: `makeup_${Date.now()}`,
      dateStr: makeupDate,
      teacherName: makeupTeacher.trim(),
      hours: hoursNum,
      notes: makeupNotes.trim()
    };

    const updated = [...makeupRecords, record];
    setMakeupRecords(updated);
    saveToLocalStorage('course_deduction_makeups', updated);

    // Clear form
    setMakeupDate('');
    setMakeupTeacher('');
    setMakeupHours('');
    setMakeupNotes('');
    showAlert('成功录入补课记录，相关教师的课消已被自动增加并重算！');
  };

  const handleDeleteMakeup = (id: string) => {
    const updated = makeupRecords.filter(m => m.id !== id);
    setMakeupRecords(updated);
    saveToLocalStorage('course_deduction_makeups', updated);
    showAlert('补课记录已删除，对应课销已扣除');
  };

  // --- SUGGESTED BONUS CONFIGS FROM IMPORTED DATA ---
  const suggestedRules = useMemo(() => {
    const suggestions: Array<{ teacherName: string; classCode: string; className: string }> = [];
    classBlocks.forEach(b => {
      if (b.teacher && b.classCode) {
        // Check if a bonus rule already exists
        const exists = bonusRules.some(r => 
          r.teacherName.trim().toLowerCase() === b.teacher.trim().toLowerCase() && 
          r.classCode === b.classCode
        );
        const alreadyInSuggestions = suggestions.some(s => 
          s.teacherName === b.teacher && 
          s.classCode === b.classCode
        );
        if (!exists && !alreadyInSuggestions) {
          suggestions.push({
            teacherName: b.teacher,
            classCode: b.classCode,
            className: b.className
          });
        }
      }
    });
    return suggestions;
  }, [classBlocks, bonusRules]);

  const addSuggestedRule = (s: { teacherName: string; classCode: string }) => {
    setNewRuleTeacher(s.teacherName);
    setNewRuleClassCode(s.classCode);
    setNewRuleBonusRate('5'); // Default suggestions of 5
    // scroll to form
    document.getElementById('bonus-config-form')?.scrollIntoView({ behavior: 'smooth' });
  };

  // --- BACKUP & RESTORE ---
  const handleBackupExport = () => {
    const backup: AppStateBackup = {
      bonusRules,
      teacherBaseRates,
      substitutionRecords,
      makeupRecords
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `课销结算配置备份_${new Date().toISOString().split('T')[0]}.json`;
    link.click();
    URL.revokeObjectURL(url);
    showAlert('配置备份文件导出成功！');
  };

  const handleBackupImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const backup: AppStateBackup = JSON.parse(event.target?.result as string);
        if (!backup.bonusRules && !backup.teacherBaseRates && !backup.substitutionRecords && !backup.makeupRecords) {
          showAlert('无效的备份文件，未检测到支持的配置类型。', 'error');
          return;
        }

        if (backup.bonusRules) {
          setBonusRules(backup.bonusRules);
          saveToLocalStorage('course_deduction_bonus_rules', backup.bonusRules);
        }
        if (backup.teacherBaseRates) {
          setTeacherBaseRates(backup.teacherBaseRates);
          saveToLocalStorage('course_deduction_base_rates', backup.teacherBaseRates);
        }
        if (backup.substitutionRecords) {
          setSubstitutionRecords(backup.substitutionRecords);
          saveToLocalStorage('course_deduction_substitutions', backup.substitutionRecords);
        }
        if (backup.makeupRecords) {
          setMakeupRecords(backup.makeupRecords);
          saveToLocalStorage('course_deduction_makeups', backup.makeupRecords);
        } else {
          setMakeupRecords([]);
          saveToLocalStorage('course_deduction_makeups', []);
        }

        showAlert('配置备份成功导入！所有课时费加成、教师单价及代课/补课记录已恢复。');
        // Reset file input
        e.target.value = '';
      } catch (err) {
        showAlert('解析备份文件失败，请确认文件格式正确。', 'error');
      }
    };
    reader.readAsText(file);
  };

  // --- REPORT GENERATION ---
  const teacherReportData = useMemo(() => {
    const report: { [teacherName: string]: {
      teacherName: string;
      baseRate: number;
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
    }} = {};

    // Initialize report structure for all unique teachers
    uniqueTeachers.forEach(tName => {
      const baseRateObj = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === tName.trim().toLowerCase());
      report[tName] = {
        teacherName: tName,
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

    // Go through all resolved lessons in the filtered date range
    resolvedLessons.forEach(item => {
      const origTeacher = item.block.teacher;
      const actTeacher = item.actualTeacher;
      const hours = item.hours;
      const baseHours = item.baseHours;
      const classCode = item.classCode;
      const date = item.lesson.dateStr;

      // 1. Taught sessions & Base Hours tracking
      if (report[origTeacher]) {
        report[origTeacher].baseHours += hours;
        report[origTeacher].baseClassHours += baseHours;
      }

      if (item.isSubstituted) {
        // Original teacher is substituted OUT
        if (report[origTeacher]) {
          report[origTeacher].substitutedOutHours += hours;
          report[origTeacher].substitutedOutClassHours += baseHours;
          report[origTeacher].substitutionDetails.push(
            `[-] ${date} 由 [${actTeacher}] 代课 ${item.block.className} (${hours} 课时 / 纯课时:${baseHours})`
          );
        }
        // Substitute teacher is substituted IN
        if (report[actTeacher]) {
          report[actTeacher].substitutedInHours += hours;
          report[actTeacher].substitutedInClassHours += baseHours;
          report[actTeacher].sessionsCount += 1;
          report[actTeacher].substitutionDetails.push(
            `[+] ${date} 代替 [${origTeacher}] 授课 ${item.block.className} (${hours} 课时 / 纯课时:${baseHours})`
          );
        }
      } else {
        // Taught by original teacher
        if (report[actTeacher]) {
          report[actTeacher].sessionsCount += 1;
        }
      }

      // 2. Bonus calculation
      // Bonus goes to the actual teacher who taught, provided they have a bonus rule for this class code
      const rule = bonusRules.find(r => 
        r.teacherName.trim().toLowerCase() === actTeacher.trim().toLowerCase() && 
        r.classCode === classCode
      );

      if (rule && report[actTeacher]) {
        const bonusPay = baseHours * rule.bonusRate;
        report[actTeacher].bonusHours += baseHours;
        report[actTeacher].bonusAmount += bonusPay;
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
      if (report[tName]) {
        report[tName].makeupHours += m.hours;
        report[tName].sessionsCount += 1;
        report[tName].makeupDetails.push(
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
  }, [resolvedLessons, teacherBaseRates, bonusRules, makeupRecords, uniqueTeachers, startDate, endDate, commissionRate]);

  // Overall KPI summaries
  const overviewStats = useMemo(() => {
    let totalClasses = classBlocks.length;
    let totalLessonsCount = resolvedLessons.length;
    
    const activeMakeups = makeupRecords.filter(m => {
      if (startDate && m.dateStr < startDate) return false;
      if (endDate && m.dateStr > endDate) return false;
      return true;
    });
    const totalMakeupHours = activeMakeups.reduce((acc, m) => acc + m.hours, 0);

    let totalHours = resolvedLessons.reduce((acc, item) => acc + item.hours, 0) + totalMakeupHours;
    let totalBaseHours = resolvedLessons.reduce((acc, item) => acc + item.baseHours, 0) + totalMakeupHours;
    let totalCourseDeduction = teacherReportData.reduce((acc, t) => acc + t.baseSalary, 0);
    let totalCommission = teacherReportData.reduce((acc, t) => acc + t.commissionAmount, 0);
    let totalBonuses = teacherReportData.reduce((acc, t) => acc + t.bonusAmount, 0);
    let totalSalary = teacherReportData.reduce((acc, t) => acc + t.totalSalary, 0);

    return {
      totalClasses,
      totalLessonsCount: totalLessonsCount + activeMakeups.length,
      totalHours,
      totalBaseHours,
      totalCourseDeduction,
      totalCommission,
      totalBonuses,
      totalSalary
    };
  }, [classBlocks, resolvedLessons, teacherReportData, makeupRecords, startDate, endDate]);

  // Export payroll report to CSV
  const handleExportCSV = () => {
    if (teacherReportData.length === 0) {
      showAlert('没有数据可以导出，请先导入Excel文件。', 'error');
      return;
    }

    const headers = [
      '老师姓名',
      '基础单价(元/课时)',
      '总授课次数',
      '名下班级课时(乘人数)',
      '名下班级纯课时(不计人数)',
      '代出课时(-, 乘人数)',
      '代出纯课时(-, 不计人数)',
      '代入课时(+, 乘人数)',
      '代入纯课时(+, 不计人数)',
      '补课增加课时(+, 不计人数)',
      '总结算课时(乘人数)',
      '总结算纯课时(不计人数)',
      '加成课时',
      '加成提成累计(元)',
      '总课销金额(元)',
      '提成比例',
      '课销提成金额(元)',
      '老师实际到手薪资(元)'
    ];

    const rows = teacherReportData.map(t => {
      const baseRateObj = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === t.teacherName.trim().toLowerCase());
      const rateOfCommission = (baseRateObj && baseRateObj.commissionRate !== undefined) ? baseRateObj.commissionRate : commissionRate;
      return [
        t.teacherName,
        t.baseRate,
        t.sessionsCount,
        t.baseHours.toFixed(1),
        t.baseClassHours.toFixed(1),
        t.substitutedOutHours.toFixed(1),
        t.substitutedOutClassHours.toFixed(1),
        t.substitutedInHours.toFixed(1),
        t.substitutedInClassHours.toFixed(1),
        t.makeupHours.toFixed(1),
        t.settlementHours.toFixed(1),
        t.settlementClassHours.toFixed(1),
        t.bonusHours.toFixed(1),
        t.bonusAmount.toFixed(1),
        t.baseSalary.toFixed(1),
        `${(rateOfCommission * 100).toFixed(0)}%`,
        t.commissionAmount.toFixed(1),
        t.totalSalary.toFixed(1)
      ];
    });

    const dateRangeStr = (startDate || endDate) ? ` (${startDate || ''}至${endDate || ''})` : '';
    exportToCSV(`培训机构月度课消薪资结算报表${dateRangeStr}.csv`, headers, rows);
    showAlert('薪资报表导出成功！');
  };

  // Helper for BOM CSV exporting
  const exportToCSV = (filename: string, headers: string[], rows: any[][]) => {
    const BOM = '\uFEFF';
    const csvContent = BOM + [
      headers.join(','),
      ...rows.map(row => row.map(val => {
        const cell = val === null || val === undefined ? '' : String(val);
        if (cell.includes(',') || cell.includes('"') || cell.includes('\n')) {
          return `"${cell.replace(/"/g, '""')}"`;
        }
        return cell;
      }).join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans antialiased">
      {/* HEADER SECTION */}
      <header className="bg-white border-b border-slate-200/80 sticky top-0 z-40 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 py-4 sm:px-6 lg:px-8 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600 text-white p-2.5 rounded-2xl shadow-md shadow-indigo-100">
              <Sparkles className="w-6 h-6" />
            </div>
            <div>
              <h1 id="app-title" className="text-2xl font-bold text-slate-900 tracking-tight">
                比伯斯课销结算系统 <span className="text-xs font-normal bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded border border-indigo-100/80 ml-2">v2.5.0</span>
              </h1>
              <p className="text-xs text-slate-500 mt-0.5">教务报表与教师课时费自动结算中心 • 智能识别、极少代课/课消意外手动修正</p>
            </div>
          </div>
          
          {/* QUICK BACKUP CONTROLS */}
          <div className="flex items-center gap-2.5 self-start md:self-auto">
            <button
              onClick={handleBackupExport}
              title="备份课时加成与教师基础课时单价，导出为JSON文件"
              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-50 shadow-sm transition cursor-pointer"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              导出配置
            </button>
            <label className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-50 shadow-sm transition cursor-pointer">
              <Upload className="w-3.5 h-3.5 text-slate-500" />
              导入配置
              <input 
                type="file" 
                accept=".json" 
                onChange={handleBackupImport} 
                className="hidden" 
              />
            </label>
          </div>
        </div>
      </header>

      {/* GLOBAL NOTIFICATION TOAST */}
      {notification && (
        <div className="fixed bottom-5 right-5 z-50 flex items-center gap-3 px-4 py-3.5 rounded-2xl shadow-lg border border-slate-100 animate-in fade-in slide-in-from-bottom-5 duration-300 max-w-md bg-white">
          {notification.type === 'success' ? (
            <CheckCircle2 className="w-5 h-5 text-emerald-500 shrink-0" />
          ) : (
            <AlertCircle className="w-5 h-5 text-rose-500 shrink-0" />
          )}
          <span className="text-sm font-semibold text-slate-700">{notification.message}</span>
        </div>
      )}

      <main className="max-w-7xl mx-auto px-4 py-6 sm:px-6 lg:px-8 space-y-6">
        
        {/* IMPORT AREA & DATE FILTERING */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* FILE EXCEL UPLOADER CARD */}
          <div className="lg:col-span-7 bg-white rounded-2xl border border-slate-200/80 p-6 shadow-sm space-y-4 flex flex-col justify-between">
            <h2 className="text-sm font-bold text-slate-900 flex items-center gap-2">
              <FileSpreadsheet className="w-4 h-4 text-indigo-600" />
              1. 导入排课课时记录 Excel 表
            </h2>
            
            <div 
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
              onDrop={onDrop}
              onClick={triggerFileSelect}
              className={`border-2 border-dashed rounded-2xl p-6 flex flex-col items-center justify-center text-center cursor-pointer transition-all duration-200 ${
                isDragging 
                  ? 'border-indigo-500 bg-indigo-50/50 scale-[0.99]' 
                  : fileName 
                    ? 'border-slate-200 bg-slate-50/40 hover:border-indigo-400' 
                    : 'border-slate-200 bg-slate-50/20 hover:border-indigo-500/50 hover:bg-slate-50/60'
              }`}
            >
              <input 
                type="file" 
                ref={fileInputRef} 
                onChange={(e) => e.target.files && handleFile(e.target.files[0])} 
                accept=".xlsx, .xls" 
                className="hidden" 
              />
              
              <div className={`p-3 rounded-full mb-3 ${fileName ? 'bg-indigo-50 text-indigo-600' : 'bg-slate-100 text-slate-400'}`}>
                <Upload className="w-6 h-6" />
              </div>
              
              {fileName ? (
                <div className="space-y-1">
                  <p className="text-sm font-bold text-slate-800 break-all px-4">{fileName}</p>
                  <p className="text-xs text-indigo-600 font-semibold">文件解析成功！点击或拖拽可重新上传更换</p>
                </div>
              ) : (
                <div className="space-y-1">
                  <p className="text-sm font-bold text-slate-700">点击或将 Excel 文件拖拽到此处上传</p>
                  <p className="text-xs text-slate-400 max-w-sm">支持包含分段、多班级多区域的常规课消统计表，自动匹配分析</p>
                </div>
              )}
            </div>

            {/* SHOW UPLOADED CLASSES SUMMARY */}
            {classBlocks.length > 0 && (
              <div className="bg-slate-50 border border-slate-100 rounded-xl p-4 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
                  <p className="text-xs text-slate-600">
                    已成功加载 <strong className="text-slate-900 font-bold">{classBlocks.length}</strong> 个班级信息，共检测到 
                    <strong className="text-slate-900 font-bold ml-1">{classBlocks.reduce((acc, b) => acc + b.lessons.length, 0)}</strong> 节上课记录。
                  </p>
                </div>
                <button
                  onClick={clearImportedData}
                  className="text-xs text-rose-600 hover:text-rose-700 font-semibold hover:underline flex items-center gap-1 cursor-pointer"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  清空数据
                </button>
              </div>
            )}
          </div>

          {/* DATE FILTERING CARD */}
          <div className="lg:col-span-5 bg-white rounded-2xl border border-slate-200/80 p-6 shadow-sm space-y-4 flex flex-col justify-between">
            <div className="space-y-3.5">
              <h2 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                <Calendar className="w-4 h-4 text-indigo-600" />
                2. 设定结算统计的时间范围
              </h2>
              <p className="text-xs text-slate-400 leading-relaxed">
                选择需要统计的课消区间。系统将自动过滤超出范围的数据，并依据此时段内的排课和代课调整生成最终报表。
              </p>
              
              <div className="grid grid-cols-2 gap-3 pt-1">
                <div>
                  <label className="block text-xs font-semibold text-slate-500 mb-1">开始日期</label>
                  <div className="relative">
                    <input 
                      type="date" 
                      value={startDate}
                      onChange={(e) => setStartDate(e.target.value)}
                      className="w-full text-sm border border-slate-200 rounded-xl px-3 py-2 focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 outline-none" 
                    />
                  </div>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-500 mb-1">结束日期</label>
                  <div className="relative">
                    <input 
                      type="date" 
                      value={endDate}
                      onChange={(e) => setEndDate(e.target.value)}
                      className="w-full text-sm border border-slate-200 rounded-xl px-3 py-2 focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 outline-none" 
                    />
                  </div>
                </div>
              </div>
            </div>

            {/* QUICK FILTER ACTIONS */}
            <div className="flex gap-2 pt-3 border-t border-slate-100">
              <button
                onClick={() => {
                  setStartDate('');
                  setEndDate('');
                  showAlert('日期筛选条件已重置，加载全部数据。');
                }}
                className="flex-1 py-2 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl text-xs font-semibold text-slate-700 transition cursor-pointer text-center"
              >
                重置时间范围
              </button>
            </div>
          </div>
        </div>

        {/* METRICS & QUICK SUMMARY WIDGETS */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm flex flex-col justify-between h-36">
            <span className="text-sm text-slate-400 font-semibold">核算班级总数</span>
            <div>
              <div className="text-3xl font-extrabold text-slate-800 tracking-tight">{overviewStats.totalClasses} <span className="text-xs font-normal text-slate-400">个</span></div>
              <p className="text-[10px] text-slate-400 mt-1">Excel中识别出的独立班级</p>
            </div>
          </div>
          
          <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm flex flex-col justify-between h-36">
            <span className="text-sm text-slate-400 font-semibold">累计授课次数</span>
            <div>
              <div className="text-3xl font-extrabold text-slate-800 tracking-tight">{overviewStats.totalLessonsCount} <span className="text-xs font-normal text-slate-400">次</span></div>
              <p className="text-[10px] text-slate-400 mt-1">选定时间区间内的总授课次数</p>
            </div>
          </div>

          <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm flex flex-col justify-between h-36">
            <span className="text-sm text-slate-400 font-semibold">累计结算课时</span>
            <div>
              <div className="text-3xl font-extrabold text-indigo-600 tracking-tight">{overviewStats.totalHours.toFixed(1)} <span className="text-xs font-normal text-slate-400">hrs</span></div>
              <p className="text-[10px] text-slate-400 mt-1">
                乘学生人数: {overviewStats.totalHours.toFixed(1)} / 不计人数纯课时: <strong className="text-slate-700 font-bold">{overviewStats.totalBaseHours.toFixed(1)}</strong>
              </p>
            </div>
          </div>

          <div className="bg-indigo-600 text-white rounded-2xl shadow-lg p-5 flex flex-col justify-between h-36 border border-indigo-700/50">
            <span className="text-sm opacity-80 uppercase tracking-widest font-semibold">教师实发课酬总支出</span>
            <div>
              <div className="text-3xl font-black tracking-tight">¥ {overviewStats.totalSalary.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</div>
              <p className="text-[10px] opacity-70 mt-1">课销提成 ({(commissionRate * 100).toFixed(0)}%) + 特定课时费加成</p>
            </div>
          </div>
        </div>

        {/* MAIN NAVIGATION TABS */}
        <div className="bg-white border border-slate-200/80 rounded-2xl shadow-sm overflow-hidden">
          <div className="bg-slate-100/60 border-b border-slate-200 px-6 py-4 flex flex-wrap items-center justify-between gap-4">
            <div className="bg-slate-200/50 p-1 rounded-xl inline-flex flex-wrap gap-1">
              <button
                onClick={() => setActiveTab('statistics')}
                className={`px-4 py-2.5 text-xs font-bold rounded-lg transition-all flex items-center gap-2 cursor-pointer ${
                  activeTab === 'statistics'
                    ? 'bg-white text-indigo-600 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <BookOpen className="w-3.5 h-3.5" />
                课消统计明细表
              </button>
              <button
                onClick={() => setActiveTab('settings')}
                className={`px-4 py-2.5 text-xs font-bold rounded-lg transition-all flex items-center gap-2 cursor-pointer ${
                  activeTab === 'settings'
                    ? 'bg-white text-indigo-600 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Settings className="w-3.5 h-3.5" />
                课时费加成 & 教师基础单价
              </button>
              <button
                onClick={() => setActiveTab('substitutions')}
                className={`px-4 py-2.5 text-xs font-bold rounded-lg transition-all flex items-center gap-2 cursor-pointer ${
                  activeTab === 'substitutions'
                    ? 'bg-white text-indigo-600 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <UserCheck className="w-3.5 h-3.5" />
                代课抵账管理
              </button>
              <button
                onClick={() => setActiveTab('makeups')}
                className={`px-4 py-2.5 text-xs font-bold rounded-lg transition-all flex items-center gap-2 cursor-pointer ${
                  activeTab === 'makeups'
                    ? 'bg-white text-indigo-600 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <CheckSquare className="w-3.5 h-3.5" />
                补课课销登记
              </button>
              <button
                onClick={() => setActiveTab('report')}
                className={`px-4 py-2.5 text-xs font-bold rounded-lg transition-all flex items-center gap-2 cursor-pointer ${
                  activeTab === 'report'
                    ? 'bg-white text-indigo-600 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <FileDown className="w-3.5 h-3.5" />
                月度课销薪资报表
              </button>
            </div>
            {classBlocks.length > 0 && activeTab === 'statistics' && (
              <div className="text-xs text-slate-500 font-mono">
                过滤后已加载: <strong className="text-indigo-600 font-bold">{resolvedLessons.length}</strong> 条记录
              </div>
            )}
          </div>

          <div className="p-6">
            
            {/* TAB 1: STATISTICS */}
            {activeTab === 'statistics' && (
              <div className="space-y-6">
                
                {/* INTRO FOR STATISTICS */}
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-indigo-50/40 p-4 rounded-2xl border border-indigo-100/50">
                  <div className="space-y-1">
                    <h3 className="text-sm font-bold text-slate-800 flex items-center gap-1.5">
                      <Info className="w-4 h-4 text-indigo-600" />
                      数据核算说明
                    </h3>
                    <p className="text-xs text-slate-500 leading-relaxed max-w-3xl">
                      系统已根据上课记录的日期分布<strong>自动判定</strong>各班是一周一次还是两次，您也可以在下方“排课频次”下拉选择框中手动微调或重算。
                      <br />
                      换算规则：
                      <span className="inline-block bg-white border border-slate-100 rounded px-1.5 py-0.5 font-mono text-[10px] mx-1">一周一次课：中教3小时 / 外教2小时</span> ； 
                      <span className="inline-block bg-white border border-slate-100 rounded px-1.5 py-0.5 font-mono text-[10px] mx-1">一周两次课：中教2小时 / 外教1小时</span> 。
                    </p>
                  </div>
                  {classBlocks.length > 0 && (
                    <button
                      onClick={autoDetectAllFrequencies}
                      className="inline-flex items-center gap-1.5 text-xs font-bold text-indigo-700 bg-indigo-100 hover:bg-indigo-200 border border-indigo-200/50 px-4 py-2.5 rounded-xl transition cursor-pointer shrink-0"
                      title="根据实际上课记录日期与每周分布再次自动计算频次"
                    >
                      <RefreshCw className="w-3.5 h-3.5" />
                      重新智能判定频次
                    </button>
                  )}
                </div>

                {classBlocks.length === 0 ? (
                  <div className="py-12 flex flex-col items-center justify-center text-center space-y-3">
                    <div className="bg-slate-100 p-4 rounded-full text-slate-400">
                      <FileSpreadsheet className="w-8 h-8" />
                    </div>
                    <div className="space-y-1">
                      <h4 className="text-sm font-bold text-slate-800">尚未导入排课 Excel 数据</h4>
                      <p className="text-xs text-slate-400 max-w-sm">
                        请在页面上方拖拽或上传您机构的 Excel 课时打卡表格，即可自动进行全自动多班级统计。
                      </p>
                    </div>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 xl:grid-cols-12 gap-6">
                    
                     {/* LESSONS MAIN TABLE */}
                    <div className="xl:col-span-8 space-y-4">
                      
                      {/* Filter Index Controls */}
                      <div className="bg-white border border-slate-200/80 rounded-2xl p-4 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="flex items-center gap-2">
                          <div className="bg-indigo-50 p-2 rounded-xl text-indigo-600">
                            <Search className="w-4 h-4" />
                          </div>
                          <div>
                            <h3 className="text-xs font-bold text-slate-800">
                              课消明细筛选索引
                            </h3>
                            <p className="text-[10px] text-slate-400">
                              当前区间共 {resolvedLessons.length} 节记录
                              {(filterTeacher || filterClassCode) && (
                                <span className="ml-1 text-indigo-600 font-semibold">
                                  （筛选出 {filteredResolvedLessons.length} 节）
                                </span>
                              )}
                            </p>
                          </div>
                        </div>

                        <div className="flex flex-wrap items-center gap-3">
                          {/* Teacher selector */}
                          <div className="flex items-center gap-1.5">
                            <label className="text-xs font-semibold text-slate-500 whitespace-nowrap">授课老师:</label>
                            <select
                              value={filterTeacher}
                              onChange={(e) => setFilterTeacher(e.target.value)}
                              className="text-xs bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2.5 py-2 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition cursor-pointer"
                            >
                              <option value="">全部老师</option>
                              {uniqueTeachers.map(t => (
                                <option key={t} value={t}>{t}</option>
                              ))}
                            </select>
                          </div>

                          {/* Class selector */}
                          <div className="flex items-center gap-1.5">
                            <label className="text-xs font-semibold text-slate-500 whitespace-nowrap">上课班级:</label>
                            <select
                              value={filterClassCode}
                              onChange={(e) => setFilterClassCode(e.target.value)}
                              className="text-xs bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2.5 py-2 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition max-w-[180px] cursor-pointer"
                            >
                              <option value="">全部班级</option>
                              {uniqueClassesList.map(c => (
                                <option key={c.classCode} value={c.classCode}>
                                  {c.className}
                                </option>
                              ))}
                            </select>
                          </div>

                          {/* Clear filters button */}
                          {(filterTeacher || filterClassCode) && (
                            <button
                              onClick={() => {
                                setFilterTeacher('');
                                setFilterClassCode('');
                              }}
                              className="text-xs text-indigo-600 hover:text-indigo-800 font-bold hover:underline transition cursor-pointer"
                            >
                              清除筛选
                            </button>
                          )}
                        </div>
                      </div>

                      <div className="overflow-hidden border border-slate-200/80 rounded-2xl bg-white shadow-sm">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="bg-slate-50/80 border-b border-slate-200 text-slate-500 font-bold text-xs uppercase tracking-wider">
                              <th className="px-4 py-3.5">日期</th>
                              <th className="px-4 py-3.5">班级名称</th>
                              <th className="px-4 py-3.5">老师</th>
                              <th className="px-3 py-3.5">中/外</th>
                              <th className="px-3 py-3.5 text-center">排课频次</th>
                              <th className="px-3 py-3.5 text-center">参课人数</th>
                              <th className="px-3 py-3.5 text-right text-indigo-600" title="标准单节课时（不乘学生人数，由频次和中/外教决定）">单节标准课时</th>
                              <th className="px-3 py-3.5 text-right" title="换算结算课时 = 参课人数 × 单节标准课时">核算结算课时</th>
                              <th className="px-4 py-3.5 text-center">操作</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-sm text-slate-700">
                            {filteredResolvedLessons.length === 0 ? (
                              <tr>
                                <td colSpan={9} className="px-4 py-10 text-center text-xs text-slate-400">
                                  {resolvedLessons.length === 0 ? '该日期区间内无符合条件的授课记录' : '无符合筛选条件的授课记录'}
                                </td>
                              </tr>
                            ) : (
                              filteredResolvedLessons.map((item, idx) => {
                                const bonusExist = bonusRules.some(r => 
                                  r.teacherName.trim().toLowerCase() === item.actualTeacher.trim().toLowerCase() && 
                                  r.classCode === item.classCode
                                );
                                const isSelected = selectedBlockId === item.block.id && selectedLessonIndex === item.lesson.index;
                                const isOverridden = !!item.lesson.teacherOverride || 
                                                     item.lesson.hoursOverride !== undefined || 
                                                     item.lesson.baseHoursOverride !== undefined;
                                
                                return (
                                  <tr 
                                    key={`${item.block.id}_lesson_${idx}`}
                                    className={`transition-colors ${
                                      isSelected 
                                        ? 'bg-indigo-50/50' 
                                        : idx % 2 === 1 
                                          ? 'bg-slate-50/30 hover:bg-slate-100/50' 
                                          : 'bg-white hover:bg-slate-100/50'
                                    }`}
                                  >
                                    <td className="px-4 py-3 font-mono text-xs whitespace-nowrap">{item.lesson.dateStr}</td>
                                    <td className="px-4 py-3">
                                      <div className="font-bold text-slate-800">{item.block.className}</div>
                                      <div className="text-[10px] text-slate-400 font-mono">编号: {item.classCode}</div>
                                    </td>
                                    <td className="px-4 py-3 whitespace-nowrap">
                                      <div className="flex flex-col">
                                        <span className="font-medium">{item.actualTeacher}</span>
                                        {item.isSubstituted && (
                                          <span className="text-[9px] bg-amber-50 text-amber-700 px-1 py-0.2 rounded border border-amber-200/50 mt-0.5 text-center truncate" title={`代替原任课老师 ${item.block.teacher} 授课`}>
                                            代课自: {item.block.teacher}
                                          </span>
                                        )}
                                      </div>
                                    </td>
                                    <td className="px-3 py-3 whitespace-nowrap">
                                      <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${
                                        item.resolvedTeacherType === '外教' 
                                          ? 'bg-amber-50 text-amber-700 border border-amber-100' 
                                          : 'bg-indigo-50 text-indigo-700 border border-indigo-100'
                                      }`}>
                                        {item.resolvedTeacherType}
                                      </span>
                                    </td>
                                    <td className="px-3 py-3 text-center whitespace-nowrap">
                                      <select 
                                        value={item.block.frequency}
                                        onChange={(e) => updateClassFrequency(item.block.id, e.target.value as 'once' | 'twice')}
                                        className="text-xs bg-slate-50 border border-slate-200 rounded px-1.5 py-0.5 text-slate-700 outline-none focus:ring-1 focus:ring-indigo-500"
                                      >
                                        <option value="once">一周一次 (单次)</option>
                                        <option value="twice">一周两次 (双次)</option>
                                      </select>
                                    </td>
                                    <td className="px-3 py-3 text-center font-medium font-mono">
                                      {item.lesson.attendedCount} <span className="text-xs text-slate-400">/ {item.lesson.totalStudents}</span>
                                    </td>
                                    <td className="px-3 py-3 text-right font-medium text-indigo-600 font-mono">
                                      {item.baseHours.toFixed(1)}
                                    </td>
                                    <td className="px-3 py-3 text-right font-bold text-slate-900 font-mono">
                                      {item.hours.toFixed(1)}
                                      {bonusExist && (
                                        <span className="inline-block text-[10px] text-emerald-600 bg-emerald-50 px-1 rounded ml-1 border border-emerald-100" title="该班级含有课时加成">
                                          +加
                                        </span>
                                      )}
                                      {isOverridden && (
                                        <span className="inline-block text-[10px] text-amber-600 bg-amber-50 px-1 rounded ml-1 border border-amber-200" title="该上课记录包含手动修正的数值">
                                          已修正
                                        </span>
                                      )}
                                    </td>
                                    <td className="px-4 py-3 text-center whitespace-nowrap">
                                      <button
                                        onClick={() => {
                                          setSelectedBlockId(item.block.id);
                                          setSelectedLessonIndex(item.lesson.index);
                                        }}
                                        className="inline-flex items-center gap-1 text-xs text-indigo-600 hover:text-indigo-800 font-medium cursor-pointer"
                                      >
                                        学生考勤
                                        <ChevronRight className="w-3.5 h-3.5" />
                                      </button>
                                    </td>
                                  </tr>
                                );
                              })
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* ATTENDANCE SIDE DETAIL PANEL */}
                    <div className="xl:col-span-4 space-y-4">
                      {selectedBlock && selectedLessonIndex !== null ? (
                        (() => {
                          const matchedLesson = selectedBlock.lessons.find(l => l.index === selectedLessonIndex);
                          if (!matchedLesson) return null;

                          const resolvedItem = resolvedLessons.find(item => item.block.id === selectedBlock.id && item.lesson.index === selectedLessonIndex);
                          const currentResolvedType = resolvedItem?.resolvedTeacherType || matchedLesson.type;

                          return (
                            <div className="bg-white border border-slate-200/80 rounded-2xl p-6 shadow-sm space-y-5 sticky top-24 animate-in fade-in duration-200">
                              <div className="border-b border-slate-100 pb-3.5 flex justify-between items-start gap-2">
                                <div className="space-y-1.5">
                                  <span className="text-[10px] bg-indigo-50 text-indigo-700 px-2.5 py-1 rounded border border-indigo-100/80 font-bold uppercase tracking-wider">
                                    考勤名单抽查
                                  </span>
                                  <h3 className="font-bold text-slate-800 text-sm mt-1">{selectedBlock.className}</h3>
                                  <p className="text-xs font-mono text-slate-400">日期: {matchedLesson.dateStr} | T. {selectedBlock.teacher}</p>
                                </div>
                                <button 
                                  onClick={() => {
                                    setSelectedBlockId('');
                                    setSelectedLessonIndex(null);
                                  }}
                                  className="p-1 rounded-lg text-slate-400 hover:bg-slate-50 transition cursor-pointer"
                                >
                                  <XCircle className="w-5 h-5" />
                                </button>
                              </div>

                              {/* ATTENDANCE LIST GRID */}
                              <div className="space-y-2.5 max-h-[360px] overflow-y-auto pr-1">
                                {selectedBlock.students.length === 0 ? (
                                  <p className="text-xs text-slate-400 text-center py-6">此班级没有学生名单</p>
                                ) : (
                                  selectedBlock.students.map(student => {
                                    const status = matchedLesson.studentStatus[student.chineseName] || '-';
                                    const isAttended = status === '是';
                                    const isLeave = ['事假', '病假'].includes(status);
                                    
                                    return (
                                      <div 
                                        key={student.chineseName}
                                        className={`flex items-center justify-between p-2.5 rounded-xl border text-xs transition-colors ${
                                          isAttended 
                                            ? 'bg-emerald-50/20 border-emerald-100/50' 
                                            : isLeave 
                                              ? 'bg-amber-50/20 border-amber-100/50'
                                              : 'bg-slate-50/50 border-slate-100'
                                        }`}
                                      >
                                        <div className="space-y-0.5">
                                          <div className="font-bold text-slate-700">{student.chineseName}</div>
                                          {student.englishName && (
                                            <div className="text-[10px] text-slate-400 font-mono">{student.englishName}</div>
                                          )}
                                        </div>
                                        
                                        <span className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-bold ${
                                          isAttended 
                                            ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' 
                                            : status === '停课' 
                                              ? 'bg-rose-50 text-rose-700 border border-rose-100'
                                              : isLeave
                                                ? 'bg-amber-50 text-amber-700 border border-amber-100'
                                                : 'bg-slate-100 text-slate-500 border border-slate-200'
                                        }`}>
                                          {isAttended && <CheckSquare className="w-3.5 h-3.5 shrink-0" />}
                                          {status}
                                        </span>
                                      </div>
                                    );
                                  })
                                )}
                              </div>

                              <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-[11px] text-slate-500 leading-relaxed">
                                <div className="font-bold text-slate-600 mb-1 flex items-center gap-1">
                                  <Info className="w-3.5 h-3.5 text-slate-400" />
                                  换算课时公式
                                </div>
                                <p>
                                  参课人数 ({matchedLesson.attendedCount}人) × 课时基数 ({calculateLessonBase(selectedBlock.frequency, currentResolvedType)}小时/课) = 
                                  <strong className="text-indigo-600 ml-1">{(matchedLesson.attendedCount * calculateLessonBase(selectedBlock.frequency, currentResolvedType)).toFixed(1)} 课时</strong>。
                                </p>
                              </div>

                              <div className="border-t border-slate-100 pt-4 space-y-3">
                                <div className="flex items-center gap-1.5 text-xs font-bold text-slate-800">
                                  <Edit className="w-3.5 h-3.5 text-indigo-600" />
                                  手动修正本节排课数据 (少数意外)
                                </div>
                                
                                <div className="grid grid-cols-2 gap-3 text-xs">
                                  {/* 课程属性 / 中外教 */}
                                  <div className="space-y-1">
                                    <label className="font-semibold text-slate-500">中/外教属性</label>
                                    <select
                                      value={editType}
                                      onChange={(e) => setEditType(e.target.value)}
                                      className="w-full bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition cursor-pointer"
                                    >
                                      <option value="中教">中教课</option>
                                      <option value="外教">外教课</option>
                                    </select>
                                  </div>

                                  {/* 参课人数 */}
                                  <div className="space-y-1">
                                    <label className="font-semibold text-slate-500">实际参课人数</label>
                                    <input
                                      type="number"
                                      min="0"
                                      value={editAttendedCount}
                                      onChange={(e) => setEditAttendedCount(Number(e.target.value))}
                                      className="w-full bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition"
                                    />
                                  </div>
                                </div>

                                <div className="space-y-1 text-xs">
                                  {/* 实际授课老师 */}
                                  <label className="font-semibold text-slate-500">实际授课老师 (可代课或修正)</label>
                                  <div className="flex gap-2">
                                    <select
                                      value={editTeacher}
                                      onChange={(e) => setEditTeacher(e.target.value)}
                                      className="flex-1 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition cursor-pointer"
                                    >
                                      <option value="">默认老师 ({selectedBlock.teacher})</option>
                                      {uniqueTeachers.filter(t => t !== selectedBlock.teacher).map(t => (
                                        <option key={t} value={t}>{t}</option>
                                      ))}
                                      <option value="__custom__">自定义输入...</option>
                                    </select>
                                    
                                    {(editTeacher === '__custom__' || (!uniqueTeachers.includes(editTeacher) && editTeacher !== '')) && (
                                      <input
                                        type="text"
                                        placeholder="输入老师姓名"
                                        value={editTeacher === '__custom__' ? '' : editTeacher}
                                        onChange={(e) => setEditTeacher(e.target.value)}
                                        className="w-28 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition"
                                      />
                                    )}
                                  </div>
                                </div>

                                <div className="grid grid-cols-2 gap-3 text-xs">
                                  {/* 单节课时基数 override */}
                                  <div className="space-y-1">
                                    <label className="font-semibold text-slate-500" title="若不填，则根据上课频次和中外教属性自动换算">
                                      单节课时 (选填)
                                    </label>
                                    <input
                                      type="text"
                                      placeholder="自动换算"
                                      value={editBaseHours}
                                      onChange={(e) => setEditBaseHours(e.target.value)}
                                      className="w-full bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2.5 py-1 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition"
                                    />
                                  </div>

                                  {/* 最终核算结算课时 override */}
                                  <div className="space-y-1">
                                    <label className="font-semibold text-slate-500" title="若不填，则通过 参课人数 × 单节课时 自动计算">
                                      核算结算课时 (选填)
                                    </label>
                                    <input
                                      type="text"
                                      placeholder="自动计算"
                                      value={editHours}
                                      onChange={(e) => setEditHours(e.target.value)}
                                      className="w-full bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl px-2.5 py-1 text-slate-700 outline-none focus:ring-2 focus:ring-indigo-100 focus:border-indigo-500 transition"
                                    />
                                  </div>
                                </div>

                                <div className="flex gap-2 pt-1">
                                  <button
                                    onClick={() => {
                                      handleUpdateLessonOverride(selectedBlock.id, matchedLesson.index, {
                                        type: editType,
                                        attendedCount: Number(editAttendedCount),
                                        teacherOverride: editTeacher === '__custom__' ? '' : editTeacher,
                                        baseHoursOverride: editBaseHours === '' ? null : Number(editBaseHours),
                                        hoursOverride: editHours === '' ? null : Number(editHours)
                                      });
                                    }}
                                    className="flex-1 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-semibold text-xs transition cursor-pointer shadow-sm text-center"
                                  >
                                    保存修正
                                  </button>
                                  <button
                                    onClick={() => {
                                      // Clear manual overrides
                                      handleUpdateLessonOverride(selectedBlock.id, matchedLesson.index, {
                                        type: undefined,
                                        teacherOverride: '',
                                        hoursOverride: null,
                                        baseHoursOverride: null
                                      });
                                      // Restore local states to default parsed values
                                      setEditType(matchedLesson.type);
                                      setEditAttendedCount(matchedLesson.attendedCount);
                                      setEditTeacher('');
                                      setEditBaseHours('');
                                      setEditHours('');
                                    }}
                                    className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-semibold text-xs transition cursor-pointer text-center"
                                    title="恢复至排课表初始自动计算状态"
                                  >
                                    恢复默认
                                  </button>
                                </div>
                              </div>
                            </div>
                          );
                        })()
                      ) : (
                        <div className="bg-slate-50 border border-slate-200/80 border-dashed rounded-2xl p-6 text-center text-slate-400 h-64 flex flex-col items-center justify-center space-y-2 sticky top-24">
                          <Users className="w-8 h-8 text-slate-300" />
                          <p className="text-xs font-bold text-slate-500">点击列表中任一行的“学生考勤”</p>
                          <p className="text-[10px] text-slate-400">可抽调查看每节课的学员打卡考勤、请假或停课明细</p>
                        </div>
                      )}
                    </div>

                  </div>
                )}
              </div>
            )}

            {/* TAB 2: CONFIGURATION */}
            {activeTab === 'settings' && (
              <div className="space-y-8 animate-in fade-in duration-200">
                
                {/* CONFIGURATION INTRODUCTION */}
                <div className="bg-indigo-50/40 p-5 rounded-2xl border border-indigo-100/50 flex gap-3">
                  <div className="bg-indigo-100 text-indigo-700 p-2 rounded-xl shrink-0 self-start">
                    <Settings className="w-5 h-5" />
                  </div>
                  <div className="space-y-1">
                    <h3 className="text-sm font-bold text-slate-800">高精准课时费匹配规则</h3>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      为了彻底避免相同名称班级导致冗余计算，课消管家对“特定班级课时费加成”采用 <strong>班级数字编号 + 任课老师姓名</strong> 这一高精准的维度匹配。
                      <br />例如：A老师的 <span className="bg-white px-1 border border-slate-200 rounded">E3.210512</span> 班加成 “5元”，只对该班级的课时数以及A老师的打卡进行核算，而不会误加给同校区B老师的班。班级名称前的英文前缀（E1, E2, B1等）在匹配时将自动被忽略。
                    </p>
                  </div>
                </div>



                <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
                  
                  {/* LEFT: ADD BONUS CONFIG & TEACHER RATES FORM */}
                  <div className="lg:col-span-4 space-y-6">
                    
                    {/* ADD BONUS RULE FORM */}
                    <div id="bonus-config-form" className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                      <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2.5 flex items-center gap-1.5">
                        <Plus className="w-4.5 h-4.5 text-indigo-600" />
                        添加特定班级课时费加成
                      </h3>

                      <form onSubmit={handleAddBonusRule} className="space-y-3.5">
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">任课教师</label>
                          <input 
                            type="text" 
                            list="teachers-datalist"
                            placeholder="如: Ann"
                            value={newRuleTeacher}
                            onChange={(e) => setNewRuleTeacher(e.target.value)}
                            className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500"
                            required
                          />
                          <datalist id="teachers-datalist">
                            {uniqueTeachers.map(t => <option key={t} value={t} />)}
                          </datalist>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">班级编号 (提取的尾部纯数字)</label>
                          <input 
                            type="text" 
                            list="classcodes-datalist"
                            placeholder="如: 210512 (不需要填E3.)"
                            value={newRuleClassCode}
                            onChange={(e) => setNewRuleClassCode(e.target.value)}
                            className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 font-mono"
                            required
                          />
                          <datalist id="classcodes-datalist">
                            {uniqueClassCodes.map(code => <option key={code} value={code} />)}
                          </datalist>
                          <p className="text-[10px] text-slate-400 mt-1">系统会自动剔除前缀。班级“E3.210512”对应的编号为 210512。</p>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">加成金额 (元/课时)</label>
                          <div className="relative">
                            <span className="absolute left-3 top-2 text-slate-400 text-xs">¥</span>
                            <input 
                              type="number" 
                              step="0.1"
                              placeholder="如: 5"
                              value={newRuleBonusRate}
                              onChange={(e) => setNewRuleBonusRate(e.target.value)}
                              className="w-full text-xs border border-slate-200 rounded-lg pl-7 pr-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500"
                              required
                            />
                          </div>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">备注说明 (可选)</label>
                          <input 
                            type="text" 
                            placeholder="如: 暑期课时加成"
                            value={newRuleNotes}
                            onChange={(e) => setNewRuleNotes(e.target.value)}
                            className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                          />
                        </div>

                        <button
                          type="submit"
                          className="w-full py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-medium shadow transition cursor-pointer"
                        >
                          确认添加课时费加成
                        </button>
                      </form>
                    </div>

                    {/* TEACHER BASE RATE STANDARD FORM */}
                    <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                      <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2.5 flex items-center gap-1.5">
                        <DollarSign className="w-4.5 h-4.5 text-indigo-600" />
                        教师基础课时单价配置
                      </h3>
                      
                      <form onSubmit={handleAddTeacherRate} className="space-y-3.5">
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">教师姓名</label>
                          <input 
                            type="text" 
                            list="teachers-datalist"
                            placeholder="如: Soup"
                            value={newTeacherName}
                            onChange={(e) => setNewTeacherName(e.target.value)}
                            className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                            required
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">基础单价 (元/课时)</label>
                          <div className="relative">
                            <span className="absolute left-3 top-2 text-slate-400 text-xs">¥</span>
                            <input 
                              type="number" 
                              placeholder="如: 120"
                              value={newTeacherBaseRate}
                              onChange={(e) => setNewTeacherBaseRate(e.target.value)}
                              className="w-full text-xs border border-slate-200 rounded-lg pl-7 pr-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                              required
                            />
                          </div>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">师资类别 (中教 / 外教)</label>
                          <select
                            value={newTeacherType}
                            onChange={(e) => setNewTeacherType(e.target.value as '中教' | '外教')}
                            className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500 bg-white cursor-pointer"
                          >
                            <option value="中教">中教 (Chinese Teacher)</option>
                            <option value="外教">外教 (Foreign Teacher)</option>
                          </select>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1">专属课消提成比例</label>
                          <select
                            value={newTeacherCommissionRate}
                            onChange={(e) => setNewTeacherCommissionRate(parseFloat(e.target.value))}
                            className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500 bg-white cursor-pointer"
                          >
                            <option value={0.07}>7% (默认)</option>
                            <option value={0.06}>6%</option>
                            <option value={0.05}>5%</option>
                            <option value={0.08}>8%</option>
                            <option value={0.09}>9%</option>
                            <option value={0.10}>10%</option>
                          </select>
                        </div>

                        <button
                          type="submit"
                          className="w-full py-2 bg-slate-800 hover:bg-slate-900 text-white rounded-lg text-xs font-medium shadow transition cursor-pointer"
                        >
                          更新或保存该教师单价
                        </button>
                      </form>
                    </div>

                  </div>

                  {/* RIGHT: LIST OF CONFIGURED RULES & SUGGESTIONS */}
                  <div className="lg:col-span-8 space-y-6">
                    
                    {/* LIST OF SPECIFIC BONUS RULES */}
                    <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                      <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2">
                        当前已配置的班级加成单价 (匹配: 教师 + 编号)
                      </h3>

                      <div className="overflow-x-auto border border-slate-100 rounded-lg">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="bg-slate-50 border-b border-slate-200 text-slate-400 text-xs font-semibold">
                              <th className="px-4 py-2.5">任课教师</th>
                              <th className="px-4 py-2.5">班级编号</th>
                              <th className="px-4 py-2.5 text-right">加成金额</th>
                              <th className="px-4 py-2.5">备注说明</th>
                              <th className="px-4 py-2.5 text-center">操作</th>
                            </tr>
                          </thead>
                          <tbody className="text-xs text-slate-700 divide-y divide-slate-100">
                            {bonusRules.length === 0 ? (
                              <tr>
                                <td colSpan={5} className="px-4 py-10 text-center text-slate-400">
                                  暂未配置任何课时费加成。您可以在下方查看“导入班级建议”一键快速配置。
                                </td>
                              </tr>
                            ) : (
                              bonusRules.map((rule) => (
                                <tr key={rule.id} className="hover:bg-slate-50">
                                  <td className="px-4 py-2.5 font-semibold text-slate-800">{rule.teacherName}</td>
                                  <td className="px-4 py-2.5 font-mono text-indigo-600 bg-indigo-50/20 rounded-md font-medium">{rule.classCode}</td>
                                  <td className="px-4 py-2.5 text-right font-bold text-emerald-600 font-mono">¥ {rule.bonusRate.toFixed(1)} /课时</td>
                                  <td className="px-4 py-2.5 text-slate-500">{rule.notes || '--'}</td>
                                  <td className="px-4 py-2.5 text-center">
                                    <button
                                      onClick={() => handleDeleteBonusRule(rule.id)}
                                      className="text-rose-600 hover:text-rose-800 p-1 cursor-pointer"
                                      title="删除此项"
                                    >
                                      <Trash2 className="w-4 h-4 mx-auto" />
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* LIST OF TEACHERS BASE RATE */}
                    <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                      <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2">
                        教师课时提成基础单价
                      </h3>
                      
                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                        {teacherBaseRates.length === 0 ? (
                          <div className="col-span-full py-6 text-center text-xs text-slate-400">
                            暂未配置任何教师课时单价，结算时将默认采用 ¥ 100/课时。
                          </div>
                        ) : (
                          teacherBaseRates.map((rate) => (
                            <div key={rate.teacherName} className="border border-slate-200 rounded-lg p-3 flex justify-between items-center bg-slate-50/50">
                              <div className="space-y-1">
                                <div className="flex items-center gap-1.5">
                                  <span className="text-xs font-bold text-slate-700">{rate.teacherName}</span>
                                  <button
                                    onClick={() => toggleTeacherType(rate.teacherName)}
                                    className={`text-[9px] px-1.5 py-0.2 rounded font-semibold border transition cursor-pointer select-none ${
                                      rate.teacherType === '外教'
                                        ? 'bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100'
                                        : 'bg-indigo-50 text-indigo-700 border-indigo-200 hover:bg-indigo-100'
                                    }`}
                                    title="点击一键切换中/外教属性"
                                  >
                                    {rate.teacherType || '中教'}
                                  </button>
                                </div>
                                <div className="flex items-center gap-4">
                                  <span className="text-xs font-bold text-indigo-600 font-mono">¥ {rate.baseRate} /课时</span>
                                  <div className="flex items-center gap-1 text-[11px] text-slate-500">
                                    <span>提成:</span>
                                    <select
                                      value={rate.commissionRate !== undefined ? rate.commissionRate : commissionRate}
                                      onChange={(e) => updateTeacherCommissionRate(rate.teacherName, parseFloat(e.target.value))}
                                      className="text-[11px] font-bold text-indigo-600 bg-transparent border-none cursor-pointer focus:ring-0 p-0 outline-none"
                                    >
                                      <option value={0.07}>7% (默认)</option>
                                      <option value={0.06}>6%</option>
                                      <option value={0.05}>5%</option>
                                      <option value={0.08}>8%</option>
                                      <option value={0.09}>9%</option>
                                      <option value={0.10}>10%</option>
                                    </select>
                                  </div>
                                </div>
                              </div>
                              <button
                                onClick={() => handleDeleteTeacherRate(rate.teacherName)}
                                className="text-slate-400 hover:text-rose-600 p-1 cursor-pointer"
                                title="清除"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          ))
                        )}
                      </div>
                    </div>

                    {/* SUGGESTIONS LIST CARD */}
                    {classBlocks.length > 0 && suggestedRules.length > 0 && (
                      <div className="bg-indigo-50/20 border border-indigo-100 rounded-xl p-5 space-y-3">
                        <div className="flex items-center gap-1.5">
                          <Sparkles className="w-4.5 h-4.5 text-indigo-600" />
                          <h4 className="text-xs font-bold text-slate-800">智能建议：从导入的数据快速配置加成</h4>
                        </div>
                        <p className="text-[11px] text-slate-500">
                          根据您刚刚导入的排课表格，以下班级拥有独立的编号和老师，您可以直接点击一键添加加成。
                        </p>
                        <div className="flex flex-wrap gap-2 pt-1">
                          {suggestedRules.map(s => (
                            <button
                              key={`${s.teacherName}_${s.classCode}`}
                              onClick={() => addSuggestedRule(s)}
                              className="bg-white border border-slate-200 rounded-lg px-2.5 py-1 text-xs font-medium text-slate-700 hover:border-indigo-400 hover:bg-indigo-50/30 transition flex items-center gap-1 cursor-pointer text-left"
                            >
                              <span className="font-semibold text-slate-800">{s.teacherName}</span>
                              <span className="text-slate-400">({s.classCode})</span>
                              <span className="text-indigo-600 ml-1 font-mono">+ 加成</span>
                            </button>
                          ))}
                        </div>
                      </div>
                    )}

                  </div>
                </div>

              </div>
            )}

            {/* TAB 3: SUBSTITUTION MANAGEMENT */}
            {activeTab === 'substitutions' && (
              <div className="space-y-6">
                
                {/* SUBSTITUTION INTRODUCTION */}
                <div className="bg-amber-50/40 p-4.5 rounded-xl border border-amber-100/50 flex gap-3">
                  <div className="bg-amber-100 text-amber-800 p-2 rounded-lg shrink-0 self-start">
                    <UserCheck className="w-5 h-5" />
                  </div>
                  <div className="space-y-1">
                    <h3 className="text-sm font-semibold text-slate-800">教师代课自动对调与结算机制</h3>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      当代课老师代上属于原老师的班级时，管理员在此处录入代课记录。
                      <br /><strong>自动重算</strong>：系统核算时会自动扣减原任课老师该节课的课时数与薪酬，并将其全额转移、合并到代课老师的月度工资报表中。
                    </p>
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
                  
                  {/* LEFT: ADD SUBSTITUTION RECORD FORM */}
                  <div className="lg:col-span-4 bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                    <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2 flex items-center gap-1.5">
                      <Plus className="w-4.5 h-4.5 text-indigo-600" />
                      新增代课抵扣记录
                    </h3>

                    <form onSubmit={handleAddSubstitution} className="space-y-4">
                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">上课日期</label>
                        <input 
                          type="date" 
                          value={subDate}
                          onChange={(e) => setSubDate(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">代课的班级编号</label>
                        <select
                          value={subClassCode}
                          onChange={(e) => {
                            setSubClassCode(e.target.value);
                            // Auto fill original teacher based on class code selection
                            const matched = classBlocks.find(b => b.classCode === e.target.value);
                            if (matched) {
                              setSubOriginalTeacher(matched.teacher);
                            }
                          }}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 bg-white outline-none focus:ring-1 focus:ring-indigo-500"
                          required
                        >
                          <option value="">-- 选择班级 --</option>
                          {classBlocks.map(b => (
                            <option key={b.id} value={b.classCode}>{b.className} (原任: {b.teacher})</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">原任课教师</label>
                        <input 
                          type="text" 
                          placeholder="原班主任教师"
                          value={subOriginalTeacher}
                          onChange={(e) => setSubOriginalTeacher(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">实际去代课教师 (Substitute)</label>
                        <input 
                          type="text" 
                          list="teachers-datalist"
                          placeholder="实际授课的代课教师"
                          value={subSubstituteTeacher}
                          onChange={(e) => setSubSubstituteTeacher(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">代课备注 (可选)</label>
                        <input 
                          type="text" 
                          placeholder="如: 病假代课"
                          value={subNotes}
                          onChange={(e) => setSubNotes(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                        />
                      </div>

                      <button
                        type="submit"
                        className="w-full py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-medium shadow transition cursor-pointer"
                      >
                        录入代课并应用
                      </button>
                    </form>
                  </div>

                  {/* RIGHT: LIST OF SUBSTITUTION RECORDS */}
                  <div className="lg:col-span-8 bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                    <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2">
                      代课置换记录明细 ({substitutionRecords.length} 项)
                    </h3>

                    <div className="overflow-x-auto border border-slate-100 rounded-lg">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200 text-slate-400 text-xs font-semibold">
                            <th className="px-4 py-2.5">代课日期</th>
                            <th className="px-4 py-2.5">班级/编号</th>
                            <th className="px-4 py-2.5">原任课老师</th>
                            <th className="px-4 py-2.5 text-indigo-600">实际代课老师</th>
                            <th className="px-4 py-2.5">备注说明</th>
                            <th className="px-4 py-2.5 text-center">操作</th>
                          </tr>
                        </thead>
                        <tbody className="text-xs text-slate-700 divide-y divide-slate-100">
                          {substitutionRecords.length === 0 ? (
                            <tr>
                              <td colSpan={6} className="px-4 py-10 text-center text-slate-400">
                                暂无代课记录。如果有老师请假代课，可在左侧进行录入，系统将全自动转移扣减并重新结算课消。
                              </td>
                            </tr>
                          ) : (
                            substitutionRecords.map((record) => (
                              <tr key={record.id} className="hover:bg-slate-50/50">
                                <td className="px-4 py-2.5 font-mono">{record.dateStr}</td>
                                <td className="px-4 py-2.5">
                                  <div className="font-semibold text-slate-800">{record.className}</div>
                                  <div className="text-[10px] text-slate-400 font-mono">编号: {record.classCode}</div>
                                </td>
                                <td className="px-4 py-2.5 text-slate-600 line-through decoration-slate-400">{record.originalTeacher}</td>
                                <td className="px-4 py-2.5 font-bold text-indigo-600">{record.substituteTeacher}</td>
                                <td className="px-4 py-2.5 text-slate-500">{record.notes || '--'}</td>
                                <td className="px-4 py-2.5 text-center">
                                  <button
                                    onClick={() => handleDeleteSubstitution(record.id)}
                                    className="text-rose-600 hover:text-rose-800 p-1 cursor-pointer"
                                    title="删除此代课"
                                  >
                                    <Trash2 className="w-4 h-4 mx-auto" />
                                  </button>
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>

              </div>
            )}

            {/* TAB 3.5: MAKEUP LESSONS MANAGEMENT */}
            {activeTab === 'makeups' && (
              <div className="space-y-6 animate-in fade-in duration-200">
                
                {/* MAKEUP INTRODUCTION */}
                <div className="bg-indigo-50/40 p-4.5 rounded-xl border border-indigo-100/50 flex gap-3">
                  <div className="bg-indigo-100 text-indigo-700 p-2 rounded-lg shrink-0 self-start">
                    <CheckSquare className="w-5 h-5" />
                  </div>
                  <div className="space-y-1">
                    <h3 className="text-sm font-semibold text-slate-800">教师补课课销核算登记</h3>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      当学生进行了课外/额外补课时，管理员在此登记补课记录。
                      <br /><strong>核算规则</strong>：不需要记录学生的名字，只需在此指定老师、上课日期及补课增加的课销数量。系统将在汇总统计时<strong>直接累加</strong>至该老师名下的总结算课时与纯课时中，并根据该老师配置的基础单价全自动结算对应的课消工资。
                    </p>
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
                  
                  {/* LEFT: ADD MAKEUP RECORD FORM */}
                  <div className="lg:col-span-4 bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                    <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2 flex items-center gap-1.5">
                      <Plus className="w-4.5 h-4.5 text-indigo-600" />
                      新增补课课销记录
                    </h3>

                    <form onSubmit={handleAddMakeup} className="space-y-4">
                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">上课日期</label>
                        <input 
                          type="date" 
                          value={makeupDate}
                          onChange={(e) => setMakeupDate(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">授课教师</label>
                        <input 
                          type="text" 
                          list="teachers-datalist"
                          placeholder="选择或输入补课教师"
                          value={makeupTeacher}
                          onChange={(e) => setMakeupTeacher(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">增加课消数量 (课时数)</label>
                        <input 
                          type="number" 
                          step="0.1"
                          placeholder="如: 1.5 或 2.0"
                          value={makeupHours}
                          onChange={(e) => setMakeupHours(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1">备注说明 (可选)</label>
                        <input 
                          type="text" 
                          placeholder="如: 某某学员个性化补课"
                          value={makeupNotes}
                          onChange={(e) => setMakeupNotes(e.target.value)}
                          className="w-full text-xs border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-1 focus:ring-indigo-500"
                        />
                      </div>

                      <button
                        type="submit"
                        className="w-full py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-medium shadow transition cursor-pointer"
                      >
                        录入补课并重算
                      </button>
                    </form>
                  </div>

                  {/* RIGHT: LIST OF MAKEUP RECORDS */}
                  <div className="lg:col-span-8 bg-white border border-slate-200 rounded-xl p-5 shadow-sm space-y-4">
                    <h3 className="text-sm font-bold text-slate-800 border-b border-slate-100 pb-2 flex items-center justify-between">
                      <span>补课课销记录明细 ({makeupRecords.length} 项)</span>
                      {(startDate || endDate) && (
                        <span className="text-[10px] text-slate-400 font-normal">
                          (已自动应用时间筛选)
                        </span>
                      )}
                    </h3>

                    <div className="overflow-x-auto border border-slate-100 rounded-lg">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200 text-slate-400 text-xs font-semibold">
                            <th className="px-4 py-2.5">补课日期</th>
                            <th className="px-4 py-2.5">授课教师</th>
                            <th className="px-4 py-2.5 text-right text-indigo-600">增加课销</th>
                            <th className="px-4 py-2.5">备注说明</th>
                            <th className="px-4 py-2.5 text-center">操作</th>
                          </tr>
                        </thead>
                        <tbody className="text-xs text-slate-700 divide-y divide-slate-100">
                          {makeupRecords.length === 0 ? (
                            <tr>
                              <td colSpan={5} className="px-4 py-10 text-center text-slate-400">
                                暂无补课记录。如果有学生进行了补课，可在左侧快捷录入，该老师的月度工资核算将全自动同步累加。
                              </td>
                            </tr>
                          ) : (
                            makeupRecords.map((record) => {
                              const isFilteredOut = (startDate && record.dateStr < startDate) || (endDate && record.dateStr > endDate);
                              return (
                                <tr key={record.id} className={`hover:bg-slate-50/50 transition-colors ${isFilteredOut ? 'opacity-40 bg-slate-50/20' : ''}`}>
                                  <td className="px-4 py-2.5 font-mono">
                                    {record.dateStr}
                                    {isFilteredOut && <span className="text-[9px] text-rose-500 font-bold ml-1.5 bg-rose-50 border border-rose-100 px-1 py-0.2 rounded">超出范围已过滤</span>}
                                  </td>
                                  <td className="px-4 py-2.5 font-semibold text-slate-800">{record.teacherName}</td>
                                  <td className="px-4 py-2.5 text-right font-bold text-indigo-600 font-mono">+{record.hours.toFixed(1)} 课时</td>
                                  <td className="px-4 py-2.5 text-slate-500 truncate max-w-xs">{record.notes || '--'}</td>
                                  <td className="px-4 py-2.5 text-center">
                                    <button
                                      onClick={() => handleDeleteMakeup(record.id)}
                                      className="text-rose-600 hover:text-rose-800 p-1 cursor-pointer"
                                      title="删除此补课记录"
                                    >
                                      <Trash2 className="w-4 h-4 mx-auto" />
                                    </button>
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>

              </div>
            )}

            {/* TAB 4: PAYROLL REPORT */}
            {activeTab === 'report' && (
              <div className="space-y-6">
                
                {/* PAYROLL ACTIONS & TITLE */}
                <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 border-b border-slate-100 pb-4">
                  <div>
                    <h3 className="text-lg font-bold text-slate-900">月度课消提成与薪资汇总表</h3>
                    <p className="text-xs text-slate-500 mt-1">
                      包含教师基础单价、符合筛选时间的扣减对换及特定加成之后的最终核算薪资
                    </p>
                  </div>
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 shadow-sm">
                      <span className="text-xs font-bold text-slate-600">提成比例：</span>
                      <select
                        id="commission-rate-select"
                        value={commissionRate}
                        onChange={(e) => {
                          const val = parseFloat(e.target.value);
                          setCommissionRate(val);
                          localStorage.setItem('course_deduction_commission_rate', String(val));
                          showAlert(`提成比例已切换为 ${(val * 100).toFixed(0)}%`);
                        }}
                        className="text-xs font-bold text-indigo-600 bg-transparent outline-none border-none cursor-pointer focus:ring-0 p-0"
                      >
                        <option value={0.07}>7% (默认)</option>
                        <option value={0.06}>6%</option>
                      </select>
                    </div>

                    <button
                      id="export-csv-btn"
                      onClick={handleExportCSV}
                      className="inline-flex items-center gap-2 px-4 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-semibold shadow-md shadow-indigo-100 transition cursor-pointer"
                    >
                      <Download className="w-4 h-4" />
                      一键导出薪资结算表 (.csv)
                    </button>
                  </div>
                </div>

                {classBlocks.length === 0 ? (
                  <div className="py-12 flex flex-col items-center justify-center text-center space-y-3">
                    <div className="bg-slate-100 p-4 rounded-full text-slate-400">
                      <FileSpreadsheet className="w-8 h-8" />
                    </div>
                    <div className="space-y-1">
                      <h4 className="text-sm font-semibold text-slate-800">未检测到任何数据</h4>
                      <p className="text-xs text-slate-400">请导入 Excel 考勤课时后查看报表。</p>
                    </div>
                  </div>
                ) : (
                  <div className="space-y-6">
                    
                    {/* SUMMARY PAYROLL TABLE */}
                    <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white shadow-sm">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200 text-slate-400 font-semibold text-xs whitespace-nowrap">
                            <th className="px-4 py-3.5">老师姓名</th>
                            <th className="px-4 py-3.5 text-right">课消基础单价</th>
                            <th className="px-4 py-3.5 text-center">授课总次数</th>
                            <th className="px-4 py-3.5 text-right">名下班级课时<span className="text-[10px] text-slate-400 block font-normal">(乘人数)</span></th>
                            <th className="px-4 py-3.5 text-right text-indigo-500">名下班级纯课时<span className="text-[10px] text-indigo-400 block font-normal">(不计人数)</span></th>
                            <th className="px-4 py-3.5 text-right text-rose-600">代出课时(-)<span className="text-[10px] text-rose-400 block font-normal">(乘人数)</span></th>
                            <th className="px-4 py-3.5 text-right text-emerald-600">代入课时(+)<span className="text-[10px] text-emerald-400 block font-normal">(乘人数)</span></th>
                            <th className="px-4 py-3.5 text-right text-indigo-500">补课课销(+)<span className="text-[10px] text-indigo-400 block font-normal">(不计人数)</span></th>
                            <th className="px-4 py-3.5 text-right font-bold text-slate-900 bg-slate-50/50">总结算课时<span className="text-[10px] text-slate-500 block font-normal">(乘人数)</span></th>
                            <th className="px-4 py-3.5 text-right font-bold text-indigo-600 bg-indigo-50/30">总结算纯课时<span className="text-[10px] text-indigo-500 block font-normal">(不计人数)</span></th>
                            <th className="px-4 py-3.5 text-right text-indigo-600">加成课时</th>
                            <th className="px-4 py-3.5 text-right text-indigo-600">加成提成累计<span className="text-[10px] text-indigo-400 block font-normal">(不计人数)</span></th>
                            <th className="px-4 py-3.5 text-right font-bold text-slate-900 bg-slate-50/55">总课销金额<span className="text-[10px] text-slate-500 block font-normal">(应结课销)</span></th>
                            <th className="px-4 py-3.5 text-right font-bold text-indigo-600 bg-indigo-50/20">课销提成<span className="text-[10px] text-indigo-500 block font-normal">(切换比例)</span></th>
                            <th className="px-4 py-3.5 text-right font-bold text-slate-950 bg-indigo-50/10">老师实际到手薪资</th>
                            <th className="px-4 py-3.5 text-center">代/补课明细</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-sm text-slate-700">
                          {teacherReportData.map((t) => {
                            return (
                              <tr key={t.teacherName} className="hover:bg-slate-50/60 transition-colors">
                                <td className="px-4 py-3.5 font-bold text-slate-800">{t.teacherName}</td>
                                <td className="px-4 py-3.5 text-right font-mono">¥ {t.baseRate} <span className="text-[10px] text-slate-400">/课时</span></td>
                                <td className="px-4 py-3.5 text-center font-mono font-medium">{t.sessionsCount} <span className="text-xs text-slate-400">次</span></td>
                                <td className="px-4 py-3.5 text-right font-mono">{t.baseHours.toFixed(1)}</td>
                                <td className="px-4 py-3.5 text-right font-mono text-indigo-500">{t.baseClassHours.toFixed(1)}</td>
                                <td className="px-4 py-3.5 text-right font-mono text-rose-600">
                                  {t.substitutedOutHours > 0 ? `-${t.substitutedOutHours.toFixed(1)}` : '0.0'}
                                </td>
                                <td className="px-4 py-3.5 text-right font-mono text-emerald-600">
                                  {t.substitutedInHours > 0 ? `+${t.substitutedInHours.toFixed(1)}` : '0.0'}
                                </td>
                                <td className="px-4 py-3.5 text-right font-mono text-indigo-600">
                                  {t.makeupHours > 0 ? `+${t.makeupHours.toFixed(1)}` : '0.0'}
                                </td>
                                <td className="px-4 py-3.5 text-right font-bold font-mono text-slate-900 bg-slate-50/50">{t.settlementHours.toFixed(1)}</td>
                                <td className="px-4 py-3.5 text-right font-bold font-mono text-indigo-600 bg-indigo-50/30">{t.settlementClassHours.toFixed(1)}</td>
                                <td className="px-4 py-3.5 text-right font-mono text-indigo-600">{t.bonusHours.toFixed(1)}</td>
                                <td className="px-4 py-3.5 text-right font-bold font-mono text-indigo-600">¥ {t.bonusAmount.toFixed(1)}</td>
                                <td className="px-4 py-3.5 text-right font-bold font-mono text-slate-900 bg-slate-50/55">
                                  ¥ {t.baseSalary.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}
                                </td>
                                <td className="px-4 py-3.5 text-right font-bold font-mono text-indigo-600 bg-indigo-50/20">
                                  <div>¥ {t.commissionAmount.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</div>
                                  <div className="text-[10px] text-indigo-500 font-normal mt-0.5">
                                    <select
                                      value={(() => {
                                        const baseRateObj = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === t.teacherName.trim().toLowerCase());
                                        return (baseRateObj && baseRateObj.commissionRate !== undefined) ? baseRateObj.commissionRate : commissionRate;
                                      })()}
                                      onChange={(e) => {
                                        const val = parseFloat(e.target.value);
                                        const baseRateObj = teacherBaseRates.find(r => r.teacherName.trim().toLowerCase() === t.teacherName.trim().toLowerCase());
                                        if (baseRateObj) {
                                          updateTeacherCommissionRate(baseRateObj.teacherName, val);
                                        } else {
                                          // Create base rate object if not exists
                                          const updated = [...teacherBaseRates, { teacherName: t.teacherName, baseRate: 100, commissionRate: val }];
                                          setTeacherBaseRates(updated);
                                          saveToLocalStorage('course_deduction_base_rates', updated);
                                          showAlert(`已为 ${t.teacherName} 创建专属配置，并设置提成比例为 ${(val * 100).toFixed(0)}%！`);
                                        }
                                      }}
                                      className="text-[10px] text-indigo-500 bg-transparent border-none cursor-pointer focus:ring-0 p-0 outline-none text-right font-semibold"
                                    >
                                      <option value={0.07}>7% (默认)</option>
                                      <option value={0.06}>6%</option>
                                      <option value={0.05}>5%</option>
                                      <option value={0.08}>8%</option>
                                      <option value={0.09}>9%</option>
                                      <option value={0.10}>10%</option>
                                    </select>
                                  </div>
                                </td>
                                <td className="px-4 py-3.5 text-right font-bold font-mono text-slate-950 bg-indigo-50/10">
                                  ¥ {t.totalSalary.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}
                                </td>
                                <td className="px-4 py-3.5 text-center">
                                  {t.substitutionDetails.length > 0 || t.makeupDetails.length > 0 ? (
                                    <div className="group relative inline-block">
                                      <span className="bg-amber-50 text-amber-700 border border-amber-200 rounded px-1.5 py-0.5 text-[10px] font-medium cursor-help">
                                        {t.substitutionDetails.length + t.makeupDetails.length} 条记录
                                      </span>
                                      <div className="hidden group-hover:block absolute right-0 bottom-full mb-2 w-72 bg-slate-900 text-white text-[10px] p-3 rounded-lg shadow-xl z-50 leading-relaxed text-left space-y-1.5">
                                        {t.substitutionDetails.length > 0 && (
                                          <>
                                            <div className="font-bold border-b border-slate-700 pb-1 text-amber-400">代课置换明细：</div>
                                            {t.substitutionDetails.map((det, di) => (
                                              <div key={di} className="truncate">{det}</div>
                                            ))}
                                          </>
                                        )}
                                        {t.makeupDetails.length > 0 && (
                                          <>
                                            <div className="font-bold border-b border-slate-700 pb-1 text-emerald-400 pt-1.5">补课课销明细：</div>
                                            {t.makeupDetails.map((det, di) => (
                                              <div key={di} className="truncate">{det}</div>
                                            ))}
                                          </>
                                        )}
                                      </div>
                                    </div>
                                  ) : (
                                    <span className="text-slate-300 text-xs">-</span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    {/* REPORT NOTES */}
                    <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-2">
                      <h4 className="text-xs font-bold text-slate-700 flex items-center gap-1">
                        <Info className="w-3.5 h-3.5 text-indigo-500" />
                        工资算法解析说明
                      </h4>
                      <ul className="list-disc pl-4 text-xs text-slate-500 space-y-1.5">
                        <li><strong>总结算课时</strong> = 名下班级总课时 - 代出课时 + 代入课时 + 录入的补课课时。</li>
                        <li><strong>总课销金额 (应结课销)</strong> = 总结算课时 × 该教师个性化配置的课消基础单价。</li>
                        <li><strong>课销提成 (提成所得)</strong> = 总课销金额 × 选择的提成比例 (可选择 6% 或 7%)，这是教师实际所得的课消提成。</li>
                        <li><strong>加成提成累计</strong> = 满足特定“课时费加成”的班级<strong>纯课时数 (无论该班级学生人数是多少人，单次课时的加成固定不乘学生人数)</strong> × 对应的加成额度。</li>
                        <li><strong>老师实际到手薪资</strong> = 课销提成 + 加成提成累计。</li>
                      </ul>
                    </div>

                  </div>
                )}
              </div>
            )}

          </div>
        </div>

      </main>

      {/* FOOTER */}
      <footer className="bg-white border-t border-slate-200 mt-12 py-6 text-center text-xs text-slate-400">
        <div className="max-w-7xl mx-auto px-4">
          <p>© 2026 课消管家 - 培训机构课消提成统计系统. All rights reserved.</p>
          <p className="mt-1">设计精美，专注于机构核算的高精准结算管家</p>
        </div>
      </footer>
    </div>
  );
}
