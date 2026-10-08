import { Download, FileSpreadsheet, Info } from 'lucide-react';
import type { PayrollRow } from '../utils/payroll';

interface Props {
  teacherReportData: PayrollRow[];
  commissionRate: number;
  onCommissionRateChange: (rate: number) => void;
  onTeacherCommissionRateChange: (name: string, rate: number) => void;
  onExport: () => void;
}
export function PayrollReport({ teacherReportData, commissionRate, onCommissionRateChange, onTeacherCommissionRateChange, onExport }: Props) {
  return (
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
              onChange={(e) => onCommissionRateChange(Number(e.target.value))}
              className="text-xs font-bold text-indigo-600 bg-transparent outline-none border-none cursor-pointer focus:ring-0 p-0"
            >
              {![0.06, 0.07].includes(commissionRate) && <option value={commissionRate}>{(commissionRate * 100).toFixed(1)}%</option>}
              <option value={0.07}>7% (默认)</option>
              <option value={0.06}>6%</option>
            </select>
          </div>

          <button
            id="export-csv-btn"
            onClick={onExport}
            className="inline-flex items-center gap-2 px-4 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-semibold shadow-md shadow-indigo-100 transition cursor-pointer"
          >
            <Download className="w-4 h-4" />
            一键导出薪资结算表 (.csv)
          </button>
        </div>
      </div>

      {teacherReportData.length === 0 ? (
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
                            value={t.commissionRate}
                            onChange={(e) => onTeacherCommissionRateChange(t.teacherName, Number(e.target.value))}
                            className="text-[10px] text-indigo-500 bg-transparent border-none cursor-pointer focus:ring-0 p-0 outline-none text-right font-semibold"
                          >
                            {![0.05, 0.06, 0.07, 0.08, 0.09, 0.10].includes(t.commissionRate) && <option value={t.commissionRate}>{(t.commissionRate * 100).toFixed(1)}%</option>}
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
  );
}
