from pathlib import Path

src = Path(r"C:\Users\HZJ\Desktop\财务管理实操\00-知识库与学习导航")
out = Path(r"C:\Users\HZJ\Desktop\Git\work\month_close_20260909\nav")
out.mkdir(parents=True, exist_ok=True)

readme = (src / "README.md").read_text(encoding="utf-8")
entry = "- [M02-001 总账会计年度工作与月结](../02-总账与月结季结年结/01-总账月结全景/00-阅读说明.md)：年度工作节奏、D-5至D+10月结日历、采购暂估、成本、往来、账税、报表勾稽、制造案例和面试表达。\n"
if entry not in readme:
    readme = readme.replace("- [M12-001 会计职业风险]", entry + "- [M12-001 会计职业风险]")
old = "**M12-001：会计职业风险。** 资料已建立，包含XMind导图、可筛选Excel、文档式手册和兼容导入文件；当前状态是待学习与练习验证，不代表已经掌握。该主题把职业底线与入职交接、月结复核、税务、资金、成本、系统权限和AI数据风险相连。"
new = "**M02-001：总账会计如何组织并完成一次月结。** 资料已建立，包含年度工作节奏、月结全景图、17天关账日历、五个重点模块、制造企业案例、面试表达和5页签执行清单；当前状态为资料已建立、待模拟练习。\n\n" + old
readme = readme.replace(old, new)
start = "## 下一学习主题\n\n**M02-001：总账会计如何组织并完成一次月结。**\n\n拟覆盖：关账前准备 → 业务截止 → 子模块核对 → 调整分录 → 全科目复核 → 报表编制与勾稽 → 复核审批 → 锁账及更正流程。\n\n拟交付：一份示例关账日历、一份总账检查清单，以及“月结发现差异如何追踪与沟通”的练习。具体日程和模块顺序须结合企业组织、ERP配置及报表要求确定。"
replacement = "## 下一学习主题\n\n**M02-002：采购暂估与发票差异。**\n\n拟覆盖：材料和服务暂估、暂估价格、次月冲回、发票三单匹配、差额分配、已耗用后收票、长期暂估清理、进项税额和企业所得税扣除凭证。"
readme = readme.replace(start, replacement)
(out / "README.md").write_text(readme, encoding="utf-8")

ledger = (src / "02-学习与问题台账.md").read_text(encoding="utf-8")
oldrow = "| M02-001 | 总账会计如何组织并完成一次月结 | M02 | M01、M03、M04—M10、M12 | 待学习 | 示例关账日历、总账检查清单、差异处理练习 | 结合具体问题开展第一轮讲解 |"
newrow = "| M02-001 | 总账会计如何组织并完成一次月结 | M02 | M01、M03、M04—M12 | 资料已建立，待模拟练习 | Word实战手册、年度日历、D-5至D+10清单、科目勾稽和问题台账 | 依次深化采购暂估、成本结转、往来清理、账税核对、报表勾稽 |"
ledger = ledger.replace(oldrow, newrow)
new_history = "| 2026-09-09 | 建立M02-001总账月结专题 | 17页Word手册、月结全景图、12个月年度节奏、D-5至D+10关账日历、5页签Excel、制造企业案例、面试表达 | D+7为示例企业口径；具体企业应按报表时限、ERP配置和主管税务机关核定期限调整 |\n"
marker = "| 2026-09-08 | 建立顾问工作约定与知识体系"
if new_history not in ledger:
    ledger = ledger.replace(marker, new_history + marker)
(out / "02-学习与问题台账.md").write_text(ledger, encoding="utf-8")
print(out)
