# 工作记录仓库

这个仓库用来备份电脑上 `C:\Users\HZJ\Desktop\Git` 文件夹里的工作内容，并记录对电脑做过的每一次操作。

**这是私有仓库。**

---

## 目录里有什么

| 文件夹 | 装什么 |
|---|---|
| `操作记录\` | 对电脑做过的每一次操作：改了什么、怎么撤销。出事时照这个退回去 |
| `work\` | 日常工作产出：财务实操手册、工具包、脚本、说明文档 |
| `WorkBuddyVersions\` | WorkBuddy 的版本记录 |

### `work\` 下按主题分

- **应付 / 应收**：`ap_review_*`、`ap_supplier_management_*`、`ar_review_*`、`ar_management_*`
- **账务与报表**：`gl_reconciliation_manual_*`、`month_close_*`、`financial_statements_*`、`erp_finance_*`
- **税务与费用**：`invoice_tax_manual_*`、`expense_invoice_cases_*`
- **资产与存货**：`fixed_asset_management_*`、`inventory_management_*`、`manufacturing_cost_inventory_*`
- **分析与风控**：`fpna-output\`、`finance_risk_*`、`finance_supervisor_*`、`audit-adjustments-output\`
- **流程与自动化**：`ai_finance_automation_*`、`full_cycle_manual_*`
- **人员相关**：`interview_onboarding_*`、`finance_records_handover_*`
- **手册排版**：`manual_pdf\`、`manual_pdf_5_7\`、`manual_pdf_8_11\`、`word_pdf\`

---

## 哪些东西**没有**放进仓库

这个文件夹本地有约 **2.8 GB**，仓库里只留了约 **44 MB**。被排除的都不是成果，而是「能重新下载」或「能重新生成」的东西：

| 排除的内容 | 为什么 |
|---|---|
| 音视频素材（847 MB） | 原始素材，体积大 |
| 文档渲染出的 PNG 截图（587 MB） | 文档里已经含图，用原文档可重新导出 |
| 语音识别包、模型（1 GB 以上） | 官方渠道能重新下载 |
| `node_modules`（12 万个文件） | 装个依赖就能重新生成 |
| Python 虚拟环境、`__pycache__` | 同上 |
| `render_*` / `lo_profile_*` 渲染中间产物 | 中间步骤，不是最终结果 |
| `.xlsx` 检查记录（`.inspect.ndjson`） | 能由原表格重新生成 |
| `video_extract_temp\`、`Claude code（DeepSeek）\` | 临时文件 |

**完整规则写在仓库根目录的 `.gitignore` 里**，每条都带中文注释。

> 注意：被排除的文件**仍然在本地电脑上**，只是不进 git。删除 `.gitignore` 里对应的那一行，它们就会重新被纳入。

---

## 出事了怎么退回去

在对话里说：「去 `操作记录\` 里找 `<任务名>`，退回到第 N 步」。

每个任务的记录里有三样东西：
- **开工前的状态** —— 动手前电脑本来是什么样
- **步骤记录** —— 每步做了什么、**怎么撤销**
- **`备份\`** —— 改动前的原始文件

### 一条必须记住的事

**git 只能回到「记录文件」的旧版本，不能帮你还原电脑。**

还原电脑靠的是 `备份\` 里的原始文件，加上记录里「怎么撤销」那一列。哪一步没写撤销方法、没做备份，那一步就退不回去。

---

## 关于敏感信息

仓库是私有的，但仍然：**账号、密码、密钥不要写进来。** 写清「改了哪一项」就够了。
