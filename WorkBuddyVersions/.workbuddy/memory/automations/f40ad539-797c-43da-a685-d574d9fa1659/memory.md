# 自动化执行记录：WorkBuddy 版本快照留档

## 固定信息
- 工作目录：`C:\Users\HZJ\Desktop\Git\WorkBuddyVersions`（独立 git 仓库，branch main）
- 入口脚本：`tools\snapshot.py`（只读采集，不修改 WorkBuddy 安装目录与用户目录）
- Python：`%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe`
- 分支约定：指纹无变化 → 跳过，不 commit；有变化 → `git add .` + `git commit -m "snapshot: WorkBuddy version record"`

## 执行历史

| 日期 | 应用版本 | 指纹 | 结果 |
| --- | --- | --- | --- |
| 2026-09-15 | 5.5.6 | 4041382d97a5e203 | 指纹未变化，未新增记录，工作树干净，无需 commit |

## 备注
- 首次执行本自动化（此前 `records\` 内已有 2026-09-14 的 v5.5.6 记录，由手动/首次抓取产生）。
- 已确认外层 `C:\Users\HZJ\Desktop\Git` 仓库未被触碰。
