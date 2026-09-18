# WorkBuddyVersions —— WorkBuddy 桌面端运行版本记录

独立仓库，用来记录 **WorkBuddy 桌面端本身**的版本与运行环境快照。

> 与 Codex 无关。本目录自带独立 `.git`，不参与、不影响外层 `Desktop\Git` 仓库的任何内容。

---

## 这个目录解决什么问题

和记录 Codex 版本的做法一样，本目录给 WorkBuddy 建立一个**可追溯的版本时间线**：

- WorkBuddy 什么时候升级到哪个版本
- 升级前后安装文件数量、主程序大小、配置有没有变
- 每次升级时开着多少个会话
- 版本指纹（指纹未变 = 版本没动，不重复记录）

## 目录结构

```
WorkBuddyVersions\
├── .git\                   独立仓库（不碰外层 Codex 仓库）
├── README.md               本文件
├── TIMELINE.md             人类可读时间线（自动追加）
├── records\                每次快照一份 JSON
│   └── 2026-09-14_204500_v5.5.6.json
└── tools\
    ├── snapshot.py         采集脚本（只读，不改 WorkBuddy）
    └── run-snapshot.bat    双击即抓（Windows）
```

## 怎么用

**双击运行**（最简单）：

```
tools\run-snapshot.bat
```

**命令行**：

```bat
python tools\snapshot.py           :: 快速模式，秒级
python tools\snapshot.py --deep    :: 深度模式，额外算主程序 SHA256
python tools\snapshot.py --force   :: 指纹没变也强制记录
```

## 采集了哪些数据

| 字段 | 来源 | 说明 |
|---|---|---|
| `app_version` | `install-manifest.json` / `renderer-version.json` | 应用版本号、安装文件数、安装生成时间 |
| `binary` | `WorkBuddy.exe` | 大小、修改时间（`--deep` 时含 SHA256） |
| `cli` | `cli/product.json` | CLI 端产品信息 |
| `runtime` | 执行快照时的机器环境 | 系统、Python/Node 版本 |
| `config` | `app-config.json` | 语气、语言等配置 |
| `sessions` | `~/WorkBuddy/*` | 会话工作区列表与数量 |
| `audit_log` | `~/.workbuddy/audit-log` | 审计日志分片统计（只统计条目数，不复制内容） |

**不含任何凭据**：不读密码、Token、Cookie、聊天内容。

## 安全边界

- 全部操作**只读**，不修改 WorkBuddy 安装目录和用户目录里的任何文件
- 无网络请求
- 无循环、无定时轮询（除非用户显式创建定时任务）
- 指纹去重：版本没变就不生成新记录，避免历史被噪音污染

## 提交习惯

每次版本变化后提交一次，commit message 建议格式：

```
snapshot: WorkBuddy v5.5.6 @ 2026-09-14
```

这样 `git log` 就是一条清晰的版本升级时间线。
