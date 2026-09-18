#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WorkBuddy 运行版本快照工具
=========================

抓取 WorkBuddy 桌面端的版本信息 + 运行环境 + 运行痕迹，生成一份带时间戳的
JSON 快照，并追加到人类可读的 TIMELINE.md。

设计原则（遵守用户的电脑协作红线）：
  * 只读，不修改 WorkBuddy 的任何文件
  * 不重复处理：如果版本指纹和上次完全一致，则不生成新记录
  * 无网络、无循环、无高频重试
  * 幂等：同一天同一版本重复运行不会污染历史

用法：
    python snapshot.py            # 快速模式（推荐，秒级）
    python snapshot.py --deep     # 深度模式（额外计算主程序 SHA256，约 10-30 秒）
    python snapshot.py --force    # 强制记录，即使指纹未变化
"""

import argparse
import datetime as _dt
import hashlib
import json
import os
import platform
import subprocess
import sys

# ---------------------------------------------------------------- 路径常量

WORKBUDDY_INSTALL = r"D:\04 Software\WorkBuddy 腾讯旗下"
WORKBUDDY_USER_DIR = r"C:\Users\HZJ\.workbuddy"
WORKBUDDY_SESSIONS = r"C:\Users\HZJ\WorkBuddy"

MANIFEST = os.path.join(WORKBUDDY_INSTALL, "resources", "install-manifest.json")
RENDERER_VERSION = os.path.join(WORKBUDDY_USER_DIR, "app", "renderer-version.json")
APP_CONFIG = os.path.join(WORKBUDDY_USER_DIR, "app", "app-config.json")
MAIN_EXE = os.path.join(WORKBUDDY_INSTALL, "WorkBuddy.exe")
CLI_PRODUCT = os.path.join(
    WORKBUDDY_INSTALL, "resources", "app.asar.unpacked", "cli", "product.json"
)
AUDIT_LOG_DIR = os.path.join(WORKBUDDY_USER_DIR, "audit-log")

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(TOOLS_DIR)          # 仓库根 = tools 的上一级
RECORDS_DIR = os.path.join(BASE_DIR, "records")
TIMELINE = os.path.join(BASE_DIR, "TIMELINE.md")


# ---------------------------------------------------------------- 工具函数

def _setup_console():
    """让控制台输出与 run-snapshot.bat 里的 chcp 65001 对齐。

    不加这层，Windows 控制台按 cp936 解码 Python 的 UTF-8 输出会乱码。
    这里强制 UTF-8，与 bat 保持一致。
    """
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


def _now():
    return _dt.datetime.now().astimezone()


def _read_json(path):
    """安全读取 JSON 文件，失败返回 None。"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as exc:  # noqa: BLE001
        return {"__error__": str(exc)}


def _file_meta(path):
    """文件大小 + 修改时间（不读内容，秒级）。"""
    try:
        st = os.stat(path)
        return {
            "exists": True,
            "size_bytes": st.st_size,
            "size_mb": round(st.st_size / 1024 / 1024, 2),
            "mtime": _dt.datetime.fromtimestamp(st.st_mtime).astimezone().isoformat(),
        }
    except Exception:  # noqa: BLE001
        return {"exists": False}


def _sha256(path, chunk=1024 * 1024):
    """流式计算 SHA256，避免一次性读入 200MB 文件。"""
    h = hashlib.sha256()
    try:
        with open(path, "rb") as f:
            while True:
                block = f.read(chunk)
                if not block:
                    break
                h.update(block)
        return h.hexdigest()
    except Exception as exc:  # noqa: BLE001
        return "error: %s" % exc


def _dir_listing(path, limit=500):
    """列出目录下的条目名（用于会话目录统计）。"""
    try:
        names = sorted(os.listdir(path))
        return names[:limit], len(names)
    except Exception:  # noqa: BLE001
        return [], 0


# ---------------------------------------------------------------- 采集模块

def collect_app_version():
    """WorkBuddy 应用版本——最核心的字段。"""
    manifest = _read_json(MANIFEST) or {}
    renderer = _read_json(RENDERER_VERSION) or {}
    return {
        "app_version": manifest.get("appVersion") or renderer.get("version"),
        "manifest_version_scheme": manifest.get("version"),
        "install_generated_at": manifest.get("generatedAt"),
        "installed_file_count": manifest.get("fileCount"),
        "renderer_version": renderer.get("version"),
    }


def collect_binary(deep=False):
    """主程序二进制指纹。"""
    meta = _file_meta(MAIN_EXE)
    if deep and meta.get("exists"):
        meta["sha256"] = _sha256(MAIN_EXE)
    return meta


def collect_cli():
    """CLI 产品信息（若存在）。"""
    data = _read_json(CLI_PRODUCT)
    if isinstance(data, dict) and "__error__" not in data:
        # 只保留小而关键的字段，避免把整包塞进记录
        keep = {k: v for k, v in data.items() if isinstance(v, (str, int, float, bool))}
        return keep
    return {"exists": os.path.exists(CLI_PRODUCT)}


def collect_runtime():
    """当前运行环境（执行快照时所在的环境）。"""
    info = {
        "os": platform.system(),
        "os_release": platform.release(),
        "os_version": platform.version(),
        "machine": platform.machine(),
        "processor": platform.processor(),
        "python_version": sys.version.split()[0],
        "snapshot_time": _now().isoformat(),
        "snapshot_host_user": os.environ.get("USERNAME") or os.environ.get("USER"),
    }
    # Node 版本（如果 PATH 里有）
    try:
        node = subprocess.run(
            ["node", "--version"], capture_output=True, text=True, timeout=5
        )
        if node.returncode == 0:
            info["node_version"] = node.stdout.strip()
    except Exception:  # noqa: BLE001
        pass
    return info


def collect_config():
    """应用配置快照（不含任何凭据）。"""
    cfg = _read_json(APP_CONFIG)
    if isinstance(cfg, dict) and "__error__" not in cfg:
        return cfg
    return {"__note__": "config unreadable", "raw": cfg}


def collect_sessions():
    """WorkBuddy 会话工作区列表。"""
    names, total = _dir_listing(WORKBUDDY_SESSIONS)
    return {"count": total, "sessions": names}


def collect_audit_summary():
    """审计日志分片统计（只统计，不复制内容）。"""
    names, total = _dir_listing(AUDIT_LOG_DIR)
    segments = []
    for name in names:
        if not name.endswith(".jsonl"):
            continue
        full = os.path.join(AUDIT_LOG_DIR, name)
        meta = _file_meta(full)
        line_count = None
        try:
            with open(full, "r", encoding="utf-8", errors="replace") as f:
                line_count = sum(1 for _ in f)
        except Exception:  # noqa: BLE001
            pass
        segments.append(
            {
                "segment": name,
                "size_bytes": meta.get("size_bytes"),
                "entries": line_count,
            }
        )
    return {"segment_count": total, "segments": segments}


# ---------------------------------------------------------------- 指纹与写入

def build_fingerprint(record):
    """版本指纹：只取决定"版本是否变化"的字段，与时间无关。

    刻意不纳入 binary.sha256 和文件 mtime —— 前者只有 --deep 模式才有，
    纳入会导致深度模式与快速模式之间出现假变化；后者只是元数据，
    文件被 touch 但内容没变不应算版本变化。
    """
    fp_src = {
        "app_version": record["app_version"]["app_version"],
        "installed_file_count": record["app_version"]["installed_file_count"],
        "install_generated_at": record["app_version"]["install_generated_at"],
        "binary_size": record["binary"].get("size_bytes"),
        "config": record["config"],
    }
    blob = json.dumps(fp_src, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]


def last_record():
    """读取最近一份记录（用于去重比对）。"""
    if not os.path.isdir(RECORDS_DIR):
        return None
    files = sorted(f for f in os.listdir(RECORDS_DIR) if f.endswith(".json"))
    if not files:
        return None
    return _read_json(os.path.join(RECORDS_DIR, files[-1]))


def write_record(record):
    """写入 records/ 目录，文件名带时间戳 + 版本号。"""
    os.makedirs(RECORDS_DIR, exist_ok=True)
    ts = _now().strftime("%Y-%m-%d_%H%M%S")
    ver = (record["app_version"].get("app_version") or "unknown").replace("/", "_")
    fname = "%s_v%s.json" % (ts, ver)
    path = os.path.join(RECORDS_DIR, fname)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)
    return path


def append_timeline(record, note=""):
    """把这次快照追加到人类可读的 TIMELINE.md。"""
    rows = []
    if os.path.exists(TIMELINE):
        with open(TIMELINE, "r", encoding="utf-8") as f:
            rows = f.read().rstrip("\n").split("\n")

    header = [
        "# WorkBuddy 运行版本时间线",
        "",
        "> 本文件由 `tools/snapshot.py` 自动追加，请勿手工改写数据行。",
        "> 每行 = 一次版本快照。指纹未变化时不会新增行。",
        "",
        "| 快照时间 | 应用版本 | 安装文件数 | 主程序大小 | 会话数 | 指纹 | 备注 |",
        "|---|---|---|---|---|---|---|",
    ]
    if not rows:
        rows = header
    elif rows[-1].strip() == "":
        rows = rows[:-1]

    av = record["app_version"]
    line = "| %s | %s | %s | %s MB | %s | `%s` | %s |" % (
        record["snapshot_time"].replace("T", " ")[:19],
        av.get("app_version"),
        av.get("installed_file_count"),
        record["binary"].get("size_mb"),
        record["sessions"]["count"],
        record["fingerprint"],
        note,
    )
    rows.append(line)

    with open(TIMELINE, "w", encoding="utf-8") as f:
        f.write("\n".join(rows) + "\n")


# ---------------------------------------------------------------- 主流程

def main():
    _setup_console()
    ap = argparse.ArgumentParser(description="WorkBuddy 运行版本快照")
    ap.add_argument("--deep", action="store_true", help="额外计算主程序 SHA256")
    ap.add_argument("--force", action="store_true", help="指纹未变也强制记录")
    args = ap.parse_args()

    record = {
        "schema": 1,
        "snapshot_time": _now().isoformat(),
        "app_version": collect_app_version(),
        "binary": collect_binary(deep=args.deep),
        "cli": collect_cli(),
        "runtime": collect_runtime(),
        "config": collect_config(),
        "sessions": collect_sessions(),
        "audit_log": collect_audit_summary(),
    }
    record["fingerprint"] = build_fingerprint(record)

    prev = last_record()
    if prev and prev.get("fingerprint") == record["fingerprint"] and not args.force:
        print("[=] 版本指纹未变化 (%s)，未新增记录。" % record["fingerprint"])
        print("    应用版本: %s" % record["app_version"].get("app_version"))
        return 0

    path = write_record(record)
    note = "首次记录" if not prev else "版本/配置变化"
    append_timeline(record, note=note)

    print("[+] 已记录快照")
    print("    应用版本 : %s" % record["app_version"].get("app_version"))
    print("    安装文件 : %s 个" % record["app_version"].get("installed_file_count"))
    print("    主程序   : %s MB%s" % (
        record["binary"].get("size_mb"),
        " (含 SHA256)" if args.deep else "",
    ))
    print("    会话数   : %s" % record["sessions"]["count"])
    print("    指纹     : %s" % record["fingerprint"])
    print("    文件     : %s" % path)
    print("    时间线   : %s" % TIMELINE)
    return 0


if __name__ == "__main__":
    sys.exit(main())
