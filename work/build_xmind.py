from __future__ import annotations

import base64
import hashlib
import json
import sys
import zipfile
from pathlib import Path


SOURCE = Path(r"C:\Users\HZJ\Desktop\Git\work\齐昊老师_学员实战案例_第1集_思维导图源.md")
OUTPUT = Path(r"C:\Users\HZJ\Desktop\Git\work\齐昊老师_学员实战案例_第1集_思维导图.xmind")


def stable_id(path: str) -> str:
    return hashlib.sha1(path.encode("utf-8")).hexdigest()[:26]


def topic(title: str, path: str, children: list[dict] | None = None, *, root: bool = False) -> dict:
    node = {
        "id": stable_id(path),
        "class": "topic",
        "title": title,
    }
    if root:
        node["structureClass"] = "org.xmind.ui.map.unbalanced"
    if children:
        node["children"] = {"attached": children}
    return node


def parse_markdown(source: Path) -> tuple[str, list[dict]]:
    root_title = ""
    main_topics: list[dict] = []
    current_main: dict | None = None
    current_sub: dict | None = None

    for raw in source.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("# "):
            root_title = line[2:].strip()
        elif line.startswith("## "):
            title = line[3:].strip()
            current_main = topic(title, f"detail/{title}")
            current_main["children"] = {"attached": []}
            main_topics.append(current_main)
            current_sub = None
        elif line.startswith("### ") and current_main is not None:
            title = line[4:].strip()
            current_sub = topic(title, f"detail/{current_main['title']}/{title}")
            current_sub["children"] = {"attached": []}
            current_main["children"]["attached"].append(current_sub)
        elif line.startswith("- ") and current_main is not None:
            title = line[2:].strip()
            parent = current_sub if current_sub is not None else current_main
            child_path = f"detail/{current_main['title']}/{parent['title']}/{title}"
            parent["children"]["attached"].append(topic(title, child_path))

    if not root_title or not main_topics:
        raise ValueError("Markdown source does not contain the expected heading hierarchy")
    return root_title, main_topics


def build_overview() -> tuple[str, list[dict]]:
    overview = {
        "1. 背景与核心问题": [
            ("企业状态", ["啤酒经销商持续亏损", "财务数据混乱", "财务加班仍频繁出错", "老板无法据此制定奖金与决策"]),
            ("底层判断", ["不是单纯的账务问题", "系统、流程、责任、资产、信用与激励共同失灵"]),
        ],
        "2. 统一系统与责任归位": [
            ("一个数据源", ["停用6套割裂系统", "统一手机APP、客户与商品口径"]),
            ("源头录入", ["业务员直接下单并选择司机", "配送轨迹与异常停留自动记录", "谁发生业务，谁负责录入"]),
            ("财务转型", ["不再替业务重复录入", "转向审核、分析、控制与经营支持"]),
        ],
        "3. 推动变革落地": [
            ("识别阻力", ["系统透明化会触碰既得利益", "表面的‘太忙’不一定是真因"]),
            ("绩效强绑定", ["录入率低于70%：相关绩效归零", "超过90%的前五名获得奖励", "先培养习惯，再增加经营指标"]),
            ("组织保障", ["重大改革由老板拍板", "财务负责人获得跨部门授权"]),
        ],
        "4. 数据与资产治理": [
            ("上线前治理", ["连续两个月精准盘点", "SKU由约2000个精简至635个", "重建仓储和物流制度"]),
            ("包装物分项核算", ["酒水、瓶子、瓶盖、箱子分别核算", "回瓶率目标约70%—80%", "赠酒仍收包装物押金"]),
        ],
        "5. 应收与客户经营": [
            ("应收三维控制", ["账期上限", "信用额度上限", "未结清最多再开3单，超限锁单"]),
            ("闭环", ["交易前授信", "交易中预警", "超限锁单", "回款后解锁"]),
            ("客户绑定", ["3个月完成8万元销量", "每箱奖2元，提前完成奖3元", "奖励折成下批赠酒且继续收押金"]),
        ],
        "6. 绩效重构与经营结果": [
            ("绩效结构", ["基本工资＋绩效工资＋销售提成", "阶段奖励＋扣除项目"]),
            ("奖励维度", ["回款、新客户、拜访频率", "高毛利产品、归口专销、利润分红"]),
            ("风险扣除", ["坏账、客户流失、风险金", "业务合规、异常订单、资产流失"]),
            ("结果", ["由亏损转为省级龙头经销商", "视频标题概括：业绩增长约5倍"]),
        ],
        "7. 财务负责人能力模型": [
            ("六项核心能力", ["业务理解", "系统与数据治理", "流程与内控设计", "管理会计", "变革推动", "经营结果意识"]),
            ("总体因果链", ["数据统一 → 流程重构 → 责任归位", "核算精细 → 风险受控 → 绩效驱动 → 经营改善"]),
        ],
    }

    branches: list[dict] = []
    for main_title, sections in overview.items():
        section_nodes = []
        for section_title, leaves in sections:
            leaf_nodes = [topic(leaf, f"overview/{main_title}/{section_title}/{leaf}") for leaf in leaves]
            section_nodes.append(topic(section_title, f"overview/{main_title}/{section_title}", leaf_nodes))
        branches.append(topic(main_title, f"overview/{main_title}", section_nodes))
    return "第1集｜亏损企业如何业绩增长5倍", branches


def sheet(title: str, root_title: str, children: list[dict], key: str) -> dict:
    return {
        "id": stable_id(f"sheet/{key}"),
        "class": "sheet",
        "title": title,
        "rootTopic": topic(root_title, f"root/{key}", children, root=True),
        "topicPositioning": "fixed",
    }


def walk(node: dict):
    yield node
    for child in node.get("children", {}).get("attached", []):
        yield from walk(child)


def main() -> None:
    detail_title, detail_children = parse_markdown(SOURCE)
    overview_title, overview_children = build_overview()

    overview_sheet = sheet("01 总览图", overview_title, overview_children, "overview")
    detail_sheet = sheet("02 详细知识树", detail_title, detail_children, "detail")
    content = [overview_sheet, detail_sheet]

    ids = [node["id"] for item in content for node in walk(item["rootTopic"])]
    if len(ids) != len(set(ids)):
        raise ValueError("Duplicate topic IDs detected")

    metadata = {
        "creator": {"name": "Xmind", "version": "24.0"},
        "activeSheetId": overview_sheet["id"],
    }
    manifest = {
        "file-entries": {
            "content.json": {},
            "metadata.json": {},
            "Thumbnails/thumbnail.png": {},
        }
    }
    transparent_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/"
        "W9n7WQAAAABJRU5ErkJggg=="
    )

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUTPUT, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("content.json", json.dumps(content, ensure_ascii=False, separators=(",", ":")))
        archive.writestr("metadata.json", json.dumps(metadata, ensure_ascii=False, separators=(",", ":")))
        archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, separators=(",", ":")))
        archive.writestr("Thumbnails/thumbnail.png", transparent_png)

    with zipfile.ZipFile(OUTPUT, "r") as archive:
        bad_member = archive.testzip()
        if bad_member:
            raise ValueError(f"Corrupt ZIP member: {bad_member}")
        required = {"content.json", "metadata.json", "manifest.json", "Thumbnails/thumbnail.png"}
        if not required.issubset(archive.namelist()):
            raise ValueError("Missing required XMind package members")
        loaded_content = json.loads(archive.read("content.json").decode("utf-8"))
        json.loads(archive.read("metadata.json").decode("utf-8"))
        json.loads(archive.read("manifest.json").decode("utf-8"))

    counts = [sum(1 for _ in walk(item["rootTopic"])) for item in loaded_content]
    print(f"Created: {OUTPUT}")
    print(f"Sheets: {len(loaded_content)}")
    print(f"Topics: overview={counts[0]}, detail={counts[1]}, total={sum(counts)}")
    print(f"Bytes: {OUTPUT.stat().st_size}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
