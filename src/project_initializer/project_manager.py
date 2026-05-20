"""项目管理：创建/加载/更新项目，维护工作区目录和索引。数据持久化到 SQLite。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import json
import yaml

from .. import database

DEFAULT_STAGES = [
    "01_前期决策", "02_设计准备", "03_招标合同", "04_施工实施",
    "05_竣工验收", "06_结算审计", "07_后评估",
]

STAGE_LABELS = {
    "01_前期决策":  "前期决策（立项/可研/审批）",
    "02_设计准备":  "设计准备（初设/概算/施工图预算）",
    "03_招标合同":  "招标合同（招标/清单/合同签订）",
    "04_施工实施":  "施工实施（进度/质量/安全/变更/材料）",
    "05_竣工验收":  "竣工验收（验收/竣工图/移交）",
    "06_结算审计":  "结算审计（结算/决算/审计）",
    "07_后评估":   "后评估（复盘/经验/归档）",
}


def create(name: str, workspace: str, project_type: str = "",
           current_stage: str = "", stages: list[str] | None = None,
           wbs_json: str = "", contract_json: str = "", **meta) -> dict[str, Any]:
    ws = Path(workspace).expanduser().resolve()

    categories = meta.pop("categories", None) or []
    stage_list = stages or DEFAULT_STAGES
    if current_stage and current_stage not in stage_list:
        stage_list = [current_stage] + stage_list

    # 解析 WBS 和合同
    wbs_tree = None
    if wbs_json:
        try:
            data = json.loads(wbs_json)
            wbs_tree = data.get("wbs", data if isinstance(data, list) else [])
        except json.JSONDecodeError:
            pass

    contract = None
    if contract_json:
        try:
            contract = json.loads(contract_json)
            if isinstance(contract, dict) and "wbs" in contract:
                if not wbs_tree:
                    wbs_tree = contract.get("wbs", [])
                contract = contract.get("contract", contract)
        except json.JSONDecodeError:
            pass

    project = database.save_project(
        name=name, workspace=str(ws), project_type=project_type,
        current_stage=current_stage, meta=meta, stages=stage_list,
        categories=categories,
    )

    # 保存 WBS 和合同到数据库
    if wbs_tree:
        try:
            database.save_wbs(name, wbs_tree)
        except Exception:
            pass
    if contract:
        try:
            database.save_contract(name, contract)
        except Exception:
            pass

    # 创建工作区目录结构
    marker = ws / ".muniuliuma"
    marker.mkdir(parents=True, exist_ok=True)
    (marker / "project.yaml").write_text(
        yaml.safe_dump({"name": name}, allow_unicode=True), encoding="utf-8"
    )
    _create_stage_dirs(ws, stage_list)
    (ws / "templates").mkdir(parents=True, exist_ok=True)

    # 施工阶段 → 按分部分项创建资料子目录
    if current_stage == "04_施工实施" and wbs_tree:
        _create_wbs_dirs(ws / current_stage, wbs_tree)
    elif current_stage == "04_施工实施":
        _create_default_construction_dirs(ws / current_stage)

    return project


def _create_stage_dirs(workspace: Path, stages: list[str]) -> None:
    for stage in stages:
        stage_dir = workspace / stage
        stage_dir.mkdir(parents=True, exist_ok=True)
        (stage_dir / "README.md").write_text(
            f"# {STAGE_LABELS.get(stage, stage)}\n", encoding="utf-8")


def _create_wbs_dirs(base: Path, wbs: list[dict[str, Any]], depth: int = 0) -> None:
    """按分部分项层级递归创建资料目录。最多展开3层，第4层不建子目录。"""
    if depth >= 3:
        return

    DOC_SUBDIRS = ["隐蔽验收影像", "实验检测报告", "施工日志", "变更签证"]

    for node in wbs:
        code = node.get("code", "")
        name = node.get("name", "")
        dirname = f"{code}_{name}" if code else name
        dirname = dirname.replace("/", "_").replace(" ", "").replace("\\", "_")[:80]
        node_dir = base / dirname
        node_dir.mkdir(parents=True, exist_ok=True)

        children = node.get("children", [])
        if children:
            # 有子项 → 递归创建
            _create_wbs_dirs(node_dir, children, depth + 1)
        else:
            # 叶子节点（分项工程/检验批） → 创建资料子目录
            for sub in DOC_SUBDIRS:
                (node_dir / sub).mkdir(parents=True, exist_ok=True)


def _create_default_construction_dirs(stage_dir: Path) -> None:
    """没有 WBS 时的默认施工资料目录。"""
    for sub in ["隐蔽验收影像资料", "施工日志", "实验检测报告", "变更签证", "材料台账"]:
        (stage_dir / sub).mkdir(parents=True, exist_ok=True)


def list_projects() -> list[dict[str, Any]]:
    return database.list_projects()


def load(name: str) -> dict[str, Any] | None:
    return database.get_project(name)


def save(project: dict[str, Any]) -> None:
    database.save_project(
        name=project["name"],
        workspace=project.get("workspace", ""),
        project_type=project.get("type", project.get("project_type", "")),
        current_stage=project.get("current_stage", ""),
        meta=project.get("meta", {}),
        stages=project.get("stages", []),
        categories=project.get("meta", {}).get("categories", []),
    )


def update(name: str, **kwargs) -> dict[str, Any] | None:
    return database.update_project(name, **kwargs)


def add_stages(name: str, new_stages: list[str]) -> dict[str, Any] | None:
    project = load(name)
    if not project:
        return None
    ws = Path(project["workspace"])
    stages = list(project.get("stages", []))
    for s in new_stages:
        if s not in stages:
            stages.append(s)
    _create_stage_dirs(ws, new_stages)
    return update(name, stages=stages, meta=project.get("meta", {}))


def get_workspace(name: str) -> Path | None:
    project = load(name)
    return Path(project["workspace"]) if project else None


def discover_project(ws_path: str) -> dict[str, Any] | None:
    ws = Path(ws_path).expanduser().resolve()
    marker = ws / ".muniuliuma" / "project.yaml"
    if not marker.exists():
        return None
    with open(marker, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    name = data.get("name", "")
    return load(name) if name else None
