"""项目管理：创建/加载/更新项目，维护工作区目录和索引。数据持久化到 SQLite。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

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
           current_stage: str = "", stages: list[str] | None = None, **meta) -> dict[str, Any]:
    ws = Path(workspace).expanduser().resolve()

    categories = meta.pop("categories", None) or []
    stage_list = stages or DEFAULT_STAGES

    project = database.save_project(
        name=name, workspace=str(ws), project_type=project_type,
        current_stage=current_stage, meta=meta, stages=stage_list,
        categories=categories,
    )

    # 创建工作区目录结构
    marker = ws / ".muniuliuma"
    marker.mkdir(parents=True, exist_ok=True)
    (marker / "project.yaml").write_text(
        yaml.safe_dump({"name": name}, allow_unicode=True), encoding="utf-8"
    )
    _create_stage_dirs(ws, stage_list)
    (ws / "templates").mkdir(parents=True, exist_ok=True)

    return project


def _create_stage_dirs(workspace: Path, stages: list[str]) -> None:
    for stage in stages:
        stage_dir = workspace / stage
        stage_dir.mkdir(parents=True, exist_ok=True)
        (stage_dir / "README.md").write_text(
            f"# {STAGE_LABELS.get(stage, stage)}\n", encoding="utf-8")


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
