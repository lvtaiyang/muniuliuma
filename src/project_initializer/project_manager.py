"""项目管理：创建/加载/更新项目，维护工作区目录和索引。"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

PROJECTS_DIR = Path.home() / ".muniuliuma" / "projects"

# 工程建设全生命周期 7 阶段
DEFAULT_STAGES = [
    "01_前期决策",
    "02_设计准备",
    "03_招标合同",
    "04_施工实施",
    "05_竣工验收",
    "06_结算审计",
    "07_后评估",
]

STAGE_LABELS = {
    "01_前期决策":   "前期决策（立项/可研/审批）",
    "02_设计准备":   "设计准备（初设/概算/施工图预算）",
    "03_招标合同":   "招标合同（招标/清单/合同签订）",
    "04_施工实施":   "施工实施（进度/质量/安全/变更/材料）",
    "05_竣工验收":   "竣工验收（验收/竣工图/移交）",
    "06_结算审计":   "结算审计（结算/决算/审计）",
    "07_后评估":    "后评估（复盘/经验/归档）",
}


def _index_path() -> Path:
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
    return PROJECTS_DIR / "_index.yaml"


def _project_index_path(project_name: str) -> Path:
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
    safe = project_name.replace("/", "_").replace(" ", "_")
    return PROJECTS_DIR / f"{safe}.yaml"


def _load_index() -> dict[str, Any]:
    path = _index_path()
    if not path.exists():
        return {"projects": {}}
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {"projects": {}}


def _save_index(data: dict[str, Any]) -> None:
    with open(_index_path(), "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


def create(
    name: str,
    workspace: str,
    project_type: str = "",
    current_stage: str = "",
    stages: list[str] | None = None,
    **meta,
) -> dict[str, Any]:
    """创建新项目。

    Args:
        name: 项目名称
        workspace: 用户指定的工作区绝对路径
        project_type: 项目类型（住宅/工业/市政/公路/水利...）
        current_stage: 当前所处阶段
        stages: 自定义阶段列表，默认 7 阶段
        **meta: 其他元数据（总投资、工期、地点等）
    """
    ws = Path(workspace).expanduser().resolve()

    # 创建索引记录
    project = {
        "name": name,
        "workspace": str(ws),
        "type": project_type,
        "current_stage": current_stage,
        "stages": stages or DEFAULT_STAGES,
        "meta": meta,
        "templates": [],
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat(),
    }

    # 保存项目索引文件
    idx_path = _project_index_path(name)
    with open(idx_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(project, f, allow_unicode=True, default_flow_style=False, sort_keys=False)

    # 创建项目根目录下的 .muniuliuma 隐藏标记
    marker = ws / ".muniuliuma"
    marker.mkdir(parents=True, exist_ok=True)
    (marker / "project.yaml").write_text(
        yaml.safe_dump({"name": name}, allow_unicode=True), encoding="utf-8"
    )

    # 创建阶段子目录
    _create_stage_dirs(ws, project["stages"])

    # 创建 templates 子目录
    (ws / "templates").mkdir(parents=True, exist_ok=True)

    # 更新全局索引
    index = _load_index()
    index["projects"][name] = {
        "workspace": str(ws),
        "type": project_type,
        "current_stage": current_stage,
        "created_at": project["created_at"],
    }
    _save_index(index)

    return project


def _create_stage_dirs(workspace: Path, stages: list[str]) -> None:
    for stage in stages:
        stage_dir = workspace / stage
        stage_dir.mkdir(parents=True, exist_ok=True)
        # 每个阶段放一个 README 说明
        readme = stage_dir / "README.md"
        label = STAGE_LABELS.get(stage, stage)
        readme.write_text(f"# {label}\n", encoding="utf-8")


def list_projects() -> list[dict[str, Any]]:
    """列出所有项目。"""
    index = _load_index()
    result = []
    for name, info in index.get("projects", {}).items():
        result.append({"name": name, **info})
    return result


def load(name: str) -> dict[str, Any] | None:
    """加载项目完整信息。"""
    path = _project_index_path(name)
    if not path.exists():
        return None
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def save(project: dict[str, Any]) -> None:
    """保存项目信息。"""
    project["updated_at"] = datetime.now().isoformat()
    name = project["name"]
    path = _project_index_path(name)
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(project, f, allow_unicode=True, default_flow_style=False, sort_keys=False)

    # 更新索引
    index = _load_index()
    if name in index.get("projects", {}):
        index["projects"][name].update({
            "type": project.get("type", ""),
            "current_stage": project.get("current_stage", ""),
            "updated_at": project["updated_at"],
        })
        _save_index(index)


def update(name: str, **kwargs) -> dict[str, Any] | None:
    """更新项目字段。"""
    project = load(name)
    if project is None:
        return None
    project.update(kwargs)
    save(project)
    return project


def add_stages(name: str, new_stages: list[str]) -> dict[str, Any] | None:
    """追加阶段目录。"""
    project = load(name)
    if project is None:
        return None
    ws = Path(project["workspace"])
    for s in new_stages:
        if s not in project["stages"]:
            project["stages"].append(s)
    _create_stage_dirs(ws, new_stages)
    save(project)
    return project


def get_workspace(name: str) -> Path | None:
    project = load(name)
    if project is None:
        return None
    return Path(project["workspace"])


def discover_project(ws_path: str) -> dict[str, Any] | None:
    """从工作区路径反向查找项目索引。"""
    ws = Path(ws_path).expanduser().resolve()
    marker = ws / ".muniuliuma" / "project.yaml"
    if not marker.exists():
        return None
    with open(marker, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    name = data.get("name", "")
    if name:
        return load(name)
    return None
