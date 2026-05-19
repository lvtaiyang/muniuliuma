"""配置管理：读写 config.yaml。"""

from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

DEFAULT_CONFIG = {
    "version": "0.0.1",
    "wechat_monitor": {
        "enabled": True,
        "groups": [],
        "last_run": None,
        "linked_project": "",  # 关联的项目名，影像资料归档到项目工作区
    },
    "llm": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "api_key": "",
        "model": "qwen-vl-max",
    },
    "archive": {
        "base_path": "~/隐蔽验收影像资料",
        "naming": "{date}_{category}_{description}",
    },
    "projects": [],
}

CONFIG_DIR = Path.home() / ".muniuliuma"


def config_path() -> Path:
    local = Path("config.yaml")
    if local.exists():
        return local.resolve()
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    return CONFIG_DIR / "config.yaml"


def load() -> dict[str, Any]:
    path = config_path()
    if not path.exists():
        return deepcopy(DEFAULT_CONFIG)
    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    cfg = deepcopy(DEFAULT_CONFIG)
    _deep_merge(cfg, data)
    return cfg


def save(cfg: dict[str, Any]) -> Path:
    path = config_path()
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
    return path


def update_wechat_last_run() -> None:
    cfg = load()
    cfg.setdefault("wechat_monitor", {})["last_run"] = datetime.now().isoformat()
    save(cfg)


def resolve_archive_path(cfg: dict[str, Any]) -> Path:
    """解析归档路径。优先使用关联项目的工作区，否则使用全局配置。"""
    linked = cfg.get("wechat_monitor", {}).get("linked_project", "")
    if linked:
        try:
            from .project_initializer import project_manager
            proj = project_manager.load(linked)
            if proj:
                ws = Path(proj["workspace"])
                archive = ws / "04_施工实施" / "隐蔽验收影像资料"
                archive.mkdir(parents=True, exist_ok=True)
                return archive
        except ImportError:
            pass
    return Path(cfg["archive"]["base_path"]).expanduser().resolve()


def validate(cfg: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    wm = cfg.get("wechat_monitor", {})
    if wm.get("enabled"):
        if not wm.get("groups"):
            errors.append("微信监控: 未配置群聊 (groups 为空)")
    if wm.get("enabled") and not cfg.get("projects"):
        errors.append("未配置项目信息 (projects 为空)")
    if not cfg.get("llm", {}).get("api_key"):
        errors.append("未配置 LLM API key")
    if not cfg.get("llm", {}).get("base_url"):
        errors.append("未配置 LLM base_url")
    for proj in cfg.get("projects", []):
        if not proj.get("name"):
            errors.append("项目缺少 name 字段")
        if not proj.get("categories"):
            errors.append(f"项目 '{proj.get('name', '?')}' 缺少 categories")
    return errors


def _deep_merge(base: dict, override: dict) -> None:
    for k, v in override.items():
        if k in base and isinstance(base[k], dict) and isinstance(v, dict):
            _deep_merge(base[k], v)
        else:
            base[k] = v
