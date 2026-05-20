"""配置管理：读写 config.yaml。"""

from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

DEFAULT_CONFIG = {
    "version": "0.0.2",
    "wechat_monitor": {
        "enabled": True,
        "groups": [],
        "last_run": None,
        "linked_project": "",
    },
    "llm": {
        # 纯文本任务：模板分析、文档生成、数据提取、台账解析
        "text": {
            "base_url": "https://api.deepseek.com",
            "api_key": "",
            "model": "deepseek-v4-flash",
        },
        # 多模态任务：照片分析、影像分类、施工图识别
        "vision": {
            "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
            "api_key": "",
            "model": "qwen3.6-plus",
        },
    },
    "archive": {
        "base_path": "~/隐蔽验收影像资料",
        "naming": "{date}_{category}_{description}",
    },
    "projects": [],
}

# 向后兼容性：load() 时自动处理旧格式
_LLM_KEYS_MIGRATED = False

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
    _migrate_old_llm_format(cfg)
    return cfg


def _migrate_old_llm_format(cfg: dict[str, Any]) -> None:
    """将旧格式 {llm: {base_url, api_key, model}} 迁移到新格式 {llm: {text: {...}, vision: {...}}}。"""
    llm = cfg.get("llm", {})
    if "base_url" in llm or "model" in llm:

        # 保留旧的 api_key 如果有
        old_key = llm.get("api_key", "")

        # 旧 model 是 vision 类的视为 vision，否则当 text
        old_model = llm.get("model", "")
        old_url = llm.get("base_url", "")

        if "vl" in old_model.lower() or "vision" in old_model.lower():
            if old_key:
                llm.setdefault("vision", {})["api_key"] = old_key
            if old_model:
                llm.setdefault("vision", {})["model"] = old_model
            if old_url:
                llm.setdefault("vision", {})["base_url"] = old_url
        else:
            if old_key:
                llm.setdefault("text", {})["api_key"] = old_key
            if old_model:
                llm.setdefault("text", {})["model"] = old_model
            if old_url:
                llm.setdefault("text", {})["base_url"] = old_url

        # 清除旧字段
        for k in ("base_url", "api_key", "model"):
            llm.pop(k, None)


def get_llm_config(task_type: str = "text") -> dict[str, str]:
    """获取 LLM 配置。

    Args:
        task_type: "text" 纯文本任务 / "vision" 多模态任务
    """
    conf = load()
    llm = conf.get("llm", {})
    if task_type == "vision":
        return {
            "base_url": llm.get("vision", {}).get("base_url", ""),
            "api_key": llm.get("vision", {}).get("api_key", ""),
            "model": llm.get("vision", {}).get("model", ""),
        }
    return {
        "base_url": llm.get("text", {}).get("base_url", ""),
        "api_key": llm.get("text", {}).get("api_key", ""),
        "model": llm.get("text", {}).get("model", ""),
    }


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
    llm = cfg.get("llm", {})
    if not llm.get("text", {}).get("api_key"):
        errors.append("未配置 LLM text.api_key（纯文本模型）")
    if not llm.get("text", {}).get("base_url"):
        errors.append("未配置 LLM text.base_url")
    if not llm.get("vision", {}).get("api_key"):
        errors.append("未配置 LLM vision.api_key（多模态模型）")
    if not llm.get("vision", {}).get("base_url"):
        errors.append("未配置 LLM vision.base_url")
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
