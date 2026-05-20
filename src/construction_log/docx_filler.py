"""Word 模板填充：通过 win32com 驱动本地 Word/WPS 替换占位符并另存。

模板使用 {{field_name}} 占位符，支持段落、表格、页眉页脚。
通过原生 Word/WPS 执行操作，完美保留合并单元格等所有格式。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .. import win32_helper


def fill_template(
    template_path: str | Path,
    data: dict[str, Any],
    output_path: str | Path,
) -> Path:
    """将数据填入 Word 模板并另存为新文件。

    通过 Word/WPS 原生 Find/Replace 替换 {{key}} 占位符。
    格式完全保留，包括合并单元格、边框、字体样式等。

    Args:
        template_path: .docx 模板文件路径
        data: 键值对，key 对应模板中的占位符名（不含大括号）
        output_path: 输出 .docx 路径
    """
    fill_data = {k: _format_value(v) for k, v in data.items()}
    return win32_helper.word_fill_placeholders(
        template_path=template_path,
        output_path=output_path,
        data=fill_data,
    )


def _format_value(value: Any) -> str:
    if isinstance(value, list):
        return "\n".join(f"    {v}" for v in value)
    if isinstance(value, dict):
        return "\n".join(f"{k}: {v}" for k, v in value.items())
    return str(value) if value is not None else ""


# ── 施工日志模板占位符约定 ──────────────────────────────────────

LOG_PLACEHOLDERS = {
    "日期": "date",
    "天气": "weather",
    "温度": "temperature",
    "施工进度": "progress",
    "材料进场": "materials",
    "质量检验": "quality",
    "安全检查": "safety",
    "关键事件": "key_events",
    "次日计划": "next_day_plan",
    "备注": "notes",
}


def build_log_data(log_json: dict[str, Any]) -> dict[str, str]:
    """将 LLM 返回的日志 JSON 转为模板填充用的扁平化键值对。"""
    data: dict[str, str] = {}

    data["日期"] = log_json.get("date", "")
    data["天气"] = log_json.get("weather", "")
    data["温度"] = log_json.get("temperature", "")

    sections = log_json.get("sections", [])
    for s in sections:
        heading = s.get("heading", "")
        content = s.get("content", "")
        key = _map_heading(heading)
        data[key] = content

    events = log_json.get("key_events", [])
    data["关键事件"] = "\n".join(f"{i+1}. {e}" for i, e in enumerate(events))

    data["次日计划"] = log_json.get("next_day_plan", "")
    data["备注"] = log_json.get("notes", "")

    return data


def _map_heading(heading: str) -> str:
    mapping = {
        "施工进度": "施工进度",
        "进度": "施工进度",
        "材料进场": "材料进场",
        "材料": "材料进场",
        "质量检验": "质量检验",
        "质量": "质量检验",
        "安全检查": "安全检查",
        "安全": "安全检查",
    }
    for pat, key in mapping.items():
        if pat in heading:
            return key
    return heading
