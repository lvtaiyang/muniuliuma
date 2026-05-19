"""施工日志生成：从影像资料或台账数据生成结构化施工日志。"""

from __future__ import annotations

import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openai import OpenAI
from .. import config as cfg
from ..wechat_monitor import summary_writer
from . import docx_filler, ledger_reader

LOG_SYSTEM_PROMPT = """你是工程施工领域的专业施工日志撰写专家。根据提供的施工影像资料和台账数据，生成规范的施工日志。

你必须返回 JSON 格式：

{
  "date": "YYYY-MM-DD",
  "weather": "晴/阴/雨/多云/...",
  "temperature": "温度描述，如 15-25℃",
  "sections": [
    {
      "heading": "标题（如：施工进度、材料进场、质量检验、安全检查、会议纪要等）",
      "content": "该段落的详细内容"
    }
  ],
  "key_events": ["今日关键事件1", "事件2"],
  "next_day_plan": "次日工作计划简述",
  "notes": "备注或特别说明"
}"""

LOG_FROM_LEDGERS_PROMPT = """你是工程施工领域的专业施工日志撰写专家。根据提供的台账资料，汇总生成施工日志。

台账包括：
- 材料进场台账：记录了当日进场的材料、规格、数量
- 实验检测台账：记录了当日进行的检测项目、结果
- 其他施工台账

你必须返回 JSON 格式：

{
  "date": "YYYY-MM-DD",
  "weather": "根据上下文推断或写未知",
  "sections": [
    {"heading": "材料进场", "content": "汇总当日进场的材料..."},
    {"heading": "实验检测", "content": "汇总当日检测情况..."},
    {"heading": "其他", "content": "..."}
  ],
  "key_events": ["..."],
  "next_day_plan": "根据施工进度推断次日计划",
  "notes": "数据来源说明"
}"""


def generate_from_images(
    project_name: str,
    log_date: str = "",
    max_images: int = 30,
    template_path: str = "",
) -> dict[str, Any]:
    """从归档影像资料生成施工日志。

    读取当天归档的图片路径，送给多模态模型分析，生成日志。
    """
    conf = cfg.load()
    llm = conf["llm"]
    archive_base = cfg.resolve_archive_path(conf)

    if not log_date:
        log_date = date.today().isoformat()

    # 收集当天归档的图片
    images = _find_images_by_date(archive_base, log_date, max_images)
    if not images:
        # 没有当天图片，尝试取最近的文件
        images = _find_recent_images(archive_base, max_images)
        if not images:
            return {"error": "没有找到可用的影像资料", "archive_path": str(archive_base)}

    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=180)

    # 构建多模态消息
    user_content: list[dict[str, Any]] = [
        {"type": "text", "text": f"请根据以下 {len(images)} 张施工照片，生成 {log_date} 的施工日志。"},
    ]
    for img_path in images:
        user_content.append({
            "type": "image_url",
            "image_url": {"url": _encode_image(img_path)},
        })

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": LOG_SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ],
        max_tokens=2000,
        temperature=0.3,
    )

    log_data = _parse_response(response.choices[0].message.content or "")

    # 保存：优先 Word 模板，否则 Markdown
    saved_path = _save_log(project_name, log_date, log_data, template_path)
    log_data["saved_to"] = str(saved_path) if saved_path else ""
    log_data["format"] = "docx" if template_path and saved_path else "md"

    return log_data


def generate_from_ledgers(
    project_name: str,
    ledger_paths: list[str],
    log_date: str = "",
    template_path: str = "",
) -> dict[str, Any]:
    """从台账文件生成施工日志。"""
    conf = cfg.load()
    llm = conf["llm"]

    if not log_date:
        log_date = date.today().isoformat()

    # 读取台账
    ledgers = ledger_reader.read_ledgers(ledger_paths)
    ledger_text = "\n\n".join(ledger_reader.format_ledger_for_llm(l) for l in ledgers)

    if not ledger_text.strip():
        return {"error": "台账文件为空或无法读取"}

    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=120)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": LOG_FROM_LEDGERS_PROMPT},
            {"role": "user", "content": f"台账数据:\n\n{ledger_text[:8000]}\n\n请生成 {log_date} 的施工日志。"},
        ],
        max_tokens=2000,
        temperature=0.3,
    )

    log_data = _parse_response(response.choices[0].message.content or "")
    log_data.setdefault("data_sources", [str(p) for p in ledger_paths])

    # 保存：优先 Word 模板，否则 Markdown
    saved_path = _save_log(project_name, log_date, log_data, template_path)
    log_data["saved_to"] = str(saved_path) if saved_path else ""
    log_data["format"] = "docx" if template_path and saved_path else "md"

    return log_data


def list_logs(project_name: str) -> list[dict[str, Any]]:
    """列出项目已有的施工日志。"""
    log_dir = _get_log_dir(project_name)
    if not log_dir or not log_dir.exists():
        return []
    logs = []
    for f in sorted(log_dir.glob("*.md"), reverse=True):
        logs.append({
            "filename": f.name,
            "date": f.stem[:10],
            "path": str(f),
            "size": f.stat().st_size,
        })
    return logs


def read_log(project_name: str, filename: str) -> str | None:
    """读取一篇施工日志内容。"""
    log_dir = _get_log_dir(project_name)
    if not log_dir:
        return None
    path = log_dir / filename
    if not path.exists():
        return None
    return path.read_text(encoding="utf-8")


# ── 内部函数 ────────────────────────────────────────────────────


def _get_log_dir(project_name: str) -> Path | None:
    """获取项目施工日志目录，不存在则创建。"""
    try:
        from ..project_initializer import project_manager
        proj = project_manager.load(project_name)
        if proj:
            log_dir = Path(proj["workspace"]) / "04_施工实施" / "施工日志"
            log_dir.mkdir(parents=True, exist_ok=True)
            return log_dir
    except ImportError:
        pass
    return None


def _save_log(project_name: str, log_date: str, log_data: dict[str, Any],
              template_path: str = "") -> Path | None:
    """保存施工日志。有模板则填 Word，否则存 Markdown。"""
    log_dir = _get_log_dir(project_name)
    if not log_dir:
        return None

    # Word 模板路径
    if template_path:
        try:
            fill_data = docx_filler.build_log_data(log_data)
            filename = f"{log_date}_施工日志.docx"
            path = log_dir / filename
            docx_filler.fill_template(template_path, fill_data, path)
            # 同时保存 MD 预览
            md_path = log_dir / f"{log_date}_施工日志_预览.md"
            md_path.write_text(_render_log_markdown(log_data), encoding="utf-8")
            return path
        except ImportError:
            pass  # 无 python-docx，回退到 MD

    # Fallback: Markdown
    md = _render_log_markdown(log_data)
    filename = f"{log_date}_施工日志.md"
    path = log_dir / filename
    path.write_text(md, encoding="utf-8")
    return path


def _render_log_markdown(log_data: dict[str, Any]) -> str:
    """将日志 JSON 渲染为 Markdown。"""
    d = log_data.get("date", date.today().isoformat())
    weather = log_data.get("weather", "")
    temp = log_data.get("temperature", "")

    lines = [
        f"# 施工日志",
        f"",
        f"**日期:** {d}",
        f"**天气:** {weather}  **温度:** {temp}",
        f"",
        f"---",
        f"",
    ]

    for section in log_data.get("sections", []):
        heading = section.get("heading", "")
        content = section.get("content", "")
        lines.append(f"## {heading}")
        lines.append("")
        lines.append(content)
        lines.append("")

    events = log_data.get("key_events", [])
    if events:
        lines.append("## 今日关键事件")
        lines.append("")
        for e in events:
            lines.append(f"- {e}")
        lines.append("")

    plan = log_data.get("next_day_plan", "")
    if plan:
        lines.append("## 次日计划")
        lines.append("")
        lines.append(plan)
        lines.append("")

    notes = log_data.get("notes", "")
    if notes:
        lines.append(f"> {notes}")
        lines.append("")

    data_sources = log_data.get("data_sources", [])
    if data_sources:
        lines.append("---")
        lines.append(f"*数据来源: {', '.join(data_sources)}*")

    sources_executed = log_data.get("saved_to", "")
    if sources_executed:
        lines.append(f"*生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}*")

    return "\n".join(lines)


def _find_images_by_date(archive_base: Path, target_date: str, max_count: int) -> list[Path]:
    """查找指定日期的归档图片。"""
    images = []
    for f in archive_base.rglob("*"):
        if f.is_file() and f.suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}:
            mtime = date.fromtimestamp(f.stat().st_mtime)
            if mtime.isoformat() == target_date:
                images.append(f)
    return sorted(images)[:max_count]


def _find_recent_images(archive_base: Path, max_count: int) -> list[Path]:
    """查找最近的归档图片。"""
    images = []
    for f in archive_base.rglob("*"):
        if f.is_file() and f.suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}:
            images.append((f.stat().st_mtime, f))
    images.sort(key=lambda x: x[0], reverse=True)
    return [f for _, f in images[:max_count]]


def _encode_image(path: Path) -> str:
    import base64
    data = path.read_bytes()
    b64 = base64.b64encode(data).decode("ascii")
    suffix = path.suffix.lower()
    mime_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp"}
    return f"data:{mime_map.get(suffix, 'image/jpeg')};base64,{b64}"


def _parse_response(content: str) -> dict[str, Any]:
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        pass
    match = re.search(r"```(?:json)?\s*\n?([\s\S]*?)\n?```", content)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass
    match = re.search(r"\{[\s\S]*\}", content)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass
    return {"raw": content, "date": date.today().isoformat(),
            "sections": [{"heading": "施工日志", "content": content}],
            "key_events": [], "notes": "LLM 返回格式异常"}
