"""模板管理：保存/列表/应用模板，生成文档实例。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any
import json

import yaml

from . import project_manager


def save_template(
    project_name: str,
    template_def: dict[str, Any],
    filename: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """保存分析后的模板到项目 templates/ 目录。

    Returns:
        {saved: true, path: "...", template: {...}, message: "..."}
    """
    project = project_manager.load(project_name)
    if project is None:
        return {"error": f"项目不存在: {project_name}"}

    ws = Path(project["workspace"])
    tmpl_dir = ws / "templates"
    tmpl_dir.mkdir(parents=True, exist_ok=True)

    tname = filename or template_def.get("template_name", "未命名")
    safe_name = tname.replace("/", "_").replace(" ", "_")
    tmpl_path = tmpl_dir / f"{safe_name}.yaml"

    if tmpl_path.exists() and not overwrite:
        return {
            "saved": False,
            "path": str(tmpl_path),
            "message": f"模板已存在: {safe_name}.yaml，使用 overwrite=true 覆盖",
        }

    # 补充元数据
    template_def.setdefault("template_name", tname)
    template_def["saved_at"] = datetime.now().isoformat()
    template_def["source_file"] = filename or ""

    with open(tmpl_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(template_def, f, allow_unicode=True, default_flow_style=False, sort_keys=False)

    # 更新项目索引中的模板引用
    tmpl_ref = {
        "name": tname,
        "path": str(tmpl_path),
        "applies_to": template_def.get("applies_to", {}),
        "fields_count": len(template_def.get("fields", [])),
    }
    project.setdefault("templates", [])
    # 去重
    existing_names = [t.get("name") for t in project["templates"]]
    if tname in existing_names:
        for t in project["templates"]:
            if t.get("name") == tname:
                t.update(tmpl_ref)
                break
    else:
        project["templates"].append(tmpl_ref)
    project_manager.save(project)

    return {
        "saved": True,
        "path": str(tmpl_path),
        "template": template_def,
        "message": f"模板已保存: {tname}",
    }


def list_templates(project_name: str) -> list[dict[str, Any]]:
    """列出项目的所有模板。"""
    project = project_manager.load(project_name)
    if project is None:
        return []
    return project.get("templates", [])


def load_template(project_name: str, template_name: str) -> dict[str, Any] | None:
    """加载模板完整定义。"""
    project = project_manager.load(project_name)
    if project is None:
        return None
    for tmpl in project.get("templates", []):
        if tmpl.get("name") == template_name:
            path = Path(tmpl["path"])
            if path.exists():
                with open(path, encoding="utf-8") as f:
                    return yaml.safe_load(f)
    return None


def apply_template(
    project_name: str,
    template_name: str,
    field_values: dict[str, Any],
    context: str = "",
) -> dict[str, Any]:
    """用模板生成一个文档实例。

    Args:
        project_name: 项目名
        template_name: 模板名
        field_values: 字段值 {字段名: 值}
        context: 场景描述（如"钢筋隐蔽验收-3号楼"）

    Returns:
        {generated: true, path: "...", filename: "..."}
    """
    project = project_manager.load(project_name)
    if project is None:
        return {"error": f"项目不存在: {project_name}"}

    template = load_template(project_name, template_name)
    if template is None:
        return {"error": f"模板不存在: {template_name}"}

    ws = Path(project["workspace"])

    # 确定输出目录
    applies_to = template.get("applies_to", {})
    stage = applies_to.get("stage", "")
    if stage and stage in project.get("stages", []):
        out_dir = ws / stage
    else:
        out_dir = ws

    # 生成文件名
    date_str = datetime.now().strftime("%Y%m%d")
    context_suffix = f"_{context}" if context else ""
    safe_tname = template_name.replace("/", "_").replace(" ", "_")
    filename = f"{safe_tname}{context_suffix}_{date_str}.md"

    # 生成文档内容（Markdown 格式）
    content = _render_template_doc(template, field_values)

    out_path = out_dir / filename
    # 重名处理
    counter = 1
    while out_path.exists():
        out_path = out_dir / f"{safe_tname}{context_suffix}_{date_str}_{counter}.md"
        counter += 1

    out_path.write_text(content, encoding="utf-8")

    # 增加使用计数
    _increment_usage(project, template_name)

    return {
        "generated": True,
        "path": str(out_path),
        "filename": out_path.name,
        "stage": stage or project.get("current_stage", ""),
        "template": template_name,
    }


def _render_template_doc(template: dict[str, Any], values: dict[str, Any]) -> str:
    """用字段值和模板定义渲染 Markdown 文档。"""
    lines = [
        f"# {template.get('template_name', '文档')}",
        "",
    ]
    desc = template.get("description", "")
    if desc:
        lines.append(f"> {desc}")
        lines.append("")

    applies = template.get("applies_to", {})
    if applies:
        lines.append(f"- 阶段: {applies.get('stage', '-')}")
        lines.append(f"- 场景: {applies.get('scenario', '-')}")
        lines.append("")

    notes = template.get("notes", "")
    if notes:
        lines.append(f"**填写说明:** {notes}")
        lines.append("")

    lines.append(f"**生成时间:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    lines.append("---")
    lines.append("")

    # 字段表格
    lines.append("## 信息记录")
    lines.append("")
    lines.append("| 字段 | 值 | 类型 |")
    lines.append("|------|-----|------|")

    for field in template.get("fields", []):
        fname = field.get("name", "")
        ftype = field.get("type", "string")
        value = values.get(fname, values.get(fname.lower(), ""))
        if isinstance(value, list):
            value = ", ".join(str(v) for v in value)
        else:
            value = str(value) if value else ""
        lines.append(f"| {fname} | {value} | {ftype} |")

    lines.append("")
    lines.append("---")
    lines.append("*此文档由木牛流马自动生成*")

    return "\n".join(lines)


def _increment_usage(project: dict[str, Any], template_name: str) -> None:
    for t in project.get("templates", []):
        if t.get("name") == template_name:
            t["usage_count"] = t.get("usage_count", 0) + 1
            project_manager.save(project)
            return


def suggest_stage_dir(project_name: str, applies_to: dict) -> str:
    """根据模板的 applies_to 推测应该在哪个阶段目录。"""
    stage = applies_to.get("stage", "")
    project = project_manager.load(project_name)
    if project and stage in project.get("stages", []):
        return str(Path(project["workspace"]) / stage)
    if project:
        return project["workspace"]
    return ""
