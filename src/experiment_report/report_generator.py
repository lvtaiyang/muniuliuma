"""实验报告生成：从台账数据 + xlsx 模板生成填充后的实验检测报告。

核心流程：
1. 读取台账数据
2. LLM 从台账中提取结构化检测数据
3. 按模板定义填充 xlsx 各单元格
4. 展开数据表区域（每样品一行）
5. 应用逻辑规则（计算、判定）
"""

from __future__ import annotations

import copy
import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openai import OpenAI
from .. import config as cfg
from ..construction_log import ledger_reader
from . import template_analyzer

DATA_EXTRACTION_PROMPT = """你是工程实验检测数据提取专家。根据模板定义的字段和台账数据，提取/整理生成实验报告所需的数据。

模板中的每个 region 都有 data_source 说明，你需要根据这些说明从台账中找到对应数据。

## 返回格式

你必须返回以下 JSON：

{
  "report_data": {
    "region_id_1": "填充的值",
    "region_id_2": "填充的值",
    ...
  },
  "data_table_rows": [
    {
      "col_A": "值",
      "col_B": "值",
      ...
    },
    ...
  ],
  "notes": "数据提取说明"
}

其中：
- report_data: 模板中每个非表格 region 的填充值，key 是 region 的 id
- data_table_rows: 数据表每一行的填充值，key 是列的 col_letter

## 重要规则
1. 如果台账中找不到对应数据，填写 "见原始记录"
2. 数据表的行数应等于台账中的实际样品数
3. 注意数值精度，保持与台账一致
"""


def generate_from_template(
    project_name: str,
    template_name: str,
    ledger_paths: list[str],
    report_date: str = "",
) -> dict[str, Any]:
    """使用已确认的模板定义 + 台账生成实验报告。

    Args:
        project_name: 项目名称
        template_name: 模板定义文件名
        ledger_paths: 台账文件路径列表
        report_date: 报告日期
    """
    if not report_date:
        report_date = date.today().isoformat()

    # 1. 加载模板定义
    tmpl_def = template_analyzer.load_template_definition(project_name, template_name)
    if not tmpl_def:
        return {"error": f"模板定义不存在: {template_name}"}

    if not tmpl_def.get("confirmed"):
        remaining = tmpl_def.get("confirmation_required", "?")
        return {
            "error": f"模板定义尚未确认（还有 {remaining} 个待确认问题），"
                     f"请先用 confirm_template_analysis 确认后再生成报告",
            "uncertainties": tmpl_def.get("uncertainties", []),
        }

    # 2. 读取台账
    ledgers = ledger_reader.read_ledgers(ledger_paths)
    ledger_text = "\n\n".join(ledger_reader.format_ledger_for_llm(l) for l in ledgers)
    if not ledger_text.strip():
        return {"error": "台账文件为空或无法读取"}

    # 3. 构建模板描述给 LLM 提取数据
    template_desc = _describe_template_for_extraction(tmpl_def)

    # 4. LLM 从台账提取数据
    conf = cfg.load()
    llm = cfg.get_llm_config("text")
    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=180)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": DATA_EXTRACTION_PROMPT},
            {"role": "user", "content": (
                f"## 模板定义\n\n{template_desc}\n\n"
                f"## 台账数据\n\n{ledger_text[:12000]}\n\n"
                f"报告日期: {report_date}\n\n"
                f"请根据模板定义从台账中提取对应的数据。"
            )},
        ],
        max_tokens=4000,
        temperature=0.2,
    )

    extracted = _parse_response(response.choices[0].message.content or "")

    # 5. 通过 win32com 填充 xlsx（保留原生格式）
    source_file = tmpl_def.get("source_file", "")
    if not source_file or not Path(source_file).exists():
        return {"error": f"模板源文件不存在: {source_file}"}

    from .. import win32_helper
    output_dir = _get_report_dir(project_name)
    if not output_dir:
        return {"error": f"项目不存在: {project_name}"}
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', tmpl_def.get("template_name", "report"))[:30]
    filename = f"{report_date}_{safe_name}.xlsx"
    output_path = (output_dir / filename).resolve()

    try:
        output_path = win32_helper.excel_fill_template(
            template_path=source_file,
            output_path=output_path,
            fill_data={
                "sheets": tmpl_def.get("sheets", []),
                "report_data": extracted.get("report_data", {}),
                "data_table_rows": extracted.get("data_table_rows", []),
                "logic_rules": tmpl_def.get("logic_rules", []),
            },
        )
    except Exception as e:
        return {"error": f"填充 xlsx 失败: {e}", "extracted_data": extracted}

    preview_path = _save_preview(project_name, report_date, tmpl_def, extracted)

    return {
        "template_name": tmpl_def.get("template_name", ""),
        "report_type": tmpl_def.get("report_type", ""),
        "saved_to": str(output_path) if output_path else "",
        "preview": str(preview_path) if preview_path else "",
        "format": "xlsx",
        "data_rows": len(extracted.get("data_table_rows", [])),
        "notes": extracted.get("notes", ""),
    }


def generate_single_from_template(
    project_name: str,
    template_name: str,
    ledger_paths: list[str],
    test_item: str,
    report_date: str = "",
    extra_context: str = "",
) -> dict[str, Any]:
    """针对特定检测项目，用模板生成单份实验报告。

    与 generate_from_template 相同，但只提取指定检测项目的数据。
    """
    if not report_date:
        report_date = date.today().isoformat()

    tmpl_def = template_analyzer.load_template_definition(project_name, template_name)
    if not tmpl_def:
        return {"error": f"模板定义不存在: {template_name}"}

    if not tmpl_def.get("confirmed"):
        remaining = tmpl_def.get("confirmation_required", "?")
        return {
            "error": f"模板定义尚未确认（还有 {remaining} 个待确认问题）",
            "uncertainties": tmpl_def.get("uncertainties", []),
        }

    ledgers = ledger_reader.read_ledgers(ledger_paths)
    ledger_text = "\n\n".join(ledger_reader.format_ledger_for_llm(l) for l in ledgers)
    if not ledger_text.strip():
        return {"error": "台账文件为空或无法读取"}

    template_desc = _describe_template_for_extraction(tmpl_def)

    conf = cfg.load()
    llm = cfg.get_llm_config("text")
    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=180)

    context = f"\n\n额外说明: {extra_context}" if extra_context else ""

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": DATA_EXTRACTION_PROMPT},
            {"role": "user", "content": (
                f"请只提取检测项目为「{test_item}」的数据。{context}\n\n"
                f"## 模板定义\n\n{template_desc}\n\n"
                f"## 台账数据\n\n{ledger_text[:12000]}\n\n"
                f"报告日期: {report_date}"
            )},
        ],
        max_tokens=4000,
        temperature=0.2,
    )

    extracted = _parse_response(response.choices[0].message.content or "")

    source_file = tmpl_def.get("source_file", "")
    if not source_file or not Path(source_file).exists():
        return {"error": f"模板源文件不存在: {source_file}"}

    from .. import win32_helper
    output_dir = _get_report_dir(project_name)
    if not output_dir:
        return {"error": f"项目不存在: {project_name}"}
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', f"{test_item}_{tmpl_def.get('template_name', 'report')}")[:40]
    filename = f"{report_date}_{safe_name}.xlsx"
    output_path = (output_dir / filename).resolve()

    try:
        output_path = win32_helper.excel_fill_template(
            template_path=source_file,
            output_path=output_path,
            fill_data={
                "sheets": tmpl_def.get("sheets", []),
                "report_data": extracted.get("report_data", {}),
                "data_table_rows": extracted.get("data_table_rows", []),
                "logic_rules": tmpl_def.get("logic_rules", []),
            },
        )
    except Exception as e:
        return {"error": f"填充 xlsx 失败: {e}", "extracted_data": extracted}

    preview_path = _save_preview(project_name, report_date, tmpl_def, extracted)

    return {
        "test_item": test_item,
        "template_name": tmpl_def.get("template_name", ""),
        "saved_to": str(output_path) if output_path else "",
        "preview": str(preview_path) if preview_path else "",
        "format": "xlsx",
        "data_rows": len(extracted.get("data_table_rows", [])),
        "notes": extracted.get("notes", ""),
    }


def list_reports(project_name: str) -> list[dict[str, Any]]:
    """列出项目已有的实验报告。"""
    report_dir = _get_report_dir(project_name)
    if not report_dir or not report_dir.exists():
        return []
    reports = []
    for f in sorted(report_dir.glob("*"), reverse=True):
        if f.name.startswith("_") or f.name.endswith("_预览.md"):
            continue
        reports.append({
            "filename": f.name,
            "date": f.stem[:10] if f.stem else "",
            "path": str(f),
            "size": f.stat().st_size,
            "format": f.suffix,
        })
    return reports


def read_report(project_name: str, filename: str) -> str | None:
    """读取一篇实验报告内容（返回 Markdown 预览，xlsx 返回结构摘要）。"""
    report_dir = _get_report_dir(project_name)
    if not report_dir:
        return None
    path = report_dir / filename
    if not path.exists():
        return None

    if path.suffix in (".xlsx", ".xlsm"):
        return _read_xlsx_preview(path)
    return path.read_text(encoding="utf-8")


# ── xlsx 填充 —— 通过 win32_helper.excel_fill_template() 完成 ──


# ── 辅助函数 ────────────────────────────────────────────────────

def _describe_template_for_extraction(tmpl_def: dict[str, Any]) -> str:
    """将模板定义转为 LLM 可读的数据提取说明。支持多 sheet 格式。"""
    lines = [f"模板名称: {tmpl_def.get('template_name', '')}"]
    lines.append(f"报告类型: {tmpl_def.get('report_type', '')}")
    lines.append(f"工作表数量: {len(tmpl_def.get('sheets', []))}")
    lines.append("")

    for sheet in tmpl_def.get("sheets", []):
        sname = sheet.get("sheet_name", "")
        srole = sheet.get("sheet_role", "")
        lines.append(f"## 工作表: {sname} ({srole})")
        lines.append(f"说明: {sheet.get('sheet_description', '')}")
        lines.append("")

        lines.append("### 需要填充的区域:")
        for region in sheet.get("regions", []):
            if region.get("type") in ("data_table_header", "data_table_body"):
                continue
            lines.append(f"- [{region.get('id', '')}] {region.get('label', '')}")
            lines.append(f"  位置: {region.get('cells', '')}")
            lines.append(f"  作用: {region.get('purpose', '')}")
            lines.append(f"  数据来源: {region.get('data_source_detail', region.get('data_source', ''))}")
            if region.get("fixed_value"):
                lines.append(f"  固定值: {region['fixed_value']}")
            lines.append("")

        data_table = sheet.get("data_table", {})
        if data_table:
            lines.append("### 数据表结构:")
            lines.append(f"表头行: 第{data_table.get('header_row', '?')}行")
            lines.append(f"数据起始行: 第{data_table.get('data_start_row', '?')}行")
            lines.append("列映射:")
            for col in data_table.get("columns", []):
                lines.append(f"  - 列{col.get('col_letter', '')}: {col.get('header_text', '')} → {col.get('data_source', '')}")
            if data_table.get("expand_logic"):
                lines.append(f"展开逻辑: {data_table['expand_logic']}")
            lines.append("")

    cross = tmpl_def.get("cross_sheet_logic", [])
    if cross:
        lines.append("## 跨工作表数据流:")
        for c in cross:
            lines.append(f"- {c.get('description', '')}: {c.get('source_sheet', '')}.{c.get('source_cells', '')} → {c.get('target_sheet', '')}.{c.get('target_cells', '')}")

    logic_rules = tmpl_def.get("logic_rules", [])
    if logic_rules:
        lines.append("\n## 逻辑规则:")
        for rule in logic_rules:
            lines.append(f"- {rule.get('description', '')}: {rule.get('expression', '')}")
            lines.append(f"  应用于: {rule.get('applies_in_sheet', '')}.{rule.get('applies_to', '')}")
        lines.append("")

    return "\n".join(lines)


def _get_report_dir(project_name: str) -> Path | None:
    try:
        from ..project_initializer import project_manager
        proj = project_manager.load(project_name)
        if proj:
            report_dir = Path(proj["workspace"]) / "04_施工实施" / "实验检测报告"
            report_dir.mkdir(parents=True, exist_ok=True)
            return report_dir
    except ImportError:
        pass
    return None


def _save_preview(
    project_name: str,
    report_date: str,
    tmpl_def: dict[str, Any],
    extracted: dict[str, Any],
) -> Path | None:
    """保存 Markdown 预览。"""
    report_dir = _get_report_dir(project_name)
    if not report_dir:
        return None

    lines = [f"# {tmpl_def.get('template_name', '实验报告')} — 预览", ""]
    lines.append(f"**日期:** {report_date}")
    lines.append(f"**报告类型:** {tmpl_def.get('report_type', '')}")
    lines.append("")

    report_data = extracted.get("report_data", {})
    if report_data:
        lines.append("## 报告信息")
        lines.append("")
        for sheet in tmpl_def.get("sheets", []):
            for region in sheet.get("regions", []):
                if region.get("type") in ("data_table_header", "data_table_body"):
                    continue
                rid = region.get("id", "")
                label = region.get("label", rid)
                value = report_data.get(rid, "")
                if value:
                    lines.append(f"- **{label}:** {value}")
        lines.append("")

    data_rows = extracted.get("data_table_rows", [])
    if data_rows:
        lines.append(f"## 检测数据 ({len(data_rows)} 条)")
        lines.append("")
        # 从 main_report 工作表获取数据表列定义
        data_table = {}
        for s in tmpl_def.get("sheets", []):
            if s.get("data_table") and s.get("sheet_role") == "main_report":
                data_table = s["data_table"]
                break
        if not data_table:
            data_table = tmpl_def.get("data_table", {})
        columns = data_table.get("columns", [])
        if columns:
            headers = [c.get("header_text", c.get("col_letter", "")) for c in columns]
            lines.append("| " + " | ".join(headers) + " |")
            lines.append("|" + "|".join(["---"] * len(headers)) + "|")
            for row in data_rows:
                vals = [str(row.get(c.get("col_letter", ""), "")) for c in columns]
                lines.append("| " + " | ".join(vals) + " |")
        lines.append("")

    notes = extracted.get("notes", "")
    if notes:
        lines.append(f"> {notes}")

    safe_name = re.sub(r'[\\/:*?"<>|]', '_', tmpl_def.get("template_name", "report"))[:20]
    preview_path = report_dir / f"{report_date}_{safe_name}_预览.md"
    preview_path.write_text("\n".join(lines), encoding="utf-8")
    return preview_path


def _read_xlsx_preview(path: Path) -> str:
    """读取 xlsx 文件返回 Markdown 预览。仅通过 COM 操作。"""
    from .. import win32_helper
    try:
        structure = win32_helper.excel_read_structure(path)
        if "error" in structure:
            return f"无法读取 xlsx: {structure['error']}"
        lines = [f"# 实验报告: {path.name}", ""]
        for sheet in structure.get("sheets", []):
            lines.append(f"## {sheet['name']}")
            for row_info in sheet.get("rows", []):
                cells = [c["value"] for c in row_info["cells"]]
                lines.append(" | ".join(cells))
            lines.append("")
        return "\n".join(lines)
    except Exception as e:
        return f"无法读取 xlsx: {e}"


def _parse_response(content: str) -> dict[str, Any]:
    from .. import json_utils
    result = json_utils.parse_llm_json(content)
    if "_raw" in result:
        return {"report_data": {}, "data_table_rows": [],
                "notes": f"LLM 返回格式异常: {content[:300]}"}
    return result
