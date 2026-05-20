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

    # 5. 加载 xlsx 模板并填充
    source_file = tmpl_def.get("source_file", "")
    if not source_file or not Path(source_file).exists():
        return {"error": f"模板源文件不存在: {source_file}，请重新分析上传模板"}

    try:
        output_path = _fill_xlsx(
            template_path=source_file,
            template_def=tmpl_def,
            extracted_data=extracted,
            report_date=report_date,
            project_name=project_name,
        )
    except Exception as e:
        return {"error": f"填充 xlsx 失败: {e}", "extracted_data": extracted}

    # 保存 Markdown 预览
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

    try:
        output_path = _fill_xlsx(
            template_path=source_file,
            template_def=tmpl_def,
            extracted_data=extracted,
            report_date=report_date,
            project_name=project_name,
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


# ── xlsx 填充引擎 (win32com) ───────────────────────────────────

def _fill_xlsx(
    template_path: str | Path,
    template_def: dict[str, Any],
    extracted_data: dict[str, Any],
    report_date: str,
    project_name: str,
) -> Path:
    """用提取的数据填充 xlsx 模板。

    Windows + Office/WPS: win32com 原生操作，完美保留格式
    其他环境: openpyxl 回退模式（用于测试逻辑）
    """
    from .. import win32_helper

    output_dir = _get_report_dir(project_name)
    if not output_dir:
        raise RuntimeError(f"项目不存在: {project_name}")

    safe_name = re.sub(r'[\\/:*?"<>|]', '_', template_def.get("template_name", "report"))[:30]
    filename = f"{report_date}_{safe_name}.xlsx"
    output_path = (output_dir / filename).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    report_data = extracted_data.get("report_data", {})
    data_table_rows = extracted_data.get("data_table_rows", [])

    with win32_helper.open_excel(template_path) as (app, wb, app_name):
        is_com = win32_helper.WIN32_AVAILABLE and app is not None

        # 获取目标工作表
        if is_com:
            ws = wb.Worksheets(1)
        else:
            ws = wb.active

        # 1. 填充固定区域
        for region in template_def.get("regions", []):
            cells = region.get("cells", "")
            rid = region.get("id", "")
            value = report_data.get(rid, "")

            if not value and region.get("fixed_value"):
                value = region["fixed_value"]
            if not value:
                continue

            target_cell = cells.split(":")[0] if ":" in cells else cells
            try:
                if is_com:
                    ws.Range(target_cell).Value = value
                else:
                    _write_cell_ol(ws, target_cell, value)
            except Exception:
                pass

        # 2. 填充数据表
        data_table = template_def.get("data_table", {})
        if data_table and data_table_rows:
            header_row = data_table.get("data_start_row", data_table.get("header_row", 5) + 1)
            columns = data_table.get("columns", [])

            if is_com and len(data_table_rows) > 1:
                for i in range(1, len(data_table_rows)):
                    ws.Rows(header_row).Insert()
            elif not is_com and len(data_table_rows) > 1:
                ws.insert_rows(header_row, len(data_table_rows) - 1)

            for i, row_data in enumerate(data_table_rows):
                current_row = header_row + i
                for col_def in columns:
                    col_letter = col_def.get("col_letter", "")
                    if not col_letter:
                        continue
                    value = row_data.get(col_letter, row_data.get(col_def.get("header_text", ""), ""))
                    try:
                        if is_com:
                            ws.Range(f"{col_letter}{current_row}").Value = value
                        else:
                            _write_cell_ol(ws, f"{col_letter}{current_row}", value)
                    except Exception:
                        pass

        # 3. 应用逻辑规则
        logic_rules = template_def.get("logic_rules", [])
        for rule in logic_rules:
            if rule.get("confidence") != "high":
                continue
            _apply_logic_rule(ws, rule, data_table, data_table_rows, is_com)

        # 4. 另存
        win32_helper.excel_save_as(wb, output_path)

    return output_path


def _write_cell_ol(ws, cell_ref: str, value: Any) -> None:
    """openpyxl 模式写单元格。"""
    from openpyxl.utils import column_index_from_string
    col_str = "".join(c for c in cell_ref if c.isalpha())
    row_str = "".join(c for c in cell_ref if c.isdigit())
    if col_str and row_str:
        ws.cell(row=int(row_str), column=column_index_from_string(col_str), value=value)


def _apply_logic_rule(
    ws,
    rule: dict[str, Any],
    data_table: dict[str, Any],
    data_rows: list[dict[str, Any]],
    is_com: bool,
) -> None:
    """应用逻辑规则（如自动判定合格/不合格）。"""
    applies_to = rule.get("applies_to", "")
    if not applies_to:
        return

    depends_on = rule.get("depends_on", [])
    if len(depends_on) < 2:
        return

    dep_cols = []
    for dep in depends_on:
        parts = dep.split(".")
        if len(parts) >= 2:
            col_letter = parts[-1].replace("col_letter_", "")
            dep_cols.append(col_letter)

    target_col = applies_to.replace("data_table.", "").replace("col_letter_", "")
    header_row = data_table.get("data_start_row", data_table.get("header_row", 5) + 1)

    for i, row_data in enumerate(data_rows):
        current_row = header_row + i
        dep_values = [row_data.get(c, "") for c in dep_cols]

        try:
            measured = float(str(dep_values[0]))
            required = float(str(dep_values[1]))
            conclusion = "合格" if measured >= required else "不合格"
        except (ValueError, TypeError):
            continue

        try:
            cell_ref = f"{target_col}{current_row}"
            if is_com:
                ws.Range(cell_ref).Value = conclusion
            else:
                _write_cell_ol(ws, cell_ref, conclusion)
        except Exception:
            pass


# ── 辅助函数 ────────────────────────────────────────────────────

def _describe_template_for_extraction(tmpl_def: dict[str, Any]) -> str:
    """将模板定义转为 LLM 可读的数据提取说明。"""
    lines = [f"模板名称: {tmpl_def.get('template_name', '')}"]
    lines.append(f"报告类型: {tmpl_def.get('report_type', '')}")
    lines.append("")

    lines.append("## 需要填充的区域:")
    for region in tmpl_def.get("regions", []):
        if region.get("type") in ("data_table_header", "data_table_body"):
            continue
        lines.append(f"- [{region.get('id', '')}] {region.get('label', '')}")
        lines.append(f"  位置: {region.get('cells', '')}")
        lines.append(f"  作用: {region.get('purpose', '')}")
        lines.append(f"  数据来源: {region.get('data_source_detail', region.get('data_source', ''))}")
        if region.get("fixed_value"):
            lines.append(f"  固定值: {region['fixed_value']}")
        lines.append("")

    data_table = tmpl_def.get("data_table", {})
    if data_table:
        lines.append("## 数据表结构:")
        lines.append(f"表头行: 第{data_table.get('header_row', '?')}行")
        lines.append(f"数据起始行: 第{data_table.get('data_start_row', '?')}行")
        lines.append("列映射:")
        for col in data_table.get("columns", []):
            lines.append(f"  - 列{col.get('col_letter', '')}: {col.get('header_text', '')} → {col.get('data_source', '')}")
        if data_table.get("expand_logic"):
            lines.append(f"展开逻辑: {data_table['expand_logic']}")
        lines.append("")

    logic_rules = tmpl_def.get("logic_rules", [])
    if logic_rules:
        lines.append("## 逻辑规则:")
        for rule in logic_rules:
            lines.append(f"- {rule.get('description', '')}: {rule.get('expression', '')}")
            lines.append(f"  应用于: {rule.get('applies_to', '')}")
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
        for region in tmpl_def.get("regions", []):
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
    """读取 xlsx 文件并返回 Markdown 预览。支持 COM 和 openpyxl 两种模式。"""
    from .. import win32_helper
    try:
        with win32_helper.open_excel(path) as (app, wb, app_name):
            is_com = win32_helper.WIN32_AVAILABLE and app is not None
            lines = [f"# 实验报告: {path.name}", ""]

            if is_com:
                ws = wb.Worksheets(1)
                used = ws.UsedRange
                max_row = min(used.Rows.Count, 100)
                max_col = min(used.Columns.Count, 30)
                for row_idx in range(1, max_row + 1):
                    cells = []
                    for col_idx in range(1, max_col + 1):
                        value = ws.Cells(row_idx, col_idx).Value
                        cells.append(str(value) if value is not None else "")
                    if any(cells):
                        lines.append(" | ".join(cells))
            else:
                ws = wb.active
                for row in ws.iter_rows(values_only=True, max_row=100):
                    cells = [str(c) if c is not None else "" for c in row]
                    if any(cells):
                        lines.append(" | ".join(cells))
            return "\n".join(lines)
    except Exception as e:
        return f"无法读取 xlsx: {e}"


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
    return {"report_data": {}, "data_table_rows": [],
            "notes": f"LLM 返回格式异常: {content[:300]}"}
