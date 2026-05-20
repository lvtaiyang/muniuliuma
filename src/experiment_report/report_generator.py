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

DATA_COMPLETION_PROMPT = """你是工程实验检测报告数据专家。根据数据映射表和台账数据，逐单元格生成完整的填充数据。

## 重要：模板中的数字只是示例

模板 xlsx 中的数值（如 0.98、98.1%）只是展示格式的占位数据，不是真实检测结果。
你需要根据映射表中每个单元格的数据类型，从台账中提取真实数据填入。

## 数据来源优先级

1. **台账提取**：台账中有对应列的数据 → 直接提取
2. **计算得出**：映射表标注为 calculated 的 → 根据 logic 列的计算公式，用台账数据真算
3. **固定值**：映射表标注为 fixed_value 的 → 直接使用固定值
4. **LLM 生成**：映射表标注为 generated/manual_input 的 → 根据行业规范生成（检测结论、备注等）
5. **无法确定**：台账没有、无法计算、也无法生成的 → 填 "见原始记录"，记录到 missing_data

## 数据行规则

- 台账中每个匹配的样品/路段 → 数据表一行
- 如果台账没有实测数据，按台账的设计指标填入标准要求列，实测值列填"待检测"

## 返回格式

{
  "report_data": { "region_id_1": "值", ...每个region_id都有值... },
  "data_table_rows": [
    {"col_A": "值1", "col_B": "值2", ...所有列都有值...},
    ...
  ],
  "statistics": { "total_points": N, "qualified_points": N, "qualification_rate": "百分比", "average_value": N, "representative_value": N },
  "conclusion": "完整规范的检测结论",
  "missing_data": ["无法确定的字段"],
  "notes": "数据来源说明"
}

## 警告
- 不要从模板中提取数值，模板数值只是示例
- 数据必须从台账提取或根据台账计算
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
        return {"error": f"模板定义不存在: {template_name}。"
                "请改用纯 I/O 路径: read_template_cells 读模板 → 自己分析 → fill_template_cells 写入。"}

    if not tmpl_def.get("confirmed"):
        remaining = tmpl_def.get("confirmation_required", "?")
        return {
            "error": f"模板定义尚未确认（还有 {remaining} 个待确认问题）。"
                     f"请改用纯 I/O 路径: read_template_cells 读模板 → 自己分析 → fill_template_cells 写入。",
            "uncertainties": tmpl_def.get("uncertainties", []),
        }

    # 2. 读取台账
    ledgers = ledger_reader.read_ledgers(ledger_paths)
    ledger_text = "\n\n".join(ledger_reader.format_ledger_for_llm(l) for l in ledgers)
    if not ledger_text.strip():
        return {"error": "台账文件为空或无法读取"}

    # 3. 构建数据映射表：每个需填充的单元格 + 类型 + 来源 + 逻辑
    data_mapping = _build_data_mapping(tmpl_def)

    # 4. LLM 数据补全：映射表 + 台账 → 完整报告数据
    conf = cfg.load()
    llm = cfg.get_llm_config("text")
    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=300)

    row_estimate = max(5, min(80, len(ledger_text.split("\n")) // 2))
    max_tok = max(4000, row_estimate * 400)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": DATA_COMPLETION_PROMPT},
            {"role": "user", "content": (
                f"## 数据映射表（模板中所有需填充的单元格）\n\n{data_mapping}\n\n"
                f"## 台账数据（真实数据源）\n\n{ledger_text[:15000]}\n\n"
                f"## 生成要求\n"
                f"- 报告日期: {report_date}\n"
                f"- 项目名称: {project_name}\n"
                f"- 模板中的数值只是示例,不要提取,真实数据必须从台账来\n"
                f"- 数据表按台账中匹配的行数展开\n"
                f"- 统计值根据数据行实际计算\n"
                f"- 检测结论要完整规范\n"
            )},
        ],
        max_tokens=max_tok,
        temperature=0.2,
    )

    extracted = _parse_response(response.choices[0].message.content or "")

    # 4.5 完整性校验：检查映射表中的 region 是否都被填充
    expected_regions = _collect_region_ids(tmpl_def)
    filled_regions = set(extracted.get("report_data", {}).keys())
    missing = expected_regions - filled_regions
    if missing:
        # 补填缺失的 region
        for rid in missing:
            val = _infer_default_value(rid, tmpl_def)
            if val:
                extracted.setdefault("report_data", {})[rid] = val

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

    data_mapping = _build_data_mapping(tmpl_def)

    conf = cfg.load()
    llm = cfg.get_llm_config("text")
    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=300)

    ctx = f"\n\n额外说明: {extra_context}" if extra_context else ""
    max_tok = max(4000, max(5, min(80, len(ledger_text.split("\n")) // 2)) * 400)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": DATA_COMPLETION_PROMPT},
            {"role": "user", "content": (
                f"只提取检测项目为「{test_item}」的数据。{ctx}\n\n"
                f"## 数据映射表\n\n{data_mapping}\n\n"
                f"## 台账数据\n\n{ledger_text[:15000]}\n\n"
                f"报告日期: {report_date}"
            )},
        ],
        max_tokens=max_tok,
        temperature=0.2,
    )

    extracted = _parse_response(response.choices[0].message.content or "")

    # 完整性校验
    expected_regions = _collect_region_ids(tmpl_def)
    filled_regions = set(extracted.get("report_data", {}).keys())
    for rid in expected_regions - filled_regions:
        val = _infer_default_value(rid, tmpl_def)
        if val:
            extracted.setdefault("report_data", {})[rid] = val

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

def _build_data_mapping(tmpl_def: dict[str, Any]) -> str:
    """构建完整的数据映射表：列出模板中每一个需要填充的单元格。

    输出格式为 Markdown 表格，每行一个填充项，LLM 可以直接对照填充。
    """
    lines = [
        f"模板: {tmpl_def.get('template_name', '')} | 类型: {tmpl_def.get('report_type', '')}",
        f"工作表数: {len(tmpl_def.get('sheets', []))}", "", "",
    ]

    for sheet in tmpl_def.get("sheets", []):
        sname = sheet.get("sheet_name", "")
        srole = sheet.get("sheet_role", "")
        lines.append(f"## 工作表: {sname} (角色: {srole})")
        lines.append(f"说明: {sheet.get('sheet_description', '')}")
        lines.append("")

        # ── 固定区域映射表 ──
        regions = sheet.get("regions", [])
        non_table_regions = [r for r in regions if r.get("type") not in ("data_table_header", "data_table_body")]
        if non_table_regions:
            lines.append("### 固定区域填充映射")
            lines.append("| region_id | 位置 | 作用 | 数据类型 | 数据来源/固定值 | 逻辑 |")
            lines.append("|-----------|------|------|---------|----------------|------|")
            for r in non_table_regions:
                rid = r.get("id", "")
                cells = r.get("cells", "")
                purpose = r.get("purpose", "")
                ds = r.get("data_source", "")
                ds_detail = r.get("data_source_detail", "")
                fixed = r.get("fixed_value", "")
                logic = r.get("logic", "")

                if ds == "fixed_value" or fixed:
                    dtype = "固定值"
                    source = fixed or ds_detail
                elif ds in ("calculated",):
                    dtype = "计算值"
                    source = ds_detail
                elif ds in ("manual_input",):
                    dtype = "需生成"
                    source = ds_detail
                elif "ledger" in ds:
                    dtype = "台账提取"
                    source = ds_detail
                elif ds in ("report_title", "project_name"):
                    dtype = "固定值"
                    source = ds
                else:
                    dtype = "台账提取/生成"
                    source = ds_detail or ds

                lines.append(f"| {rid} | {cells} | {purpose} | {dtype} | {source} | {logic} |")
            lines.append("")

        # ── 数据表映射 ──
        dt = sheet.get("data_table", {})
        if dt:
            columns = dt.get("columns", [])
            if columns:
                lines.append("### 数据表列映射")
                lines.append("| 列 | 表头 | 数据类型 | 数据来源 |")
                lines.append("|-----|------|---------|---------|")
                for col in columns:
                    cl = col.get("col_letter", "")
                    hdr = col.get("header_text", "")
                    cds = col.get("data_source", "")
                    lines.append(f"| {cl} | {hdr} | 台账提取 | {cds} |")
                lines.append("")
                lines.append(f"**数据表位置:** 表头第{dt.get('header_row', '?')}行, 数据从第{dt.get('data_start_row', '?')}行开始")
                lines.append(f"**展开规则:** {dt.get('expand_logic', '每个样品一行')}")
                lines.append("")

    # ── 跨表数据流 ──
    cross = tmpl_def.get("cross_sheet_logic", [])
    if cross:
        lines.append("## 跨工作表数据流")
        for c in cross:
            lines.append(f"- {c.get('description', '')}: "
                         f"{c.get('source_sheet', '')} 的 {c.get('source_cells', '')} "
                         f"→ {c.get('target_sheet', '')} 的 {c.get('target_cells', '')} "
                         f"({c.get('logic', '')})")
        lines.append("")

    # ── 逻辑规则 ──
    rules = tmpl_def.get("logic_rules", [])
    if rules:
        lines.append("## 判定与计算规则")
        for rule in rules:
            lines.append(f"- [{rule.get('rule_id', '')}] {rule.get('description', '')}: "
                         f"`{rule.get('expression', '')}` "
                         f"→ 填入 {rule.get('applies_in_sheet', '')} 的 {rule.get('applies_to', '')}")
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


def _collect_region_ids(tmpl_def: dict[str, Any]) -> set[str]:
    """收集模板中所有需要填充的 region_id。"""
    ids = set()
    for sheet in tmpl_def.get("sheets", []):
        for r in sheet.get("regions", []):
            if r.get("type") not in ("data_table_header", "data_table_body"):
                ids.add(r.get("id", ""))
    ids.discard("")
    return ids


def _infer_default_value(rid: str, tmpl_def: dict[str, Any]) -> str | None:
    """对 LLM 遗漏的 region，尝试从模板定义推断默认值。"""
    for sheet in tmpl_def.get("sheets", []):
        for r in sheet.get("regions", []):
            if r.get("id") == rid:
                if r.get("fixed_value"):
                    return r["fixed_value"]
                if r.get("data_source") == "fixed_value":
                    return r.get("data_source_detail", "")
    return "见原始记录"


def _parse_response(content: str) -> dict[str, Any]:
    from .. import json_utils
    result = json_utils.parse_llm_json(content)
    if "_raw" in result:
        return {"report_data": {}, "data_table_rows": [],
                "notes": f"LLM 返回格式异常: {content[:300]}"}
    return result
