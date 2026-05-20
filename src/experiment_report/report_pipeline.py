"""实验报告完整生成管线 — 单步工具，锁死流程，不依赖智能体能力。

流程:
  ① 读模板所有单元格
  ② 读台账所有数据
  ③ 把模板+台账送给 LLM，要求逐单元格返回填充值
  ④ COM 逐单元格写入
  ⑤ 输出报告 + 缺失数据清单

智能体只需调用 generate_complete_report，传模板路径+台账路径+输出路径。
"""

from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path
from typing import Any

from openai import OpenAI
from .. import config as cfg
from .. import win32_helper

PIPELINE_PROMPT = """你是工程实验检测报告生成专家。现在要填写一份实验报告。

## 数据源

1. **模板原始数据** — xlsx 中当前的所有单元格和值。数值只是示例，不要当作真实数据提取
2. **台账数据** — 检测计划/路段信息（桩号、部位、设计指标等元数据）

## 你的任务

逐单元格决定填充内容。每个单元格只属于以下5种情况之一：

### 情况1: 固定文本
标签、标题、单位名称、标准号、签名行标签等 → **原样保留**

### 情况2: 台账有对应的数据
桩号、部位、设计指标等 → **从台账精确提取**

### 情况3: 需要计算
平均值、合格率、代表值等 → **严格计算**
- 例如: 平均值 = sum(所有实测值) / count
- 例如: 合格率 = 合格点数 / 总点数 × 100%

### 情况4: 台账缺失
台账没有对应数据 → **填入 "见原始记录"**
**绝不保留模板示例值**。模板中的 0.98, 98.1%, K9+370.58 等都是示例，必须替换。

### 情况5: 需生成文本
检测结论、备注等 → **根据行业规范生成**
- 检测结论格式: "经检测，该[层次]的压实度检测[N]点，检测值为：[列出各值]，检测结果[合格/不合格]"

## 返回格式

{
  "cell_values": {
    "工作表名": {
      "A1": "保留原值或填充新值",
      "B3": "新值",
      ...
    },
    ...
  },
  "unchanged": ["未修改的单元格说明"],
  "missing": ["台账缺失无法填充的项目"],
  "summary": "填充总结"
}

## 关键规则
- **模板中的 0.98, 98.1%, K9+370.58 等数字必须被替换，不能保留**
- 对每个有值的单元格做判断：是固定文本(保留)还是示例数据(替换)
- 台账有多行匹配时，数据表展开对应行数
- 实在无法确定的填 "见原始记录"
- 不要编造数据
"""


def generate_complete_report(
    template_path: str | Path,
    ledger_paths: list[str],
    output_dir: str | Path,
    report_date: str = "",
) -> dict[str, Any]:
    """一步生成完整实验报告。

    Args:
        template_path: xlsx 模板路径
        ledger_paths: 台账文件路径列表
        output_dir: 输出目录
        report_date: 报告日期，默认今天
    """
    template_path = Path(template_path)
    output_dir = Path(output_dir)

    if not template_path.exists():
        return {"error": f"模板不存在: {template_path}"}

    # ── 步骤 1: 读模板 ──
    tmpl_structure = win32_helper.excel_read_structure(template_path)
    if "error" in tmpl_structure:
        return {"error": f"读取模板失败: {tmpl_structure['error']}"}

    # ── 步骤 2: 读台账 ──
    from ..construction_log import ledger_reader
    ledgers = ledger_reader.read_ledgers(ledger_paths)
    ledger_text = "\n\n".join(ledger_reader.format_ledger_for_llm(l) for l in ledgers)
    if not ledger_text.strip():
        return {"error": "台账文件为空或无法读取"}

    # ── 步骤 3: 构建上下文 ──
    # 提取模板所有单元格（只提取前2个sheet，第3个是参考表）
    cells_text = []
    for sheet in tmpl_structure.get("sheets", [])[:2]:
        cells_text.append(f"## 工作表: {sheet['name']}")
        for row_info in sheet.get("rows", []):
            for c in row_info["cells"]:
                cells_text.append(f"  {c['cell']}: \"{c['value']}\"")
        cells_text.append("")
    template_cell_text = "\n".join(cells_text)

    if not report_date:
        report_date = date.today().isoformat()

    # ── 步骤 4: LLM 分析 ──
    llm = cfg.get_llm_config("text")
    if not llm.get("api_key"):
        return {"error": "需要 LLM text API Key。请在 config.yaml 中配置。"}

    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=300)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": PIPELINE_PROMPT},
            {"role": "user", "content": (
                f"## 模板当前单元格数据\n\n{template_cell_text}\n\n"
                f"## 台账数据\n\n{ledger_text[:15000]}\n\n"
                f"## 生成要求\n"
                f"- 报告日期: {report_date}\n"
                f"- 模板中所有数值(0.98, 98.1%, K9+370.58等)是示例，必须替换\n"
                f"- 台账有的提取，台账缺的填\"见原始记录\"\n"
                f"- 检测结论要完整规范\n"
            )},
        ],
        max_tokens=8000,
        temperature=0.1,
    )

    llm_result = _parse_pipeline_response(response.choices[0].message.content or "")
    cell_values = llm_result.get("cell_values", {})

    if not cell_values:
        return {"error": "LLM 未返回有效的填充数据", "raw": str(llm_result)[:500]}

    # ── 步骤 5: COM 写入 ──
    output_dir.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', template_path.stem)[:30]
    output_path = output_dir / f"{report_date}_{safe_name}.xlsx"

    try:
        fill_data = _cell_values_to_fill_data(cell_values)
        win32_helper.excel_fill_template(template_path, output_path, fill_data)
    except Exception as e:
        return {"error": f"COM 写入失败: {e}", "llm_result": llm_result}

    # ── 步骤 6: 校验（检查示例值是否被替换）─
    sample_values = ["0.98", "0.981", "98%", "98.1%"]
    validation = _validate_output(output_path, sample_values)

    return {
        "saved": str(output_path),
        "cells_filled": sum(len(v) for v in cell_values.values()),
        "unchanged": llm_result.get("unchanged", []),
        "missing": llm_result.get("missing", []),
        "summary": llm_result.get("summary", ""),
        "validation": validation,
        "warning": "仍有示例值残留，请检查" if validation["sample_values_found"] else None,
    }


def _cell_values_to_fill_data(cell_values: dict[str, dict[str, str]]) -> dict[str, Any]:
    """将 {sheet: {cell: value}} 转为 fill_data 格式。"""
    sheets = []
    for sname, cells in cell_values.items():
        if not isinstance(cells, dict):
            continue
        regions = []
        for cr, val in cells.items():
            regions.append({
                "id": f"cell_{cr}", "cells": cr,
                "type": "other", "data_source": "fixed_value",
                "fixed_value": str(val), "confidence": "high",
            })
        sheets.append({
            "sheet_name": sname, "sheet_role": "main_report",
            "regions": regions, "data_table": {},
        })
    return {"sheets": sheets, "report_data": {}, "data_table_rows": [], "logic_rules": []}


def _validate_output(output_path: Path, sample_values: list[str]) -> dict[str, Any]:
    """校验输出文件：检查示例值是否被替换。"""
    try:
        structure = win32_helper.excel_read_structure(output_path)
        found = []
        for sheet in structure.get("sheets", []):
            for row_info in sheet.get("rows", []):
                for c in row_info["cells"]:
                    val = c["value"]
                    for sv in sample_values:
                        if sv in val:
                            found.append({"cell": c["cell"], "value": val, "matched_sample": sv})
        return {
            "sample_values_found": len(found),
            "details": found[:10],
            "ok": len(found) == 0,
        }
    except Exception as e:
        return {"error": str(e)}


def _parse_pipeline_response(content: str) -> dict[str, Any]:
    from .. import json_utils
    result = json_utils.parse_llm_json(content)
    if "_raw" in result:
        return {"cell_values": {}, "error": str(result.get("_raw", ""))[:500]}
    return result
