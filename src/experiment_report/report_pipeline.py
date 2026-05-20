"""实验报告完整生成管线 — 单步工具，锁死流程。

适用任何类型的检测报告(压实度/材料/现场检测等)。
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

PIPELINE_PROMPT = """你是工程实验检测报告生成专家。根据台账数据，生成一份完整可用的检测报告。

## 你的任务

### 1. 理解报告类型
看模板内容判断这是什么类型的检测报告(压实度/材料/强度/...)。
根据报告类型，确定检测项目的行业标准和合格判定规则。

### 2. 区分固定内容 vs 需填充内容
- **固定文本**: 标题、单位名、标准号、签名行标签、表格列头等 → 原样保留
- **需填充**: 工程部位、检测日期、检测数据、统计值、结论 → 必须填充新值
- **模板示例数值**: 0.98、98.1%、K9+370.58 等 → 必须替换，绝不保留

### 3. 从台账提取数据
- 工程部位/桩号范围 → 台账的"桩号及部位"或"桩号"列
- 设计指标/标准值 → 台账的"设计指标"列
- 施工日期/检测日期 → 台账的"施工日期"/"检测日期"列
- 厚度/层次 → 台账对应列

### 4. 生成检测点数据(台账没有实测值时)
根据设计指标和行业常识，为数据表生成合理的检测点:
- 在设计值附近取2-3个点
- 压实度: 通常大于等于设计值, 范围设计值~设计值+1.5%
- 强度: 通常大于等于设计值, 范围设计值~设计值+20%
- 桩号位置: 在路段范围内均匀分布
- 每个点标注距边距离
- 根据行业标准判定每个点合格/不合格

### 5. 计算统计值
- 检测点数 = 实际生成的行数
- 合格点数 = 判定为合格的行数
- 合格率 = 合格点数/检测点数×100%
- 平均值 = sum(所有实测值)/检测点数
- 根据行业标准计算代表值

### 6. 生成检测结论
完整规范的结论，包含: 检测依据、检测数量、各点实测值、综合判定。

### 7. 填充所有工作表
- 报告正页(Sheet1): 完整填写
- 原始记录表(Sheet2): 如有，也完整填写
- 参考表(Sheet3): 如有检测数据相关，也更新

## 返回格式

{
  "cell_values": {
    "报告正页": {"A3": "值", "B5": "值", "C12": "值", ...},
    "原始记录表": {"A3": "值", ...},
    ...
  },
  "test_points": [{"桩号": "...", "检测值": ..., "判定": "合格"}, ...],
  "summary": "生成了N个检测点的完整报告"
}

## 禁止
- ❌ 保留模板示例数值
- ❌ 填入"见原始记录"
- ❌ 跳过任何工作表
- ❌ 编造台账中已有的信息
"""


def generate_complete_report(
    template_path: str | Path,
    ledger_paths: list[str],
    output_dir: str | Path,
    report_date: str = "",
) -> dict[str, Any]:
    template_path = Path(template_path)
    output_dir = Path(output_dir)
    if not template_path.exists():
        return {"error": f"模板不存在: {template_path}"}

    # 1. 读模板
    tmpl_structure = win32_helper.excel_read_structure(template_path)
    if "error" in tmpl_structure:
        return {"error": f"读取模板失败: {tmpl_structure['error']}"}

    # 2. 读台账
    from ..construction_log import ledger_reader
    ledgers = ledger_reader.read_ledgers(ledger_paths)
    ledger_text = "\n\n".join(ledger_reader.format_ledger_for_llm(l) for l in ledgers)
    if not ledger_text.strip():
        return {"error": "台账文件为空或无法读取"}

    # 3. 构建模板上下文(所有sheet)
    cells_lines = []
    for sheet in tmpl_structure.get("sheets", []):
        cells_lines.append(f"## 工作表: {sheet['name']}")
        for ri in sheet.get("rows", []):
            for c in ri["cells"]:
                cells_lines.append(f"  {c['cell']}: \"{c['value']}\"")
        cells_lines.append("")
    template_text = "\n".join(cells_lines)

    if not report_date:
        report_date = date.today().isoformat()

    # 4. LLM
    llm = cfg.get_llm_config("text")
    if not llm.get("api_key"):
        return {"error": "需要 LLM text API Key。请在 config.yaml 中配置。"}

    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=300)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": PIPELINE_PROMPT},
            {"role": "user", "content": (
                f"## 模板所有单元格\n\n{template_text}\n"
                f"## 台账数据\n\n{ledger_text[:15000]}\n\n"
                f"报告日期: {report_date}\n"
                f"请生成完整的填充数据。模板中所有数值是示例，必须替换。"
                f"所有工作表都要填充。不要留\"见原始记录\"。"
            )},
        ],
        max_tokens=8000,
        temperature=0.1,
    )

    llm_result = _parse(response.choices[0].message.content or "")
    cell_values = llm_result.get("cell_values", {})
    if not cell_values:
        return {"error": "LLM 未返回填充数据", "raw": str(llm_result)[:500]}

    # 5. COM 写入
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r'[\\/:*?"<>|]', '_', template_path.stem)[:30]
    output_path = output_dir / f"{report_date}_{stem}.xlsx"

    try:
        fill_data = _to_fill_data(cell_values)
        win32_helper.excel_fill_template(template_path, output_path, fill_data)
    except Exception as e:
        return {"error": f"COM 写入失败: {e}"}

    # 6. 校验: 对比模板，列出仍相同的单元格
    unchanged = _find_unchanged(template_path, output_path)

    return {
        "saved": str(output_path),
        "cells_written": sum(len(v) for v in cell_values.values() if isinstance(v, dict)),
        "test_points": llm_result.get("test_points", []),
        "summary": llm_result.get("summary", ""),
        "unchanged_from_template": unchanged,
        "ok": len(unchanged) < 5,  # 允许少量固定文本不变
    }


def _to_fill_data(cell_values: dict[str, dict[str, str]]) -> dict[str, Any]:
    sheets = []
    for sname, cells in cell_values.items():
        if not isinstance(cells, dict):
            continue
        regions = [{"id": f"c_{cr}", "cells": cr, "fixed_value": str(v),
                     "data_source": "fixed_value", "confidence": "high"}
                   for cr, v in cells.items()]
        sheets.append({"sheet_name": sname, "sheet_role": "main_report",
                        "regions": regions, "data_table": {}})
    return {"sheets": sheets, "report_data": {}, "data_table_rows": [], "logic_rules": []}


def _find_unchanged(template_path: Path, output_path: Path) -> list[dict[str, str]]:
    """找出输出文件中与模板值完全相同的单元格（可能未填充）。"""
    try:
        tmpl = win32_helper.excel_read_structure(template_path)
        out = win32_helper.excel_read_structure(output_path)
    except Exception:
        return [{"error": "无法读取文件进行校验"}]

    unchanged = []
    for ts, os in zip(tmpl.get("sheets", []), out.get("sheets", [])):
        for tr, or_ in zip(ts.get("rows", []), os.get("rows", [])):
            for tc, oc in zip(tr.get("cells", []), or_.get("cells", [])):
                tv = tc["value"].strip()
                ov = oc["value"].strip()
                if tv == ov and tv and not _is_fixed_label(tv):
                    unchanged.append({
                        "cell": tc["cell"], "sheet": ts["name"],
                        "value": tv[:60],
                    })
    return unchanged[:20]


def _is_fixed_label(text: str) -> bool:
    """判断是否为固定标签文本（不需要变化的）。"""
    fixed_patterns = [
        "试验室名称", "施工单位", "工程名称", "检测依据", "判定依据",
        "检测", "审核", "批准", "日期", "序号", "桩号", "位置",
        "报告编号", "记录编号", "样品名称", "样品编号",
        "试验检测报告", "检测记录表", "允许偏差", "规定极值",
        "检测点数", "合格点数", "合格率", "保证率", "标准差",
        "压实度", "检测结论", "附加声明", "备注", "填料类型",
        "道路等级", "填筑层次", "填筑厚度", "工程部位", "样品信息",
        "主要仪器", "试验条件", "试验检测日期", "最大干密度", "最佳含水率",
        "盒号", "盒重", "层", "次", "项", "目",
    ]
    return any(p in text for p in fixed_patterns)


def _parse(content: str) -> dict[str, Any]:
    from .. import json_utils
    result = json_utils.parse_llm_json(content)
    return result if "_raw" not in result else {"cell_values": {}}
