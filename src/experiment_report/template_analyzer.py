"""实验报告 xlsx 模板分析：用 LLM 逐单元格理解模板结构和数据逻辑。

核心理念：
- 读取出 xlsx 的完整结构（单元格位置、内容、合并区域、格式）
- 送 LLM 逐单元格/区域分析：每个格子填什么、数据从哪来、跟其他格子有什么逻辑关系
- LLM 不确定的地方 → 生成问题让用户确认
- 确认后的模板定义保存下来，供填充使用
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from openai import OpenAI
from .. import config as cfg

# ── LLM Prompt ──────────────────────────────────────────────────

TEMPLATE_ANALYSIS_SYSTEM_PROMPT = """你是工程建设领域实验检测报告模板的分析专家。你的任务是深度理解一份实验报告 xlsx 模板，逐单元格分析其含义和数据逻辑。

## 背景知识

工程实验检测报告通常包含以下信息：
- 报告标题（如"钢筋原材力学性能检测报告"）
- 工程基本信息：工程名称、施工单位、监理单位、检测单位
- 样品信息：样品编号、样品名称、规格型号、取样部位、取样日期
- 检测信息：检测项目、检测依据（标准编号）、检测日期、检测环境（温度、湿度）
- 检测数据表：样品编号、检测参数、标准要求值、实测值、单项判定
- 综合结论
- 签章区：检测人员、审核人、批准人、日期
- 备注

## 你的任务

仔细阅读提供的 xlsx 模板结构（每个单元格的位置、内容、合并区域）。对模板做逐区域分析：

1. **识别每个有意义的单元格/区域**：它在报告中扮演什么角色？
2. **确定数据来源**：该格子的数据从台账的哪个字段来？是固定值？是计算得出？还是需要人工填写？
3. **发现逻辑关系**：哪些格子的值依赖于其他格子？（如"单项判定 = IF(实测值 符合 标准要求, '合格', '不合格')"）
4. **标记不确定的地方**：对任何不能 100% 确定的内容，生成一个明确的、可回答的问题

## 关键原则

- **宁可多问，不可猜错**。如果你对某个格子的含义、数据来源、判断逻辑不确定，必须生成问题。
- 每个问题必须具体、可回答，不要问模糊的问题。
- 对于数据表区域，要说明：表头各列含义、数据从台账哪些列映射、是否需要按样品展开多行
- 如果模板中有公式或计算逻辑的痕迹，要明确指出

## 返回格式

你必须严格返回以下 JSON：

{
  "template_name": "简短准确的模板名称",
  "description": "模板用途描述",
  "report_type": "检测报告类型，如 材料检测/现场检测/见证取样/平行检测",
  "regions": [
    {
      "id": "region_1",
      "cells": "A1:F1 或 A5",
      "type": "title|project_info|sample_info|test_info|data_table_header|data_table_body|conclusion|signature|notes|other",
      "label": "该区域的显示标签/内容",
      "purpose": "该区域在报告中的作用",
      "data_source": "report_title|project_name|ledger.sample_no|ledger.test_parameter|ledger.required_value|ledger.measured_value|calculated|fixed_value|manual_input",
      "data_source_detail": "具体的数据来源说明，如'台账的样品编号列'或'计算：实测值 vs 标准要求'",
      "fixed_value": "如果是固定值，这里填写固定内容",
      "logic": "如果有计算/判断逻辑，这里描述（如 IF measured_value >= required_value THEN '合格' ELSE '不合格'）",
      "confidence": "high|medium|low",
      "question": "如果不确定，生成一个具体问题（confidence为high时可为空字符串）"
    }
  ],
  "data_table": {
    "header_row": 5,
    "data_start_row": 6,
    "columns": [
      {
        "col_letter": "A",
        "header_text": "样品编号",
        "data_source": "ledger.sample_no",
        "confidence": "high",
        "question": ""
      }
    ],
    "expand_direction": "down",
    "expand_logic": "每个样品一行，行数=台账中该检测项目的样品数"
  },
  "logic_rules": [
    {
      "rule_id": "rule_1",
      "description": "单项判定逻辑",
      "expression": "IF measured_value meets required_value THEN '合格' ELSE '不合格'",
      "depends_on": ["data_table.col_letter_D", "data_table.col_letter_E"],
      "applies_to": "data_table.col_letter_F",
      "confidence": "high",
      "question": ""
    }
  ],
  "uncertainties": [
    {
      "id": "q1",
      "region_id": "region_X",
      "question": "具体的问题描述",
      "context": "为什么不确定（给出上下文帮助用户理解）",
      "suggested_answer": "你建议的答案（可选）"
    }
  ],
  "notes": "任何补充说明"
}
"""


def analyze_template(file_path: str | Path) -> dict[str, Any]:
    """分析实验报告 xlsx 模板。

    Args:
        file_path: xlsx 模板文件路径

    Returns:
        结构化分析结果，包含 regions, data_table, logic_rules, uncertainties
    """
    file_path = Path(file_path)
    if not file_path.exists():
        return {"error": f"文件不存在: {file_path}"}
    if file_path.suffix.lower() not in (".xlsx", ".xlsm"):
        return {"error": f"不支持的文件格式: {file_path.suffix}，实验报告模板必须是 .xlsx"}

    conf = cfg.load()
    llm = cfg.get_llm_config("text")
    if not llm.get("api_key"):
        return {"error": "未配置 LLM API key，请先 setup_monitoring"}

    # 读取 xlsx 完整结构
    xlsx_structure = _read_xlsx_structure(file_path)

    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=180)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": TEMPLATE_ANALYSIS_SYSTEM_PROMPT},
            {"role": "user", "content": _build_analysis_prompt(file_path.name, xlsx_structure)},
        ],
        max_tokens=4000,
        temperature=0.1,
    )

    result = _parse_response(response.choices[0].message.content or "")
    result["source_file"] = str(file_path)
    result["analyzed_at"] = _now_iso()

    # 检查是否有需要确认的问题
    has_uncertainties = bool(result.get("uncertainties"))
    low_confidence_regions = [
        r for r in result.get("regions", [])
        if r.get("confidence") in ("medium", "low") and r.get("question")
    ]
    result["needs_confirmation"] = has_uncertainties or bool(low_confidence_regions)
    result["confirmation_required"] = (
        len(result.get("uncertainties", [])) + len(low_confidence_regions)
    )

    return result


def save_template_definition(
    project_name: str,
    template_definition: dict[str, Any],
    filename: str = "",
) -> dict[str, Any]:
    """将确认后的模板定义保存到项目中。

    Args:
        project_name: 项目名称
        template_definition: 确认后的模板定义 JSON
        filename: 保存的文件名（可选）
    """
    template_dir = _get_template_dir(project_name)
    if not template_dir:
        return {"error": f"项目不存在: {project_name}"}

    name = filename or f"{template_definition.get('template_name', 'experiment_template')}.json"
    if not name.endswith(".json"):
        name += ".json"

    template_definition["saved_at"] = _now_iso()
    template_definition["project_name"] = project_name

    path = template_dir / name
    path.write_text(
        json.dumps(template_definition, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return {
        "saved": str(path),
        "template_name": template_definition.get("template_name", ""),
    }


def load_template_definition(project_name: str, template_name: str) -> dict[str, Any] | None:
    """加载已保存的模板定义。"""
    template_dir = _get_template_dir(project_name)
    if not template_dir:
        return None
    path = template_dir / template_name
    if not path.exists():
        # 尝试加 .json 后缀
        path = template_dir / f"{template_name}.json"
        if not path.exists():
            return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, IOError):
        return None


def confirm_template(
    analysis_result: dict[str, Any],
    answers_json: str,
) -> dict[str, Any]:
    """将用户对不确定问题的回答合并到模板分析结果中。

    Args:
        analysis_result: analyze_template 返回的结果
        answers_json: 用户回答的 JSON 字符串，格式:
            [{"question_id": "q1", "answer": "..."}, ...]
            或 {"region_id": "region_X", "resolution": "..."}

    Returns:
        合并后的完整模板定义（uncertainties 已清除，confidence 已更新）
    """
    try:
        answers = json.loads(answers_json)
    except json.JSONDecodeError:
        return {"error": "answers_json 不是有效 JSON"}

    if not isinstance(answers, list):
        answers = [answers]

    # 建立答案索引
    answer_map: dict[str, str] = {}
    for a in answers:
        qid = a.get("question_id") or a.get("id") or a.get("region_id") or ""
        ans = a.get("answer") or a.get("resolution") or ""
        if qid:
            answer_map[qid] = ans

    # 处理 uncertainties 列表
    resolved_uncertainties = []
    for u in analysis_result.get("uncertainties", []):
        uid = u.get("id", "")
        if uid in answer_map:
            u["resolution"] = answer_map[uid]
            u["resolved"] = True
            resolved_uncertainties.append(u)
        else:
            # 未回答的保留
            u["resolved"] = False

    # 处理 regions 中的 question
    for region in analysis_result.get("regions", []):
        rid = region.get("id", "")
        if rid in answer_map:
            region["resolution"] = answer_map[rid]
            region["confidence"] = "high"
            region["question"] = ""

    # 处理 logic_rules 中的 question
    for rule in analysis_result.get("logic_rules", []):
        rid = rule.get("rule_id", "")
        if rid in answer_map:
            rule["resolution"] = answer_map[rid]
            rule["confidence"] = "high"
            rule["question"] = ""

    # 清除已解决的 uncertainties
    remaining = [u for u in resolved_uncertainties if not u.get("resolved")]
    analysis_result["uncertainties"] = remaining
    analysis_result["needs_confirmation"] = len(remaining) > 0
    analysis_result["confirmation_required"] = len(remaining)
    analysis_result["confirmed"] = len(remaining) == 0
    analysis_result["confirmed_at"] = _now_iso()

    return analysis_result


def list_template_definitions(project_name: str) -> list[dict[str, Any]]:
    """列出项目已保存的模板定义。"""
    template_dir = _get_template_dir(project_name)
    if not template_dir or not template_dir.exists():
        return []
    templates = []
    for f in sorted(template_dir.glob("*.json")):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            templates.append({
                "filename": f.name,
                "template_name": data.get("template_name", f.stem),
                "report_type": data.get("report_type", ""),
                "saved_at": data.get("saved_at", ""),
                "path": str(f),
            })
        except (json.JSONDecodeError, IOError):
            templates.append({"filename": f.name, "template_name": f.stem, "error": "无法读取"})
    return templates


# ── 内部函数 ────────────────────────────────────────────────────

def _read_xlsx_structure(file_path: Path) -> dict[str, Any]:
    """通过 win32com Excel/WPS 读取 xlsx 完整结构信息。

    包括单元格位置、内容、合并区域。完美读取所有格式信息。
    """
    from .. import win32_helper

    if not win32_helper.WIN32_AVAILABLE:
        raise RuntimeError("win32com 不可用，请在 Windows 上运行")

    return win32_helper.excel_read_structure(file_path)


def _build_analysis_prompt(filename: str, structure: dict[str, Any]) -> str:
    """构建发给 LLM 的分析请求。"""
    parts = [f"请分析以下实验报告模板: {filename}", ""]

    for sheet in structure.get("sheets", []):
        parts.append(f"## 工作表: {sheet['name']}")
        parts.append(f"尺寸: {sheet.get('dimensions', '')}")
        parts.append(f"合并单元格: {len(sheet.get('merged_detail', []))} 处")
        for m in sheet.get("merged_detail", []):
            parts.append(f"  - {m['range']} (从 {m['top_left']} 到 {m['bottom_right']})")
        parts.append("")

        # 用网格形式展示有内容的单元格
        parts.append("### 逐单元格内容（含精确位置）:")
        parts.append("")
        for row_info in sheet.get("rows", []):
            row_num = row_info["row"]
            for c in row_info["cells"]:
                parts.append(f"  {c['cell']}: \"{c['value']}\"")
        parts.append("")

    parts.append("请逐区域分析这个模板，特别注意：")
    parts.append("1. 数据表中每列与台账字段的对应关系")
    parts.append("2. 单元格间的计算/判定逻辑")
    parts.append("3. 任何你不确定的地方，请在 uncertainties 中生成具体问题")

    return "\n".join(parts)


def _get_template_dir(project_name: str) -> Path | None:
    """获取项目模板存储目录。"""
    try:
        from ..project_initializer import project_manager
        proj = project_manager.load(project_name)
        if proj:
            template_dir = Path(proj["workspace"]) / "04_施工实施" / "实验检测报告" / "_模板定义"
            template_dir.mkdir(parents=True, exist_ok=True)
            return template_dir
    except ImportError:
        pass
    return None


def _parse_response(content: str) -> dict[str, Any]:
    """解析 LLM 返回的 JSON。"""
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
    return {
        "raw_response": content,
        "template_name": "解析失败",
        "regions": [],
        "uncertainties": [{"id": "q0", "question": "LLM 返回格式异常，请检查原始输出",
                           "context": content[:500]}],
        "needs_confirmation": True,
        "confirmation_required": 1,
    }


def _now_iso() -> str:
    from datetime import datetime
    return datetime.now().isoformat()
