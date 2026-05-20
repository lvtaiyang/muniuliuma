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
from .. import config as cfg, json_utils

# ── LLM Prompt ──────────────────────────────────────────────────

TEMPLATE_ANALYSIS_SYSTEM_PROMPT = """你是工程建设领域实验检测报告模板的分析专家。你的任务是深度理解一份实验报告 xlsx 模板的**所有工作表**，逐单元格分析其含义和数据逻辑。

## 重要：必须分析所有工作表

模板可能包含多个工作表，你必须**逐个分析每一个工作表**，不能只分析第一个。

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

对**每个工作表**做逐区域分析：

1. **判断工作表角色**：main_report（报告正页）/ raw_data（原始记录）/ reference（参考数据）/ other
2. **识别每个有意义的单元格/区域**：它在报告中扮演什么角色？
3. **确定数据来源**：该格子的数据从台账的哪个字段来？是固定值？是计算得出？还是需要人工填写？
4. **发现逻辑关系**：哪些格子的值依赖于其他格子？哪些数据从 raw_data 表汇总到 main_report 表？
5. **标记不确定的地方**：对任何不能 100% 确定的内容，生成一个明确的、可回答的问题

## 关键原则

- **宁可多问，不可猜错**。如果你对某个格子的含义、数据来源、判断逻辑不确定，必须生成问题。
- 每个问题必须具体、可回答，不要问模糊的问题。
- 对于数据表区域，要说明：表头各列含义、数据从台账哪些列映射、是否需要按样品展开多行
- 如果模板中有公式或计算逻辑的痕迹，要明确指出
- **必须分析所有工作表**，在 sheets 数组中逐个输出

## 返回格式

你必须严格返回以下 JSON（注意：sheets 是数组，包含所有工作表的分析）：

{
  "template_name": "简短准确的模板名称",
  "description": "模板用途描述",
  "report_type": "检测报告类型，如 压实度检测/材料检测/现场检测",
  "sheets": [
    {
      "sheet_name": "工作表名称",
      "sheet_role": "main_report|raw_data|reference|other",
      "sheet_description": "该工作表在报告中的作用说明",
      "regions": [
        {
          "id": "sheet1_region_1",
          "cells": "A1:F1 或 A5",
          "type": "title|project_info|sample_info|test_info|data_table_header|data_table_body|conclusion|signature|notes|other",
          "label": "该区域的显示标签/内容",
          "purpose": "该区域在报告中的作用",
          "data_source": "report_title|project_name|ledger.sample_no|ledger.test_parameter|ledger.required_value|ledger.measured_value|calculated|fixed_value|manual_input",
          "data_source_detail": "具体的数据来源说明",
          "fixed_value": "如果是固定值，这里填写固定内容",
          "logic": "如果有计算/判断逻辑，这里描述",
          "confidence": "high|medium|low",
          "question": "不确定时的问题（confidence为high时可为空）"
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
        "expand_logic": "每个样品一行"
      }
    }
  ],
  "cross_sheet_logic": [
    {
      "description": "Sheet2的检测结果填入Sheet1的数据表",
      "source_sheet": "原始记录表",
      "source_cells": "C33",
      "target_sheet": "报告正页",
      "target_cells": "I12",
      "logic": "压实度值从原始记录表引用到报告正页的压实度列"
    }
  ],
  "logic_rules": [
    {
      "rule_id": "rule_1",
      "description": "单项判定逻辑",
      "expression": "IF measured_value >= required_value THEN '合格' ELSE '不合格'",
      "depends_on": ["data_table.col_I"],
      "applies_to": "data_table.col_L",
      "applies_in_sheet": "报告正页",
      "confidence": "high",
      "question": ""
    }
  ],
  "uncertainties": [
    {
      "id": "q1",
      "sheet_name": "所属工作表",
      "region_id": "所属区域id",
      "question": "具体的问题描述",
      "context": "为什么不确定",
      "suggested_answer": "你建议的答案（可选）"
    }
  ],
  "notes": "任何补充说明"
}

注意：sheets 数组必须包含**所有**工作表，每个工作表各自有 regions、data_table。cross_sheet_logic 描述不同工作表之间的数据流动关系。uncertainties 汇总所有工作表的不确定问题。

你必须输出纯粹的 JSON，不要加任何代码块标记（```json），不要加解释文字。
JSON 中所有字符串值内的双引号必须用 \\" 转义。确保最后一个字段后面没有逗号。
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

    # 按工作表数量动态调整 max_tokens：每个 sheet 至少 3000 tokens
    sheet_count = len(xlsx_structure.get("sheets", []))
    max_tok = max(4000, sheet_count * 10000)

    client = OpenAI(base_url=llm["base_url"], api_key=llm["api_key"], timeout=300)

    response = client.chat.completions.create(
        model=llm["model"],
        messages=[
            {"role": "system", "content": TEMPLATE_ANALYSIS_SYSTEM_PROMPT},
            {"role": "user", "content": _build_analysis_prompt(file_path.name, xlsx_structure)},
        ],
        max_tokens=max_tok,
        temperature=0.1,
    )

    result = _parse_response(response.choices[0].message.content or "")
    result["source_file"] = str(file_path)
    result["analyzed_at"] = _now_iso()

    # 统计每个 sheet 的区域数
    total_regions = 0
    for s in result.get("sheets", []):
        total_regions += len(s.get("regions", []))

    # 收集不确定问题：来自 uncertainties 列表 + 各 sheet 的 regions 中低信心项
    uncertain_count = len(result.get("uncertainties", []))
    for s in result.get("sheets", []):
        for r in s.get("regions", []):
            if r.get("confidence") in ("medium", "low") and r.get("question"):
                uncertain_count += 1

    result["sheet_count"] = len(result.get("sheets", []))
    result["total_regions"] = total_regions
    result["needs_confirmation"] = uncertain_count > 0
    result["confirmation_required"] = uncertain_count

    # 向后兼容：如果 LLM 返回了旧格式（无 sheets），包装一下
    if "regions" in result and "sheets" not in result:
        result["sheets"] = [{
            "sheet_name": "Sheet1",
            "sheet_role": "main_report",
            "sheet_description": "",
            "regions": result.pop("regions", []),
            "data_table": result.pop("data_table", {}),
        }]
        result.setdefault("cross_sheet_logic", [])

    return result


def save_template_definition(
    project_name: str,
    template_definition: dict[str, Any],
    filename: str = "",
) -> dict[str, Any]:
    """将确认后的模板定义保存到数据库。"""
    from .. import database
    tname = filename or template_definition.get("template_name", "experiment_template")
    if tname.endswith(".json"):
        tname = tname[:-5]
    template_definition["saved_at"] = _now_iso()
    template_definition["project_name"] = project_name
    try:
        tid = database.save_template(
            project_name=project_name, template_name=tname,
            definition=template_definition,
            report_type=template_definition.get("report_type", ""),
            source_file=template_definition.get("source_file", ""),
        )
        return {"saved": True, "id": tid, "template_name": tname}
    except Exception as e:
        return {"error": f"保存失败: {e}"}


def load_template_definition(project_name: str, template_name: str) -> dict[str, Any] | None:
    """从数据库加载模板定义。"""
    from .. import database
    if template_name.endswith(".json"):
        template_name = template_name[:-5]
    tmpl = database.get_template(project_name, template_name)
    return tmpl["definition"] if tmpl else None


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

    # 处理各 sheet 中 regions 的 question
    for sheet in analysis_result.get("sheets", []):
        for region in sheet.get("regions", []):
            rid = region.get("id", "")
            if rid in answer_map:
                region["resolution"] = answer_map[rid]
                region["confidence"] = "high"
                region["question"] = ""

    # 向后兼容旧格式（无 sheets 包裹）
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
    """列出项目数据库中的模板定义。"""
    from .. import database
    return database.list_templates(project_name)


# ── 内部函数 ────────────────────────────────────────────────────

def _read_xlsx_structure(file_path: Path) -> dict[str, Any]:
    """读取 xlsx 完整结构信息。win32com 优先，无则回退 openpyxl。"""
    from .. import win32_helper
    return win32_helper.excel_read_structure(file_path)


def _build_analysis_prompt(filename: str, structure: dict[str, Any]) -> str:
    """构建发给 LLM 的分析请求。"""
    sheets = structure.get("sheets", [])
    parts = [
        f"请分析以下实验报告模板: {filename}",
        f"该模板包含 {len(sheets)} 个工作表，请逐一分析每一个工作表，不要遗漏。",
        "",
    ]

    for i, sheet in enumerate(sheets, 1):
        parts.append(f"## 工作表 {i}/{len(sheets)}: {sheet['name']}")
        merged = sheet.get("merged_ranges", sheet.get("merged_detail", []))
        parts.append(f"合并单元格: {len(merged)} 处")
        for m in merged[:30]:
            parts.append(f"  - {m}")
        parts.append("")

        parts.append("### 逐单元格内容（含精确位置）:")
        parts.append("")
        for row_info in sheet.get("rows", []):
            for c in row_info["cells"]:
                parts.append(f"  {c['cell']}: \"{c['value']}\"")
        parts.append("")

    parts.append("---")
    parts.append(f"以上共 {len(sheets)} 个工作表。")
    parts.append("请对每个工作表分别分析（sheets 数组中每个元素一个），特别注意：")
    parts.append("1. 判断每个工作表的角色（main_report/raw_data/reference/other）")
    parts.append("2. 数据表中每列与台账字段的对应关系")
    parts.append("3. 单元格间的计算/判定逻辑")
    parts.append("4. 不同工作表间的数据流动关系（填入 cross_sheet_logic）")
    parts.append("5. 任何不确定的地方，在 uncertainties 中生成具体问题")

    return "\n".join(parts)




def _parse_response(content: str) -> dict[str, Any]:
    """解析 LLM 返回的 JSON（jiter + repair 兜底）。"""
    result = json_utils.parse_llm_json(content)
    if "_raw" in result:
        return {
            "raw_response": content,
            "template_name": "解析失败",
            "sheets": [],
            "uncertainties": [{"id": "q0", "question": "LLM 返回格式异常，请检查原始输出",
                               "context": content[:500]}],
            "needs_confirmation": True,
            "confirmation_required": 1,
        }
    return result


def _now_iso() -> str:
    from datetime import datetime
    return datetime.now().isoformat()
