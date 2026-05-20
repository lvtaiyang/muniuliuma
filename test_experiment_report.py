"""实验报告生成 — 测试脚本

用法:
    python3 test_experiment_report.py              # 只跑不需要 LLM 的测试
    python3 test_experiment_report.py --full       # 完整测试（需要 LLM API key）
    python3 test_experiment_report.py --setup KEY  # 先配置 API key
"""

import json
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src import config
from src.experiment_report import template_analyzer
from src.construction_log import ledger_reader

TEST_DIR = "/mnt/e/木牛流马/实验报告生成测试"
TEMPLATE_FILE = f"{TEST_DIR}/1-K9+320.00～K9+520.00左幅水稳底基层试验段.xlsx"
LEDGER_FILE = f"{TEST_DIR}/水稳基层压实度检测记录_2023.01.01.xlsx"


def section(title: str):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


def step1_check_environment():
    """步骤1：检测运行环境"""
    section("步骤1: 环境检测")
    from src import win32_helper
    result = win32_helper.check_available()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def step2_check_files():
    """步骤2：检测文件是否存在"""
    section("步骤2: 文件检测")
    import os
    for label, path in [("模板文件", TEMPLATE_FILE), ("台账文件", LEDGER_FILE)]:
        exists = os.path.exists(path)
        size = os.path.getsize(path) if exists else 0
        print(f"  {label}: {path}")
        print(f"    存在: {exists}, 大小: {size:,} bytes")
    return os.path.exists(TEMPLATE_FILE) and os.path.exists(LEDGER_FILE)


def step3_read_ledger():
    """步骤3：读取台账数据"""
    section("步骤3: 读取台账")
    result = ledger_reader.read_ledger(LEDGER_FILE)
    if "error" in result:
        print(f"  错误: {result['error']}")
        return None

    print(f"  文件: {result.get('file', '')}")
    sheets = result.get("sheets", [])
    for s in sheets:
        rows = s.get("rows", [])
        print(f"  工作表 [{s['name']}]: {len(rows)} 行")
        if rows:
            print(f"    表头: {' | '.join(rows[0])}")
            print(f"    前3行数据:")
            for row in rows[1:4]:
                print(f"      {' | '.join(row[:6])}...")

    # 格式化给 LLM 看的部分
    formatted = ledger_reader.format_ledger_for_llm(result)
    print(f"\n  格式化后长度: {len(formatted)} 字符")
    return result


def step4_read_template_structure():
    """步骤4：读取模板结构（不用 LLM）"""
    section("步骤4: 读取模板结构")
    from src import win32_helper
    structure = win32_helper.excel_read_structure(TEMPLATE_FILE)
    if "error" in structure:
        print(f"  错误: {structure['error']}")
        return None

    print(f"  文件: {structure.get('filename', '')}")
    for sheet in structure.get("sheets", []):
        n_cells = sum(len(r["cells"]) for r in sheet.get("rows", []))
        n_merged = len(sheet.get("merged_ranges", []))
        print(f"  工作表 [{sheet['name']}]: "
              f"{len(sheet.get('rows', []))} 行有值, "
              f"{n_cells} 个单元格, "
              f"{n_merged} 处合并")

    # 展示第一个工作表的关键单元格
    first_sheet = structure["sheets"][0]
    print(f"\n  [{first_sheet['name']}] 关键内容:")
    for row_info in first_sheet.get("rows", [])[:40]:
        for c in row_info["cells"]:
            print(f"    {c['cell']}: {c['value'][:80]}")

    print(f"\n  合并区域 ({len(first_sheet.get('merged_ranges', []))} 处):")
    for m in first_sheet.get("merged_ranges", [])[:20]:
        print(f"    {m}")
    if len(first_sheet.get("merged_ranges", [])) > 20:
        print(f"    ... 还有 {len(first_sheet['merged_ranges']) - 20} 处")

    return structure


def step5_analyze_template():
    """步骤5：LLM 分析模板（需要 API key）"""
    section("步骤5: LLM 分析模板")

    conf = config.load()
    llm = config.get_llm_config("text")
    if not llm.get("api_key"):
        print("  跳过: 未配置 LLM text.api_key")
        print("  设置方法: python3 test_experiment_report.py --setup YOUR_DEEPSEEK_API_KEY")
        return None

    print("  正在调用 LLM 分析模板...")
    result = template_analyzer.analyze_template(TEMPLATE_FILE)

    if "error" in result:
        print(f"  错误: {result['error']}")
        return None

    print(f"  模板名称: {result.get('template_name', '')}")
    print(f"  报告类型: {result.get('report_type', '')}")
    print(f"  区域数量: {len(result.get('regions', []))}")
    print(f"  逻辑规则: {len(result.get('logic_rules', []))}")
    print(f"  需要确认: {result.get('needs_confirmation', False)}")
    print(f"  待确认问题: {result.get('confirmation_required', 0)}")

    # 展示关键区域
    print(f"\n  数据表定义:")
    dt = result.get("data_table", {})
    if dt:
        print(f"    表头行: {dt.get('header_row', '?')}")
        print(f"    数据起始行: {dt.get('data_start_row', '?')}")
        for col in dt.get("columns", []):
            print(f"    列{col.get('col_letter', '')}: {col.get('header_text', '')} → {col.get('data_source', '')} (confidence: {col.get('confidence', '')})")

    # 展示不确定问题
    uncertainties = result.get("uncertainties", [])
    if uncertainties:
        print(f"\n  ⚠ 需要确认的问题 ({len(uncertainties)} 个):")
        for u in uncertainties:
            print(f"    [{u.get('id', '')}] {u.get('question', '')}")
            if u.get("suggested_answer"):
                print(f"    建议: {u['suggested_answer']}")

    # 保存结果
    output_path = "/tmp/template_analysis_result.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n  完整结果已保存到: {output_path}")

    return result


def step6_data_extraction_demo(template_analysis):
    """步骤6：演示数据提取逻辑（不用 LLM）"""
    section("步骤6: 数据提取分析")

    conf = config.load()
    llm = config.get_llm_config("text")
    if not llm.get("api_key"):
        print("  跳过: 未配置 LLM text.api_key")
        return

    # 构建发送给 LLM 的 prompt 预览
    from src.experiment_report import report_generator
    template_desc = report_generator._describe_template_for_extraction(template_analysis)

    ledgers = ledger_reader.read_ledger(LEDGER_FILE)
    ledger_text = "\n\n".join(ledger_reader.format_ledger_for_llm(l) for l in [ledgers])

    print(f"  模板描述长度: {len(template_desc)} 字符")
    print(f"  台账数据长度: {len(ledger_text)} 字符")
    print(f"\n  --- 发给 LLM 的模板描述（前 1500 字符）---")
    print(template_desc[:1500])
    print(f"  ... (共 {len(template_desc)} 字符)")


def main():
    full_mode = "--full" in sys.argv
    setup_key = None

    for i, arg in enumerate(sys.argv):
        if arg == "--setup" and i + 1 < len(sys.argv):
            setup_key = sys.argv[i + 1]

    if setup_key:
        conf = config.load()
        conf["llm"]["text"]["api_key"] = setup_key
        config.save(conf)
        print(f"API key 已配置 (text)")
        full_mode = True

    # 检查是否需要 full mode
    conf = config.load()
    if config.get_llm_config("text").get("api_key"):
        full_mode = True
        print("检测到已配置 text API key，自动启用 --full 模式")

    env = step1_check_environment()
    files_ok = step2_check_files()
    if not files_ok:
        print("\n请确认测试文件路径正确")
        return

    ledger = step3_read_ledger()
    structure = step4_read_template_structure()

    if full_mode:
        analysis = step5_analyze_template()
        if analysis:
            step6_data_extraction_demo(analysis)


if __name__ == "__main__":
    main()
