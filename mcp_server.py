"""木牛流马 MCP Server — 工程行业 AI 技能工具箱。

每个模块独立注册自己的 MCP 工具，框架无关。

启动方式:
    python mcp_server.py              # stdio 模式（本地智能体）
    python mcp_server.py --sse        # SSE HTTP 模式（远程/WSL 智能体，默认 :8700）
    python mcp_server.py --sse --port 8700 --host 0.0.0.0

MCP 配置:
    stdio: {"command": "python", "args": ["path/to/mcp_server.py"]}
    SSE:   {"url": "http://localhost:8700/sse"}
"""

from __future__ import annotations

import json
import sys
import os
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src import config as cfg
from src import database
from mcp.server.lowlevel import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# ── 注册模块 ────────────────────────────────────────────────────
# 新模块：在 src/ 下新建子包，在此导入并注册工具即可。

MODULES = {}

# 加载微信影像监控模块
try:
    from src.wechat_monitor import wechat_reader, monitor, summary_writer

    MODULES["wechat_monitor"] = {
        "name": "微信隐蔽验收影像归档",
        "tools": {},
    }
except ImportError:
    pass

# 加载建设期项目初始化模块
try:
    from src.project_initializer import (
        project_manager,
        template_analyzer,
        template_manager,
    )

    MODULES["project_initializer"] = {
        "name": "建设期项目初始化与模板管理",
        "tools": {},
    }
except ImportError:
    pass

# 加载现有资料整理模块
try:
    from src.image_organizer import organizer, scanner as img_scanner

    MODULES["image_organizer"] = {
        "name": "现有影像资料整理",
        "tools": {},
    }
except ImportError:
    pass

# 加载施工日志模块
try:
    from src.construction_log import log_generator, ledger_reader

    MODULES["construction_log"] = {
        "name": "施工日志自动生成",
        "tools": {},
    }
except ImportError:
    pass

# 加载实验报告模块
try:
    from src.experiment_report import report_generator, template_analyzer

    MODULES["experiment_report"] = {
        "name": "实验检测报告自动生成",
        "tools": {},
    }
except ImportError:
    pass

# 加载 WBS 解析器
try:
    from src.project_initializer import wbs_parser

    MODULES["wbs_parser"] = {
        "name": "分部分项解析",
        "tools": {},
    }
except ImportError:
    pass

APP = Server("muniuliuma")


# ── 环境检测与安装工具 ──────────────────────────────────────────

async def tool_check_setup() -> str:
    """检测 wechat-cli 运行环境状态。"""
    result = wechat_reader.check_setup()
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_install_wechat_cli() -> str:
    """自动安装 wechat-cli（需要 Node.js >= 14）。"""
    result = wechat_reader.install()
    return json.dumps(result, ensure_ascii=False, indent=2)


# ── 核心工具 ────────────────────────────────────────────────────

async def tool_list_wechat_groups() -> str:
    """列出微信群聊。"""
    try:
        groups = wechat_reader.list_groups(limit=200)
    except Exception as exc:
        return json.dumps({"error": str(exc)}, ensure_ascii=False)
    return json.dumps(
        [{"name": g.get("chat", ""), "username": g.get("username", ""),
          "last_message": (g.get("last_message", "") or "")[:80]}
         for g in groups],
        ensure_ascii=False, indent=2,
    )


async def tool_setup(groups: str = "", projects: str = "",
                     api_key: str = "", base_url: str = "",
                     model: str = "", archive_path: str = "",
                     project_name: str = "") -> str:
    """配置监控参数。"""
    conf = cfg.load()

    if groups:
        try:
            gl = json.loads(groups)
            if isinstance(gl, list):
                conf.setdefault("wechat_monitor", {})["groups"] = gl
        except json.JSONDecodeError:
            return json.dumps({"error": "groups 不是有效 JSON"}, ensure_ascii=False)

    if projects:
        try:
            pl = json.loads(projects)
            if isinstance(pl, list):
                conf["projects"] = pl
        except json.JSONDecodeError:
            return json.dumps({"error": "projects 不是有效 JSON"}, ensure_ascii=False)

    if project_name:
        conf.setdefault("wechat_monitor", {})["linked_project"] = project_name

    if api_key:
        conf["llm"].setdefault("text", {})["api_key"] = api_key
        conf["llm"].setdefault("vision", {})["api_key"] = api_key
    if base_url:
        conf["llm"].setdefault("text", {})["base_url"] = base_url
        conf["llm"].setdefault("vision", {})["base_url"] = base_url
    if model:
        conf["llm"].setdefault("text", {})["model"] = model
        conf["llm"].setdefault("vision", {})["model"] = model
    if archive_path:
        conf["archive"]["base_path"] = archive_path

    saved = cfg.save(conf)
    errors = cfg.validate(conf)
    archive = cfg.resolve_archive_path(conf)
    return json.dumps({
        "saved": str(saved),
        "config": {
            "groups": conf.get("wechat_monitor", {}).get("groups", []),
            "linked_project": conf.get("wechat_monitor", {}).get("linked_project", ""),
            "projects": [p["name"] for p in conf.get("projects", [])],
            "llm_text_model": conf.get("llm", {}).get("text", {}).get("model", ""),
            "llm_vision_model": conf.get("llm", {}).get("vision", {}).get("model", ""),
            "archive_path": str(archive),
        },
        "warnings": errors,
    }, ensure_ascii=False, indent=2)


async def tool_run_monitor() -> str:
    """执行一次监控管线。"""
    result = monitor.run_monitor()
    return json.dumps({
        "success": result.success,
        "skipped": result.skipped,
        "errors": result.errors,
        "entries": [
            {"project": e.get("project"), "category": e.get("category"),
             "filename": e.get("filename"), "description": e.get("description"),
             "quality_notes": e.get("quality_notes")}
            for e in result.entries
        ],
    }, ensure_ascii=False, indent=2)


async def tool_get_summary() -> str:
    """获取归档汇总 MD。"""
    conf = cfg.load()
    archive_base = cfg.resolve_archive_path(conf)
    entries = summary_writer.collect_entries(archive_base)
    return summary_writer.generate_summary(archive_base, entries)


async def tool_get_config() -> str:
    """查看当前配置。"""
    conf = cfg.load()
    # 脱敏处理
    for key_type in ("text", "vision"):
        sub = conf.get("llm", {}).get(key_type, {})
        if sub.get("api_key"):
            sub["api_key"] = sub["api_key"][:8] + "****"
    return json.dumps(conf, ensure_ascii=False, indent=2)


async def tool_update_config(field: str, value: str) -> str:
    """更新单个配置项。"""
    conf = cfg.load()
    try:
        parsed = json.loads(value)
    except (json.JSONDecodeError, TypeError):
        parsed = value
    keys = field.split(".")
    target = conf
    for k in keys[:-1]:
        target.setdefault(k, {})
        if not isinstance(target[k], dict):
            return json.dumps({"error": f"路径 {field} 类型冲突"}, ensure_ascii=False)
        target = target[k]
    target[keys[-1]] = parsed
    saved = cfg.save(conf)
    return json.dumps({"saved": str(saved), "field": field}, ensure_ascii=False)


# ── 项目初始化工具 ──────────────────────────────────────────────

async def tool_project_init(
    name: str, workspace: str,
    project_type: str = "", current_stage: str = "",
    meta_json: str = "{}",
) -> str:
    """创建新工程项目，搭建全流程工作区。"""
    try:
        meta = json.loads(meta_json) if meta_json else {}
    except json.JSONDecodeError:
        meta = {}
    project = project_manager.create(
        name=name,
        workspace=workspace,
        project_type=project_type,
        current_stage=current_stage,
        **meta,
    )
    # 同时创建隐蔽验收影像资料目录
    ws = Path(project["workspace"])
    image_dir = ws / "04_施工实施" / "隐蔽验收影像资料"
    image_dir.mkdir(parents=True, exist_ok=True)
    project["image_archive_dir"] = str(image_dir)
    return json.dumps(project, ensure_ascii=False, indent=2)


async def tool_project_list() -> str:
    """列出所有已创建项目。"""
    projects = project_manager.list_projects()
    return json.dumps(projects, ensure_ascii=False, indent=2)


async def tool_project_load(name: str) -> str:
    """加载项目完整信息。"""
    project = project_manager.load(name)
    if project is None:
        return json.dumps({"error": f"项目不存在: {name}"}, ensure_ascii=False)
    return json.dumps(project, ensure_ascii=False, indent=2)


async def tool_project_update(name: str, updates_json: str) -> str:
    """更新项目信息。"""
    try:
        updates = json.loads(updates_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "updates_json 不是有效 JSON"}, ensure_ascii=False)
    result = project_manager.update(name, **updates)
    if result is None:
        return json.dumps({"error": f"项目不存在: {name}"}, ensure_ascii=False)
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_template_analyze(file_path: str) -> str:
    """分析用户上传的文档模板（Word/Excel/PDF）。"""
    conf = cfg.load()
    llm = cfg.get_llm_config("text")
    if not llm.get("api_key"):
        return json.dumps({"error": "未配置 LLM text.api_key，请先执行 setup_monitoring 配置模型"}, ensure_ascii=False)
    result = template_analyzer.analyze_document(
        file_path,
        base_url=llm["base_url"],
        api_key=llm["api_key"],
        model=llm["model"],
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_template_save(
    project_name: str, template_json: str,
    filename: str = "",
) -> str:
    """保存确认后的模板到项目。"""
    try:
        template_def = json.loads(template_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "template_json 不是有效 JSON"}, ensure_ascii=False)
    result = template_manager.save_template(
        project_name, template_def, filename=filename,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_template_list(project_name: str) -> str:
    """列出项目所有模板。"""
    templates = template_manager.list_templates(project_name)
    return json.dumps(templates, ensure_ascii=False, indent=2)


async def tool_template_apply(
    project_name: str, template_name: str,
    field_values_json: str, context: str = "",
) -> str:
    """用模板生成一个文档实例。"""
    try:
        field_values = json.loads(field_values_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "field_values_json 不是有效 JSON"}, ensure_ascii=False)
    result = template_manager.apply_template(
        project_name, template_name, field_values, context=context,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_project_discover(workspace: str) -> str:
    """从工作区路径反向查找关联项目。"""
    project = project_manager.discover_project(workspace)
    if project is None:
        return json.dumps({"error": "该目录不是木牛流马工作区"}, ensure_ascii=False)
    return json.dumps(project, ensure_ascii=False, indent=2)


# ── 影像资料整理工具 ────────────────────────────────────────────

async def tool_organize_images(
    source_dir: str, recursive: str = "true",
    limit: str = "50",
) -> str:
    """扫描并整理指定目录中的现有影像资料。"""
    result = organizer.organize_images(
        source_dir=source_dir,
        recursive=recursive.lower() != "false",
        limit=int(limit),
    )
    return json.dumps({
        "total_found": result.total_found,
        "new_images": result.new_images,
        "success": result.success,
        "skipped": result.skipped,
        "errors": result.errors,
        "entries": [
            {"project": e.get("project"), "category": e.get("category"),
             "filename": e.get("filename"), "description": e.get("description"),
             "quality_notes": e.get("quality_notes")}
            for e in result.entries
        ],
    }, ensure_ascii=False, indent=2)


async def tool_scan_directory(
    source_dir: str, recursive: str = "true",
) -> str:
    """扫描目录，预览图片列表（不进行分析）"""
    images = img_scanner.scan_directory(source_dir, recursive.lower() != "false")
    return json.dumps({
        "total": len(images),
        "files": [{"name": i["name"], "relative_path": i["relative_path"],
                    "size": i["size"]} for i in images[:200]],
    }, ensure_ascii=False, indent=2)


# ── 施工日志工具 ────────────────────────────────────────────────

async def tool_generate_daily_log(
    project_name: str, log_date: str = "",
    max_images: str = "30", template_path: str = "",
) -> str:
    """从当天归档的影像资料生成施工日志。有 Word 模板则输出 .docx。"""
    result = log_generator.generate_from_images(
        project_name=project_name,
        log_date=log_date,
        max_images=int(max_images),
        template_path=template_path,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_generate_log_from_ledgers(
    project_name: str, ledger_paths_json: str,
    log_date: str = "", template_path: str = "",
) -> str:
    """从台账文件生成施工日志。有 Word 模板则输出 .docx。"""
    try:
        paths = json.loads(ledger_paths_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "ledger_paths_json 不是有效 JSON"}, ensure_ascii=False)
    result = log_generator.generate_from_ledgers(
        project_name=project_name,
        ledger_paths=paths,
        log_date=log_date,
        template_path=template_path,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_list_logs(project_name: str) -> str:
    """列出项目已有的施工日志。"""
    logs = log_generator.list_logs(project_name)
    return json.dumps(logs, ensure_ascii=False, indent=2)


async def tool_read_log(project_name: str, filename: str) -> str:
    """读取一篇施工日志内容。"""
    content = log_generator.read_log(project_name, filename)
    if content is None:
        return json.dumps({"error": f"日志不存在: {filename}"}, ensure_ascii=False)
    return content


# ── 实验报告工具 ────────────────────────────────────────────────

async def tool_analyze_experiment_template(file_path: str) -> str:
    """分析实验报告 xlsx 模板：逐单元格理解含义、数据来源、逻辑关系、不确定处生成确认问题。"""
    result = template_analyzer.analyze_template(file_path)
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_confirm_template_analysis(
    project_name: str, analysis_result_json: str,
    answers_json: str, save_filename: str = "",
) -> str:
    """用回答确认模板分析中的不确定问题，确认后保存模板定义到项目。"""
    try:
        analysis_result = json.loads(analysis_result_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "analysis_result_json 不是有效 JSON"}, ensure_ascii=False)

    confirmed = template_analyzer.confirm_template(analysis_result, answers_json)
    if confirmed.get("error"):
        return json.dumps(confirmed, ensure_ascii=False)

    if confirmed.get("confirmed"):
        saved = template_analyzer.save_template_definition(
            project_name, confirmed, filename=save_filename,
        )
        confirmed["saved"] = saved

    return json.dumps(confirmed, ensure_ascii=False, indent=2)


async def tool_list_experiment_templates(project_name: str) -> str:
    """列出项目已保存的实验报告模板定义。"""
    templates = template_analyzer.list_template_definitions(project_name)
    return json.dumps(templates, ensure_ascii=False, indent=2)


async def tool_generate_experiment_report(
    project_name: str, template_name: str,
    ledger_paths_json: str, report_date: str = "",
) -> str:
    """用已确认的模板定义 + 台账数据生成实验报告（xlsx 格式）。"""
    try:
        paths = json.loads(ledger_paths_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "ledger_paths_json 不是有效 JSON"}, ensure_ascii=False)
    result = report_generator.generate_from_template(
        project_name=project_name,
        template_name=template_name,
        ledger_paths=paths,
        report_date=report_date,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_generate_single_experiment_report(
    project_name: str, template_name: str,
    test_item: str, ledger_paths_json: str,
    report_date: str = "", extra_context: str = "",
) -> str:
    """用模板为特定检测项目生成单份实验报告（如钢筋原材力学性能检测报告）。"""
    try:
        paths = json.loads(ledger_paths_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "ledger_paths_json 不是有效 JSON"}, ensure_ascii=False)
    result = report_generator.generate_single_from_template(
        project_name=project_name,
        template_name=template_name,
        test_item=test_item,
        ledger_paths=paths,
        report_date=report_date,
        extra_context=extra_context,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_list_experiment_reports(project_name: str) -> str:
    """列出项目已有的实验报告。"""
    reports = report_generator.list_reports(project_name)
    return json.dumps(reports, ensure_ascii=False, indent=2)


async def tool_read_experiment_report(project_name: str, filename: str) -> str:
    """读取一篇实验报告内容。xlsx 格式返回 Markdown 表格预览。"""
    content = report_generator.read_report(project_name, filename)
    if content is None:
        return json.dumps({"error": f"报告不存在: {filename}"}, ensure_ascii=False)
    return content


# ── 纯 I/O 工具（智能体驱动，不调 LLM）─────────────────────────

async def tool_read_template_cells(file_path: str) -> str:
    """读取 xlsx 模板的所有单元格位置和内容。纯数据提取，不调 LLM。智能体拿到后自己分析结构、匹配台账列、决定填充内容。"""
    from src import win32_helper
    result = win32_helper.excel_read_structure(file_path)
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_fill_template_cells(
    template_path: str, output_path: str, cell_values_json: str,
) -> str:
    """向 xlsx 模板指定单元格写入值并另存。纯数据操作。

    cell_values_json:
      {"Sheet1": {"A3": "填的值", "B5": "填的值"}, "Sheet2": {"C1": "值"}}
      或 {"cells": {"A1": "值", "B2": "值"}}
    """
    from src import win32_helper
    try:
        values = json.loads(cell_values_json)
    except json.JSONDecodeError:
        return json.dumps({"error": "cell_values_json 不是有效 JSON"}, ensure_ascii=False)

    tp = Path(template_path)
    op = Path(output_path)
    if not tp.exists():
        return json.dumps({"error": f"模板不存在: {tp}"}, ensure_ascii=False)
    op.parent.mkdir(parents=True, exist_ok=True)

    try:
        fill_data = _build_simple_fill_data(values)
        win32_helper.excel_fill_template(tp, op, fill_data)
        return json.dumps({"saved": str(op), "cells_written": _count_cells(values)}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"填充失败: {e}"}, ensure_ascii=False)


def _build_simple_fill_data(values: dict[str, Any]) -> dict[str, Any]:
    cells = values.get("cells", {})
    sheets = []
    if not cells:
        for sname, svals in values.items():
            if isinstance(svals, dict):
                regions = [{"id": f"c_{cr}", "cells": cr, "fixed_value": str(v),
                            "data_source": "fixed_value", "confidence": "high"}
                           for cr, v in svals.items()]
                sheets.append({"sheet_name": sname, "sheet_role": "main_report",
                               "regions": regions, "data_table": {}})
    else:
        regions = [{"id": f"c_{cr}", "cells": cr, "fixed_value": str(v),
                    "data_source": "fixed_value", "confidence": "high"}
                   for cr, v in cells.items()]
        sheets = [{"sheet_name": "", "sheet_role": "main_report", "regions": regions, "data_table": {}}]
    return {"sheets": sheets, "report_data": {}, "data_table_rows": [], "logic_rules": []}


def _count_cells(values: dict[str, Any]) -> int:
    cells = values.get("cells", {})
    if cells: return len(cells)
    return sum(len(v) for v in values.values() if isinstance(v, dict))


# ── 活动日志工具 ────────────────────────────────────────────────

# ── 分部分项解析工具 ────────────────────────────────────────────

async def tool_parse_project_document(file_path: str) -> str:
    """解析项目文档（分部分项清单/合同/量单），返回结构化 WBS 和合同信息。"""
    result = wbs_parser.parse_project_document(file_path)
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_parse_project_text(text: str, context: str = "") -> str:
    """从用户文本描述中解析项目分部分项结构。"""
    result = wbs_parser.parse_from_text(text, context)
    return json.dumps(result, ensure_ascii=False, indent=2)


async def tool_get_project_wbs(project_name: str) -> str:
    """获取项目的分部分项树。"""
    wbs = database.get_wbs_tree(project_name)
    contract = database.get_contract(project_name)
    return json.dumps({
        "project": project_name,
        "wbs": wbs,
        "contract": contract,
    }, ensure_ascii=False, indent=2)


async def tool_get_activity_log(
    limit: str = "20", tool_name: str = "",
    module: str = "", status: str = "",
    since: str = "",
) -> str:
    """查询 MCP 工具的调用历史。"""
    logs = database.query_calls(
        limit=int(limit), tool_name=tool_name,
        module=module, status=status, since=since,
    )
    compact = [{
        "time": e["called_at"], "tool": e["tool_name"],
        "module": e["module"], "status": e["status"],
        "duration_ms": e["duration_ms"], "summary": e["result_summary"][:200],
    } for e in logs]
    return json.dumps({
        "count": len(compact), "logs": compact,
        "hint": "用 since='1h'/'24h' 筛选时间，用 module='experiment_report' 筛选模块",
    }, ensure_ascii=False, indent=2)


async def tool_get_activity_summary(since: str = "24h") -> str:
    """获取最近活动概要。"""
    summary = database.get_call_summary(since=since)
    return json.dumps(summary, ensure_ascii=False, indent=2)


# ── 工具注册表 ──────────────────────────────────────────────────

TOOLS = [
    # ── 系统工具 ──
    # ── 纯 I/O 工具（智能体驱动）──
    Tool(
        name="read_template_cells",
        description="读取 xlsx 模板的所有单元格位置和内容（纯数据，不调 LLM）。智能体拿到原始数据后自己分析结构、匹配合账、决定填充内容。",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "xlsx 文件路径"},
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="fill_template_cells",
        description="向 xlsx 模板的指定单元格写入值并另存（纯数据操作，不调 LLM）。cell_values_json 格式: {\"Sheet1\": {\"A3\": \"值\", \"B5\": \"值\"}} 或 {\"cells\": {\"A1\": \"值\"}}",
        inputSchema={
            "type": "object",
            "properties": {
                "template_path": {"type": "string", "description": "模板 xlsx 路径"},
                "output_path": {"type": "string", "description": "输出 xlsx 路径"},
                "cell_values_json": {"type": "string", "description": "单元格值 JSON"},
            },
            "required": ["template_path", "output_path", "cell_values_json"],
        },
    ),
    Tool(
        name="get_activity_log",
        description="查询 MCP 工具调用历史日志。智能体在开始任务前应先查看此日志，了解之前做了什么、避免重复运行。支持按工具名、模块、状态、时间过滤。",
        inputSchema={
            "type": "object",
            "properties": {
                "limit": {"type": "string", "description": "返回条数，默认 20"},
                "tool_name": {"type": "string", "description": "按工具名筛选，如 'analyze_experiment_template'"},
                "module": {"type": "string", "description": "按模块筛选，如 'experiment_report'"},
                "status": {"type": "string", "description": "按状态筛选: success / error"},
                "since": {"type": "string", "description": "时间范围: 1h / 24h / 7d"},
            },
        },
    ),
    Tool(
        name="get_activity_summary",
        description="获取最近 MCP 活动概要：调用次数、成功率、各模块/工具使用统计、最近错误。适合在开始工作前快速了解当前状态。",
        inputSchema={
            "type": "object",
            "properties": {
                "since": {"type": "string", "description": "时间范围: 1h / 24h / 7d，默认 24h"},
            },
        },
    ),
    Tool(
        name="check_wechat_setup",
        description="检测 wechat-cli 运行环境：Node.js 是否安装、wechat-cli 是否可用、依赖是否就绪。首次使用前应先调用此工具。",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="install_wechat_cli",
        description="自动安装 wechat-cli。需要 Node.js >= 14。会尝试在项目内置依赖中安装，失败则尝试全局安装。",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="list_wechat_groups",
        description="列出微信中可用的群聊会话，供用户选择监控哪些群。",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="setup_monitoring",
        description="配置微信影像监控：群聊、项目信息、LLM 模型、归档路径。设置 project_name 可将影像资料自动归档到项目工作区的 04_施工实施/隐蔽验收影像资料/ 目录。",
        inputSchema={
            "type": "object",
            "properties": {
                "groups": {"type": "string", "description": "群聊名称 JSON 数组"},
                "projects": {"type": "string", "description": "项目列表 JSON: [{\"name\":\"项目A\",\"categories\":[\"地基验槽\"]}]"},
                "project_name": {"type": "string", "description": "关联的项目名称，设置后影像自动归档到项目工作区"},
                "api_key": {"type": "string", "description": "多模态 LLM API Key"},
                "base_url": {"type": "string", "description": "LLM API 地址 (OpenAI 兼容协议)"},
                "model": {"type": "string", "description": "模型名称"},
                "archive_path": {"type": "string", "description": "归档根目录（不关联项目时使用，默认 ~/隐蔽验收影像资料）"},
            },
        },
    ),
    Tool(
        name="run_monitor",
        description="手动执行一次微信群影像归档：拉取图片 → 多模态分析 → 分类归档 → 更新汇总。",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="get_summary",
        description="获取当前影像资料归档汇总（Markdown）。",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="get_config",
        description="查看当前配置（API key 脱敏）。",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="update_config",
        description="更新指定配置项，如 'llm.model'。",
        inputSchema={
            "type": "object",
            "properties": {
                "field": {"type": "string", "description": "配置路径"},
                "value": {"type": "string", "description": "新值"},
            },
            "required": ["field", "value"],
        },
    ),
    # ── 项目初始化工具 ──
    Tool(
        name="project_init",
        description="创建新工程项目。按阶段搭建目录，施工阶段按分部分项(WBS)创建资料子目录。wbs_json 为 parse_project_document 返回的 wbs 部分，contract_json 为合同部分。",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "项目名称"},
                "workspace": {"type": "string", "description": "工作区绝对路径，如 /home/user/万科城市花园"},
                "project_type": {"type": "string", "description": "项目类型：住宅/工业/市政/公路/水利/其他"},
                "current_stage": {"type": "string", "description": "当前阶段代码，如 04_施工实施"},
                "meta_json": {"type": "string", "description": "额外元数据 JSON"},
                "wbs_json": {"type": "string", "description": "分部分项 JSON（parse_project_document 返回的 wbs 数组或完整结果）"},
                "contract_json": {"type": "string", "description": "合同信息 JSON（parse_project_document 返回的 contract 对象）"},
            },
            "required": ["name", "workspace"],
        },
    ),
    Tool(
        name="project_list",
        description="列出所有已创建的项目。",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="project_load",
        description="加载项目完整信息：阶段、模板、元数据。",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "项目名称"},
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="project_update",
        description="更新项目信息，如切换当前阶段、更新元数据等。updates_json 是包含要更新字段的 JSON 对象。",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "项目名称"},
                "updates_json": {"type": "string", "description": "更新字段 JSON，如 {\"current_stage\": \"05_竣工验收\"}"},
            },
            "required": ["name", "updates_json"],
        },
    ),
    Tool(
        name="project_discover",
        description="从工作区路径反向查找关联的项目，判断一个目录是否已经是木牛流马工作区。",
        inputSchema={
            "type": "object",
            "properties": {
                "workspace": {"type": "string", "description": "目录路径"},
            },
            "required": ["workspace"],
        },
    ),
    Tool(
        name="template_analyze",
        description="分析用户上传的文档模板（Word/Excel/PDF/CSV/TXT），用 LLM 提取字段名、类型、使用场景，返回结构化模板定义供用户确认。",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文档文件的绝对路径"},
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="template_save",
        description="将用户确认后的模板定义保存到指定项目。template_json 是 template_analyze 返回的结果（可修改后传入）。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "template_json": {"type": "string", "description": "模板定义 JSON"},
                "filename": {"type": "string", "description": "可选的模板文件名"},
            },
            "required": ["project_name", "template_json"],
        },
    ),
    Tool(
        name="template_list",
        description="列出项目的所有模板及其使用次数。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
            },
            "required": ["project_name"],
        },
    ),
    Tool(
        name="template_apply",
        description="用已有模板生成一个文档实例（如用'材料检验表'模板生成某个具体验收记录）。field_values_json 为字段名到值的映射。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "template_name": {"type": "string", "description": "模板名称"},
                "field_values_json": {"type": "string", "description": "字段值 JSON: {\"材料名称\":\"钢筋\",\"检验结论\":\"合格\"}"},
                "context": {"type": "string", "description": "实例场景描述，如'3号楼钢筋隐蔽验收'"},
            },
            "required": ["project_name", "template_name", "field_values_json"],
        },
    ),
    # ── 分部分项解析工具 ──
    Tool(
        name="parse_project_document",
        description="分析项目文档（合同/分部分项清单/工程量清单），用LLM提取合同关键信息和分部分项结构（单位工程→分部工程→分项工程→检验批）。支持xlsx/docx/csv/txt。",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文档文件路径"},
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="parse_project_text",
        description="从文本描述中解析项目分部分项结构。用户口头描述项目划分时使用。",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {"type": "string", "description": "项目结构文本描述"},
                "context": {"type": "string", "description": "补充说明，如项目类型"},
            },
            "required": ["text"],
        },
    ),
    Tool(
        name="get_project_wbs",
        description="获取项目的分部分项树和合同信息。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
            },
            "required": ["project_name"],
        },
    ),
    # ── 影像资料整理工具 ──
    Tool(
        name="scan_directory",
        description="扫描指定目录中的图片文件，返回文件列表（不进行分析）。用于预览确认后再调用 organize_images 执行分析归档。",
        inputSchema={
            "type": "object",
            "properties": {
                "source_dir": {"type": "string", "description": "要扫描的目录路径"},
                "recursive": {"type": "string", "description": "是否递归扫描子目录，默认 true"},
            },
            "required": ["source_dir"],
        },
    ),
    Tool(
        name="organize_images",
        description="整理指定目录中的现有影像资料：扫描 → 多模态分析 → 分类 → 归档 → 更新汇总。支持增量（跳过已归档的）。",
        inputSchema={
            "type": "object",
            "properties": {
                "source_dir": {"type": "string", "description": "图片源目录路径"},
                "recursive": {"type": "string", "description": "是否递归扫描，默认 true"},
                "limit": {"type": "string", "description": "最多处理张数，默认 50"},
            },
            "required": ["source_dir"],
        },
    ),
    # ── 施工日志工具 ──
    Tool(
        name="generate_daily_log",
        description="从归档的施工影像资料，用多模态模型分析并生成施工日志。有 Word 模板则填入模板输出 .docx，否则输出 .md。保存到项目工作区的 04_施工实施/施工日志/。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "log_date": {"type": "string", "description": "日志日期 YYYY-MM-DD，默认今天"},
                "max_images": {"type": "string", "description": "最多分析图片张数，默认 30"},
                "template_path": {"type": "string", "description": "用户上传的施工日志 Word 模板路径（.docx），有则填入模板输出"},
            },
            "required": ["project_name"],
        },
    ),
    Tool(
        name="generate_log_from_ledgers",
        description="从材料进场台账、实验检测台账等施工文档，用 LLM 分析汇总生成施工日志。有 Word 模板则输出 .docx。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "ledger_paths_json": {"type": "string", "description": "台账文件路径 JSON 数组: [\"/path/to/材料台账.xlsx\", \"/path/to/检测台账.csv\"]"},
                "log_date": {"type": "string", "description": "日志日期 YYYY-MM-DD，默认今天"},
                "template_path": {"type": "string", "description": "用户上传的施工日志 Word 模板路径（.docx）"},
            },
            "required": ["project_name", "ledger_paths_json"],
        },
    ),
    Tool(
        name="list_logs",
        description="列出项目已有的施工日志文件。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
            },
            "required": ["project_name"],
        },
    ),
    Tool(
        name="read_log",
        description="读取一篇施工日志的完整内容。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "filename": {"type": "string", "description": "日志文件名"},
            },
            "required": ["project_name", "filename"],
        },
    ),
    # ── 实验报告工具 ──
    Tool(
        name="analyze_experiment_template",
        description="分析用户上传的实验报告 xlsx 模板。逐单元格/区域深度理解：每个格子填什么、数据从台账哪个字段来、单元格间的计算/判定逻辑。不确定的地方会生成具体问题让用户确认。",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "xlsx 模板文件的绝对路径"},
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="confirm_template_analysis",
        description="回答模板分析中的不确定问题，确认后保存模板定义为可复用的实验报告模板。answers_json 格式: [{\"question_id\": \"q1\", \"answer\": \"...\"}] 或 [{\"region_id\": \"region_X\", \"resolution\": \"...\"}]。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "analysis_result_json": {"type": "string", "description": "analyze_experiment_template 返回的完整 JSON（可修改其中不确定的部分后再传入）"},
                "answers_json": {"type": "string", "description": "对不确定问题的回答 JSON 数组"},
                "save_filename": {"type": "string", "description": "保存的模板文件名，默认自动生成"},
            },
            "required": ["project_name", "analysis_result_json", "answers_json"],
        },
    ),
    Tool(
        name="list_experiment_templates",
        description="列出项目已保存的实验报告模板定义。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
            },
            "required": ["project_name"],
        },
    ),
    Tool(
        name="generate_experiment_report",
        description="使用已确认的模板定义 + 实验检测台账，用 LLM 提取数据并填充 xlsx 模板生成实验报告。支持数据表行展开、逻辑判定计算。保存到项目工作区的 04_施工实施/实验检测报告/。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "template_name": {"type": "string", "description": "模板定义文件名（由 confirm_template_analysis 保存的）"},
                "ledger_paths_json": {"type": "string", "description": "台账文件路径 JSON 数组: [\"/path/to/实验台账.xlsx\"]"},
                "report_date": {"type": "string", "description": "报告日期 YYYY-MM-DD，默认今天"},
            },
            "required": ["project_name", "template_name", "ledger_paths_json"],
        },
    ),
    Tool(
        name="generate_single_experiment_report",
        description="用模板为特定检测项目生成单份实验报告。如只生成'钢筋原材力学性能检测报告'，LLM 会从台账中只提取该检测项目的数据。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "template_name": {"type": "string", "description": "模板定义文件名"},
                "test_item": {"type": "string", "description": "检测项目名称，如'钢筋原材力学性能检测'、'混凝土抗压强度检测'"},
                "ledger_paths_json": {"type": "string", "description": "台账文件路径 JSON 数组"},
                "report_date": {"type": "string", "description": "报告日期 YYYY-MM-DD，默认今天"},
                "extra_context": {"type": "string", "description": "额外说明，如检测部位、批次信息等"},
            },
            "required": ["project_name", "template_name", "test_item", "ledger_paths_json"],
        },
    ),
    Tool(
        name="list_experiment_reports",
        description="列出项目已生成的实验检测报告文件（xlsx/md）。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
            },
            "required": ["project_name"],
        },
    ),
    Tool(
        name="read_experiment_report",
        description="读取一篇实验报告内容。xlsx 格式报告以 Markdown 表格形式返回预览。",
        inputSchema={
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目名称"},
                "filename": {"type": "string", "description": "报告文件名"},
            },
            "required": ["project_name", "filename"],
        },
    ),
]

TOOL_MAP = {
    # 活动日志
    "read_template_cells": tool_read_template_cells,
    "fill_template_cells": tool_fill_template_cells,
    "get_activity_log": tool_get_activity_log,
    "get_activity_summary": tool_get_activity_summary,
    "check_wechat_setup": tool_check_setup,
    "install_wechat_cli": tool_install_wechat_cli,
    "list_wechat_groups": tool_list_wechat_groups,
    "setup_monitoring": tool_setup,
    "run_monitor": tool_run_monitor,
    "get_summary": tool_get_summary,
    "get_config": tool_get_config,
    "update_config": tool_update_config,
    # 分部分项解析
    "parse_project_document": tool_parse_project_document,
    "parse_project_text": tool_parse_project_text,
    "get_project_wbs": tool_get_project_wbs,
    # 项目初始化
    "project_init": tool_project_init,
    "project_list": tool_project_list,
    "project_load": tool_project_load,
    "project_update": tool_project_update,
    "project_discover": tool_project_discover,
    "template_analyze": tool_template_analyze,
    "template_save": tool_template_save,
    "template_list": tool_template_list,
    "template_apply": tool_template_apply,
    # 影像资料整理
    "scan_directory": tool_scan_directory,
    "organize_images": tool_organize_images,
    # 施工日志
    "generate_daily_log": tool_generate_daily_log,
    "generate_log_from_ledgers": tool_generate_log_from_ledgers,
    "list_logs": tool_list_logs,
    "read_log": tool_read_log,
    # 实验报告
    "analyze_experiment_template": tool_analyze_experiment_template,
    "confirm_template_analysis": tool_confirm_template_analysis,
    "list_experiment_templates": tool_list_experiment_templates,
    "generate_experiment_report": tool_generate_experiment_report,
    "generate_single_experiment_report": tool_generate_single_experiment_report,
    "list_experiment_reports": tool_list_experiment_reports,
    "read_experiment_report": tool_read_experiment_report,
}


@APP.list_tools()
async def list_tools():
    return TOOLS


@APP.call_tool()
async def call_tool(name: str, arguments: dict):
    import time as _time
    _start = _time.time()
    handler = TOOL_MAP.get(name)
    if not handler:
        raise ValueError(f"Unknown tool: {name}")

    try:
        result = await handler(**arguments)
        _elapsed = int((_time.time() - _start) * 1000)
        # 记录成功调用
        database.record_call(
            tool_name=name,
            params=arguments,
            result_summary=_summarize_result(name, result),
            status="success",
            duration_ms=_elapsed,
            module=database._classify_module(name),
        )
        return [TextContent(type="text", text=result)]
    except Exception as exc:
        _elapsed = int((_time.time() - _start) * 1000)
        database.record_call(
            tool_name=name,
            params=arguments,
            result_summary=str(exc)[:500],
            status="error",
            duration_ms=_elapsed,
            module=database._classify_module(name),
        )
        raise


def _summarize_result(tool_name: str, result: str) -> str:
    """从工具返回的 JSON 中提取摘要信息。"""
    if not result:
        return ""
    try:
        data = json.loads(result)
    except json.JSONDecodeError:
        return result[:200]

    parts = []
    if "error" in data:
        return f"[错误] {str(data['error'])[:200]}"
    if "template_name" in data:
        parts.append(f"模板: {data['template_name']}")
    if "sheet_count" in data:
        parts.append(f"{data['sheet_count']}个工作表")
    if "report_type" in data:
        parts.append(f"类型: {data['report_type']}")
    if "report_count" in data:
        parts.append(f"生成 {data['report_count']} 份报告")
    if "needs_confirmation" in data:
        parts.append(f"待确认: {data.get('confirmation_required', 0)} 个问题")
    if "total" in data:
        parts.append(f"共 {data['total']} 项")
    if "projects" in data:
        if isinstance(data.get("projects"), list):
            parts.append(f"{len(data['projects'])} 个项目")
    if "saved_to" in data:
        parts.append(f"已保存: {data['saved_to']}")
    if "saved" in data:
        parts.append(f"已保存: {data['saved']}")
    if "config" in data:
        parts.append("配置已更新")
    if "groups" in data:
        parts.append(f"{len(data.get('groups', []))} 个群聊")

    return "; ".join(parts) if parts else result[:200]


async def amain_stdio():
    async with stdio_server() as (read, write):
        await APP.run(read, write, APP.create_initialization_options())


def run_sse(host: str = "127.0.0.1", port: int = 8700):
    """SSE HTTP 模式：让 WSL/远程智能体通过 HTTP 连接 Windows 上的 MCP Server。"""
    try:
        from mcp.server.sse import SseServerTransport
        from starlette.applications import Starlette
        from starlette.routing import Route
        import uvicorn
    except ImportError as e:
        print(f"SSE 模式需要额外依赖: {e}")
        print("  pip install starlette uvicorn")
        import sys; sys.exit(1)

    sse = SseServerTransport("/messages/")

    async def handle_sse(request):
        async with sse.connect_sse(request.scope, request.receive, request._send) as streams:
            await APP.run(streams[0], streams[1], APP.create_initialization_options())

    async def handle_messages(request):
        await sse.handle_post_message(request.scope, request.receive, request._send)

    starlette_app = Starlette(routes=[
        Route("/sse", endpoint=handle_sse),
        Route("/messages/", endpoint=handle_messages, methods=["POST"]),
    ])

    print(f"木牛流马 MCP Server (SSE 模式)")
    print(f"监听: http://{host}:{port}")
    print(f"MCP 配置: {{\"url\": \"http://{host}:{port}/sse\"}}")
    uvicorn.run(starlette_app, host=host, port=port, log_level="warning")


def run():
    import asyncio
    import sys

    if "--sse" in sys.argv:
        host = "127.0.0.1"
        port = 8700
        for i, arg in enumerate(sys.argv):
            if arg == "--host" and i + 1 < len(sys.argv):
                host = sys.argv[i + 1]
            if arg == "--port" and i + 1 < len(sys.argv):
                port = int(sys.argv[i + 1])
        run_sse(host=host, port=port)
    else:
        asyncio.run(amain_stdio())


if __name__ == "__main__":
    run()
