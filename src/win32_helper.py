"""COM 自动化助手 — 通过 win32com 驱动本地 Word/Excel/WPS 操作文档。

路径策略：
- 原生 Windows Python → 直接 import win32com
- WSL 内 Python → 通过 Windows Python + com_bridge.py 桥接
- 不提供 openpyxl/python-docx 写入回退（会破坏格式）
"""

from __future__ import annotations

import json
import os
import platform
import subprocess
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Any

# win32com 仅在原生 Windows Python 上可用
try:
    import win32com.client as win32
    from win32com.client import Dispatch, constants
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

BRIDGE_SCRIPT = Path(__file__).parent / "com_bridge.py"


def _is_wsl() -> bool:
    """检测是否运行在 WSL 内。"""
    try:
        return "microsoft" in platform.uname().release.lower() or "WSL" in platform.uname().release
    except Exception:
        return False


def _find_windows_python() -> str | None:
    """查找可用的 Windows Python 解释器。"""
    candidates = [
        "python.exe",
        "python3.exe",
    ]
    for c in candidates:
        try:
            result = subprocess.run([c, "--version"], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                return c
        except Exception:
            continue
    return None


def _wsl_to_windows_path(wsl_path: str) -> str:
    """将 WSL 路径转为 Windows 路径。"""
    try:
        result = subprocess.run(
            ["wslpath", "-w", str(wsl_path)],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return wsl_path


def _call_bridge(cmd_args: list[str]) -> dict[str, Any]:
    """通过 Windows Python 调用 COM 桥接脚本。"""
    windows_py = _find_windows_python()
    if not windows_py:
        return {"error": "WSL 环境中未找到 Windows Python。请确认 Windows 上已安装 Python 并加入 PATH"}

    bridge_path = _wsl_to_windows_path(str(BRIDGE_SCRIPT))
    full_cmd = [windows_py, bridge_path] + cmd_args

    try:
        result = subprocess.run(full_cmd, capture_output=True, timeout=120)

        # Windows 中文环境 stdout 可能是 GBK，先试 UTF-8 再试 GBK
        stdout = ""
        stderr = ""
        for data, target in [(result.stdout, "stdout"), (result.stderr, "stderr")]:
            if data:
                for encoding in ("utf-8", "gbk", "gb2312", "latin-1"):
                    try:
                        decoded = data.decode(encoding)
                        if target == "stdout":
                            stdout = decoded
                        else:
                            stderr = decoded
                        break
                    except (UnicodeDecodeError, LookupError):
                        continue

        if result.returncode != 0:
            return {"error": f"COM 桥接失败: {stderr[:500]}"}
        return json.loads(stdout)
    except subprocess.TimeoutExpired:
        return {"error": "COM 操作超时（120秒），请检查 Office/WPS 是否正常"}
    except json.JSONDecodeError:
        return {"error": f"COM 桥接返回非 JSON: {stdout[:300] if stdout else 'empty'}"}
    except Exception as e:
        return {"error": f"桥接调用失败: {e}"}


# ── 环境检测 ────────────────────────────────────────────────────

def check_available() -> dict[str, Any]:
    """检测 COM 环境是否可用。"""
    if WIN32_AVAILABLE:
        word_app, word_name = _detect_word_app()
        excel_app, excel_name = _detect_excel_app()
        result = {
            "win32_available": True,
            "mode": "native_win32com",
            "wsl": False,
            "word": {"available": bool(word_app), "app_name": word_name},
            "excel": {"available": bool(excel_app), "app_name": excel_name},
        }
        if word_app:
            try: word_app.Quit()
            except Exception: pass
        if excel_app:
            try: excel_app.Quit()
            except Exception: pass
        return result

    if _is_wsl():
        win_py = _find_windows_python()
        return {
            "win32_available": False,
            "mode": "wsl_bridge",
            "wsl": True,
            "windows_python": win_py,
            "bridge_script": str(BRIDGE_SCRIPT),
            "word": {"available": win_py is not None, "method": "com_bridge via Windows Python"},
            "excel": {"available": win_py is not None, "method": "com_bridge via Windows Python"},
            "hint": "WSL 环境将通过 Windows Python 调用 COM 操作" if win_py
                    else "WSL 环境但未找到 Windows Python，请安装: winget install python",
        }

    return {
        "win32_available": False,
        "mode": "unsupported",
        "wsl": False,
        "error": "非 Windows 非 WSL 环境，Word/Excel COM 操作不可用。请将 MCP Server 部署到 Windows 或 WSL",
    }


def _detect_word_app():
    if not WIN32_AVAILABLE:
        return None, ""
    for prog_id, name in [("Kwps.Application", "WPS Word"), ("Word.Application", "Microsoft Word")]:
        try:
            app = Dispatch(prog_id)
            app.Visible = False
            return app, name
        except Exception:
            continue
    return None, ""


def _detect_excel_app():
    if not WIN32_AVAILABLE:
        return None, ""
    for prog_id, name in [("Ket.Application", "WPS Excel"), ("Excel.Application", "Microsoft Excel")]:
        try:
            app = Dispatch(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            return app, name
        except Exception:
            continue
    return None, ""


# ── Word 操作 ───────────────────────────────────────────────────

def word_fill_placeholders(
    template_path: str | Path,
    output_path: str | Path,
    data: dict[str, str],
) -> Path:
    """替换 Word 模板中的 {{key}} 占位符并另存。仅通过 COM 操作，无回退。"""
    template_path = Path(template_path).resolve()
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if WIN32_AVAILABLE:
        return _word_fill_native(template_path, output_path, data)

    if _is_wsl():
        # 通过 Windows Python 桥接
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
            data_file = f.name

        try:
            result = _call_bridge([
                "word_fill",
                _wsl_to_windows_path(str(template_path)),
                _wsl_to_windows_path(str(output_path)),
                _wsl_to_windows_path(data_file),
            ])
            if "error" in result:
                raise RuntimeError(result["error"])
            return output_path
        finally:
            try: os.unlink(data_file)
            except Exception: pass

    raise RuntimeError(
        "Word 模板填充需要 Windows + Office/WPS。当前环境不支持 COM 操作。\n"
        "解决方案：\n"
        "  1. 在 Windows 上运行 MCP Server\n"
        "  2. 或在 WSL 中运行（需 Windows Python + Office/WPS 已安装）"
    )


def _word_fill_native(template_path: Path, output_path: Path, data: dict[str, str]) -> Path:
    app, app_name = _detect_word_app()
    if not app:
        raise RuntimeError("未检测到 Word 或 WPS")
    app.Visible = False
    app.DisplayAlerts = False
    doc = None
    try:
        doc = app.Documents.Open(str(template_path))
        story_ranges = _get_all_story_ranges(doc)
        for story_range in story_ranges:
            find = story_range.Find
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                find.Text = placeholder
                find.Replacement.Text = str(value) if value else ""
                find.Forward = True
                find.Wrap = 1
                find.Format = False
                find.MatchCase = True
                find.MatchWholeWord = False
                try:
                    find.Execute(Replace=2)
                except Exception:
                    pass
        doc.SaveAs2(str(output_path))
        return output_path
    finally:
        try:
            if doc: doc.Close(SaveChanges=False)
        except Exception: pass
        try: app.Quit()
        except Exception: pass


def _get_all_story_ranges(doc):
    ranges = [doc.Content]
    try:
        for i in range(1, 12):
            try:
                story = doc.StoryRanges(i)
                while story:
                    ranges.append(story)
                    story = story.NextStoryRange
            except Exception:
                continue
    except Exception:
        pass
    return ranges


# ── Excel 操作 ──────────────────────────────────────────────────

@contextmanager
def open_excel(template_path: str | Path):
    """上下文管理器：打开 Excel 工作簿。

    原生 Windows → win32com 直接操作
    WSL → Windows Python 桥接（仅支持读取和填充，不支持交互式操作）
    """
    template_path = Path(template_path).resolve()
    if WIN32_AVAILABLE:
        app, app_name = _detect_excel_app()
        if not app:
            raise RuntimeError("未检测到 Excel 或 WPS 表格")
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        wb = None
        try:
            wb = app.Workbooks.Open(str(template_path))
            yield app, wb, app_name
        finally:
            try:
                if wb: wb.Close(SaveChanges=False)
            except Exception: pass
            try: app.Quit()
            except Exception: pass
    else:
        raise RuntimeError(
            "Excel 操作需要 Windows + Office/WPS。当前环境不支持。\n"
            "请在 Windows 或 WSL 中运行 MCP Server。"
        )


def excel_read_structure(template_path: str | Path) -> dict[str, Any]:
    """读取 xlsx 完整结构信息。通过 COM 操作。"""
    template_path = Path(template_path).resolve()
    if not template_path.exists():
        return {"error": f"文件不存在: {template_path}"}

    if WIN32_AVAILABLE:
        return _excel_read_structure_native(template_path)
    if _is_wsl():
        return _call_bridge(["excel_read_structure", _wsl_to_windows_path(str(template_path))])

    raise RuntimeError("Excel 读取需要 Windows + Office/WPS")


def _excel_read_structure_native(file_path: Path) -> dict[str, Any]:
    app, app_name = _detect_excel_app()
    if not app:
        raise RuntimeError("未检测到 Excel 或 WPS")
    app.Visible = False
    app.DisplayAlerts = False
    structure = {"filename": file_path.name, "sheets": []}
    wb = None
    try:
        wb = app.Workbooks.Open(str(file_path))
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            sheet_info = {"name": ws.Name, "rows": [], "merged_ranges": []}
            try:
                used = ws.UsedRange
                max_row = min(used.Rows.Count, 100)
                max_col = min(used.Columns.Count, 50)
                for row_idx in range(1, max_row + 1):
                    cells_info = []
                    for col_idx in range(1, max_col + 1):
                        cell = ws.Cells(row_idx, col_idx)
                        value = cell.Value
                        if value is not None:
                            col_letter = _col_to_letter(col_idx)
                            cells_info.append({
                                "cell": f"{col_letter}{row_idx}",
                                "value": str(value)[:200],
                                "column": col_letter,
                                "row": row_idx,
                            })
                            if cell.MergeCells:
                                merge_range = cell.MergeArea.Address.replace("$", "")
                                if merge_range not in sheet_info["merged_ranges"]:
                                    sheet_info["merged_ranges"].append(merge_range)
                    if cells_info:
                        sheet_info["rows"].append({"row": row_idx, "cells": cells_info})
            except Exception as e:
                sheet_info["error"] = str(e)
            structure["sheets"].append(sheet_info)
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
        except Exception: pass
        try: app.Quit()
        except Exception: pass
    return structure


def excel_save_as(wb, output_path: str | Path) -> Path:
    """工作簿另存为。"""
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if WIN32_AVAILABLE:
        wb.SaveAs(str(output_path))
    else:
        raise RuntimeError("Excel 另存需要 COM，当前环境不支持")
    return output_path


def excel_write_cell(wb, sheet_name: str, cell_ref: str, value: Any) -> None:
    """向指定单元格写入值。"""
    if WIN32_AVAILABLE:
        ws = wb.Worksheets(sheet_name)
        ws.Range(cell_ref).Value = value
    else:
        raise RuntimeError("Excel 写入需要 COM，当前环境不支持")


def excel_fill_template(
    template_path: str | Path,
    output_path: str | Path,
    fill_data: dict[str, Any],
) -> Path:
    """填充 xlsx 模板并另存。通过 COM 操作，保留所有格式。

    Args:
        template_path: 模板 .xlsx 路径
        output_path: 输出 .xlsx 路径
        fill_data: {'sheets': [...], 'report_data': {...}, 'data_table_rows': [...], 'logic_rules': [...]}
    """
    template_path = Path(template_path).resolve()
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if WIN32_AVAILABLE:
        return _excel_fill_native(template_path, output_path, fill_data)

    if _is_wsl():
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump(fill_data, f, ensure_ascii=False)
            data_file = f.name
        try:
            result = _call_bridge([
                "excel_fill",
                _wsl_to_windows_path(str(template_path)),
                _wsl_to_windows_path(str(output_path)),
                _wsl_to_windows_path(data_file),
            ])
            if "error" in result:
                raise RuntimeError(result["error"])
            return output_path
        finally:
            try: os.unlink(data_file)
            except Exception: pass

    raise RuntimeError("Excel 模板填充需要 Windows + Office/WPS。当前环境不支持。")


def _excel_fill_native(template_path: Path, output_path: Path, data: dict[str, Any]) -> Path:
    app, app_name = _detect_excel_app()
    if not app:
        raise RuntimeError("未检测到 Excel 或 WPS")
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    wb = None
    try:
        wb = app.Workbooks.Open(str(template_path))
        _fill_sheets(wb, data)
        wb.SaveAs(str(output_path))
        return output_path
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
        except Exception: pass
        try: app.Quit()
        except Exception: pass


def _fill_sheets(wb, data: dict[str, Any]) -> None:
    """在 COM 工作簿中填充数据（安全写入：合并单元格写左上角）。"""
    sheets_def = data.get("sheets", [])
    report_data = data.get("report_data", {})
    data_rows = data.get("data_table_rows", [])

    for si, sheet_def in enumerate(sheets_def):
        sname = sheet_def.get("sheet_name", "")
        try:
            ws = wb.Worksheets(sname) if sname else wb.Worksheets(si + 1)
        except Exception:
            continue

        for region in sheet_def.get("regions", []):
            cells = region.get("cells", "")
            rid = region.get("id", "")
            value = report_data.get(rid, "")
            if not value and region.get("fixed_value"):
                value = region["fixed_value"]
            if not value:
                continue
            target_cell = cells.split(":")[0] if ":" in cells else cells
            try:
                _com_safe_write(ws, target_cell, value)
            except Exception:
                pass

        dt = sheet_def.get("data_table", {})
        if dt and data_rows and sheet_def.get("sheet_role") == "main_report":
            header_row = dt.get("data_start_row", dt.get("header_row", 5) + 1)
            columns = dt.get("columns", [])
            for i, rd in enumerate(data_rows):
                cr = header_row + i
                for col_def in columns:
                    col_letter = col_def.get("col_letter", "")
                    if not col_letter:
                        continue
                    value = rd.get(col_letter, "")
                    try:
                        _com_safe_write(ws, f"{col_letter}{cr}", value)
                    except Exception:
                        pass

        for rule in data.get("logic_rules", []):
            if rule.get("confidence") != "high":
                continue
            applies_to = rule.get("applies_to", "")
            depends_on = rule.get("depends_on", [])
            if len(depends_on) < 2:
                continue
            rule_sheet_name = rule.get("applies_in_sheet", "")
            if rule_sheet_name and rule_sheet_name != sname:
                continue
            dep_cols = [d.split(".")[-1].replace("col_letter_", "") for d in depends_on]
            target_col = applies_to.replace("data_table.", "").replace("col_letter_", "")
            row_start = dt.get("data_start_row", dt.get("header_row", 5) + 1)
            for i, rd in enumerate(data_rows):
                cr = row_start + i
                try:
                    measured = float(str(rd.get(dep_cols[0], "0")))
                    required = float(str(rd.get(dep_cols[1], "0")))
                    conclusion = "合格" if measured >= required else "不合格"
                    _com_safe_write(ws, f"{target_col}{cr}", conclusion)
                except (ValueError, TypeError):
                    pass


def _com_safe_write(ws, cell_ref: str, value) -> None:
    """安全写入单元格。若在合并区域内，写入左上角。"""
    rng = ws.Range(cell_ref)
    try:
        if rng.MergeCells:
            rng.MergeArea.Cells(1, 1).Value = value
        else:
            rng.Value = value
    except Exception:
        rng.Value = value


def excel_read_ledger(file_path: str | Path, max_row: int = 500) -> dict[str, Any]:
    """读取 xlsx 台账数据。通过 COM 操作。"""
    file_path = Path(file_path).resolve()
    if not file_path.exists():
        return {"error": f"文件不存在: {file_path}"}

    if WIN32_AVAILABLE:
        return _excel_read_ledger_native(file_path, max_row)
    if _is_wsl():
        return _call_bridge(["excel_read", _wsl_to_windows_path(str(file_path)), str(max_row)])

    raise RuntimeError("台账读取需要 Windows + Office/WPS")


def _excel_read_ledger_native(file_path: Path, max_row: int) -> dict[str, Any]:
    app, app_name = _detect_excel_app()
    if not app:
        raise RuntimeError("未检测到 Excel 或 WPS")
    app.Visible = False
    app.DisplayAlerts = False
    sheets = []
    wb = None
    try:
        wb = app.Workbooks.Open(str(file_path))
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            rows = []
            try:
                used = ws.UsedRange
                r_max = min(used.Rows.Count, max_row)
                c_max = min(used.Columns.Count, 50)
                for row_idx in range(1, r_max + 1):
                    cells = []
                    for col_idx in range(1, c_max + 1):
                        value = ws.Cells(row_idx, col_idx).Value
                        cells.append(str(value) if value is not None else "")
                    if any(cells):
                        rows.append(cells)
            except Exception:
                pass
            sheets.append({"name": ws.Name, "rows": rows})
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
        except Exception: pass
        try: app.Quit()
        except Exception: pass
    return {"file": str(file_path), "sheets": sheets}


def _col_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result
