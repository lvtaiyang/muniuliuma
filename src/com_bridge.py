"""COM 操作桥接脚本 — 由 Windows Python 执行。

WSL 中的 MCP Server 通过此脚本调用 Windows COM 操作 Office/WPS。
独立运行，不依赖项目其他模块。

用法:
    python.exe com_bridge.py excel_read_structure <file.xlsx>
    python.exe com_bridge.py excel_read <file.xlsx> [max_row]
    python.exe com_bridge.py excel_fill <template> <output> <data.json>
    python.exe com_bridge.py word_fill <template> <output> <data.json>
"""

import json
import sys
import os

try:
    import win32com.client
    from win32com.client import Dispatch
except ImportError:
    print(json.dumps({"error": "pywin32 未安装，请运行: pip install pywin32"}))
    sys.exit(1)


def _detect_excel():
    for prog_id, name in [("Ket.Application", "WPS"), ("Excel.Application", "MS Excel")]:
        try:
            app = Dispatch(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            return app, name
        except Exception:
            continue
    return None, ""


def _detect_word():
    for prog_id, name in [("Kwps.Application", "WPS"), ("Word.Application", "MS Word")]:
        try:
            app = Dispatch(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            return app, name
        except Exception:
            continue
    return None, ""


def _col_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ── 操作 1: 读取 xlsx 结构 ──────────────────────────────────────

def cmd_excel_read_structure(file_path: str):
    if not os.path.exists(file_path):
        print(json.dumps({"error": f"文件不存在: {file_path}"}, ensure_ascii=False))
        return

    app, app_name = _detect_excel()
    if not app:
        print(json.dumps({"error": "未检测到 Excel 或 WPS 表格"}, ensure_ascii=False))
        return

    structure = {"filename": os.path.basename(file_path), "sheets": []}
    wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(file_path))
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
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

    print(json.dumps(structure, ensure_ascii=False))


# ── 操作 2: 读取 xlsx 台账 ─────────────────────────────────────

def cmd_excel_read(file_path: str, max_row: int = 500):
    if not os.path.exists(file_path):
        print(json.dumps({"error": f"文件不存在: {file_path}"}, ensure_ascii=False))
        return

    app, app_name = _detect_excel()
    if not app:
        print(json.dumps({"error": "未检测到 Excel 或 WPS"}, ensure_ascii=False))
        return

    sheets = []
    wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(file_path))
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
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

    print(json.dumps({"file": file_path, "sheets": sheets}, ensure_ascii=False))


# ── 操作 3: 填充 xlsx 模板 ─────────────────────────────────────


def _safe_range_write(ws, cell_ref: str, value) -> None:
    """安全写入。合并单元格写左上角。"""
    rng = ws.Range(cell_ref)
    try:
        if rng.MergeCells:
            rng.MergeArea.Cells(1, 1).Value = value
        else:
            rng.Value = value
    except Exception:
        rng.Value = value

def cmd_excel_fill(template_path: str, output_path: str, data_json_path: str):
    if not os.path.exists(template_path):
        print(json.dumps({"error": f"模板不存在: {template_path}"}, ensure_ascii=False))
        return

    with open(data_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    app, app_name = _detect_excel()
    if not app:
        print(json.dumps({"error": "未检测到 Excel 或 WPS"}, ensure_ascii=False))
        return

    wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(template_path))

        sheets_def = data.get("sheets", [])
        report_data = data.get("report_data", {})
        data_rows = data.get("data_table_rows", [])

        for si, sheet_def in enumerate(sheets_def):
            sname = sheet_def.get("sheet_name", "")
            try:
                ws = wb.Worksheets(sname) if sname else wb.Worksheets(si + 1)
            except Exception:
                continue

            # 填充固定区域
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
                    _safe_range_write(ws, target_cell, value)
                except Exception:
                    pass

            # 填充数据表
            dt = sheet_def.get("data_table", {})
            if dt and data_rows and sheet_def.get("sheet_role") == "main_report":
                header_row = dt.get("data_start_row", dt.get("header_row", 5) + 1)
                columns = dt.get("columns", [])

                for i, row_data in enumerate(data_rows):
                    cr = header_row + i
                    for col_def in columns:
                        col_letter = col_def.get("col_letter", "")
                        if not col_letter:
                            continue
                        value = row_data.get(col_letter, "")
                        try:
                            _safe_range_write(ws, f"{col_letter}{cr}", value)
                        except Exception:
                            pass

            # 应用逻辑规则
            for rule in data.get("logic_rules", []):
                if rule.get("confidence") != "high":
                    continue
                applies_to = rule.get("applies_to", "")
                depends_on = rule.get("depends_on", [])
                if len(depends_on) < 2:
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
                        _safe_range_write(ws, f"{target_col}{cr}", conclusion)
                    except (ValueError, TypeError):
                        pass

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.SaveAs(os.path.abspath(output_path))
        print(json.dumps({"saved": output_path}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False))
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


# ── 操作 4: 填充 Word 模板 ─────────────────────────────────────

def cmd_word_fill(template_path: str, output_path: str, data_json_path: str):
    if not os.path.exists(template_path):
        print(json.dumps({"error": f"模板不存在: {template_path}"}, ensure_ascii=False))
        return

    with open(data_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    app, app_name = _detect_word()
    if not app:
        print(json.dumps({"error": "未检测到 Word 或 WPS"}, ensure_ascii=False))
        return

    doc = None
    try:
        doc = app.Documents.Open(os.path.abspath(template_path))
        story_ranges = [doc.Content]
        try:
            for i in range(1, 12):
                try:
                    story = doc.StoryRanges(i)
                    while story:
                        story_ranges.append(story)
                        story = story.NextStoryRange
                except Exception:
                    continue
        except Exception:
            pass

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
                    find.Execute(Replace=2)  # wdReplaceAll
                except Exception:
                    pass

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.SaveAs2(os.path.abspath(output_path))
        print(json.dumps({"saved": output_path}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False))
    finally:
        try:
            if doc:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


# ── 主入口 ─────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "缺少命令"}, ensure_ascii=False))
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "excel_read_structure":
        cmd_excel_read_structure(sys.argv[2])

    elif cmd == "excel_read":
        max_row = int(sys.argv[3]) if len(sys.argv) > 3 else 500
        cmd_excel_read(sys.argv[2], max_row)

    elif cmd == "excel_fill":
        cmd_excel_fill(sys.argv[2], sys.argv[3], sys.argv[4])

    elif cmd == "word_fill":
        cmd_word_fill(sys.argv[2], sys.argv[3], sys.argv[4])

    else:
        print(json.dumps({"error": f"未知命令: {cmd}"}, ensure_ascii=False))
        sys.exit(1)
