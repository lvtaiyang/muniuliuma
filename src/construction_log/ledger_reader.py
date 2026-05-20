"""台账读取：解析材料进场台账、实验检测台账等施工文档。"""

from __future__ import annotations

from pathlib import Path
from typing import Any


def read_ledger(file_path: str | Path) -> dict[str, Any]:
    """读取一个台账文件，返回结构化内容。

    支持: .xlsx, .xlsm, .csv, .yaml, .json, .txt, .md
    """
    path = Path(file_path)
    if not path.exists():
        return {"error": f"文件不存在: {path}", "file": str(path)}

    ext = path.suffix.lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            return {"file": str(path), "sheets": _read_xlsx(path)}
        elif ext == ".csv":
            return {"file": str(path), "content": _read_csv(path)}
        elif ext in (".yaml", ".yml"):
            return {"file": str(path), "content": _read_yaml(path)}
        elif ext == ".json":
            return {"file": str(path), "content": _read_json(path)}
        elif ext in (".txt", ".md", ".markdown"):
            return {"file": str(path), "content": _read_text(path)}
        else:
            return {"error": f"不支持的文件格式: {ext}", "file": str(path)}
    except ImportError as e:
        return {"error": f"缺少依赖: {e}", "file": str(path), "hint": "pip install openpyxl pyyaml"}
    except Exception as e:
        return {"error": str(e), "file": str(path)}


def read_ledgers(file_paths: list[str]) -> list[dict[str, Any]]:
    """批量读取多个台账文件。"""
    return [read_ledger(p) for p in file_paths]


def _read_xlsx(path: Path) -> list[dict[str, Any]]:
    from .. import win32_helper
    result = win32_helper.excel_read_ledger(path)
    if "error" in result:
        raise RuntimeError(result["error"])
    return result.get("sheets", [])


def _read_csv(path: Path) -> list[list[str]]:
    import csv
    with open(path, encoding="utf-8-sig", errors="replace") as f:
        reader = csv.reader(f)
        return [row for row in reader if any(row)]


def _read_yaml(path: Path) -> Any:
    import yaml
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def _read_json(path: Path) -> Any:
    import json
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def format_ledger_for_llm(ledger_data: dict[str, Any]) -> str:
    """将台账数据格式化为 LLM 可读的文本。"""
    if "error" in ledger_data:
        return f"[读取失败] {ledger_data['file']}: {ledger_data['error']}"

    lines = [f"### 文件: {ledger_data.get('file', '')}", ""]

    if "sheets" in ledger_data:
        for sheet in ledger_data["sheets"]:
            lines.append(f"#### 工作表: {sheet['name']}")
            for row in sheet["rows"][:200]:
                lines.append(" | ".join(row))
            lines.append("")

    elif "content" in ledger_data:
        content = ledger_data["content"]
        if isinstance(content, list):
            for row in content[:200]:
                lines.append(" | ".join(str(c) for c in row) if isinstance(row, list) else str(row))
        elif isinstance(content, dict):
            import yaml
            lines.append(yaml.dump(content, allow_unicode=True))
        else:
            lines.append(str(content)[:5000])

    return "\n".join(lines)
