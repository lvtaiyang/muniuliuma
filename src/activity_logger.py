"""活动日志 — 用 SQLite 记录每次 MCP 工具调用，给智能体提供"短期记忆"。

纯标准库，零额外依赖。日志存储在 ~/.muniuliuma/activity.db。
"""

from __future__ import annotations

import json
import sqlite3
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

DB_DIR = Path.home() / ".muniuliuma"
DB_PATH = DB_DIR / "activity.db"

_local = threading.local()


def _get_conn() -> sqlite3.Connection:
    """获取线程本地数据库连接。"""
    if not hasattr(_local, "conn") or _local.conn is None:
        DB_DIR.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(DB_PATH))
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        _local.conn = conn
        _init_tables(conn)
    return _local.conn


def _init_tables(conn: sqlite3.Connection) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS calls (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tool_name TEXT NOT NULL,
            module TEXT DEFAULT '',
            params_json TEXT DEFAULT '{}',
            result_summary TEXT DEFAULT '',
            status TEXT DEFAULT 'success',
            duration_ms INTEGER DEFAULT 0,
            called_at TEXT NOT NULL,
            session_id TEXT DEFAULT ''
        )
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_calls_tool ON calls(tool_name)
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_calls_time ON calls(called_at)
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_calls_session ON calls(session_id)
    """)
    conn.commit()


def log_call(
    tool_name: str,
    params: dict[str, Any] | None = None,
    result_summary: str = "",
    status: str = "success",
    duration_ms: int = 0,
    module: str = "",
    session_id: str = "",
) -> int:
    """记录一次工具调用。返回记录 ID。

    Args:
        tool_name: 工具名称
        params: 调用参数（敏感字段自动脱敏）
        result_summary: 结果摘要（截取前 500 字符）
        status: success / error
        duration_ms: 执行耗时（毫秒）
        module: 所属模块名
        session_id: 会话标识（同一智能体会话共享）
    """
    conn = _get_conn()
    safe_params = _sanitize_params(params or {})
    summary = (result_summary or "")[:500]

    cursor = conn.execute(
        """INSERT INTO calls (tool_name, module, params_json, result_summary, status, duration_ms, called_at, session_id)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            tool_name,
            module,
            json.dumps(safe_params, ensure_ascii=False),
            summary,
            status,
            duration_ms,
            datetime.now().isoformat(timespec="seconds"),
            session_id,
        ),
    )
    conn.commit()
    return cursor.lastrowid


def query_logs(
    limit: int = 20,
    tool_name: str = "",
    module: str = "",
    status: str = "",
    since: str = "",
    session_id: str = "",
) -> list[dict[str, Any]]:
    """查询最近的调用记录。

    Args:
        limit: 返回条数上限
        tool_name: 按工具名筛选
        module: 按模块筛选
        status: 按状态筛选 (success/error)
        since: 起始时间，如 "1h" "24h" "7d"
        session_id: 按会话筛选
    """
    conn = _get_conn()
    conditions: list[str] = []
    params: list[Any] = []

    if tool_name:
        conditions.append("tool_name = ?")
        params.append(tool_name)
    if module:
        conditions.append("module = ?")
        params.append(module)
    if status:
        conditions.append("status = ?")
        params.append(status)
    if since:
        cutoff = _parse_since(since)
        if cutoff:
            conditions.append("called_at >= ?")
            params.append(cutoff.isoformat(timespec="seconds"))
    if session_id:
        conditions.append("session_id = ?")
        params.append(session_id)

    where = (" WHERE " + " AND ".join(conditions)) if conditions else ""
    query = f"SELECT * FROM calls{where} ORDER BY called_at DESC LIMIT ?"
    params.append(limit)

    rows = conn.execute(query, params).fetchall()
    return [dict(r) for r in rows]


def get_summary(since: str = "24h", session_id: str = "") -> dict[str, Any]:
    """获取运行摘要。

    Returns:
        {
          "total_calls": 总调用次数,
          "success_rate": "成功率",
          "by_tool": {"tool_name": count, ...},
          "by_module": {"module": count, ...},
          "errors": [{"tool": ..., "summary": ...}, ...],
          "recent_tools": ["最近使用的5个工具"],
          "period": "统计时间范围"
        }
    """
    conn = _get_conn()
    cutoff = _parse_since(since)
    cutoff_str = cutoff.isoformat(timespec="seconds") if cutoff else ""

    params: list[Any] = []
    session_filter = ""
    if session_id:
        session_filter = " AND session_id = ?"
        params = [session_id]
    else:
        params = []

    if cutoff_str:
        time_filter = f" AND called_at >= ?"
        time_params = [cutoff_str]
    else:
        time_filter = ""
        time_params = []

    all_params = params + time_params

    # 总调用次数
    row = conn.execute(
        f"SELECT COUNT(*) as cnt FROM calls WHERE 1=1{session_filter}{time_filter}",
        all_params,
    ).fetchone()
    total = row["cnt"] if row else 0

    # 错误计数
    err_params = all_params + []
    row = conn.execute(
        f"SELECT COUNT(*) as cnt FROM calls WHERE status='error'{session_filter}{time_filter}",
        all_params,
    ).fetchone()
    errors_count = row["cnt"] if row else 0

    # 按工具统计
    rows = conn.execute(
        f"SELECT tool_name, COUNT(*) as cnt FROM calls WHERE 1=1{session_filter}{time_filter} GROUP BY tool_name ORDER BY cnt DESC",
        all_params,
    ).fetchall()
    by_tool = {r["tool_name"]: r["cnt"] for r in rows}

    # 按模块统计
    rows = conn.execute(
        f"SELECT module, COUNT(*) as cnt FROM calls WHERE 1=1{session_filter}{time_filter} GROUP BY module ORDER BY cnt DESC",
        all_params,
    ).fetchall()
    by_module = {r["module"]: r["cnt"] for r in rows if r["module"]}

    # 最近的错误
    error_rows = conn.execute(
        f"SELECT tool_name, result_summary, called_at FROM calls WHERE status='error'{session_filter}{time_filter} ORDER BY called_at DESC LIMIT 5",
        all_params,
    ).fetchall()
    errors = [{"tool": r["tool_name"], "summary": r["result_summary"], "time": r["called_at"]} for r in error_rows]

    # 最近使用的工具
    recent_rows = conn.execute(
        f"SELECT DISTINCT tool_name FROM calls WHERE 1=1{session_filter}{time_filter} ORDER BY called_at DESC LIMIT 5",
        all_params,
    ).fetchall()
    recent_tools = [r["tool_name"] for r in recent_rows]

    # 会话数
    row = conn.execute(
        f"SELECT COUNT(DISTINCT session_id) as cnt FROM calls WHERE session_id != ''{session_filter}{time_filter}",
        all_params,
    ).fetchall()
    active_sessions = row[0]["cnt"] if row else 0

    return {
        "total_calls": total,
        "errors_count": errors_count,
        "success_rate": f"{(total - errors_count) / total * 100:.1f}%" if total > 0 else "N/A",
        "by_tool": by_tool,
        "by_module": by_module,
        "errors": errors,
        "recent_tools": recent_tools,
        "active_sessions": active_sessions,
        "period": f"过去 {since}" if since else "全部",
    }


def _sanitize_params(params: dict[str, Any]) -> dict[str, Any]:
    """脱敏参数：遮蔽 api_key, token, password 等。"""
    sensitive_keys = {"api_key", "token", "password", "secret", "key", "auth"}
    safe = {}
    for k, v in params.items():
        if any(s in k.lower() for s in sensitive_keys):
            s = str(v)
            safe[k] = s[:4] + "****" if len(s) > 4 else "****"
        elif isinstance(v, str) and len(v) > 200:
            safe[k] = v[:200] + "..."
        else:
            safe[k] = v
    return safe


def _parse_since(since: str) -> datetime | None:
    """解析时间范围字符串：1h, 24h, 7d, 30d。"""
    if not since:
        return None
    now = datetime.now()
    match = __import__("re").match(r"(\d+)([hd])", since.lower())
    if match:
        num = int(match.group(1))
        unit = match.group(2)
        if unit == "h":
            return now - timedelta(hours=num)
        elif unit == "d":
            return now - timedelta(days=num)
    return None


def _classify_module(tool_name: str) -> str:
    """根据工具名推断所属模块。"""
    mapping = {
        "wechat": "wechat_monitor",
        "monitor": "wechat_monitor",
        "project": "project_initializer",
        "template_analyze": "project_initializer",
        "template_save": "project_initializer",
        "template_list": "project_initializer",
        "template_apply": "project_initializer",
        "scan_directory": "image_organizer",
        "organize_images": "image_organizer",
        "daily_log": "construction_log",
        "log_from_ledgers": "construction_log",
        "list_logs": "construction_log",
        "read_log": "construction_log",
        "experiment": "experiment_report",
        "get_config": "system",
        "update_config": "system",
        "check_wechat": "system",
        "install_wechat": "system",
        "setup": "system",
        "get_summary": "system",
    }
    for key, mod in mapping.items():
        if key in tool_name:
            return mod
    return "other"
