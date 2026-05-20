"""活动日志 — 已迁入 database.py，本文件仅保留兼容接口。"""

from __future__ import annotations

from . import database

_log_call = database.record_call
_query_logs = database.query_calls
_get_summary = database.get_call_summary
_classify_module = database._classify_module


def log_call(tool_name: str, params=None, result_summary="", status="success",
             duration_ms=0, module="", session_id=""):
    return _log_call(tool_name, params, result_summary, status, duration_ms, module, session_id)


def query_logs(limit=20, tool_name="", module="", status="", since="", session_id=""):
    return _query_logs(limit, tool_name, module, status, since, session_id)


def get_summary(since="24h", session_id=""):
    return _get_summary(since, session_id)
