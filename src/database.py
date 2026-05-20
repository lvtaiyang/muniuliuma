"""统一数据层 — SQLite 单文件存储项目、模板、归档、报告、调用日志。

数据库: ~/.muniuliuma/muniuliuma.db
替代: activity.db + projects/*.yaml + config.yaml projects 数组
"""

from __future__ import annotations

import json
import sqlite3
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

DB_DIR = Path.home() / ".muniuliuma"
DB_PATH = DB_DIR / "muniuliuma.db"

_local = threading.local()


def _get_conn() -> sqlite3.Connection:
    """线程本地数据库连接。"""
    if not hasattr(_local, "conn") or _local.conn is None:
        DB_DIR.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(DB_PATH))
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA foreign_keys=ON")
        _local.conn = conn
        _init_tables(conn)
        _migrate_if_needed(conn)
    return _local.conn


def _init_tables(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            workspace TEXT NOT NULL,
            project_type TEXT DEFAULT '',
            current_stage TEXT DEFAULT '',
            meta_json TEXT DEFAULT '{}',
            stages_json TEXT DEFAULT '[]',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS project_categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL REFERENCES projects(id),
            category TEXT NOT NULL,
            UNIQUE(project_id, category)
        );

        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT NOT NULL,
            template_name TEXT NOT NULL,
            report_type TEXT DEFAULT '',
            definition_json TEXT DEFAULT '{}',
            source_file TEXT DEFAULT '',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS archives (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT NOT NULL,
            category TEXT DEFAULT '',
            filename TEXT NOT NULL,
            path TEXT NOT NULL,
            original_path TEXT DEFAULT '',
            description TEXT DEFAULT '',
            quality_notes TEXT DEFAULT '',
            analysis_json TEXT DEFAULT '{}',
            archived_at TEXT NOT NULL
        );

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
        );

        CREATE INDEX IF NOT EXISTS idx_projects_name ON projects(name);
        CREATE INDEX IF NOT EXISTS idx_templates_project ON templates(project_name);
        CREATE INDEX IF NOT EXISTS idx_archives_project ON archives(project_name);
        CREATE INDEX IF NOT EXISTS idx_archives_category ON archives(project_name, category);
        CREATE INDEX IF NOT EXISTS idx_calls_tool ON calls(tool_name);
        CREATE INDEX IF NOT EXISTS idx_calls_time ON calls(called_at);
        CREATE INDEX IF NOT EXISTS idx_calls_session ON calls(session_id);
    """)
    conn.commit()


def _migrate_if_needed(conn: sqlite3.Connection) -> None:
    """从旧存储迁移数据到 SQLite。"""
    # 1. 迁移 activity.db → muniuliuma.db
    old_db = DB_DIR / "activity.db"
    if old_db.exists():
        try:
            old_conn = sqlite3.connect(str(old_db))
            old_conn.row_factory = sqlite3.Row
            rows = old_conn.execute("SELECT * FROM calls").fetchall()
            if rows:
                for row in rows:
                    try:
                        conn.execute(
                            "INSERT OR IGNORE INTO calls (id, tool_name, module, params_json, result_summary, status, duration_ms, called_at, session_id) VALUES (?,?,?,?,?,?,?,?,?)",
                            [row[c] for c in ["id", "tool_name", "module", "params_json", "result_summary", "status", "duration_ms", "called_at", "session_id"]],
                        )
                    except Exception:
                        pass
                conn.commit()
            old_conn.close()
            old_db.rename(old_db.with_suffix(".db.bak"))
        except Exception:
            pass

    # 2. 迁移 projects/*.yaml → SQLite
    projects_dir = DB_DIR / "projects"
    if projects_dir.exists():
        for f in projects_dir.glob("*.yaml"):
            if f.name.startswith("_"):
                continue
            try:
                import yaml
                with open(f, encoding="utf-8") as fh:
                    data = yaml.safe_load(fh) or {}
                if data.get("name"):
                    _save_project_from_yaml(conn, data)
            except Exception:
                pass
        # 备份
        try:
            for f in projects_dir.glob("*.yaml"):
                f.rename(f.with_suffix(".yaml.bak"))
        except Exception:
            pass


def _save_project_from_yaml(conn: sqlite3.Connection, data: dict[str, Any]) -> None:
    now = datetime.now().isoformat()
    name = data.get("name", "")
    if not name:
        return

    meta = data.get("meta", {})
    categories = []
    if isinstance(meta, dict):
        cats = meta.pop("categories", None) or meta.pop("_categories", None) or []
        if isinstance(cats, list):
            categories = cats

    conn.execute(
        """INSERT OR REPLACE INTO projects (name, workspace, project_type, current_stage, meta_json, stages_json, created_at, updated_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            name,
            data.get("workspace", ""),
            data.get("type", data.get("project_type", "")),
            data.get("current_stage", ""),
            json.dumps(meta, ensure_ascii=False),
            json.dumps(data.get("stages", []), ensure_ascii=False),
            data.get("created_at", now),
            data.get("updated_at", now),
        ),
    )

    proj_id = conn.execute("SELECT id FROM projects WHERE name=?", (name,)).fetchone()
    if proj_id:
        for cat in categories:
            conn.execute(
                "INSERT OR IGNORE INTO project_categories (project_id, category) VALUES (?, ?)",
                (proj_id[0], str(cat)),
            )
    conn.commit()


# ── 项目 CRUD ──────────────────────────────────────────────────

def save_project(name: str, workspace: str, project_type: str = "",
                 current_stage: str = "", meta: dict[str, Any] | None = None,
                 stages: list[str] | None = None, categories: list[str] | None = None) -> dict[str, Any]:
    conn = _get_conn()
    now = datetime.now().isoformat()

    existing = conn.execute("SELECT id FROM projects WHERE name=?", (name,)).fetchone()
    meta_json = json.dumps(meta or {}, ensure_ascii=False)
    stages_json = json.dumps(stages or [], ensure_ascii=False)

    if existing:
        conn.execute(
            """UPDATE projects SET workspace=?, project_type=?, current_stage=?,
               meta_json=?, stages_json=?, updated_at=? WHERE name=?""",
            (workspace, project_type, current_stage, meta_json, stages_json, now, name),
        )
    else:
        conn.execute(
            """INSERT INTO projects (name, workspace, project_type, current_stage, meta_json, stages_json, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (name, workspace, project_type, current_stage, meta_json, stages_json, now, now),
        )

    proj_row = conn.execute("SELECT id FROM projects WHERE name=?", (name,)).fetchone()
    if proj_row and categories:
        conn.execute("DELETE FROM project_categories WHERE project_id=?", (proj_row[0],))
        for cat in categories:
            conn.execute(
                "INSERT OR IGNORE INTO project_categories (project_id, category) VALUES (?, ?)",
                (proj_row[0], cat),
            )

    conn.commit()
    return get_project(name) or {}


def get_project(name: str) -> dict[str, Any] | None:
    conn = _get_conn()
    row = conn.execute("SELECT * FROM projects WHERE name=?", (name,)).fetchone()
    if not row:
        return None
    d = dict(row)

    # 加载 categories
    cats = conn.execute("SELECT category FROM project_categories WHERE project_id=?", (d["id"],)).fetchall()
    meta = json.loads(d.get("meta_json", "{}"))
    meta["categories"] = [c["category"] for c in cats]

    return {
        "name": d["name"],
        "workspace": d["workspace"],
        "project_type": d["project_type"],
        "current_stage": d["current_stage"],
        "meta": meta,
        "stages": json.loads(d.get("stages_json", "[]")),
        "created_at": d["created_at"],
        "updated_at": d["updated_at"],
    }


def list_projects() -> list[dict[str, Any]]:
    conn = _get_conn()
    rows = conn.execute("SELECT name, workspace, project_type, current_stage, meta_json, created_at, updated_at FROM projects ORDER BY name").fetchall()
    result = []
    for r in rows:
        cats = conn.execute(
            "SELECT category FROM project_categories WHERE project_id=(SELECT id FROM projects WHERE name=?)", (r["name"],)
        ).fetchall()
        meta = json.loads(r["meta_json"] or "{}")
        meta["categories"] = [c["category"] for c in cats]
        result.append({
            "name": r["name"],
            "workspace": r["workspace"],
            "project_type": r["project_type"],
            "current_stage": r["current_stage"],
            "meta": meta,
            "created_at": r["created_at"],
            "updated_at": r["updated_at"],
        })
    return result


def update_project(name: str, **kwargs) -> dict[str, Any] | None:
    proj = get_project(name)
    if not proj:
        return None

    categories = kwargs.pop("categories", None) or kwargs.pop("meta_categories", None)
    meta = kwargs.pop("meta", proj.get("meta", {}))

    proj.update(kwargs)
    proj["meta"] = meta
    return save_project(
        name=proj["name"],
        workspace=proj.get("workspace", ""),
        project_type=proj.get("project_type", ""),
        current_stage=proj.get("current_stage", ""),
        meta=meta,
        stages=proj.get("stages", []),
        categories=categories or meta.get("categories", []),
    )


# ── 模板 CRUD ──────────────────────────────────────────────────

def save_template(project_name: str, template_name: str, definition: dict[str, Any],
                  report_type: str = "", source_file: str = "") -> int:
    conn = _get_conn()
    now = datetime.now().isoformat()
    existing = conn.execute(
        "SELECT id FROM templates WHERE project_name=? AND template_name=?",
        (project_name, template_name),
    ).fetchone()
    if existing:
        conn.execute(
            """UPDATE templates SET definition_json=?, report_type=?, source_file=?, created_at=?
               WHERE project_name=? AND template_name=?""",
            (json.dumps(definition, ensure_ascii=False), report_type, source_file, now,
             project_name, template_name),
        )
    else:
        conn.execute(
            """INSERT INTO templates (project_name, template_name, report_type, definition_json, source_file, created_at)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (project_name, template_name, report_type,
             json.dumps(definition, ensure_ascii=False), source_file, now),
        )
    conn.commit()
    return conn.execute("SELECT id FROM templates WHERE project_name=? AND template_name=?",
                        (project_name, template_name)).fetchone()["id"]


def get_template(project_name: str, template_name: str) -> dict[str, Any] | None:
    conn = _get_conn()
    row = conn.execute(
        "SELECT * FROM templates WHERE project_name=? AND template_name=?",
        (project_name, template_name),
    ).fetchone()
    if not row:
        return None
    d = dict(row)
    return {
        "id": d["id"],
        "project_name": d["project_name"],
        "template_name": d["template_name"],
        "report_type": d["report_type"],
        "definition": json.loads(d["definition_json"] or "{}"),
        "source_file": d["source_file"],
        "created_at": d["created_at"],
    }


def list_templates(project_name: str) -> list[dict[str, Any]]:
    conn = _get_conn()
    rows = conn.execute(
        "SELECT id, template_name, report_type, source_file, created_at FROM templates WHERE project_name=? ORDER BY created_at DESC",
        (project_name,),
    ).fetchall()
    return [{"id": r["id"], "template_name": r["template_name"], "report_type": r["report_type"],
             "source_file": r["source_file"], "created_at": r["created_at"]} for r in rows]


# ── 归档 CRUD ──────────────────────────────────────────────────

def save_archive(project_name: str, category: str, filename: str, path: str,
                 original_path: str = "", description: str = "", quality_notes: str = "",
                 analysis: dict[str, Any] | None = None) -> int:
    conn = _get_conn()
    now = datetime.now().isoformat()
    conn.execute(
        """INSERT INTO archives (project_name, category, filename, path, original_path, description, quality_notes, analysis_json, archived_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (project_name, category, filename, path, original_path, description, quality_notes,
         json.dumps(analysis or {}, ensure_ascii=False), now),
    )
    conn.commit()
    return conn.execute("SELECT last_insert_rowid()").fetchone()[0]


def get_archives(project_name: str = "", category: str = "", limit: int = 100) -> list[dict[str, Any]]:
    conn = _get_conn()
    conditions = []
    params: list[Any] = []
    if project_name:
        conditions.append("project_name = ?")
        params.append(project_name)
    if category:
        conditions.append("category = ?")
        params.append(category)
    where = (" WHERE " + " AND ".join(conditions)) if conditions else ""
    rows = conn.execute(
        f"SELECT * FROM archives{where} ORDER BY archived_at DESC LIMIT ?",
        params + [limit],
    ).fetchall()
    return [dict(r) for r in rows]


# ── 调用日志 ───────────────────────────────────────────────────

def record_call(tool_name: str, params: dict[str, Any] | None = None,
                result_summary: str = "", status: str = "success",
                duration_ms: int = 0, module: str = "", session_id: str = "") -> int:
    conn = _get_conn()
    safe_params = json.dumps(_sanitize_params(params or {}), ensure_ascii=False)
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        """INSERT INTO calls (tool_name, module, params_json, result_summary, status, duration_ms, called_at, session_id)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (tool_name, module, safe_params, (result_summary or "")[:500], status, duration_ms, now, session_id),
    )
    conn.commit()
    return conn.execute("SELECT last_insert_rowid()").fetchone()[0]


def query_calls(limit: int = 20, tool_name: str = "", module: str = "",
                status: str = "", since: str = "", session_id: str = "") -> list[dict[str, Any]]:
    conn = _get_conn()
    conditions: list[str] = []
    params: list[Any] = []
    if tool_name:
        conditions.append("tool_name = ?"); params.append(tool_name)
    if module:
        conditions.append("module = ?"); params.append(module)
    if status:
        conditions.append("status = ?"); params.append(status)
    if since:
        cutoff = _parse_since(since)
        if cutoff:
            conditions.append("called_at >= ?"); params.append(cutoff.isoformat(timespec="seconds"))
    if session_id:
        conditions.append("session_id = ?"); params.append(session_id)
    where = (" WHERE " + " AND ".join(conditions)) if conditions else ""
    rows = conn.execute(
        f"SELECT * FROM calls{where} ORDER BY called_at DESC LIMIT ?",
        params + [limit],
    ).fetchall()
    return [dict(r) for r in rows]


def get_call_summary(since: str = "24h", session_id: str = "") -> dict[str, Any]:
    conn = _get_conn()
    cutoff = _parse_since(since)
    cutoff_str = cutoff.isoformat(timespec="seconds") if cutoff else ""
    params: list[Any] = []
    sf = ""
    if session_id:
        sf = " AND session_id = ?"
        params = [session_id]
    tf = ""
    tp: list[Any] = []
    if cutoff_str:
        tf = " AND called_at >= ?"
        tp = [cutoff_str]

    all_p = params + tp
    total = conn.execute(f"SELECT COUNT(*) as c FROM calls WHERE 1=1{sf}{tf}", all_p).fetchone()["c"]
    errors_count = conn.execute(f"SELECT COUNT(*) as c FROM calls WHERE status='error'{sf}{tf}", all_p).fetchone()["c"]

    by_tool = {r["tool_name"]: r["cnt"] for r in conn.execute(
        f"SELECT tool_name, COUNT(*) as cnt FROM calls WHERE 1=1{sf}{tf} GROUP BY tool_name ORDER BY cnt DESC", all_p
    ).fetchall()}

    by_module = {r["module"]: r["cnt"] for r in conn.execute(
        f"SELECT module, COUNT(*) as cnt FROM calls WHERE 1=1{sf}{tf} AND module != '' GROUP BY module ORDER BY cnt DESC", all_p
    ).fetchall()}

    errors = [{"tool": r["tool_name"], "summary": r["result_summary"], "time": r["called_at"]}
              for r in conn.execute(
        f"SELECT tool_name, result_summary, called_at FROM calls WHERE status='error'{sf}{tf} ORDER BY called_at DESC LIMIT 5", all_p
    ).fetchall()]

    recent = [r["tool_name"] for r in conn.execute(
        f"SELECT DISTINCT tool_name FROM calls WHERE 1=1{sf}{tf} ORDER BY called_at DESC LIMIT 5", all_p
    ).fetchall()]

    return {
        "total_calls": total,
        "errors_count": errors_count,
        "success_rate": f"{(total - errors_count) / total * 100:.1f}%" if total > 0 else "N/A",
        "by_tool": by_tool,
        "by_module": by_module,
        "errors": errors,
        "recent_tools": recent,
        "period": f"过去 {since}" if since else "全部",
    }


# ── helpers ─────────────────────────────────────────────────────

def _sanitize_params(params: dict[str, Any]) -> dict[str, Any]:
    sensitive = {"api_key", "token", "password", "secret", "key", "auth"}
    safe = {}
    for k, v in params.items():
        if any(s in k.lower() for s in sensitive):
            s = str(v); safe[k] = s[:4] + "****" if len(s) > 4 else "****"
        elif isinstance(v, str) and len(v) > 200:
            safe[k] = v[:200] + "..."
        else:
            safe[k] = v
    return safe


def _parse_since(since: str) -> datetime | None:
    import re
    m = re.match(r"(\d+)([hd])", since.lower())
    if m:
        n = int(m.group(1)); u = m.group(2)
        return datetime.now() - timedelta(hours=n if u == "h" else n * 24)
    return None


def _classify_module(tool_name: str) -> str:
    mapping = {
        "wechat": "wechat_monitor", "monitor": "wechat_monitor",
        "project": "project_initializer", "template": "project_initializer",
        "scan_directory": "image_organizer", "organize_images": "image_organizer",
        "daily_log": "construction_log", "log_from_ledgers": "construction_log",
        "list_logs": "construction_log", "read_log": "construction_log",
        "experiment": "experiment_report",
        "get_config": "system", "update_config": "system", "check_wechat": "system",
        "install_wechat": "system", "setup": "system", "get_summary": "system",
        "get_activity": "system",
    }
    for key, mod in mapping.items():
        if key in tool_name: return mod
    return "other"
