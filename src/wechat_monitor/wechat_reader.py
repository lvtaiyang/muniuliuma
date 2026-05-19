"""wechat-cli 封装 — 自动检测/安装，无需用户手动配置。

优先级：
1. 检查 PATH 中已有的 wechat-cli
2. 使用项目内置的 vendor/wechat-cli，自动 npm install
3. 没有 Node.js 时给出清晰的安装指引
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from pathlib import Path
from typing import Any

_VENDOR_DIR = Path(__file__).resolve().parent.parent.parent / "vendor" / "wechat-cli"


def _which(cmd: str) -> str | None:
    return shutil.which(cmd)


def _has_node() -> bool:
    return _which("node") is not None


def _has_npm() -> bool:
    return _which("npm") is not None


def _ensure_wechat_cli() -> list[str]:
    """确保 wechat-cli 可用，返回命令行前缀。

    Returns:
        ["wechat-cli"] — 系统已安装
        ["node", "vendor/wechat-cli/bin/wechat-cli.js"] — 使用内置版本

    Raises:
        RuntimeError: Node.js 未安装，并给出安装指引。
    """
    # 1. 系统已安装
    if _which("wechat-cli"):
        return ["wechat-cli"]

    # 2. 使用内置 vendor 版本
    js_path = _VENDOR_DIR / "bin" / "wechat-cli.js"
    if js_path.exists():
        if not _has_node():
            raise RuntimeError(
                "wechat-cli 需要 Node.js 运行环境。\n"
                "请安装 Node.js (>= 14): https://nodejs.org/\n"
                "安装后重新运行即可自动完成后续配置。"
            )
        # 确保 npm 依赖已安装（下载平台二进制）
        node_modules = _VENDOR_DIR / "node_modules"
        if not node_modules.exists():
            _npm_install()
        return ["node", str(js_path)]

    raise RuntimeError(
        "wechat-cli 未安装。请先安装 Node.js: https://nodejs.org/\n"
        "然后执行: npm install -g @canghe_ai/wechat-cli"
    )


def _npm_install() -> None:
    """在 vendor 目录安装 npm 依赖（自动下载对应平台二进制）。"""
    proc = subprocess.run(
        ["npm", "install", "--no-audit", "--no-fund"],
        cwd=str(_VENDOR_DIR),
        capture_output=True,
        text=True,
        timeout=120,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"wechat-cli 自动安装失败:\n{proc.stderr}\n"
            f"请手动执行: cd vendor/wechat-cli && npm install"
        )


def _run_wechat(args: list[str]) -> list[dict[str, Any]]:
    """执行 wechat-cli 并解析 JSON 输出。"""
    cmd_prefix = _ensure_wechat_cli()
    cmd = [*cmd_prefix, *args, "--format", "json"]

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    except FileNotFoundError:
        raise RuntimeError(
            "无法找到 Node.js 运行环境，请安装 Node.js >= 14: https://nodejs.org/"
        )

    stdout = proc.stdout.strip()
    stderr = proc.stderr.strip()

    if proc.returncode != 0 and not stdout:
        raise RuntimeError(f"wechat-cli 执行失败:\n{stderr}")

    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        if stdout:
            return [{"raw": line} for line in stdout.split("\n") if line.strip()]
        return []


# ── 公开 API ────────────────────────────────────────────────────


def check_setup() -> dict[str, Any]:
    """检测 wechat-cli 环境状态，供智能体判断下一步操作。"""
    return {
        "wechat_cli_in_path": _which("wechat-cli") is not None,
        "node_available": _has_node(),
        "npm_available": _has_npm(),
        "vendor_bundled": (_VENDOR_DIR / "bin" / "wechat-cli.js").exists(),
        "vendor_deps_installed": (_VENDOR_DIR / "node_modules").exists(),
    }


def install() -> dict[str, Any]:
    """尝试自动安装 wechat-cli。"""
    result = {"success": False, "method": "", "error": None}

    if not _has_node():
        result["error"] = "Node.js 未安装，请先访问 https://nodejs.org/ 安装 (>= 14)"
        return result

    # 优先安装到 vendor 目录
    js_path = _VENDOR_DIR / "bin" / "wechat-cli.js"
    if js_path.exists():
        result["method"] = "vendor_npm_install"
        try:
            _npm_install()
            result["success"] = (_VENDOR_DIR / "node_modules").exists()
        except Exception as e:
            result["error"] = str(e)
        return result

    # 尝试全局安装
    if _has_npm():
        result["method"] = "global_npm_install"
        try:
            subprocess.run(
                ["npm", "install", "-g", "@canghe_ai/wechat-cli"],
                capture_output=True, text=True, timeout=120,
            )
            result["success"] = _which("wechat-cli") is not None
        except Exception as e:
            result["error"] = str(e)
        return result

    result["error"] = "无法自动安装，请手动执行: npm install -g @canghe_ai/wechat-cli"
    return result


def list_sessions(limit: int = 50) -> list[dict[str, Any]]:
    return _run_wechat(["sessions", "--limit", str(limit)])


def list_groups(limit: int = 50) -> list[dict[str, Any]]:
    sessions = list_sessions(limit)
    return [s for s in sessions if s.get("is_group")]


def fetch_images(
    chat_name: str,
    start_time: str | None = None,
    end_time: str | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    args = ["history", chat_name, "--type", "image", "--media", "--limit", str(limit)]
    if start_time:
        args.extend(["--start-time", start_time])
    if end_time:
        args.extend(["--end-time", end_time])
    result = _run_wechat(args)
    if isinstance(result, dict):
        return result.get("messages", [])
    return result


def parse_image_path(raw_message: str | dict) -> str | None:
    text = raw_message if isinstance(raw_message, str) else raw_message.get("raw", "")
    if "[图片]" in text:
        parts = text.split("[图片]")
        if len(parts) >= 2:
            path = parts[-1].strip()
            if path:
                return path
    return None


def format_time_for_cli(dt: Any) -> str:
    from datetime import datetime
    if isinstance(dt, str):
        dt = datetime.fromisoformat(dt)
    return dt.strftime("%Y-%m-%d %H:%M:%S")
