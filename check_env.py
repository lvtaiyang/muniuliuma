#!/usr/bin/env python3
"""木牛流马 环境检测 — 纯标准库，零依赖，克隆后第一步运行。

用法:
    python check_env.py          # 检测环境
    python check_env.py --json   # JSON 输出（供其他工具解析）
"""

import importlib
import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


REQUIRED = ["mcp", "openai", "yaml", "httpx"]
OPTIONAL = {
    "python-docx": "docx",        # pip 包名 → import 名
    "openpyxl": "openpyxl",
    "pdfplumber": "pdfplumber",
}
WIN32_PKG = "pywin32"


def check_py_version() -> dict:
    v = sys.version_info
    ok = (v.major, v.minor) >= (3, 10)
    return {
        "current": f"{v.major}.{v.minor}.{v.micro}",
        "required": ">=3.10",
        "ok": ok,
    }


def check_package(import_name: str) -> dict:
    try:
        mod = importlib.import_module(import_name)
        version = getattr(mod, "__version__", "?")
        return {"installed": True, "version": version}
    except ImportError:
        return {"installed": False, "version": None}


def check_nodejs() -> dict:
    path = shutil.which("node") or shutil.which("nodejs")
    if not path:
        return {"installed": False, "version": None, "path": None}
    try:
        out = subprocess.check_output([path, "--version"], text=True, timeout=5)
        return {"installed": True, "version": out.strip(), "path": path}
    except Exception:
        return {"installed": True, "version": "?", "path": path}


def check_win32() -> dict:
    """检测 Windows COM 环境。"""
    is_win = platform.system() == "Windows"
    if not is_win:
        return {"available": False, "note": "非 Windows，走 openpyxl/python-docx 回退模式"}

    pkg = check_package("win32com.client")
    if not pkg["installed"]:
        return {"available": False, "note": "pywin32 未安装，走 openpyxl/python-docx 回退"}

    # 探测 Office/WPS
    try:
        import win32com.client
        for prog_id, name in [("Ket.Application", "WPS Excel"), ("Excel.Application", "Microsoft Excel")]:
            try:
                app = win32com.client.Dispatch(prog_id)
                app.Quit()
                return {"available": True, "excel": name, "method": "win32com"}
            except Exception:
                continue
        for prog_id, name in [("Kwps.Application", "WPS Word"), ("Word.Application", "Microsoft Word")]:
            try:
                app = win32com.client.Dispatch(prog_id)
                app.Quit()
                return {"available": True, "word": name, "method": "win32com"}
            except Exception:
                continue
        return {"available": True, "method": "win32com", "note": "COM OK，但未探测到 Office/WPS 应用"}
    except Exception:
        return {"available": True, "method": "win32com", "note": "COM OK，探测应用失败"}


def check_config() -> dict:
    """检查 config.yaml 是否已配置 API key。"""
    config_path = Path(__file__).parent / "config.yaml"
    if not config_path.exists():
        return {"exists": False, "text_configured": False, "vision_configured": False}

    try:
        with open(config_path) as f:
            content = f.read()

        text_key = "api_key:" in content and "sk-" in content
        vision_key = "api_key:" in content

        import yaml
        data = yaml.safe_load(content)
        llm = data.get("llm", {}) if data else {}
        text_ok = bool(llm.get("text", {}).get("api_key", ""))
        vision_ok = bool(llm.get("vision", {}).get("api_key", ""))

        return {
            "exists": True,
            "text_configured": text_ok,
            "vision_configured": vision_ok,
            "text_model": llm.get("text", {}).get("model", ""),
            "vision_model": llm.get("vision", {}).get("model", ""),
        }
    except Exception:
        return {"exists": True, "text_configured": False, "vision_configured": False, "error": "解析失败"}


def get_mcp_config() -> str:
    self_path = Path(__file__).parent / "mcp_server.py"
    return json.dumps({
        "mcpServers": {
            "muniuliuma": {
                "command": sys.executable,
                "args": [str(self_path.resolve())],
            }
        }
    }, ensure_ascii=False, indent=2)


def run():
    is_json = "--json" in sys.argv

    py = check_py_version()
    node = check_nodejs()
    win32 = check_win32()
    config = check_config()

    required = {name: check_package(name) for name in REQUIRED}
    optional = {pip_name: check_package(import_name) for pip_name, import_name in OPTIONAL.items()}

    all_required_ok = all(v["installed"] for v in required.values())
    tools_available = (
        sum(v["installed"] for v in required.values()) +
        sum(v["installed"] for v in optional.values())
    )

    missing_required = [name for name, v in required.items() if not v["installed"]]
    missing_optional = [name for name, v in optional.items() if not v["installed"]]

    if is_json:
        print(json.dumps({
            "python": py,
            "nodejs": node,
            "win32com": win32,
            "config": config,
            "packages": {
                "required": {k: v["installed"] for k, v in required.items()},
                "optional": {k: v["installed"] for k, v in optional.items()},
            },
            "ready": all_required_ok and py["ok"],
            "missing_required": missing_required,
            "mcp_config": json.loads(get_mcp_config()) if all_required_ok else None,
        }, ensure_ascii=False, indent=2))
        return

    # ── 人类可读输出 ───────────────────────────────────────────

    print("木牛流马 环境检测")
    print("=" * 50)

    # Python
    icon = "✓" if py["ok"] else "✗"
    print(f"\n  {icon} Python {py['current']} (需要 {py['required']})")

    # 核心包
    print(f"\n  核心依赖:")
    all_core_ok = True
    for name, info in required.items():
        icon = "✓" if info["installed"] else "✗"
        ver = f" ({info['version']})" if info["installed"] else ""
        print(f"    {icon} {name}{ver}")
        if not info["installed"]:
            all_core_ok = False
    if missing_required:
        print(f"    缺: {' '.join(missing_required)} → pip install {' '.join(missing_required)}")

    # 可选包
    print(f"\n  可选依赖 (文档处理):")
    for name, info in optional.items():
        icon = "✓" if info["installed"] else "-"
        ver = f" ({info['version']})" if info["installed"] else " (未安装)"
        print(f"    {icon} {name}{ver}")

    # Windows COM
    print(f"\n  Word/Excel 操作:")
    if win32["available"]:
        print(f"    ✓ win32com — {win32.get('excel', win32.get('word', win32.get('method','')))}{win32.get('note','')}")
    else:
        print(f"    - {win32.get('note', '不可用')}")

    # Node.js
    print(f"\n  Node.js (微信监控):")
    if node["installed"]:
        print(f"    ✓ {node['version']} ({node['path']})")
    else:
        print(f"    - 未安装，微信监控模块不可用")

    # API Key
    print(f"\n  API Key:")
    if config.get("text_configured"):
        print(f"    ✓ text ({config.get('text_model', '?')}) 已配置")
    else:
        print(f"    ✗ text 未配置 — 模板分析/报告生成不可用")
    if config.get("vision_configured"):
        print(f"    ✓ vision ({config.get('vision_model', '?')}) 已配置")
    else:
        print(f"    ✗ vision 未配置 — 照片分析/影像归档不可用")

    # 总结
    print(f"\n{'=' * 50}")

    if all_required_ok and py["ok"]:
        print("✓ 环境就绪，可以启动 MCP Server")
        print(f"\n  MCP 配置 (复制到 ~/.claude/mcp.json 或 .cursor/mcp.json):")
        print(f"  {get_mcp_config()}")
    else:
        print("✗ 缺少核心依赖，先安装：")
        if not py["ok"]:
            print("  需要 Python >= 3.10")
        if missing_required:
            print(f"  pip install {' '.join(missing_required)}")

    if not config["text_configured"] and not config["vision_configured"]:
        print(f"\n  然后编辑 config.yaml 填入 API key")

    print()


if __name__ == "__main__":
    run()
