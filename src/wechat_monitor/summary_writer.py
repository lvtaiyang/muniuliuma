"""MD 汇总：生成和更新归档影像资料汇总文件。"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any


def generate_summary(
    archive_base: Path,
    entries: list[dict[str, Any]],
) -> str:
    """从当前归档条目生成完整 MD 汇总。"""
    today = date.today().isoformat()

    lines = [
        f"# 隐蔽验收影像资料汇总",
        f"",
        f"> 最后更新: {today}",
        f"",
    ]

    if not entries:
        lines.append("暂无归档记录。")
        return "\n".join(lines)

    # 按项目分组
    by_project: dict[str, list[dict]] = {}
    for e in entries:
        proj = e.get("project", "未分类")
        by_project.setdefault(proj, []).append(e)

    # 目录
    lines.append("## 项目索引")
    lines.append("")
    for proj_name in sorted(by_project):
        count = len(by_project[proj_name])
        anchor = proj_name.replace(" ", "-")
        lines.append(f"- [{proj_name}](#{anchor}) ({count} 份)")
    lines.append("")

    # 每个项目的详细列表
    for proj_name in sorted(by_project):
        items = by_project[proj_name]
        lines.append(f"## {proj_name}")
        lines.append("")

        # 按类别分组
        by_cat: dict[str, list[dict]] = {}
        for e in items:
            by_cat.setdefault(e.get("category", "其他"), []).append(e)

        for cat_name in sorted(by_cat):
            cat_items = by_cat[cat_name]
            lines.append(f"### {cat_name}")
            lines.append("")
            lines.append("| 文件名 | 描述 | 质量判断 | 归档日期 |")
            lines.append("|--------|------|----------|----------|")
            for e in cat_items:
                fname = e.get("filename", "-")
                desc = e.get("description", "-")
                quality = e.get("quality_notes", "-")
                arch_date = e.get("date", today)
                lines.append(f"| {fname} | {desc} | {quality} | {arch_date} |")
            lines.append("")

    # 统计
    lines.append("## 统计")
    lines.append("")
    lines.append(f"- 归档总数: {len(entries)} 份")
    lines.append(f"- 涉及项目: {len(by_project)} 个")
    lines.append(f"- 最后更新: {today}")
    lines.append("")

    return "\n".join(lines)


def write_summary(archive_base: Path, entries: list[dict[str, Any]]) -> Path:
    """写入 MD 汇总文件到归档根目录。"""
    content = generate_summary(archive_base, entries)
    path = archive_base / "影像资料汇总.md"
    path.write_text(content, encoding="utf-8")
    return path


def collect_entries(archive_base: Path) -> list[dict[str, Any]]:
    """扫描归档目录，收集已有条目元数据。"""
    entries = []
    if not archive_base.exists():
        return entries

    for proj_dir in sorted(archive_base.iterdir()):
        if not proj_dir.is_dir():
            continue
        for cat_dir in sorted(proj_dir.iterdir()):
            if not cat_dir.is_dir():
                continue
            for img_file in sorted(cat_dir.iterdir()):
                if img_file.suffix.lower() in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"):
                    entries.append({
                        "project": proj_dir.name,
                        "category": cat_dir.name,
                        "filename": img_file.name,
                        "path": str(img_file),
                        "description": "",
                        "quality_notes": "",
                        "date": date.fromtimestamp(img_file.stat().st_mtime).isoformat(),
                    })
    return entries
