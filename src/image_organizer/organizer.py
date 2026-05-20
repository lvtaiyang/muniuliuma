"""现有资料整理管线：扫描 → 分析 → 分类 → 归档 → 汇总。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from .. import config as cfg
from ..wechat_monitor import image_analyzer, classifier, archiver, summary_writer
from . import scanner


class OrganizeResult:
    def __init__(self):
        self.total_found: int = 0
        self.new_images: int = 0
        self.success: int = 0
        self.skipped: int = 0
        self.errors: list[dict[str, Any]] = []
        self.entries: list[dict[str, Any]] = []


def organize_images(
    source_dir: str,
    recursive: bool = True,
    limit: int = 50,
) -> OrganizeResult:
    """整理指定目录中的现有影像资料。

    Args:
        source_dir: 图片源目录
        recursive: 是否递归扫描子目录
        limit: 最多处理张数（防止一次 API 调用过多）
    """
    result = OrganizeResult()
    conf = cfg.load()

    errors = cfg.validate(conf)
    if errors:
        for e in errors:
            result.errors.append({"stage": "validation", "error": e})
        return result

    llm_cfg = cfg.get_llm_config("vision")
    archive_base = cfg.resolve_archive_path(conf)
    projects = conf["projects"]
    naming_template = conf["archive"].get("naming", "{date}_{category}_{description}")

    # 1. 扫描目录
    all_images = scanner.scan_directory(source_dir, recursive=recursive)
    result.total_found = len(all_images)

    if not all_images:
        return result

    # 2. 去重
    new_images = scanner.filter_new_images(all_images, archive_base)
    result.new_images = len(new_images)

    if not new_images:
        return result

    # 3. 逐张处理（限制数量）
    to_process = new_images[:limit]
    for i, img in enumerate(to_process):
        try:
            src_path = Path(img["path"])

            # 分析
            analysis = image_analyzer.analyze_image(
                src_path,
                base_url=llm_cfg["base_url"],
                api_key=llm_cfg["api_key"],
                model=llm_cfg["model"],
            )

            # 分类
            clf = classifier.classify(analysis, projects)

            # 归档
            arch = archiver.archive_image(
                src_path, archive_base, naming_template,
                analysis, clf, i + 1,
            )

            result.entries.append({
                **arch,
                "description": analysis.get("description", ""),
                "quality_notes": analysis.get("quality_notes", ""),
                "source_dir": source_dir,
            })
            result.success += 1

        except Exception as exc:
            result.errors.append({
                "stage": "process_image",
                "file": img["name"],
                "error": str(exc),
            })

    # 4. 更新汇总
    all_entries = summary_writer.collect_entries(archive_base)
    for e in result.entries:
        for existing in all_entries:
            if existing.get("path") == e.get("archived_path"):
                existing.update(e)
                break
        else:
            all_entries.append({
                "project": e.get("project", ""),
                "category": e.get("category", ""),
                "filename": e.get("filename", ""),
                "path": e.get("archived_path", ""),
                "description": e.get("description", ""),
                "quality_notes": e.get("quality_notes", ""),
                "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })

    summary_writer.write_summary(archive_base, all_entries)

    return result
