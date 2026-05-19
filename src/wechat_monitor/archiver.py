"""图片归档：复制到项目目录，按统一规则命名。"""

from __future__ import annotations

import re
import shutil
from datetime import date
from pathlib import Path
from typing import Any


def sanitize_filename(s: str) -> str:
    """清理字符串中的非法文件名字符。"""
    s = re.sub(r"[\\/:*?\"<>|]", "", s)
    s = re.sub(r"\s+", "_", s)
    return s[:50]  # 限制长度


def build_filename(
    naming_template: str,
    analysis: dict[str, Any],
    classification: dict[str, str],
    index: int,
) -> str:
    """根据命名模板生成目标文件名。

    Args:
        naming_template: 如 "{date}_{category}_{description}"
        analysis: LLM 分析结果
        classification: 分类结果
        index: 图片序号（同批次去重用）
    """
    today = date.today().isoformat()
    date_hint = analysis.get("date_hint", "") or today
    category = sanitize_filename(classification.get("category", "其他"))
    description = sanitize_filename(analysis.get("description", "未命名"))
    project = sanitize_filename(classification.get("project", ""))

    filename = naming_template.format(
        date=date_hint,
        date_today=today,
        category=category,
        description=description,
        project=project,
        index=index,
    )
    return filename


def archive_image(
    src_path: Path,
    archive_base: Path,
    naming_template: str,
    analysis: dict[str, Any],
    classification: dict[str, str],
    index: int,
) -> dict[str, Any]:
    """将单张图片归档到目标目录。

    Returns:
        {"archived_path": "...", "project": "...", "category": "...", "filename": "..."}
    """
    project_name = sanitize_filename(classification["project"])
    category_name = sanitize_filename(classification["category"])

    # 构建目标目录: base_path/项目名/类别/
    target_dir = archive_base / project_name / category_name
    target_dir.mkdir(parents=True, exist_ok=True)

    # 生成文件名
    filename = build_filename(naming_template, analysis, classification, index)

    # 保留原始扩展名
    ext = src_path.suffix.lower()
    if ext in (".dat", ".decrypted"):
        # 从解密后文件推断
        ext = ".jpg"
    target_name = f"{filename}{ext}"
    target_path = target_dir / target_name

    # 处理重名
    counter = 1
    while target_path.exists():
        target_path = target_dir / f"{filename}_{counter}{ext}"
        counter += 1

    # 复制
    shutil.copy2(src_path, target_path)

    return {
        "archived_path": str(target_path),
        "project": project_name,
        "category": category_name,
        "filename": target_path.name,
        "original_path": str(src_path),
    }
