"""目录扫描：递归查找图片文件，去重和增量处理。"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif", ".heic"}


def scan_directory(
    source_dir: str | Path,
    recursive: bool = True,
) -> list[dict[str, Any]]:
    """扫描目录中的所有图片文件。

    Returns:
        [{"path": "/abs/path/img.jpg", "name": "img.jpg", "size": 12345, "hash": "abc..."}, ...]
    """
    source = Path(source_dir).expanduser().resolve()
    if not source.exists():
        return []

    images = []

    if recursive:
        files = source.rglob("*")
    else:
        files = source.glob("*")

    for f in files:
        if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS:
            stat = f.stat()
            images.append({
                "path": str(f),
                "name": f.name,
                "size": stat.st_size,
                "mtime": stat.st_mtime,
                "relative_path": str(f.relative_to(source)),
            })

    images.sort(key=lambda x: x["name"])
    return images


def compute_hash(file_path: str | Path) -> str:
    """计算文件 MD5 哈希，用于去重。"""
    h = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def filter_new_images(
    images: list[dict[str, Any]],
    archive_base: Path,
) -> list[dict[str, Any]]:
    """过滤出尚未归档的图片（按文件名去重）。

    Args:
        images: 扫描结果
        archive_base: 归档根目录

    Returns:
        尚未归档的图片列表（附加 hash）
    """
    # 收集归档目录中已有的文件名
    existing_names: set[str] = set()
    if archive_base.exists():
        for f in archive_base.rglob("*"):
            if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS:
                existing_names.add(f.name)

    new_images = []
    for img in images:
        if img["name"] not in existing_names:
            img["hash"] = compute_hash(img["path"])
            new_images.append(img)

    return new_images
