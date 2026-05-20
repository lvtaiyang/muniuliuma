"""核心管线：编排图片拉取 → 解密 → 分析 → 分类 → 归档 → 汇总的完整流程。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from .. import config as cfg
from . import archiver
from . import classifier
from . import image_analyzer
from . import image_decryptor
from . import summary_writer
from . import wechat_reader


class MonitorResult:
    def __init__(self):
        self.success: int = 0
        self.skipped: int = 0
        self.errors: list[dict[str, Any]] = []
        self.entries: list[dict[str, Any]] = []

    @property
    def total_processed(self) -> int:
        return self.success + self.skipped + len(self.errors)


def run_monitor() -> MonitorResult:
    """执行一次完整的监控-归档管线。"""
    result = MonitorResult()
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

    wm_conf = conf.get("wechat_monitor", {})
    start_time = wm_conf.get("last_run")

    for group_name in wm_conf.get("groups", []):
        _process_group(
            group_name, start_time, conf, llm_cfg,
            archive_base, projects, naming_template, result,
        )

    # 更新汇总
    all_entries = summary_writer.collect_entries(archive_base)
    # 合并本次新归档的条目信息
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

    cfg.update_wechat_last_run()
    return result


def _process_group(
    group_name: str,
    start_time: str | None,
    conf: dict[str, Any],
    llm_cfg: dict[str, str],
    archive_base: Path,
    projects: list[dict[str, Any]],
    naming_template: str,
    result: MonitorResult,
) -> None:
    """处理单个群聊的图片。"""
    try:
        # 查找匹配的群聊
        sessions = wechat_reader.list_groups(limit=200)
        matched = None
        for s in sessions:
            chat = s.get("chat", "")
            if group_name.lower() in chat.lower():
                matched = s
                break

        if not matched:
            result.errors.append({
                "stage": "find_group",
                "group": group_name,
                "error": f"未找到匹配的群聊",
            })
            return

        actual_name = matched["chat"]

        # 拉取图片消息
        time_arg = None
        if start_time:
            time_arg = wechat_reader.format_time_for_cli(start_time)

        messages = wechat_reader.fetch_images(
            actual_name,
            start_time=time_arg,
            limit=200,
        )

        if not messages:
            return

        for i, msg in enumerate(messages):
            try:
                img_path_str = wechat_reader.parse_image_path(msg)
                if not img_path_str:
                    result.skipped += 1
                    continue

                src_path = Path(img_path_str)
                if not src_path.exists():
                    result.errors.append({
                        "stage": "find_file",
                        "group": actual_name,
                        "error": f"文件不存在: {img_path_str}",
                    })
                    continue

                # 解密
                decrypted_path = image_decryptor.decrypt_file(src_path)

                # 分析
                analysis = image_analyzer.analyze_image(
                    decrypted_path,
                    base_url=llm_cfg["base_url"],
                    api_key=llm_cfg["api_key"],
                    model=llm_cfg["model"],
                )

                # 分类
                clf = classifier.classify(analysis, projects)

                # 归档
                arch = archiver.archive_image(
                    decrypted_path, archive_base, naming_template,
                    analysis, clf, i + 1,
                )

                result.entries.append({
                    **arch,
                    "description": analysis.get("description", ""),
                    "quality_notes": analysis.get("quality_notes", ""),
                    "sender": _extract_sender(msg),
                    "source_group": actual_name,
                })
                result.success += 1

            except Exception as exc:
                result.errors.append({
                    "stage": "process_image",
                    "group": actual_name,
                    "error": str(exc),
                })

    except Exception as exc:
        result.errors.append({
            "stage": "fetch_group",
            "group": group_name,
            "error": str(exc),
        })


def _extract_sender(msg: str | dict) -> str:
    """从消息中提取发送者。"""
    if isinstance(msg, dict):
        text = msg.get("raw", "")
    else:
        text = str(msg)
    if "] " in text:
        parts = text.split("] ", 1)
        if len(parts) >= 2:
            sender_part = parts[1].split(":")[0]
            return sender_part.strip()
    return "未知"
