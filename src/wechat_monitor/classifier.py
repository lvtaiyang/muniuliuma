"""自动分类：将 LLM 分析结果匹配到用户配置的项目和类别。"""

from __future__ import annotations

import re
from typing import Any


def classify(
    analysis: dict[str, Any],
    projects: list[dict[str, Any]],
) -> dict[str, str]:
    """将分析结果匹配到项目/类别。

    Args:
        analysis: LLM 分析结果，含 project_hint, category, keywords 等。
        projects: 用户配置的项目列表，每项含 name, categories。

    Returns:
        {"project": "项目名", "category": "类别名", "confidence": "high|medium|low"}
    """
    proj_hint = (analysis.get("project_hint") or "").strip()
    category = (analysis.get("category") or "").strip()
    keywords = analysis.get("keywords", [])

    best_project = None
    best_score = 0

    for proj in projects:
        score = 0

        # 项目名出现在 hint 中
        if proj_hint and _fuzzy_match(proj["name"], proj_hint):
            score += 10

        # 类别在项目的 categories 中
        proj_cats = [c.strip() for c in proj.get("categories", [])]
        if category in proj_cats:
            score += 5
        else:
            # 模糊匹配类别
            for cat in proj_cats:
                if _fuzzy_match(cat, category):
                    score += 3
                    category = cat  # 修正为精确名称
                    break

        # 关键词匹配项目名或类别
        for kw in keywords:
            if _fuzzy_match(proj["name"], kw):
                score += 2
            for cat in proj_cats:
                if _fuzzy_match(cat, kw):
                    score += 1

        if score > best_score:
            best_score = score
            best_project = proj["name"]

    if best_project is None:
        # 无匹配，归入第一个项目或 "未分类"
        if projects:
            best_project = projects[0]["name"]
            best_score = 0
        else:
            best_project = "未分类"

    if not category:
        category = "其他"

    confidence = "low"
    if best_score >= 10:
        confidence = "high"
    elif best_score >= 5:
        confidence = "medium"

    return {
        "project": best_project,
        "category": category,
        "confidence": confidence,
    }


def _fuzzy_match(a: str, b: str) -> bool:
    """模糊匹配：忽略大小写/空格/标点后，检查 a 是否包含 b 或 b 包含 a。"""
    def normalize(s: str) -> str:
        return re.sub(r"\s+", "", s.lower())

    na, nb = normalize(a), normalize(b)
    if not na or not nb:
        return False
    return na in nb or nb in na
