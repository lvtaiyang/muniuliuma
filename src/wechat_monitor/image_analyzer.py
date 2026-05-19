"""多模态模型图片分析。

使用 OpenAI 兼容 API 调用国内多模态模型（通义千问、DeepSeek 等），
从工程隐蔽验收照片中提取结构化信息。
"""

from __future__ import annotations

import base64
import json
import re
from pathlib import Path
from typing import Any

from openai import OpenAI

SYSTEM_PROMPT = """你是一个工程隐蔽验收影像资料分析专家。分析施工照片，提取关键信息。

你必须严格返回 JSON 格式，不要包含任何其他文字：

{
  "project_hint": "从图片中推断的项目名称或施工地点，无法判断则为空字符串",
  "category": "验收类别，如：地基验槽、钢筋隐蔽工程、防水工程、混凝土浇筑、砌体工程、抹灰工程、桩基工程、土方回填、模板工程、钢结构、管线预埋、屋面工程、其他",
  "description": "简洁描述图片中的施工内容（20字以内）",
  "date_hint": "从图片水印/文字中识别到的日期，YYYY-MM-DD 格式，无法识别则为空",
  "quality_notes": "对施工质量的初步判断：合格/存在缺陷/无法判断",
  "keywords": ["关键词1", "关键词2"]
}"""


def analyze_image(
    image_path: str | Path,
    base_url: str,
    api_key: str,
    model: str,
) -> dict[str, Any]:
    """对单张工程照片进行多模态分析，返回结构化信息。"""
    image_path = Path(image_path)

    mime_type = _mime_type(image_path)
    data_url = _encode_image(image_path, mime_type)

    client = OpenAI(base_url=base_url, api_key=api_key, timeout=120)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": data_url},
                    },
                    {
                        "type": "text",
                        "text": "请分析这张工程隐蔽验收照片，返回 JSON。",
                    },
                ],
            },
        ],
        max_tokens=500,
        temperature=0.1,
    )

    content = response.choices[0].message.content or ""
    return _parse_response(content)


def _mime_type(path: Path) -> str:
    suffix = path.suffix.lower()
    mime_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp"}
    return mime_map.get(suffix, "image/jpeg")


def _encode_image(path: Path, mime_type: str) -> str:
    data = path.read_bytes()
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{mime_type};base64,{b64}"


def _parse_response(content: str) -> dict[str, Any]:
    """从 LLM 响应中提取 JSON。"""
    # 尝试直接解析
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        pass

    # 尝试提取 ```json ... ``` 代码块
    match = re.search(r"```(?:json)?\s*\n?([\s\S]*?)\n?```", content)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass

    # 尝试提取 { ... } 对象
    match = re.search(r"\{[\s\S]*\}", content)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass

    # 返回原始文本作为兜底
    return {"raw_response": content, "category": "无法识别", "description": "分析失败",
            "project_hint": "", "date_hint": "", "quality_notes": "无法判断", "keywords": []}
