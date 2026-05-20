"""LLM JSON 输出解析 — jiter 优先 + 修复 + strict 提示。"""

from __future__ import annotations

import re
from typing import Any

# jiter 随 openai 包安装，容错性比 json 好
try:
    import jiter
    _HAS_JITER = True
except ImportError:
    _HAS_JITER = False


def parse_llm_json(content: str) -> dict[str, Any]:
    """从 LLM 输出中解析 JSON。jiter partial_mode → json → repair。"""
    if not content or not content.strip():
        return {}

    text = content.strip()

    # 1. 去掉 ```json ... ``` 包裹
    text = re.sub(r'```(?:json)?\s*\n?([\s\S]*?)\n?```', r'\1', text, flags=re.DOTALL)

    # 2. jiter 尝试（带 partial_mode 容错）
    if _HAS_JITER:
        try:
            result = jiter.from_json(text.encode("utf-8"), partial_mode=True)
            if isinstance(result, dict):
                return result
            if isinstance(result, list):
                return {"_list": result}
        except Exception:
            pass

    # 3. 标准 json 尝试
    try:
        import json
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 4. JSON 修复再试
    repaired = _repair_json(text)
    try:
        import json
        return json.loads(repaired)
    except json.JSONDecodeError:
        pass

    # 5. 兜底：尝试提取最外层 { ... }
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        try:
            import json
            return json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            pass

    # 6. 完全放弃，返回原始文本
    return {"_raw": content, "_error": "JSON 解析失败"}


def _repair_json(text: str) -> str:
    """修复常见 JSON 格式错误。"""
    # 截取 { ... } 范围
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start:end + 1]

    # 去掉尾随逗号
    text = re.sub(r',\s*}', '}', text)
    text = re.sub(r',\s*]', ']', text)

    # 修复未转义的引号：在字符串值中查找未转义的双引号
    # 匹配 "key": "value" 模式，其中 value 可能含未转义的 "
    # 策略：用正则匹配 "(\w+)":\s*"([^"]*)"，逐行修复
    def _fix_unquoted_val(m):
        key = m.group(1)
        val = m.group(2)
        # 如果 value 里还有引号，用转义引号替换
        val = val.replace('"', '\\"')
        return f'"{key}": "{val}"'

    # 修复：在一行之内的 "key": "value" 模式
    text = re.sub(r'"(?P<key>[^"]+)":\s*"(?P<val>[^"]*)"', _fix_unquoted_val, text)

    return text


# Prompt 后缀：让 LLM 输出更纯净的 JSON
STRICT_JSON_SUFFIX = """

你必须输出纯粹的 JSON，不要加任何代码块标记（\`\`\`json），不要加解释文字。
JSON 中所有字符串值内的双引号必须用 \\" 转义。
确保最后一个字段后面没有逗号。"""
