"""WeChat .dat 图片解密。

WeChat PC 将图片加密存储为 .dat 文件，采用单字节或多字节 XOR 加密。
通过对比已知图片文件头自动推算密钥。
"""

from __future__ import annotations

from pathlib import Path

# 常见图片格式的魔数 (magic bytes) 用于逆向推算密钥
MAGIC_BYTES = {
    "jpg": bytes([0xFF, 0xD8, 0xFF]),
    "png": bytes([0x89, 0x50, 0x4E, 0x47]),
    "gif": bytes([0x47, 0x49, 0x46, 0x38]),
    "bmp": bytes([0x42, 0x4D]),
    "webp": b"RIFF",
}


def _detect_xor_key(data: bytes) -> int | None:
    """通过匹配图片文件头，自动检测单字节 XOR 密钥。"""
    if len(data) < 8:
        return None
    for magic in MAGIC_BYTES.values():
        key = data[0] ^ magic[0]
        # 验证 key 对其他 magic bytes 也成立
        if all(data[i] ^ key == magic[i] for i in range(min(len(magic), 4))):
            return key
    return None


def _detect_multi_byte_key(data: bytes) -> tuple[bytes, str] | None:
    """尝试检测多字节循环 XOR 密钥（部分 WeChat 版本使用）。"""
    for fmt, magic in MAGIC_BYTES.items():
        if len(data) < len(magic):
            continue
        key = bytes(data[i] ^ magic[i] for i in range(len(magic)))
        if len(key) >= 2 and key[0] == key[-1]:
            continue  # 退化为单字节，交给 _detect_xor_key
        # 用这个 key 验证
        key_len = len(key)
        for i in range(len(magic)):
            if data[i] ^ key[i % key_len] != magic[i]:
                break
        else:
            return key, fmt
    return None


def decrypt_file(src: str | Path) -> Path:
    """解密 .dat 文件，返回解密后图片路径 (.jpg/.png 等)。

    自动检测加密方式并解密，解密后的文件与原文件同目录，
    文件名为原名去掉 .dat 后缀或替换为正确扩展名。
    """
    src = Path(src)
    if not src.exists():
        raise FileNotFoundError(f"图片文件不存在: {src}")

    data = src.read_bytes()

    # 先检查是否未加密（极少情况）
    for fmt, magic in MAGIC_BYTES.items():
        if data[: len(magic)] == magic:
            return src  # 已经是明文图片

    # 尝试单字节 XOR
    key = _detect_xor_key(data)
    if key is not None:
        decrypted = bytes(b ^ key for b in data)
        ext = _detect_format(decrypted) or "jpg"
        out_path = src.with_suffix(f".decrypted.{ext}")
        out_path.write_bytes(decrypted)
        return out_path

    # 尝试多字节 XOR
    multi = _detect_multi_byte_key(data)
    if multi is not None:
        key_bytes, fmt = multi
        key_len = len(key_bytes)
        decrypted = bytes(data[i] ^ key_bytes[i % key_len] for i in range(len(data)))
        out_path = src.with_suffix(f".decrypted.{fmt}")
        out_path.write_bytes(decrypted)
        return out_path

    # 暴力尝试常见单字节密钥
    for key in range(256):
        test = bytes(b ^ key for b in data[:8])
        fmt = _detect_format(test)
        if fmt:
            decrypted = bytes(b ^ key for b in data)
            out_path = src.with_suffix(f".decrypted.{fmt}")
            out_path.write_bytes(decrypted)
            return out_path

    raise ValueError(f"无法解密文件 (未知加密方式): {src}")


def _detect_format(data: bytes) -> str | None:
    """根据文件头判断图片格式。"""
    if data[:3] == b"\xFF\xD8\xFF":
        return "jpg"
    if data[:4] == b"\x89PNG":
        return "png"
    if data[:4] == b"GIF8":
        return "gif"
    if data[:2] == b"BM":
        return "bmp"
    if data[:4] == b"RIFF":
        return "webp"
    return None
