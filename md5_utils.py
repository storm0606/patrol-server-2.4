"""MD5 哈希工具模块"""

import hashlib


def calculate_md5(text: str) -> str:
    """
    计算字符串的 MD5 哈希值
    
    Args:
        text: 输入字符串
        
    Returns:
        MD5 哈希值（32 位十六进制字符串）
    """
    md5_hash = hashlib.md5(text.encode('utf-8'))
    return md5_hash.hexdigest()


if __name__ == "__main__":
    # 测试示例
    test_string = "Hello, World!"
    result = calculate_md5(test_string)
    print(f"字符串 '{test_string}' 的 MD5 哈希值：{result}")
