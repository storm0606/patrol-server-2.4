#!/usr/bin/env python3
"""
设备指纹验证中间件模块

提供设备指纹验证功能，用于巡更管理系统的设备授权管理。
从数据库读取授权设备指纹列表，验证传入的指纹是否在授权列表中。

参考 auth.py 的代码风格实现。
"""

import sqlite3
import os
import sys
from typing import Dict, List, Optional


def get_resource_path(relative_path: str) -> str:
    """获取资源文件的正确路径 (支持 PyInstaller 打包)"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时目录（只读，用于读取资源文件）
        base_path = sys._MEIPASS
    else:
        # 普通 Python 运行
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def get_writable_path(relative_path: str) -> str:
    """获取可写路径 (支持 PyInstaller 打包)
    
    打包后：返回 EXE 同级目录（可写），供运行时数据库写入
    未打包：返回源码目录（可写）
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后：使用 EXE 所在目录（可写）
        base_path = os.path.dirname(sys.executable)
    else:
        # 普通 Python 运行：使用源码目录
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


# 数据库路径：使用可写路径（打包后持久化）
DB_PATH = get_writable_path(os.path.join("data", "patrol.db"))


def get_authorized_fingerprints(db_path: Optional[str] = None) -> List[str]:
    """从数据库中获取所有授权设备的指纹列表
    
    尝试从 authorized_devices 表或 device_fingerprints 表读取授权设备信息。
    
    Args:
        db_path: 数据库路径，如果为 None 则使用默认路径
        
    Returns:
        list: 授权设备指纹的列表，每个元素为32字符的SHA256哈希值
    """
    if db_path is None:
        db_path = DB_PATH
    
    fingerprints: List[str] = []
    conn = None
    
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        
        # 尝试从 authorized_devices 表获取
        cursor = conn.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='authorized_devices'
        """)
        
        if cursor.fetchone():
            cursor = conn.execute(
                "SELECT fingerprint FROM authorized_devices WHERE status = 'active'"
            )
            fingerprints = [row[0] for row in cursor.fetchall()]
        
        # 如果 authorized_devices 表不存在或为空，尝试 device_fingerprints 表
        if not fingerprints:
            cursor = conn.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name='device_fingerprints'
            """)
            
            if cursor.fetchone():
                cursor = conn.execute(
                    "SELECT fingerprint FROM device_fingerprints WHERE is_authorized = 1"
                )
                fingerprints = [row[0] for row in cursor.fetchall()]
                
    except sqlite3.Error as e:
        print(f"[ERROR] get_authorized_fingerprints: {e}")
    except FileNotFoundError as e:
        print(f"[ERROR] get_authorized_fingerprints: 数据库文件不存在 - {e}")
    finally:
        if conn:
            conn.close()
    
    return fingerprints


def verify_device_fingerprint(fingerprint: str) -> Dict[str, bool | str]:
    """验证设备指纹是否在授权列表中
    
    从数据库读取授权设备指纹列表，验证传入的指纹是否在授权列表中。
    支持从 authorized_devices 表或 device_fingerprints 表读取授权设备信息。
    
    设备指纹格式：
    - 32字符的SHA256哈希值
    - 由 CPU ID、主板序列号、硬盘序列号、机器SID、MAC地址组合生成
    
    Args:
        fingerprint: 设备指纹字符串，应为32字符的SHA256哈希值
        
    Returns:
        dict: 包含验证结果的字典
            - 'valid': bool，表示指纹是否有效
            - 'message': str，验证结果的描述信息
    
    Raises:
        TypeError: 如果 fingerprint 参数不是字符串类型
    
    Example:
        >>> result = verify_device_fingerprint("a1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6")
        >>> print(result)
        {'valid': True, 'message': '设备指纹验证成功'}
    """
    # 参数类型检查
    if not isinstance(fingerprint, str):
        raise TypeError(f"fingerprint 参数必须是字符串类型，当前类型为: {type(fingerprint).__name__}")
    
    # 参数值验证 - 空字符串检查
    if not fingerprint:
        return {'valid': False, 'message': '设备指纹不能为空'}
    
    # 验证指纹格式：32字符的十六进制字符串
    if len(fingerprint) != 32:
        return {'valid': False, 'message': '设备指纹格式无效，应为32字符的十六进制字符串'}
    
    if not all(c in '0123456789abcdef' for c in fingerprint.lower()):
        return {'valid': False, 'message': '设备指纹格式无效，应为32字符的十六进制字符串'}
    
    # 将指纹转为小写以便比较
    fingerprint = fingerprint.lower()
    
    try:
        # 获取授权设备指纹列表
        authorized_fingerprints = get_authorized_fingerprints(DB_PATH)
        
        # 如果没有授权设备（数据库为空或未配置），返回验证通过（允许首次注册）
        if not authorized_fingerprints:
            return {'valid': True, 'message': '未配置授权设备，验证通过'}
        
        # 验证指纹是否在授权列表中
        if fingerprint in authorized_fingerprints:
            return {'valid': True, 'message': '设备指纹验证成功'}
        else:
            return {'valid': False, 'message': '设备未授权，请先进行设备注册'}
            
    except sqlite3.Error as e:
        print(f"[ERROR] verify_device_fingerprint 数据库操作失败: {e}")
        return {'valid': False, 'message': f'数据库操作失败: {str(e)}'}
    except Exception as e:
        print(f"[ERROR] verify_device_fingerprint 验证过程出错: {e}")
        return {'valid': False, 'message': f'设备指纹验证失败: {str(e)}'}


def verify_device_components(mac_address: str = None, disk_serial: str = None, 
                              cpu_id: str = None) -> Dict[str, bool | str]:
    """验证设备硬件组件是否在授权列表中
    
    支持通过 MAC 地址、硬盘序列号、CPU ID 组合验证设备。
    至少需要提供其中一个组件信息。
    
    Args:
        mac_address: MAC 地址字符串，格式为 XX:XX:XX:XX:XX:XX
        disk_serial: 硬盘序列号字符串
        cpu_id: CPU ID 字符串
        
    Returns:
        dict: 包含验证结果的字典
            - 'valid': bool，表示设备是否有效
            - 'message': str，验证结果的描述信息
    
    Example:
        >>> result = verify_device_components(mac_address="00:1B:44:11:3A:B7")
        >>> print(result)
        {'valid': True, 'message': '设备验证成功'}
    """
    # 至少需要一个组件信息
    if not mac_address and not disk_serial and not cpu_id:
        return {'valid': False, 'message': '至少需要提供 MAC 地址、硬盘序列号或 CPU ID 之一'}
    
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        
        # 检查 authorized_devices 表是否存在
        cursor = conn.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='authorized_devices'
        """)
        
        if not cursor.fetchone():
            return {'valid': True, 'message': '未配置授权设备表，验证通过'}
        
        # 构建查询条件
        conditions = []
        params = []
        
        if mac_address:
            conditions.append("mac_address = ?")
            params.append(mac_address.lower().replace(':', '').replace('-', ''))
        
        if disk_serial:
            conditions.append("disk_serial = ?")
            params.append(disk_serial)
        
        if cpu_id:
            conditions.append("cpu_id = ?")
            params.append(cpu_id)
        
        if not conditions:
            return {'valid': False, 'message': '无效的验证参数'}
        
        # 查询授权设备
        query = f"""
            SELECT id, device_name, status FROM authorized_devices 
            WHERE {' OR '.join(conditions)} AND status = 'active'
            LIMIT 1
        """
        
        cursor = conn.execute(query, params)
        row = cursor.fetchone()
        
        if row:
            return {'valid': True, 'message': f'设备验证成功: {row["device_name"]}'}
        else:
            return {'valid': False, 'message': '设备未授权，请先进行设备注册'}
            
    except sqlite3.Error as e:
        print(f"[ERROR] verify_device_components 数据库操作失败: {e}")
        return {'valid': False, 'message': f'数据库操作失败: {str(e)}'}
    except Exception as e:
        print(f"[ERROR] verify_device_components 验证过程出错: {e}")
        return {'valid': False, 'message': f'设备验证失败: {str(e)}'}
    finally:
        if conn:
            conn.close()


# 测试代码
if __name__ == '__main__':
    print("设备指纹验证中间件测试")
    print("=" * 50)
    
    # 测试 verify_device_fingerprint 函数
    test_fingerprints = [
        "a1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6",  # 合法格式
        "A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6",  # 大写合法格式
        "invalid_fingerprint",                  # 非法格式（长度不对）
        "g1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6",   # 非法字符
        "",                                     # 空字符串
    ]
    
    print("\n--- 测试 verify_device_fingerprint ---")
    for fp in test_fingerprints:
        try:
            result = verify_device_fingerprint(fp)
            print(f"指纹: {fp}")
            print(f"验证结果: {result}")
            print()
        except TypeError as e:
            print(f"指纹: {fp}")
            print(f"验证错误: {e}")
            print()
    
    # 测试 verify_device_components 函数
    print("\n--- 测试 verify_device_components ---")
    test_components = [
        {"mac_address": "00:1B:44:11:3A:B7", "disk_serial": None, "cpu_id": None},
        {"mac_address": None, "disk_serial": "WD-ABC123", "cpu_id": None},
        {"mac_address": None, "disk_serial": None, "cpu_id": "BFEBFBFF000906E9"},
        {"mac_address": None, "disk_serial": None, "cpu_id": None},  # 无参数
    ]
    
    for components in test_components:
        result = verify_device_components(**components)
        print(f"输入: {components}")
        print(f"验证结果: {result}")
        print()