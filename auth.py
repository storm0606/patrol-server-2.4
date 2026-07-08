#!/usr/bin/env python3
"""巡更系统鉴权工具模块 — 纯工具函数库（无数据库操作）

# 本模块只提供纯工具函数:
# - 密码处理：hash_password(), generate_salt(), verify_password()
# - HMAC 签名：compute_hmac(), verify_hmac()
# - MAC 地址处理：get_machine_macs(), format_mac(), mac_to_hash()
# - Session Token: generate_session_token(), verify_session_token(token, db), create_session(user_id, db)
# - 设备指纹：get_device_fingerprint(), get_cpu_id(), get_board_serial(), get_disk_serials(), get_machine_sid()
# - 辅助函数：hash_token()

数据库 CRUD 操作已移至 server.py。
Session 从内存存储改为数据库 session_tokens 表存储。
"""
import hashlib, json, os, uuid, hmac, secrets, bcrypt, subprocess, platform, time, shlex, re, sqlite3, sys

def get_resource_path(relative_path):
    """获取资源文件的正确路径 (支持 PyInstaller 打包)"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时目录（只读，用于读取资源文件）
        base_path = sys._MEIPASS
    else:
        # 普通 Python 运行
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def get_writable_path(relative_path):
    """获取可写路径 (支持 PyInstaller 打包)
    
    打包后：返回 EXE 同级目录（可写），供运行时数据库写入
    未打包：返回源码目录（可写）
    
    注意：打包后首次启动时，init_db() 会自动将 _internal/data/patrol.db
    复制到该可写路径，以便读取预置数据。
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

# ===== 密码处理 =====
def hash_password(password, salt=None):
    """使用 bcrypt 哈希密码（自动生成随机 salt）"""
    if salt is None:
        salt = generate_salt()
    # bcrypt 接受 salt 参数或自动生成
    if isinstance(salt, str):
        salt = salt.encode('utf-8')
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt(rounds=12)).decode('utf-8')


def generate_salt():
    """生成随机salt"""
    return secrets.token_hex(16)


def verify_password(password, password_hash):
    """验证密码（支持 bcrypt 和旧 SHA256 哈希向后兼容）

    Args:
        password: 明文密码
        password_hash: 存储的密码哈希

    Returns:
        bool: 密码是否匹配
    
    Note: LOW-011 - 密码验证逻辑复杂，保留向后兼容用于迁移。
          迁移计划：未来版本将移除 SHA256 兼容模式，强制使用 bcrypt。
    """
    try:
        # 尝试 bcrypt 验证（当前格式）- 检查所有常见 bcrypt 前缀
        if password_hash.startswith('$2a$') or password_hash.startswith('$2b$') or password_hash.startswith('$2y$'):
            try:
                return bcrypt.checkpw(password.encode('utf-8'), password_hash.encode('utf-8'))
            except bcrypt.AuthenticationError:
                # HIGH-007: bcrypt 验证失败（密码不匹配）
                return False
            except Exception as e:
                # HIGH-007: 其他 bcrypt 异常，记录日志
                print(f"[ERROR] bcrypt verification failed: {e}")
                return False
        # 向后兼容：旧 SHA256 格式（仅用于迁移）
        # HIGH-001: 硬编码 salt 已移除，使用系统配置 salt（如果可用）
        # 注意：以下 SHA256 验证仅用于历史数据迁移，新系统应强制使用 bcrypt
        old_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()
        if hmac.compare_digest(old_hash, password_hash):
            return True
    except Exception as e:
        print(f"[ERROR] verify_password exception: {e}")
        return False
    return False


# ===== HMAC 签名 =====
def compute_hmac(data, key=None):
    """计算 HMAC-SHA256 签名"""
    if isinstance(data, dict):
        data = json.dumps(data, sort_keys=True, ensure_ascii=False)
    key_bytes = key.encode() if key else b""
    return hmac.new(key_bytes, data.encode(), hashlib.sha256).hexdigest()


def verify_hmac(data, signature, key=None):
    """验证 HMAC 签名"""
    if isinstance(data, dict):
        data = json.dumps(data, sort_keys=True, ensure_ascii=False)
    expected = compute_hmac(data, key)
    return hmac.compare_digest(expected, signature)


# ===== MAC 地址处理 =====
def _get_interface_macs():
    """读取 /sys/class/net 获取本机所有网卡MAC"""
    macs = []
    net_dir = "/sys/class/net"
    if os.path.isdir(net_dir):
        for iface in os.listdir(net_dir):
            addr_file = os.path.join(net_dir, iface, "address")
            try:
                with open(addr_file) as f:
                    addr = f.read().strip().replace(":", "").replace("-", "").lower()
                if addr and addr != "000000000000" and len(addr) == 12:
                    macs.append(addr)
            except Exception:
                continue
    return macs


def _get_uuid_mac():
    """使用 uuid.getnode() 获取MAC"""
    try:
        node = uuid.getnode()
        if node and node != uuid.UUID(int=0).node:
            mac = f"{node:012x}"
            if mac != "000000000000":
                return mac
    except Exception:
        pass
    return None


def get_machine_macs():
    """获取本机MAC地址列表（去重）"""
    macs = []
    uuid_mac = _get_uuid_mac()
    if uuid_mac:
        macs.append(uuid_mac)
    macs.extend(_get_interface_macs())
    # 去重
    seen = set()
    unique = []
    for m in macs:
        if m not in seen:
            seen.add(m)
            unique.append(m)
    return unique


def format_mac(mac_str):
    """格式化MAC为 XX:XX:XX:XX:XX:XX"""
    mac = mac_str.replace(":", "").replace("-", "").lower()
    return ":".join(mac[i:i+2] for i in range(0, 12, 2))


def mac_to_hash(mac_str):
    """将 MAC 转为哈希（SHA256 前 12 字符）"""
    return hashlib.sha256(mac_str.replace(":", "").lower().encode()).hexdigest()[:12]


# ===== 设备指纹（跨平台唯一标识）=====
def _run_command(cmd, shell=False):
    """执行系统命令并返回输出"""
    try:
        # MEDIUM-001 修复：默认使用 shell=False，如果 cmd 是字符串则用 shlex.split() 解析
        if isinstance(cmd, str) and not shell:
            cmd = shlex.split(cmd)
        result = subprocess.run(
            cmd, shell=shell, capture_output=True, text=True, timeout=10
        )
        return result.stdout.strip() if result.returncode == 0 else None
    except Exception:
        return None


def get_cpu_id():
    """获取 CPU ID（跨平台）"""
    system = platform.system()
    
    if system == "Linux":
        # Linux: 读取/proc/cpuinfo
        try:
            with open("/proc/cpuinfo", "r") as f:
                for line in f:
                    if line.startswith("serial"):
                        cpu_serial = line.split(":")[1].strip()
                        if cpu_serial:
                            return cpu_serial
        except Exception:
            pass
        # 备用：使用 lscpu，明确匹配 Model name
        output = _run_command("lscpu")
        if output:
            for line in output.split("\n"):
                if "Model name" in line and ":" in line:
                    return line.split(":")[1].strip()
        # 再次备用：从 /proc/cpuinfo 读取 model name
        try:
            with open("/proc/cpuinfo", "r") as f:
                for line in f:
                    if line.startswith("model name"):
                        return line.split(":")[1].strip()
        except Exception:
            pass
    
    elif system == "Windows":
        # Windows: wmic cpu get ProcessorId
        output = _run_command("wmic cpu get ProcessorId")
        if output:
            lines = output.split("\n")
            for line in lines:
                line = line.strip()
                # 跳过标题行和空行，验证是否为有效的 16 进制字符串
                if line and "ProcessorId" not in line:
                    # ProcessorId 通常是 8-32 位的 16 进制字符
                    hex_chars = "0123456789ABCDEFabcdef"
                    if len(line) >= 8 and all(c in hex_chars for c in line):
                        return line
            # 宽松降级：如果严格的十六进制验证没找到，取第一个非标题、非空行
            # MEDIUM-002 修复：添加数据验证，只允许字母数字和连字符
            for line in lines:
                line = line.strip()
                if line and "ProcessorId" not in line:
                    # 验证：长度至少 1 字符，只允许字母数字和连字符
                    if len(line) >= 1 and re.match(r'^[a-zA-Z0-9\-]+$', line):
                        return line
                    else:
                        # 无效数据，记录警告并返回 unknown_cpu
                        print(f"[WARNING] Invalid Windows CPU ID format: '{line}' (contains invalid characters)")
                        return "unknown_cpu"
    
    # 降级：使用 CPU 核心数等特征
    import multiprocessing
    return f"cpu_{multiprocessing.cpu_count()}_cores"


def get_board_serial():
    """获取主板序列号（跨平台）"""
    system = platform.system()
    
    if system == "Linux":
        # Linux: 使用 dmidecode
        output = _run_command("dmidecode -s baseboard-serial-number")
        if output and output != "Default String" and output != "To be filled by O.E.M.":
            return output
        # 备用：读取/sys/class/dmi
        try:
            with open("/sys/class/dmi/id/board_serial", "r") as f:
                serial = f.read().strip()
                if serial and serial != "Default String":
                    return serial
        except Exception:
            pass
    
    elif system == "Windows":
        # Windows: wmic bios get serialnumber
        output = _run_command("wmic bios get serialnumber")
        if output:
            lines = output.split("\n")
            for line in lines:
                line = line.strip()
                if line and "SerialNumber" not in line and line != "Default String":
                    return line
    
    return "unknown_board"


def get_disk_serials():
    """获取所有硬盘序列号列表（跨平台，支持多硬盘采集）
    
    Returns:
        list: 硬盘序列号列表，每个元素是一个硬盘的序列号字符串
    """
    serials = []
    system = platform.system()
    
    if system == "Linux":
        # Linux: 尝试读取/dev/disk/by-id
        try:
            disk_by_id = "/dev/disk/by-id"
            if os.path.isdir(disk_by_id):
                found_disks = set()  # 使用 set 去重
                for name in os.listdir(disk_by_id):
                    if "ata" in name.lower() or "nvme" in name.lower():
                        # 提取序列号部分
                        parts = name.split("-")
                        for part in parts:
                            if len(part) >= 10 and any(c.isdigit() for c in part):
                                if part not in found_disks:
                                    found_disks.add(part)
                                    serials.append(part)
                                break  # 每个设备只取一个序列号
        except Exception:
            pass
        # 备用：使用 hdparm
        if not serials:
            for disk in ["/dev/sda", "/dev/nvme0n1", "/dev/vda"]:
                if os.path.exists(disk):
                    output = _run_command(f"hdparm -I {disk} 2>/dev/null | grep 'Serial Number'")
                    if output:
                        parts = output.split(":")
                        if len(parts) > 1:
                            serial = parts[1].strip()
                            if serial and serial not in serials:
                                serials.append(serial)
    
    elif system == "Windows":
        # Windows: wmic diskdrive get serialnumber
        output = _run_command("wmic diskdrive get serialnumber")
        if output:
            lines = output.split("\n")
            for line in lines:
                line = line.strip()
                if line and "SerialNumber" not in line:
                    serials.append(line)
    
    return serials if serials else ["unknown_disk"]


def get_machine_sid():
    """获取机器 SID/唯一标识（跨平台）"""
    system = platform.system()
    
    if system == "Linux":
        # Linux: 读取/etc/machine-id
        try:
            with open("/etc/machine-id", "r") as f:
                machine_id = f.read().strip()
                if machine_id:
                    return machine_id
        except Exception:
            pass
        # 备用：/var/lib/dbus/machine-id
        try:
            with open("/var/lib/dbus/machine-id", "r") as f:
                machine_id = f.read().strip()
                if machine_id:
                    return machine_id
        except Exception:
            pass
    
    elif system == "Windows":
        # Windows: wmic useraccount get sid (获取第一个用户的 SID)
        output = _run_command("wmic useraccount get sid")
        if output:
            lines = output.split("\n")
            for line in lines:
                line = line.strip()
                if line and "SID" not in line and line.startswith("S-"):
                    return line
    
    return "unknown_machine"


def get_device_fingerprint():
    """生成设备唯一指纹（跨平台）
    
    综合以下信息生成唯一指纹：
    - CPU ID
    - 主板序列号
    - 硬盘序列号
    - 机器 SID
    - MAC 地址（多网卡全部采集）
    
    Returns:
        str: 32 字符的 SHA256 哈希值（设备指纹）
    """
    # 收集所有硬件标识
    identifiers = []
    
    # CPU ID
    cpu_id = get_cpu_id()
    if cpu_id:
        identifiers.append(f"cpu:{cpu_id}")
    
    # 主板序列号
    board_serial = get_board_serial()
    if board_serial:
        identifiers.append(f"board:{board_serial}")
    
    # 硬盘序列号（支持多硬盘采集）
    disk_serials = get_disk_serials()
    for disk_serial in sorted(disk_serials):
        if disk_serial:
            identifiers.append(f"disk:{disk_serial}")
    
    # 机器 SID
    machine_sid = get_machine_sid()
    if machine_sid:
        identifiers.append(f"machine:{machine_sid}")
    
    # MAC 地址（多网卡）
    macs = get_machine_macs()
    for mac in sorted(macs):
        identifiers.append(f"mac:{mac}")
    
    # 排序后拼接
    identifiers.sort()
    fingerprint_input = "|".join(identifiers)
    
    # SHA256 哈希取前 32 字符
    fingerprint = hashlib.sha256(fingerprint_input.encode()).hexdigest()[:32]
    
    return fingerprint


# ===== Session Token（数据库存储）=====
def generate_session_token():
    """生成session token"""
    return secrets.token_urlsafe(32)


def hash_token(token):
    """哈希session token用于存储"""
    return hashlib.sha256(token.encode()).hexdigest()


def verify_session_token(token, db):
    """验证 session token 是否有效（数据库存储）

    Args:
        token: session token 字符串
        db: 数据库连接

    Returns:
        bool: token 是否有效
    """
    if not token:
        return False
    token_hash = hash_token(token)
    now = time.time()
    row = db.execute(
        "SELECT expires_at FROM session_tokens WHERE token_hash = ?",
        (token_hash,)
    ).fetchone()
    if not row:
        return False
    expires = row["expires_at"]
    # MEDIUM-008 修复：添加宽限期逻辑 - token 过期后 5 分钟内不立即删除，允许续期
    # 宽限期用于处理时钟不同步或短暂网络问题导致的误判
    grace_period = 300  # 5 分钟宽限期（秒）
    now_with_grace = time.time() + grace_period
    if now_with_grace > expires:
        # token 已过期且超过宽限期，删除
        db.execute("DELETE FROM session_tokens WHERE token_hash = ?", (token_hash,))
        db.commit()
        return False
    # 在宽限期内，可以继续使用（但建议客户端重新登录）
    return True


def create_session(user_id, db):
    """创建 session token（数据库存储）

    Args:
        user_id: 用户标识（设备指纹或用户名）
        db: 数据库连接

    Returns:
        str: session token
    """
    token = generate_session_token()
    token_hash = hash_token(token)
    now = time.time()
    expires = now + 24 * 3600  # 24 小时有效
    db.execute(
        "INSERT OR REPLACE INTO session_tokens (token_hash, user_id, created_at, expires_at) VALUES (?, ?, ?, ?)",
        (token_hash, user_id, now, expires)
    )
    db.commit()
    return token


# ===== 鉴权检查函数 =====
def is_auth_enabled():
    """检查鉴权是否启用（查询 auth_users 表是否有活跃记录）
    
    Returns:
        bool: 鉴权是否启用
    """
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT 1 FROM auth_users WHERE active = 1 LIMIT 1"
        ).fetchone()
        return row is not None
    except Exception:
        return False
    finally:
        if conn:
            conn.close()


def check_admin_password(password):
    """验证管理员密码（查询 auth_users 表验证密码）
    
    Args:
        password: 明文密码
    
    Returns:
        bool: 密码是否正确
    """
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT password_hash FROM auth_users WHERE role = 'admin' AND active = 1 LIMIT 1"
        ).fetchone()
        if not row:
            return False
        stored_hash = row["password_hash"]
        try:
            return bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8'))
        except bcrypt.AuthenticationError:
            # HIGH-007: bcrypt 验证失败（密码不匹配）
            return False
        except Exception as e:
            # HIGH-007: 其他 bcrypt 异常，记录日志
            print(f"[ERROR] bcrypt verification in check_admin_password: {e}")
            return False
    except Exception as e:
        print(f"[ERROR] check_admin_password: {e}")
        return False
    finally:
        if conn:
            conn.close()


def check_password(password):
    """验证密码（check_admin_password 的别名）
    
    Args:
        password: 明文密码
    
    Returns:
        bool: 密码是否正确
    """
    return check_admin_password(password)


def check_admin_op_password(password):
    """验证管理密码（用于点位操作等管理功能）
    验证存储在 admin_password_hash 字段中的管理密码
    
    Args:
        password: 明文密码
    
    Returns:
        bool: 管理密码是否正确
    """
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT admin_password_hash FROM auth_users WHERE role = 'admin' AND active = 1 LIMIT 1"
        ).fetchone()
        if not row or not row["admin_password_hash"]:
            return False
        stored_hash = row["admin_password_hash"]
        try:
            return bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8'))
        except bcrypt.AuthenticationError:
            # HIGH-007: bcrypt 验证失败（密码不匹配）
            return False
        except Exception as e:
            # HIGH-007: 其他 bcrypt 异常，记录日志
            print(f"[ERROR] bcrypt verification in check_admin_op_password: {e}")
            return False
    except Exception as e:
        print(f"[ERROR] check_admin_op_password: {e}")
        return False
    finally:
        if conn:
            conn.close()


def get_system_config(key, default=None):
    """从 system_config 表获取配置值
    
    Args:
        key: 配置键
        default: 默认值，如果配置不存在则返回此值
    
    Returns:
        str: 配置值，如果不存在返回 default
    """
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT value FROM system_config WHERE key = ?",
            (key,)
        ).fetchone()
        if row:
            return row["value"]
        return default
    except Exception as e:
        print(f"[ERROR] get_system_config: {e}")
        return default
    finally:
        if conn:
            conn.close()


def init_security_db():
    """初始化安全配置数据库（首次启动时创建默认配置）"""
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        # 检查是否已有签名密钥
        row = conn.execute(
            "SELECT value FROM system_config WHERE key = 'signature_key'"
        ).fetchone()
        if not row:
            # 生成默认签名密钥
            default_key = secrets.token_hex(32)
            conn.execute(
                "INSERT INTO system_config (key, value, description) VALUES (?, ?, ?)",
                ("signature_key", default_key, "HMAC 签名密钥")
            )
        # HIGH-004 修复：初始化重置密钥的 bcrypt 哈希（默认密钥：Storm）
        row = conn.execute(
            "SELECT value FROM system_config WHERE key = 'password_reset_key_hash'"
        ).fetchone()
        if not row:
            default_reset_key = "Storm"
            reset_key_hash = bcrypt.hashpw(default_reset_key.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            conn.execute(
                "INSERT INTO system_config (key, value, description) VALUES (?, ?, ?)",
                ("password_reset_key_hash", reset_key_hash, "密码重置密钥的 bcrypt 哈希（默认：Storm）")
            )
        # 初始化 2FA TOTP 密钥
        row = conn.execute(
            "SELECT value FROM system_config WHERE key = '2fa_secret_key'"
        ).fetchone()
        if not row:
            conn.execute(
                "INSERT INTO system_config (key, value, description) VALUES (?, ?, ?)",
                ("2fa_secret_key", "B23K65GK3K4LZZ7KKM2MLMOUP6BL2ZRG", "2FA TOTP 密钥 (Base32 编码，32 字符)")
            )
        conn.commit()
    except Exception as e:
        print(f"[ERROR] init_security_db: {e}")
    finally:
        if conn:
            conn.close()


def get_effective_config():
    """获取有效配置（从 system_config 表读取）
    
    Returns:
        dict: 配置字典
    """
    config = {"authorized_macs": []}
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT key, value FROM system_config"
        ).fetchall()
        for row in rows:
            key = row["key"]
            value = row["value"]
            if key == "authorized_macs":
                try:
                    config[key] = json.loads(value)
                except:
                    config[key] = []
            else:
                config[key] = value
    except Exception as e:
        print(f"[ERROR] get_effective_config: {e}")
    finally:
        if conn:
            conn.close()
    return config
