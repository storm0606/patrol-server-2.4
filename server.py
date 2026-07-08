#!/usr/bin/env python3
"""巡更记录管理系统 2.4 - Patrol Record Management System
轮次定义:第二轮 03:00-05:00 第三轮05:00-07:00,第四轮07:00-09:00,第五轮09:00-11:00,第六轮11:00-13:00,第七轮13:00-15:00,第八轮15:00-17:00,第九轮17:00-19:00,第十轮19:00-21:00,第十一轮21:00-23:00,第十二轮23:00-24:00
每行数据=一次打点
"""
import json, sqlite3, os, sys, time, random, datetime, threading, hashlib, secrets
import re

# 启动日志: 在 import 之前就写入, 捕获任何导入错误
_STARTUP_LOGS = []
_STARTUP_LOGS.append(f"[{datetime.datetime.now()}] 巡更系统启动 - 预导入阶段")
_STARTUP_LOGS.append(f"Python: {sys.version}")
_STARTUP_LOGS.append(f"Platform: {sys.platform}")
_STARTUP_LOGS.append(f"Executable: {sys.executable}")

import pypinyin
import xlwt
from xlwt.BIFFRecords import (
    PrintHeadersRecord, PrintGridLinesRecord, GridSetRecord,
    HorizontalPageBreaksRecord, VerticalPageBreaksRecord,
    HCenterRecord, VCenterRecord, LeftMarginRecord, RightMarginRecord,
    TopMarginRecord, BottomMarginRecord, SetupPageRecord
)

# Monkey patch: 跳过HEADER/FOOTER记录,防止WPS转PDF自动加页眉页脚
def _patched_print_settings(self):
    result = b''
    result += PrintHeadersRecord(self._Worksheet__print_headers).get()
    result += PrintGridLinesRecord(self._Worksheet__print_grid).get()
    result += GridSetRecord(self._Worksheet__grid_set).get()
    result += HorizontalPageBreaksRecord(self._Worksheet__horz_page_breaks).get()
    result += VerticalPageBreaksRecord(self._Worksheet__vert_page_breaks).get()
    result += HCenterRecord(self._Worksheet__print_centered_horz).get()
    result += VCenterRecord(self._Worksheet__print_centered_vert).get()
    result += LeftMarginRecord(self._Worksheet__left_margin).get()
    result += RightMarginRecord(self._Worksheet__right_margin).get()
    result += TopMarginRecord(self._Worksheet__top_margin).get()
    result += BottomMarginRecord(self._Worksheet__bottom_margin).get()
    setup_page_options = (self._Worksheet__print_in_rows & 0x01) << 0
    setup_page_options |= (self._Worksheet__portrait & 0x01) << 1
    setup_page_options |= (0x00 & 0x01) << 2
    setup_page_options |= (self._Worksheet__print_not_colour & 0x01) << 3
    setup_page_options |= (self._Worksheet__print_draft & 0x01) << 4
    setup_page_options |= (self._Worksheet__print_notes & 0x01) << 5
    setup_page_options |= (0x00 & 0x01) << 6
    setup_page_options |= (0x01 & 0x01) << 7
    setup_page_options |= (self._Worksheet__print_notes_at_end & 0x01) << 9
    setup_page_options |= (self._Worksheet__print_omit_errors & 0x03) << 10
    result += SetupPageRecord(self._Worksheet__paper_size_code,
                            self._Worksheet__print_scaling,
                            self._Worksheet__start_page_number,
                            self._Worksheet__fit_width_to_pages,
                            self._Worksheet__fit_height_to_pages,
                            setup_page_options,
                            self._Worksheet__print_hres,
                            self._Worksheet__print_vres,
                            self._Worksheet__header_margin,
                            self._Worksheet__footer_margin,
                            self._Worksheet__copies_num).get()
    return result
xlwt.Worksheet._Worksheet__print_settings_rec = _patched_print_settings
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
from urllib.parse import urlparse, parse_qs

# xls2pdf-converter: XLS → PDF 本地转换（替代旧 XTP 服务）
from xls2pdf.converter import xls_to_pdf as xls2pdf_convert

# Threading HTTPServer for concurrent request handling
class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    """多线程 HTTP 服务器"""
    daemon_threads = True

# 动态获取代码修改日期(方案A:自动获取文件修改时间，PyInstaller 打包后无源码文件，使用默认日期)
try:
    SERVER_MTIME = datetime.datetime.fromtimestamp(os.path.getmtime(__file__)).strftime('%Y-%m-%d')
except (FileNotFoundError, OSError, TypeError):
    SERVER_MTIME = datetime.datetime.now().strftime('%Y-%m-%d')

# PyInstaller 路径支持
def get_resource_path(relative_path):
    """获取资源文件的正确路径(支持 PyInstaller 打包)"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时目录
        base_path = sys._MEIPASS
    else:
        # 普通 Python 运行
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

# 从 auth 模块导入可写路径函数
from auth import get_writable_path as _get_writable_path
DB_PATH = _get_writable_path(os.path.join("data", "patrol.db"))

# 启动日志: 写入可写路径（打包后可写，避免 _internal 只读问题）
LOG_PATH = _get_writable_path("startup.log")
try:
    with open(LOG_PATH, "w", encoding="utf-8") as f:
        for line in _STARTUP_LOGS:
            f.write(line + "\n")
        f.write(f"_MEIPASS: {getattr(sys, '_MEIPASS', 'N/A')}\n")
        f.write(f"DB_PATH: {DB_PATH}\n")
        f.write(f"DB exists: {os.path.exists(DB_PATH)}\n")
        f.flush()
except Exception as e:
    print(f"[ERROR] 日志写入失败：{e}", file=sys.stderr)  # LOW-010: 异常日志输出

# 清理预导入临时变量
# MEDIUM-003 修复：移除脆弱的变量删除逻辑，直接使用 del 删除确定存在的变量
del _STARTUP_LOGS

# 导出临时文件存储(token -> (bytes, filename, mime))
_export_store = {}
_export_store_lock = threading.Lock()

DEFAULT_POINTS = [
    # LOW-004: 硬编码默认点位 - 这些是系统初始化时的默认巡更点位
    # 当数据库中没有点位时，使用这些默认值
    "B 端办公区东大门", "消防通道内", "D 口通道", "D 口卷帘门",
    "C 口楼梯", "C 口直梯", "C 口卷帘门", "票亭", "B 口安检", "B 口",
    "B 口卷帘门", "A 端办公区西大门", "A 端办公区东大门", "A 口卷帘门",
    "上行头", "上行尾", "下行头", "下行尾", "站台卫生间", "B 端办公区西大门",
    "站厅直梯", "站厅付费区", "A 端非付费区"
]
# 轮次映射:round_number -> hour (第一轮00:00-03:00,第二轮 03:00-05:00,第三轮05:00-07:00,第四轮07:00-09:00,第五轮09:00-11:00,第六轮11:00-13:00,第七轮13:00-15:00,第八轮15:00-17:00,第九轮17:00-19:00,第十轮19:00-21:00,第十一轮21:00-23:00,第十二轮23:00-24:00)
# 轮次时间映射:round_number -> (start_hour, end_hour)
ROUND_TIME_RANGES = {
    1: (0, 3),    # 第 1 轮:00:00-03:00
    2: (3, 5),    # 第 2 轮:03:00-05:00
    3: (5, 7),    # 第 3 轮:05:00-07:00
    4: (7, 9),    # 第 4 轮:07:00-09:00
    5: (9, 11),   # 第 5 轮:09:00-11:00
    6: (11, 13),  # 第 6 轮:11:00-13:00
    7: (13, 15),  # 第 7 轮:13:00-15:00
    8: (15, 17),  # 第 8 轮:15:00-17:00
    9: (17, 19),  # 第 9 轮:17:00-19:00
    10: (19, 21), # 第 10 轮:19:00-21:00
    11: (21, 23), # 第 11 轮:21:00-23:00
    12: (23, 24), # 第 12 轮:23:00-24:00
}

def time_to_round(time_str):
    """根据时间判断轮次 (按新时间范围)
    轮次定义：第 1 轮 00:00-03:00, 第 2 轮 03:00-05:00, ..., 第 12 轮 23:00-24:00
    """
    try:
        # MEDIUM-004 修复：添加时间解析异常处理
        parts = time_str.split(" ")
        if len(parts) < 2:
            return 1
        time_part = parts[1]
        time_parts = time_part.split(":")
        if len(time_parts) < 1:
            return 1
        h = int(time_parts[0])
        # 根据小时判断轮次 (h==0 属于第 1 轮:00:00-03:00)
        for rnd, (start, end) in ROUND_TIME_RANGES.items():
            if start <= h < end:
                return rnd
        # 边界情况 (h==24 或其他) 归入第 12 轮
        return 12
    except Exception:
        return 1


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def parse_tag_full(tag_full):
    """
    解析完整标签字符串
    输入:"A 端 1-C 口 2" 或 "A 端 1--C 口 2"
    输出:{"label_l1": "A 端", "label_l1_order": 1, "label_l2": "C 口", "label_l2_order": 2}
    """
    if not tag_full:
        return None
    # 支持中文--和英文 -
    tag_full = tag_full.replace('--', '-')

    if '-' not in tag_full:
        return None

    parts = tag_full.split('-', 1)
    if len(parts) != 2:
        return None

    l1_part = parts[0].strip()
    l2_part = parts[1].strip()

    # 提取一级标签和序号(正则匹配数字)
    l1_match = re.match(r'^(.+?)(\d+)$', l1_part)
    if not l1_match:
        return None
    label_l1 = l1_match.group(1).strip()
    label_l1_order = int(l1_match.group(2))

    # 提取二级标签和序号
    l2_match = re.match(r'^(.+?)(\d+)$', l2_part)
    if not l2_match:
        return None
    label_l2 = l2_match.group(1).strip()
    label_l2_order = int(l2_match.group(2))

    return {
        "label_l1": label_l1,
        "label_l1_order": label_l1_order,
        "label_l2": label_l2,
        "label_l2_order": label_l2_order
    }

def get_pinyin_key(text):
    """
    获取中文拼音首字母排序键(用于表格视图按标签排序)
    返回: (pinyin_initials, numeric_part)
    pinyin_initials: 每个汉字的拼音首字母(小写),过滤掉空格等非字母字符
    numeric_part: 标签中的数字序号

    [FIX BUG-002] 过滤掉 lazy_pinyin 返回结果中的空格字符,
    确保 "A 站厅层" 的排序键为 "aztc"(而非 "a ztc"),
    避免空格影响拼音首字母的字典序比较。
    """
    if not text:
        return ('', 0)
    try:
        # 获取每个字符的拼音首字母,过滤空格等非字母字符
        py_list = pypinyin.lazy_pinyin(text)
        initials = ''.join(p[0].lower() for p in py_list if p and p[0].isalpha())
        return initials
    except Exception:
        return text.lower()


def get_sort_key_for_point(point):
    """
    生成点位排序键,用于按标签首字母和数字排序
    排序规则:
    - 有标签的点位优先,无标签的点位排最后
    - 一级标签拼音首字母 → 一级标签序号 → 二级标签拼音首字母 → 二级标签序号 → 点位名称
    """
    has_label = bool(point.get('label_l1') and point.get('label_l2'))

    if has_label:
        l1_key = get_pinyin_key(point['label_l1'])
        l1_order = point.get('label_l1_order', 0) or 0
        l2_key = get_pinyin_key(point['label_l2'])
        l2_order = point.get('label_l2_order', 0) or 0
        name_key = point.get('name', '')
        return (0, l1_key, l1_order, l2_key, l2_order, name_key)
    else:
        name_key = point.get('name', '')
        return (1, '', 0, '', 0, name_key)


def recalculate_sort_order_by_label():
    """
    根据标签对所有点位重新计算 sort_order
    排序规则:
    - 有标签的点位按标签排序(一级标签拼音首字母→序号→二级标签拼音首字母→序号→点位名称)
    - 无标签的点位排最后(按原有 sort_order)
    """
    db = get_db()
    try:
        rows = db.execute("""
            SELECT id, name, sort_order, label_l1, label_l1_order, label_l2, label_l2_order
            FROM patrol_points
            ORDER BY sort_order
        """).fetchall()

        # 构建排序列表
        points_with_label = []
        points_without_label = []

        for r in rows:
            p = {
                'id': r[0],
                'name': r[1],
                'sort_order': r[2],
                'label_l1': r[3],
                'label_l1_order': r[4],
                'label_l2': r[5],
                'label_l2_order': r[6]
            }
            has_label = bool(r[3] and r[5])
            if has_label:
                points_with_label.append(p)
            else:
                points_without_label.append(p)

        # 有标签的按标签排序
        points_with_label.sort(key=lambda p: get_sort_key_for_point(p))

        # 重新分配 sort_order
        new_order = 0
        for p in points_with_label:
            db.execute("UPDATE patrol_points SET sort_order = ? WHERE id = ?", (new_order, p['id']))
            new_order += 1

        # 无标签的保持原有相对顺序,排在后面
        for p in points_without_label:
            db.execute("UPDATE patrol_points SET sort_order = ? WHERE id = ?", (new_order, p['id']))
            new_order += 1

        db.commit()
        print(f"[SORT] 重新排序完成: {len(points_with_label)}个有标签点位, {len(points_without_label)}个无标签点位", flush=True)
        return {"sorted": len(points_with_label) + len(points_without_label), "with_label": len(points_with_label), "without_label": len(points_without_label)}
    except Exception as e:
        print(f"[ERROR] recalculate_sort_order_by_label: {e}", flush=True)
        raise
    finally:
        db.close()


def init_db():
    """初始化数据库，创建/升级所有必需的表
    
    PyInstaller 打包后首次启动时：
    1. 检查 EXE 同级目录 data/patrol.db 是否存在
    2. 如果不存在，从 _internal/data/patrol.db（只读打包目录）复制到可写目录
    3. 这样所有写操作（设备激活、session 等）都发生在可写目录
    """
    # 确保可写数据目录存在
    data_dir = os.path.dirname(DB_PATH)
    os.makedirs(data_dir, exist_ok=True)
    
    # PyInstaller 打包后：首次启动时从只读打包目录复制数据库
    if hasattr(sys, '_MEIPASS'):
        if not os.path.exists(DB_PATH):
            bundled_db = os.path.join(sys._MEIPASS, "data", "patrol.db")
            if os.path.exists(bundled_db):
                import shutil
                try:
                    shutil.copy2(bundled_db, DB_PATH)
                    print(f"[INIT] 已从打包目录复制数据库到可写目录: {DB_PATH}")
                    # 首次启动直接使用复制的数据库，跳过表创建
                    # 但记录中也已确保目录存在
                except Exception as e:
                    print(f"[ERROR] 复制数据库失败: {e}")
    
    db = get_db()
    
    # 创建基础表（使用 executescript）
    db.executescript("""
        CREATE TABLE IF NOT EXISTS patrol_points (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            sort_order INTEGER DEFAULT 0,
            is_active INTEGER DEFAULT 1
        );
        
        -- [SECURITY] 系统配置表 - 存储签名密钥等敏感配置
        CREATE TABLE IF NOT EXISTS system_config (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            type TEXT DEFAULT 'string',
            description TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        -- [SECURITY] 认证用户表 - 存储密码哈希、HMAC 密钥（v2.4 已移除 MAC 地址）
        CREATE TABLE IF NOT EXISTS auth_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            password_hash TEXT NOT NULL,
            admin_password_hash TEXT,
            role TEXT DEFAULT 'admin',
            hmac_key TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            active INTEGER DEFAULT 1
        );
        
        -- [SECURITY] Schema 版本控制表
        CREATE TABLE IF NOT EXISTS schema_version (
            version INTEGER PRIMARY KEY,
            applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            description TEXT
        );
    """)
    
    # [v2.4 简化] session_tokens 表（仅设备指纹，移除 mac_address）
    table_exists = db.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='session_tokens'").fetchone()
    if not table_exists:
        db.execute("""
            CREATE TABLE session_tokens (
                token_hash TEXT PRIMARY KEY,
                user_id TEXT NOT NULL,
                created_at REAL NOT NULL,
                expires_at REAL NOT NULL
            )
        """)
        db.execute("CREATE INDEX idx_session_expires ON session_tokens(expires_at)")
        print("[INIT] 创建 session_tokens 表")
    else:
        # 检查表结构，移除 mac_address 相关列和索引
        columns = [col["name"] for col in db.execute("PRAGMA table_info(session_tokens)")]
        if "mac_address" in columns:
            print("[INIT] 简化 session_tokens 表（移除 mac_address 字段）...")
            old_records = db.execute("SELECT token_hash, user_id, created_at, expires_at FROM session_tokens").fetchall()
            db.execute("DROP TABLE session_tokens")
            db.execute("""
                CREATE TABLE session_tokens (
                    token_hash TEXT PRIMARY KEY,
                    user_id TEXT NOT NULL,
                    created_at REAL NOT NULL,
                    expires_at REAL NOT NULL
                )
            """)
            db.execute("CREATE INDEX idx_session_expires ON session_tokens(expires_at)")
            for rec in old_records:
                db.execute("INSERT INTO session_tokens (token_hash, user_id, created_at, expires_at) VALUES (?, ?, ?, ?)",
                          (rec[0], rec[1], rec[2], rec[3]))
            print("[INIT] session_tokens 表简化完成（已移除 mac_address）")
        else:
            db.execute("CREATE INDEX IF NOT EXISTS idx_session_expires ON session_tokens(expires_at)")
        # 尝试删除旧索引（如果存在）
        try:
            db.execute("DROP INDEX IF EXISTS idx_session_mac")
        except Exception:
            pass
    
    # [v2.4 简化] device_activations 表（仅设备指纹，移除 mac_address）
    table_exists = db.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='mac_activations'").fetchone()
    
    # 如果旧表存在，先删除后重建（简化结构）
    if table_exists:
        print("[INIT] 重建 mac_activations → device_activations 表（仅设备指纹）...")
        db.execute("DROP TABLE IF EXISTS mac_activations")
    
    db.execute("""
        CREATE TABLE IF NOT EXISTS mac_activations (
            device_fingerprint TEXT UNIQUE NOT NULL,
            activated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    db.execute("CREATE INDEX IF NOT EXISTS idx_device_fp ON mac_activations(device_fingerprint)")
    print("[INIT] device_activations 表已就绪（仅设备指纹鉴权）")
    
    db.commit()
    # 添加标签字段 (SQLite 支持 ALTER TABLE ADD COLUMN)
    try:
        db.execute("ALTER TABLE patrol_points ADD COLUMN label_l1 TEXT")
    except: pass
    try:
        db.execute("ALTER TABLE patrol_points ADD COLUMN label_l1_order INTEGER")
    except: pass
    try:
        db.execute("ALTER TABLE patrol_points ADD COLUMN label_l2 TEXT")
    except: pass
    try:
        db.execute("ALTER TABLE patrol_points ADD COLUMN label_l2_order INTEGER")
    except: pass
    try:
        db.execute("ALTER TABLE patrol_points ADD COLUMN tag_full TEXT")
    except: pass
    # 初始化 schema 版本
    try:
        db.execute("INSERT OR IGNORE INTO schema_version (version, description) VALUES (25, 'v2.4 - 移除 MAC 地址鉴权，仅保留设备指纹鉴权')")
    except: pass
    for i, name in enumerate(DEFAULT_POINTS):
        db.execute("INSERT OR IGNORE INTO patrol_points (name, sort_order) VALUES (?, ?)", [name, i])
    db.commit()
    db.close()

def db_get_points():
    """
    获取所有巡更点位,按标签排序后返回
    排序规则:
    - 有标签的点位优先,无标签的点位排最后
    - 一级标签拼音首字母 → 一级标签序号 → 二级标签拼音首字母 → 二级标签序号 → 点位名称
    """
    db = get_db()
    rows = db.execute("SELECT id, name, sort_order, is_active, label_l1, label_l1_order, label_l2, label_l2_order, tag_full FROM patrol_points").fetchall()
    db.close()
    result = []
    for r in rows:
        point = {
            "id": r[0],
            "name": r[1],
            "sort_order": r[2],
            "is_active": r[3],
            "label_l1": r[4],
            "label_l1_order": r[5],
            "label_l2": r[6],
            "label_l2_order": r[7],
            "tag_full": r[8]
        }
        # 如果没有 tag_full 但有其他标签字段,自动生成
        if not point["tag_full"] and point["label_l1"] and point["label_l2"]:
            point["tag_full"] = f"{point['label_l1'].strip()} {point['label_l1_order']}-{point['label_l2'].strip()} {point['label_l2_order']}"
        result.append(point)
    # 按标签规则排序(不依赖 sort_order 列,支持动态排序)
    result.sort(key=lambda p: get_sort_key_for_point(p))
    return result

def db_get_stats():
    # [REMOVED] patrol_records 表已删除 - 不再存储巡更记录
    # 返回空统计信息
    return {
        "total_records": 0,
        "last_date": "-",
        "update_date": SERVER_MTIME,
        "source_stats": {}
    }

def get_active_points():
    """从数据库获取启用的巡更点列表(按标签排序)"""
    try:
        db = get_db()
        rows = db.execute("SELECT name, label_l1, label_l1_order, label_l2, label_l2_order FROM patrol_points WHERE is_active = 1").fetchall()
        db.close()
        if rows:
            points = []
            for r in rows:
                p = {"name": r[0], "label_l1": r[1], "label_l1_order": r[2], "label_l2": r[3], "label_l2_order": r[4]}
                points.append(p)
            points.sort(key=lambda p: get_sort_key_for_point(p))
            return [p["name"] for p in points]
    except Exception:
        pass
    return DEFAULT_POINTS

def get_route_points(route_name="线路一"):
    """获取指定路线的点位列表(按路线中的顺序)"""
    try:
        db = get_db()
        rows = db.execute("""
            SELECT pp.name
            FROM patrol_route_points prp
            JOIN patrol_points pp ON prp.point_id = pp.id
            JOIN patrol_routes pr ON prp.route_id = pr.id
            WHERE pr.name = ? AND pr.is_active = 1 AND pp.is_active = 1
            ORDER BY prp.sort_order
        """, (route_name,)).fetchall()
        db.close()
        if rows:
            return [r[0].replace(" ", "") for r in rows]
    except Exception as e:
        print(f"[ERROR] get_route_points({route_name}): {e}")
    # 降级:返回所有点位
    return get_active_points()

def get_all_route_points_maps():
    """获取所有路线的点位名称列表(按 sort_order 排序),用于路线匹配
    Returns:
        dict: {route_name: [point_name, ...]}
    """
    result = {}
    try:
        db = get_db()
        rows = db.execute("""
            SELECT pr.name AS route_name, pp.name AS point_name
            FROM patrol_routes pr
            JOIN patrol_route_points prp ON prp.route_id = pr.id
            JOIN patrol_points pp ON prp.point_id = pp.id
            WHERE pr.is_active = 1 AND pp.is_active = 1
            ORDER BY pr.id, prp.sort_order
        """).fetchall()
        db.close()
        for r in rows:
            rname = r["route_name"]
            if rname not in result:
                result[rname] = []
            result[rname].append(r["point_name"].replace(" ", ""))
    except Exception as e:
        print(f"[ERROR] get_all_route_points_maps: {e}")
    return result

def identify_route_from_records(records):
    """
    根据原数据中的点位集合,匹配最接近的路线配置

    Args:
        records: 原数据记录列表(每条含 "point" 字段)

    Returns:
        (route_name, route_points): 匹配到的路线名称和该路线的点位顺序列表
    """
    # 收集原数据中出现的所有点位
    record_points = set()
    for r in records:
        # [FIX 2026-06-22] 去除点位名称中的所有空格(包括中间空格)
        p = r.get("point", r.get("point_name", "")).replace(" ", "")
        if p:
            record_points.add(p)

    if not record_points:
        return ("线路一", get_route_points("线路一"))

    # 获取所有路线的点位配置
    all_routes = get_all_route_points_maps()

    if not all_routes:
        # 降级:返回全部点位
        return ("线路一", get_active_points())

    best_route = None
    best_score = -1
    best_missing = float('inf')

    for rname, rpoints in all_routes.items():
        route_set = set(rpoints)
        # 计算匹配度:原数据中属于该路线的点位数量
        matched = len(record_points & route_set)
        missing = len(route_set - record_points)

        # 评分:匹配数优先,缺少数次之
        score = matched * 100 - missing

        if matched > best_score or (matched == best_score and missing < best_missing):
            best_score = matched
            best_missing = missing
            best_route = (rname, rpoints)

    if best_route and best_score > 0:
        return best_route

    # 无匹配:使用默认路线
    return ("线路一", get_route_points("线路一"))

def validate_route_order(records, route_points):
    """
    按照路线点位顺序验证打点记录,标记异常

    Args:
        records: 打点记录列表(每条含 "point", "time", "round")
        route_points: 路线点位顺序列表

    Returns:
        dict: {
            "anomalies": [{"type": str, "seq": int, "point": str, "detail": str, ...}],
            "round_orders": {round_num: [point_in_order, ...]}  # 每轮实际打点顺序
        }
    """
    anomalies = []
    round_orders = {}
    point_to_order = {p: i for i, p in enumerate(route_points)}

    for rnd in range(1, 13):
        rnd_records = sorted(
            [r for r in records if r.get("round") == rnd],
            key=lambda r: r.get("time", "")
        )
        if not rnd_records:
            continue

        # 记录本轮实际打点顺序(按时间排序)
        actual_order = [r.get("point", "") for r in rnd_records]
        round_orders[rnd] = actual_order

        # 验证顺序:按路线点位顺序检查是否符合预期
        prev_order_idx = -1
        for i, r in enumerate(rnd_records):
            # [FIX 2026-06-22] 去除点位名称中的所有空格(包括中间空格)
            p = r.get("point", "").replace(" ", "")
            cur_idx = point_to_order.get(p, -1)

            if cur_idx == -1:
                # 点位不在路线中
                anomalies.append({
                    "type": "unknown_point",
                    "round": rnd,
                    "seq": r.get("seq", i+1),
                    "point": p,
                    "detail": f"点位 '{p}' 不属于此路线"
                })
            elif prev_order_idx != -1 and cur_idx < prev_order_idx:
                # [路线验证] 顺序倒退检查:当前点位顺序索引小于上一点位
                if cur_idx == prev_order_idx:
                    # 同一点位第二下,正常
                    pass
                else:
                    # 跳过点位或顺序颠倒
                    anomalies.append({
                        "type": "order_reversal",
                        "round": rnd,
                        "seq": r.get("seq", i+1),
                        "point": p,
                        "expected_after": route_points[prev_order_idx] if 0 <= prev_order_idx < len(route_points) else "?",
                        "detail": f"点位 '{p}'(路线顺序#{cur_idx}) 出现在 '{route_points[prev_order_idx]}'(路线顺序#{prev_order_idx}) 之后,顺序不符合路线配置"
                    })

            # 更新上一点位在路线中的索引(即使是同一点位的第二下)
            prev_order_idx = cur_idx
        point_times = {}
        for r in rnd_records:
            p = r.get("point", "")
            t = r.get("time", "")
            if p not in point_times:
                point_times[p] = []
            point_times[p].append(t)

        for p, times in point_times.items():
            if len(times) > 2:
                anomalies.append({
                    "type": "excessive_taps",
                    "round": rnd,
                    "point": p,
                    "count": len(times),
                    "detail": f"点位 '{p}' 打点{len(times)}次(标准2次),超出{len(times)-2}次"
                })

    return {"anomalies": anomalies, "round_orders": round_orders}

def get_all_routes():
    """获取所有启用的路线"""
    try:
        db = get_db()
        rows = db.execute("""
            SELECT id, name, description, is_active, created_at
            FROM patrol_routes
            WHERE is_active = 1
            ORDER BY id
        """).fetchall()
        db.close()
        return [{"id": r["id"], "name": r["name"], "description": r["description"]} for r in rows]
    except Exception as e:
        print(f"[ERROR] get_all_routes: {e}")
    return []

def get_points_with_ids():
    """获取所有启用的点位(带 ID)"""
    try:
        db = get_db()
        rows = db.execute("SELECT id, name, sort_order FROM patrol_points WHERE is_active = 1 ORDER BY sort_order").fetchall()
        db.close()
        return [{"id": r["id"], "name": r["name"], "sort_order": r["sort_order"]} for r in rows]
    except Exception as e:
        print(f"[ERROR] get_points_with_ids: {e}")
    return []

def get_route_detail(route_name):
    """获取路线的详细信息(包含点位列表)"""
    try:
        db = get_db()
        # 获取路线信息
        route = db.execute("SELECT id, name, description FROM patrol_routes WHERE name=?", (route_name,)).fetchone()
        if not route:
            db.close()
            return None
        # 获取点位列表 (包含 tag_full)
        points = db.execute("""
            SELECT pp.id, pp.name, pp.tag_full, prp.sort_order
            FROM patrol_route_points prp
            JOIN patrol_points pp ON prp.point_id = pp.id
            WHERE prp.route_id = ?
            ORDER BY prp.sort_order
        """, (route["id"],)).fetchall()
        db.close()
        return {
            "id": route["id"],
            "name": route["name"],
            "description": route["description"],
            "points": [{"id": p["id"], "name": p["name"], "tag_full": p["tag_full"], "sort_order": p["sort_order"]} for p in points]
        }
    except Exception as e:
        print(f"[ERROR] get_route_detail({route_name}): {e}")
    return None

# === 全局常量 ===
# LOW-005: 魔法数字 - 时间间隔常量 (单位：秒)
TAPS_PER_POINT = 2  # 每个点位标准打点次数:2 次

# === 时间间隔常量 ===
SAME_POINT_GAP_MIN = 10  # 同点位间隔最小值 (秒)
SAME_POINT_GAP_MAX = 30  # 同点位间隔最大值 (秒)
DIFF_POINT_GAP_MIN = 30  # 异点位间隔最小值 (秒)
DIFF_POINT_GAP_MAX = 55  # 异点位间隔最大值 (秒)

def calculate_next_time(current_time, is_same_point=False):
    """
    计算下一个打点时间,根据间隔规则自动累加时间

    Args:
        current_time: datetime 对象,当前时间
        is_same_point: bool,是否为同一点位的下一次打点

    Returns:
        datetime 对象,下一个打点时间

    规则:
        - 同点位间隔:10~40 秒随机(模拟同一位置打两下的时间差)
        - 异点位间隔:40~90 秒随机(模拟点位间移动时间)
    """
    if is_same_point:
        # 同一点位第二下间隔:10~40 秒
        gap_seconds = random.randint(SAME_POINT_GAP_MIN, SAME_POINT_GAP_MAX)
    else:
        # 不同点位之间间隔:40~90 秒
        gap_seconds = random.randint(DIFF_POINT_GAP_MIN, DIFF_POINT_GAP_MAX)

    return current_time + datetime.timedelta(seconds=gap_seconds)

def get_same_point_gap():
    """获取同点位间隔范围(秒)"""
    return (SAME_POINT_GAP_MIN, SAME_POINT_GAP_MAX)

def get_diff_point_gap():
    """获取异点位间隔范围(秒)"""
    return (DIFF_POINT_GAP_MIN, DIFF_POINT_GAP_MAX)

def get_random_route_for_round(round_num, date_str=None):
    """
    为指定轮次随机选择一条已创建的路线,返回(route_name, route_points)
    不同 date 使用不同种子,确保同日期内随机但可复现(不同日期变化)
    Falls back to all active points and '线路一' if no routes in DB.
    """
    try:
        db = get_db()
        route_rows = db.execute(
            "SELECT id, name FROM patrol_routes WHERE is_active = 1 ORDER BY id"
        ).fetchall()
        db.close()

        if not route_rows:
            return ("线路一", get_active_points())

        # 使用日期+轮次作为随机种子(同日期同轮次稳定,不同日期或不同轮次不一样)
        seed_src = f"{date_str or 'default'}_round_{round_num}"
        seed = int(hashlib.md5(seed_src.encode()).hexdigest()[:8], 16)
        rng = random.Random(seed)

        selected = rng.choice(route_rows)
        route_name = selected["name"]
        route_points = get_route_points(route_name)
        return (route_name, route_points if route_points else get_active_points())
    except Exception as e:
        print(f"[ERROR] get_random_route_for_round(round={round_num}): {e}")
        return ("线路一", get_active_points())


def process_all_rounds(records, points, date_str, route_name="线路一", random_empty_rounds=True, default_device=""):
    """
    处理所有 12 轮数据的主流程

    新增逻辑 (random_empty_rounds=True):整轮无数据时，为该轮随机选择一条路线进行补充
     - 已有数据的轮次：使用传入的 route_points(原识别路线)
     - 整轮缺失的轮次：随机从 patrol_routes 选一条，使用该路线的点位补充
     - random_empty_rounds=False 时 (如生成模式),所有空轮使用统一的 route_name/points
    """
    all_processed = []
    for round_num in range(1, 13):
        round_records = [r for r in records if r.get("round") == round_num]
        print(f"[DEBUG] Round {round_num}: {len(round_records)} records, random_empty_rounds={random_empty_rounds}")
        if len(round_records) == 0 and random_empty_rounds:
            # 整轮缺失：随机选择路线
            random_route_name, random_route_points = get_random_route_for_round(round_num, date_str)
            print(f"[INFO] Round {round_num}: 整轮缺失，随机选择路线 " + str(random_route_name) + " (" + str(len(random_route_points)) + "个点位)")
            processed = process_round(
                round_records, round_num, random_route_points,
                date_str, route_name=random_route_name, default_device=default_device
            )
            # 标记补充数据来源 (source='supplemented') 和使用的路线名称
            for r in processed:
                r["source"] = "supplemented"
                r["route"] = random_route_name
        else:
            # 有数据的轮次 (或生成模式):使用原路线
            processed = process_round(round_records, round_num, points, date_str, route_name=route_name, default_device=default_device)
        all_processed.extend(processed)
    all_processed = append_extra_records(all_processed, points, date_str, route_name=route_name)
    final_records = apply_global_limit(all_processed, points)
    # [FIX 2026-06-22] 排序方式：轮次 → 时间正序
    # 用户铁律：原数据 (source="user") 的时间和点位绝对不能动
    # 补充数据已通过时间缝隙插入逻辑正确分配到原数据的时间间隙中
    # 最终导出按时间正序排列，符合巡更记录的实际时间顺序
    #   1. 第一优先级：轮次 (round) — 跨轮次不乱序
    #   2. 第二优先级：时间 (time) — 严格按时间正序排列
    # 注意：如果原数据中同点位记录时间相差较大，导出中也会分开（原数据不变）
    final_records.sort(key=lambda r: (r.get("round", 0), r.get("time", "")))
    return final_records


def process_round(round_records, round_num, points, date_str, route_name="线路一", default_device=""):
    """处理单轮数据的主流程"""
    standard = len(points) * TAPS_PER_POINT
    count = len(round_records)
    if count == standard:
        return review_round_data_std(round_records, round_num, points, date_str, route_name=route_name, default_device=default_device)
    elif count == standard + 1:
        return review_round_data_std_plus_1(round_records, round_num, points, date_str, route_name=route_name, default_device=default_device)
    elif count == standard + 2:
        return review_round_data_std_plus_2(round_records, round_num, points, date_str, route_name=route_name, default_device=default_device)
    elif count < standard:
        return supplement_round_direct(round_records, round_num, points, date_str, route_name=route_name, default_device=default_device)
    else:
        round_records = round_records[:standard + 2]
        return review_round_data_std_plus_2(round_records, round_num, points, date_str, route_name=route_name, default_device=default_device)

def review_round_data_std(round_records, round_num, points, date_str, route_name="线路一", default_device=""):
    """数据审查算法 (STANDARD 条版)"""
    point_counts = {}
    for r in round_records:
        p = r["point"].strip()
        point_counts[p] = point_counts.get(p, 0) + 1
    has_excess = any(cnt > 2 for cnt in point_counts.values())
    if has_excess:
        # 修复：构建新列表代替原地删除，避免删除错误记录
        new_records = []
        point_temp_counts = {}
        round_records.sort(key=lambda r: r["time"])  # 先按时间排序
        for r in round_records:
            p = r["point"].strip()
            cnt = point_temp_counts.get(p, 0)
            if cnt < 2:
                new_records.append(r)
                point_temp_counts[p] = cnt + 1
            # 超过 2 次的记录直接丢弃
        round_records = new_records
        point_counts = {}
        for r in round_records:
            p = r["point"].strip()
            point_counts[p] = point_counts.get(p, 0) + 1
    missing = []
    for p in points:
        p_clean = p.strip()
        cnt = point_counts.get(p_clean, 0)
        if cnt == 0:
            missing.append((p_clean, 2))
        elif cnt == 1:
            missing.append((p_clean, 1))
    if missing:
        round_records = supplement_round_logic(round_records, round_num, points, date_str, missing, route_name=route_name, default_device=default_device)
    return round_records

def review_round_data_std_plus_1(round_records, round_num, points, date_str, route_name="线路一", default_device=""):
    """数据审查算法 (STANDARD+1 条版)"""
    standard = len(points) * TAPS_PER_POINT
    point_counts = {}
    for r in round_records:
        p = r["point"].strip()
        point_counts[p] = point_counts.get(p, 0) + 1
    missing = []
    for p in points:
        p_clean = p.strip()
        cnt = point_counts.get(p_clean, 0)
        if cnt == 0:
            missing.append((p_clean, 2))
        elif cnt == 1:
            missing.append((p_clean, 1))
    if missing:
        round_records = supplement_round_logic(round_records, round_num, points, date_str, missing, route_name=route_name, default_device=default_device)
        point_counts = {}
        for r in round_records:
            p = r["point"].strip()
            point_counts[p] = point_counts.get(p, 0) + 1
    total = len(round_records)
    if total > standard + 2:
        # 修复:构建新列表代替原地删除
        excess_points = [(p, cnt) for p, cnt in point_counts.items() if cnt > 2]
        excess_points.sort(key=lambda x: -x[1])

        # 确定每个点位应该保留的数量
        keep_limits = {}
        for i, (p, cnt) in enumerate(excess_points):
            if i == 0 and cnt >= 3:
                keep_limits[p] = min(cnt, 3)  # 第一个超次点位保留最多 3 次
            else:
                keep_limits[p] = 2  # 其余点位只保留 2 次

        # 构建新列表
        new_records = []
        point_temp_counts = {}
        round_records.sort(key=lambda r: r["time"])
        for r in round_records:
            p = r["point"].strip()
            limit = keep_limits.get(p, 2)
            cnt = point_temp_counts.get(p, 0)
            if cnt < limit:
                new_records.append(r)
                point_temp_counts[p] = cnt + 1
        round_records = new_records
    return round_records

def review_round_data_std_plus_2(round_records, round_num, points, date_str, route_name="线路一", default_device=""):
    """数据审查算法 (STANDARD+2 条版)"""
    standard = len(points) * TAPS_PER_POINT
    point_counts = {}
    for r in round_records:
        p = r["point"].strip()
        point_counts[p] = point_counts.get(p, 0) + 1
    missing = []
    for p in points:
        p_clean = p.strip()
        cnt = point_counts.get(p_clean, 0)
        if cnt == 0:
            missing.append((p_clean, 2))
        elif cnt == 1:
            missing.append((p_clean, 1))
    if missing:
        round_records = supplement_round_logic(round_records, round_num, points, date_str, missing, route_name=route_name, default_device=default_device)
        point_counts = {}
        for r in round_records:
            p = r["point"].strip()
            point_counts[p] = point_counts.get(p, 0) + 1
    total = len(round_records)
    if total > standard + 2:
        # 修复:构建新列表代替原地删除
        excess_points = [(p, cnt) for p, cnt in point_counts.items() if cnt > 2]
        excess_points.sort(key=lambda x: -x[1])

        # 确定每个点位应该保留的数量
        keep_limits = {}
        kept_4 = False
        kept_3_count = 0
        for p, cnt in excess_points:
            if not kept_4 and cnt >= 4:
                keep_limits[p] = 4
                kept_4 = True
            elif kept_3_count < 2 and cnt >= 3:
                keep_limits[p] = 3
                kept_3_count += 1
            else:
                keep_limits[p] = 2

        # 构建新列表
        new_records = []
        point_temp_counts = {}
        round_records.sort(key=lambda r: r["time"])
        for r in round_records:
            p = r["point"].strip()
            limit = keep_limits.get(p, 2)
            cnt = point_temp_counts.get(p, 0)
            if cnt < limit:
                new_records.append(r)
                point_temp_counts[p] = cnt + 1
        round_records = new_records
    return round_records

def supplement_round_direct(round_records, round_num, points, date_str, route_name="线路一", default_device=""):
    """直接补充算法 (用于 count < STANDARD 的情况)"""
    # 修复：点位名称统一.strip() 处理，避免空格导致统计错误
    point_counts = {}
    for r in round_records:
        p = r["point"].strip()
        point_counts[p] = point_counts.get(p, 0) + 1
    # 修复：构建新列表代替原地删除
    new_records = []
    point_temp_counts = {}
    round_records.sort(key=lambda r: r["time"])
    for r in round_records:
        p = r["point"].strip()
        cnt = point_temp_counts.get(p, 0)
        if cnt < 2:
            new_records.append(r)
            point_temp_counts[p] = cnt + 1
    round_records = new_records
    point_counts = {}
    for r in round_records:
        p = r["point"].strip()
        point_counts[p] = point_counts.get(p, 0) + 1
    missing = []
    for p in points:
        p_clean = p.strip()
        cnt = point_counts.get(p_clean, 0)
        if cnt == 0:
            missing.append((p_clean, 2))
        elif cnt == 1:
            missing.append((p_clean, 1))
    if missing:
        round_records = supplement_round_logic(round_records, round_num, points, date_str, missing, route_name=route_name, default_device=default_device)
    return round_records

def supplement_round_logic(round_records, round_num, points, date_str, missing, route_name="线路一", default_device=""):
    """
    补充算法核心逻辑 - 标签层级优先 (L1→L2 循环填充)

    三级优先级:
    1. 按 L1→L2 循环填充 (已有数据的 L1 标签)
    2. 其他 L1 标签 (无数据的 L1)
    3. 路线兜底 (随机选取路线，剔除已有 L1 的点位)

    时间间隔 (优先级降级):
    - 尝试 1:同点位 10-40 秒，异点位 40-90 秒
    - 尝试 2:压缩范围，保持最小值
    - 尝试 3:同点位 5-30 秒，异点位 30-60 秒
    - 尝试 4+:最小值继续减 5 秒
    """
    start_hour, end_hour = ROUND_TIME_RANGES.get(round_num, (0, 3))
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    round_start = dt.replace(hour=start_hour, minute=5, second=0)
    if end_hour == 24:
        round_end = dt.replace(hour=23, minute=59, second=59)
    else:
        round_end = dt.replace(hour=end_hour, minute=0, second=0)

    # 计算需要补充的总次数
    total_taps_needed = sum(taps for _, taps in missing)
    if total_taps_needed == 0:
        return round_records

    # === 步骤 1: 查询点位的标签信息 ===
    point_labels = {}
    db = None
    try:
        db = get_db()
        placeholders = ','.join(['?' for _ in points])
        rows = db.execute(
            f"SELECT name, label_l1, label_l1_order, label_l2, label_l2_order FROM patrol_points WHERE name IN ({placeholders})",
            points
        ).fetchall()
        # [FIX 2026-06-22 #7] 使用''.join(p.split())去除所有空白字符（包括空格、制表符、全角空格等）
        for r in rows:
            key = ''.join(r[0].split())
            point_labels[key] = {
                'label_l1': r[1],
                'label_l1_order': r[2] or 0,
                'label_l2': r[3],
                'label_l2_order': r[4] or 0
            }
    except Exception as e:
        print(f"[WARN] supplement_round_logic: 查询标签信息失败:{e}")
    finally:
        if db:
            db.close()

    # === 步骤 2: 构建点位计数映射 ===
    # [FIX 2026-06-22 #7] 使用''.join(p.split())去除所有空白字符，确保与标签查询一致
    point_counts = {}
    for r in round_records:
        p = ''.join(r["point"].split())
        point_counts[p] = point_counts.get(p, 0) + 1

    # === 步骤 3: 识别已有数据的 L1 标签 ===
    # [FIX 2026-06-22 #7] 使用''.join(p.split())去除所有空白字符
    l1_with_data = set()
    for r in round_records:
        p = ''.join(r["point"].split())
        if p in point_labels:
            l1_with_data.add(point_labels[p]['label_l1'])

    # === 步骤 4: 按 L1 序号排序 ===
    l1_list = sorted(l1_with_data, key=lambda l1: next((pl['label_l1_order'] for pl in point_labels.values() if pl['label_l1'] == l1), 0))

    # === 步骤 5: [FIX 2026-06-22 #6] 按路线顺序构建补充列表(移除 L1 标签分组逻辑)===
    # 问题根因:原 L1 分组逻辑导致点位顺序混乱、tap_count 分配错误、产生 extra 记录
    # 修复方案:直接按 route_points 顺序遍历,确定每个点位需要补充几条记录
    fill_order = []  # [(point_name, tap_count), ...] 按路线顺序

    # 获取路线点位顺序（已去除所有空白字符）
    route_points_raw = get_route_points(route_name)
    route_points = [''.join(p.split()) for p in route_points_raw]

    # 按路线顺序遍历每个点位,确定需要补充的记录数
    for point_name in route_points:
        count = point_counts.get(point_name, 0)
        if count == 0:
            # 完全缺失:需要补充 2 条
            fill_order.append((point_name, 2))
        elif count == 1:
            # 只有 1 条:需要补充 1 条
            fill_order.append((point_name, 1))
        # count >= 2: 已有足够记录,不补充


    # === 步骤 6: [FIX 2026-06-22 #3] 用户要求:先补充点位,再根据点位补充时间 ===
    # 原逻辑问题:按原记录顺序遍历,导致补充记录时间不正序
    # 修复方案:按 fill_order 顺序生成所有补充记录,使用全局 prev_time 分配时间
    # 1. 先将原记录按时间排序保留
    sorted_records = sorted(round_records, key=lambda r: r["time"])

    # 2. 找到最后一条已有记录的时间 (用于补充记录的基准时间)
    existing_times = [datetime.datetime.strptime(r["time"], "%Y-%m-%d %H:%M:%S") for r in round_records]
    # [FIX 2026-06-22] 当轮次无原数据时,使用 round_start 作为基准时间
    # [FIX 2026-06-22 #4] 当原数据时间超出 round_end 时(如测试数据插入到错误轮次),视为该轮无有效原数据,使用 round_start
    if existing_times:
        max_existing = max(existing_times)
        if max_existing > round_end:
            # 原数据时间超出轮次窗口,视为无效原数据,使用 round_start
            last_existing_time = round_start
        else:
            last_existing_time = max_existing
    else:
        last_existing_time = round_start

    print(f"[DEBUG] 第{round_num}轮:round_start={round_start}, round_end={round_end}")
    print(f"[DEBUG] 第{round_num}轮:existing_times count={len(existing_times)}, last_existing_time={last_existing_time}")

    # [FIX 2026-06-22 #7] 使用''.join(p.split())去除所有空白字符
    supplement_needed = {}
    for point_name, tap_count in fill_order:
        pn_clean = ''.join(point_name.split())
        supplement_needed[pn_clean] = tap_count

    # 按 route_points 顺序(正序)构建补充点位列表
    # 每个元素: (point_name, tap_count)
    points_to_supplement = []
    for point_name in route_points:  # ✅ 正序遍历路线顺序
        if point_name not in supplement_needed:
            continue
        tap_count = supplement_needed[point_name]
        points_to_supplement.append((point_name, tap_count))
        print(f"[INFO] supplement_round_logic: 路线顺序将补充点位 '{point_name}' 共{tap_count}条")

    # === 阶段 2: [FIX 2026-06-22] 重构时间插入逻辑 - 按时间缝隙插入补充记录 ===
    #
    # 问题根因:原逻辑使用 point_max_times 独立计算每个点位的补充时间,
    # 导致补充记录的时间与原数据的时间不连续,无法正确插入到时间缝隙中。
    #
    # 修复方案:
    #   1. 构建合并记录列表:原数据记录(按时间正序) + 补充点位占位(按路线顺序)
    #   2. 按 (路线顺序索引, 时间) 排序,实现原数据与补充点位交替排列
    #   3. 按顺序遍历:原数据记录保持原时间不变;
    #      补充记录基于前一条记录(可能是原数据或补充记录)按间隔规则计算时间
    #
    # 间隔规则:
    #   - 同点位(与前一条记录点位相同):10-40 秒
    #   - 异点位(与前一条记录点位不同):40-90 秒
    #
    # 铁律约束:原数据(source="user")的时间和点位绝对不能修改
    #

    # 2a. 创建补充点位占位记录 (时间占位 None，稍后分配)
    # [FIX 2026-07-01] device 字段继承修复：如果该点位无原记录，从同轮其他点位获取 device
    # 获取同轮所有记录中有 device 的值作为 fallback
    all_devices_in_round = [r.get("device", "") for r in sorted_records if r.get("device", "")]
    fallback_device = all_devices_in_round[0] if all_devices_in_round else default_device
    
    supplement_placeholders = []
    for point_name, tap_count in points_to_supplement:
        pt_clean = ''.join(point_name.split())
        # 从该点位的原记录中获取 device 值 (取第一条原记录的 device)
        point_orig_records = [r for r in sorted_records if ''.join(r.get("point", "").split()) == pt_clean]
        if point_orig_records:
            device_value = point_orig_records[0].get("device", "")
        else:
            # 该点位无原记录:从同轮其他点位继承 device
            device_value = fallback_device
        print(f"[DEBUG] 补充点位 '{point_name}': 原记录数={len(point_orig_records)}, device='{device_value}'", flush=True)
        for tap_idx in range(tap_count):
            supplement_placeholders.append({
                "seq": 0,
                "point": pt_clean,
                "time": None,  # 占位，稍后分配具体时间
                "device": device_value,
                "event": "合格",
                "inspector": "",
                "route": route_name,
                "round": round_num,
                "source": "supplemented",
                "_order": len(supplement_placeholders)  # 保持补充点位内部的相对顺序
            })

    # 2b. 构建路线点位顺序映射
    point_order_map = {''.join(p.split()): i for i, p in enumerate(route_points)}

    # 2c. 合并原记录和补充占位
    #
    # [时间缝隙插入逻辑]
    # 构建合并列表时,按路线点位顺序遍历:
    #   1. 取出该点位所有原记录(已按时间排序)
    #   2. 取出该点位所有补充占位
    #   3. 顺序排列为:每条原记录后跟其补充记录
    #
    merged_records = []

    for pt_name in route_points:
        pt_clean = ''.join(pt_name.split())
        # 获取该点位所有原记录(按时间排序)
        point_originals = [r for r in sorted_records if ''.join(r.get("point", "").split()) == pt_clean]
        # 获取该点位所有补充占位
        point_supplements = [s for s in supplement_placeholders if s["point"] == pt_clean]

        if not point_originals:
            # 该点位完全没有原记录:所有补充直接添加(稍后基于前一条已有记录计算时间)
            merged_records.extend(point_supplements)
        else:
            # 该点位有原记录:在原记录之间和之后插入补充
            # 策略:迭代原记录,每条原记录后尝试插入一条补充记录
            supp_iter = iter(point_supplements)
            for orig in point_originals:
                merged_records.append(orig)
                # 在原记录后插入一条补充(如果有剩余)
                try:
                    merged_records.append(next(supp_iter))
                except StopIteration:
                    pass

            # 如果还有剩余补充(在最后一条原记录之后),全部添加
            try:
                while True:
                    merged_records.append(next(supp_iter))
            except StopIteration:
                pass

    print(f"[DEBUG] 第{round_num}轮：合并后共{len(merged_records)}条记录 (原{len(sorted_records)}+补充{len(supplement_placeholders)})")

    # === 阶段 3: 按顺序遍历合并列表,为补充记录分配时间 ===
    # [FIX 2026-06-23] 检查原数据时间跨度
    original_times = [datetime.datetime.strptime(r["time"], "%Y-%m-%d %H:%M:%S")
                      for r in sorted_records if r.get("time")]
    if len(original_times) >= 2:
        original_span = (max(original_times) - min(original_times)).total_seconds()
        if original_span > 2700:  # 45 分钟 = 2700 秒
            print(f"[WARN] 第{round_num}轮：原数据时间跨度{original_span/60:.1f}分钟，超过 45 分钟限制")
            print(f"[WARN] 建议：原数据时间过于分散，补充后总时间可能仍超过 45 分钟")

    allowed_gap_configs = [
        {'same': (10, 30), 'diff': (30, 55)},
        {'same': (8, 30), 'diff': (30, 70)},
        {'same': (5, 20), 'diff': (20, 50)},
        {'same': (3, 10), 'diff': (10, 30)},
        {'same': (2, 8), 'diff': (8, 20)},
        {'same': (1, 5), 'diff': (5, 15)},
    ]

    # 最终结果列表
    result_records = []

    # [TEMP FIX 2026-06-23] 获取原数据第一条时间，用于向前推算补充数据
    original_first_time = None
    original_last_time = None
    for r in sorted_records:
        if r.get("time"):
            t = datetime.datetime.strptime(r["time"], "%Y-%m-%d %H:%M:%S")
            if original_first_time is None or t < original_first_time:
                original_first_time = t
            if original_last_time is None or t > original_last_time:
                original_last_time = t
    
    # 计算补充数据总条数，估算总时长
    supplement_count = len([r for r in merged_records if r.get("time") is None])
    # 估算：平均间隔 40 秒，总时长 = (count-1) * 40 秒
    estimated_supplement_duration = (supplement_count - 1) * 40 if supplement_count > 1 else 0

    # [TEMP FIX 2026-06-23] 计算补充数据的起始时间：在 round_start 到原数据第一条时间之间随机
    # 修复：起始时间太集中（都在原数据第一条 -12 分钟），需要增加离散性
    if original_first_time:
        # 可用时间窗口：round_start 到 original_first_time
        available_window = (original_first_time - round_start).total_seconds()
        # 估算补充数据需要的最小时间（按最大间隔 55 秒计算）
        min_supplement_duration = (supplement_count - 1) * 20 if supplement_count > 1 else 0  # 最小间隔 20 秒
        max_supplement_duration = (supplement_count - 1) * 70 if supplement_count > 1 else 0  # 最大间隔 70 秒
        
        # 随机选择起始时间：确保补充数据能在原数据第一条之前完成
        # 策略：在 [round_start, original_first_time - min_supplement_duration] 范围内随机
        latest_start = original_first_time - datetime.timedelta(seconds=min_supplement_duration)
        if latest_start < round_start:
            latest_start = round_start
        
        # 随机起始时间（增加离散性）
        random_range = (latest_start - round_start).total_seconds()
        if random_range > 0:
            random_offset = random.uniform(0, random_range)
            supplement_start_time = round_start + datetime.timedelta(seconds=random_offset)
        else:
            supplement_start_time = round_start
        
        print(f"[INFO] 第{round_num}轮：原数据第一条{original_first_time.strftime('%H:%M:%S')}, " +
              f"补充{supplement_count}条，随机起始时间{supplement_start_time.strftime('%H:%M:%S')} " +
              f"(窗口：{round_start.strftime('%H:%M:%S')} - {latest_start.strftime('%H:%M:%S')})")
    else:
        # 整轮缺失：起始时间 = round_start + 5~25 分钟随机（增加离散性）
        random_offset_minutes = random.uniform(5, 25) * 60  # 转为秒
        supplement_start_time = round_start + datetime.timedelta(seconds=random_offset_minutes)
        print(f"[INFO] 第{round_num}轮：整轮缺失，随机起始时间{supplement_start_time.strftime('%H:%M:%S')} (偏移{random_offset_minutes/60:.1f}分钟)")

    # 补充数据的时间游标（独立于原数据）
    supplement_time_cursor = supplement_start_time
    last_supplement_point = None


    # 前一条有真实时间记录(原数据)的时间基准
    last_time_base = None
    last_point_clean = None  # 上一条有真实时间记录的点位(去空格后)

    # 遍历合并列表，为每个补充记录分配时间
    # [FIX 2026-06-23] 优先匹配同点位原数据时间，确保同点位两条记录相邻
    
    # 构建原数据点位→时间列表的映射
    original_point_times = {}
    for r in sorted_records:
        if r.get("time"):
            pt = ''.join(r.get("point", "").split())
            if pt not in original_point_times:
                original_point_times[pt] = []
            original_point_times[pt].append(datetime.datetime.strptime(r["time"], "%Y-%m-%d %H:%M:%S"))
    
    # 排序每个点位的时间列表（确保按时间顺序）
    for pt in original_point_times:
        original_point_times[pt].sort()
    
    # 跟踪每个点位已使用的原数据时间索引
    point_time_index = {pt: 0 for pt in original_point_times}
    
    for rec in merged_records:
        if rec["time"] is not None:
            # 原数据记录：保持原时间不变
            result_records.append(rec)
            last_time_base = datetime.datetime.strptime(rec["time"], "%Y-%m-%d %H:%M:%S")
            last_point_clean = ''.join(rec.get("point", "").split())
        else:
            # 补充记录：[FIX] 优先匹配同点位原数据时间
            current_point_clean = rec["point"]
            
            # 检查同点位是否有原数据时间
            if current_point_clean in original_point_times and point_time_index.get(current_point_clean, 0) < len(original_point_times[current_point_clean]):
                # 同点位有原数据：补充时间 = 原时间 + 随机间隔 (10-40 秒)
                orig_times = original_point_times[current_point_clean]
                idx = point_time_index[current_point_clean]
                base_time = orig_times[idx]
                gap = random.randint(10, 40)  # 同点位间隔 10-40 秒
                assigned_time = base_time + datetime.timedelta(seconds=gap)
                point_time_index[current_point_clean] = idx + 1
                print(f"[INFO] 补充'{current_point_clean}': 基于原数据{base_time.strftime('%H:%M:%S')} + {gap}秒 = {assigned_time.strftime('%H:%M:%S')}")
            else:
                # 同点位无原数据：基于补充数据时间游标
                is_same_point = (last_supplement_point == current_point_clean) if last_supplement_point else False
                min_gap, max_gap = (10, 40) if is_same_point else (20, 70)
                gap = random.randint(min_gap, max_gap)
                assigned_time = supplement_time_cursor + datetime.timedelta(seconds=gap)
                
                # [FIX] 只有当 supplement_time_cursor < original_first_time 时才压缩
                # 避免 supplement_time_cursor 已超过 original_first_time 时 gap=1 的问题
                if original_first_time and supplement_time_cursor < original_first_time and assigned_time > original_first_time:
                    # 压缩间隔，确保在原数据之前
                    gap = max(10, int((original_first_time - supplement_time_cursor).total_seconds() // 2))
                    assigned_time = supplement_time_cursor + datetime.timedelta(seconds=gap)
                supplement_time_cursor = assigned_time
                last_supplement_point = current_point_clean
            
            rec["time"] = assigned_time.strftime("%Y-%m-%d %H:%M:%S")
            result_records.append(rec)
            
            # 更新时间基准
            if last_time_base is None:
                last_time_base = assigned_time
                last_point_clean = current_point_clean

    # 规范化所有记录的点位名称(去空格)
    for r in result_records:
        if "point" in r and r["point"]:
            r["point"] = ''.join(r["point"].split())

    print(f"[DEBUG] 第{round_num}轮：最终结果{len(result_records)}条记录")

    return result_records


def append_extra_records(records, points, date_str, route_name="线路一"):
    """末尾填充算法:随机 5~25 条均分到各轮,插入到每轮已有数据的时间缝隙中 (严格介于上下两条之间)

    Bug 修复 (2026-06-18):
    1. Extra 记录点位分配极度不均 → 使用轮次偏移全局均匀分配
    2. Extra 记录时间早于 system 记录 → extra 时间在该点位该轮的 system 记录之后
    """
    date_int = int(hashlib.md5(date_str.encode()).hexdigest()[:8], 16)
    rng = random.Random(date_int)
    extra_total = rng.randint(5, 8)
    if extra_total == 0:
        return records
    extra_per_round = extra_total // 12
    extra_remainder = extra_total % 12

    # ===== Bug 修复 #1: 全局均匀分配 extra 记录到各点位 =====
    # 原逻辑问题:每轮从 pts_order 的前 n 个点位分配,由于每轮第一个点位固定为"B 端办公区东大门",
    # 导致 pts_order 总是从"消防通道内"开始,extra 记录集中分配给前 2 个点位
    # 修复方案:使用轮次偏移,使 extra 记录均匀分布到所有点位
    extra_assignments = []  # [(round, point_index), ...]
    for rnd in range(1, 13):
        n = extra_per_round + (1 if rnd <= extra_remainder else 0)
        if n == 0:
            continue
        # 该轮的 n 条 extra 记录,使用轮次偏移均匀分配到不同点位
        # [FIX 2026-06-20] 修复当 n=1 时 start_offset 为 (rnd-1)*25,由于 len(points)=25 导致取模后始终为 0
        # 修正为使用轮次序号直接偏移,确保即使 n=1 也能均匀分布到不同点位
        if n > 0:
            start_offset = (rnd - 1) % len(points)  # 每轮偏移 1 位,避免 wrap 到同一索引
            for i in range(n):
                pt_idx = (start_offset + i) % len(points)
                extra_assignments.append((rnd, pt_idx))
    # ===== 修复结束 =====

    new_records = []
    for rnd in range(1, 13):
        n = extra_per_round + (1 if rnd <= extra_remainder else 0)
        if n == 0:
            continue
        rnd_records = [r for r in records if r.get("round") == rnd]
        rnd_records.sort(key=lambda r: r.get("time", ""))
        if len(rnd_records) < 2:
            continue
        # 获取该轮时间范围

        # ===== Bug 修复 #2: extra 记录的时间应该晚于该点位该轮的 system 记录 =====
        # 原逻辑问题:extra 记录的时间基于该轮所有点位的时间缝隙,可能早于该点位的 system 记录
        # 例如:extra 记录分配给"消防通道内",但时间缝隙是基于"B 端办公区东大门"的时间计算的
        # 导致 extra 记录时间 (00:05:57) 早于"消防通道内"的第一条 system 记录 (00:06:11)
        # 修复方案:extra 记录的时间应该在该点位该轮的最后一条 system 记录之后
        for assigned_rnd, pt_idx in extra_assignments:
            if assigned_rnd != rnd:
                continue
            pt = points[pt_idx]
            # 找到该点位该轮的 system 记录
            pt_rnd_records = [r for r in rnd_records if r["point"] == pt]
            # P2 修复：点位在该轮无记录时，使用轮次最后一条记录的时间作为基准
            # MEDIUM-004 修复：添加时间解析异常处理
            if not pt_rnd_records:
                if not rnd_records:
                    continue  # 整轮都无记录，跳过
                # 使用轮次最后一条记录的时间作为基准
                try:
                    last_pt_time = datetime.datetime.strptime(rnd_records[-1]["time"], "%Y-%m-%d %H:%M:%S")
                except ValueError as e:
                    print(f"[WARNING] MEDIUM-004: 时间解析失败 (rnd_records): {rnd_records[-1]['time']}, error: {e}", flush=True)
                    last_pt_time = datetime.datetime.now()
            else:
                # extra 记录的时间应该在该点位该轮的最后一条 system 记录之后
                try:
                    last_pt_time = datetime.datetime.strptime(pt_rnd_records[-1]["time"], "%Y-%m-%d %H:%M:%S")
                except ValueError as e:
                    print(f"[WARNING] MEDIUM-004: 时间解析失败 (pt_rnd_records): {pt_rnd_records[-1]['time']}, error: {e}", flush=True)
                    last_pt_time = datetime.datetime.now()
            # 在该点位该轮的最后一条 system 记录之后插入
            # 在 last_pt_time 之后随机增加 5-30 秒
            offset_seconds = rng.randint(10, 40)
            current_time = last_pt_time + datetime.timedelta(seconds=offset_seconds)
            time_str = current_time.strftime("%Y-%m-%d %H:%M:%S")
            # 从该点位该轮的原记录中获取 device 值 (取第一条记录的 device)
            # [FIX 2026-07-01] device 字段继承修复:如果该点位无原记录,从同轮其他点位获取 device
            if pt_rnd_records:
                device_value = pt_rnd_records[0].get("device", "")
            else:
                # 该点位无原记录:从同轮其他点位继承 device
                all_devices_in_rnd = [r.get("device", "") for r in rnd_records if r.get("device", "")]
                device_value = all_devices_in_rnd[0] if all_devices_in_rnd else ""
            new_records.append({
                "seq": 0, "point": pt, "time": time_str, "device": device_value,
                "event": "合格", "inspector": "", "route": route_name,
                "round": rnd, "source": "extra"
            })

    if new_records:
        records = [r for r in records if r.get("source") != "extra"]

        # ===== 关键修复:将 records 按点位重新排列,使同一站点的所有记录连续 =====
        # 原始 records 按时间排序,同一点位的多次打点会被其他点位隔开,
        # 导致后续连续分组逻辑将同一点位拆成多个独立组。
        # 修复:按巡更路线顺序重组,同点位记录连续排列,组内按时间排序。
        grouped = {}
        for r in records:
            pt = r["point"]
            grouped.setdefault(pt, []).append(r)

        reordered = []
        # 按巡更路线(points)顺序输出各点位组
        for pt in points:
            if pt in grouped:
                pt_records = grouped.pop(pt)
                pt_records.sort(key=lambda r: r.get("time", ""))
                reordered.extend(pt_records)
        # 处理不在 points 列表中的点位(容错)
        for pt, pt_records in grouped.items():
            pt_records.sort(key=lambda r: r.get("time", ""))
            reordered.extend(pt_records)
        records = reordered
        # ===== 修复结束 =====

        # 按点位分组(现在同点位记录已连续,分组逻辑可正确工作)
        groups = []
        i = 0
        while i < len(records):
            pt = records[i]["point"]
            j = i
            while j < len(records) and records[j]["point"] == pt:
                j += 1
            groups.append((pt, i, j))
            i = j

        extra_by_point = {}
        for extra in new_records:
            pt = extra["point"]
            extra_by_point.setdefault(pt, []).append(extra)

        extra_remaining = []
        for pt, extras in extra_by_point.items():
            found = False
            for g_idx, (g_pt, g_start, g_end) in enumerate(groups):
                if g_pt == pt:
                    insert_idx = g_end
                    for extra in extras:
                        records.insert(insert_idx, extra)
                        insert_idx += 1
                    shift = len(extras)
                    groups[g_idx] = (g_pt, g_start, g_end + shift)
                    for gi in range(g_idx + 1, len(groups)):
                        gp, gs, ge = groups[gi]
                        groups[gi] = (gp, gs + shift, ge + shift)
                    found = True
                    break
            if not found:
                extra_remaining.extend(extras)

        if extra_remaining:
            records.extend(extra_remaining)

    return records

def apply_global_limit(all_records, points):
    """全局上限控制(只在全局末尾执行一次)- 保持原始顺序,只删除超出的记录"""
    max_total = len(points) * TAPS_PER_POINT * 12 + 30
    total_count = len(all_records)
    if total_count <= max_total:
        return all_records
    excess = total_count - max_total
    # 保持原始顺序,从末尾删除优先删除 extra
    result = list(all_records)
    deleted = 0
    # 从后往前遍历,优先删除 extra 记录
    i = len(result) - 1
    while i >= 0 and deleted < excess:
        if result[i].get("source") == "extra":
            result.pop(i)
            deleted += 1
        i -= 1
    # 如果还不够,删除 system
    i = len(result) - 1
    while i >= 0 and deleted < excess:
        if result[i].get("source") == "system":
            result.pop(i)
            deleted += 1
        i -= 1
    # 最后删除 user
    i = len(result) - 1
    while i >= 0 and deleted < excess:
        result.pop(i)
        deleted += 1
        i -= 1
    return result


def generate_patrol_data(date_str, num_records, route_name="线路一", rule="odd", gap_min=10, gap_max=40, tap=2):
    """
    ╔══════════════════════════════════════════════════════════════════╗
    ║ [生成模块] 生成巡更数据 - 仅对「生成」选项卡有效              ║
    ║ 与「一键修改」模块无关,一键修改使用 /api/auto-fix 接口         ║
    ╚══════════════════════════════════════════════════════════════════╝
    无论输入多少数量,都生成完整 12 轮数据
    使用新的 process_all_rounds 主流程处理
    按指定路线的点位顺序生成
    """
    # 使用路线点位而非所有点位
    points = get_route_points(route_name)
    records = []  # 生成模式无原数据
    records = process_all_rounds(records, points, date_str, route_name=route_name, random_empty_rounds=False)
    return records


# [REMOVED] save_records_to_db 函数已删除 - patrol_records 表已删除

def parse_uploaded_xls(file_bytes):
    """解析上传的 xls 文件，自动识别轮次"""
    import xlrd
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheet_by_index(0)
    records = []
    last_point = ""  # 用于处理合并单元格 (点位名为空时继承上一行)
    
    # 读取 G2 单元格的值（第 2 行，第 7 列=G 列）
    g2_value = ""
    try:
        g2_cell = ws.cell(1, 6)  # row=1 (0-indexed), column=6 (G 列)
        if g2_cell and g2_cell.value:
            g2_value = str(g2_cell.value).strip()
        print(f"[DEBUG] G2 单元格值：'{g2_value}'", flush=True)
    except Exception as e:
        print(f"[DEBUG] 读取 G2 单元格失败：{e}", flush=True)
        pass
    
    for row in range(12, ws.nrows):
        seq = ws.cell(row, 2).value
        point = ws.cell(row, 3).value
        ptime = ws.cell(row, 4).value
        device = ws.cell(row, 7).value
        event = ws.cell(row, 9).value
        inspector = ws.cell(row, 12).value
        # 处理合并单元格:点位名为空时继承上一行的值
        if not point:
            point = last_point
        if not point or not ptime:
            break
        last_point = str(point).strip()  # 更新 last_point 供下一行使用
        if isinstance(ptime, (int, float)):
            try:
                dt_tuple = xlrd.xldate_as_tuple(ptime, wb.datemode)
                ptime = datetime.datetime(*dt_tuple).strftime("%Y-%m-%d %H:%M:%S")
            except ValueError as e:
                # MEDIUM-005 修复：添加日志和默认日期值
                print(f"[WARNING] MEDIUM-005: Excel 日期转换失败，row={row}, value={ptime}, error: {e}", flush=True)
                ptime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            except Exception as e:
                # 捕获其他异常
                print(f"[WARNING] MEDIUM-005: Excel 日期转换未知错误，row={row}, value={ptime}, error: {e}", flush=True)
                ptime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ptime_str = str(ptime).strip()
        round_num = time_to_round(ptime_str)
        records.append({
            "seq": int(seq) if seq else len(records)+1,
            "point": str(point).strip(),
            "time": ptime_str,
            "device": str(device).strip() if device else "",
            "event": str(event).strip() if event else "合格",
            "inspector": str(inspector).strip() if inspector else "",
            "round": round_num,
            "route": g2_value,  # 使用 G2 单元格的值
            "_added": False
        })
    return records




class PatrolHandler(BaseHTTPRequestHandler):
    # LOW-006 注释：do_POST 方法超过 800 行，包含多个 API 端点处理
    # 结构说明：
    #   - 1896-1913: /api/upload 文件上传处理
    #   - 1941-1972: /api/login 登录处理
    #   - 1975-2001: /api/auth/activate 设备激活 (含 rate limiting)
    #   - 2004-2052: /api/admin/reset-password 密码重置 (含 rate limiting 和审计日志)
    #   - 2059-2071: /api/points 点位管理
    #   - 2073-2194: /api/routes/* 路线管理
    #   - 2196-2214: /api/points/sort-by-label 标签排序
    #   - 2216-2223: /api/records/save 记录保存 (已移除)
    #   - 2224-2320: /api/auto-fix 自动修复
    #   - 2322-2329: /api/save-modified 保存修改 (已移除)
    #   - 2331-2502: /api/export-prepare Excel 导出
    #   - 2504-2757: /api/export-pdf PDF 导出
    #   - 2759-2777: /api/points/toggle 点位启用/禁用
    #   - 2769-2777: /api/points/delete 点位删除
    def log_message(self, format, *args):
        """记录 HTTP 请求日志到文件和控制台"""
        # 格式化日志消息
        msg = format % args
        # 记录到错误日志文件
        log_error(f"[HTTP] {self.client_address[0]} - {msg}")
        # 同时打印到控制台（便于调试）
        print(f"[HTTP] {self.client_address[0]} - {msg}", flush=True)

    def _send_json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, html, code=200):
        body = html.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, filename, code=200):
        """发送静态HTML文件（支持 PyInstaller 打包路径）"""
        filepath = get_resource_path(filename)
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                body = f.read().encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
            self.end_headers()
            self.wfile.write(body)
        except FileNotFoundError:
            self.send_error(404, f"文件 {filename} 不存在")

    def _handle_upload(self):
        """处理 multipart/form-data 文件上传"""
        try:
            content_type = self.headers.get('Content-Type', '')
            content_len = int(self.headers.get('Content-Length', 0))

            # 解析 boundary
            if 'boundary=' not in content_type:
                self._send_json({"error": "请求格式错误:缺少 boundary 参数"}); return
            boundary = content_type.split('boundary=')[1].strip().encode()

            raw = self.rfile.read(content_len)
            parts = raw.split(b'--' + boundary)
            file_bytes = None

            for part in parts:
                if b'Content-Disposition' in part and b'filename' in part:
                    # 找到头部和内容的分隔符
                    header_end = part.find(b'\r\n\r\n')
                    if header_end > 0:
                        file_bytes = part[header_end+4:]
                        # 移除末尾的边界标记和 CRLF
                        if file_bytes.endswith(b'\r\n'):
                            file_bytes = file_bytes[:-2]
                        if file_bytes.endswith(b'--\r\n'):
                            file_bytes = file_bytes[:-4]
                        elif file_bytes.endswith(b'--'):
                            file_bytes = file_bytes[:-2]
                        break

            if not file_bytes:
                self._send_json({"error": "未找到文件数据"}); return

            # 调试日志
            print(f"[UPLOAD] 文件大小:{len(file_bytes)} 字节", flush=True)

            records = parse_uploaded_xls(file_bytes)
            if not records:
                self._send_json({"error": "文件中未找到有效记录,请确认文件格式正确"}); return

            print(f"[UPLOAD] 解析成功:{len(records)} 条记录", flush=True)
            self._send_json({"records": records, "count": len(records)})
        except IndexError as e:
            print(f"[UPLOAD] boundary 解析错误:{e}", flush=True)
            self._send_json({"error": "请求格式错误:无法解析 multipart 数据"})
        except Exception as e:
            print(f"[UPLOAD] 上传错误:{e}", flush=True)
            self._send_json({"error": f"文件解析失败:{str(e)}"})


    def _serve_static_file(self, filename, content_type):
        """提供静态文件服务（支持 PyInstaller 打包路径）"""
        filepath = get_resource_path(filename)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                html_content = f.read()
            self._send_html(html_content)
        except FileNotFoundError:
            self._send_json({"error": f"文件未找到:{filename}"}, 404)
        except Exception as e:
            self._send_json({"error": f"文件读取失败:{str(e)}"}, 500)

    # LOW-007 修复：提取鉴权逻辑为辅助方法
    def _check_auth(self):
        """验证 session token，返回是否已认证"""
        from auth import verify_session_token, is_auth_enabled
        if not is_auth_enabled():
            return True
        token = self.headers.get('X-Auth-Token', '')
        if not token:
            return False
        db = get_db()
        try:
            return verify_session_token(token, db)
        finally:
            db.close()

    def do_GET(self):
        parsed = urlparse(self.path)
        print(f"[GET] {parsed.path} from {self.client_address}", flush=True)
        path = parsed.path
        params = parse_qs(parsed.query)

        # 静态页面和登录接口无需鉴权
        if path in ("/", "/index.html"):
            self._serve_static_file("index.html", "text/html; charset=utf-8")
            return
        if path == "/api/login":
            from auth import is_auth_enabled
            self._send_json({"enabled": is_auth_enabled()})
            return
        if path.startswith("/api/export-download/"):
            # 导出下载无需额外鉴权 (token 在 URL 路径中)
            token = parsed.path.split("/")[-1].split(".")[0]
            with _export_store_lock:
                if token not in _export_store:
                    self.send_error(404, "File not found or expired")
                    return
                body_bytes, filename, mime = _export_store.pop(token)
            self.send_response(200)
            self.send_header('Content-Type', mime)
            # 中文文件名编码 (RFC 5987)
            import urllib.parse
            encoded_filename = urllib.parse.quote(filename, encoding='utf-8')
            self.send_header('Content-Disposition', f'attachment; filename*=UTF-8\'\'{encoded_filename}')
            self.send_header('Content-Length', str(len(body_bytes)))
            self.end_headers()
            self.wfile.write(body_bytes)
            return

        # ===== 公开 API(无需鉴权)=====
        if path == "/api/round-ranges":
            # 返回轮次时间范围定义,供前端动态引用
            round_ranges = {str(k): {"start": v[0], "end": v[1], "label": f"第{k}轮 ({v[0]}:00-{v[1]}:00)"} for k, v in ROUND_TIME_RANGES.items()}
            self._send_json({"round_ranges": round_ranges, "total_rounds": len(ROUND_TIME_RANGES)})
            return

        if path == "/api/auth/check-device":
            # HIGH-004 修复：添加 rate limiting（10 次/分钟）
            import time as _time
            if not hasattr(self.__class__, '_check_device_rate_limit'):
                self.__class__._check_device_rate_limit = {}  # {ip: [timestamps]}
            client_ip = self.client_address[0]
            now = _time.time()
            if client_ip not in self.__class__._check_device_rate_limit:
                self.__class__._check_device_rate_limit[client_ip] = []
            # 清理 60 秒前的记录
            self.__class__._check_device_rate_limit[client_ip] = [
                t for t in self.__class__._check_device_rate_limit[client_ip] if now - t < 60
            ]
            if len(self.__class__._check_device_rate_limit[client_ip]) >= 10:
                self._send_json({"error": "请求过于频繁，请稍后再试"}, 429)
                return
            self.__class__._check_device_rate_limit[client_ip].append(now)
            
            # v2.4 简化：返回设备指纹和授权状态（仅指纹，无 MAC）
            from auth import get_device_fingerprint
            
            try:
                fingerprint = get_device_fingerprint()
            except Exception as e:
                print(f"[WARN] 获取设备指纹失败：{e}")
                fingerprint = None
            
            authorized = is_device_authorized()
            
            self._send_json({
                "authorized": authorized,
                "fingerprint": fingerprint
            })
            return

        # ===== 获取 2FA 信息 API（无需登录，用于显示密钥和二维码）=====
        if path == "/api/auth/2fa-info":
            import pyotp
            
            # 从数据库读取 2FA 密钥
            secret = get_system_config("2fa_secret_key")
            if not secret:
                # 首次使用时生成并保存 2FA 密钥（Base32 编码，20 字节=160 位）
                secret = pyotp.random_base32()
                db = get_db()
                try:
                    db.execute(
                        "INSERT OR REPLACE INTO system_config (key, value, description) VALUES (?, ?, ?)",
                        ("2fa_secret_key", secret, "2FA TOTP 密钥 (Base32 编码，32 字符)")
                    )
                    db.commit()
                finally:
                    db.close()
            
            # 生成 otpauth:// URI
            # 格式：otpauth://totp/Issuer:AccountName?secret=SECRET&issuer=Issuer&algorithm=SHA1&digits=6&period=30
            issuer = "PatrolSystem"
            account_name = "admin"
            otpauth_uri = f"otpauth://totp/{issuer}:{account_name}?secret={secret}&issuer={issuer}&algorithm=SHA1&digits=6&period=30"
            
            self._send_json({
                "success": True,
                "secret_key": secret,
                "otpauth_uri": otpauth_uri
            })
            return

        # ===== 其他 GET API 需要鉴权 =====
        # LOW-007: 使用提取的 _check_auth 方法，避免重复鉴权逻辑
        if not self._check_auth():
            self._send_json({"error": "请先登录"}, 401)
            return

        if path == "/api/overview":
            self._send_json(db_get_stats())
        elif path == "/api/points":
            self._send_json(db_get_points())
        elif path == "/api/routes":
            self._send_json({"routes": get_all_routes()})
        elif path == "/api/routes/detail":
            route_name = params.get("name", [None])[0]
            if not route_name:
                self._send_json({"error": "需要 name 参数"}); return
            detail = get_route_detail(route_name)
            if detail:
                self._send_json(detail)
            else:
                self._send_json({"error": "路线不存在"}, 404)
        elif path == "/api/records":
            # [REMOVED] patrol_records 表已删除 - 不再存储巡更记录
            # 返回空记录列表
            q_route = params.get("route", ["线路一"])[0]
            q_source = params.get("source_type", ["actual"])[0]
            self._send_json({"records": [], "count": 0, "used_route": q_route, "source_type": q_source})
        elif path == "/api/generate":
            date_str = params.get("date", [None])[0]
            count = int(params.get("count", ["480"])[0])
            route = params.get("route", ["线路一"])[0]
            rule = params.get("rule", ["odd"])[0]
            tap = int(params.get("tap", ["2"])[0])
            gap_min = int(params.get("gap_min", ["5"])[0])
            gap_max = int(params.get("gap_max", ["8"])[0])
            preview = params.get("preview", ["0"])[0] == "1"
            save = params.get("save", ["0"])[0] == "1"
            if not date_str:
                self._send_json({"error": "需要date参数"}); return
            # 随机路线选择
            if route == "random":
                all_routes = get_all_routes()
                if all_routes:
                    route = random.choice(all_routes)["name"]
                    print(f"[随机路线] 选择:{route}", flush=True)
                else:
                    route = "线路一"  # fallback to default if no routes

            records = generate_patrol_data(date_str, count, route, rule, gap_min, gap_max, tap)
            # [REMOVED] save_records_to_db 调用已删除 - patrol_records 表已删除
            if save:
                print("[INFO] save 参数被忽略 - patrol_records 表已删除，不再存储巡更记录", flush=True)
            self._send_json({"records": records, "count": len(records), "used_route": route})
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        print(f"[POST] {parsed.path} from {self.client_address}", flush=True)
        
        # 导入鉴权函数（用于密码验证）
        from auth import check_admin_op_password
        
        # /api/upload 使用 multipart/form-data,不走 JSON 解析
        if parsed.path == "/api/upload":
            # 验证 session token
            from auth import verify_session_token, is_auth_enabled
            # 鉴权未启用时，允许所有访问
            if is_auth_enabled():
                token = self.headers.get('X-Auth-Token', '')
                if not token:
                    self._send_json({"error": "请先登录"}, 401)
                    return
                db = get_db()
                try:
                    if not verify_session_token(token, db):
                        self._send_json({"error": "请先登录"}, 401)
                        return
                finally:
                    db.close()
            self._handle_upload()
            return

        try:
            content_len = int(self.headers.get('Content-Length', 0))
            raw_body = self.rfile.read(content_len) if content_len > 0 else b''
            body = json.loads(raw_body) if raw_body else {}
        except Exception as e:
            self._send_json({"error": f"请求数据解析失败: {e}"}, 400)
            return

        # ===== 登录/鉴权检查 =====
        def require_auth():
            """验证 session token，未登录返回 False
            当鉴权未启用时 (auth.json 不存在)，直接返回 True 允许访问"""
            from auth import verify_session_token, is_auth_enabled
            # 鉴权未启用时，允许所有访问
            if not is_auth_enabled():
                return True
            token = self.headers.get('X-Auth-Token', '')
            if not token:
                return False
            db = get_db()
            try:
                return verify_session_token(token, db)
            finally:
                db.close()

        # ===== 登录/鉴权端点 (无需验证)=====
        if parsed.path == "/api/login":
            from auth import check_password, create_session, is_auth_enabled, get_device_fingerprint
            # GET 请求：返回鉴权是否启用
            if self.command == "GET":
                self._send_json({"enabled": is_auth_enabled()})
                return
            # POST 请求：验证密码 + 设备授权
            password = body.get("password", "")
            if not password:
                self._send_json({"enabled": is_auth_enabled(), "error": "请输入密码"}, 401)
                return
            # 检查设备授权（仅设备指纹）
            try:
                fingerprint = get_device_fingerprint()
            except Exception:
                fingerprint = None
            if fingerprint and not is_device_authorized():
                self._send_json({"ok": False, "error": "设备未授权，请联系管理员激活", "code": "DEVICE_NOT_AUTHORIZED"}, 403)
                return
            # 设备授权通过，验证密码
            if check_password(password):
                db = get_db()
                try:
                    # 使用设备指纹作为 user_id
                    user_id = fingerprint if fingerprint else "admin"
                    token = create_session(user_id, db)
                finally:
                    db.close()
                self._send_json({"ok": True, "token": token})
            else:
                self._send_json({"ok": False, "error": "密码错误"}, 401)
            return

        # ===== MAC 激活 API（无需登录，但有 rate limiting）=====
        if parsed.path == "/api/auth/activate":
            # Rate limiting: 最多 5 次/分钟（防止 TOTP 暴力破解）
            import time as _time
            if not hasattr(self.__class__, '_activate_rate_limit'):
                self.__class__._activate_rate_limit = {}  # {ip: [timestamps]}
            client_ip = self.client_address[0]
            now = _time.time()
            if client_ip not in self.__class__._activate_rate_limit:
                self.__class__._activate_rate_limit[client_ip] = []
            # 清理 60 秒前的记录
            self.__class__._activate_rate_limit[client_ip] = [
                t for t in self.__class__._activate_rate_limit[client_ip] if now - t < 60
            ]
            # MEDIUM-009 修复：当 IP 记录列表为空时删除该 IP 的键，防止内存泄漏
            if not self.__class__._activate_rate_limit[client_ip]:
                del self.__class__._activate_rate_limit[client_ip]
            if client_ip not in self.__class__._activate_rate_limit:
                self.__class__._activate_rate_limit[client_ip] = []
            if len(self.__class__._activate_rate_limit[client_ip]) >= 5:
                self._send_json({"success": False, "activated_at": None, "error": "请求过于频繁，请稍后再试"}, 429)
                return
            self.__class__._activate_rate_limit[client_ip].append(now)

            # 兼容前端参数名：code / otp_code
            otp_code = body.get("otp_code") or body.get("code", "")
            if not otp_code:
                self._send_json({"success": False, "activated_at": None, "error": "缺少授权密钥参数"}, 400)
                return
            result = activate_device(otp_code)
            code = 200 if result.get("success") else 403
            self._send_json(result, code)
            return

        # ===== 重置密码 API（无需登录，但有 rate limiting 和密钥验证）=====
        if parsed.path == "/api/admin/reset-password":
            # Rate limiting: 最多 5 次/分钟
            import time as _time
            if not hasattr(self.__class__, '_reset_rate_limit'):
                self.__class__._reset_rate_limit = {}  # {ip: [timestamps]}
            client_ip = self.client_address[0]
            now = _time.time()
            if client_ip not in self.__class__._reset_rate_limit:
                self.__class__._reset_rate_limit[client_ip] = []
            # 清理 60 秒前的记录
            self.__class__._reset_rate_limit[client_ip] = [
                t for t in self.__class__._reset_rate_limit[client_ip] if now - t < 60
            ]
            # MEDIUM-009 修复：当 IP 记录列表为空时删除该 IP 的键，防止内存泄漏
            if not self.__class__._reset_rate_limit[client_ip]:
                del self.__class__._reset_rate_limit[client_ip]
            if client_ip not in self.__class__._reset_rate_limit:
                self.__class__._reset_rate_limit[client_ip] = []
            if len(self.__class__._reset_rate_limit[client_ip]) >= 5:
                self._send_json({"error": "请求过于频繁，请稍后再试"}, 429)
                return
            self.__class__._reset_rate_limit[client_ip].append(now)

            reset_key = body.get("reset_key", "")
            new_password = body.get("new_password", "")
            if not reset_key or not new_password:
                self._send_json({"error": "请提供重置密钥和新密码"}, 400)
                return
            if len(new_password) < 3:
                self._send_json({"error": "新密码长度至少 3 位"}, 400)
                return
            # HIGH-004 修复：从数据库获取重置密钥的 bcrypt 哈希
            from auth import get_system_config
            import bcrypt
            stored_reset_key_hash = get_system_config("password_reset_key_hash")
            if not stored_reset_key_hash:
                self._send_json({"error": "重置密钥未配置，请联系管理员"}, 500)
                return
            # 使用 bcrypt 验证重置密钥
            try:
                if not bcrypt.checkpw(reset_key.encode('utf-8'), stored_reset_key_hash.encode('utf-8')):
                    self._send_json({"error": "重置密钥错误"}, 403)
                    return
            except Exception:
                self._send_json({"error": "重置密钥验证失败"}, 500)
                return
            # 验证通过，更新密码
            new_hash = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            db = get_db()
            db.execute("UPDATE auth_users SET password_hash = ? WHERE role = 'admin' AND active = 1", (new_hash,))
            db.commit()
            db.close()
            # HIGH-003 修复：添加审计日志
            print(f"[AUDIT] {datetime.now()} - 密码重置 by IP:{client_ip}")
            self._send_json({"ok": True, "message": "密码重置成功"})
            return

        # ===== 其他API需要登录 =====
        if not require_auth():
            self._send_json({"error": "请先登录"}, 401)
            return

        if parsed.path == "/api/points":
            name = body.get("name", "").strip()
            pwd = body.get("pwd", "")
            if not check_admin_op_password(pwd):
                self._send_json({"error": "密码错误"}, 403); return
            if not name: self._send_json({"error": "名称不能为空"}); return
            db = get_db()
            try:
                mx = db.execute("SELECT MAX(sort_order) FROM patrol_points").fetchone()[0] or 0
                db.execute("INSERT INTO patrol_points (name, sort_order) VALUES (?, ?)", [name, mx+1])
                db.commit(); db.close()
                self._send_json({"ok": True})
            except: db.close(); self._send_json({"error": "点位已存在"})

        elif parsed.path == "/api/routes":
            # GET: 获取所有路线
            routes = get_all_routes()
            self._send_json({"routes": routes})

        elif parsed.path == "/api/routes/save":
            # POST: 保存路线配置
            route_name = body.get("name", "").strip()
            description = body.get("description", "")
            point_ids = body.get("point_ids", [])  # 点位 ID 列表,按顺序
            pwd = body.get("pwd", "")
            if not check_admin_op_password(pwd):
                self._send_json({"error": "密码错误"}, 403); return
            if not route_name:
                self._send_json({"error": "路线名称不能为空"}); return
            # 如果没有传 point_ids,自动包含所有点位(新增路线场景)
            print(f"[DEBUG] point_ids before: {point_ids}", flush=True)
            if not point_ids:
                db = get_db()
                all_points = db.execute("SELECT id FROM patrol_points ORDER BY sort_order").fetchall()
                db.close()
                point_ids = [p[0] for p in all_points]
                print(f"[DEBUG] Loaded {len(point_ids)} points", flush=True)
                if not point_ids:
                    self._send_json({"error": "系统中没有可用点位"}); return
            db = get_db()
            try:
                # 插入或更新路线
                db.execute("INSERT OR REPLACE INTO patrol_routes (name, description) VALUES (?, ?)", (route_name, description))
                route_id = db.execute("SELECT id FROM patrol_routes WHERE name=?", (route_name,)).fetchone()[0]
                # 删除旧的点位关联
                db.execute("DELETE FROM patrol_route_points WHERE route_id=?", (route_id,))
                # 插入新的点位关联
                for idx, point_id in enumerate(point_ids):
                    db.execute("INSERT INTO patrol_route_points (route_id, point_id, sort_order) VALUES (?, ?, ?)",
                               (route_id, point_id, idx))
                db.commit()
                db.close()
                self._send_json({"ok": True, "route_id": route_id})
            except Exception as e:
                db.close()
                self._send_json({"error": f"保存失败:{e}"})

        elif parsed.path == "/api/routes/update":
            # POST: 更新路线 (支持修改名称和描述)
            old_name = body.get("old_name", "").strip()
            new_name = body.get("new_name", "").strip()
            description = body.get("description", "").strip()
            pwd = body.get("pwd", "")
            if not check_admin_op_password(pwd):
                self._send_json({"error": "密码错误"}, 403); return
            if not old_name:
                self._send_json({"error": "原路线名称不能为空"}); return
            if not new_name:
                self._send_json({"error": "新路线名称不能为空"}); return
            # 检查新名称是否与其他路线重名
            if new_name != old_name:
                db = get_db()
                dup = db.execute("SELECT id FROM patrol_routes WHERE name=? AND name!=?", (new_name, old_name)).fetchone()
                if dup:
                    self._send_json({"error": "路线名称已存在"}); return
            db = get_db()
            try:
                # 检查路线是否存在
                route = db.execute("SELECT id FROM patrol_routes WHERE name=?", (old_name,)).fetchone()
                if not route:
                    self._send_json({"error": "路线不存在"}); return
                db.execute("UPDATE patrol_routes SET name=?, description=? WHERE name=?", (new_name, description, old_name))
                db.commit()
                self._send_json({"ok": True})
            except Exception as e:
                self._send_json({"error": f"更新失败:{e}"})

        elif parsed.path == "/api/routes/delete":
            # POST: 删除路线
            route_name = body.get("name", "")
            pwd = body.get("pwd", "")
            if not check_admin_op_password(pwd):
                self._send_json({"error": "密码错误"}, 403); return
            if not route_name:
                self._send_json({"error": "路线名称不能为空"}); return
            if route_name == "线路一":
                self._send_json({"error": "默认路线不能删除"}); return
            db = get_db()
            try:
                route_id = db.execute("SELECT id FROM patrol_routes WHERE name=?", (route_name,)).fetchone()
                if route_id:
                    route_id = route_id[0]
                    db.execute("DELETE FROM patrol_route_points WHERE route_id=?", (route_id,))
                    db.execute("DELETE FROM patrol_routes WHERE id=?", (route_id,))
                    db.commit()
                db.close()
                self._send_json({"ok": True})
            except Exception as e:
                db.close()
                self._send_json({"error": f"删除失败:{e}"})

        elif parsed.path == "/api/routes/save-order":
            # POST: 保存点位顺序
            route_id = body.get("route_id")
            point_ids = body.get("point_ids", [])
            pwd = body.get("pwd", "")
            if not check_admin_op_password(pwd):
                self._send_json({"error": "密码错误"}, 403); return
            if not route_id:
                self._send_json({"error": "需要 route_id"}); return
            if not point_ids:
                self._send_json({"error": "点位列表不能为空"}); return
            db = get_db()
            try:
                # MEDIUM-006 修复：使用 BEGIN 开始事务，确保删除和插入操作原子性
                db.execute("BEGIN")
                # 删除旧的点位关联
                db.execute("DELETE FROM patrol_route_points WHERE route_id=?", (route_id,))
                # 插入新的点位关联 (按提供的顺序)
                for idx, point_id in enumerate(point_ids):
                    db.execute("INSERT INTO patrol_route_points (route_id, point_id, sort_order) VALUES (?, ?, ?)",
                               (route_id, point_id, idx))
                db.commit()
                db.close()
                self._send_json({"ok": True})
            except Exception as e:
                db.close()
                self._send_json({"error": f"保存失败:{e}"})

        elif parsed.path == "/api/points/sort-by-label":
            # POST: 手动触发按标签排序所有点位
            # [FIX BUG-001] 统一密码验证通道:支持 body.password / body.pwd / Authorization Bearer
            # 与 PUT/DELETE /api/points/:id/label 的密码验证逻辑保持一致
            auth_password = body.get("password") or body.get("pwd") or self.headers.get('Authorization', '').replace('Bearer ', '')
            if not check_admin_op_password(auth_password):
                self._send_json({"error": "密码错误"}, 403)
                return

            try:
                result = recalculate_sort_order_by_label()
                self._send_json({
                    "success": True,
                    "message": f"排序完成: {result['sorted']} 个点位已重新排序",
                    "details": result
                })
            except Exception as e:
                self._send_json({"error": f"排序失败: {e}"}, 500)
            return

        elif parsed.path == "/api/records/save":
            # [REMOVED] patrol_records 表已删除 - 不再存储巡更记录
            q_date = body.get("date", "")
            q_route = body.get("route", "线路一")
            records = body.get("records", [])
            print(f"[INFO] /api/records/save 请求被忽略 - patrol_records 表已删除 (date={q_date}, route={q_route}, count={len(records)})", flush=True)
            self._send_json({"ok": True, "count": 0, "message": "巡更记录不再存储到数据库"})

        elif parsed.path == "/api/auto-fix":
            records = body.get("records", [])
            if not records: self._send_json({"error": "无记录"}); return
            date_str = body.get("date", records[0].get("time", "")[:10] if records else "")

            # [MODIFY 2026-06-22] 步骤0:规范化点位名称,统一去除首尾空白字符
            # Bug修复:同点位记录被分开的根因
            # 前端传入的原始数据点位名称可能含有尾随空格(如'票亭 '),而补充算法生成的点位名称来自
            # route_points(已去空格)。process_all_rounds 的最终排序 key 使用 r.get('point','')原始值,
            # 导致'票亭 '和'票亭'被视为不同字符串,排在不同位置(第3条和第36条)。
            # 修复方案:在入口处统一规范化所有点位名称,确保所有下游处理使用一致的点位字符串。
            for r in records:
                if "point" in r:
                    r["point"] = ''.join(r["point"].split())

            # [MODIFY] 步骤1:根据原数据点位集合识别所属路线
            route_name, route_points = identify_route_from_records(records)
            # [FIX 2026-07-01] 修复点位数量硬编码问题
            # 无论上传数据包含多少个点位,补充数据时应使用数据库中所有活跃点位,
            # 而不是根据上传数据匹配到的路线点位数量(可能少于活跃点位总数)
            all_active_points = get_active_points()
            if all_active_points:
                route_points = all_active_points
                # 如果之前识别的路线名称为空或未识别,使用默认路线
                if not route_name:
                    route_name = "线路一"
            if not route_points:
                self._send_json({"error": "无活跃巡更点"}); return

            # [MODIFY] 步骤 2:给每条记录加 round 字段和 source 标记
            # 从原记录中提取第一个非空的 device 值作为默认值 (用于继承)
            default_device = ""
            for r in records:
                dev = r.get("device", "")
                if dev:
                    default_device = dev
                    break
            for r in records:
                if "round" not in r or not r["round"]:
                    r["round"] = time_to_round(r.get("time", ""))
                if "source" not in r:
                    r["source"] = "user"  # 原数据标记
                # 如果记录没有 route 字段或为空，则使用识别出的路线名称
                if "route" not in r or not r["route"]:
                    r["route"] = route_name
                # 如果记录没有 device 字段或为空，则使用默认 device 值 (继承)
                if "device" not in r or not r["device"]:
                    r["device"] = default_device

            # [MODIFY] 步骤 3:按照路线点位顺序验证原数据，检测异常
            original_records = list(records)
            order_validation = validate_route_order(original_records, route_points)

            # [MODIFY] 步骤 4:使用路线点位作为标准顺序，调用主流程进行数据补充
            original_count = len(records)
            records = process_all_rounds(records, route_points, date_str, route_name=route_name, default_device=default_device)

            # 生成日志(包含路线识别信息和异常信息)
            fix_log = [f'识别路线: {route_name} ({len(route_points)}个点位)']
            if order_validation["anomalies"]:
                fix_log.append(f'异常数: {len(order_validation["anomalies"])}条')
                for anom in order_validation["anomalies"][:5]:  # 最多显示前5条
                    fix_log.append(f'  - [{anom["type"]}] {anom["detail"]}')
                if len(order_validation["anomalies"]) > 5:
                    fix_log.append(f'  ... 还有{len(order_validation["anomalies"])-5}条异常')
            else:
                fix_log.append('点位顺序验证: 正常')
            # 统计整轮缺失的随机路线分配
            empty_rounds_info = []
            for rnd in range(1, 13):
                rnd_records = [r for r in records if r.get("round") == rnd]
                rnd_original = [r for r in original_records if time_to_round(r.get("time","")) == rnd]
                added = len(rnd_records) - len(rnd_original)
                if len(rnd_original) == 0 and len(rnd_records) > 0:
                    # 整轮缺失:获取该轮使用的路线
                    route_used = route_name
                    for r in rnd_records:
                        if r.get("route"):
                            route_used = r["route"]
                            break
                    empty_rounds_info.append(f'第{rnd}轮:随机路线"{route_used}",补充{added}条')
                fix_log.append(f'第{rnd}轮:{len(rnd_records)}条(原{len(rnd_original)}条+补{added})')
            if empty_rounds_info:
                fix_log.append(f'整轮缺失随机路线分配:')
                for info in empty_rounds_info:
                    fix_log.append(f'  - {info}')

            self._send_json({
                "ok": True,
                "records": records,
                "log": fix_log,
                "added": len(records) - original_count,
                "route_name": route_name,
                "route_points": route_points,
                "anomalies": order_validation["anomalies"],
                "round_orders": order_validation["round_orders"]
            })

        elif parsed.path == "/api/save-modified":
            # [REMOVED] patrol_records 表已删除 - 不再存储巡更记录
            records = body.get("records", [])
            date_str = body.get("date", "")
            if not records:
                self._send_json({"error": "没有记录可保存"}); return
            print(f"[INFO] /api/save-modified 请求被忽略 - patrol_records 表已删除 (date={date_str}, count={len(records)})", flush=True)
            self._send_json({"ok": True, "count": 0, "message": "巡更记录不再存储到数据库"})

        elif parsed.path == "/api/export-prepare":
            # 准备导出，返回下载 token
            records = body.get("records", [])
            route_name = body.get("route_name", "")  # 从原数据 G2 单元格继承
            original_filename = body.get("original_filename", "")  # 从原数据文件名继承
            fmt = body.get("format", "xls")
            if not records:
                self._send_json({"error": "没有数据可导出"})
                return

            # 生成 token (LOW-012: 使用 secrets 模块生成安全随机 token)
            token = f"export_{int(time.time()*1000)}_{secrets.token_urlsafe(16)}"
            # 优先使用原文件名，否则使用默认名称
            if original_filename:
                filename = f"{original_filename}.{fmt}"
            else:
                filename = f"patrol_records.{fmt}"
            mime = "application/vnd.ms-excel"

            try:
                # 生成文件内容(复用export-preview的逻辑)
                if fmt == 'xls':
                    # 从头创建XLS(不用模板),以正确支持网格线控制
                    import xlwt, io as _io
                    _wb = xlwt.Workbook()
                    _wb.set_colour_RGB(9, 0xD3, 0xD3, 0xD3)  # 索引9覆盖为灰色边框(211,211,211)
                    _ws = _wb.add_sheet('Rpt_Records_ens')
                    _ws.show_grid = False  # 关闭网格线(与金模板一致)
                    # 冻结窗格:第10列(J列)固定,K列起可滚动
                    _ws.panes_frozen = True
                    _ws.horz_split_pos = 10
                    _ws.horz_split_first_visible = 10
                    _ws.split_active_pane = 2

                    # === 样式定义 ===
                    # 金模板方案:关网格线 + 数据区显式边框 + 非数据区无边框无填充

                    # G2样式:20号微软雅黑加粗,居中顶端
                    g2_style = xlwt.XFStyle()
                    g2_style.font.name = '微软雅黑'; g2_style.font.height = 400; g2_style.font.bold = True
                    g2_style.alignment.horz = 2; g2_style.alignment.vert = 0

                    # B6样式:14号微软雅黑加粗,左上角
                    b6_style = xlwt.XFStyle()
                    b6_style.font.name = '微软雅黑'; b6_style.font.height = 280; b6_style.font.bold = True
                    b6_style.alignment.horz = 0; b6_style.alignment.vert = 0

                    # K7样式:微软雅黑10号,左上角,locale日期格式
                    k7_style = xlwt.XFStyle()
                    k7_style.font.name = '微软雅黑'; k7_style.font.height = 200
                    k7_style.alignment.horz = 0; k7_style.alignment.vert = 0
                    k7_style.num_format_str = '[$-010804]yyyy/m/d h:mm:ss'

                    # L3编号样式:微软雅黑10号,左上角
                    l3_style = xlwt.XFStyle()
                    l3_style.font.name = '微软雅黑'; l3_style.font.height = 200
                    l3_style.alignment.horz = 0; l3_style.alignment.vert = 0

                    # A10样式:Arial 10号(B8日期/时间,family=2 Swiss, charset=0 ANSI)
                    a10 = xlwt.XFStyle()
                    a10.font.name = 'Arial'; a10.font.height = 200
                    a10.font.family = 2; a10.font.charset = 0; a10.font.colour_index = 8
                    a10.alignment.horz = 0; a10.alignment.vert = 0

                    # 数据区边框(thin black)
                    cell_border = xlwt.Borders()
                    cell_border.left = xlwt.Borders.THIN
                    cell_border.right = xlwt.Borders.THIN
                    cell_border.top = xlwt.Borders.THIN
                    cell_border.bottom = xlwt.Borders.THIN
                    cell_border.left_colour = 9
                    cell_border.right_colour = 9
                    cell_border.top_colour = 9
                    cell_border.bottom_colour = 9

                    # 表头样式:Arial 10号,居中顶端,带边框
                    hdr_style = xlwt.XFStyle()
                    hdr_style.font.name = 'Arial'; hdr_style.font.height = 200
                    hdr_style.alignment.horz = 2; hdr_style.alignment.vert = 0; hdr_style.alignment.wrap = 1
                    hdr_style.borders = cell_border

                    # 数据行样式:微软雅黑10号,居中顶端,带灰色边框
                    data_border = xlwt.Borders()
                    data_border.left = xlwt.Borders.THIN; data_border.left_colour = 9
                    data_border.right = xlwt.Borders.THIN; data_border.right_colour = 9
                    data_border.top = xlwt.Borders.THIN; data_border.top_colour = 9
                    data_border.bottom = xlwt.Borders.THIN; data_border.bottom_colour = 9
                    data_style = xlwt.XFStyle()
                    data_style.font.name = '微软雅黑'; data_style.font.height = 200
                    data_style.alignment.horz = 2; data_style.alignment.vert = 0; data_style.alignment.wrap = 1
                    data_style.borders = data_border

                    # Row 10:空行,但A10:N10有黑色顶部边框(视觉上像下边框)
                    row10_border = xlwt.Borders()
                    row10_border.top = xlwt.Borders.THIN; row10_border.top_colour = 8  # RGB(0,0,0)黑色
                    row10_style = xlwt.XFStyle()
                    row10_style.borders = row10_border
                    for ci in range(14):
                        _ws.write(10, ci, '', row10_style)

                    # === 设置列宽(与XLSX一致,xlwt单位=1/256字符宽)===
                    col_widths = [548, 292, 1133, 4571, 877, 1170, 3876, 2560, 841, 512, 2048, 475, 3108, 548]
                    for ci, w in enumerate(col_widths):
                        _ws.col(ci).width = w

                    # === 设置行高(匹配原始模板)===
                    row_heights = [197, 180, 285, 107, 23, 360, 90, 340, 20, 8, 235, 340]
                    # 金模板的height_mismatch标志(1=该行需要,0=不需要)
                    height_mismatch_flags = [1, 0, 1, 0, 1, 0, 0, 0, 0, 1, 1, 0]
                    for ri, h in enumerate(row_heights):
                        _ws.row(ri).height = h
                        if ri < len(height_mismatch_flags) and height_mismatch_flags[ri]:
                            _ws.row(ri).height_mismatch = True

                    # === 按照原始模板格式生成 ===
                    num_data_rows = len(records)
                    data_row_end_0 = 12 + num_data_rows
                    max_cols = 14

                    # === 非数据区：直接写入内容，无需白色填充 ===
                    # G2 单元格：显示线路名称（从请求参数继承）
                    g2_value = route_name if route_name else '线路一'
                    _ws.write_merge(1, 3, 6, 7, g2_value, g2_style)  # G2:H4
                    _ws.write_merge(5, 6, 1, 4, '记录查询', b6_style)  # B6:E7
                    if records:
                        ft = records[0].get('time', ''); lt = records[-1].get('time', '')
                        ds = ft.split(' ')[0]; de = lt.split(' ')[0]
                        ldt = datetime.datetime.strptime(lt, '%Y-%m-%d %H:%M:%S')
                        if end_date:
                            # 前端传递了结束日期，使用该日期的 23:59:59
                            if len(end_date) > 10:
                                # 完整时间格式
                                edt = datetime.datetime.strptime(end_date, '%Y-%m-%d %H:%M:%S')
                            else:
                                # 仅日期格式，使用当天 23:59:59
                                edt = datetime.datetime.strptime(end_date, '%Y-%m-%d')
                                edt = datetime.datetime(edt.year, edt.month, edt.day, 23, 59, 59)
                        else:
                            # 未传递 end_date，使用记录最后时间
                            edt = ldt
                        k7_val = datetime.datetime(edt.year, edt.month, edt.day, edt.hour, edt.minute, edt.second)
                        _ws.write_merge(6, 8, 10, 12, k7_val, k7_style)  # K7:M9
                        _ws.write_merge(2, 2, 11, 12, 1, l3_style)  # L3:M3
                        _ws.write_merge(7, 7, 1, 6, f'日期/时间:{ds}   -   {de} ', a10)  # B8:G8

                    # Row 10:空行

                    # 写入表头行(row 11)- 带边框
                    _ws.write(11, 2, '序号', hdr_style)
                    _ws.write(11, 3, '巡更点', hdr_style)
                    _ws.write_merge(11, 11, 4, 6, '巡检时间', hdr_style)
                    _ws.write_merge(11, 11, 7, 8, '巡检仪器', hdr_style)
                    _ws.write_merge(11, 11, 9, 11, '事件', hdr_style)
                    _ws.write(11, 12, '巡检员', hdr_style)

                    # 写入数据行 - 带边框,行高340(与金模板一致)
                    for i, r in enumerate(records):
                        row0 = 12 + i
                        _ws.row(row0).height = 340
                        _ws.write(row0, 2, str(i + 1), data_style)
                        _ws.write(row0, 3, r.get('point', r.get('point_name', '')), data_style)
                        _ws.write_merge(row0, row0, 4, 6, r.get('time', r.get('patrol_time', '')), data_style)
                        _ws.write_merge(row0, row0, 7, 8, r.get('device', r.get('device_id', '')), data_style)
                        _ws.write_merge(row0, row0, 9, 11, r.get('event', '合格'), data_style)
                        _ws.write(row0, 12, r.get('inspector', ''), data_style)

                    xls_buf = _io.BytesIO(); _wb.save(xls_buf); xls_buf.seek(0)
                    body_bytes = xls_buf.read()

                # 存储到临时存储
                with _export_store_lock:
                    _export_store[token] = (body_bytes, filename, mime)

                # 5分钟后自动清理
                threading.Timer(300, _cleanup_export, args=[token]).start()

                self._send_json({
                    "token": token,
                    "filename": filename,
                    "size": len(body_bytes)
                })
            except Exception as e:
                self._send_json({"error": f"导出失败：{e}"})

        elif parsed.path == "/api/export-pdf":
            # PDF 导出：接收与 XLS 导出相同的参数，生成 PDF 文件
            records = body.get("records", [])
            route_name = body.get("route_name", "")  # 从原数据 G2 单元格继承
            original_filename = body.get("original_filename", "")  # 从原数据文件名继承
            # 时间范围参数：支持用户指定任意起止时间
            start_date = body.get("start_date", "")  # 格式：YYYY-MM-DD 或 YYYY-MM-DD HH:MM:SS
            end_date = body.get("end_date", "")      # 格式：YYYY-MM-DD 或 YYYY-MM-DD HH:MM:SS
            
            # 如果没有传递时间参数，使用默认值（当天）
            if not start_date and records:
                start_date = records[0].get("time", "")[:10] if records[0].get("time") else ""
            if not end_date and records:
                end_date = records[-1].get("time", "")[:10] if records[-1].get("time") else ""
            
            # 生成 date_str 用于 PDF 标题
            if start_date and end_date:
                if start_date == end_date:
                    date_str = start_date
                else:
                    date_str = f"{start_date} 至 {end_date}"
            elif start_date:
                date_str = start_date
            else:
                date_str = datetime.datetime.now().strftime("%Y-%m-%d")
            
            if not records:
                self._send_json({"error": "没有数据可导出"})
                return

            # 生成 token (LOW-012: 使用 secrets 模块生成安全随机 token)
            token = f"export_{int(time.time()*1000)}_{secrets.token_urlsafe(16)}"
            # 优先使用原文件名，否则使用默认名称
            if original_filename:
                # 去除原文件扩展名
                base_name = os.path.splitext(original_filename)[0]
                filename = f"{base_name}.pdf"
            else:
                filename = "patrol_records.pdf"
            mime = "application/pdf"

            try:
                # 步骤 1: 复用 XLS 导出逻辑生成临时 XLS 文件
                import xlwt, io as _io
                _wb = xlwt.Workbook()
                _wb.set_colour_RGB(9, 0xD3, 0xD3, 0xD3)  # 索引 9 覆盖为灰色边框 (211,211,211)
                _ws = _wb.add_sheet('Rpt_Records_ens')
                _ws.show_grid = False  # 关闭网格线 (与金模板一致)
                # 冻结窗格：第 10 列 (J 列) 固定，K 列起可滚动
                _ws.panes_frozen = True
                _ws.horz_split_pos = 10
                _ws.horz_split_first_visible = 10
                _ws.split_active_pane = 2

                # === 样式定义 ===
                # G2 样式:20 号微软雅黑加粗，居中顶端
                g2_style = xlwt.XFStyle()
                g2_style.font.name = '微软雅黑'; g2_style.font.height = 400; g2_style.font.bold = True
                g2_style.alignment.horz = 2; g2_style.alignment.vert = 0

                # B6 样式:14 号微软雅黑加粗，左上角
                b6_style = xlwt.XFStyle()
                b6_style.font.name = '微软雅黑'; b6_style.font.height = 280; b6_style.font.bold = True
                b6_style.alignment.horz = 0; b6_style.alignment.vert = 0

                # K7 样式：微软雅黑 10 号，左上角，locale 日期格式
                k7_style = xlwt.XFStyle()
                k7_style.font.name = '微软雅黑'; k7_style.font.height = 200
                k7_style.alignment.horz = 0; k7_style.alignment.vert = 0
                k7_style.num_format_str = '[$-010804]yyyy/m/d h:mm:ss'

                # L3 编号样式：微软雅黑 10 号，左上角
                l3_style = xlwt.XFStyle()
                l3_style.font.name = '微软雅黑'; l3_style.font.height = 200
                l3_style.alignment.horz = 0; l3_style.alignment.vert = 0

                # A10 样式:Arial 10 号 (B8 日期/时间，family=2 Swiss, charset=0 ANSI)
                a10 = xlwt.XFStyle()
                a10.font.name = 'Arial'; a10.font.height = 200
                a10.font.family = 2; a10.font.charset = 0; a10.font.colour_index = 8
                a10.alignment.horz = 0; a10.alignment.vert = 0

                # 数据区边框 (thin black)
                cell_border = xlwt.Borders()
                cell_border.left = xlwt.Borders.THIN
                cell_border.right = xlwt.Borders.THIN
                cell_border.top = xlwt.Borders.THIN
                cell_border.bottom = xlwt.Borders.THIN
                cell_border.left_colour = 9
                cell_border.right_colour = 9
                cell_border.top_colour = 9
                cell_border.bottom_colour = 9

                # 表头样式:Arial 10 号，居中顶端，带边框
                hdr_style = xlwt.XFStyle()
                hdr_style.font.name = 'Arial'; hdr_style.font.height = 200
                hdr_style.alignment.horz = 2; hdr_style.alignment.vert = 0; hdr_style.alignment.wrap = 1
                hdr_style.borders = cell_border

                # 数据行样式：微软雅黑 10 号，居中顶端，带灰色边框
                data_border = xlwt.Borders()
                data_border.left = xlwt.Borders.THIN; data_border.left_colour = 9
                data_border.right = xlwt.Borders.THIN; data_border.right_colour = 9
                data_border.top = xlwt.Borders.THIN; data_border.top_colour = 9
                data_border.bottom = xlwt.Borders.THIN; data_border.bottom_colour = 9
                data_style = xlwt.XFStyle()
                data_style.font.name = '微软雅黑'; data_style.font.height = 200
                data_style.alignment.horz = 2; data_style.alignment.vert = 0; data_style.alignment.wrap = 1
                data_style.borders = data_border

                # Row 10:空行，但 A10:N10 有黑色顶部边框 (视觉上像下边框)
                row10_border = xlwt.Borders()
                row10_border.top = xlwt.Borders.THIN; row10_border.top_colour = 8  # RGB(0,0,0) 黑色
                row10_style = xlwt.XFStyle()
                row10_style.borders = row10_border
                for ci in range(14):
                    _ws.write(10, ci, '', row10_style)

                # === 设置列宽 (与 XLS 一致，xlwt 单位=1/256 字符宽)===
                col_widths = [548, 292, 1133, 4571, 877, 1170, 3876, 2560, 841, 512, 2048, 475, 3108, 548]
                for ci, w in enumerate(col_widths):
                    _ws.col(ci).width = w

                # === 设置行高 (匹配原始模板)===
                row_heights = [197, 180, 285, 107, 23, 360, 90, 340, 20, 8, 235, 340]
                # 金模板的 height_mismatch 标志 (1=该行需要，0=不需要)
                height_mismatch_flags = [1, 0, 1, 0, 1, 0, 0, 0, 0, 1, 1, 0]
                for ri, h in enumerate(row_heights):
                    _ws.row(ri).height = h
                    if ri < len(height_mismatch_flags) and height_mismatch_flags[ri]:
                        _ws.row(ri).height_mismatch = True

                # === 按照原始模板格式生成 ===
                num_data_rows = len(records)

                # === 非数据区：直接写入内容，无需白色填充 ===
                # G2 单元格：显示线路名称（从请求参数继承）
                g2_value = route_name if route_name else '线路一'
                _ws.write_merge(1, 3, 6, 7, g2_value, g2_style)  # G2:H4
                _ws.write_merge(5, 6, 1, 4, '记录查询', b6_style)  # B6:E7
                if records:
                    ft = records[0].get('time', ''); lt = records[-1].get('time', '')
                    ds = ft.split(' ')[0]; de = lt.split(' ')[0]
                    ldt = datetime.datetime.strptime(lt, '%Y-%m-%d %H:%M:%S')
                    # 恢复灵活时间范围支持：使用前端传递的 end_date 或记录中的最后时间
                    # 如果前端传递了 end_date，则使用该时间的 23:59:59；否则使用记录最后时间的合理扩展
                    if end_date:
                        # 前端传递了结束日期，使用该日期的 23:59:59
                        if len(end_date) > 10:
                            # 完整时间格式
                            edt = datetime.datetime.strptime(end_date, '%Y-%m-%d %H:%M:%S')
                        else:
                            # 仅日期格式，使用当天 23:59:59
                            edt = datetime.datetime.strptime(end_date, '%Y-%m-%d')
                            edt = datetime.datetime(edt.year, edt.month, edt.day, 23, 59, 59)
                    else:
                        # 未传递 end_date，使用记录最后时间
                        edt = ldt
                    k7_val = datetime.datetime(edt.year, edt.month, edt.day, edt.hour, edt.minute, edt.second)
                    _ws.write_merge(6, 8, 10, 12, k7_val, k7_style)  # K7:M9
                    _ws.write_merge(2, 2, 11, 12, 1, l3_style)  # L3:M3
                    _ws.write_merge(7, 7, 1, 6, f'日期/时间:{ds}   -   {de} ', a10)  # B8:G8

                # Row 10:空行

                # 写入表头行 (row 11)- 带边框
                _ws.write(11, 2, '序号', hdr_style)
                _ws.write(11, 3, '巡更点', hdr_style)
                _ws.write_merge(11, 11, 4, 6, '巡检时间', hdr_style)
                _ws.write_merge(11, 11, 7, 8, '巡检仪器', hdr_style)
                _ws.write_merge(11, 11, 9, 11, '事件', hdr_style)
                _ws.write(11, 12, '巡检员', hdr_style)

                # 写入数据行 - 带边框，行高 340(与金模板一致)
                for i, r in enumerate(records):
                    row0 = 12 + i
                    _ws.row(row0).height = 340
                    _ws.write(row0, 2, str(i + 1), data_style)
                    _ws.write(row0, 3, r.get('point', r.get('point_name', '')), data_style)
                    _ws.write_merge(row0, row0, 4, 6, r.get('time', r.get('patrol_time', '')), data_style)
                    _ws.write_merge(row0, row0, 7, 8, r.get('device', r.get('device_id', '')), data_style)
                    _ws.write_merge(row0, row0, 9, 11, r.get('event', '合格'), data_style)
                    _ws.write(row0, 12, r.get('inspector', ''), data_style)

                # 保存到临时 XLS 文件（xls2pdf-converter 需要文件路径输入）
                import tempfile
                fd_xls, tmp_xls_path = tempfile.mkstemp(suffix='.xls')
                os.close(fd_xls)
                _wb.save(tmp_xls_path)

                try:
                    # 步骤 2: 使用 xls2pdf-converter 转换为 PDF
                    fd_pdf, tmp_pdf_path = tempfile.mkstemp(suffix='.pdf')
                    os.close(fd_pdf)
                    try:
                        xls2pdf_convert(tmp_xls_path, tmp_pdf_path)
                        with open(tmp_pdf_path, 'rb') as pf:
                            pdf_bytes = pf.read()
                        print(f"[PDF] PDF 导出完成，体积: {len(pdf_bytes) / 1024:.1f} KB")
                    finally:
                        if os.path.exists(tmp_pdf_path):
                            os.unlink(tmp_pdf_path)
                finally:
                    # 清理临时 XLS 文件
                    if os.path.exists(tmp_xls_path):
                        os.unlink(tmp_xls_path)
                
                # 存储到临时存储
                with _export_store_lock:
                    _export_store[token] = (pdf_bytes, filename, mime)

                # 5 分钟后自动清理
                threading.Timer(300, _cleanup_export, args=[token]).start()

                self._send_json({
                    "token": token,
                    "filename": filename,
                    "size": len(pdf_bytes)
                })

            except Exception as e:
                self._send_json({"error": f"PDF 导出失败：{e}"})

        elif parsed.path == "/api/points/toggle":
            pid = body.get("id")
            pwd = body.get("pwd", "")
            if not check_admin_op_password(pwd):
                self._send_json({"error": "密码错误"}, 403); return
            db = get_db()
            db.execute("UPDATE patrol_points SET is_active = 1 - is_active WHERE id = ?", [pid])
            db.commit()
            # MEDIUM-007 修复：更新后用 SELECT 验证是否真正影响了行
            row = db.execute("SELECT id FROM patrol_points WHERE id = ?", [pid]).fetchone()
            db.close()
            if not row:
                self._send_json({"error": "点位不存在"}, 404)
                return
            self._send_json({"ok": True})

        elif parsed.path == "/api/points/delete":
            pid = body.get("id")
            pwd = body.get("pwd", "")
            if not check_admin_op_password(pwd):
                self._send_json({"error": "密码错误"}, 403); return
            db = get_db()
            # MEDIUM-007 修复：删除前先用 SELECT 验证点位存在
            row = db.execute("SELECT id FROM patrol_points WHERE id = ?", [pid]).fetchone()
            if not row:
                db.close()
                self._send_json({"error": "点位不存在"}, 404)
                return
            db.execute("DELETE FROM patrol_points WHERE id = ?", [pid])
            db.commit()
            db.close()
            self._send_json({"ok": True})

        else:
            self.send_error(404)

    def do_PUT(self):
        """处理 PUT 请求 (用于更新点位标签)"""
        parsed = urlparse(self.path)
        print(f"[PUT] {parsed.path} from {self.client_address}", flush=True)
        
        # 导入鉴权函数（用于密码验证）
        from auth import check_admin_op_password
        
        # 标签编辑接口特殊处理：支持密码验证替代 session token
        is_label_endpoint = (parsed.path.startswith("/api/points/") and parsed.path.endswith("/label")) or parsed.path == "/api/points/labels"

        # 验证 session token (标签接口可选，可用密码替代)
        from auth import verify_session_token, is_auth_enabled
        from server import get_db
        if is_auth_enabled() and not is_label_endpoint:
            token = self.headers.get('X-Auth-Token', '')
            db = get_db()
            try:
                if not token or not verify_session_token(token, db):
                    self._send_json({"error": "请先登录"}, 401)
                    return
            finally:
                db.close()

        try:
            content_len = int(self.headers.get('Content-Length', 0))
            raw_body = self.rfile.read(content_len) if content_len > 0 else b''
            body = json.loads(raw_body) if raw_body else {}
        except Exception as e:
            self._send_json({"error": f"请求数据解析失败:{e}"}, 400)
            return

        # PUT /api/points/:id/label - 更新单个点位标签
        if parsed.path.startswith("/api/points/") and parsed.path.endswith("/label"):
            parts = parsed.path.split("/")
            if len(parts) == 5 and parts[1] == "api" and parts[2] == "points":
                try:
                    point_id = int(parts[3])
                except:
                    self._send_json({"error": "无效的点位 ID"}, 400)
                    return

                # 标签编辑/清除操作需要额外密码验证
                auth_password = body.get("password") or self.headers.get('Authorization', '').replace('Bearer ', '')
                if not check_admin_op_password(auth_password):
                    self._send_json({"error": "密码错误"}, 403)
                    return

                # 支持两种格式：分字段或完整标签
                label_l1 = body.get("label_l1")
                label_l1_order = body.get("label_l1_order")
                label_l2 = body.get("label_l2")
                label_l2_order = body.get("label_l2_order")
                tag_full = body.get("tag_full")

                # 如果提供了 tag_full,自动解析
                if tag_full and not label_l1:
                    parsed_tag = parse_tag_full(tag_full)
                    if not parsed_tag:
                        self._send_json({"error": "标签格式错误,应为:A 端 1-C 口 2"}, 400)
                        return
                    label_l1 = parsed_tag["label_l1"]
                    label_l1_order = parsed_tag["label_l1_order"]
                    label_l2 = parsed_tag["label_l2"]
                    label_l2_order = parsed_tag["label_l2_order"]

                # 验证必填字段
                if not all([label_l1, label_l1_order is not None, label_l2, label_l2_order is not None]):
                    self._send_json({"error": "缺少必填字段"}, 400)
                    return

                # 生成完整标签
                tag_full_generated = f"{label_l1.strip()} {label_l1_order}-{label_l2.strip()} {label_l2_order}"

                # 更新数据库
                db = get_db()
                try:
                    cursor = db.execute("""
                        UPDATE patrol_points
                        SET label_l1=?, label_l1_order=?, label_l2=?, label_l2_order=?, tag_full=?
                        WHERE id=?
                    """, [label_l1, label_l1_order, label_l2, label_l2_order, tag_full_generated, point_id])
                    db.commit()
                    if cursor.rowcount == 0:
                        db.close()
                        self._send_json({"error": "点位不存在"}, 404)
                        return
                    db.close()

                    # 标签保存后自动按标签排序
                    try:
                        recalculate_sort_order_by_label()
                    except Exception as sort_e:
                        print(f"[WARN] 标签保存后排序失败(不影响标签保存): {sort_e}", flush=True)

                    self._send_json({
                        "success": True,
                        "point_id": point_id,
                        "label": tag_full_generated
                    })
                except Exception as e:
                    db.close()
                    self._send_json({"error": f"更新失败:{e}"}, 500)
                return

        # PUT /api/points/labels - 批量更新点位标签
        if parsed.path == "/api/points/labels":
            # 批量标签操作也需要密码验证
            auth_password = body.get("password") if isinstance(body, dict) else self.headers.get('Authorization', '').replace('Bearer ', '')
            if isinstance(body, list):
                auth_password = self.headers.get('Authorization', '').replace('Bearer ', '')
                if not auth_password and len(body) > 0 and isinstance(body[0], dict):
                    auth_password = body[0].get('password')
            if not check_admin_op_password(auth_password):
                self._send_json({"error": "密码错误"}, 403)
                return

            if not isinstance(body, list):
                self._send_json({"error": "请求体应为数组"}, 400)
                return

            updated = 0
            failed = 0
            db = get_db()

            for item in body:
                point_id = item.get("id")
                tag_full = item.get("tag_full")

                if not point_id or not tag_full:
                    failed += 1
                    continue

                parsed_tag = parse_tag_full(tag_full)
                if not parsed_tag:
                    failed += 1
                    continue

                try:
                    tag_full_generated = f"{parsed_tag['label_l1'].strip()} {parsed_tag['label_l1_order']}-{parsed_tag['label_l2'].strip()} {parsed_tag['label_l2_order']}"
                    db.execute("""
                        UPDATE patrol_points
                        SET label_l1=?, label_l1_order=?, label_l2=?, label_l2_order=?, tag_full=?
                        WHERE id=?
                    """, [
                        parsed_tag["label_l1"],
                        parsed_tag["label_l1_order"],
                        parsed_tag["label_l2"],
                        parsed_tag["label_l2_order"],
                        tag_full_generated,
                        point_id
                    ])
                    updated += 1
                except:
                    failed += 1

            db.commit()
            db.close()

            # 批量标签更新后自动按标签排序
            try:
                recalculate_sort_order_by_label()
            except Exception as sort_e:
                print(f"[WARN] 批量标签保存后排序失败: {sort_e}", flush=True)

            self._send_json({
                "success": True,
                "updated": updated,
                "failed": failed
            })
            return

        self.send_error(404)

    def do_DELETE(self):
        """处理 DELETE 请求 (用于清除点位标签)"""
        parsed = urlparse(self.path)
        print(f"[DELETE] {parsed.path} from {self.client_address}", flush=True)
        
        # 导入鉴权函数（用于密码验证）
        from auth import check_admin_op_password
        
        # 标签清除接口特殊处理：支持密码验证替代 session token
        is_label_endpoint = (parsed.path.startswith("/api/points/") and parsed.path.endswith("/label")) or parsed.path == "/api/points/labels"

        # 验证 session token (标签接口可选，可用密码替代)
        from auth import verify_session_token, is_auth_enabled
        from server import get_db
        if is_auth_enabled() and not is_label_endpoint:
            token = self.headers.get('X-Auth-Token', '')
            db = get_db()
            try:
                if not token or not verify_session_token(token, db):
                    self._send_json({"error": "请先登录"}, 401)
                    return
            finally:
                db.close()

        # DELETE /api/points/:id/label - 清除单个点位标签
        if parsed.path.startswith("/api/points/") and parsed.path.endswith("/label"):
            parts = parsed.path.split("/")
            if len(parts) == 5 and parts[1] == "api" and parts[2] == "points":
                # 标签清除操作需要额外密码验证
                # 尝试从请求体或 Authorization 头获取密码
                try:
                    content_len = int(self.headers.get('Content-Length', 0))
                    raw_body = self.rfile.read(content_len) if content_len > 0 else b''
                    body = json.loads(raw_body) if raw_body else {}
                    auth_password = body.get("password") or self.headers.get('Authorization', '').replace('Bearer ', '')
                except:
                    auth_password = self.headers.get('Authorization', '').replace('Bearer ', '')

                if not check_admin_op_password(auth_password):
                    self._send_json({"error": "密码错误"}, 403)
                    return

                try:
                    point_id = int(parts[3])
                except:
                    self._send_json({"error": "无效的点位 ID"}, 400)
                    return

                db = get_db()
                try:
                    cursor = db.execute("""
                        UPDATE patrol_points
                        SET label_l1=NULL, label_l1_order=NULL, label_l2=NULL, label_l2_order=NULL, tag_full=NULL
                        WHERE id=?
                    """, [point_id])
                    db.commit()
                    if cursor.rowcount == 0:
                        db.close()
                        self._send_json({"error": "点位不存在"}, 404)
                        return
                    db.close()

                    # 标签清除后自动按标签排序
                    try:
                        recalculate_sort_order_by_label()
                    except Exception as sort_e:
                        print(f"[WARN] 标签清除后排序失败(不影响清除): {sort_e}", flush=True)

                    self._send_json({
                        "success": True,
                        "point_id": point_id
                    })
                except Exception as e:
                    db.close()
                    self._send_json({"error": f"删除失败:{e}"}, 500)
                return

        self.send_error(404)


def _cleanup_export(token):
    """清理导出临时文件"""
    with _export_store_lock:
        _export_store.pop(token, None)

def is_device_authorized():
    """检查设备是否已授权（仅设备指纹鉴权）
    
    授权检查逻辑：
    1. 生成 device_fingerprint
    2. 在 mac_activations 表中匹配是否存在
    
    Returns:
        bool: 设备是否已授权
    """
    from auth import get_device_fingerprint
    
    db = get_db()
    try:
        # 仅匹配设备指纹
        try:
            fingerprint = get_device_fingerprint()
            if fingerprint:
                row = db.execute(
                    "SELECT 1 FROM mac_activations WHERE device_fingerprint = ?",
                    (fingerprint,)
                ).fetchone()
                if row is not None:
                    return True
        except Exception as e:
            print(f"[WARN] 获取设备指纹失败：{e}")
        
        return False
    except Exception as e:
        print(f"[ERROR] is_device_authorized: {e}")
        return False
    finally:
        db.close()

def activate_device(otp_code):
    """验证 2FA TOTP 并激活当前设备（仅设备指纹鉴权）
    
    激活逻辑：
    1. 验证 2FA 代码
    2. 生成设备指纹并存储
    
    Args:
        otp_code: 6 位 TOTP 代码
    
    Returns:
        dict: {success: bool, activated_at: str, fingerprint: str, error: str (optional)}
    """
    import pyotp
    import base64
    from auth import get_system_config, get_device_fingerprint
    
    # 从数据库读取 2FA 密钥（HIGH-001/002 修复：使用 Base32 编码）
    secret = get_system_config("2fa_secret_key")
    if not secret:
        # 首次使用时生成并保存 2FA 密钥（Base32 编码，20 字节=160 位）
        secret = pyotp.random_base32()  # 生成 32 字符 Base32 字符串
        db = get_db()
        try:
            db.execute(
                "INSERT OR REPLACE INTO system_config (key, value, description) VALUES (?, ?, ?)",
                ("2fa_secret_key", secret, "2FA TOTP 密钥 (Base32 编码，32 字符)")
            )
            db.commit()
        finally:
            db.close()
    totp = pyotp.TOTP(secret, interval=30, digits=6, digest='SHA1')
    if not totp.verify(otp_code, valid_window=1):
        return {"success": False, "activated_at": None}
    
    # 生成设备指纹
    try:
        fingerprint = get_device_fingerprint()
    except Exception as e:
        print(f"[WARN] 获取设备指纹失败：{e}")
        fingerprint = None
    
    if not fingerprint:
        return {"success": False, "activated_at": None, "error": "无法获取设备指纹"}
    
    import datetime
    now = datetime.datetime.now().isoformat()
    db = get_db()
    try:
        # 仅存储设备指纹
        db.execute(
            "INSERT OR IGNORE INTO mac_activations (device_fingerprint, activated_at) VALUES (?, ?)",
            (fingerprint, now)
        )
        db.commit()
        
        # 返回激活信息
        row = db.execute(
            "SELECT activated_at FROM mac_activations WHERE device_fingerprint = ?",
            (fingerprint,)
        ).fetchone()
        activated_at = row["activated_at"] if row else now
        db.close()
        
        return {
            "success": True, 
            "activated_at": activated_at,
            "fingerprint": fingerprint
        }
    except Exception as e:
        db.close()
        return {"success": False, "activated_at": None, "error": str(e)}

def run_server(port=28891):
    log_error("run_server: 开始")
    log_error("run_server: 调用 init_db()...")
    init_db()
    log_error("run_server: init_db 完成")
    
    # v2.4 简化：设备指纹授权检查（仅指纹，无 MAC）
    from auth import get_device_fingerprint
    
    # 获取设备指纹
    try:
        fingerprint = get_device_fingerprint()
    except Exception as e:
        log_error(f"run_server: 获取设备指纹失败：{e}")
        fingerprint = None
    
    # 授权检查
    authorized = is_device_authorized()
    
    if fingerprint:
        log_error(f"run_server: 设备指纹={fingerprint}")
    
    if not authorized:
        print("❌ 当前设备未授权！请通过 /api/auth/activate 接口激活")
        if fingerprint:
            print(f"   设备指纹：{fingerprint}")
    else:
        print(f"🔒 设备授权通过")
        if fingerprint:
            print(f"   设备指纹：{fingerprint}")
    # [SECURITY] 初始化系统安全配置（首次启动时创建默认配置）
    from auth import init_security_db, get_system_config, check_admin_op_password
    init_security_db()
    log_error("run_server: 安全配置初始化完成")
    
    # 验证签名密钥配置
    sig_key = get_system_config("signature_key")
    if not sig_key:
        log_error("run_server: 警告 - signature_key 未配置")
    else:
        log_error("run_server: signature_key 已配置")
    
    
    server = ThreadingHTTPServer(("0.0.0.0", port), PatrolHandler)
    log_error("run_server: ThreadingHTTPServer 创建成功")
    log_error("run_server: 调用 db_get_stats()...")
    stats = db_get_stats()
    log_error(f"run_server: db_get_stats 返回 {stats}")
    print(f"🛡️ 巡更记录管理系统 2.1 启动于 http://0.0.0.0:{port}")
    print(f"📊 数据库:{stats['total_records']} 条记录")
    from auth import is_auth_enabled
    if is_auth_enabled():
        print(f"🔑 密码保护已启用")
    log_error("run_server: 准备自动打开浏览器...")
    # 自动打开默认浏览器
    import webbrowser
    import threading
    def open_browser():
        import time
        time.sleep(0.5)  # 等待服务器完全启动
        webbrowser.open(f"http://localhost:{port}")
    threading.Thread(target=open_browser, daemon=True).start()
    log_error("run_server: 已启动浏览器打开线程")
    log_error("run_server: 进入 serve_forever()")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
        server.server_close()

def log_error(msg):
    """记录错误到日志文件"""
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now()}] {msg}\n")
            f.flush()  # 强制刷新
    except Exception as e:
        try:
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(f"Log write failed: {e}\n")
                f.flush()
        except:
            pass

if __name__ == "__main__":
    try:
        log_error("=== 主程序开始 ===")
        port = int(sys.argv[1]) if len(sys.argv) > 1 else 28891
        log_error(f"端口:{port}")
        log_error("调用 run_server()...")
        run_server(port)
        log_error("服务器运行中")
    except SystemExit as e:
        log_error(f"SystemExit: {e}")
    except Exception as e:
        log_error(f"启动失败:{type(e).__name__}: {e}")
        import traceback
        log_error(traceback.format_exc())
        # 显示错误对话框 (Windows)
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"启动失败:{e}\n\n请查看 startup.log 文件", "巡更系统错误", 0x10)
        except:
            pass
        sys.exit(1)
