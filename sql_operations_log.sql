-- ============================================================
-- SQL 操作记录 - patrol_records 表删除
-- 执行时间：2026-07-01
-- 数据库：/root/patrol-server-2.0/data/patrol.db
-- ============================================================

-- 1. 检查 patrol_records 表是否存在
-- 结果：表不存在（可能之前已被删除）
-- SELECT name FROM sqlite_master WHERE type='table' AND name='patrol_records';

-- 2. 如果表存在，执行删除操作（实际未执行，因为表已不存在）
-- DROP TABLE IF EXISTS patrol_records;

-- 3. 删除相关索引（如果存在）
-- DROP INDEX IF EXISTS idx_patrol_date;
-- DROP INDEX IF EXISTS idx_patrol_point;

-- ============================================================
-- 数据库状态确认
-- ============================================================

-- 当前数据库表列表:
-- auth_users, patrol_points, patrol_routes, patrol_route_points, session_tokens, system_config

-- patrol_points 表状态:
-- 总记录数：25
-- ID 范围：1-84730
-- sqlite_sequence 值：84730

-- 点位标签状态:
-- 所有 25 个点位都有完整的 label_l1, label_l2, tag_full 字段

-- ============================================================
-- 代码修改说明
-- ============================================================

-- 由于 patrol_records 表已删除，以下代码修改确保系统正常运行:
-- 1. 移除 CREATE TABLE patrol_records 语句
-- 2. db_get_stats() 返回空统计信息
-- 3. save_records_to_db() 函数已删除
-- 4. /api/records 端点返回空记录列表
-- 5. /api/records/save 和 /api/save-modified 端点不存储数据

-- ============================================================
-- 备份文件
-- ============================================================
-- 代码备份：server.py.bak.before_delete_records
-- 数据库备份：data/patrol.db.bak.before_delete_records
