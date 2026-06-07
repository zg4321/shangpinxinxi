import sqlite3
import os
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS import_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    import_time TIMESTAMP,
                    source_file TEXT,
                    source_sheet TEXT,
                    total_rows INTEGER,
                    target_files TEXT,
                    import_type TEXT DEFAULT 'data'
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS import_details (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    history_id INTEGER,
                    import_type TEXT DEFAULT 'data',
                    target_file TEXT,
                    source_row INTEGER,
                    product_name TEXT,
                    style TEXT,
                    color TEXT,
                    status TEXT,
                    message TEXT,
                    source_row_data TEXT,
                    FOREIGN KEY(history_id) REFERENCES import_history(id)
                )
            ''')
            cursor.execute("PRAGMA table_info(import_history)")
            columns = [col[1] for col in cursor.fetchall()]
            if "import_type" not in columns:
                cursor.execute("ALTER TABLE import_history ADD COLUMN import_type TEXT DEFAULT 'data'")
            cursor.execute("PRAGMA table_info(import_details)")
            columns = [col[1] for col in cursor.fetchall()]
            if "import_type" not in columns:
                cursor.execute("ALTER TABLE import_details ADD COLUMN import_type TEXT DEFAULT 'data'")
            if "source_row_data" not in columns:
                cursor.execute("ALTER TABLE import_details ADD COLUMN source_row_data TEXT")

    def record_history(self, source_file, source_sheet, total_rows, target_files_list, import_type='data'):
        local_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO import_history (import_time, source_file, source_sheet, total_rows, target_files, import_type)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (local_time,
                  os.path.basename(source_file),
                  source_sheet,
                  total_rows,
                  ','.join(target_files_list),
                  import_type))
            return cursor.lastrowid

    def record_detail(self, history_id, target_file, source_row, product_name, style, color, status, message, row_data_json, import_type='data'):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO import_details (history_id, import_type, target_file, source_row, product_name, style, color, status, message, source_row_data)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (history_id, import_type, target_file, source_row, product_name, style, color, status, message, row_data_json))

    def get_history(self, filters=None):
        conditions = []
        params = []
        if filters:
            if filters.get('start_date'):
                conditions.append("import_time >= ?")
                params.append(filters['start_date'])
            if filters.get('end_date'):
                conditions.append("import_time <= ?")
                params.append(filters['end_date'] + " 23:59:59")
            if filters.get('target_file') and filters['target_file'] != '全部':
                conditions.append("id IN (SELECT DISTINCT history_id FROM import_details WHERE target_file = ?)")
                params.append(filters['target_file'])
            if filters.get('import_type'):
                conditions.append("import_type = ?")
                params.append(filters['import_type'])
        where_clause = " AND ".join(conditions) if conditions else "1"
        query = f"SELECT id, import_time, source_file, source_sheet, total_rows, target_files, import_type FROM import_history WHERE {where_clause} ORDER BY import_time DESC"
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return cursor.fetchall()

    def get_details(self, history_id):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT source_row, target_file, product_name, style, color, status, message, source_row_data FROM import_details WHERE history_id=?", (history_id,))
            return cursor.fetchall()

    def get_distinct_target_files(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT target_file FROM import_details")
            return [row[0] for row in cursor.fetchall()]

    def delete_history(self, history_id):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM import_details WHERE history_id = ?", (history_id,))
            cursor.execute("DELETE FROM import_history WHERE id = ?", (history_id,))

    def vacuum(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("VACUUM")