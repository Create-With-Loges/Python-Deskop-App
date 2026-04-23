import sqlite3
import os

import sys
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    base_path = os.path.dirname(sys.executable)
else:
    # Running as script
    base_path = os.path.dirname(os.path.abspath(__file__))

DB_NAME = os.path.join(base_path, "exam_system.db")

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Admin Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        )
    ''')
    
    # Insert default admin if not exists
    cursor.execute("SELECT * FROM admin WHERE user_id = 'admin@cia'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO admin (user_id, password) VALUES ('admin@cia', 'admin@gasc')")

    # 2. Staff Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS staff (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            department TEXT NOT NULL,
            designation TEXT NOT NULL,
            joining_date TEXT NOT NULL
        )
    ''')

    # 3. Exam Hall Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS halls (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hall_number TEXT UNIQUE NOT NULL
        )
    ''')
    
    # Migration: Remove gender_type if it exists
    try:
        # Check if gender_type column exists
        cursor.execute("SELECT gender_type FROM halls LIMIT 1")
        # If no error, column exists. We need to restructure.
        print("Migrating database: Removing gender_type column...")
        
        cursor.execute("ALTER TABLE halls RENAME TO halls_old")
        cursor.execute('''
            CREATE TABLE halls (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                hall_number TEXT UNIQUE NOT NULL
            )
        ''')
        
        # Copy data
        cursor.execute("INSERT INTO halls (id, hall_number) SELECT id, hall_number FROM halls_old")
        cursor.execute("DROP TABLE halls_old")
        print("Migration complete.")
        
    except sqlite3.OperationalError:
        # Column likely doesn't exist, which is good.
        pass
    except Exception as e:
        print(f"Migration warning: {e}")

    # 4. Allotment History (Optional, for record keeping)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS allotments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_date TEXT,
            hall_number TEXT,
            staff_name TEXT,
            session_details TEXT
        )
    ''')

    conn.commit()
    conn.close()
    print("Database initialized successfully.")

if __name__ == "__main__":
    init_db()
