import sqlite3
from flask import g, current_app
from datetime import datetime

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(
            current_app.config['DATABASE'],
            detect_types=sqlite3.PARSE_DECLTYPES
        )
        g.db.row_factory = sqlite3.Row
    return g.db

def close_db(e=None):
    db = g.pop('db', None)
    if db is not None:
        db.close()

def init_db():
    db = get_db()
    
    # Employees table schema
    db.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            employee_id TEXT NOT NULL,
            designation TEXT,
            dob TEXT,
            date_of_joining TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    # Leave requests table schema
    db.execute('''
        CREATE TABLE IF NOT EXISTS leave_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT,
            dates TEXT,
            num_days INTEGER,
            reason TEXT,
            description TEXT,
            request_date TEXT,
            status TEXT NOT NULL DEFAULT 'Pending',
            submitted_by TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    # Users table schema for dynamic web-based user management
    db.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            linked_employee_id INTEGER REFERENCES employees(id),
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    # Attendance table: one row per employee per calendar day.
    db.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL REFERENCES employees(id),
            work_date TEXT NOT NULL,
            check_in TEXT,
            check_out TEXT,
            status TEXT NOT NULL DEFAULT 'Present',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(employee_id, work_date)
        );
    ''')

    # Holidays table: company-wide holidays shown on every employee's calendar.
    db.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            holiday_date TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    # Ensure default admin user always exists (recreates it if deleted)
    admin_exists = db.execute('SELECT id FROM users WHERE username = ?', ('admin',)).fetchone()
    if not admin_exists:
        local_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        db.execute(
            'INSERT INTO users (username, password, role, created_at) VALUES (?, ?, ?, ?)',
            ('admin', 'admin123', 'admin', local_timestamp)
        )
    
    # Safely migrate existing databases if columns are missing
    cursor = db.execute("PRAGMA table_info(leave_requests);")
    columns = [col['name'] for col in cursor.fetchall()]
    
    if 'reason' not in columns:
        db.execute('ALTER TABLE leave_requests ADD COLUMN reason TEXT;')
        
    if 'description' not in columns:
        db.execute('ALTER TABLE leave_requests ADD COLUMN description TEXT;')

    user_columns = [col['name'] for col in db.execute("PRAGMA table_info(users);").fetchall()]
    if 'linked_employee_id' not in user_columns:
        db.execute('ALTER TABLE users ADD COLUMN linked_employee_id INTEGER REFERENCES employees(id);')

    attendance_columns = [col['name'] for col in db.execute("PRAGMA table_info(attendance);").fetchall()]
    if 'created_at' not in attendance_columns:
        db.execute('ALTER TABLE attendance ADD COLUMN created_at TEXT;')

    db.commit()

def init_app(app):
    app.teardown_appcontext(close_db)
    with app.app_context():
        init_db()
