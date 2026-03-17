import sqlite3
import os
from datetime import datetime

# ── Database path ──────────────────────────────────────────────
# On Render with a persistent disk: set DB_PATH=/data/database.db
# Local dev: falls back to ~/InterviewAutomation/database.db
DB_PATH = os.environ.get(
    'DB_PATH',
    os.path.join(
        os.path.expanduser("~"), "InterviewAutomation", "database.db"
    )
)

# Ensure the directory exists (important for first Render boot)
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)


def init_db():
    """Initialize the SQLite database and create tables if they don't exist."""
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()

        # Users table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                name           TEXT    NOT NULL,
                signup_date    TEXT    NOT NULL,
                location       TEXT,
                distance       REAL,
                attempt_number TEXT,
                dob            TEXT,
                created_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Backward-compatibility: add columns that may be missing in older DBs
        for col, col_type in [
            ('location',       'TEXT'),
            ('distance',       'REAL'),
            ('attempt_number', 'TEXT'),
            ('dob',            'TEXT'),
        ]:
            try:
                cursor.execute(f'ALTER TABLE users ADD COLUMN {col} {col_type}')
            except sqlite3.OperationalError:
                pass  # Column already exists

        # Typing results table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS typing_results (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name  TEXT    NOT NULL,
                wpm        REAL    NOT NULL,
                accuracy   REAL    NOT NULL,
                time_limit INTEGER NOT NULL,
                test_date  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        conn.commit()


def insert_user(name, signup_date, location, distance, attempt_number, dob):
    """Insert a new candidate into the users table."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO users (name, signup_date, location, distance, attempt_number, dob)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (name, signup_date, location, float(distance), attempt_number, dob))
        conn.commit()


def insert_typing_result(user_name, wpm, accuracy, time_limit):
    """Insert a typing test result into the typing_results table."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO typing_results (user_name, wpm, accuracy, time_limit, test_date)
            VALUES (?, ?, ?, ?, ?)
        ''', (user_name, wpm, accuracy, time_limit,
              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()