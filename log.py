import sqlite3
from datetime import datetime

def log_upload(file_name, row_counts):
    conn = sqlite3.connect('uploads_log.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS uploads (
            timestamp TEXT,
            file_name TEXT,
            transactions INT,
            customers INT,
            products INT
        )
    ''')
    c.execute('INSERT INTO uploads VALUES (?, ?, ?, ?, ?)', (
        datetime.now().isoformat(),
        file_name,
        row_counts['Transactions'],
        row_counts['Customers'],
        row_counts['Products']
    ))
    conn.commit()
    conn.close()
