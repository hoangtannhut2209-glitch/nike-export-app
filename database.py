# -*- coding: utf-8 -*-
"""
Database models and initialization for Nike Export Web App
"""

import sqlite3
import os
from datetime import datetime

DATABASE_FILE = 'nike_export.db'

def get_db_connection():
    """Get database connection"""
    conn = sqlite3.connect(DATABASE_FILE)
    conn.row_factory = sqlite3.Row  # Enable dict-like access to rows
    return conn

def init_db():
    """Initialize database with required tables"""
    conn = get_db_connection()
    
    # Create invoices table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_number TEXT UNIQUE NOT NULL,
            date_processed TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'pending',
            pl_file_path TEXT,
            booking_file_path TEXT,
            invoice_file_path TEXT,
            data_json TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Create extracted_data table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS extracted_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id INTEGER,
            field_name TEXT NOT NULL,
            field_value TEXT,
            source_file TEXT,
            confidence REAL DEFAULT 1.0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (invoice_id) REFERENCES invoices (id)
        )
    ''')
    
    # Create templates table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            description TEXT,
            file_path TEXT NOT NULL,
            is_active BOOLEAN DEFAULT 1,
            placeholders TEXT,  -- JSON list of placeholders
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Create export_jobs table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS export_jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT UNIQUE NOT NULL,
            status TEXT DEFAULT 'pending',
            invoice_ids TEXT,  -- JSON array of invoice IDs
            template_ids TEXT,  -- JSON array of template IDs
            export_mode TEXT NOT NULL,
            output_path TEXT,
            progress REAL DEFAULT 0.0,
            error_message TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            completed_at TIMESTAMP
        )
    ''')
    
    # Create processing_logs table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS processing_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT,
            invoice_number TEXT,
            log_level TEXT DEFAULT 'INFO',
            message TEXT NOT NULL,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Insert default templates if not exist
    default_templates = [
        ('CO-FORM A', 'Certificate of Origin Form A', 'templates/CO_FORM_A.xlsx'),
        ('CO-FORM E', 'Certificate of Origin Form E', 'templates/CO_FORM_E.xlsx'),
        ('CO-FORM EVFTA', 'Certificate of Origin EVFTA', 'templates/CO_FORM_EVFTA.xlsx'),
        ('CO-FORM CPTPP', 'Certificate of Origin CPTPP', 'templates/CO_FORM_CPTPP.xlsx'),
        ('CO-FORM RCEP', 'Certificate of Origin RCEP', 'templates/CO_FORM_RCEP.xlsx'),
    ]
    
    for name, desc, path in default_templates:
        conn.execute('''
            INSERT OR IGNORE INTO templates (name, description, file_path, placeholders)
            VALUES (?, ?, ?, ?)
        ''', (name, desc, path, '[]'))
    
    conn.commit()
    conn.close()

def add_invoice(invoice_number, status='pending'):
    """Add new invoice to database"""
    conn = get_db_connection()
    try:
        conn.execute('''
            INSERT INTO invoices (invoice_number, status)
            VALUES (?, ?)
        ''', (invoice_number, status))
        conn.commit()
        invoice_id = conn.lastrowid
        return invoice_id
    except sqlite3.IntegrityError:
        # Invoice already exists, get existing ID
        row = conn.execute('''
            SELECT id FROM invoices WHERE invoice_number = ?
        ''', (invoice_number,)).fetchone()
        return row['id'] if row else None
    finally:
        conn.close()

def update_invoice_status(invoice_id, status, data_json=None):
    """Update invoice status and data"""
    conn = get_db_connection()
    if data_json:
        conn.execute('''
            UPDATE invoices 
            SET status = ?, data_json = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ''', (status, data_json, invoice_id))
    else:
        conn.execute('''
            UPDATE invoices 
            SET status = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ''', (status, invoice_id))
    conn.commit()
    conn.close()

def get_invoice_by_number(invoice_number):
    """Get invoice by invoice number"""
    conn = get_db_connection()
    row = conn.execute('''
        SELECT * FROM invoices WHERE invoice_number = ?
    ''', (invoice_number,)).fetchone()
    conn.close()
    return dict(row) if row else None

def get_all_invoices():
    """Get all invoices"""
    conn = get_db_connection()
    rows = conn.execute('''
        SELECT * FROM invoices ORDER BY created_at DESC
    ''').fetchall()
    conn.close()
    return [dict(row) for row in rows]

def add_extracted_data(invoice_id, field_name, field_value, source_file=None):
    """Add extracted data field"""
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO extracted_data (invoice_id, field_name, field_value, source_file)
        VALUES (?, ?, ?, ?)
    ''', (invoice_id, field_name, field_value, source_file))
    conn.commit()
    conn.close()

def get_extracted_data(invoice_id):
    """Get all extracted data for an invoice"""
    conn = get_db_connection()
    rows = conn.execute('''
        SELECT * FROM extracted_data WHERE invoice_id = ?
    ''', (invoice_id,)).fetchall()
    conn.close()
    return [dict(row) for row in rows]

def log_processing(job_id, invoice_number, message, level='INFO'):
    """Add processing log entry"""
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO processing_logs (job_id, invoice_number, log_level, message)
        VALUES (?, ?, ?, ?)
    ''', (job_id, invoice_number, level, message))
    conn.commit()
    conn.close()

if __name__ == '__main__':
    init_db()
    print("Database initialized successfully!")