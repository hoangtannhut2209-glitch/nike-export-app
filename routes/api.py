# -*- coding: utf-8 -*-
"""
API routes for Nike Export Web App
"""

from flask import Blueprint, request, jsonify, send_file, current_app
from werkzeug.utils import secure_filename
import os
import json
import uuid
from datetime import datetime
import threading
import time

# Import core modules
from core.pdf_processor import process_pdf_files, get_invoice_data_dict
from database import (
    add_invoice, update_invoice_status, get_all_invoices, 
    get_invoice_by_number, add_extracted_data, log_processing
)

api_bp = Blueprint('api', __name__)

# In-memory job tracking (in production, use Redis or database)
active_jobs = {}

@api_bp.route('/upload', methods=['POST'])
def upload_files():
    """Handle file upload and processing"""
    try:
        # Get form data
        invoice_number = request.form.get('invoice_number', '').strip()
        if not invoice_number:
            return jsonify({'success': False, 'message': 'Invoice number is required'}), 400
        
        # Get files
        pl_file = request.files.get('pl_file')
        booking_file = request.files.get('booking_file')
        additional_files = request.files.getlist('additional_files')
        
        if not pl_file and not booking_file:
            return jsonify({'success': False, 'message': 'At least one PDF file is required'}), 400
        
        # Get options
        auto_process = request.form.get('auto_process', 'false').lower() == 'true'
        extract_dest = request.form.get('extract_dest', 'true').lower() == 'true'
        validate_data = request.form.get('validate_data', 'true').lower() == 'true'
        save_to_db = request.form.get('save_to_db', 'true').lower() == 'true'
        
        # Create upload directory
        upload_dir = os.path.join(current_app.config['UPLOAD_FOLDER'], invoice_number)
        os.makedirs(upload_dir, exist_ok=True)
        
        # Save files
        saved_files = {}
        if pl_file and pl_file.filename:
            pl_filename = secure_filename(pl_file.filename)
            pl_path = os.path.join(upload_dir, pl_filename)
            pl_file.save(pl_path)
            saved_files['pl_file'] = pl_path
        
        if booking_file and booking_file.filename:
            booking_filename = secure_filename(booking_file.filename)
            booking_path = os.path.join(upload_dir, booking_filename)
            booking_file.save(booking_path)
            saved_files['booking_file'] = booking_path
        
        # Save additional files
        for file in additional_files:
            if file and file.filename:
                filename = secure_filename(file.filename)
                file_path = os.path.join(upload_dir, filename)
                file.save(file_path)
                saved_files[f'additional_{filename}'] = file_path
        
        # Add invoice to database
        if save_to_db:
            invoice_id = add_invoice(invoice_number, 'uploaded')
        else:
            invoice_id = None
        
        # Process immediately or create background job
        if auto_process:
            if invoice_id:
                # Create background job
                job_id = str(uuid.uuid4())
                active_jobs[job_id] = {
                    'status': 'processing',
                    'progress': 0.0,
                    'invoice_id': invoice_id,
                    'invoice_number': invoice_number,
                    'files': saved_files,
                    'options': {
                        'extract_dest': extract_dest,
                        'validate_data': validate_data,
                        'save_to_db': save_to_db
                    },
                    'started_at': datetime.now(),
                    'result': None,
                    'error': None
                }
                
                # Start background processing
                thread = threading.Thread(
                    target=process_files_background,
                    args=(job_id, saved_files, invoice_id, invoice_number, {
                        'extract_dest': extract_dest,
                        'validate_data': validate_data,
                        'save_to_db': save_to_db
                    })
                )
                thread.start()
                
                return jsonify({
                    'success': True,
                    'message': 'Files uploaded and processing started',
                    'job_id': job_id,
                    'invoice_id': invoice_id,
                    'files_saved': list(saved_files.keys())
                })
            else:
                # Process immediately without database
                result = process_pdf_files(
                    pl_file=saved_files.get('pl_file'),
                    booking_file=saved_files.get('booking_file')
                )
                
                return jsonify({
                    'success': True,
                    'message': 'Files processed successfully',
                    'extracted_data': result.get('pl_data', {}),
                    'asi_description': result.get('asi_description', ''),
                    'dest': result.get('dest', ''),
                    'errors': result.get('errors', []),
                    'files_saved': list(saved_files.keys())
                })
        else:
            return jsonify({
                'success': True,
                'message': 'Files uploaded successfully',
                'invoice_id': invoice_id,
                'files_saved': list(saved_files.keys())
            })
            
    except Exception as e:
        current_app.logger.error(f"Upload error: {str(e)}")
        return jsonify({'success': False, 'message': f'Upload failed: {str(e)}'}), 500

def process_files_background(job_id, files, invoice_id, invoice_number, options):
    """Background processing function"""
    try:
        # Update job status
        active_jobs[job_id]['status'] = 'processing'
        active_jobs[job_id]['progress'] = 0.1
        
        # Log start
        if options.get('save_to_db'):
            log_processing(job_id, invoice_number, 'Started PDF processing', 'INFO')
            update_invoice_status(invoice_id, 'processing')
        
        # Process PDF files
        active_jobs[job_id]['progress'] = 0.3
        result = process_pdf_files(
            pl_file=files.get('pl_file'),
            booking_file=files.get('booking_file')
        )
        
        active_jobs[job_id]['progress'] = 0.6
        
        # Save extracted data to database
        if options.get('save_to_db') and invoice_id:
            pl_data = result.get('pl_data', {})
            for field_name, field_value in pl_data.items():
                if field_value:
                    add_extracted_data(invoice_id, field_name, str(field_value), 'pl_file')
            
            # Add ASI description
            if result.get('asi_description'):
                add_extracted_data(invoice_id, 'DESCRIPTION', result['asi_description'], 'booking_file')
            
            # Add destination
            if result.get('dest'):
                add_extracted_data(invoice_id, 'DEST', result['dest'], 'pl_file')
        
        active_jobs[job_id]['progress'] = 0.9
        
        # Prepare result data
        extracted_data = result.get('pl_data', {})
        if result.get('asi_description'):
            extracted_data['DESCRIPTION'] = result['asi_description']
        if result.get('dest'):
            extracted_data['DEST'] = result['dest']
        
        # Update job completion
        active_jobs[job_id]['status'] = 'completed'
        active_jobs[job_id]['progress'] = 1.0
        active_jobs[job_id]['result'] = {
            'extracted_data': extracted_data,
            'errors': result.get('errors', []),
            'invoice_numbers': result.get('invoice_numbers', [])
        }
        
        # Update database
        if options.get('save_to_db') and invoice_id:
            data_json = json.dumps(extracted_data)
            update_invoice_status(invoice_id, 'completed', data_json)
            log_processing(job_id, invoice_number, 'PDF processing completed successfully', 'INFO')
        
    except Exception as e:
        # Handle errors
        active_jobs[job_id]['status'] = 'failed'
        active_jobs[job_id]['error'] = str(e)
        
        if options.get('save_to_db') and invoice_id:
            update_invoice_status(invoice_id, 'failed')
            log_processing(job_id, invoice_number, f'PDF processing failed: {str(e)}', 'ERROR')

@api_bp.route('/job/<job_id>/status')
def get_job_status(job_id):
    """Get job processing status"""
    if job_id not in active_jobs:
        return jsonify({'success': False, 'message': 'Job not found'}), 404
    
    job = active_jobs[job_id]
    
    response = {
        'success': True,
        'job_id': job_id,
        'status': job['status'],
        'progress': job['progress'],
        'started_at': job['started_at'].isoformat(),
    }
    
    if job['status'] == 'completed':
        response.update(job['result'])
        response['invoice_id'] = job.get('invoice_id')
    elif job['status'] == 'failed':
        response['error_message'] = job.get('error', 'Unknown error')
    
    # Add status message
    status_messages = {
        'processing': 'Đang xử lý PDF files...',
        'completed': 'Xử lý hoàn tất!',
        'failed': 'Xử lý thất bại!'
    }
    response['status_message'] = status_messages.get(job['status'], 'Unknown status')
    
    return jsonify(response)

@api_bp.route('/invoices')
def list_invoices():
    """List all invoices via API"""
    try:
        invoices = get_all_invoices()
        return jsonify({
            'success': True,
            'invoices': invoices,
            'count': len(invoices)
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@api_bp.route('/invoices/<invoice_number>')
def get_invoice(invoice_number):
    """Get specific invoice data"""
    try:
        invoice = get_invoice_by_number(invoice_number)
        if not invoice:
            return jsonify({'success': False, 'message': 'Invoice not found'}), 404
        
        return jsonify({
            'success': True,
            'invoice': invoice
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@api_bp.route('/dashboard/stats')
def dashboard_stats():
    """Get dashboard statistics"""
    try:
        invoices = get_all_invoices()
        
        stats = {
            'total_invoices': len(invoices),
            'processed': len([i for i in invoices if i['status'] == 'completed']),
            'pending': len([i for i in invoices if i['status'] in ['pending', 'uploaded']]),
            'processing': len([i for i in invoices if i['status'] == 'processing']),
            'errors': len([i for i in invoices if i['status'] == 'failed'])
        }
        
        return jsonify({
            'success': True,
            'stats': stats
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@api_bp.route('/dashboard/recent-activity')
def recent_activity():
    """Get recent activity for dashboard"""
    try:
        invoices = get_all_invoices()
        
        # Get last 10 invoices
        recent = sorted(invoices, key=lambda x: x['updated_at'] or x['created_at'], reverse=True)[:10]
        
        activities = []
        for invoice in recent:
            status_map = {
                'pending': 'Chờ xử lý',
                'uploaded': 'Đã upload',
                'processing': 'Đang xử lý',
                'completed': 'Hoàn thành',
                'failed': 'Thất bại'
            }
            
            activities.append({
                'invoice_number': invoice['invoice_number'],
                'message': status_map.get(invoice['status'], invoice['status']),
                'timestamp': invoice['updated_at'] or invoice['created_at']
            })
        
        return jsonify({
            'success': True,
            'activities': activities
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@api_bp.route('/process/<invoice_number>', methods=['POST'])
def process_invoice(invoice_number):
    """Process a specific invoice"""
    try:
        invoice = get_invoice_by_number(invoice_number)
        if not invoice:
            return jsonify({'success': False, 'message': 'Invoice not found'}), 404
        
        # Find files for this invoice
        upload_dir = os.path.join(current_app.config['UPLOAD_FOLDER'], invoice_number)
        if not os.path.exists(upload_dir):
            return jsonify({'success': False, 'message': 'No files found for this invoice'}), 404
        
        files = os.listdir(upload_dir)
        pl_file = None
        booking_file = None
        
        for file in files:
            file_path = os.path.join(upload_dir, file)
            if 'pl' in file.lower() and file.endswith('.pdf'):
                pl_file = file_path
            elif 'booking' in file.lower() and file.endswith('.pdf'):
                booking_file = file_path
        
        if not pl_file and not booking_file:
            return jsonify({'success': False, 'message': 'No PDF files found'}), 404
        
        # Process files
        result = process_pdf_files(pl_file=pl_file, booking_file=booking_file)
        
        # Update database
        extracted_data = result.get('pl_data', {})
        if result.get('asi_description'):
            extracted_data['DESCRIPTION'] = result['asi_description']
        if result.get('dest'):
            extracted_data['DEST'] = result['dest']
        
        data_json = json.dumps(extracted_data)
        status = 'completed' if not result.get('errors') else 'failed'
        update_invoice_status(invoice['id'], status, data_json)
        
        return jsonify({
            'success': True,
            'message': 'Invoice processed successfully',
            'extracted_data': extracted_data,
            'errors': result.get('errors', [])
        })
        
    except Exception as e:
        current_app.logger.error(f"Process invoice error: {str(e)}")
        return jsonify({'success': False, 'message': str(e)}), 500