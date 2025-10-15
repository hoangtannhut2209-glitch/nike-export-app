# -*- coding: utf-8 -*-
"""
Main routes for Nike Export C/O
"""

from flask import Blueprint, render_template, request, jsonify, redirect, url_for, flash
from database import get_all_invoices, get_invoice_by_number
import os

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    """Home page with dashboard"""
    return render_template('index.html')

@main_bp.route('/upload')
def upload():
    """Upload page for PDF files"""
    return render_template('upload.html')

@main_bp.route('/invoices')
def invoices():
    """List all invoices"""
    invoices_list = get_all_invoices()
    return render_template('invoices.html', invoices=invoices_list)

@main_bp.route('/invoices/<int:invoice_id>')
def invoice_detail(invoice_id):
    """Invoice detail page"""
    # Get invoice by ID (would need to modify database.py to support this)
    return render_template('invoice_detail.html', invoice_id=invoice_id)

@main_bp.route('/about')
def about():
    """About page"""
    return render_template('about.html')
