# -*- coding: utf-8 -*-
"""
Nike Export Web App
Main Flask application for processing Nike export documents
"""

import os
import logging
from flask import Flask, render_template, request, jsonify, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from datetime import datetime
import sqlite3

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def create_app():
    app = Flask(__name__)
    app.config.from_object('config.Config')
    
    # Ensure upload directories exist
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    os.makedirs(app.config['TEMPLATE_FOLDER'], exist_ok=True)
    
    # Initialize database
    from database import init_db
    init_db()
    
    # Register blueprints
    from routes.main import main_bp
    from routes.api import api_bp
    from routes.templates import templates_bp
    
    app.register_blueprint(main_bp)
    app.register_blueprint(api_bp, url_prefix='/api')
    app.register_blueprint(templates_bp, url_prefix='/templates')
    
    @app.errorhandler(404)
    def not_found_error(error):
        return render_template('errors/404.html'), 404
    
    @app.errorhandler(500)
    def internal_error(error):
        return render_template('errors/500.html'), 500
    
    @app.context_processor
    def inject_builtin_functions():
        return dict(chr=chr, ord=ord, len=len)
    
    return app

if __name__ == '__main__':
    import os
    app = create_app()
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)