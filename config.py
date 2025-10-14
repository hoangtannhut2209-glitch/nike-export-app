import os
from datetime import timedelta

class Config:
    """Configuration class for Flask app"""
    
    # Flask settings
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'nike-export-secret-key-2025'
    DEBUG = os.environ.get('DEBUG', 'True').lower() == 'true'
    
    # Upload settings
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max file size
    UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
    OUTPUT_FOLDER = os.path.join(os.getcwd(), 'outputs')
    TEMPLATE_FOLDER = os.path.join(os.getcwd(), 'excel_templates')
    
    # Allowed file extensions
    ALLOWED_EXTENSIONS = {
        'pdf': ['pdf'],
        'excel': ['xlsx', 'xlsm', 'xls'],
        'image': ['png', 'jpg', 'jpeg']
    }
    
    # Database
    DATABASE_URL = os.environ.get('DATABASE_URL') or 'sqlite:///nike_export.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # Session settings
    PERMANENT_SESSION_LIFETIME = timedelta(hours=2)
    
    # Processing settings
    PDF_PROCESSING_TIMEOUT = 300  # 5 minutes
    MAX_INVOICES_PER_BATCH = 100
    
    # Template settings
    DEFAULT_TEMPLATE_SHEETS = [
        'CO-FORM A',
        'CO-FORM E', 
        'CO-FORM EVFTA',
        'CO-FORM CPTPP',
        'CO-FORM RCEP'
    ]
    
    # Export modes
    EXPORT_MODES = {
        'EXCEL': 'Export to Excel files',
        'PDF': 'Export to PDF files',
        'BOTH': 'Export to both Excel and PDF',
        'PREVIEW': 'Preview only (no download)'
    }

class DevelopmentConfig(Config):
    """Development configuration"""
    DEBUG = True
    
class ProductionConfig(Config):
    """Production configuration"""
    DEBUG = False
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'production-secret-key-change-in-production'

# Configuration mapping
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}