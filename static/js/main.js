// Main JavaScript for Nike Export Web App

// Global variables
window.NikeExport = {
    apiBaseUrl: '/api',
    currentUser: null,
    settings: {
        autoRefresh: true,
        refreshInterval: 30000 // 30 seconds
    }
};

// Utility functions
const Utils = {
    // Format file size
    formatFileSize: function(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    },

    // Format date
    formatDate: function(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('vi-VN') + ' ' + date.toLocaleTimeString('vi-VN');
    },

    // Show notification
    showNotification: function(message, type = 'info') {
        const alertClass = type === 'error' ? 'alert-danger' : `alert-${type}`;
        const alert = document.createElement('div');
        alert.className = `alert ${alertClass} alert-dismissible fade show position-fixed`;
        alert.style.cssText = 'top: 20px; right: 20px; z-index: 9999; min-width: 300px;';
        alert.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        
        document.body.appendChild(alert);
        
        // Auto dismiss after 5 seconds
        setTimeout(() => {
            if (alert.parentNode) {
                alert.remove();
            }
        }, 5000);
    },

    // Copy text to clipboard
    copyToClipboard: function(text) {
        navigator.clipboard.writeText(text).then(() => {
            this.showNotification('Đã copy vào clipboard', 'success');
        }).catch(() => {
            this.showNotification('Không thể copy', 'error');
        });
    },

    // Validate file type
    validateFileType: function(file, allowedTypes) {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        return allowedTypes.includes(fileExtension);
    },

    // Generate unique ID
    generateId: function() {
        return Date.now().toString(36) + Math.random().toString(36).substr(2);
    }
};

// API Helper
const API = {
    // Base fetch function
    fetch: async function(endpoint, options = {}) {
        const url = `${window.NikeExport.apiBaseUrl}${endpoint}`;
        const defaultOptions = {
            headers: {
                'Content-Type': 'application/json',
            },
        };

        // Don't set Content-Type for FormData
        if (options.body instanceof FormData) {
            delete defaultOptions.headers['Content-Type'];
        }

        try {
            const response = await fetch(url, { ...defaultOptions, ...options });
            const data = await response.json();
            
            if (!response.ok) {
                throw new Error(data.message || `HTTP error! status: ${response.status}`);
            }
            
            return data;
        } catch (error) {
            console.error('API Error:', error);
            throw error;
        }
    },

    // GET request
    get: function(endpoint) {
        return this.fetch(endpoint);
    },

    // POST request
    post: function(endpoint, data) {
        return this.fetch(endpoint, {
            method: 'POST',
            body: data instanceof FormData ? data : JSON.stringify(data),
        });
    },

    // PUT request
    put: function(endpoint, data) {
        return this.fetch(endpoint, {
            method: 'PUT',
            body: JSON.stringify(data),
        });
    },

    // DELETE request
    delete: function(endpoint) {
        return this.fetch(endpoint, {
            method: 'DELETE',
        });
    }
};

// File Upload Handler
const FileUpload = {
    // Setup drag and drop
    setupDragAndDrop: function(element, callback) {
        element.addEventListener('dragover', function(e) {
            e.preventDefault();
            e.stopPropagation();
            element.classList.add('dragover');
        });

        element.addEventListener('dragleave', function(e) {
            e.preventDefault();
            e.stopPropagation();
            element.classList.remove('dragover');
        });

        element.addEventListener('drop', function(e) {
            e.preventDefault();
            e.stopPropagation();
            element.classList.remove('dragover');
            
            const files = Array.from(e.dataTransfer.files);
            if (callback) callback(files);
        });
    },

    // Upload files with progress
    uploadWithProgress: function(files, endpoint, onProgress, onComplete) {
        const formData = new FormData();
        files.forEach((file, index) => {
            formData.append(`file_${index}`, file);
        });

        const xhr = new XMLHttpRequest();
        
        xhr.upload.addEventListener('progress', function(e) {
            if (e.lengthComputable) {
                const percentComplete = (e.loaded / e.total) * 100;
                if (onProgress) onProgress(percentComplete);
            }
        });

        xhr.addEventListener('load', function() {
            if (xhr.status === 200) {
                try {
                    const response = JSON.parse(xhr.responseText);
                    if (onComplete) onComplete(null, response);
                } catch (error) {
                    if (onComplete) onComplete(error);
                }
            } else {
                if (onComplete) onComplete(new Error(`Upload failed: ${xhr.status}`));
            }
        });

        xhr.addEventListener('error', function() {
            if (onComplete) onComplete(new Error('Upload failed'));
        });

        xhr.open('POST', `${window.NikeExport.apiBaseUrl}${endpoint}`);
        xhr.send(formData);
    }
};

// Progress Bar Handler
const ProgressBar = {
    update: function(elementId, percent, text = '') {
        const progressBar = document.getElementById(elementId);
        if (progressBar) {
            progressBar.style.width = `${percent}%`;
            progressBar.setAttribute('aria-valuenow', percent);
            
            if (text) {
                const textElement = progressBar.querySelector('.progress-text') || 
                                 progressBar.parentNode.querySelector('.progress-text');
                if (textElement) {
                    textElement.textContent = text;
                }
            }
        }
    },

    show: function(containerId) {
        const container = document.getElementById(containerId);
        if (container) {
            container.classList.remove('d-none');
        }
    },

    hide: function(containerId) {
        const container = document.getElementById(containerId);
        if (container) {
            container.classList.add('d-none');
        }
    }
};

// Table Handler
const TableHandler = {
    // Sort table
    sortTable: function(table, columnIndex, ascending = true) {
        const tbody = table.querySelector('tbody');
        const rows = Array.from(tbody.querySelectorAll('tr'));
        
        rows.sort((a, b) => {
            const aVal = a.cells[columnIndex].textContent.trim();
            const bVal = b.cells[columnIndex].textContent.trim();
            
            // Try to parse as numbers
            const aNum = parseFloat(aVal);
            const bNum = parseFloat(bVal);
            
            if (!isNaN(aNum) && !isNaN(bNum)) {
                return ascending ? aNum - bNum : bNum - aNum;
            }
            
            // String comparison
            return ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        });
        
        // Re-append sorted rows
        rows.forEach(row => tbody.appendChild(row));
    },

    // Filter table
    filterTable: function(table, searchTerm) {
        const tbody = table.querySelector('tbody');
        const rows = tbody.querySelectorAll('tr');
        
        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            const shouldShow = text.includes(searchTerm.toLowerCase());
            row.style.display = shouldShow ? '' : 'none';
        });
    }
};

// Form Validation
const FormValidator = {
    // Validate required fields
    validateRequired: function(form) {
        const requiredFields = form.querySelectorAll('[required]');
        let isValid = true;
        
        requiredFields.forEach(field => {
            if (!field.value.trim()) {
                this.showFieldError(field, 'Trường này là bắt buộc');
                isValid = false;
            } else {
                this.clearFieldError(field);
            }
        });
        
        return isValid;
    },

    // Show field error
    showFieldError: function(field, message) {
        field.classList.add('is-invalid');
        
        let errorDiv = field.parentNode.querySelector('.invalid-feedback');
        if (!errorDiv) {
            errorDiv = document.createElement('div');
            errorDiv.className = 'invalid-feedback';
            field.parentNode.appendChild(errorDiv);
        }
        errorDiv.textContent = message;
    },

    // Clear field error
    clearFieldError: function(field) {
        field.classList.remove('is-invalid');
        const errorDiv = field.parentNode.querySelector('.invalid-feedback');
        if (errorDiv) {
            errorDiv.remove();
        }
    }
};

// Initialize app
document.addEventListener('DOMContentLoaded', function() {
    // Initialize tooltips
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function(tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });

    // Initialize popovers
    const popoverTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="popover"]'));
    popoverTriggerList.map(function(popoverTriggerEl) {
        return new bootstrap.Popover(popoverTriggerEl);
    });

    // Setup global error handling
    window.addEventListener('unhandledrejection', function(event) {
        console.error('Unhandled promise rejection:', event.reason);
        Utils.showNotification('Đã có lỗi xảy ra. Vui lòng thử lại.', 'error');
    });

    // Auto-refresh data if enabled
    if (window.NikeExport.settings.autoRefresh) {
        setInterval(function() {
            // Refresh dashboard stats if on dashboard
            if (window.location.pathname === '/' && typeof loadDashboardStats === 'function') {
                loadDashboardStats();
            }
        }, window.NikeExport.settings.refreshInterval);
    }
});

// Export utilities to global scope
window.Utils = Utils;
window.API = API;
window.FileUpload = FileUpload;
window.ProgressBar = ProgressBar;
window.TableHandler = TableHandler;
window.FormValidator = FormValidator;