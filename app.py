"""
CO-PO Attainment Sheet Generator
Flask Web Application
"""
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from werkzeug.utils import secure_filename
import os
from pathlib import Path
import uuid
from datetime import datetime

from utils.template_mapper import TemplateMapper
from utils.data_parser import DataParser
from utils.validator import Validator
from utils.excel_handler import ExcelHandler

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'co-po-attainment-secret-key-2026')

# Configuration
BASE_DIR = Path(__file__).parent
UPLOAD_FOLDER = BASE_DIR / 'uploads'
OUTPUT_FOLDER = BASE_DIR / 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Ensure directories exist
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
app.config['OUTPUT_FOLDER'] = str(OUTPUT_FOLDER)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Initialize utils
template_mapper = TemplateMapper(BASE_DIR)
parser = DataParser()
validator = Validator()
excel_handler = ExcelHandler(BASE_DIR)


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def cleanup_old_files(folder, max_age_hours=24):
    """Remove files older than max_age_hours"""
    import time
    current_time = time.time()
    for file_path in Path(folder).glob('*'):
        if file_path.is_file():
            file_age = current_time - file_path.stat().st_mtime
            if file_age > (max_age_hours * 3600):
                try:
                    file_path.unlink()
                except:
                    pass


@app.route('/')
def index():
    """Home page with upload form"""
    regulations = template_mapper.get_available_regulations()
    return render_template('index.html', regulations=regulations)


@app.route('/api/categories/<regulation>')
def get_categories(regulation):
    """API endpoint to get categories for a regulation"""
    categories = template_mapper.get_available_categories(regulation)
    return jsonify({'categories': categories})


@app.route('/api/dept_types/<regulation>/<category>')
def get_dept_types(regulation, category):
    """API endpoint to get department types for a regulation and category"""
    dept_types = template_mapper.get_available_dept_types(regulation, category)
    return jsonify({'dept_types': dept_types})


@app.route('/api/required_inputs/<regulation>/<category>')
def get_required_inputs(regulation, category):
    """API endpoint to get required input files"""
    try:
        inputs = template_mapper.get_required_inputs(regulation, category)
        return jsonify({'inputs': inputs})
    except ValueError as e:
        return jsonify({'error': str(e)}), 400


@app.route('/generate', methods=['POST'])
def generate():
    """Generate attainment sheet from uploaded files"""
    try:
        # Get form data
        regulation = request.form.get('regulation', 'R17')
        category = request.form.get('category', 'theory')
        dept_type = request.form.get('dept_type', 'dept')
        
        # Get required inputs for this category
        required_inputs = template_mapper.get_required_inputs(regulation, category)
        
        # Create unique session folder for uploads
        session_id = str(uuid.uuid4())[:8]
        session_folder = UPLOAD_FOLDER / session_id
        session_folder.mkdir(exist_ok=True)
        
        # Process uploaded files
        eval_files = {}
        uploaded_paths = []
        
        for input_type in required_inputs:
            file_key = f'file_{input_type.lower()}'
            
            if file_key not in request.files:
                flash(f'Missing file for {input_type}', 'error')
                return redirect(url_for('index'))
            
            file = request.files[file_key]
            
            if file.filename == '':
                flash(f'No file selected for {input_type}', 'error')
                return redirect(url_for('index'))
            
            if not allowed_file(file.filename):
                flash(f'Invalid file type for {input_type}. Only .xlsx and .xls allowed.', 'error')
                return redirect(url_for('index'))
            
            # Save file
            filename = secure_filename(f"{input_type}_{file.filename}")
            file_path = session_folder / filename
            file.save(str(file_path))
            
            eval_files[input_type] = str(file_path)
            uploaded_paths.append(str(file_path))
        
        # Validate consistency
        validation_result = validator.validate_all(uploaded_paths, regulation)
        
        if not validation_result.is_valid:
            flash('Validation failed: ' + '; '.join(validation_result.errors), 'error')
            return redirect(url_for('index'))
        
        # Show warnings if any
        if validation_result.warnings:
            for warning in validation_result.warnings[:5]:  # Show first 5 warnings
                flash(f'Warning: {warning}', 'warning')
        
        # Extract course info for output filename
        first_file = list(eval_files.values())[0]
        course_info = parser.extract_validation_fields(first_file)
        
        course_code = course_info.get('course_code', 'UNKNOWN')
        course_name = course_info.get('course_name', 'Course')
        
        # Clean course name for filename
        safe_course_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '' for c in course_name)
        safe_course_name = safe_course_name[:50]  # Limit length
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"{course_code}_{safe_course_name}_{regulation}_Attainment_{timestamp}.xlsx"
        output_path = OUTPUT_FOLDER / output_filename
        
        # Generate attainment sheet
        result = excel_handler.generate_attainment_sheet(
            regulation=regulation,
            category=category,
            dept_type=dept_type,
            eval_files=eval_files,
            output_path=str(output_path),
            course_info=course_info
        )
        
        if not result['success']:
            flash(f'Generation failed: {result.get("error", "Unknown error")}', 'error')
            return redirect(url_for('index'))
        
        # Cleanup session files
        for path in uploaded_paths:
            try:
                Path(path).unlink()
            except:
                pass
        try:
            session_folder.rmdir()
        except:
            pass
        
        flash(f'Successfully generated attainment sheet with {result["students_count"]} students!', 'success')
        
        return render_template('result.html', 
                             filename=output_filename,
                             course_code=course_code,
                             course_name=course_name,
                             students_count=result['students_count'],
                             regulation=regulation,
                             category=category)
    
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/download/<filename>')
def download(filename):
    """Download generated attainment sheet"""
    file_path = OUTPUT_FOLDER / secure_filename(filename)
    
    if not file_path.exists():
        flash('File not found. It may have been deleted.', 'error')
        return redirect(url_for('index'))
    
    return send_file(
        str(file_path),
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/validate', methods=['POST'])
def api_validate():
    """API endpoint to validate uploaded files before generation"""
    try:
        regulation = request.form.get('regulation', 'R17')
        
        # Get all uploaded files
        uploaded_paths = []
        session_id = str(uuid.uuid4())[:8]
        session_folder = UPLOAD_FOLDER / session_id
        session_folder.mkdir(exist_ok=True)
        
        for key, file in request.files.items():
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = session_folder / filename
                file.save(str(file_path))
                uploaded_paths.append(str(file_path))
        
        if not uploaded_paths:
            return jsonify({'valid': False, 'errors': ['No valid files uploaded']})
        
        # Validate
        result = validator.validate_all(uploaded_paths, regulation)
        
        # Cleanup
        for path in uploaded_paths:
            try:
                Path(path).unlink()
            except:
                pass
        try:
            session_folder.rmdir()
        except:
            pass
        
        return jsonify({
            'valid': result.is_valid,
            'errors': result.errors,
            'warnings': result.warnings
        })
    
    except Exception as e:
        return jsonify({'valid': False, 'errors': [str(e)]})


@app.errorhandler(413)
def too_large(e):
    """Handle file too large error"""
    flash('File too large. Maximum size is 16MB.', 'error')
    return redirect(url_for('index'))


@app.errorhandler(500)
def server_error(e):
    """Handle server error"""
    flash('An unexpected error occurred. Please try again.', 'error')
    return redirect(url_for('index'))


# Cleanup old files on startup
cleanup_old_files(UPLOAD_FOLDER)
cleanup_old_files(OUTPUT_FOLDER)


if __name__ == '__main__':
    print("=" * 60)
    print("CO-PO Attainment Sheet Generator")
    print("=" * 60)
    print(f"Upload folder: {UPLOAD_FOLDER}")
    print(f"Output folder: {OUTPUT_FOLDER}")
    print(f"Available regulations: {template_mapper.get_available_regulations()}")
    print("=" * 60)
    print("Starting server at http://localhost:5000")
    print("=" * 60)
    
    app.run(debug=True, host='0.0.0.0', port=5000)
