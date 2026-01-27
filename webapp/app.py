"""
Translation Processor Web Application
Flask-based web server for processing translation documents
"""
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session
from werkzeug.utils import secure_filename
from datetime import datetime
from pathlib import Path
import os
import json
import secrets

# Import utilities
from utils.translator_manager import TranslatorManager
from utils.document_processor import DocumentProcessor

# Initialize Flask app
app = Flask(__name__)
app.secret_key = secrets.token_hex(16)  # For session and flash messages

# Configuration
UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
OUTPUT_FOLDER = Path(__file__).parent / 'output'
DATA_FOLDER = Path(__file__).parent / 'data'
ASSETS_FOLDER = Path(__file__).parent / 'assets'
ALLOWED_EXTENSIONS = {'docx'}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB

# Ensure directories exist
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)
DATA_FOLDER.mkdir(exist_ok=True)
(DATA_FOLDER / 'signatures').mkdir(exist_ok=True)
ASSETS_FOLDER.mkdir(exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Initialize managers
translator_manager = TranslatorManager(data_dir=DATA_FOLDER)
document_processor = DocumentProcessor()

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_common_languages():
    """Return list of common languages"""
    return [
        "English", "Spanish", "French", "German", "Italian",
        "Portuguese", "Chinese", "Japanese", "Korean", "Arabic",
        "Russian", "Hindi", "Bengali", "Polish", "Dutch",
        "Turkish", "Vietnamese", "Thai", "Greek", "Hebrew"
    ]

@app.route('/')
def index():
    """Main page - translation processor"""
    translators = translator_manager.get_translator_names()
    languages = get_common_languages()

    return render_template(
        'index.html',
        translators=translators,
        languages=languages,
        current_date=datetime.now().strftime("%B %d, %Y")
    )

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload"""

    # Check if file was uploaded
    if 'document' not in request.files:
        flash('No file uploaded', 'error')
        return redirect(url_for('index'))

    file = request.files['document']

    # Check if file was selected
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))

    # Validate file
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload a .docx file', 'error')
        return redirect(url_for('index'))

    # Save file
    filename = secure_filename(file.filename)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    unique_filename = f"{timestamp}_{filename}"
    filepath = app.config['UPLOAD_FOLDER'] / unique_filename

    try:
        file.save(filepath)

        # Store filename in session
        session['uploaded_file'] = unique_filename
        session['original_filename'] = filename

        flash(f'File uploaded successfully: {filename}', 'success')
        return redirect(url_for('index'))

    except Exception as e:
        flash(f'Error uploading file: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/process', methods=['POST'])
def process_document():
    """Process the translation document"""

    # Get uploaded file from session
    uploaded_file = session.get('uploaded_file')

    if not uploaded_file:
        flash('No file uploaded. Please upload a document first.', 'error')
        return redirect(url_for('index'))

    # Get form data
    case_number = request.form.get('case_number', '').strip()
    source_language = request.form.get('source_language', '').strip()
    target_language = request.form.get('target_language', '').strip()
    translator_name = request.form.get('translator', '').strip()
    date = request.form.get('date', '').strip()

    # Validate inputs
    errors = []
    if not case_number:
        errors.append('Case number is required')
    if not source_language:
        errors.append('Source language is required')
    if not target_language:
        errors.append('Target language is required')
    if not translator_name:
        errors.append('Translator is required')
    if not date:
        errors.append('Date is required')

    if errors:
        for error in errors:
            flash(error, 'error')
        return redirect(url_for('index'))

    # Get translator info
    translator = translator_manager.get_translator(translator_name)

    if not translator:
        flash('Selected translator not found', 'error')
        return redirect(url_for('index'))

    # Prepare metadata
    metadata = {
        'case_number': case_number,
        'source_language': source_language,
        'target_language': target_language,
        'language_pair': f"{source_language} > {target_language} Translation",
        'translator_name': translator_name,
        'translator_signature': translator.get('signature_path', ''),
        'date': date,
    }

    # Process document
    input_path = app.config['UPLOAD_FOLDER'] / uploaded_file
    output_filename = f"Park_Case_{case_number}_Certified.pdf"
    output_path = app.config['OUTPUT_FOLDER'] / output_filename

    try:
        success = document_processor.process_translation(
            input_path,
            output_path,
            metadata
        )

        if success:
            # Clean up uploaded file
            try:
                input_path.unlink()
            except:
                pass

            # Clear session
            session.pop('uploaded_file', None)
            session.pop('original_filename', None)

            # Store output file in session for download
            session['output_file'] = output_filename

            flash('Translation processed successfully!', 'success')
            return redirect(url_for('download'))
        else:
            flash('Failed to process document. Check server logs for details.', 'error')
            return redirect(url_for('index'))

    except Exception as e:
        flash(f'Error processing document: {str(e)}', 'error')
        import traceback
        traceback.print_exc()
        return redirect(url_for('index'))

@app.route('/download')
def download():
    """Download page"""
    output_file = session.get('output_file')

    if not output_file:
        flash('No processed document available', 'error')
        return redirect(url_for('index'))

    return render_template('download.html', filename=output_file)

@app.route('/download/<filename>')
def download_file(filename):
    """Serve the processed PDF for download"""

    # Security: only allow downloading from output folder
    safe_filename = secure_filename(filename)
    filepath = app.config['OUTPUT_FOLDER'] / safe_filename

    if not filepath.exists():
        flash('File not found', 'error')
        return redirect(url_for('index'))

    return send_file(
        filepath,
        as_attachment=True,
        download_name=safe_filename
    )

@app.route('/translators')
def translators():
    """Translator management page"""
    all_translators = translator_manager.get_all_translators()
    languages = get_common_languages()

    return render_template(
        'translators.html',
        translators=all_translators,
        languages=languages
    )

@app.route('/translators/add', methods=['POST'])
def add_translator():
    """Add new translator"""

    name = request.form.get('name', '').strip()
    language_pairs = request.form.get('language_pairs', '').strip()

    if not name:
        flash('Translator name is required', 'error')
        return redirect(url_for('translators'))

    if not language_pairs:
        flash('At least one language pair is required', 'error')
        return redirect(url_for('translators'))

    # Parse language pairs (one per line)
    pairs = [line.strip() for line in language_pairs.split('\n') if line.strip()]

    # Handle signature upload
    signature_path = None
    if 'signature' in request.files:
        signature_file = request.files['signature']
        if signature_file and signature_file.filename:
            # Save to temporary upload folder first
            original_filename = secure_filename(signature_file.filename)
            file_ext = os.path.splitext(original_filename)[1]
            temp_filename = f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}_{name.replace(' ', '_')}{file_ext}"
            temp_path = UPLOAD_FOLDER / temp_filename
            signature_file.save(temp_path)
            signature_path = str(temp_path)

    # Add translator
    success, message = translator_manager.add_translator(name, pairs, signature_path)

    # Clean up temporary file
    if signature_path and Path(signature_path).exists():
        try:
            Path(signature_path).unlink()
        except:
            pass

    if success:
        flash(message, 'success')
    else:
        flash(message, 'error')

    return redirect(url_for('translators'))

@app.route('/translators/update/<old_name>', methods=['POST'])
def update_translator(old_name):
    """Update existing translator"""

    new_name = request.form.get('name', '').strip()
    language_pairs = request.form.get('language_pairs', '').strip()

    if not new_name:
        flash('Translator name is required', 'error')
        return redirect(url_for('translators'))

    if not language_pairs:
        flash('At least one language pair is required', 'error')
        return redirect(url_for('translators'))

    # Parse language pairs (one per line)
    pairs = [line.strip() for line in language_pairs.split('\n') if line.strip()]

    # Handle signature upload
    signature_path = None
    if 'signature' in request.files:
        signature_file = request.files['signature']
        if signature_file and signature_file.filename:
            # Save to temporary upload folder first
            original_filename = secure_filename(signature_file.filename)
            file_ext = os.path.splitext(original_filename)[1]
            temp_filename = f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}_{new_name.replace(' ', '_')}{file_ext}"
            temp_path = UPLOAD_FOLDER / temp_filename
            signature_file.save(temp_path)
            signature_path = str(temp_path)

    # Update translator
    success, message = translator_manager.update_translator(old_name, new_name, pairs, signature_path)

    # Clean up temporary file
    if signature_path and Path(signature_path).exists():
        try:
            Path(signature_path).unlink()
        except:
            pass

    if success:
        flash(message, 'success')
    else:
        flash(message, 'error')

    return redirect(url_for('translators'))

@app.route('/translators/delete/<name>', methods=['POST'])
def delete_translator(name):
    """Delete translator"""

    success, message = translator_manager.delete_translator(name)

    if success:
        flash(message, 'success')
    else:
        flash(message, 'error')

    return redirect(url_for('translators'))

@app.route('/api/translators')
def api_translators():
    """API endpoint to get translators as JSON"""
    translators = translator_manager.get_all_translators()
    return jsonify(translators)

@app.route('/clear', methods=['POST'])
def clear():
    """Clear uploaded file and form"""

    # Clean up uploaded file if exists
    uploaded_file = session.get('uploaded_file')
    if uploaded_file:
        filepath = app.config['UPLOAD_FOLDER'] / uploaded_file
        try:
            if filepath.exists():
                filepath.unlink()
        except:
            pass

    # Clear session
    session.pop('uploaded_file', None)
    session.pop('original_filename', None)
    session.pop('output_file', None)

    flash('Form cleared', 'info')
    return redirect(url_for('index'))

@app.errorhandler(413)
def file_too_large(e):
    """Handle file size limit exceeded"""
    flash('File is too large. Maximum size is 50 MB.', 'error')
    return redirect(url_for('index'))

@app.errorhandler(404)
def not_found(e):
    """Handle 404 errors"""
    flash('Page not found', 'error')
    return redirect(url_for('index'))

@app.errorhandler(500)
def internal_error(e):
    """Handle 500 errors"""
    flash('An internal error occurred. Please try again.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Run on all network interfaces so it's accessible from other machines
    # Your IT team can configure this properly when deploying
    app.run(
        host='0.0.0.0',
        port=5000,
        debug=True  # Set to False in production
    )
