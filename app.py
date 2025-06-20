import os
import json
import logging # Import logging
import time # Still used for time.sleep if needed, but not for threading.Thread
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import excel_processor

app = Flask(__name__)

# --- Configure logging ---
# This basic configuration will send logs to stderr, which Gunicorn will capture
# and which will then be accessible via journalctl.
app.logger.setLevel(logging.INFO)
handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
app.logger.addHandler(handler)

# --- Configure upload and processed file directories using absolute paths ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
PROCESSED_FOLDER = os.path.join(BASE_DIR, 'processed_files')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
app.logger.info(f"Uploads directory: {UPLOAD_FOLDER}")
app.logger.info(f"Processed files directory: {PROCESSED_FOLDER}")


ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    """Checks if the file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Removed delete_files_after_delay function and threading.Thread usage.
# For robust background file deletion in production, a dedicated task queue (e.g., Celery)
# should be used, not direct threading within the Flask app under Gunicorn.
# Files will now remain in the 'processed_files' directory until manually cleared or
# until a proper background task system is implemented.

@app.route('/')
def serve_index():
    """Serves the index.html file from the static directory."""
    return send_from_directory(os.path.join(BASE_DIR, 'static'), 'index.html')

@app.route('/process-excel', methods=['POST'])
def process_excel_file():
    """
    Handles the uploaded Excel file, processes it with dynamic keywords and input sheet name,
    and returns download links.
    """
    if 'excelFile' not in request.files:
        app.logger.warning('No file part in request')
        return jsonify({'error': 'No file part'}), 400

    file = request.files['excelFile']

    if file.filename == '':
        app.logger.warning('No selected file')
        return jsonify({'error': 'No selected file'}), 400

    # Get keywords from the form data
    keywords_json = request.form.get('keywords')
    keywords_list = []
    if keywords_json:
        try:
            keywords_list = json.loads(keywords_json)
            if not isinstance(keywords_list, list):
                raise ValueError("Keywords not in expected list format.")
            app.logger.info(f"Received keywords: {keywords_list}")
        except json.JSONDecodeError:
            app.logger.error('Invalid keywords JSON format')
            return jsonify({'error': 'Invalid keywords format'}), 400
        except ValueError as e:
            app.logger.error(f'Keywords format error: {e}')
            return jsonify({'error': str(e)}), 400
    if not keywords_list:
        app.logger.info("No keywords provided, proceeding without specific keyword filtering.")
        keywords_list = []

    # Get input sheet name from the form data
    input_sheet_name = request.form.get('inputSheetName', 'Sheet1').strip()
    if not input_sheet_name:
        input_sheet_name = 'Sheet1'
    app.logger.info(f"Using input sheet name: '{input_sheet_name}'")

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        try:
            file.save(input_filepath)
            app.logger.info(f"File saved to: {input_filepath}")
        except Exception as e:
            app.logger.error(f"Failed to save uploaded file {filename}: {e}")
            return jsonify({'error': f'Failed to save uploaded file: {str(e)}'}), 500

        cleaned_output_filename = f"cleaned_{filename}"
        excluded_output_filename = f"excluded_{filename}"
        cleaned_output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], cleaned_output_filename)
        excluded_output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], excluded_output_filename)

        try:
            app.logger.info(f"Starting Excel processing for {filename}...")
            excel_processor.process_data_excel(
                input_filepath,
                cleaned_output_filepath,
                excluded_output_filepath,
                keywords_list,
                input_sheet_name
            )
            app.logger.info(f"Finished Excel processing for {filename}.")

            # Files are no longer automatically deleted by this Flask app.
            # Implement a dedicated task queue for background deletion if needed.
            
            return jsonify({
                'message': 'File processed successfully',
                'cleaned_url': f'/downloads/{cleaned_output_filename}',
                'excluded_url': f'/downloads/{excluded_output_filename}'
            }), 200

        except Exception as e:
            app.logger.error(f"Error during processing of {filename}: {e}", exc_info=True) # exc_info to log full traceback
            return jsonify({'error': f'File processing failed: {str(e)}'}), 500
    else:
        app.logger.warning(f"Invalid file type uploaded: {file.filename}")
        return jsonify({'error': 'Invalid file type. Only .xlsx files are allowed.'}), 400

@app.route('/downloads/<filename>')
def download_file(filename):
    """Serves the processed files for download."""
    # Ensure processed files are served from the correct absolute path
    app.logger.info(f"Serving download for: {filename}")
    return send_from_directory(os.path.join(BASE_DIR, 'processed_files'), filename, as_attachment=True)

# The if __name__ == '__main__': block is for local development only and is ignored by Gunicorn.
# No changes are needed or made here for production deployment.
