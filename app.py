import os
import json
import threading
import time
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import excel_processor

app = Flask(__name__)

# --- Configure upload and processed file directories using absolute paths ---
# Get the base directory where app.py is located
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
PROCESSED_FOLDER = os.path.join(BASE_DIR, 'processed_files')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    """Checks if the file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# NOTE FOR PRODUCTION DEPLOYMENT:
# The following function and its threading usage are NOT recommended for
# production Flask applications served by WSGI servers like Gunicorn.
# For robust background tasks in production, consider:
# 1. Task Queue: Celery with Redis/RabbitMQ.
# 2. Cron Job: A separate script run periodically by cron.
def delete_files_after_delay(file_paths, delay_seconds=15):
    """
    Deletes a list of files after a specified delay.
    This function runs in a separate thread.
    """
    time.sleep(delay_seconds)
    for filepath in file_paths:
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                # In production, use app.logger.info instead of print
                print(f"Deleted processed file after {delay_seconds}s delay: {filepath}")
            except Exception as e:
                # In production, use app.logger.error instead of print
                print(f"Error deleting processed file {filepath}: {e}")
        else:
            # In production, use app.logger.warning instead of print
            print(f"File not found for delayed deletion: {filepath}")

@app.route('/')
def serve_index():
    """Serves the index.html file."""
    # Ensure static files are served from the correct absolute path
    return send_from_directory(os.path.join(BASE_DIR, 'static'), 'index.html')

@app.route('/process-excel', methods=['POST'])
def process_excel_file():
    """
    Handles the uploaded Excel file, processes it with dynamic keywords and input sheet name,
    returns download links, and schedules the deletion of processed files.
    """
    if 'excelFile' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['excelFile']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    # Get keywords from the form data
    keywords_json = request.form.get('keywords')
    keywords_list = []
    if keywords_json:
        try:
            keywords_list = json.loads(keywords_json)
            if not isinstance(keywords_list, list):
                raise ValueError("Keywords not in expected list format.")
        except json.JSONDecodeError:
            return jsonify({'error': 'Invalid keywords format'}), 400
        except ValueError as e:
            return jsonify({'error': str(e)}), 400
    if not keywords_list: # Default if empty or invalid JSON
        keywords_list = []

    # Get input sheet name from the form data
    input_sheet_name = request.form.get('inputSheetName', 'Sheet1').strip()
    if not input_sheet_name: # Ensure it's not an empty string after strip
        input_sheet_name = 'Sheet1'

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_filepath)

        cleaned_output_filename = f"cleaned_{filename}"
        excluded_output_filename = f"excluded_{filename}"
        cleaned_output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], cleaned_output_filename)
        excluded_output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], excluded_output_filename)

        try:
            # Pass the keywords_list AND input_sheet_name to the data processing function
            excel_processor.process_data_excel(
                input_filepath,
                cleaned_output_filepath,
                excluded_output_filepath,
                keywords_list,
                input_sheet_name
            )

            deletion_files = [cleaned_output_filepath, excluded_output_filepath]
            # Consider removing threading.Thread for production as discussed
            deleter_thread = threading.Thread(target=delete_files_after_delay, args=(deletion_files, 15))
            deleter_thread.start()

            return jsonify({
                'message': 'File processed successfully',
                'cleaned_url': f'/downloads/{cleaned_output_filename}',
                'excluded_url': f'/downloads/{excluded_output_filename}'
            }), 200

        except Exception as e:
            # Use app.logger.error for production
            print(f"Error during processing: {e}")
            return jsonify({'error': f'File processing failed: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Invalid file type. Only .xlsx files are allowed.'}), 400

@app.route('/downloads/<filename>')
def download_file(filename):
    """Serves the processed files for download."""
    # Ensure processed files are served from the correct absolute path
    return send_from_directory(os.path.join(BASE_DIR, 'processed_files'), filename, as_attachment=True)

if __name__ == '__main__':
    # For local development ONLY. Gunicorn will ignore this.
    app.run(debug=True, host='0.0.0.0', port=6967)
