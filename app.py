import os
import json # Import json to parse the keywords string
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import excel_processor # Import your data processing script

app = Flask(__name__)

# Configure upload and processed file directories
# Ensure these directories exist and the server has write permissions
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed_files'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Create the directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Allowed extensions for uploaded files
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    """Checks if the file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def serve_index():
    """Serves the index.html file."""
    # This assumes index.html is directly in the static folder
    return send_from_directory('static', 'index.html')

@app.route('/process-excel', methods=['POST'])
def process_excel_file():
    """
    Handles the uploaded Excel file, processes it with dynamic keywords,
    and returns download links.
    """
    if 'excelFile' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['excelFile']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    # Get keywords from the form data
    keywords_json = request.form.get('keywords')
    if not keywords_json:
        # Fallback to default if no keywords are provided (should be handled by frontend)
        keywords_list = ['gopay', 'dijual','']
    else:
        try:
            keywords_list = json.loads(keywords_json)
            if not isinstance(keywords_list, list):
                raise ValueError("Keywords not in expected list format.")
        except json.JSONDecodeError:
            return jsonify({'error': 'Invalid keywords format'}), 400
        except ValueError as e:
            return jsonify({'error': str(e)}), 400


    if file and allowed_file(file.filename):
        # Secure the filename to prevent directory traversal attacks
        filename = secure_filename(file.filename)
        input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_filepath)

        # Define output file paths for cleaned and excluded data
        cleaned_output_filename = f"cleaned_{filename}"
        excluded_output_filename = f"excluded_{filename}"
        cleaned_output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], cleaned_output_filename)
        excluded_output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], excluded_output_filename)

        try:
            # Pass the keywords_list to the data processing function
            excel_processor.process_data_excel(
                input_filepath,
                cleaned_output_filepath,
                excluded_output_filepath,
                keywords_list # Pass keywords here
            )

            # Return the URLs for downloading the processed files
            return jsonify({
                'message': 'File processed successfully',
                'cleaned_url': f'/downloads/{cleaned_output_filename}',
                'excluded_url': f'/downloads/{excluded_output_filename}'
            }), 200

        except Exception as e:
            # General error handling during processing
            print(f"Error during processing: {e}")
            return jsonify({'error': f'File processing failed: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Invalid file type. Only .xlsx files are allowed.'}), 400

@app.route('/downloads/<filename>')
def download_file(filename):
    """Serves the processed files for download."""
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    # For development, run with debug=True
    # For production, use a WSGI server like Gunicorn
    app.run(debug=True, host='0.0.0.0', port=5000)
