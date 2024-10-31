import os
import time  # Optional delay for troubleshooting
import logging
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, session
from werkzeug.utils import secure_filename
from data_processing import run_report_generation

# Initialize Flask app
app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Set a secret key for session management
app.config['UPLOAD_FOLDER'] = 'uploads'  # Directory for file uploads
app.config['GENERATED_REPORTS'] = 'generated_reports'  # Directory for generated reports

# Ensure the upload and report folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_REPORTS'], exist_ok=True)

# Set up logging
logging.basicConfig(level=logging.INFO)

# Home route (for file upload form)
@app.route('/')
def index():
    return render_template('index.html')  # Render the home page template

# Process file upload and report generation
@app.route('/process', methods=['POST'])
def process_file():
    # Check if the file is in the request
    if 'file' not in request.files:
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        return redirect(url_for('index'))

    # Save the uploaded file
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    # Get class names from the form
    class_name1 = request.form.get('class_name1', '')
    class_name2 = request.form.get('class_name2', '')

    # Generate a session ID to track progress
    session_id = os.urandom(8).hex()
    session[session_id] = {"progress": 0, "status": "Initializing"}

    # Run the report generation process
    result = run_report_generation(file_path, class_name1, class_name2, session_id)

    # Error handling
    if result["status"] == "error":
        logging.error("Error during report generation.")
        return jsonify(result), 400

    # Use the path directly from result to avoid adding "uploads" twice
    report_path = result["report_path"]
    session[session_id]["report_path"] = report_path

    # Check if the file exists and list directory contents for debugging
    if os.path.isfile(report_path):
        logging.info(f"Report generated and stored at: {report_path}")
    else:
        logging.error("Report file not found immediately after generation: %s", report_path)
        logging.info("Listing contents of the directory for verification:")
        logging.info(os.listdir(os.path.dirname(report_path)))
        return jsonify({"status": "error", "message": "Report generation failed"}), 500

    return jsonify({"status": "success", "session_id": session_id})

# Check progress of report generation
@app.route('/progress/<session_id>')
def check_progress(session_id):
    progress_info = session.get(session_id, {"progress": 0, "status": "No progress data"})
    return jsonify(progress_info)

# Download the generated report
@app.route('/download/<session_id>')
def download_report(session_id):
    report_path = session.get(session_id, {}).get("report_path")
    logging.info(f"Attempting to download report from: {report_path}")

    # Optional delay for troubleshooting
    time.sleep(0.5)  # Half-second delay to ensure file is accessible

    # Ensure the path exists and is correct
    if report_path and os.path.isfile(report_path):
        return send_file(report_path, as_attachment=True)
    else:
        logging.error("File not found at the specified path: %s", report_path)
        return "File not found", 404

# Run the app
if __name__ == '__main__':
    app.run(debug=True)
