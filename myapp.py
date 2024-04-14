import os
import tempfile
from flask import Flask, render_template, request, send_file
from cv_parser import process_bundle_cv

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "No file part"
    files = request.files.getlist('file')  # Handle multiple files
    if not files:
        return "No files selected"
    temp_dir = tempfile.mkdtemp()
    for file in files:
        if file.filename == '':
            continue
        file_path = os.path.join(temp_dir, file.filename)
        file.save(file_path)
    output_file = process_bundle_cv(temp_dir)
    if os.path.exists(output_file):
        return send_file(output_file, as_attachment=True)
    else:
        return "Error: Output file not found"

if __name__ == '__main__':
    app.run(debug=True)
