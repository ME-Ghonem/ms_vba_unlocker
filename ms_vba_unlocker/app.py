from flask import Flask, request, send_file, render_template, jsonify
import zipfile
import os
import re
import shutil
from io import BytesIO

app = Flask(__name__)

# Paths for storing uploaded files and the modified files
UPLOAD_FOLDER = 'uploads'
MODIFIED_FOLDER = 'modified_files'

# Ensure the directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(MODIFIED_FOLDER, exist_ok=True)

# Extract VBA content from the Office file
def get_vba_from_zip(file_path):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        # Check for common locations of vbaProject.bin
        for name in ['xl/vbaProject.bin', 'word/vbaProject.bin', 'ppt/vbaProject.bin']:
            if name in zip_ref.namelist():
                return name, zip_ref.read(name)
    return None, None

# Add the modified VBA content back to the ZIP file
def add_vba_to_zip(file_path, modified_vba, original_vba_path):
    # Create a temporary ZIP file to modify
    temp_zip_path = file_path + '.modified'
    
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        with zipfile.ZipFile(temp_zip_path, 'w') as temp_zip:
            # Copy everything from the original ZIP except the old vbaProject.bin
            for item in zip_ref.infolist():
                if item.filename != original_vba_path:
                    temp_zip.writestr(item, zip_ref.read(item.filename))
            
            # Add the modified vbaProject.bin back
            temp_zip.writestr(original_vba_path, modified_vba)

    # Replace the original file with the new modified ZIP
    shutil.move(temp_zip_path, file_path)

# Handle the uploaded file and process it
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        # Save the uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        # Process the file to remove VBA password protection
        original_vba_path, vba_data = get_vba_from_zip(file_path)
        if vba_data is None:
            return jsonify({'error': 'No VBA project found or file is corrupted'}), 400

        # If the VBA is protected (contains DPB=), we modify it
        if b'DPB=' in vba_data:
            vba_data = vba_data.replace(b'DPB=', b'DPx=')

        # Save the modified file
        modified_file_path = os.path.join(MODIFIED_FOLDER, file.filename)
        try:
            add_vba_to_zip(file_path, vba_data, original_vba_path)
            # Return the modified file for download
            return send_file(file_path, as_attachment=True)
        except Exception as e:
            return jsonify({'error': f'Error processing the file: {str(e)}'}), 500

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
