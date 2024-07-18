from flask import Flask, request, render_template

from app.routes.core.insert_rows import insert_rows
from flask import send_file
from io import BytesIO

import config

# Create a Flask app
app = Flask(__name__)
app.use_static_for = 'static'

# Load config
app.config.from_object(config)

# Validate extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['VALID_FILE_EXTENSIONS']

@app.route('/generate_file', methods=['GET', 'POST'])
def generate_file_route():
    if 'pccs' in request.files:
        file = request.files['pccs']
        try:
            if allowed_file(file.filename):
                output_buffer = BytesIO()
                insert_rows(file, output_buffer=output_buffer)
                output_buffer.seek(0)
                return send_file(output_buffer, attachment_filename='pccs.xlsx', as_attachment=True)
        except Exception as e:
            return render_template('error.html', error_message=str(e)), 500  # Render error template for server errors
    return render_template('error.html', error_message='Invalid file extension'), 400