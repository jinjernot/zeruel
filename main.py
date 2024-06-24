
from flask import Flask, request, render_template
from app.core.insert_rows import insert_rows
from app.config import config

from flask import send_file
from io import BytesIO

app = Flask(__name__)
app.use_static_for = 'static'

# Configuration
app.config.from_object(config)

# Validate file extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['VALID_FILE_EXTENSIONS']

@app.route('/app5')
def index():
    """Homepage with a button to generate DOCX files"""
    return render_template('index.html')    

@app.route('/generate_file', methods=['POST'])
def generate_file():
    
    if 'MAT' in request.files:
        file = request.files['MAT']
        try:
            if allowed_file(file.filename):
                output_buffer = BytesIO()
                insert_rows(file, output_buffer=output_buffer)
                output_buffer.seek(0)
                return send_file(output_buffer, attachment_filename='pccs.xlsx', as_attachment=True)
        except Exception as e:
            print(e)
            return render_template('error.html'), 500
    return render_template('error.html'), 400

if __name__ == "__main__":
    app.run(debug=True)