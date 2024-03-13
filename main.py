
from flask import Flask, request, render_template, send_from_directory
from app.core.insert_rows import insert_rows
from app.config import config

app = Flask(__name__)
app.use_static_for = 'static'

# Configuration
app.config.from_object(config)

###############################
### Validate file extension ###
###############################

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['VALID_FILE_EXTENSIONS']

###############################
### Validate file extension ###
###############################

@app.route('/app5')
def index():
    """Homepage with a button to generate DOCX files"""
    return render_template('index.html')    

@app.route('/pccs/generate_file', methods=['POST'])
def generate_xlsx():
    
    if 'MAT' in request.files:
        file = request.files['MAT']
        try:
            if allowed_file(file.filename):
                insert_rows(file)
                return send_from_directory('.', 'PCCS.xlsx', as_attachment=True)
        except Exception as e:
            print(e)
            return render_template('error.html'), 500
    return render_template('error.html'), 400

if __name__ == "__main__":
    app.run(debug=True)