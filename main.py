from flask import Flask, render_template

import config

app = Flask(__name__)
app.use_static_for = 'static'
app.config.from_object(config)

# Validate file extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['VALID_FILE_EXTENSIONS']

@app.route('/app5')
def index():
    return render_template('index.html')    

if __name__ == "__main__":
    app.run(debug=True)