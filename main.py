from flask import Flask, request, render_template,send_file
from data.queryReport import queryReport

app = Flask(__name__)

app.use_static_for = 'static'

ALLOWED_EXTENSIONS = {'xlsx','xls'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'Regular' in request.files and 'Other' in request.files:
            file1 = request.files['Regular']
            file2 = request.files['Other']
            try:
                if allowed_file(file1.filename) and allowed_file(file2.filename):
                    queryReport(file1)
                    return send_file('QueryReport.xlsx', as_attachment=True)
            except Exception as e:
                print(e)
                return render_template('error.html')
        return render_template('error.html')
    return render_template('index.html')


def main():
    upload_file()

if __name__ == '__main__':
    app.run(debug=True)
    main()