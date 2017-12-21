import os, shutil
from flask import Flask, render_template, request, send_from_directory, flash
from werkzeug.utils import secure_filename
from app import AutoExcel
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = set(['xls','xlsx'])

app = Flask(__name__)
app.secret_key = 'autoexcel'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/fileupload',methods =['POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'ipfile' not in request.files:
            flash('No file part')
            return render_template("index.html")
        file = request.files['ipfile']
        # if user does not select file, browser also
        # submit a empty part without filename
        if file.filename == '':
            flash('No selected file')
            return render_template("index.html")
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(filename)
            print('Converting...')
            opFile = AutoExcel.To_xlsx.convert(filename)
            return send_from_directory(app.config['UPLOAD_FOLDER'], opFile)
        else:
            return 'Not a Valid File. Upload only .xls files'
    else:
        return 'Failed'

if __name__ == '__main__':
    app.run(port=5001, debug=True)