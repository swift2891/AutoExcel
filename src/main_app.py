import os, shutil
from flask import Flask, render_template, request, send_from_directory, flash
from werkzeug.utils import secure_filename
from CheckXLS import XLSCheck
from Manipulator import AutoExcel

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = set(['xls','xlsx'])

app = Flask(__name__)
app.secret_key = 'autoexcel'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Cleans Uploads Folder:
def clean():
    for the_file in os.listdir('uploads'):
        file_path = os.path.join('uploads', the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(e)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/fileupload',methods =['POST'])
def upload_file():
    clean() # Cleans uploads
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
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            sheetsList = ManipulateFile()
            file_Selected = filename
            return render_template('get_config.html', listOfSheets=sheetsList, file_Selected=file_Selected)
        else:
            return 'Not a Valid File. Upload only .xls files'
    else:
        return 'Failed'

@app.route('/output',methods =['POST'])
def manipulate_excel():
    sheet_selected=request.form['sheetSelect']
    potential_ip = request.form['potential_ip']
    current_ip = request.form['current_ip']
    time_ip = request.form['time_ip']
    capacity_ip = request.form['capacity_ip']
    rowstart_ip = request.form['rowstart_ip']
    gap_ip = request.form['gap_ip']
    valList=[]
    valList.append(sheet_selected)
    valList.append(potential_ip)
    valList.append(current_ip)
    valList.append(capacity_ip)
    valList.append(time_ip)
    valList.append(rowstart_ip)
    valList.append(gap_ip)
    for v in valList:
        print(v)
    AutoExcel.initialize(valList)
    AutoExcel.mainApp()
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'output.xlsx')

def ManipulateFile():
    listOfSheets = []
    print('Manipulating...')
    # Check if a file exist
    for fname in os.listdir('uploads'):
        if fname.endswith('xlsx') or fname.endswith('xls'):
            print('Input File uploaded Successfully')
            targetFile = XLSCheck.checkInput(fname)
            print('Converted File: '+targetFile)
            listOfSheets = AutoExcel.loadSheets()
            if len(listOfSheets)>1:
                return listOfSheets
            break
    else:
        print('Not good. no file in uploads dir')

@app.before_first_request
def initialize():
    clean()

if __name__ == '__main__':
    app.run(port=5001, debug=True)