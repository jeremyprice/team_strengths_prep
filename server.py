#!/usr/bin/env python3

from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash, send_file
import team_strengths_prep
import os
import logging
import logging.handlers
from werkzeug.utils import secure_filename

DEBUG = False
app_log = logging.getLogger('team_strengths')
app = Flask(__name__)


def setup_paths():
    for req_path in ('logs/', 'zips/', 'uploads/'):
        try:
            os.mkdir(req_path)
        except FileExistsError:
            pass


def setup_logging():
    if DEBUG:
        loglevel = logging.DEBUG
    else:
        loglevel = logging.INFO
    formatter = logging.Formatter('%(asctime)s %(name)-12s %(levelname)-8s %(message)s')
    # send messages to a rotated log file - rotated every monday
    flask_handler = logging.handlers.TimedRotatingFileHandler('./logs/flask.log', when='W0')
    flask_handler.setFormatter(formatter)
    flask_handler.setLevel(loglevel)
    app.logger.addHandler(flask_handler)
    app.logger.setLevel(loglevel)
    app_handler = logging.handlers.TimedRotatingFileHandler('./logs/app.log', when='W0')
    app_handler.setFormatter(formatter)
    app_handler.setLevel(loglevel)
    app_log.addHandler(app_handler)
    app_log.setLevel(loglevel)


UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads/')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
ALLOWED_EXTENSIONS = {'xls', 'pdf'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return send_file('index.html')

@app.route('/generate', methods=['POST', 'GET'])
def generate():
    if request.method == 'GET':  # if we have a GET, send them to the root
        return redirect('/')
    # load info from the post data and generate a PDF and return it
    app_log.info('Got a submission')
    prefix = request.form.get('prefixInput')
    if 'listFile' not in request.files or 'reportFile' not in request.files:
        flash('List and Report files must be specified')
        return redirect(request.url)
    listFile = request.files['listFile']
    reportFile = request.files['reportFile']
    if listFile.filename == '' or reportFile.filename == '':
        flash('List and Report files must be specified')
        return redirect(request.url)
    if listFile and allowed_file(listFile.filename):
        listFilename = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(listFile.filename))
        listFile.save(listFilename)
        listFile.close()
    if reportFile and allowed_file(reportFile.filename):
        reportFilename = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(reportFile.filename))
        reportFile.save(reportFilename)
        reportFile.close()
    zip_results = team_strengths_prep.prep(reportFilename, listFilename, prefix)
    return send_file(zip_results, as_attachment=True)


@app.route('/python/')
def python_dir_list():
    files = sorted(os.listdir('python/'))
    return render_template('files.html', files=files, path='Python')


@app.route('/python/<fname>')
def serve_python(fname):
    return send_from_directory('python/', fname)


setup_paths()
setup_logging()

if __name__ == '__main__':
    app_log.info('Application started')
    app.run(debug=DEBUG)
