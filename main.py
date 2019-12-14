import os
import re
import sys
import subprocess
from uuid import uuid4
from flask import Flask, render_template, request, jsonify, send_from_directory


app = Flask(__name__)

app.config['UPLOADS_FOLDER'] = "./UPLOAD_FOLDER"
app.config['EXPORT_PDF_FOLDER'] = "./PDF"


def save_to(folder, file):
    os.makedirs(folder, exist_ok=True)
    save_path = os.path.join(folder, secure_filename(file.filename))
    file.save(save_path)
    return save_path



def convert_to(folder, source, timeout=None):
    args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', folder, source]

    process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    filename = re.search('-> (.*?) using filter', process.stdout.decode())

    if filename is None:
        raise LibreOfficeError(process.stdout.decode())
    else:
        return filename.group(1)

def libreoffice_exec():
    # TODO: Provide support for more platforms
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'

class LibreOfficeError(Exception):
    def __init__(self, output):
        self.output = output



@app.route('/doc2pdf', methods=['POST'])
def upload_file(re):
    upload_id = str(uuid4())
    source = save_to(os.path.join(app.config['UPLOADS_FOLDER'], 'source', upload_id), request.files['file'])

    try:
        result = convert_to(os.path.join(app.config['UPLOADS_FOLDER'], 'pdf', upload_id), source, timeout=15)
    except LibreOfficeError:
        raise InternalServerErrorError({'message': 'Error when converting file to PDF'})
    except TimeoutExpired:
        raise InternalServerErrorError({'message': 'Timeout when converting file to PDF'})

    return jsonify({'result': {'source': uploads_url(source), 'pdf': uploads_url(result)}})


@app.route('/pdf/<path:path>', methods=['GET'])
def serve_uploads(path):
    return send_from_directory(app.config['EXPORT_PDF_FOLDER'], path)


@app.errorhandler(500)
def handle_500_error():
    return InternalServerErrorError().to_response()


@app.errorhandler(RestAPIError)
def handle_rest_api_error(error):
    return error.to_response()


if __name__ == '__main__':
    app.run(host='0.0.0.0', threaded=True)
