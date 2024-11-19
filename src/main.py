from flask import Flask, request, jsonify, render_template
import os

template_folder = os.path.join(os.path.dirname(__file__), 'templates')
app = Flask(__name__, template_folder=template_folder)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'file1' not in request.files or 'file2' not in request.files or 'file3' not in request.files or 'file4' not in request.files or 'zipFile' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file1 = request.files['file1']
    file2 = request.files['file2']
    file3 = request.files['file3']
    file4 = request.files['file4']
    zip_file = request.files['zipFile']

    file1.save(os.path.join(app.config['UPLOAD_FOLDER'], file1.filename))
    file2.save(os.path.join(app.config['UPLOAD_FOLDER'], file2.filename))
    file3.save(os.path.join(app.config['UPLOAD_FOLDER'], file3.filename))
    file4.save(os.path.join(app.config['UPLOAD_FOLDER'], file4.filename))
    zip_file.save(os.path.join(app.config['UPLOAD_FOLDER'], zip_file.filename))

    return jsonify({'message': 'Files uploaded successfully'}), 200

if __name__ == '__main__':
    app.run(debug=True)
