from flask import Flask, render_template, request
import os

app = Flask(__name__)


# app.config['UPLOAD_FOLDER'] = 'uploads'


# os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('form.html')


@app.route('/login', methods=['POST'])
def login():

    id = request.form.get('id')
    pw = request.form.get('pw')


    print(id)
    print(pw)
    return render_template('login.html', name=id)

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['photo']

    print(file)
    
    filename = file.filename
    filepath = os.path.join(app.frgs['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    
    return "파일 잘 받음"


if __name__ == '__main__':
    app.run(debug=True)
