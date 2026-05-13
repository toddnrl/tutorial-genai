from flask import Flask

app = Flask(__name__)

@app.route('/add/<int:a>/<int:b>')
def add(a, b):
    return f"결과: {a+b}"

@app.route('/mul/<int:a>/<int:b>')
def mul(a, b):
    return f"결과: {a*b}"

if __name__ == '__main__':
    app.run(debug=True)