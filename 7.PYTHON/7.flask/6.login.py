from flask import Flask, request

app = Flask(__name__)

@app.route('/')
def home():
    return """
    <form action="/login" method="post">
        <input type="text" name="id">
        <input type="password" name="pw">
        <button>로그인</button>
    </form>
    """

@app.route('/login', methods=['POST'])
def login():

    user_id = request.form['id']
    user_pw = request.form['pw']

    return f"아이디: {user_id}"

if __name__ == '__main__':
    app.run(debug=True)