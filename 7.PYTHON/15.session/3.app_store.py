from flask import Flask, session
from flask_session import Session

app = Flask(__name__)
app.secret_key = 'your_secret_key' # 나만 아는 나의 세션 암호화 키
app.config['SESSION_TYPE'] = 'filesystem'  # 나의 세션을 파일
app.config['SESSION_FILE_DIR'] = './.sessions' # 내가 정한 폴더명

app.config['SESSION_PERMANENT'] = False
app.config['SESSION_USE_SIGNER'] = True

Session(app)

@app.route('/')
def main():

    if 'username' in session :
        return f"세션에서 당신의 정보를 찾았습니다 {session['username'],session['fullname'],session['hobby']}"
    
    session['username'] = 'spc2026'
    session['fullname'] = '이상욱'
    session['hobby'] = '돈쓰기'
    session['dob'] = '2000/05/03'

    return "첫방문."
    
if __name__ == "__main__":
    app.run(debug=True)