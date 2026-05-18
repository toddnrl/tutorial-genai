from flask import Flask, render_template, request
from flask import session, redirect, url_for

# Session은 더이상 안함 -> DB로 대체

app = Flask(__name__)
app.secret_key = 'my-random-key'

users = [
    {'name':'Alice', 'id': 'alice', 'pw':'alice'},
    {'name': 'Bob', 'id': 'bob', 'pw': 'bob123'},
    {'name': 'Charlie', 'id': 'charlie', 'pw': 'charlie123'},
    {'name': 'David', 'id': 'david', 'pw': 'david123'},
    {'name': 'Eve', 'id': 'eve', 'pw': 'eve123'}

]

@app.route('/dashboard')
def welcome():
    user = session.get('user')
    return render_template('dashboard.html', name=user['name'])



@app.route('/', methods=['GET'])
def home():
    if session.get('user'):
        return redirect(url_for('welcome'))
    

    return render_template('index.html')


@app.route('/', methods=['POST'])

def login():
    if request.method == 'POST':
        # 요청에서 id/pw 가져온다
        id = request.form.get('id')
        pw = request.form.get('pw')

        # 2. user DB에서 이 사용자를 매칭

        user = next ((u     for u in users      if u['id'] == id and u['pw'] == pw), None)

        # 3. 사용자가 있으면?

        if user :
            session['user'] = user
            error = None
            return redirect(url_for('welcome'))
        else :
            error = "Invaild ID or password"

        return render_template('index.html', error=error)
    

# 1 사용자가 비밀번호를 바꾸도록
# 1-1 method를 post로 확장
# 1-2 users 안에서 나의 비번을 바꾼다
# 1-3 성공적으로 변경되면 나의 profile에서 확인한다
# 1-4 비밀번호 변경을 눌렀을때 성공적으로 변경되었음을 알려준다


@app.route('/profile', methods=['GET', 'POST'])
def profile():
    user = session.get('user')   # 세션 안에 있는 데이터를 변경 안했음
    if not user :
        return redirect(url_for('home'))
    message = None

    
    if request.method == 'POST':

        new_pw = request.form.get('new_pw')

        for u in users:

                if u['id'] == user['id']:

                    u['pw'] = new_pw

                    # 따라서 세션도 최신화
                    session['user'] = u # 세션정보 구 -> 신 버전으로 갱신

                    message = "비밀번호가 변경되었습니다."
                    # return render_template('profile.html', user=user, message=message)
                    return redirect(url_for('profile'))
    
    return render_template(
        'profile.html',
        user=session.get('user'),
        message=message
    )





@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('home'))


if __name__ == "__main__":
    app.run(debug=True)