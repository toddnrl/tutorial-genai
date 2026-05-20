from flask import Flask, render_template, request
from flask import redirect, url_for
from flask import session, flash

from datetime import timedelta


import sqlite3

app = Flask(__name__)
app.secret_key = "andrea144"  ## == 내가 정한 시크릿키  커밋 안함   .env에 넣고 사용   .env 는 커밋 안함

app.permanent_session_lifetime = timedelta(minutes=5)


DATABASE = 'users.sqlite3'  ## 나의 파일명


def get_db_connection():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row      ## 나의 결과를 다 dict 포맷으로 관리하겠다
                                        ## row[0] 이렇게 접근해야하는걸 row['id'] 이런식으로 사용 가능
    return conn



def init_db():
    with app.app_context():
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                password TEXT NOT NULL                   
            )
        ''')

        cur.execute("SELECT COUNT(*) AS count FROM users")
        count = cur.fetchone()['count']
        if count == 0:
            cur.execute("INSERT INTO users (username, password) VALUES (?, ?)", ("user1", "password1"))
            cur.execute("INSERT INTO users (username, password) VALUES (?, ?)", ("user2", "password2"))
         
        cur.execute('SELECT * FROM users')
        rows = cur.fetchall()

        print('-' * 30)
        for row in rows:
            print(row['id'], row['username'], row['password'])

        print('-' * 30)

        conn.commit()
        conn.close()


@app.route('/')
def home():
    return render_template('index.html')





@app.route('/profile', methods=['GET', 'POST'])
def profile():
    if "user" not in session:
        flash("로그인이 필요합니다")
        return redirect(url_for("login"))

    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        if not username or not password:
            flash("아이디와 비밀번호를 입력해주세요")
            conn.close()
            return redirect(url_for("profile"))

        # 다른 사람이 이미 쓰는 아이디인지 확인
        cur.execute(
            "SELECT * FROM users WHERE username = ? AND username != ?",
            (username, session["user"])
        )
        existing_user = cur.fetchone()

        if existing_user:
            flash("해당 ID는 사용할 수 없음")
            conn.close()
            return redirect(url_for("profile"))

        # 현재 로그인한 유저 정보 수정
        cur.execute(
            "UPDATE users SET username = ?, password = ? WHERE username = ?",
            (username, password, session["user"])
        )
        conn.commit()

        # 세션 아이디도 새 아이디로 변경
        session["user"] = username

        flash("수정 성공!")
        conn.close()
        return redirect(url_for("profile"))

    cur.execute("SELECT * FROM users WHERE username = ?", (session["user"],))
    user = cur.fetchone()
    conn.close()

    return render_template("profile.html", user=user)










@app.route('/signin', methods=['GET', 'POST'])
def signin():

    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        conn = get_db_connection()
        cur = conn.cursor()

        if not username or not password:
            flash("아이디와 비밀번호를 입력해주세요")
            conn.close()
            return redirect(url_for("signin"))


        cur.execute("SELECT * FROM users WHERE username = ?", (username,))
        existing_user = cur.fetchone()



        if existing_user:
            flash("해당 ID는 사용할 수 없음")
            conn.close()
            return redirect(url_for('signin'))

        
        cur.execute(
            "INSERT INTO users (username, password) VALUES (?, ?)",
            (username, password)
        )
        conn.commit()
        conn.close()

        flash("회원가입 성공! 로그인해주세요")
        return redirect(url_for("login"))


    return render_template('signin.html')



@app.route('/login', methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT *  FROM users WHERE username = ? AND password = ?", (username,password))
        user_data = cur.fetchone()
        conn.close()

        if user_data:
            session['user'] = username
            flash("로그인에 성공")
            return redirect(url_for("home"))
        else:
            flash("로그인 실패")
            return redirect(url_for("login"))

    return render_template('login.html')

@app.route('/logout')
def logout():
    flash("성공적으로 로그아웃이 되었습니다")
    session.pop("user", None)
    return redirect(url_for("home"))







if __name__ == "__main__":
    init_db()
    app.run(debug=True)