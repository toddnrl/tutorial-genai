from flask import Flask, render_template, request, redirect
import sqlite3


app = Flask(__name__)

def get_db():
    conn = sqlite3.connect('memo.db')
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS memos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            message TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

memos = []

@app.route('/')
def index():
    conn = get_db()
    memos = conn.execute('SELECT * FROM memos ORDER BY id DESC').fetchall()
    conn.close()

    return render_template('index.html', memos=memos)


@app.route('/create', methods=['POST'])
def create():
    title = request.form.get('title')
    message = request.form.get('message')

    conn = get_db()
    conn.execute(
        'INSERT INTO memos (title, message) VALUES(?, ?)',
        (title, message)
    )
    conn.commit()
    conn.close()

    return redirect('/')


@app.route('/delete/<int:memo_id>')
def delete(memo_id):
    conn = get_db()

    conn.execute(
        'DELETE FROM memos WHERE id = ?',
        (memo_id,)
    )

    conn.commit()
    conn.close()

    return redirect('/')


if __name__ == "__main__":
    init_db()
    app.run(debug=True)