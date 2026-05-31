from flask import Flask, render_template, request, redirect, session, jsonify
from flask import jsonify

from dotenv import load_dotenv
from openai import OpenAI

import os
import sqlite3

load_dotenv()

client = OpenAI(
    api_key=os.getenv('OPENAI_API_KEY')
)

app = Flask(__name__)
app.secret_key = 'my-secret-key'


DB_NAME = 'chatbot.db'



def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def save_message(user_id, role, content):
    conn = get_db()

    conn.execute(
        '''
        INSERT INTO chat_messages
        (user_id, role, content)
        VALUES (?, ?, ?)
        ''',
        (user_id, role, content)
    )
    
    conn.commit()
    conn.close()


def load_user_memory(user_id):
    conn = get_db()

    rows = conn.execute(
        '''
        SELECT role, content
        FROM chat_messages
        WHERE user_id = ?
        ORDER BY id ASC
        ''',
        (user_id,)
    ).fetchall()

    conn.close()

    messages = []

    for row in rows :
        messages.append({
            'role':row['role'],
            'content':row['content']
        })

    return messages

def init_db():
    conn = get_db()

    conn.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL
    )
    ''')

    conn.execute('''
        CREATE TABLE IF NOT EXISTS chat_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')


    conn.commit()
    conn.close()



@app.route('/api/chat', methods=['POST'])
def api_chat():
    if 'user_id' not in session:
        return jsonify({
            'error':'로그인이 필요합니다'
        }), 401
    
    user_id = session['user_id']
    user_message = request.json.get('message')

    save_message(
        user_id,
        'user',
        user_message
    )

    memory = load_user_memory(user_id)

    messages = [
        {
            'role' : 'system',
            'content': '당신은 친절한 AI 챗봇입니다.'
        },
        *memory
    ]

    response = client.chat.completions.create(
        model='gpt-4o-mini',
        messages=messages
    )

    ai_message = response.choices[0].message.content

    save_message(user_id, 'assistant', ai_message)
    

    return jsonify({
        'reply': ai_message
    })



@app.route('/signup', methods=['POST', 'GET'])

def signup():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')

        conn = get_db()

        conn.execute(
            'INSERT INTO users (username, password) VALUES (?, ?)',
            (username, password)
        )

        conn.commit()
        conn.close()

        return redirect('/')
    return render_template('signup.html')



@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')

        conn = get_db()

        user = conn.execute(
            '''
            SELECT * FROM users WHERE username = ? AND password = ?
        ''',
        (username, password)
        ).fetchone()

        conn.close()

        if user: 
            session['user_id'] = user['id']
            session['username'] = user['username']

            return redirect('/')
        
        return '실패'
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')



@app.route('/')
def index() :
    if 'user_id' not in session :
        return redirect('/login')
    return redirect('/chat')


@app.route('/chat')
def chat_page():

    if 'user_id' not in session:
        return redirect('/login')
    
    return render_template('chat.html', username=session['username'])

if __name__ == '__main__':
    init_db()
    app.run(debug=True)