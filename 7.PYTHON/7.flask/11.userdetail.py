from flask import Flask, render_template

app = Flask(__name__)

users = [
    {'name' : '홍길동', 'age':25, 'phone':'23-123-234'},
    {'name' : '김길동', 'age':24, 'phone':'123-234-666'},
    {'name' : '이길동', 'age':23, 'phone':'345-123-53'},
    {'name' : '고길동', 'age':23, 'phone':'234-555-656'}

]

@app.route('/')
def index():
    final_html = render_template('users_detail.html', users=users)
    print(final_html)
    return final_html



if __name__=='__main__':
    app.run(debug=True)