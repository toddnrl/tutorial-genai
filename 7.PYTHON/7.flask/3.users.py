from flask import Flask, jsonify

app = Flask(__name__)

users = [
    {'name' : 'Alice', 'age':25, 'phone':'23-123-234'},
    {'name' : 'Bob', 'age':24, 'phone':'123-234-666'},
    {'name' : 'Charlie', 'age':23, 'phone':'345-123-53'},
    {'name' : 'David', 'age':23, 'phone':'234-555-656'}

]
# 나이가 같은 두명을 반환

@app.route('/')
def main():
    return jsonify(users) 

@app.route('/user/<name>')
def get_user_by_name(name):
    print("사용자입력값: ", name)
    user = None
    for u in users:
        if u['name'].lower() == name.lower() :
            user = u

    if user:
        return jsonify(user)
    else : 
        return jsonify({"message": "user not found"})

@app.route('/user/<int:age>')
def get_user_by_age(age):
    print("사용자 입력값: ", age)

    user = []

    for u in users :
        if u['age'] == age:
            user.append(u)

    user = None
    for u in users :
        if u['age'] == age:
            user = u
        if user :
            return jsonify(user)
        else :
            return jsonify({"message" : "found"})

if __name__ == '__main__':
    app.run(debug=True)