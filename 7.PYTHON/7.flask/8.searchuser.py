from flask import Flask
from flask import jsonify
from flask import request

app = Flask(__name__)




users = [
    {'name' : 'Alice', 'age':25, 'phone':'23-123-234'},
    {'name' : 'Bob', 'age':24, 'phone':'123-234-666'},
    {'name' : 'Charlie', 'age':23, 'phone':'345-123-53'},
    {'name' : 'David', 'age':23, 'phone':'234-555-656'}

]


@app.route('/search')
def search_user():
    name = request.args.get('name')
    age = request.args.get('age')
    phone = request.args.get('phone')
    result = users

    if name:
        result = [u for u in users if name.lower() in u['name'].lower()]

    if age:
        result = [u for u in result if int(age) == u['age']]

    if phone:
        # result = [u for u in result if phone == u['phone']]
        result = [u for u in result if u['phone'].startswith(phone)]
    # 쿼리 파라미터로 name age phone 로 검색해서 결과를 반환 

    return jsonify(result)

if __name__ == '__main__':
    app.run(debug=True)



# from flask import Flask, jsonify, request

# app = Flask(__name__)

# users = [
#     {'name': 'Alice', 'age': 25, 'phone': '23-123-234'},
#     {'name': 'Bob', 'age': 24, 'phone': '123-234-666'},
#     {'name': 'Charlie', 'age': 23, 'phone': '345-123-53'},
#     {'name': 'David', 'age': 23, 'phone': '234-555-656'}
# ]

# @app.route('/search')
# def search_user():

#     # 쿼리 파라미터 가져오기
#     name = request.args.get('name')
#     age = request.args.get('age')
#     phone = request.args.get('phone')

#     result = []

#     for user in users:

#         # 조건 검사
#         match = True

#         if name:
#             if user['name'].lower() != name.lower():
#                 match = False

#         if age:
#             if user['age'] != int(age):
#                 match = False

#         if phone:
#             if user['phone'] != phone:
#                 match = False

#         # 모든 조건 만족하면 추가
#         if match:
#             result.append(user)

#     return jsonify(result)

# if __name__ == '__main__':
#     app.run(debug=True)
