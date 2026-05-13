from flask import Flask, request

app = Flask(__name__)

@app.route('/search')
def search():

    keyword = request.args.get('keyword')

    return f"검색어: {keyword}"

if __name__ == '__main__':
    app.run(debug=True)