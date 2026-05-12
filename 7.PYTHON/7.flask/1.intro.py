from flask import Flask

app = Flask(__name__)

@app.route('/')

def home():
    5/0
    return """
    <html>
        <head>
            <title>제목</title>
            <style>
                p{
                    color:red;
                }
            </style>
        </head>
        <body>
            <h1>컴 투 마이 홈</h1>
            <p>여기는 텍스트 본문1</p>
            <p>여기는 텍스트 본문2</p>
        </body>
    </html>
    
    <h1>컴 투 마이 홈</h1>
    """

if __name__ =='__main__':
    app.run(debug=True) # 배포, 운영할땐 꼭 제거