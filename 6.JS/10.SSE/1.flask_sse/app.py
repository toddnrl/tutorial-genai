from flask import Flask, send_from_directory, request, Response

from queue import Queue

app = Flask(__name__)

# 연결 사용자 관리
clients = []


@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

# 클라이언트에게 응답을 보낼 API _ SSE 방식으로 보낼 API
@app.route('/stream')
def stream():
    print('클라이언트 연결됨 - 누가 이 API를 듣고있음')

    def event_stream():
        q = Queue()
        clients.append(q) # 응답을 보낼 사용자 목록에 이 새로운 사용자를 추가
        try:
            yield f"data :서버에 연결되었습니다 \n\n"

            while True:
                message = q.get()
                if message is None:
                    break
                yield f"data: {message}\n\n"
        except GeneratorExit:
            print("클라 연결 종료")
        finally:
            if q in clients:
                clients.remove(q)

    return Response(event_stream(), mimetype='text/event-stream')
@app.route('/send', methods=["POST"])
def send():
    message = request.form.get('msg', "")
    print('클라이언트 메세지:' , message)
    for q in clients:
        q.put(f'서버가 받은 메세지:{message}' )
    return ("", 204)



if __name__ == "__main__":
    app.run(debug=True)