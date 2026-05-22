from flask import Flask, request, jsonify, send_from_directory
import openai, os
from dotenv import load_dotenv

load_dotenv()

client = openai.OpenAI(api_key=os.getenv("OPEN_API_KEY"))

app= Flask(__name__)

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/api/chat', methods=['POST'])
def chat():
    data = request.get_json()
    chat_message = data.get('chatMessage','')
    print('사용자 입력값: ', chat_message)

    get_reply = ask_chatgpt(chat_message)

    return jsonify({'reply' : f'답변:{get_reply}'})




def ask_chatgpt(chat_message):
    response = client.chat.completions.create(
        model='gpt-3.5-turbo',
        messages=[
            {'role': 'system', 'content': '당신의 나의 질문에 답변을 잘 하는 챗봇입니다.'},
            {'role': 'user', 'content': chat_message}
        ]
    )

    final_response = response.choices[0].message.content
    return final_response


# while True:
#     chat_message = input("\n질문: ").strip()
#     chatbot_response = ask_chatgpt(chat_message)
#     print("챗봇응답: ", chatbot_response)


if __name__ == '__main__':
    app.run(debug=True)


