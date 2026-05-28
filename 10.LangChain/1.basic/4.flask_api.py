from dotenv import load_dotenv
load_dotenv()

from langchain_openai import ChatOpenAI 
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage

from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__)
llm = ChatOpenAI(model='gpt-4o-mini')


@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/name')
def name():
    prompt = [
        SystemMessage(content="you are a creative breanding  expert"),
        HumanMessage(content="what's a good company name that makes computer games")
    ]
    result = llm.invoke(prompt)
    return jsonify({'result':'success', 'chatbot':result.content})


@app.route('/api/name', methods=['POST'])
def name2():
    data = request.get_json()  # 사용자 입력값읽기
    product = data.get('product')
    user_prompt = f"what's a good company name that makes {product}. just give me a name"
    print(user_prompt)
    
    prompt = [
        SystemMessage(content="you are a creative breanding  expert"),
        HumanMessage(content=user_prompt)
    ]
    result = llm.invoke(prompt)
    names = [line.strip() for line in result.content.split()]

    return jsonify({'result':'success', 'chatbot': names})




@app.route('/api/dinner')
def dinner():
    prompt = [
        SystemMessage(content="당신은 경력 10년차 호텔 쉐프입니다"),
        HumanMessage(content="오늘 저녁을 추천해줘")
    ]

        
    result = llm.invoke(prompt)
    # print(result.content)

    return jsonify({'result':'success', 'chatbot':result.content})



if __name__ == "__main__":
    app.run(debug=True)












