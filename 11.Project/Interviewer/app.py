from flask import Flask, render_template, request
import os
from dotenv import load_dotenv
from openai import OpenAI


load_dotenv()
client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

app = Flask(__name__)



@app.route('/', methods=['POST','GET'])
def index():
    
    questions = None

    if request.method == 'POST':

        resume = request.form.get('resume')

        response = client.chat.completions.create(
            model='gpt-4o-mini',
            messages=[
                {
                    'role':'system',
                    'content':'당신은 백엔드 개발자 면접관입니다.'
                },
                {
                    'role':'user',
                    'content':f'''
                                다음 자기소개서를 보고 면접 질문 3개를 만들어주세요
                                자기소개서: {resume}
                                '''
                }
            ]
        )

        questions = response.choices[0].message.content
    return render_template('index.html', questions=questions)


@app.route('/answer', methods=['POST'])
def answer():
    questions = request.form.get('questions')
    answer = request.form.get('answer')

    response = client.chat.completions.create(
        model='gpt-4o-mini',
        messages=[
            {
                'role':'system',
                'content':'당신은 백엔드 개발자 면접 답변을 평가하는 면접관입니다'
            },
            {
                'role':'user',
                'content':f'''
                                다음 면접 질문과 지원자의 답변을 보고 평가해주세요
                                면접 질문 :
                                {questions}
                                면접 답변 :
                                {answer}
                                아래 형식으로 답변해주세요.
                                점수: 100점 만점
                                좋은 점:
                                개선할 점:
                                모범 답변:
                            '''
            }
        ]
    )

    feedback = response.choices[0].message.content

    return render_template('result.html', feedback=feedback)


if __name__ == '__main__':
    app.run(debug=True)

