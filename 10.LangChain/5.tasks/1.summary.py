# 목적: 긴 문장을 받아서 짧게 요약한다

from dotenv import load_dotenv

from langchain_core.prompts import (
    ChatPromptTemplate,
    HumanMessagePromptTemplate,
    SystemMessagePromptTemplate,
    AIMessagePromptTemplate
)

from langchain_openai import ChatOpenAI
from langchain_core.runnables import RunnableLambda

load_dotenv()

template = '다음의 긴 내용을 3개의 문장으로 요약하시오: \n\n{article}'
chat_prompt = ChatPromptTemplate.from_messages([
    SystemMessagePromptTemplate.from_template('당신은 전문 문장 요약가입니다.'),
    HumanMessagePromptTemplate.from_template(template)
])

llm = ChatOpenAI(model='gpt-4o-mini', temperature=0.3)

chain = chat_prompt | llm | RunnableLambda(lambda x: {'summary': x.content.strip()})

input_text = {
    'article':'네이버는 인공지능(AI) 시대에 발맞춘 콘텐츠 생태계 구축을 위해 향후 5년간 1조 원을 투입할 계획이라고 28일 밝혔다. AI 플랫폼 경쟁력이 양질의'
    '콘텐츠를 생성하는 우수한 창작자에 있다고 보고 기술 외적인 서비스 혁신과 창작자 지원 등에 대규모 투자를 단행하겠다는 취지다.'
    '김광현 네이버 최고데이터·콘텐츠책임자(CDO)는 이날 서울 중구 더플라자호텔 서울에서 열린 간담회에서 “창작자 생태계와 외부 파트너십을 통해 실행형'
    '에이전트의 기반이 되는 양질의 데이터를 잘 쌓고, 이를 AI와 연결해 차별화된 사용자 경험을 제공하며 경쟁에 더 과감히 뛰어들겠다”고 밝혔다.'

}

result = chain.invoke(input_text)
print('요약 결과:', result['summary'])