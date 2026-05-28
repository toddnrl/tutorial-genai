# pip install langchain langchain-openai
import os
from dotenv import load_dotenv

from langchain_openai import OpenAI
load_dotenv()

openai_api_key = os.environ.get('OPENAI_API_KEY')

llm = OpenAI(model='gpt-4o-mini')
llm = OpenAI(model='gpt-4o-mini', temperature=0.0)  # 1.0으로 갈수록 창의적인 답변이 나옴
llm = OpenAI(model='gpt-4o-mini', openai_api_key=openai_api_key)
llm = OpenAI(model='gpt-4o-mini', api_key=openai_api_key)


print(llm)

prompt = '오늘 저녁은 무엇을 먹을까요'
result = llm.invoke(prompt)

print(result)
