from dotenv import load_dotenv

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser

load_dotenv()

llm = ChatOpenAI(model='gpt-4o-mini')

parser = StrOutputParser()

prompt = ChatPromptTemplate.from_messages([
    ('system', '당신은 상품명을 지어주는 기획자'),
    ('user', '{company} 회사에서 {product}를 만드는데 이 제품명 만들어')

])

chain = prompt | llm | parser   # << = 이걸 LCEL라고 부른다


inputs = {'company': 'ai 첨단 기술 회사', 'product': '화장품'}

result = chain.invoke(inputs)

print('최종결과 :' ,result)