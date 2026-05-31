from dotenv import load_dotenv

from langchain_openai import ChatOpenAI

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnableParallel


load_dotenv()

llm = ChatOpenAI(model='gpt-4o-mini')

prompt1 = ChatPromptTemplate.from_template('다음 뉴스를 2~3문장으로 요약해줘\n\n {news}')
prompt2 = ChatPromptTemplate.from_template('')

summary_chain = prompt1 | llm | StrOutputParser()


summary_chain = (
    ChatPromptTemplate.from_template('다음 뉴스의 전체적 감성을 한 단어로 분석해줘 (긍정, 부정, 중립)\n\n{news}')
    | llm
    | StrOutputParser()
)

sentiment_chain = (
    ChatPromptTemplate.from_template(
        '다음 뉴스의 카테고리를 분석하시오 \n'
        '경제, 정치, 사회, IT, 국제 중 하나로 답하시오. \n{news}')
    | llm
    | StrOutputParser()
)

category_chain = ()

final_chain = RunnableParallel({
    'summary' : summary_chain,
    'sentiment': sentiment_chain,
    'category': category_chain,
})

news = '뉴스' 

result = final_chain.invoke({'news':news})
print('원문: {news}')