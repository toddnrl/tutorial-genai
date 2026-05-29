# 목적 - 뉴스를 분석한다
# 뉴스입력 - 요약 - 감성분석 - 카테고리 분석
# runnableparallel
from dotenv import load_dotenv
from langchain_core.prompts import (
    ChatPromptTemplate, # LangChain에서 프롬프트를 만들기 위한 도구
    HumanMessagePromptTemplate, # 시스템 메시지와 사용자 메시지를 묶어서 하나의 채팅 프롬프트로 만드는 클래스
    SystemMessagePromptTemplate, # 사용자가 입력하는 메시지 형식의 프롬프트를 만드는 클래스
    AIMessagePromptTemplate # AI의 역할이나 성격을 정하는 시스템 메시지를 만드는 클래스
)

from langchain_openai import ChatOpenAI
from langchain_core.output_parsers import PydanticOutputParser

from pydantic import BaseModel, Field # AI가 어떤 구조로 답해야 하는지 클래스로 정의하기 위해

load_dotenv()

class NewsAnalysis(BaseModel):
    '''뉴스 분석 결과'''
    summary: str = Field(description='뉴스 요약 1문장')
    sentiment: str = Field(description='감성 분류: 긍정, 부정, 중립')
    category: str = Field(description='뉴스 카테고리: 경제, 정치, 사회, IT, 국제, 문화')
    keywords: list[str] = Field(description='핵심 키워드 3개')

llm = ChatOpenAI(model='gpt-4o-mini', temperature=0.1)

parser = PydanticOutputParser(pydantic_object=NewsAnalysis)
# AI 응답을 NewsAnalysis 형식으로 변환할 파서를 만듭니다.


prompt = ChatPromptTemplate.from_template(
    '''
    다음 뉴스를 분석해주세요

    뉴스:{article}  

    분석 항목: 
    - 한 문장 요약
    - 감성 분석
    - 뉴스 카테고리
    - 핵심 키워드

    {format_instructions}
    '''
)

chain = prompt | llm | parser

input_text = {
    'article':'올해 정부가 출자기관으로부터 2조 8,000억 원에 육박하는 역대 최대 규모의 배당금을 받았다. '
    '3대 국책은행이 전체 배당을 이끈 가운데 평균 배당성향 역시 정부의 목표치인 40%를 넘어서며 최고치를 기록했다. '
    '재정경제부는 28일 올해 40개 정부출자기관 중 20개 기관에 대한 정부 배당액이 전년 대비 4,964억 원 증가한 2조 7,951억 원으로 확정됐다고 밝혔다. '
    '당기순이익 대비 총배당금인 평균 배당성향은 40.90%로 1년 전보다 1.18%포인트(p) 상승했다. '
    '이는 배당액과 배당성향 모두 역대 최대 규모로, 정부 출자기관 배당성향 목표치인 40%를 달성한 수치다. '
    '실적을 이끈 것은 3대 국책은행이다. 한국산업은행, 중소기업은행, 한국수출입은행의 정부 배당은 1조 9,536억 원으로 전체 배당액의 69.9%를 차지했다. '
    '기관별로는 산업은행의 배당액이 8,806억 원으로 가장 많았다. 여기에는 현금 출자한 정책금융 모펀드 회수자금 2,494억 원이 포함됐다. '
    '이어 중소기업은행 5,968억 원, 한국수출입은행 4,762억 원, 인천국제공항공사 3,194억 원 순으로 큰 규모를 기록했다. '
    '에너지 공기업들의 배당 참여도 두드러졌다. '
    '장기간 적자 늪에 빠져있던 한국전력공사는 지난해 산업용 전기요금 인상 등으로 역대 두 번째로 많은 당기순이익을 거두며 '
    '1,802억 원을 지급해 2년 연속 배당을 이어갔다. 한국가스공사 279억 원과 한국지역난방공사 246억 원도 각각 배당에 참여했다.'
}

result = chain.invoke({
    'article': input_text['article'],   # prompt 안에 article 빈칸에 이 값이 들어감
    'format_instructions': parser.get_format_instructions() # 여기도 동일
})


print('요약:', result.summary)
print('감성:', result.sentiment)
print('카테고리:', result.category)
print('키워드:', result.keywords)

