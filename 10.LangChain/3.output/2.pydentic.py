from dotenv import load_dotenv

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import PydanticOutputParser

from pydantic import BaseModel, Field

load_dotenv()

class MovieReview(BaseModel):
    '''영화 리뷰 분석 결과'''
    title: str = Field(description='영화 제목')
    sentiment: str = Field(description='감성 분류: 긍정, 부정, 중립')
    score: int = Field(description='1~10 점수')
    summary: str = Field(description='리뷰 요약 (1~2 문장)')
    keywords: list[str] = Field(description='핵심 키워드 3개')

llm = ChatOpenAI(model='gpt-4o-mini')

parser = PydanticOutputParser(pydantic_object=MovieReview)
# print('포멧 명령문:')
# print(parser.get_format_instructions())

prompt = ChatPromptTemplate.from_template(
    '''다음 영화 리뷰를 분석해 주세요.
    리뷰: {review}
    {format_instructions}
    '''
)

chain = prompt | llm | parser 

reviews = [
    '우주를 배경으로 한 SF 영화인데도 인간적인 감정선이 굉장히 섬세하게 살아있다. 라이언 고슬링의 연기가 몰입감을 끌어올렸고, 후반부 전개는 생각보다 훨씬 감동적이었다. 과학적인 설정과 대중성이 균형을 잘 맞춘 작품.',
    '스타워즈 팬이라면 반가운 요소가 많지만, 전체적인 스토리는 다소 안전하게 흘러간다. 액션과 비주얼은 훌륭했지만 예전 시리즈만큼의 강렬한 임팩트는 부족했다. 그래도 그로구의 매력 하나만으로 충분히 볼 가치가 있다',
    '저예산 공포영화라는 게 믿기지 않을 정도로 긴장감 연출이 뛰어나다. 단순히 놀래키는 방식이 아니라 심리적으로 압박해오는 분위기가 인상적이었다. 올해 가장 의외의 화제작이라는 말이 괜히 나온 게 아니다'
]

for review in reviews:
    result = chain.invoke({
        'review': review,
        'format_instructions':parser.get_format_instructions()
    })

    print(f'제목: {result.title}')
    print(f'감성: {result.sentiment} (점수: {result.score}/10)')
    print(f'요약: {result.summary}')
    print(f'키워드: {result.keywords}')
    print('-' * 30)