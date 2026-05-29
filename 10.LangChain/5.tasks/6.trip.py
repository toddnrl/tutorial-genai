# 목적 - 여행계획 작성
# 도시 작성 -> 음식 추천 관광지 추천 호텔 추천
# 

# runnableparallel, branch
from dotenv import load_dotenv
from langchain_core.prompts import ChatPromptTemplate # LangChain에서 프롬프트를 만들기 위한 도구
from langchain_core.runnables import RunnableParallel, RunnableBranch
from langchain_core.output_parsers import StrOutputParser
from langchain_openai import ChatOpenAI

load_dotenv()

llm = ChatOpenAI(model='gpt-4o-mini', temperature=0.7)

parser = StrOutputParser()


# 공통 
food_prompt = ChatPromptTemplate.from_template(
    '{city} 여행에서 꼭 먹어야 할 대표 음식 3가지를 추천해줘'
)
food_chain = food_prompt | llm | parser 

# 장소 추천
spot_prompt = ChatPromptTemplate.from_template(
    '{city} 여행에서 꼭 가봐야 할 관광지 3곳을 추천해줘'
)
spot_chain = spot_prompt | llm | parser 

# 호텔 추천
##

# parallel 음식, 관광지, 호텔 동시 실행

travel_parallel_chain = RunnableParallel(
    food = food_chain,
    spot = spot_chain
)

# 일본 전용 체인
japan_prompt = ChatPromptTemplate.from_template(
    '{city}는 일본 도시입니다. 일본 여행 초보자 기준 여행 팁을 한 문장으로 알려줘'
)
japan_chain = RunnableParallel(
    tip = japan_prompt | llm | parser,
    plan = travel_parallel_chain
)

europe_prompt = ChatPromptTemplate.from_template(
    '{city}는 유럽 도시입니다. 유럽 여행 팁을 한 문장으로 알려줘'
)
europe_chain = RunnableParallel(
    tip = europe_prompt | llm | parser,
    plan = travel_parallel_chain
)

default_prompt = ChatPromptTemplate.from_template(
    '{city} 여행을 준비하는 사람에게 기본 여행 팁을 한 문장으로 알려줘'
)
default_chain = RunnableParallel(
    tip = default_prompt | llm | parser,
    plan = travel_parallel_chain
)


branch_chain = RunnableBranch(
    (
        lambda x: x['city'] in ['도쿄', '오사카', '교토', '후쿠오카'],
        japan_chain
    ),
    (
        lambda x: x['city'] in ['파리', '런던', '로마', '베네치아'],
        europe_chain
    ),
    default_chain
)


result = branch_chain.invoke({
    'city' : '도쿄'
})

print('여행 팁:', result['tip'])
print('\n음식 추천:', result['plan']['food'])
print('\n관광지 추천:', result['plan']['spot'])