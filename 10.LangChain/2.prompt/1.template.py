from langchain_core.prompts import PromptTemplate
from langchain_core.prompts import ChatPromptTemplate

template = "당신은 작명가 입니다. {product} 만드는 회사 이름을 지어주세요 "
#  프롬프트 안에다가 변수명을 집어넣음

prompt = PromptTemplate(input_variables=['product'], template=template)

filled_prompt = prompt.format(product='스마트폰')

print('완성된 프롬프트 :', filled_prompt)

filled_prompt = prompt.format(product='자율주행 자동차')

print('완성된 프롬프트 :', filled_prompt)

print('-' * 50)

test_products = [
    '모바일 게임',
    '로봇 장난감',
    '가방',
    '영어 교육 플랫폼',
    '전기 자전거'
]

for product in test_products:
    final_prompt = prompt.format(product=product)
    print(f"[{product}] {final_prompt}")