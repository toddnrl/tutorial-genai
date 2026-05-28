import os 

import json

from pydantic import BaseModel

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))


class CityInfo(BaseModel):
    name: str
    population: int
    area_km2: float


response = client.chat.completions.parse(
    model='gpt-4o-mini',
    messages=[
        {'role':'system', 'content':'질문에 대해 JSON으로만 답변하시오'},
        {'role':'user', 'content':'서울의 인구와 면적을 알려주세요'}
    ],
    response_format=CityInfo
)

answer = response.choices[0].message.parsed
print(answer)

data =answer
print(f'도시의 이름은: {data.name} - 인구 : {data.population}명, 면적 : {data.area_km2}km2')
