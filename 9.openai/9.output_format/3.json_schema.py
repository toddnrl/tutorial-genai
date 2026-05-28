import os 
import json

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

# 내가 원하는 출력 형식 - 즉 자료구조를 정의

city_schema = {
    'type': 'object',
    'properties':{
        'name': {'type': 'string'},
        'population':{'type':'integer'},
        'area_km2':{'type':'number'},
    },
    'required': ['name','population', 'area_km2'],
    'additionalProperties':False,
}

response = client.chat.completions.create(
    model='gpt-4o-mini',
    messages=[
        {'role':'system', 'content':'질문에 대해 JSON으로만 답변하시오'},
        {'role':'user', 'content':'서울의 인구와 면적을 알려주세요'}
    ],
    response_format={
        'type':'json_schema',
        'json_schema' : {
            'name':'city_info',
            'strict' :True,
            'schema' : city_schema
        }
        
    }
)

answer = response.choices[0].message.content
print(answer)

data = json.loads(answer)
print(f'도시의 이름은: {data['name']} - 인구 : {data['population']:,}명, 면적 : {data['area_km2']}km2')