import wikipedia

from dotenv import load_dotenv

from langchain_openai import ChatOpenAI
from langchain_community.agent_toolkits.load_tools import load_tools
from langchain_community.utilities.wikipedia import WikipediaAPIWrapper
from langchain_community.tools.wikipedia.tool import WikipediaQueryRun
from langchain.agents import create_agent  

load_dotenv()


wiki_en = WikipediaQueryRun(
    api_wrapper=WikipediaAPIWrapper(lang='en', top_k_results=3,
                                    doc_content_chars_max=200, description='English wikipedia')

)



llm = ChatOpenAI(model='gpt-4o-mini')

system_prompt = '''
    당신은 위키피디아를 활용해 정보를 조회하고 답변하는 챗봇입니다
    영어 검색 결과인 경우 한국어로 번역해서 답하세요
'''


agent = create_agent(llm, [wiki_en] , system_prompt=system_prompt)
quuestion = ['파이썬은 누가 만든거야']

result = agent.invoke({'messages':[('user', quuestion[0])]})
