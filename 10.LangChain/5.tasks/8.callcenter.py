from dotenv import load_dotenv

from langchain_openai import ChatOpenAI

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnableBranch


load_dotenv()

llm = ChatOpenAI(model='gpt-4o-mini')

def make_chain(role):
    return (
        ChatPromptTemplate.from_messages([
            ('system', role),
            ('user', '{question}')
        ])
        | llm
        | StrOutputParser()
    )

payment_chain = make_chain('당신은 상담원 입니다 친절하게 안내하세요')

delivary_chain = make_chain('당신은 배송 상담원 입니다 친절하게 안내하세요')

general_chain = make_chain('당신은 일반 상담원 입니다 친절하게 안내하세요')

branch = RunnableBranch(
    (lambda x : any(k in x['question'] for k in ['결제', '한불', '청구']), payment_chain),
    (lambda x : any(k in x['question'] for k in ['배송', '택배', '반품']), delivary_chain),
    general_chain,
)

questions = [
    '배송이 안돼요 언제쯤 도착할까요',
    '결제가 두번 됐는데 환불 될까요'
]

for q in questions :
    print('-' * 60)
    print(f'고객: {q}')
    print(f'상담원: {branch.invoke({'question': q})}')