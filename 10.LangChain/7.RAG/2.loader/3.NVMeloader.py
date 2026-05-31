import os
from dotenv import load_dotenv

from openai import OpenAI

from langchain_community.document_loaders import TextLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS

load_dotenv()

# 1. TXT 로드
loader= TextLoader('./NVMe.txt', encoding='utf-8')
documents = loader.load()

print("문서 개수:", len(documents))
print("문서 앞부분:", documents[0].page_content[:200])
print("metadata:", documents[0].metadata)


# 2. 청크 분할
splitter = RecursiveCharacterTextSplitter(
    chunk_size = 500,
    chunk_overlap = 100
)
chunks = splitter.split_documents(documents)
print("청크 개수:", len(chunks))
print("첫 청크 글자수:", len(chunks[0].page_content))


# 3. 임베딩 + FAISS 저장
embeddings = OpenAIEmbeddings(
    model="text-embedding-3-small"
)

db = FAISS.from_documents(
    chunks,
    embeddings
)
print("FAISS 저장 완료")

# 4. 질문 검색
questions = 'NVMe의 장점은 뭐야?'

docs = db.similarity_search(
    questions,
    k=3
)

print('\n검색된 문서: ')
for i, doc in enumerate(docs, start=1):
    print(f'\n--검색--결과-- {i}')
    print(doc.page_content)

# 5. GPT 답변

client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

context = '\n\n'.join([doc.page_content for doc in docs])

prompt = f'''
아래 문서를 참고해서 질문에 답변해줘.

문서:
{context}

질문:
{questions}
'''

response = client.chat.completions.create(
    model='gpt-4o-mini',
    messages=[
        {'role':'user', 'content':prompt}
    ]
)
print('\nGPT 답변: ')
print(response.choices[0].message.content)