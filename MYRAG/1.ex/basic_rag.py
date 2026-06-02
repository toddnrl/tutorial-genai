import os
import numpy as np
import faiss

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

client = OpenAI(api_key=os.getenv('OPENAPI_AI_KEY'))


# 문서 데이터 준비
documents = [
    "SPC는 한국소프트웨어저작권협회의 약자입니다.",
    "Python은 웹 개발, 데이터 분석, AI 개발에 많이 사용됩니다.",
    "RAG는 검색 증강 생성이라는 뜻입니다.",
    "RAG는 질문과 관련된 문서를 먼저 찾고, 그 내용을 바탕으로 답변합니다.",
    "FAISS는 벡터 검색을 빠르게 해주는 라이브러리입니다."
]

# 임베딩 함수
def get_embedding(text):
    response = client.embeddings.create(
        model="text-embedding-3-small",
        input=text
    )

    return np.array(response.data[0].embedding, dtype="float32")

# 문서들을 벡터로 변환
embeddings = []

for doc in documents:
    emb = get_embedding(doc)
    embeddings.append(emb)

embeddings = np.array(embeddings)

# FAISS 벡터 DB 생성

dimension = embeddings.shape[1]

index = faiss.IndexFlatL2(dimension)
index.add(embeddings)

print("벡터 저장 완료")
print("저장된 문서 개수:", index.ntotal)

