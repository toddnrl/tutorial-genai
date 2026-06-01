import os
from dotenv import load_dotenv

from flask import Flask, request, jsonify, render_template

from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnablePassthrough

load_dotenv()

UPLOAD_DIR = './uploads'
DB_DIR = './chroma_db'
COLLECTIONS_NAME = 'pdf_rag'

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DB_DIR, exist_ok=True)

embeddings = OpenAIEmbeddings(model='text-embedding-3-small')

store = Chroma(
    collection_name=COLLECTIONS_NAME,
    embedding_function=embeddings,
    persist_directory=DB_DIR
)

retriever = store.as_retriever(
    search_kwargs={"k": 3}
)

def format_docs(docs):
    return "\n\n".join(doc.page_content for doc in docs)

llm = ChatOpenAI(
    model="gpt-4o-mini",
    temperature=0
)

prompt = ChatPromptTemplate.from_messages([
    (
        "system",
        """
당신은 PDF 문서를 기반으로 답변하는 AI입니다.

아래 문서를 참고해서만 답변하세요.

문서:
{context}

문서에 없는 내용은
"문서에서 찾을 수 없습니다."
라고 답변하세요.
"""
    ),
    ("user", "{question}")
])


chain = (
    RunnablePassthrough.assign(
        context=lambda x: format_docs(
            retriever.invoke(x["question"])
        )
    )
    | prompt
    | llm
    | StrOutputParser()
)

# 랭체인 기본 불러오기

# 문서 파서 기본 불러오기 (PyPDFLoader)

# 1. 백터스토어 셋업

# 2. 랭체인 셋업한다 (LCEL)



##########
# Flask 
##########
app = Flask(__name__)

@app.get('/')
def index():
    return render_template('index.html')

@app.post('/upload')
def upload():
    file = request.files.get('file')
    
    if file is None:
        return jsonify({'message': '파일이 없습니다'}), 400
    
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    file.save(file_path)

    loader = PyPDFLoader(file_path)
    docs = loader.load()

    splitter = RecursiveCharacterTextSplitter(
        chunk_size = 500,
        chunk_overlap=100
    )

    chunks = splitter.split_documents(docs)

    for chunk in chunks:
        chunk.metadata['source'] = file.filename

    store.add_documents(chunks)

    return jsonify({
        'message':'업로드 완료',
        'chunks': len(chunks)
    })



@app.post('/ask')
def ask():
    data = request.get_json()
    question = data.get("question", "")

    if not question:
        return jsonify({"answer": "질문이 비어 있습니다."}), 400

    answer = chain.invoke({
        "question": question
    })

    return jsonify({
        "answer": answer
    })


if __name__ == "__main__":
    app.run(debug=True)