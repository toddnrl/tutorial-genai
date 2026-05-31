from langchain_community.document_loaders import PyPDFLoader

loader = PyPDFLoader('./Javascript_Secure_Coding.pdf')
docs = loader.load()

print(f'PDF 페이지수 : {len(docs)}\n')

