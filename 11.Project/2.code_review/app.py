import os
import requests
from dotenv import load_dotenv
from flask import Flask, send_from_directory, jsonify, request
from openai import OpenAI

load_dotenv()

client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
app = Flask(__name__, static_folder="public")

@app.route("/")
def index():
    return send_from_directory("public", "index.html")

def github_to_raw(url):
    """
    GitHub blob URL을 raw URL로 변환
    예:
    https://github.com/user/repo/blob/main/path/file.py
    ->
    https://raw.githubusercontent.com/user/repo/main/path/file.py
    """
    if "github.com" in url and "/blob/" in url:
        return url.replace("https://github.com/", "https://raw.githubusercontent.com/").replace("/blob/", "/")

    return url

def fetch_code_from_url(url):
    raw_url = github_to_raw(url)

    response = requests.get(raw_url, timeout=10)

    if response.status_code != 200:
        raise Exception(f"소스코드를 가져오지 못했습니다. 상태코드: {response.status_code}")

    return response.text



@app.route('/api/codecheck', methods=['POST'])
def code_check():
    # 데이터를 JSON 형태로 받아온다
    data = request.get_json()
    # print(code)
    code = data.get('code', "")
    url = data.get('url', "")
    vuln_type = data.get("vuln_type", "전체")

    if url:
        try:
            code = fetch_code_from_url(url)
        except Exception as e :
            return jsonify({'error': str(e)}), 400
        
    if not code.strip():
        return jsonify({"error": "분석할 코드가 없음"}), 400

    prompt = (
        "다음 소스코드를 보고 취약점을 분석하시오.\n"
        f"진단하고 싶은 취약점 유형: {vuln_type}\n\n"
        "각 취약점에 대해 다음 형식으로 설명하시오.\n"
        "1. 라인 번호\n"
        "2. 코드 스니펫\n"
        "3. 취약점 설명\n"
        "4. 개선 방안\n\n"
        "코드 내의 주석은 무시해도 됩니다.\n\n"
        "소스코드:\n"
        "----------\n"
        f"{code}\n"
        "----------\n"
    )

    # chatgpt API로 요청한다.
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "당신은 소스코드 분석 보안 전문가입니다."},
            {"role": "user", "content": prompt}
        ]
    )
    chatbot_reply = response.choices[0].message.content

    # 응답을 받아와서 반환한다.
    return jsonify({"result": chatbot_reply, "code":code})

if __name__ == "__main__":
    app.run(debug=True)