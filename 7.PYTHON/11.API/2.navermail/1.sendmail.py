import smtplib, os
from email.mime.text import MIMEText

from dotenv import load_dotenv

load_dotenv()

SMTP_SERVER = 'smtp.naver.com'
SMTP_PORT = 587

NAVER_ID = os.getenv('NAVER_MAIL_ID')
NAVER_PASSWORD = os.getenv('NAVER_MAIL_APP_SECRET')
NAVER_EMAIL = f'{NAVER_ID}@naver.com'

subject = "네이버 메일 보내기 테스트중"
body = "<body><h1>심각한 컴퓨터 감염 발쌩!!!!!!!!!!!!! 원인은 땅콩!!!!!!!!!!!!!!!!</h1></body>"

message = MIMEText(body,'html', _charset='utf-8')
message['Subject'] = subject
message['From'] = NAVER_EMAIL
message['To'] = NAVER_EMAIL

try:

    smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    smtp.starttls()
    smtp.login(NAVER_ID, NAVER_PASSWORD)
    smtp.sendmail(NAVER_EMAIL, message['TO'], message.as_string())
    print('메일이 성공적으로 전달되었습니다')
except Exception as e :
    print(f'메일 전송중 오류 발생: {e}')
finally:
    smtp.quit() 