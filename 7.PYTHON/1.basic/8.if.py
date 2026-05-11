print("---if----구문")

score = 70
if score >= 80 :
    print("성적은 A+입니다.")
elif score >= 70 :
    print("성적은 B입니다.")
elif score >= 60 :
    print("성적은 C입니다 ")
else : 
    print("성적은 F입니다")

score = 70
if score >= 80 :
    grade = 'A'
elif score >= 70 :
    grade = 'B'
elif score >= 60 :
    grade = 'C'
else : 
    grade = 'F'

print(f"이 학생의 점수는 {score}이고 학점은 {grade}입니다")


month = 7
if month in [12, 1, 2]:
    print("겨울입니다")
elif month in [3,4,5]:
    print("봄입니다")
elif month in [6,7,8]:
    print("여름입니다")
elif month in [9, 10, 11]:
    print("가을입니다")
else :
    print("잘못된 month입니다.")


month = 3
if month in [12, 1, 2]:
    season = "겨울"
elif month in [3,4,5]:
    season = "봄"
elif month in [6,7,8]:
    season = "여름"
elif month in [9, 10, 11]:
    season = "가을"
else :
    print("잘못된 month입니다.")

print(f"{month}월은 {season}입니다")


height = 177
weight = 80
bmi = weight / ((height / 100) **2)

if bmi < 18.5:
    category = "저체중"
elif bmi < 25:
    category = "정상"
elif bmi < 30:
    category = "과체중"
else:
    category = "비만"

print(f"{category}입니다")

username = ''
password = ''

if username and password :
    if username == 'admin' and password == '1234':
        print("관리자로 로그인 완료 ")
    elif username == 'user' and password == '1234':
        print("일반 사용자로 로그인 돠었습니다")
    else :
        print("잘못된 정보")
else:
    print("정보를 입력하세요")

