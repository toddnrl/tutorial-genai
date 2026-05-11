student = {
    "이상욱":100,
    "ㄴ상욱":80,
    "ㄷ상욱":40,
    "ㄹ상욱":70,
    "ㅁ상욱":83,
    "ㅂ상욱":55,
    "ㅅ상욱":64,
    "ㅇ상욱":57,
    "ㅈ상욱":78,
    "ㅊ상욱":88
}

print(student)

def get_a_student(student):
    a_student = []
    for name, score in student.items():
        if score >= 90:
            a_student.append(name)
    return a_student

print(get_a_student(student))