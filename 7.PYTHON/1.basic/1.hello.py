print('hello, python')
print('hello, ' + 'python')
print('"hello," ' + 'python')
num = 5
name = "홍길동"
print("hello, {}".format(name))
print("hello, {}.my lucky number is {}".format(name, num))
print("hello, {0}.my lucky number is {1}".format(name, num))
print("hello, {1}.my lucky number is {0}".format(name, num))
print("hello, %s" % name)
print("hello, %s" % name, end="")
print("hello, %s" % name, end="1")
print("hello, %s" % name, end=",")
"""
멀티라인 문자열
주석보단 여러줄의 문자열
"""
multyline = """이상욱"""

print(multyline)