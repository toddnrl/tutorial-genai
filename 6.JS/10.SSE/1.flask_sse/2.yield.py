

def test():
    print('A')
    yield 1   # 여기에서 멈춘다

    print('B')
    yield 2   # 여기에서 멈춘다

    print('C')
    yield 3   # 여기에서 멈춘다
   

x = test()  # generator 라는 것이 만들어짐 - 동적으로 바뀌는 데이터를 전달하는 객체

# print(x)
# print(next(x))
# print(next(x))
# print(next(x))

try:
    while True:
        print(next(x))

except StopIteration:
    print("모든 데이터 사용 완료")