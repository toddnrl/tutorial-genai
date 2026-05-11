

sum = add_numbers(3,4)
print(f"두 수의 합은 {sum}입니다")

def add_numbers2(a,b):
    return a, b, a+b

input1, input2, 



def calculate_all(a,b):
    add = a+b
    sub = a - b
    mult = a * b
    divi = a / b 

    return add, sub, mult, divi

add, _, mult, _ = calculate_all(3,4)
print(f"덧셈은 {add}, 곱셈은 {mult}")

