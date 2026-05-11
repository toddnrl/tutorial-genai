import math
import random

print(math.pi)

print(math.sqrt(16))

print(math.sin(0))
print(math.sin(math.pi))



import datetime
fruits = ['apple', 'banana', 'cherry', 'grape', 'orange', 'pinapple']
def pick_fruits():
    my_number = random.randint(0,len(fruits) - 1)
    my_pick = fruits[my_number]
    return my_pick


def pick_fruits2():
    return random.choice(fruits)

print("내 과일은 : ", pick_fruits())

print("내 과일은 : ", pick_fruits2())