class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

    def greet(self):
        print(f"안녕하세요 저는 {self.name}입니다")

person1 = Person("Alice", 25)
person2 = Person("Bob", 27)

person1.greet()
person2.greet()
