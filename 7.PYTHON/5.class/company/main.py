from employee import Employee
from person import Person
from driver import Driver

employee1 = Employee("James", 25, "Samsung")

employee2 = Employee("Sang", 27, "LG")

employee3 = Person("Alice", 23)

employee4 = Driver("홍길동", 40, "BMW")


employee1.greet()
employee2.greet()
employee3.greet()

employee3.get_age(40)
employee3.greet()
print(employee3.get_name())
employee4.greet()

employee4.driver()