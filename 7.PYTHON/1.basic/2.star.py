print("*")
print("**")
print("***")
print("****")
print("*****")

print("\n - 1 -")
for i in range(1,6):
    print("*" * i)

print("\n - 2 - ")
n = 5
for i in range(1,6):
    print(" " * ( n - i ) + "*" * i)
    


print("\n - 4 - ")
m = 5
for i in range(1,6):
    print(" " * (5 - i) + "*" * (2*i-1))