# with open("file.txt", "r", encoding="utf-8") as file:
#     data = file.read()
#     print("파일내용:", data)

# file = open("file.txt", "r")
# data = file.read()
# file.close()

# print(data)

with open("file.txt", "r", encoding="utf-8") as file:
    lines = file.readlines()

    for line in lines:
        print("파일내용: ", line)