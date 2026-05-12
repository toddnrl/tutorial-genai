import csv
filename = "data.csv"

# with open(filename, "r") as file :
#     data = file.read()
#     print(data)

data = []

# 옛날 방식 리스트 형태로 복원
with open(filename, "r") as file :
    csv_reader = csv.reader(file)
    for row in csv_reader:
        data.append(row)

print(data)

# 모던 방식 딕셔너리 형태로 복원

data2 = []
with open(filename, "r") as file:
    csv_reader = csv.DictReader(file)
    for row in csv_reader:
        data2.append(row)

print(data2)

