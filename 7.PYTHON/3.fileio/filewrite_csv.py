import csv

data = [
    ["Name", "Age", "City"],
    ["John", 23, "Busan"],
    ["Bob", 25, "Seoul"]
]

filename = "data.csv"

with open(filename,"w", newline="") as file:
    csv_writer = csv.writer(file)
    csv_writer.writerows(data)
    # file.write(data)

# 모던방식 딕셔너리로 데이터 관리
data2 = [
    {"Name":"John", "Age":23, "City":"Busan"},
    {"Name":"Bob", "Age":25, "City":"Seoul"}
]

with open(filename, "w", newline="") as file :
    # headers = ["Name", "Age", "City"]
    headers = data2[0].keys()
    csv_writer = csv.DictWriter(file, fieldnames=headers)
    csv_writer.writeheader()
    csv_writer.writerows(data2)

