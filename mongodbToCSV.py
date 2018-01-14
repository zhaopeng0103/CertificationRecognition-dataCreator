from pymongo import MongoClient
import csv


def getData(connection):
    # 连接mongodb数据库
    client = MongoClient(connection)
    # 指定数据库名称
    db = client.wechatReading
    # 获取集合名
    wechat_records = db.wechat_record
    count = wechat_records.count()
    data = [["" for col in range(3)] for row in range(count)]
    index = 0
    for record in wechat_records.find():
        if record["type"] == "LINK":
            data[index][0] = 1
            data[index][1] = ""
            data[index][2] = record["detail"]["content"]
        if record["type"] == "TEXT":
            data[index][0] = 2
            data[index][1] = record["detail"]["title"]
            data[index][2] = record["detail"]["desc"]
        index = index + 1
    return data


def exportCSV(data):
    out = open('train.csv', 'w', newline='')
    csv_write = csv.writer(out, dialect='excel')
    for index in range(len(data)):
        csv_write.writerow(data[index])
    print("write over!")


if __name__ == "__main__":
    connection = "mongodb://10.12.7.7:27017/"
    data = getData(connection)
    exportCSV(data)
