#!/usr/bin/env python3
# encoding=UTF-8
import os
import xlwt
import uuid
from data import produceData

fontUnderline = [xlwt.Font.UNDERLINE_NONE, xlwt.Font.UNDERLINE_SINGLE, xlwt.Font.UNDERLINE_SINGLE_ACC,
                 xlwt.Font.UNDERLINE_DOUBLE, xlwt.Font.UNDERLINE_DOUBLE_ACC]
fontEscapement = [xlwt.Font.ESCAPEMENT_NONE, xlwt.Font.ESCAPEMENT_SUPERSCRIPT, xlwt.Font.ESCAPEMENT_SUBSCRIPT]
fontFamily = [xlwt.Font.FAMILY_NONE, xlwt.Font.FAMILY_ROMAN, xlwt.Font.FAMILY_SWISS, xlwt.Font.FAMILY_MODERN,
              xlwt.Font.FAMILY_SCRIPT, xlwt.Font.FAMILY_DECORATIVE]
cellBorders = [xlwt.Borders.NO_LINE, xlwt.Borders.THIN, xlwt.Borders.MEDIUM, xlwt.Borders.DASHED, xlwt.Borders.DOTTED,
               xlwt.Borders.THICK, xlwt.Borders.DOUBLE, xlwt.Borders.HAIR, xlwt.Borders.MEDIUM_DASHED,
               xlwt.Borders.THIN_DASH_DOTTED, xlwt.Borders.MEDIUM_DASH_DOTTED, xlwt.Borders.THIN_DASH_DOT_DOTTED,
               xlwt.Borders.MEDIUM_DASH_DOT_DOTTED, xlwt.Borders.SLANTED_MEDIUM_DASH_DOTTED]
colorIndex = { "Black": 0, "White" : 1, "Red" : 2, "Green" : 3, "Blue" : 4, "Yellow" : 5, "Magenta" : 6,
               "Cyan" : 7, "Maroon" : 16, "Dark Green" : 17, "Dark Blue" : 18, "Dark Yellow" : 19,
               "Dark Magenta" : 20, "Teal" : 21, "Light Gray" : 22, "Dark Gray" : 23 }
start_index = 2
merge_num = 4

def getResultFromDB(host, username, password, db, sqlStr):
    # 连接数据库，执行sql
    result = os.popen('mysql -h' + host + ' -u' + username + ' -p' + password + ' -D' + db + ' -e "' + sqlStr + '"').read().strip().split('\n')
    # 获取列名
    datas = result.split('\t')
    return result


def createWorkbook():
    # 创建一个excel工作簿，编码utf-8，表格中支持中文
    workbook = xlwt.Workbook(encoding='utf-8')
    return workbook


def saveWorkbook(workbook):
    workbook.save("output\\" + str(uuid.uuid4()) + ".xls")
    print("Save OK")


def createSheet(workbook):
    sheet = workbook.add_sheet('sheet cert')
    return sheet


def setStyle(fontHeight, colorIndex, fontBold, fontUnderline, fontEscapement, fontFamily, cellBorders):
    style = xlwt.XFStyle()

    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.height = fontHeight
    font.colour_index = colorIndex
    font.bold = fontBold
    font.italic = False
    font.struck_out = False
    font.underline = fontUnderline
    font.escapement = fontEscapement
    font.family = fontFamily

    borders = xlwt.Borders()
    borders.left = cellBorders
    borders.right = cellBorders
    borders.top = cellBorders
    borders.bottom = cellBorders

    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER

    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def writeData(sheet, rowKey, dataValue, titleStyle, keyStyle, valueStyle):
    rows = len(rowKey)

    # 设置列的宽度
    sheet.col(0).width = 2000
    sheet.col(1).width = 5000
    sheet.col(2).width = 20000
    sheet.col(3).width = 4000
    sheet.col(4).width = 5000
    sheet.col(5).width = 20000
    sheet.col(6).width = 2000

    sheet.write_merge(start_index, start_index + merge_num - 1, 1, 2, "  12  （2017）  苏州市  不动产权第  " + dataValue[3].split(" ")[3] + "  号  ", titleStyle)
    for i in range(rows - 1):
        sheet.write_merge(merge_num*(i + 2) - 2, merge_num*(i + 3) - 3, 1, 1, rowKey[i], keyStyle)
        sheet.write_merge(merge_num*(i + 2) - 2, merge_num*(i + 3) - 3, 2, 2, dataValue[i], valueStyle)
    sheet.write_merge(merge_num*(rows + 1) - 2, merge_num*(rows + 4) - 3, 1, 1, rowKey[rows - 1], keyStyle)
    sheet.write_merge(merge_num*(rows + 1) - 2, merge_num*(rows + 4) - 3, 2, 2, dataValue[rows - 1], valueStyle)

    sheet.write_merge(start_index, start_index + merge_num - 1, 4, 5, "附    记", titleStyle)
    sheet.write_merge(start_index + merge_num, merge_num*(rows + 4) - 3, 4, 5, "1、" + dataValue[rows], valueStyle)

    sheet.write_merge(merge_num*(rows + 4) - 2, merge_num*(rows + 4), 0, 6, "", titleStyle)


if __name__ == "__main__":
    host = 'localhost'
    username = 'root'
    password = 'mimanicaibudao'
    db = 'cert'
    sqlStr = 'SELECT * FROM RealEstateCert;'

    for index in range(2):
        rowKey = ['权利人', '共有情况', '坐落', '不动产单元号', '权利类型', '权利性质', '用途', '面积', '使用期限', '权利其他情况']
        # dataValue = ['赵鹏', '共同共有', '彭义里40号', '320508 024003 GB01700 12457896', '国有建设用地使用权/房屋（构建物）所有权', '划拨', '城镇住宅用地/成套住宅', '分摊土地面积24.80平方米/房屋建筑面积63.7平方米', '截止到2030年12月31日', '房屋结构：混合']
        dataValue = produceData()

        titleStyle = setStyle(350, colorIndex["Black"], True, fontUnderline[0], fontEscapement[0], fontFamily[0], cellBorders[0])
        keyStyle = setStyle(300, colorIndex["Black"], True, fontUnderline[0], fontEscapement[0], fontFamily[0], cellBorders[2])
        valueStyle = setStyle(300, colorIndex["Black"], False, fontUnderline[0], fontEscapement[0], fontFamily[0], cellBorders[2])

        workbook = createWorkbook()
        sheet = createSheet(workbook)
        writeData(sheet, rowKey, dataValue, titleStyle, keyStyle, valueStyle)
        saveWorkbook(workbook)
