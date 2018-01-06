#!/usr/bin/env python3
# encoding=UTF-8
import os
import time
import xlwt
from enum import Enum, unique

@unique
class fontUnderline(Enum):
    xlwt.Font.UNDERLINE_NONE = 1
    xlwt.Font.UNDERLINE_SINGLE = 2
    xlwt.Font.UNDERLINE_SINGLE_ACC = 3
    xlwt.Font.UNDERLINE_DOUBLE = 4
    xlwt.Font.UNDERLINE_DOUBLE_ACC = 5

@unique
class fontEscapement(Enum):
    xlwt.Font.ESCAPEMENT_NONE = 1
    xlwt.Font.ESCAPEMENT_SUPERSCRIPT = 2
    xlwt.Font.ESCAPEMENT_SUBSCRIPT = 3

@unique
class fontFamily(Enum):
    xlwt.Font.FAMILY_NONE = 1
    xlwt.Font.FAMILY_ROMAN = 2
    xlwt.Font.FAMILY_SWISS = 3
    xlwt.Font.FAMILY_MODERN = 4
    xlwt.Font.FAMILY_SCRIPT = 5
    xlwt.Font.FAMILY_DECORATIVE = 6

@unique
class cellBorders(Enum):
    xlwt.Borders.NO_LINE = 1
    xlwt.Borders.THIN = 2
    xlwt.Borders.MEDIUM = 3
    xlwt.Borders.DASHED = 4
    xlwt.Borders.DOTTED = 5
    xlwt.Borders.THICK = 6
    xlwt.Borders.DOUBLE = 7
    xlwt.Borders.HAIR = 8
    xlwt.Borders.MEDIUM_DASHED = 9
    xlwt.Borders.THIN_DASH_DOTTED = 10
    xlwt.Borders.MEDIUM_DASH_DOTTED = 11
    xlwt.Borders.THIN_DASH_DOT_DOTTED = 12
    xlwt.Borders.MEDIUM_DASH_DOT_DOTTED = 13
    xlwt.Borders.SLANTED_MEDIUM_DASH_DOTTED = 14


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
    excelTime = time.strftime("%Y%m%d")
    workbook.save(excelTime + "cert.xlsx")
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


def writeData(sheet, rowKey, dataValue, style):
    rows = len(rowKey)

    # 设置列的宽度
    sheet.col(0).width = 2000
    sheet.col(1).width = 4000
    sheet.col(2).width = 20000
    sheet.col(3).width = 6000
    sheet.col(4).width = 4000
    sheet.col(5).width = 20000
    sheet.col(6).width = 2000


    for i in range(rows - 1):
        sheet.write(i + 2, 1, rowKey[i], style)
        sheet.write(i + 2, 2, dataValue[i], style)
    sheet.write_merge(rows - 1, rows + 2, 1, 1, rowKey[rows - 1], style)
    sheet.write_merge(rows - 1, rows + 2, 2, 2, dataValue[rows - 1], style)

    sheet.write_merge(2, rows + 2, 4, 5, "", style)
    return sheet


if __name__ == "__main__":
    host = 'localhost'
    username = 'root'
    password = 'mimanicaibudao'
    db = 'cert'
    sqlStr = 'SELECT * FROM RealEstateCert;'

    rowKey = ['权利人', '共有情况', '坐落', '不动产单元号', '权利类型', '权利性质', '用途', '面积', '使用期限', '权利其他情况']
    dataValue = ['赵鹏', '共同共有', '彭义里40号', '320508 024003 GB01700 12457896', '国有建设用地使用权/房屋（构建物）所有权', '划拨', '城镇住宅用地/成套住宅', '分摊土地面积24.80平方米/房屋建筑面积63.7平方米', '截止到2030年12月31日', '房屋结构：混合']

    colorIndex = { 0 : "Black", 1 : "White", 2 : "Red", 3 : "Green", 4 : "Blue", 5 : "Yellow", 6 : "Magenta", 7 : "Cyan", 16 : "Maroon", 17 : "Dark Green", 18 : "Dark Blue", 19 : "Dark Yellow", 20 : "Dark Magenta", 21 : "Teal", 22 : "Light Gray", 23 : "Dark Gray" }

    titleStyle = setStyle()
    keyStyle = setStyle(260, colorIndex[0], True, fontUnderline())

    workbook = createWorkbook()
    sheet = createSheet(workbook)
    writeData(sheet, rowKey, dataValue, style)




