#
# author:Owen
# import excel to mysql
#

import pymysql
# xlrd 为 python 中读取 excel 的库，支持.xls 和 .xlsx 文件
import xlrd

# openpyxl 库支持 .xlsx 文件的读写
from openpyxl.reader.excel import load_workbook
from builtins import int


# cur 是数据库的游标链接，path 是 excel 文件的路径
def importExcelToMysql(cur, path):

    ### xlrd版本
    # 读取excel文件
     workbook = xlrd.open_workbook(path)
     sheets = workbook.sheet_names()
    # 读取excel的第一个表
     worksheet = workbook.sheet_by_name(sheets[0])

    # 将表中数据读到 sqlstr 数组中
    # 第二行开始
     for i in range(1, worksheet.nrows):
         #row = worksheet.row(i)

         sqlstr = []

         for j in range(0, worksheet.ncols):
             print(worksheet.cell_value(i, j))
             sqlstr.append(worksheet.cell_value(i, j))

     valuestr = [str(sqlstr[0]), str(sqlstr[1]), str(sqlstr[2]), str(sqlstr[3]), str(sqlstr[4])]

     cur.execute("insert into t(品号, 品名, 规格, 单位,会计) values(%s, %s, %s, %s, %s)", valuestr)

    ### openpyxl版本
    # 读取excel文件
    # workbook = load_workbook(path)
    # 获得所有工作表的名字
    # sheets = workbook.get_sheet_names()
    # 获得第一张表
    # worksheet = workbook.get_sheet_by_name(sheets[0])

    # 将表中每一行数据读到 sqlstr 数组中
    # for row in worksheet.rows:

       # sqlstr = []

       # for cell in row:
       #    sqlstr.append(cell.value)


# 输出数据库中内容
def readTable(cursor):
    # 选择全部
    cursor.execute("select * from t")
    # 获得返回值，返回多条记录，若没有结果则返回()
    results = cursor.fetchall()

    for i in range(0, results.__len__()):
        for j in range(0, 5):
            print(results[i][j], end='\n')
        print('\n')


if __name__ == '__main__':
    # 和数据库建立连接
    conn = pymysql.connect('127.0.0.1', 'root', '123456', charset='utf8')
    # 创建游标链接
    cur = conn.cursor()

    # 新建一个database
    # 删除数据库
    # cur.execute("drop database if exists tb")
    # 如果不存在数据库tb就创建
    cur.execute("create database if not exists tb")
    # 选择 tb 这个数据库
    cur.execute("use tb")

    # sql中的内容为创建一个名为t的表
    sql = """CREATE TABLE IF NOT EXISTS `t` (
                `品号` VARCHAR (80),
                `品名` VARCHAR (80),
                `规格` VARCHAR (80),
                `单位` VARCHAR (80),
                `会计` VARCHAR (80)
              )"""
    # 如果存在t这个表则删除
    cur.execute("drop table if exists t")
    # 创建表
    cur.execute(sql)

    # 将 excel 中的数据导入数据库中
    importExcelToMysql(cur, "newerp.xls")
    readTable(cur)

    # 关闭游标链接
    cur.close()
    conn.commit()
    # 关闭数据库服务器连接，释放内存
    conn.close()