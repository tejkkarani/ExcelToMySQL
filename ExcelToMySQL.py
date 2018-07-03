from xlrd import open_workbook
from pprint import pprint
import mysql.connector
from mysql.connector import errorcode,Error
cnx = mysql.connector.connect(user='root', password='', host='localhost', port='3306', database='Student')
cursor = cnx.cursor()
wb = open_workbook('test.xlsx')
values = []
for s in wb.sheets():
    for row in range(1, s.nrows):
        col_names = s.row(0)
        col_value = {}
        for name, col in zip(col_names, range(s.ncols)):
            value = s.cell(row,col).value
            try:
                value = str(int(value))
            except :
                pass
            col_value[name.value]=value
        values.append(col_value)
pprint(values)