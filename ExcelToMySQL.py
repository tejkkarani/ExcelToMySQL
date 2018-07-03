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
add_Student = ("INSERT INTO ieee " "(FormNo, MembershipDate, Status, Surname, FirstName, MiddleName, MothersName, Year, Class, RollNo, RegNo, DOB, PassingYear, EmailID, PhoneNo, Address, Pincode) " "VALUES (%s, %s, %s, %s,%s, %s, %s, %s,%s, %s, %s, %s,%s, %s, %s, %s, %s)")
data_Student = (values[0]['Form No.'], values[0]['Date of Membership Taken'], values[0]['Status'], values[0]['Surname'], values[0]['First Name'], values[0]["Father's Name"], values[0]["Mother's Name"], values[0]['Year'], values[0]['Class(Only DIV)'], values[0]['Roll No.'], values[0]['Registration No.'], values[0]['Date Of Birth'], values[0]['Year Of Passing'], values[0]['Email ID'], values[0]['Contact No.'], values[0]['Address'], values[0]['Pincode'])
cursor.execute(add_Student, data_Student)
cnx.commit()