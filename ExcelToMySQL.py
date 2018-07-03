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
                    value = int(value)
            col_value[name.value]=value
        values.append(col_value)
pprint(values)
for data in values:
    add_Student = ("INSERT INTO ieee " "(FormNo, MembershipDate, Status, Surname, FirstName, MiddleName, MothersName, Year, Class, RollNo, RegNo, DOB, PassingYear, EmailID, PhoneNo, Address, Pincode) " "VALUES (%s, %s, %s, %s,%s, %s, %s, %s,%s, %s, %s, %s,%s, %s, %s, %s, %s)")
    data_Student = (data['Form No.'], data['Date of Membership Taken'], data['Status'], data['Surname'], data['First Name'], data["Father's Name"], data["Mother's Name"], data['Year'], data['Class(Only DIV)'], data['Roll No.'], data['Registration No.'], data['Date Of Birth'], data['Year Of Passing'], data['Email ID'], data['Contact No.'], data['Address'], data['Pincode'])
    cursor.execute(add_Student, data_Student)
    cnx.commit()