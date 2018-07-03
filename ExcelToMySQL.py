from xlrd import open_workbook
import xlrd
import datetime
import mysql.connector
# change the username, password, host, port and database below
cnx = mysql.connector.connect(user='root', password='', host='localhost', port='3306', database='loginsystem')
cursor = cnx.cursor()
wb = open_workbook('test.xlsx')
values = []
for s in wb.sheets():
    for row in range(1, s.nrows):
        col_names = s.row(0)
        col_value = {}
        for name, col in zip(col_names, range(s.ncols)):
            value = s.cell(row,col).value
            if type(value) is float:
                if len(str(value)) == 7 and name.value!='Registration No.':
                    value = datetime.datetime(*xlrd.xldate_as_tuple(value, wb.datemode))
                    value = value.strftime('%Y-%m-%d')
                else:
                    value = int(value)
            col_value[name.value]=value
        values.append(col_value)
for data in values:
    # Check name of the table columns matches with this
    add_Student = "INSERT INTO users (user_firstname, user_lastname, user_fathersname, user_mothersname, user_class, user_division, user_rollno, user_dob, user_yearofpassing, user_email, user_pass, user_contact, user_address, user_pincode) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    data_Student = (data['First Name'], data['Surname'], data["Father's Name"], data["Mother's Name"], data['Year'], data['Class(Only DIV)'], data['Roll No.'], data['Date Of Birth'], data['Year Of Passing'], data['Email ID'], data['Registration No.'], data['Contact No.'], data['Address'], data['Pincode'])
    cursor.execute(add_Student, data_Student)
    cnx.commit()
