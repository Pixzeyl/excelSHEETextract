import xlsxwriter as xs
import random

with xs.Workbook('test.xlsx') as file_write:
    work = file_write.add_worksheet('Test')
    head = ['Student ID', 'Name', 'Subject', 'Marks']
    for i,j in enumerate(head):
        work.write(0,i,j)
    for i in range(26):
        sub = ['Maths','DBMS','OS','MP','CG']
        data = [i+1,chr((i%26+65)),random.choice(sub),random.randint(50,100)]
        for j in range(4):
            work.write(i+1,j,data[j])
