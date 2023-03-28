import sqlite3
from openpyxl import Workbook
from datetime import datetime

T = 1
while T > 0:
    print("\nWhat do you want to do?")
    print("1 - add expense\n2 - export expenses to Excel\n3 - exit\n ")
    x = int(input())
    if x == 1:
        con = sqlite3.connect('Expenses.sqlite')
        cur = con.cursor()
        for row in cur.execute("SELECT * FROM Category ORDER BY CategoryID"):
            print(row)
        print("To which category belongs this expense? Choose number between 1-11")

        Cat = input()
        print("What is te amount of the expense?")
        Ex = input()
        now = datetime.now()
        date_time = now.strftime("%Y-%m-%d")

        # Insert a row of data
        cur.execute("INSERT INTO Expense(Amount, CategoryId, Date) VALUES (?,?,?)", (Ex, Cat, date_time))
        # Save (commit) the changes
        con.commit()
        con.close()
        print("Done")
        continue
    if x == 2:

        wb = Workbook()
        worksheet = wb.active
        worksheet.title = "Expenses"
        con = sqlite3.connect('Expenses.sqlite')
        cur = con.cursor()
        now = datetime.now()
        date_time = now.strftime("%Y-%m-%d")
        dest_filename = ('%s.xlsx' % date_time)

        mysel = cur.execute(
            "SELECT Amount, Category.Name, Date FROM Expense INNER JOIN Category ON Expense.CategoryId = Category.CategoryID ")
        for i, row in enumerate(mysel):
            for j, values in enumerate(row):
                worksheet.cell(row=i + 1, column=j + 1).value = values

        wb.save(dest_filename)
        wb.close()
        print("Done")
        continue
    else:
        T = -1
        break
print("End")



