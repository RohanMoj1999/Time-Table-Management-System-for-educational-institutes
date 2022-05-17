from distutils.log import error
from re import I
import tkinter
from tkinter import *

# database connectivity
# import pymysql for MySQL
# import pyodbc for SQL server management studio
import pyodbc


''' --------------------- DESIGNING GUI of the window (text and text Fields) --------------------- '''

examWindow = tkinter.Tk()

examWindow.geometry("800x480")  # windows size
L = Label(examWindow, text="Enter class Schedule", font=(
    'Times New Roman', 20, 'bold'), fg='blue', highlightbackground='yellow')
L.grid(row=0, column=1)

LSubMon = Label(examWindow, text="Subject Name (MON):", font=('arial', 15), fg='black')
LSubMon.grid(row=1, column=0)
ESubMon = Entry(examWindow, bd=5, width=20)  # border = bd
ESubMon.grid(row=1, column=1)

LSubTues = Label(examWindow, text="Subject Name (TUES):",font=('arial', 15), fg='black')
LSubTues.grid(row=2, column=0)
ESubTues = Entry(examWindow, bd=5, width=20)  # border = bd
ESubTues.grid(row=2, column=1)

LTime1 = Label(examWindow, text="Time:", font=('arial', 15), fg='black')
LTime1.grid(row=3, column=0)
ETime1 = Entry(examWindow, bd=5, width=20)  # border = bd
ETime1.grid(row=3, column=1)
LTime2 = Label(examWindow, text="(format: 'hh:mm')",font=('arial', 10), fg='black')
LTime2.grid(row=3, column=2)


''' --------------------- DATABASE CONNECTIVITY --------------------- '''

# events


def myButtonEvent(selection):

    # fetching values for printing in python terminal
    # print("Subject:", ESub1.get())
    # print("Subject Code:", ESub2.get())
    # print("Date:", EDate1.get())
    # print("Start Time:", ETime1.get())
    # print("End Time:", ETime2.get())

    subMon = ESubMon.get()
    subTues = ESubTues.get()
    classTime = ETime1.get()

    if selection in ("Insert"):
        # database connectivity code must be written under button method
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER={YOUR Ms SQL SERVER NAME};'
            r'DATABASE={DATABASE NAME};'
            r'Trusted_Connection=yes;'
        )
        cnxn = pyodbc.connect(conn_str)

        # creating table if not exists
        cnxn.execute('''if OBJECT_ID(N'classSchedule',N'U') is null begin create table classSchedule(Mon char(10) NOT NULL, Tues char(10) NOT NULL, TIME time(7) NOT NULL) end''')  # sql commands
        cnxn.commit()

        insertQuery = "insert into classSchedule values('%s', '%s', '%s')" %(subMon, subTues, classTime)
        cnxn.execute(insertQuery)
        
        cnxn.commit()
        cnxn.close()

        print("Data Inserted Successfully !!")

    elif selection in ("Update"):

        # database connectivity code must be written under button method
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER={YOUR Ms SQL SERVER NAME};'
            r'DATABASE={DATABASE NAME};'
            r'Trusted_Connection=yes;'
        )
        cnxn = pyodbc.connect(conn_str)

        cnxn.execute("update classSchedule set Mon='%s'" %(subMon)+ " and Tues = '%s'" %(subTues)+"and Time = '%s'"%(classTime))
        cnxn.commit()
        cnxn.close()

        print("Data Updated Successfully !!")

    elif selection in ("Delete"):

        # database connectivity code must be written under button method
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER={YOUR Ms SQL SERVER NAME};'
            r'DATABASE={DATABASE NAME};'
            r'Trusted_Connection=yes;'
        )
        cnxn = pyodbc.connect(conn_str)

        cnxn.execute("delete from classSchedule where Mon='%s'" %(subMon)+ " and Tues = '%s'"%(subTues))
        cnxn.commit()
        cnxn.close()

        print("Data Deleted Successfully !!")


''' --------------------- DESIGNING GUI of the window (BUTTONS) --------------------- '''

insertButton = tkinter.Button(text='Insert', fg="black", bg='yellow', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Insert'))
insertButton.grid(row=6, column=1)

updateButton = tkinter.Button(text='Update', fg="white", bg='orange', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Update'))
updateButton.grid(row=7, column=1)

deleteButton = tkinter.Button(text='Delete', fg="red", bg='white', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Delete'))
deleteButton.grid(row=8, column=1)

mainloop()
