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

examWindow.geometry("800x480")  #   windows size
L = Label(examWindow, text="Enter Exam Schedule", font = ('Times New Roman',20,'bold'), fg = 'blue',highlightbackground='yellow')
L.grid(row = 0, column = 1)

LSub1 = Label(examWindow, text="Subject Name:", font = ('arial',15), fg = 'black')
LSub1.grid(row = 1, column = 0)
ESub1 = Entry(examWindow, bd = 5, width=30)     #   border = bd
ESub1.grid(row = 1, column = 1)

LSub1 = Label(examWindow, text="Subject Code:", font = ('arial',15), fg = 'black')
LSub1.grid(row = 2, column = 0)
ESub2 = Entry(examWindow, bd = 5, width=20)     #   border = bd
ESub2.grid(row = 2, column = 1)

LDate1 = Label(examWindow, text="Date:", font = ('arial',15), fg = 'black')
LDate1.grid(row = 3, column = 0)
EDate1 = Entry(examWindow, bd = 5, width=20)     #   border = bd
EDate1.grid(row = 3, column = 1)
LDate1 = Label(examWindow, text="(format: 'yyyy-mm-dd')", font=('arial', 10), fg='black')
LDate1.grid(row=3, column=2)

LTime1 = Label(examWindow, text="Start Time:", font = ('arial',15), fg = 'black')
LTime1.grid(row = 4, column = 0)
ETime1 = Entry(examWindow, bd = 5, width=20)     #   border = bd
ETime1.grid(row=4, column=1)
LDate1 = Label(examWindow, text="(format: 'hh:mm')",font=('arial', 10), fg='black')
LDate1.grid(row=4, column=2)

LTime2 = Label(examWindow, text="End Time:", font = ('arial',15), fg = 'black')
LTime2.grid(row = 5, column = 0)
ETime2 = Entry(examWindow, bd = 5, width=20)     #   border = bd
ETime2.grid(row = 5, column = 1)
LDate1 = Label(examWindow, text="(format: 'hh:mm')", font=('arial', 10), fg='black')
LDate1.grid(row=5, column=2)


''' --------------------- DATABASE CONNECTIVITY --------------------- '''

# events

def myButtonEvent(selection):
    
    # fetching values for printing in python terminal
    # print("Subject:", ESub1.get())
    # print("Subject Code:", ESub2.get())
    # print("Date:", EDate1.get())
    # print("Start Time:", ETime1.get())
    # print("End Time:", ETime2.get())
    
    SubName = ESub1.get()
    SubCode = ESub2.get()
    examDate = EDate1.get()
    startTime = ETime1.get()
    endTime = ETime2.get()


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
        cnxn.execute('''if OBJECT_ID(N'examSchedule',N'U') is null begin create table examSchedule(SubName char(20) Not Null, SubCode varchar(10), examDate DATE NOT NULL, startTime time(7) NOT NULL, endTime time(7) NOT NULL) end''')  # sql commands
        cnxn.commit()

        insertQuery = "insert into examSchedule values('%s', '%s', '%s','%s','%s')" % (
            SubName, SubCode, examDate, startTime, endTime)
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
        
        cnxn.execute("update examSchedule set examDate='%s'" % (examDate)+", startTime='%s'" %(startTime)+", endTime='%s'" % (endTime)+" where SubCode = '%s'" % (SubCode))
        cnxn.commit()
        cnxn.close()
        
        print("Data of %s Updated Successfully !!"%(SubCode))
        
        
    elif selection in ("Delete"):

        # database connectivity code must be written under button method
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER={YOUR Ms SQL SERVER NAME};'
            r'DATABASE={DATABASE NAME};'
            r'Trusted_Connection=yes;'
        )
        cnxn = pyodbc.connect(conn_str)

        cnxn.execute("delete from examSchedule where SubCode='%s'"% (SubCode))
        cnxn.commit()
        cnxn.close()

        print("Data of row %s Deleted Successfully !!" % (SubCode))


''' --------------------- DESIGNING GUI of the window (BUTTONS) --------------------- '''

insertButton = tkinter.Button(text='Insert', fg="black", bg='yellow',font=('arial', 10,'bold'), command=lambda:myButtonEvent('Insert'))
insertButton.grid(row=6, column=1)

updateButton = tkinter.Button(text='Update', fg="white", bg='orange',font=('arial', 10,'bold'), command=lambda:myButtonEvent('Update'))
updateButton.grid(row=7, column=1)

deleteButton = tkinter.Button(text='Delete', fg="red", bg='white', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Delete'))
deleteButton.grid(row=8, column=1)

mainloop()
