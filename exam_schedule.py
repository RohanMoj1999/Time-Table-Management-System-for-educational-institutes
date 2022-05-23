'''
        PROJECT NAME: TIME TABLE MANAGEMENT SYSTEM FOR EDUCATION INSTITUTES
        PROJECT MODULE: CONNECTING SQL WITH PYTHON AND CREATING GUI
'''
import tkinter
from tkinter import *
from tkinter import messagebox

#PIL module is for inserting images in your GUI
import PIL
from PIL import Image, ImageTk

# database connectivity
# import pymysql for MySQL
# import pyodbc for SQL server management studio
import pyodbc

# for mail purpose
from fileinput import filename
import os
import smtplib
from email.message import EmailMessage

#   extracting data from excel
import pandas as pd


''' --------------------- DESIGNING GUI of the window (text and text Fields) --------------------- '''

examWindow = tkinter.Tk()

# windows size
examWindow.geometry('1000x800')
#examWindow.title is used to name your GUI
examWindow.title("sagacity institution")

img = ImageTk.PhotoImage(
    file='C:/Users/Rohan/Desktop/wipro internship/apple-book.jpg')

lab = Label(examWindow, image=img).place(x=0, y=1)


'''setting up text charactristics
    Label () is used normal text 
'''
L = Label(examWindow, text="Enter Exam Schedule", font = ('Times New Roman',20,'bold'), fg = 'blue',highlightbackground='yellow')
#   location of that field
L.place(x=400,y=20)

LSub1 = Label(examWindow, text="Subject Name:", font = ('arial',15), fg = 'black')
LSub1.place(x=100,y=100)

''' setting up textfield characteristics
    Entry() is used to set up a textfield
'''
ESub1 = Entry(examWindow, bd = 5, width=30)     #   border = bd
ESub1.place(x=400,y=100)

LSub1 = Label(examWindow, text="Subject Code:", font = ('arial',15), fg = 'black')
LSub1.place(x=100,y=150)
ESub2 = Entry(examWindow, bd = 5, width=20)     #   border = bd
ESub2.place(x=400,y=150)

LDate1 = Label(examWindow, text="Date (format: 'yyyy-mm-dd'):", font = ('arial',15), fg = 'black')
LDate1.place(x=100,y=200)
EDate1 = Entry(examWindow, bd = 5, width=20)     #   border = bd
EDate1.place(x=400,y=200)

LTime1 = Label(examWindow, text="Start Time ('hh:mm'):", font = ('arial',15), fg = 'black')
LTime1.place(x=100,y=250)
ETime1 = Entry(examWindow, bd = 5, width=20)     #   border = bd
ETime1.place(x=400,y=250)

LTime2 = Label(examWindow, text="End Time ('hh:mm'):", font = ('arial',15), fg = 'black')
LTime2.place(x=100,y=300)
ETime2 = Entry(examWindow, bd = 5, width=20)     #   border = bd
ETime2.place(x=400,y=300)

''' --------------------- DATABASE CONNECTIVITY --------------------- '''

# events

def myButtonEvent(selection):
    
    SubName = ESub1.get()
    SubCode = ESub2.get()
    examDate = EDate1.get()
    startTime = ETime1.get()
    endTime = ETime2.get()


    '''Inserting data'''
    if selection in ("Insert"):
        # database connectivity code must be written under button method
        conn_str = (
                r'DRIVER={SQL Server};'
                r'SERVER=ROHAN\SQLEXPRESS;'
                r'DATABASE=pythonConnection;'
                r'Trusted_Connection=yes;'
            )
        cnxn = pyodbc.connect(conn_str)

        # executing sql codes with cnxn.execute
        cnxn.execute('''if OBJECT_ID(N'examSchedule',N'U') is null begin create table examSchedule(SubName char(20) Not Null, SubCode varchar(10), examDate DATE NOT NULL, startTime time(7) NOT NULL, endTime time(7) NOT NULL) end''')  # sql commands
        cnxn.commit()

        insertQuery = "insert into examSchedule values('%s', '%s', '%s','%s','%s')" % (SubName, SubCode, examDate, startTime, endTime)
        cnxn.execute(insertQuery)
        cnxn.commit()
        
        # closing the connection of MsSQL
        cnxn.close()

        messagebox.showinfo(title="EXAM SCHEDULE", message="Data Inserted successfully !")
    
    # UPDATING DATA
    
    elif selection in ("Update"):
        
        # database connectivity code must be written under button method
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER=ROHAN\SQLEXPRESS;'
            r'DATABASE=pythonConnection;'
            r'Trusted_Connection=yes;'
        )
        cnxn = pyodbc.connect(conn_str)
        
        cnxn.execute("update examSchedule set examDate='%s'" % (examDate)+", startTime='%s'" %(startTime)+", endTime='%s'" % (endTime)+" where SubCode = '%s'" % (SubCode))
        cnxn.commit()
        cnxn.close()
        
        messagebox.showinfo(title="EXAM SCHEDULE", message="Data of %s Updated Successfully !" %(SubCode))
        
    # DELETING DATA   
        
    elif selection in ("Delete"):

        # database connectivity code must be written under button method
        conn_str = (
            r'DRIVER={SQL Server};'
            r'SERVER=ROHAN\SQLEXPRESS;'
            r'DATABASE=pythonConnection;'
            r'Trusted_Connection=yes;'
        )
        cnxn = pyodbc.connect(conn_str)

        cnxn.execute("delete from examSchedule where SubCode='%s'"% (SubCode))
        cnxn.commit()
        cnxn.close()
        
        messagebox.showinfo(title="EXAM SCHEDULE", message="Data Deleted successfully !")
        
# SEND MAIL

    elif selection in ("senDmail"):

        # grab the email_address and password from environment vairables for safety
        EMAIL_ADDRESS = os.environ.get('email_user')
        EMAIL_PASSWORD = os.environ.get('EMAIL_PASS')

        #   extracting EMAILS from EXCEL sheet
        filedata = pd.read_excel('emails.xlsx','Sheet1')  # table name, sheet name 
        emailsList = filedata['email'].values.tolist()
        msg = EmailMessage()
        # subject of mail
        msg['Subject'] = 'Python Test Mail'
        # sender
        msg['From'] = 'EMAIL_ADDRESS'
        # receiver
        msg['To'] = ",".join(emailsList)

        # body of the mail
        msg.set_content('Exam Scheduled attached...')
        # path of an image 
        with open('C:/Users/Rohan/Desktop/wipro internship/emails.xlsx', 'rb') as f:  # 'rb' --> read byte
            file_data = f.read()
            file_name = f.name

        msg.add_attachment(file_data, maintype='application', subtype='xlsx', filename=file_name)
        # subtype = 'pdf'       -->  for pdf

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)  # emaiID, password

            # send mail
            smtp.send_message(msg)

            messagebox.showinfo(title="EXAM SCHEDULE", message="Mail Sent Successfully !")


''' --------------------- DESIGNING GUI of the window (BUTTONS) --------------------- '''

''' 
tkinter.Button() is used to create and manipulate a button
connecting myButtonEvent() using command '''

insertButton = tkinter.Button(text='Insert', fg="black", bg='yellow',font=('arial', 10,'bold'), command=lambda:myButtonEvent('Insert'))
insertButton.place(x=200,y=500)

updateButton = tkinter.Button(text='Update', fg="white", bg='orange',font=('arial', 10,'bold'), command=lambda:myButtonEvent('Update'))
updateButton.place(x=400,y=500)

deleteButton = tkinter.Button(text='Delete', fg="red", bg='white', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Delete'))
deleteButton.place(x=600,y=500)

sendMail = tkinter.Button(text='SHARE', fg="white", bg='red', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('senDmail'))
sendMail.place(x=400, y=600)


mainloop()
