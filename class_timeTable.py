
'''
        PROJECT NAME: TIME TABLE MANAGEMENT SYSTEM FOR EDUCATION INSTITUTES
        PROJECT MODULE: CONNECTING SQL WITH PYTHON AND CREATING GUI
'''

# tkinter is a GUI library used to create GUI
# Tk of tkinter package is used to create and manipulating GUI window


import tkinter
from tkinter import *
from tkinter import messagebox

#PIL module is for inserting images in your GUI

import PIL
from PIL import Image,ImageTk

# import pyodbc for SQL server management studio (connects MsSQL with python)
import pyodbc

# for mail purpose
from fileinput import filename
import os
import smtplib
from email.message import EmailMessage

#   extracting data from excel
import pandas as pd

''' --------------------- DESIGNING GUI of the window (text and text Fields) --------------------- '''


examWindow=tkinter.Tk()

# windows size
examWindow.geometry('1000x800')
#examWindow.title is used to name your GUI
examWindow.title("sagacity institution")

#commands to insert image

img = ImageTk.PhotoImage(file='C:/Users/Rohan/Desktop/wipro internship/apple-book.jpg')

lab=Label(examWindow,image=img).place(x=0,y=1)



'''setting up text charactristics
    Label () is used normal text 
'''

L = Label(examWindow, text="Enter class Schedule", font=('Calisto MT', 20, 'bold'), fg='black',bg='white')
L.place(x=400,y=20)
#   location of that field


LSubMon = Label(examWindow, text="Subject Name (MON):", font=('Comic Sans MS', 15), fg='white',bg='green')
LSubMon.place(x=100,y=100)

''' setting up textfield characteristics
    Entry() is used to set up a textfield
'''
ESubMon = Entry(examWindow, bd=5, width=40)  # border = bd
ESubMon.place(x=400,y=100)

LSubTues = Label(examWindow, text="Subject Name (TUES):",font=('Comic Sans MS', 15), fg='white',bg='green')
LSubTues.place(x=100,y=150)
ESubTues = Entry(examWindow, bd=5, width=40)  # border = bd
ESubTues.place(x=400,y=150)

LTime1 = Label(examWindow, text="Time(hh:mm)", font=('Comic Sans MS', 15), fg='white',bg='green')
LTime1.place(x=100,y=200)
ETime1 = Entry(examWindow, bd=5, width=40)  # border = bd
ETime1.place(x=400,y=200)

''' --------------------- DATABASE CONNECTIVITY --------------------- '''

# events
def myButtonEvent(selection):

    # get() is used to fetch the entered data
    subMon = ESubMon.get()
    subTues = ESubTues.get()
    classTime = ETime1.get()

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
        cnxn.execute('''if OBJECT_ID(N'classSchedule',N'U') is null begin create table classSchedule(Mon char(10) NOT NULL, Tues char(10) NOT NULL, TIME time(7) NOT NULL) end''')
        cnxn.commit()

        insertQuery = "insert into classSchedule values('%s', '%s', '%s')" %(subMon, subTues, classTime)
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

        cnxn.execute("update classSchedule set Mon='%s'" %(subMon)+ ", Tues = '%s'" %(subTues)+" where TIME = '%s'"%(classTime))
        cnxn.commit()
        cnxn.close()
        messagebox.showinfo(title="CLASS SCHEDULE", message="Data Updated Successfully !")

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

        cnxn.execute("delete from classSchedule where Mon='%s'" %(subMon)+ " and Tues = '%s'"%(subTues))
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
        msg['Subject'] = 'Class Time Table'
        # sender
        msg['From'] = 'EMAIL_ADDRESS'
        # receiver
        msg['To'] = ",".join(emailsList)

        # body of the mail
        msg.set_content('Class schedule attached...')
        # path of an image 
        with open('C:/Users/Rohan/Desktop/wipro internship/class schedule.xls', 'rb') as f:  # 'rb' --> read byte
            file_data = f.read()
            file_name = f.name

        msg.add_attachment(file_data, maintype='application', subtype='xls', filename=file_name)
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

insertButton = tkinter.Button(text='INSERT', fg="white", bg='#E13102', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Insert'))
insertButton.place(x=200,y=500)

updateButton = tkinter.Button(text='UPDATE', fg="white", bg='#0AAC00', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Update'))
updateButton.place(x=400,y=500)

deleteButton = tkinter.Button(text='DELETE', fg="white", bg='#002050', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('Delete'))
deleteButton.place(x=600,y=500)

sendMail = tkinter.Button(text='SHARE', fg="white", bg='red', font=('arial', 10, 'bold'), command=lambda: myButtonEvent('senDmail'))
sendMail.place(x=400, y=600)



mainloop()
