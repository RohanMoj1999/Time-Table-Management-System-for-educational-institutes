# Time-Table-Management-System-for-educational-institutes
#Front-End: Python, Back-End: Ms SQL
In this project, I've designed a scheduling system with the help of Python and Ms SQL.
With python, I've created a GUI from where user can enter data and the data will be stored in Ms SQL.
Features include:
1. Inserting data through python based GUI data-fields
2. storing the data in Ms SQL database table
3. modify or delete data from that table
4. share the database table to the recipients if required via python email system.

### Designing GUI
To design the GUI window, we need certain libraries:
1. ```tkinter```: to design the GUI window
2. ```pillow```: to set background image for the GUI window

### Database Connectivity
To connect Ms SQL, we need:
1. ```pyodbc```: to establish connection with Ms SQL
2. provide the exact _database_ name to create table under that specific database
> import ```pymysql``` to establish connection with MySQL

### Establish Email Connection
1. import the library ```smtplib```
> Here I have imported email IDs from an excel file. For reading excel file, we have imported ```pandas``` library.

### Output
Results of some features are shown below:
1. Inserting data
![Data Uploaded successfully](https://user-images.githubusercontent.com/62896383/169952775-d1ab7014-54ec-476a-a3c4-f6f0e433bf6b.png)
2. Database output
![Table result](https://user-images.githubusercontent.com/62896383/169952911-ea124580-2daa-47d3-93d9-d856fcd9ff67.png)
3. Shared the excel format of the MsSQL table via email
![Mail received](https://user-images.githubusercontent.com/62896383/169953338-0e8c996d-761e-4d82-b86d-156f9547f31e.png)
4. Video output of some features of this project

https://user-images.githubusercontent.com/62896383/169953615-1e3102af-4f0a-4adc-8e79-ae2e962275be.mp4


