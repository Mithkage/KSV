import csv
import sqlite3
import tkinter as tk
from tkinter import filedialog

# def filePath(file_path):
#     input("-- Press Enter to enable selection of Cable Details --")
#     file_path = filedialog.askopenfilename()
#     return file_path

# file_path = filedialog.askopenfilename()
# f = open(file_path,"r")
# f = open(file_path,"r", encoding='ANSI', errors='ignore')

file_path='TrainDataUTF.CSV'
# f = open(file_path,"r")
# csv_f= csv.reader(f)

# with open(file_path, 'r', encoding='ANSI', errors='ignore') as infile:
#     with open('my_file_utf8.csv', 'w') as outfile:
#      outfile.write(infile.read())

conn = sqlite3.connect(":memory:")

# Open a file: file
file = open('TrainDataUTF.CSV',mode='r')
 
# read all lines at once
all_of_it = file.read()
 
# close the file
file.close()

# Open and Read the CSV file
with open(file_path, 'r', encoding='UTF-8') as infile:
    reader = csv.reader(infile)
    
    firstRow = True
    for row in reader:
        if (firstRow):
            #Remove blank Rows
            fields=""
            for element in row:
                newelement=element.replace('\ufeff', '')
                fields+="'"+newelement+"', "
            conn.execute("CREATE TABLE cableSchedule ("+fields[:-2]+")")
            firstRow = False
        elif (not(row[7]=="" and row[8]=="")):
            values=""
            for element in row:
                newelement=element.replace('\ufeff', '')
                values+="'"+newelement+"', "

            conn.execute("INSERT INTO cableSchedule VALUES ("+values[:-2]+")")

for row in conn.execute("SELECT `Cable Code`, `Cable Configuration` FROM cableSchedule WHERE NOT(`Cable Reference` = '') ORDER BY `Cable Reference` ASC"):
    number1=int(row[0])
    number2=int(row[1])
    newNumber = number1*number2
    print(str(number1) + " | "+ str(number2) + " | " + str(newNumber))

# xlsxwriter







