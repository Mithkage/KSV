import csv
import sqlite3
# import string
import tkinter as tk
from tkinter import filedialog
import pandas as pd


# def filePath(file_path):
#     input("-- Press Enter to enable selection of Cable Details --")
#     file_path = filedialog.askopenfilename()
#     return file_path

# file_path = filedialog..askopenfilename()
file_path = filedialog.askopenfilename()
# f = open(file_path,"r")
# f = open(file_path,"r", encoding='ANSI', errors='ignore')

# csv_f= csv.reader(f)

# with open(file_path, 'r', encoding='ANSI', errors='ignore') as infile:
#     with open('my_file_utf8.csv', 'w') as outfile:
#      outfile.write(infile.read())

# Open and Read the CSV file
with open(file_path, 'r', encoding='ANSI', errors='ignore') as infile:
    with open('my_file_utf8.csv', 'w', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)
        #Remove blank Rows
        for row in reader:
            print(row[0])
            #Remove blank Rows
            if (not(row[7]=="" and row[8]=="")):
                writer.writerow(row)
                # writer.writerow("\n")
            # firstRow = True

# for row in csv_f:
#     if (firstRow):
#         firstRow = False
#     elif (not(row[7]=="" and row[8]=="")):
#         print(row)

#Create Database
    conn = sqlite3.connect(":memory:")
    # # Create table
    conn.execute('CREATE TABLE studentResults (StudentID, Form, House, SubjectName, Teacher, Result, ResultType, TermNumber, SubjectCode)')
# print(f)







