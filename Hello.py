import csv
import sqlite3
import tkinter as tk
from tkinter import filedialog
import pandas as pd

input("-- Press Enter to enable selection of Student Details --")
file_path= filedialog.askopenfilename()
f = open(file_path,"r")
f = open(file_path,"r", encoding='ANSI', errors='ignore')

csv_f= csv.reader(f)

#Create Database
    # conn = sqlite3.connect(":memory:")
    # # Create table
    # conn.execute('CREATE TABLE studentResults (StudentID, Form, House, SubjectName, Teacher, Result, ResultType, TermNumber, SubjectCode)')
# print(f)



firstRow = True

for row in csv_f:
    if (firstRow):
        firstRow = False
    elif (not(row[7]=="" and row[8]=="")):
        print(row)


