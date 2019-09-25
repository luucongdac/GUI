#!usr/bin/python
import codecs
import openpyxl
import sys, os
from time import gmtime, strftime
from pathlib import Path
from openpyxl.utils import get_column_letter, column_index_from_string
import array as arr
import numpy as np
from shutil import copyfile
from tkinter import *
from tkinter import messagebox
import random
import time

import sys, string, os
from new_colect_testcase import run

#---------------------------------------------------------------------


#--------------------------------------------------------------------
window = Tk()
 
window.title("Analyzis TestCase")
 
window.geometry('600x800')

lb_status = Label(window, text="Status", font=("Arial Bold", 15),fg="red")
lb_status.grid(column=0, row=0)

lbl = Label(window, text="Type path excel file report: ")
lbl.grid(column=0, row=2)
txt_path = Entry(window,width=100)
txt_path.grid(column=0, row=3)

lb2 = Label(window, text="Type path Source .c test: ")
lb2.grid(column=0, row=5)
txt_func = Entry(window,width=70)
txt_func.grid(column=0, row=6)

path_excel = ''
source = ''
#


def task():
    box_1 = txt_path.get() 
    box_2 = txt_func.get()
    print(box_1)
    print(box_2)
    print('\n')
    file = open("D:/tp.txt", "w+") 
    file.write(box_1 + '$$$' + box_2)
    file.close() 
    #os.system("tempt.py 1")
    #import tempt.py
    #os.system("D:\auto.exe")
    #run(path_excel, source, run_in_excel, print_log)
    run(box_1, box_2, 0, 0)
    lb_status.configure(text="Finish")
    
def clicked():
    lb_status.configure(text="Running")
    #messagebox.showinfo('Message','Starting....')
    task()
 
btn = Button(window, text="Run...", command=clicked,bg="orange", fg="red")
btn.grid(column=0, row=20)


#while 1:
window.mainloop()
 #print('hello')
 #time.sleep(5)
