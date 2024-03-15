from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import mysql.connector


Treeview = 1 #replace this

mydb = mysql

root =Tk()

wrapper1 = LabelFrame(root, text='Guest List')
wrapper3 = LabelFrame(root, text='Guest Data')
wrapper2 = LabelFrame(root, text='Search')


wrapper1.pack(fill='both', expand='yes', padx=20, pady=10)
wrapper3.pack(fill='both', expand='yes', padx=20, pady=10)
wrapper2.pack(fill='both', expand='yes', padx=20, pady=10)

trv = Treeview(wrapper1, columns=(1,2,3,4), show='headings', height='6')
trv.pack()

trv.heading(1, text='ID')
trv.heading(2, text='Name')
trv.heading(3, text='Riel')
trv.heading(4, text='USD')





root.title('Bao\'s Simple Wed')
root.geometry('1800x900')
root.mainloop()