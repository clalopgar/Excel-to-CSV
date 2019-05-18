from tkinter import *
import pandas as pd
import tkinter.messagebox
from tkinter import filedialog
import os
import csv, operator

def OpenFileDialog(filter_opd):
	"""
	Abre una ventana emeregente para seleccionar un archivo
	devuelve ruta a ese archivo
	"""
	root = Tk()
	root.withdraw()
	filez = filedialog.askopenfilenames(parent=root,title='Choose a file',filetypes=filter_opd)

	root.destroy()
	return list(filez)

def GetSep():
	sep=';'
	if opcion.get() == 1:
		sep=','
	return sep


def csv_to_excel():
	filter_opd=[("Csv files","*.csv")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !="":
		for path in paths:
			name=os.path.split(path)[len(os.path.split(path))-1]
			xlsx_name=name.replace(".csv",".xlsx")
			xlsx=path.replace(name,xlsx_name)
			df = pd.read_csv(path, sep=GetSep(),encoding='cp1252')
			df.to_excel(xlsx,encoding='cp1252',index=False)
			
			





"""GUI"""

root = Tk()
root.title("Convert Files")
root.resizable(0,0)
root.geometry("350x150")
"""Frame"""
frame =Frame(root)
frame.pack(fill="both",expand="True")
frame.config(width="250",height="150")
"""Radio button"""
lb=Label(frame,text="CSV")
lb.grid(row=0,column=0)
opcion = IntVar()
rd1=Radiobutton(frame, text="delimitado por coma",variable=opcion, value=1)
rd1.grid(row=2,column=2)

rd2=Radiobutton(frame, text="delimitado por punto y coma", variable=opcion, value=2)
rd2.grid(row=2,column=1)
rd2.select()

"""Radio button"""
buttonExceltoCSV= Button(frame,text="Excel a CSV")
buttonCSVtoExcel= Button(frame,text="CSV a Excel",command=csv_to_excel)
buttonExceltoCSV.grid(row=6,column=1)
buttonCSVtoExcel.grid(row=6,column=2)

csv_file="C:\\ejemplos\\direccionamiento\\datos104_prueba2.csv"



root.mainloop()
