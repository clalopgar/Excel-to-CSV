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
	if opcion.get()==1:
		sep=','

	return sep

def SepCSV():
	sep=GetSep()
	if sep ==';':
		sep =','
	else:
		sep=';'
	return sep



def csv_to_excel():
	filter_opd=[("Csv files","*.csv")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !=[]:
		for path in paths:
			try:
				name=os.path.split(path)[len(os.path.split(path))-1]
				xlsx_name=name.replace(".csv",".xlsx")
				xlsx=path.replace(name,xlsx_name)
				df = pd.read_csv(path, sep=GetSep(),encoding='cp1252')
				df.to_excel(xlsx,encoding='cp1252',index=False)
			except:
				tkinter.messagebox.showerror("Error","Error con el fichero: "+name)
		tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso"+GetSep())
		paths=None
			
def excel_to_csv():
	filter_opd=[("Csv files","*.xlsx")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !=[]:

		for path in paths:
			try:
				name=os.path.split(path)[len(os.path.split(path))-1]
				csv_name=name.replace(".xlsx",".csv")
				csv=path.replace(name,csv_name)
				df = pd.read_excel(path,encoding='cp1252')
				df.to_csv(csv,encoding='cp1252',index=False, sep=GetSep())
			except:
				tkinter.messagebox.showerror("Error","Error con el fichero: "+name)	
		tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")
		paths=None	

def csv_to_csv():
	filter_opd=[("Csv files","*.csv")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !=[]:
		for path in paths:
			try:
				name=os.path.split(path)[len(os.path.split(path))-1]
				df = pd.read_csv(path, sep=GetSep(),encoding='cp1252')
				df.to_csv(path,encoding='cp1252',index=False,sep=SepCSV())
			except:
				tkinter.messagebox.showerror("Error","Error con el fichero: "+name)
		tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")	
		paths=None



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
lb=Label(frame,text="CSV separador:")
lb.grid(row=0,column=1,sticky=W,padx=30)



opcion = IntVar()
rd1=Radiobutton(frame, text="delimitado por coma",variable=opcion, value=1)
rd1.grid(row=1,column=1,sticky=W,padx=35)

rd2=Radiobutton(frame, text="delimitado por punto y coma", variable=opcion, value=2)
rd2.grid(row=2,column=1,sticky=W,padx=35)
rd2.select()


"""Radio button"""

buttonExceltoCSV= Button(root,text="Excel a CSV",command=excel_to_csv)
buttonCSVtoExcel= Button(root,text="CSV a Excel",command=csv_to_excel)
buttonCSVtoCSV= Button(root,text="CSV a CSV",command=csv_to_csv)

buttonExceltoCSV.pack(padx=20, pady=5, side='left')
buttonCSVtoExcel.pack(padx=20, pady=5, side='left')
buttonCSVtoCSV.pack(padx=20, pady=15, side='left')


root.mainloop()
