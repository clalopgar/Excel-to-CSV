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
				df = pd.read_csv(path, sep=GetSep(),encoding='cp1252',dtype=str)
				df.to_excel(xlsx,encoding='cp1252',index=False)
				tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso"+GetSep())
			except:
				tkinter.messagebox.showerror("Error","Error con el fichero: "+name)
		
		paths=None
			
def excel_to_csv():
	filter_opd=[("Excel files","*.xlsx")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !=[]:

		for path in paths:
			try:
				name=os.path.split(path)[len(os.path.split(path))-1]
				csv_name=name.replace(".xlsx",".csv")
				csv=path.replace(name,csv_name)
				df = pd.read_excel(path,encoding='cp1252',dtype=str)
				df.to_csv(csv,encoding='cp1252',index=False, sep=GetSep())
				tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")
			except:
				tkinter.messagebox.showerror("Error","Error con el fichero: "+name)	
		
		paths=None	

def csv_to_csv():
	filter_opd=[("Csv files","*.csv")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !=[]:
		for path in paths:
			try:
				name=os.path.split(path)[len(os.path.split(path))-1]
				df = pd.read_csv(path, sep=GetSep(),encoding='cp1252',dtype=str)
				df.to_csv(path,encoding='cp1252',index=False,sep=SepCSV())
				tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")	
			except:
				tkinter.messagebox.showerror("Error","Error con el fichero: "+name)
		
		paths=None

def mergeCSV():
	
	dirname = filedialog.askdirectory(parent=root,initialdir="/",title='Please select a directory')
	print (dirname)
	if len(dirname)>0:
	
		dirSave = os.path.join(dirname,"Merge")
		
		dirs = os.listdir(dirname)
		if len(dirs)>=2:
			
			try:
				os.mkdir(dirSave)
			except:
				tkinter.messagebox.showerror("Error","Ya existe carpeta Merge")
			
			dir_1=os.path.join(dirname,str(dirs[0]))
			dirs_1, subdirs_1, files_1 = next(os.walk(dir_1))
			dir_2=os.path.join(dirname,str(dirs[1]))
			dirs_2, subdirs_2, files_2 = next(os.walk(dir_2))
		
			if files_1 == files_2:
				
				
				for f in files_1:
					if str(f).endswith(".csv"):
						print(f)
						s=','
						df1 = pd.read_csv(os.path.join(dirs_1,str(f)),error_bad_lines=False,sep=s,encoding='cp1252',dtype=str)
						df2 = pd.read_csv(os.path.join(dirs_2,str(f)),error_bad_lines=False,sep=s,encoding='cp1252',dtype=str)
						out = df1.append(df2)
						s=';'
						#print(out)
						csv =os.path.join(dirSave,str(f))
						sep=";"
						out.to_csv(csv,encoding='cp1252',index=False,sep=sep)
			tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")	

  			



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
"""
buttonExceltoCSV= Button(root,text="Excel a CSV",command=excel_to_csv)
buttonCSVtoExcel= Button(root,text="CSV a Excel",command=csv_to_excel)
buttonCSVtoCSV= Button(root,text="CSV a CSV",command=csv_to_csv)
mergeCSV= Button(root,text="Merge",command=csv_to_csv)
buttonExceltoCSV.pack(padx=20, pady=5, side='left')
buttonCSVtoExcel.pack(padx=20, pady=5, side='left')
buttonCSVtoCSV.pack(padx=20, pady=15, side='left')
mergeCSV.pack()
"""

pane = Frame(root) 
pane.pack(fill = BOTH, expand = True) 
  
# button widgets which can also expand and fill 
# in the parent widget entirely 
b1 = Button(pane, text = "Excel a CSV",command=excel_to_csv) 
b1.pack(side = LEFT, expand = True, fill = BOTH) 
  
b2 = Button(pane, text = "CSV a Excel",command=csv_to_excel) 
b2.pack(side = LEFT, expand = True, fill = BOTH) 
  
b3 = Button(pane, text = "CSV a CSV",command=csv_to_csv) 
b3.pack(side = LEFT, expand = True, fill = BOTH) 
b4 = Button(pane, text = "Merge",command=mergeCSV ) 
b4.pack(side = LEFT, expand = True, fill = BOTH) 

root.mainloop()