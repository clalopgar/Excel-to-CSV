import pandas as pd
import tkinter as tk
from tkinter import filedialog

#Open file dialog python----------------

def OpenFileDialog():
	"""
	Abre una ventana emeregente para seleccionar un archivo
	devuelve ruta a ese archivo
	"""
	root = tk.Tk()
	root.withdraw()
	return filedialog.askopenfilename(filetypes={("Excel files","*.xlsx"),("Excel Worksheets","*.xls"),("All files","*")})
	
#----------------------------------------

excel_file=OpenFileDialog()

if excel_file != None and excel_file !="":
	print(excel_file)	
	data_xls = pd.read_excel(excel_file, index_col=None)
	data_xls.to_csv('C:\\ejemplos\\new_104\\csvfile.csv',sep=';', encoding='utf-8', index=False)
	#print(data_xls)
else:
	print("Nada seleccionado")




#data_xls = pd.read_excel('excelfile.xlsx', 'Sheet2', index_col=None)
#data_xls.to_csv('csvfile.csv', encoding='utf-8', index=False)