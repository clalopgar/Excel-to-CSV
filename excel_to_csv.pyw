import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox

#Open file dialog python----------------

def OpenFileDialog():
	"""
	Abre una ventana emeregente para seleccionar un archivo
	devuelve ruta a ese archivo
	"""
	root = tk.Tk()
	root.withdraw()
	#f = filedialog.asksaveasfile(mode='w', defaultextension=".txt")
	path= filedialog.askopenfilename(defaultextension=".xlsx",filetypes=[("Excel files","*.xlsx"),("Excel Worksheets","*.xls"),("All files","*")])
	root.destroy()
	return path
	
#----------------------------------------

def display_and_print():  
    tkinter.messagebox.showinfo("Info","Just so you know")
    tkinter.messagebox.showwarning("Warning","Better be careful")
    tkinter.messagebox.showerror("Error","Something went wrong")
 
    okcancel = tkinter.messagebox.askokcancel("What do you think?","Should we go ahead?")
    print(okcancel)
 
    yesno = tkinter.messagebox.askyesno("What do you think?","Please decide")
    print(yesno)
 
    retrycancel = tkinter.messagebox.askretrycancel("What do you think?","Should we try again?")
    print(retrycancel)
 
    answer = tkinter.messagebox.askquestion("What do you think?","What's your answer?")
    print(answer)


def main():
	excel_file=OpenFileDialog()

	if excel_file != None and excel_file !="":
		print(excel_file.endswith('.xlsx'))
		#display_and_print()
		print(type(excel_file))

		path =os.path.dirname(excel_file)
		file_name= os.path.split(excel_file)[len(os.path.split(excel_file))-1]


		print("path: ",path)
		
		print("Archivo: ",file_name)
		print("Ruta completa: ",len(os.path.split(excel_file)))
		#print(os.path.dirname(excel_file)) #Obtine el path de la ruta a un fichero
		#data_xls = pd.read_excel(excel_file, index_col=None)
		
		#data_xls.to_csv('C:\\ejemplos\\new_104\\csvfile.csv',sep=';', encoding='utf-8', index=False)
		#print(data_xls)

	else:
		print("Nada seleccionado")


#Interfáz gráfica
root = tk.Tk()
root.resizable(0,0)
root.geometry("250x150")

root.mainloop()

#data_xls = pd.read_excel('excelfile.xlsx', 'Sheet2', index_col=None)
#data_xls.to_csv('csvfile.csv', encoding='utf-8', index=False)