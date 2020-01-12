from tkinter import *
import pandas as pd
import tkinter.messagebox
from tkinter import simpledialog
from tkinter import filedialog
import os
import csv, operator
import sys

class popupWindow(object):
    def __init__(self,master):
        top=self.top=Toplevel(master)
        self.l=Label(top,text="Hello World")
        self.l.pack()
        self.e=Entry(top)
        self.e.pack()
        self.b=Button(top,text='Ok',command=self.cleanup)
        self.b.pack()
    def cleanup(self):
        self.value=self.e.get()
        self.top.destroy()
class mainWindow(object):
    def __init__(self,master):
        self.master=master
        self.b=Button(master,text="click me!",command=self.popup)
        self.b.pack()
        self.b2=Button(master,text="print value",command=lambda: sys.stdout.write(self.entryValue()+'\n'))
        self.b2.pack()

    def popup(self):
        self.w=popupWindow(self.master)
        self.b["state"] = "disabled" 
        self.master.wait_window(self.w.top)
        self.b["state"] = "normal"

    def entryValue(self):
        return self.w.value


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
		d=True
		
		
		for path in paths:
			try:



				name=os.path.split(path)[len(os.path.split(path))-1]
				xlsx_name=name.replace(".csv",".xlsx")
				xlsx=path.replace(name,xlsx_name)
				
				df = pd.read_csv(path, sep=GetSep(),encoding='cp1252',dtype=str,low_memory=False,header=None)
				#df.to_excel(xlsx,encoding='cp1252',index=False, options={'strings_to_numbers': True})
				writer = pd.ExcelWriter(xlsx, engine='xlsxwriter', options={'strings_to_numbers': True})
				df.to_excel(writer, sheet_name='Sheet1', index=False,header=None,encoding='cp1252')
				writer.save()
				writer.close()
			except Exception as error:
				tkinter.messagebox.showerror("Error",str(error))
				d=False
		if d:
			tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")	
		
		paths=None
			
def excel_to_csv():
	filter_opd=[("Excel files","*.xlsx")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !=[]:
		d=True
		for path in paths:
			try:
				name=os.path.split(path)[len(os.path.split(path))-1]
				csv_name=name.replace(".xlsx",".csv")
				csv=path.replace(name,csv_name)
				df = pd.read_excel(path,encoding='cp1252',dtype=str)
				df.to_csv(csv,encoding='cp1252',index=False, sep=GetSep())
				
			except Exception as error:
				tkinter.messagebox.showerror("Error",str(error))
				d=False
		if d:
			tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")	
		paths=None
		
		paths=None	

def csv_to_csv():
	filter_opd=[("Csv files","*.csv")]
	paths=OpenFileDialog(filter_opd)
	if paths != None and paths !=[]:
		d=True
		for path in paths:
			try:
				name=os.path.split(path)[len(os.path.split(path))-1]
				df = pd.read_csv(path, sep=GetSep(),encoding='cp1252',dtype=str)
				df.to_csv(path,encoding='cp1252',index=False,sep=SepCSV())
				
			except Exception as error:
				tkinter.messagebox.showerror(name,str(error))
				d=False
		if d:
			tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")	
		paths=None


def ReadLine(path):
	f= open(path,'r')
	message= f.readline()
	f.close()
	return message
def FileSeparator(line):
	csvSep=None
	if line.find(",") !=-1:
		return ","
	elif line.find(";") !=-1:
		return ";"
	return csvSep

def NumberIT(itfile):
	if itfile.find("INDICATION_TYPE_CODE") !=-1:
		return "03"
	elif itfile.find("MEASURAND_TYPE_CODE") !=-1:
		return "04"
	elif itfile.find("SETPOINT_TYPE_CODE") !=-1:
		return "05"
	elif itfile.find("ACCUMULATOR_TYPE_CODE") !=-1:
		return "06"
	elif itfile.find("PI_NAME") !=-1 and itfile.find("PI_POLICY") !=-1 and itfile.find("NOTE_TEXT_1") !=-1:
		return "07"
	elif itfile.find("GEO_SR_EXTERNAL_IDENTITY") !=-1 and itfile.find("SPISYS_ACRONYM") !=-1:
		return "02"
	elif itfile.find("ELECTRICAL_STATION") !=-1:
		return "11"
	elif itfile.find("SUBSYSTEM_ACRONYM") !=-1 and itfile.find("WIND_PARK") !=-1 and itfile.find("DENTIFICATION_TEXT") !=-1 and itfile.find("P_NOMINAL") ==-1:
		return "13"
	elif itfile.find("STATION") !=-1 and itfile.find("SUBSYSTEM") !=-1 and itfile.find("SUBSYSTEM_ACRONYM") ==-1:
		return "02B"
	else:
		return None

def SelectBy(path,sepCsv,filterby,conditions):
	df = pd.read_csv(path, sep=sepCsv,encoding='cp1252',dtype=str)
	return df[df[filterby].isin(conditions) ]


def Contains(path,sepCsv,filterby,conditions):
	df = pd.read_csv(path, sep=sepCsv,encoding='cp1252',dtype=str)
	return df[df[filterby].str.contains('|'.join(conditions)) ]

def SelectUniqueBy(df,field):
	return pd.unique(df[field]).tolist()

def AppendToList(listUniq,objects):
	listNew=[]
	if objects is not None:
		listNew=objects[:]
	print("lista leida",objects)
	if listUniq is not None:
		listNew =listUniq[:]
		listNew = listNew.extend(objects)
	return listNew

def SaveCsv(path):
	name=os.path.split(path)[len(os.path.split(path))-1]
	dirname=os.path.join(os.path.dirname(path),"csv_filtrados") 


def Filter(station):
	if station is not None:
		
		print("Hello", station)
		
		
		filter_opd=[("Csv files","*.csv")]
		paths=OpenFileDialog(filter_opd)
		version=[]
		acronyms=[]
		pgTypes=[]
		path_1 =None
		path_8 =None
		path_9 =None
		path_10 =None

		if paths != None and paths !=[]:
			d=True
			name_excel=os.path.split(paths[0])[len(os.path.split(paths[0]))-1]
			path_excel=paths[0].replace(name_excel,"ExtraccionFiltrada.xlsx")
			print("Ruta", path_excel)
			

			writer = pd.ExcelWriter(path_excel, engine='xlsxwriter', options={'strings_to_numbers': True})
				#df.to_excel(writer, sheet_name='Sheet1', index=False,header=None,encoding='cp1252')
			for path in paths:
				try:
					
					header=ReadLine(path)
					sepCsv = FileSeparator(header)
					itNumber = NumberIT(header)
					name=os.path.split(path)[len(os.path.split(path))-1]
					dirname=os.path.join(os.path.dirname(path),"csv_filtrados")
					gu =os.path.join(dirname,name)
					print("Ruta", gu)
					print("Nombre", name)
					print("Cabecera", header)
					print("separador", sepCsv)
					print("pestaña", itNumber)
					hasStationAcr = header.find("STATION_ACRONYM") !=-1
					hasStation= header.find("STATION") !=-1
					hasExternalId =header.find("EXTERNAL_IDENTITY") !=-1
					hasAcronym = header.find("SUBSYSTEM_ACRONYM") !=-1
					is1=header.find("SUBSYSTEM_ACRONYM") !=-1 and header.find("SUBSYSTEM_TEXT") !=-1 and header.find("IMP_ID")!=-1 and not header.find("EXTERNAL_IDENTITY")!=-1
					is8=header.find("INFO_PICTURE_NAME") !=-1 and header.find("STATE_0_TEXT") !=-1 and header.find("STATE_11_TEXT")!=-1 and header.find("EXTERNAL_IDENTITY")!=-1
					is9=header.find("POINT_GROUP_TYPE_NAME") !=-1 and header.find("ELEMENT_NUMBER") !=-1 and header.find("ALARM_POINT_GROUP_TYPE_NAME")!=-1 and header.find("SUFFIX")!=-1
					is10=header.find("TYPE") !=-1 and header.find("SUBCODE") !=-1 and header.find("RESETTABLE")!=-1 and header.find("FAULT_GONE_EVENT_PROC")!=-1
					is12=header.find("WIND_PARK") !=-1 and header.find("VERSION") !=-1 and header.find("P_NOMINAL")
					is14=header.find("CLASSIFICATION") !=-1 and header.find("POINT_GROUP_NAME") !=-1
					is15=hasStationAcr and header.find("NM_OBJECT_TYPE") !=-1

					if is12:
						itNumber="12"
					elif is14:
						itNumber="14"
					elif is15:
						itNumber="15"



					if is1:
						path_1=path
						itNumber="01"
					elif is8:
						path_8=path
						itNumber="08"
					elif is9:
						path_9=path
						itNumber="09"
					elif is10:
						path_10=path
						itNumber="10"
					elif hasStationAcr and not is15:
						df=SelectBy(path,sepCsv,"STATION_ACRONYM",station)
						#df.to_csv(path,encoding='cp1252',index=False,sep=sepCsv)
						if itNumber == None:
							itNumber= name.replace('.csv','')
						df.to_excel(writer, sheet_name=itNumber, index=False,encoding='cp1252')
						if hasAcronym:
							acronyms=AppendToList(acronyms,SelectUniqueBy(df,"SUBSYSTEM_ACRONYM"))
					elif is15:
						df=Contains(path,sepCsv,"POINT_GROUP_NAME",station)

						df.to_excel(writer, sheet_name="15", index=False,encoding='cp1252')
						#df.to_csv(path,encoding='cp1252',index=False,sep=sepCsv)
						if hasAcronym:
							acronyms=AppendToList(acronyms,SelectUniqueBy(df,"SUBSYSTEM_ACRONYM"))
					elif is14:
						df=Contains(path,sepCsv,"POINT_GROUP_NAME",station)
						df.to_excel(writer, sheet_name="14", index=False,encoding='cp1252')
						#df.to_csv(path,encoding='cp1252',index=False,sep=sepCsv)
						pgTypes = SelectUniqueBy(df,"POINT_GROUP_TYPE")
						if hasAcronym:
							acronyms=AppendToList(acronyms,SelectUniqueBy(df,"SUBSYSTEM_ACRONYM"))
					elif hasStation:
						df=SelectBy(path,sepCsv,"STATION",station)
						if itNumber == None:
							itNumber= name.replace('.csv','')
						df.to_excel(writer, sheet_name=itNumber, index=False,encoding='cp1252')
						#df.to_csv(path,encoding='cp1252',index=False,sep=sepCsv)
						if hasAcronym:
							acronyms=AppendToList(acronyms,SelectUniqueBy(df,"SUBSYSTEM_ACRONYM"))
					elif hasExternalId:
						df=Contains(path,sepCsv,"EXTERNAL_IDENTITY",station)
						if itNumber == None:
							itNumber= name.replace('.csv','')
						df.to_excel(writer, sheet_name=itNumber, index=False,encoding='cp1252')
						#df.to_csv(path,encoding='cp1252',index=False,sep=sepCsv)
						if hasAcronym:
							acronyms=AppendToList(acronyms,SelectUniqueBy(df,"SUBSYSTEM_ACRONYM"))
						if is12:
							version = SelectUniqueBy(df,"VERSION")


					
				except Exception as error:
					tkinter.messagebox.showerror(name,str(error))
					d=False
			print(version)
			try:
				if path_1 is not None and acronyms !=[]:
					header=ReadLine(path_1)
					sepCsv = FileSeparator(header)
					df=SelectBy(path_1,sepCsv,"SUBSYSTEM_ACRONYM",acronyms)
					df.to_excel(writer, sheet_name="01", index=False,encoding='cp1252')
					#df.to_csv(path_1,encoding='cp1252',index=False,sep=sepCsv)
				if path_8 is not None and pgTypes != []:
					header=ReadLine(path_8)
					sepCsv = FileSeparator(header)
					df=SelectBy(path_8,sepCsv,"EXTERNAL_IDENTITY",pgTypes)
					df.to_excel(writer, sheet_name="08", index=False,encoding='cp1252')
					#df.to_csv(path_8,encoding='cp1252',index=False,sep=sepCsv)
				if path_9 is not None and pgTypes != []:
					print("PGSs", pgTypes)
					header=ReadLine(path_9)
					sepCsv = FileSeparator(header)
					df=SelectBy(path_9,sepCsv,"POINT_GROUP_TYPE_NAME",pgTypes)
					df.to_excel(writer, sheet_name="09", index=False,encoding='cp1252')
					#df.to_csv(path_9,encoding='cp1252',index=False,sep=sepCsv)
				if path_10 is not None and pgTypes != []:
					header=ReadLine(path_10)
					sepCsv = FileSeparator(header)
					df=SelectBy(path_10,sepCsv,"POINT_GROUP_TYPE_NAME",pgTypes)
					if version !=[]:
						
						print("version___>", version)
						print(pd.isnull(df["VERSION"]))
						df=df[df["VERSION"].isin(version) | pd.isnull(df["VERSION"])]
					df.to_excel(writer, sheet_name="10", index=False,encoding='cp1252')
					#df.to_csv(path_10,encoding='cp1252',index=False,sep=sepCsv)

			except Exception as error:
				tkinter.messagebox.showerror(name,str(error))
				d=False
			writer.save()
			writer.close()
			print("acronyms", acronyms)
			if d:
				tkinter.messagebox.showinfo("Información","Ha finalizado correctamento el proceso")	
			paths=None




def FilterCSV():

	station =StationPopUp()
	Filter(station)


def StationPopUp():
	STFrame = Tk()
	STFrame.withdraw()
	# the input dialog
	STA_INP = simpledialog.askstring(title="Filtrar IT por Estación",
                                  prompt="Nombre de los Parques\n*Usar (;) para introducir varias:" )
	STFrame.destroy()
	if STA_INP == None:
		return None
	else:
		return STA_INP.upper().split(";")
	



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
						df1 = pd.read_csv(os.path.join(dirs_1,str(f)),error_bad_lines=False,sep=SepCSV(),encoding='cp1252',dtype=str)
						df2 = pd.read_csv(os.path.join(dirs_2,str(f)),error_bad_lines=False,sep=SepCSV(),encoding='cp1252',dtype=str)
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
b4 = Button(pane, text = "Filtrar",command=FilterCSV ) 
b4.pack(side = LEFT, expand = True, fill = BOTH) 

root.mainloop()