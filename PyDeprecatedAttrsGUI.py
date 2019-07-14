# cd C:\Users\Ujjwal\AppData\Local\Programs\Python\Python37-32\
# python "C:\Users\1297418\AppData\Local\Programs\Python\Python37-32\Codepy\PyDeprecatedAttr\Code\PyDeprecatedAttrs.py"
# Purpose : Reading the Multi Table File and create Excel file with Meta data used as input for CSVtoTCXML Utility
# Creator : Ujjwal Kumar
# History : Version 1.1 (Date : 11th April, 2019)
# 			Version 1.2 (Date : 16th April, 2019) :: Fixed the bug for unprintable character
#			Version 1.3 (Date : 28th April, 2019) :: Command Line to GUI
# Warning : Do not hange identation as this python code which is based on identation
#		  : First three column will be used for File name. ITEM_ID_REV_TYPE
# Improvement Scope : UI Based or removal of few hard-coded path
# E:\\WorkInProgress\\OnePLM\\Python\\Code\\DeprecatedAttr
# 

# import openpyxl and tkinter modules 
#from openpyxl import *
from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import tkinter.messagebox as tkMessageBox
import os
import pandas as pd
import datetime
import string
import csv

# Function for clearing the 
# contents of text entry boxes 
def clear(): 
	
	# clear the content of text entry box 
	#Base_Folder_field.delete(0, END)
	print('###### Task Completed || Check Folder for data ########')
	textPad.insert(INSERT, '--------------------------------------------------------\n')
	textPad.insert(INSERT, '###### Task Completed || Check Folder for data ########')
	textPad.insert(INSERT, '\n--------------------------------------------------------\n')
 


# Function to take data from GUI 
# window and write to an excel file 
def browse_button():
	global folder_path
	global folderName
	#global filename = filedialog.askdirectory()
	folderName = filedialog.askdirectory()
	folder_path.set(folderName)
	textPad.insert(INSERT, "Base Folder Name =>  "+folderName+"\n")
	print(folderName)
	
def PyDeprecatedAttrs():
	
	try:
		textPad.insert(INSERT, "Base Folder Name =>  "+folderName+"\n")
	except:
		textPad.insert(INSERT, "\n######## Base Folder Name is not selected #######\n")
		textPad.insert(INSERT, "######## <<<<<< Error Caught >>>>>>> #######\n")
	#print('Input Folder Name :: '+folderName)
	# if user not fill any entry 
	# then print "empty input" 
	if (folderName == "" or folderName is None): 	
		print("Base Folder Location Not Selected")
		textPad.insert(INSERT, "\n######## Base Folder Name is not selected #######\n")
		textPad.insert(INSERT, "######## <<<<<< Error Caught >>>>>>> #######\n")		
	else: 
		print("Base Folder Location "+ folderName)
		#textPad.insert(INSERT, "<<<<< Base Folder Name >>>>>> "+folderName+"\n")		
		Base_Folder=folderName
		print('########  Execution Started #######')
		print ('Start Date Time :: '+str(datetime.datetime.now())+'\n')
		##Declaring file writer for meta data file
		MetaDataFileWrite = open(Base_Folder+'\\TargetMetaData\\BDMTData_Item_Rev_Type.dsv', "w+")
		DeprecatedAttrCSVtoXML = open(Base_Folder+'\\TargetMetaData\\DeprecatedAttrCSVtoXML.csv', "w+")

		#Print the Header File for Meta data and csvtoTCXML file data
		MetaDataFileWrite.write('ITEM_ID|REVISION|ITEM_TYPE\n')
		DeprecatedAttrCSVtoXML.write('!Dataset:object_name|Dataset:object_desc|Dataset:dataset_type|ImanFile:file_name|Item:item_id|Item:object_type|ItemRevision:item_revision_id|ImanRelation:relation_type|creation_date|last_mod_date|volume|sdpath\n')
	
		for root, dirs, files in os.walk(Base_Folder+'\\RawFile'):
			print('########  Execution Started #######')
			textPad.insert(INSERT, "########  <<<<< Execution Started for base Folder >>>>>> #######\n")
			#print ('Start Date Time :: '+str(datetime.datetime.now())+'\n')
			textPad.insert(INSERT, "<<<<<< Start Date Time :: "+str(datetime.datetime.now())+" >>>>>\n\n")

			#Reading all files in RawFile Folder
			for filename in files:
				print('Processing FileName :: '+filename)
				textPad.insert(INSERT, " Processing FileName >>>>>> "+filename+"\n") 
				HeaderCount = 0
				HeaderData = ''
				filepathname = Base_Folder+'\\RawFile\\'  + filename
				#Reading file
				with open(filepathname,'r') as file:
					for line in file:
						#print(HeaderCount)
						#Storing data Header File 
						if HeaderCount == 0 :
							HeaderData = line
							#print('##')
							#print(HeaderData)
							HeaderCount=HeaderCount+1
						else :
							#Storing data for other Row excluding header
							#print(line.split('|')[0])
							#print(line.split('|')[1])
							#print(line.split('|')[2])
							MetaDataFileWrite.write(line.split('|')[0]+'|')
							MetaDataFileWrite.write(line.split('|')[1]+'|')
							MetaDataFileWrite.write(line.split('|')[2]+'\n')
							## These value will be used as File name for each File
							UniqueFileName = line.split('|')[0]+'_'+line.split('|')[1]+'_'+line.split('|')[2]+'.dsv'
							#print(UniqueFileName)
							UniqueFileNamewithPath = Base_Folder+'\\Source\\'+ UniqueFileName
							UniqueFileWrite = open(UniqueFileNamewithPath, "w")
							#UniqueFileWrite.write(line.split('\n', 1)[0])
							UniqueFileWrite.write(HeaderData)
							#UniqueFileWrite.write('\n')
							#LineZ = filter(lambda x: x in string.printable, line)
							UniqueFileWrite.write(strip_non_ascii(line))
							HeaderCount=HeaderCount+1
							
							UniqueFileWrite.close()
							##Generate Physical Files
							UniqueFileXlsxFileName = Base_Folder+'\\TargetFiles\\'+line.split('|')[0]+'_'+line.split('|')[1]+'_'+line.split('|')[2]+'.xlsx'
							df=pd.read_csv(UniqueFileNamewithPath,dtype=str,encoding='utf-8',quoting=csv.QUOTE_NONE,sep="|").T
							df.to_excel(UniqueFileXlsxFileName,index=True,header=False)
							
							##Writing the value for CSV2TCXML
							DeprecatedAttrCSVtoXML.write(line.split('|')[0]+'/'+line.split('|')[1]+'|'+'DEPRECATED ATTRIBTES|MSExcel|'+line.split('|')[0]+'_'+line.split('|')[1]+'_'+line.split('|')[2]+'.xlsx'+'|'+line.split('|')[0]+'|'+line.split('|')[2]+'|'+line.split('|')[1]+'|IMAN_specification|2019/01/01 01:01:01|2019/01/01 01:01:01|volume|dba_5b14f752')
							DeprecatedAttrCSVtoXML.write('\n')
			print ('\nEnd  Date  Time :: '+str(datetime.datetime.now()))
			textPad.insert(INSERT, "\n<<<<<< End Date Time :: "+str(datetime.datetime.now())+" >>>>>\n")
			#print('File Processed for selected folder :'+ Base_Folder)
		# assigning the max row and max column 
		# value upto which data is written 
		# in an excel sheet to the variable 
		#current_row = sheet.max_row 
		#current_column = sheet.max_column 

		# get method returns current text 
		# as string which we write into 
		# excel spreadsheet at particular location 
		#sheet.cell(row=current_row + 1, column=1).value = Base_Folder_field.get() 

		# save the file 
		##wb.save('C:\\Users\\Admin\\Desktop\\excel.xlsx') 

		# set focus on the Base_Folder_field box 
		#Base_Folder_field.focus_set() 
		
		#Calling of Function to Run Generic Code
		
		# call the clear() function 
		clear() 

def strip_non_ascii(string):
    ''' Returns the string without non ASCII characters'''
    stripped = (c for c in string if 0 < ord(c) < 127)
    return ''.join(stripped)
	
	
# Driver code 
if __name__ == "__main__": 
	
	# create a GUI window 
	root = Tk() 
	# set the background colour of GUI window 
	root.configure(background='light Grey') 
	# set the title of GUI window 
	root.title("Deprecated Attribute Interface - Ahilya Version ") 
	# set the configuration of GUI window 
	root.geometry("900x500") 

	# create a Form label 
	heading = Label(root, text="Use to Create Meta Data and Excel sheet", bg="light Grey") 
	
	# create a Form label 
	FooterCprt = Label(root, text="\u00A9 Tata Consultancy Services Limited - 2019", fg="Dark Blue",bg="light Grey")

	# create a Name label 
	Base_Folder_Name = Label(root, text="Base Folder Location",fg="Black", bg="light Grey") 

	#root = Tk()
	#v = StringVar()
	#BrowseButton = Button(root,text="Browse",bg="light Grey",command=browse_button)
	folder_path = StringVar()
	lbl1 = Label(master=root,textvariable=folder_path)
	lbl1.grid(row=1, column=1)
	BrowseButton = Button(text="Browse",bg="light Grey", command=browse_button)
	#button2.grid(row=0, column=3)
	#folderName = filedialog.askdirectory()
	#print(folderName)
	#root = tk.Tk()
	textPad = ScrolledText(root)
	#textPad.pack()
	textPad.insert(INSERT, "-----------------------------------------\n")
	textPad.insert(INSERT, "###### Utilty Successfully Started #####\n")
	textPad.insert(INSERT, "-----------------------------------------\n")
	#textPad.insert(END, " in ScrolledText")
	#root.mainloop()
	# grid method is used for placing 
	# the widgets at respective positions 
	# in table like structure . 
	heading.grid(row=0, column=1) 
	Base_Folder_Name.grid(row=1, column=0)
	BrowseButton.grid(row=1, column=4)

	# create a text entry box 
	# for typing the information 
	#Base_Folder_field = Entry(root) 


	# grid method is used for placing 
	# the widgets at respective positions 
	# in table like structure . 
	#Base_Folder_field.grid(row=1, column=1, ipadx="100") 

	# create a Submit Button and place into the root window 
	submit = Button(root, text="Run Utility", fg="Black",bg="Red", command=PyDeprecatedAttrs) 
	submit.grid(row=8, column=1) 
	textPad.grid(row=10, column=1)
	FooterCprt.grid(row=12, column=1)
	# start the GUI 
	root.mainloop() 
