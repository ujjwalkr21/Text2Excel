# cd C:\Users\Ujjwal\AppData\Local\Programs\Python\Python37-32\
# python "C:\Users\1297418\AppData\Local\Programs\Python\Python37-32\Codepy\PyDeprecatedAttr\Code\PyDeprecatedAttrs.py"
# Purpose : Reading the Multi Table File and create Excel file with Meta data used as input for CSVtoTCXML Utility
# Creator : Ujjwal Kumar
# History : Version 1.1 (Date : 11th April, 2019)
# 			Version 1.2 (Date : 16th April, 2019) :: Fixed the bug for unprintable character
# Warning : Do not hange identation as this python code which is based on identation
#		  : First three column will be used for File name. ITEM_ID_REV_TYPE
# Improvement Scope : UI Based or removal of few hard-coded path
# E:\\WorkInProgress\\OnePLM\\Python\\Code\\DeprecatedAttr

#Importing all required package
import os
import pandas as pd
import datetime
import string
import csv

''' Reading all files in RawFile Folder '''
def strip_non_ascii(string):
    ''' Returns the string without non ASCII characters'''
    stripped = (c for c in string if 0 < ord(c) < 127)
    return ''.join(stripped)

''' Set the value of base folder '''
Base_Folder='E:\\WorkInProgress\\OnePLM\\DeprecatedAttr'

''' Declaring file writer for meta data file '''
MetaDataFileWrite = open(Base_Folder+'\\TargetMetaData\\BDMTData_Item_Rev_Type.dsv', "w+")
DeprecatedAttrCSVtoXML = open(Base_Folder+'\\TargetMetaData\\DeprecatedAttrCSVtoXML.csv', "w+")	
for root, dirs, files in os.walk(Base_Folder+'\\RawFile'):
	print('########  Execution Started #######')
	print ('Start Date Time :: '+str(datetime.datetime.now())+'\n')
	''' Declaring file writer for meta data file '''
	''' Changing append type to write type '''
	#MetaDataFileWrite = open('C:\\Users\\1297418\\AppData\\Local\\Programs\\Python\\Python37-32\\Codepy\\PyDeprecatedAttr\\TargetMetaData\\BDMTData_Item_Rev_Type.dsv', "a+")
	#DeprecatedAttrCSVtoXML = open('C:\\Users\\1297418\\AppData\\Local\\Programs\\Python\\Python37-32\\Codepy\\PyDeprecatedAttr\\TargetMetaData\\DeprecatedAttrCSVtoXML.csv', "a+")
	
	''' Print the Header File for Meta data and csvtoTCXML file data '''
	MetaDataFileWrite.write('ITEM_ID|REVISION|ITEM_TYPE\n')
	DeprecatedAttrCSVtoXML.write('!Dataset:object_name|Dataset:object_desc|Dataset:dataset_type|ImanFile:file_name|Item:item_id|Item:object_type|ItemRevision:item_revision_id|ImanRelation:relation_type|creation_date|last_mod_date|volume|sdpath\n')	
	
	''' Reading all files in RawFile Folder '''
	for filename in files:
		print('Processing FileName :: '+filename)
		HeaderCount = 0
		HeaderData = ''
		filepathname = Base_Folder+'\\RawFile\\'  + filename
		''' Reading file '''
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
					''' Storing data for other Row excluding header '''
					#print(line.split('|')[0])
					#print(line.split('|')[1])
					#print(line.split('|')[2])
					MetaDataFileWrite.write(line.split('|')[0]+'|')
					MetaDataFileWrite.write(line.split('|')[1]+'|')
					MetaDataFileWrite.write(line.split('|')[2]+'\n')
					
					''' These value will be used as File name for each File '''
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
					''' Generate Physical Files '''
					UniqueFileXlsxFileName = Base_Folder+'\\TargetFiles\\'+line.split('|')[0]+'_'+line.split('|')[1]+'_'+line.split('|')[2]+'.xlsx'
					df=pd.read_csv(UniqueFileNamewithPath,dtype=str,encoding='utf-8',quoting=csv.QUOTE_NONE,sep="|").T
					df.to_excel(UniqueFileXlsxFileName,index=True,header=False)

					''' Writing the value for CSV2TCXML '''
					DeprecatedAttrCSVtoXML.write(line.split('|')[0]+'/'+line.split('|')[1]+'|'+'DEPRECATED ATTRIBTES|MSExcel|'+line.split('|')[0]+'_'+line.split('|')[1]+'_'+line.split('|')[2]+'.xlsx'+'|'+line.split('|')[0]+'|'+line.split('|')[2]+'|'+line.split('|')[1]+'|IMAN_specification|2019/01/01 01:01:01|2019/01/01 01:01:01|volume|dba_5b14f752')
					DeprecatedAttrCSVtoXML.write('\n')
		
print ('\nEnd  Date  Time :: '+str(datetime.datetime.now()))
print('#########  End Game #########')
		