
import loadingExcel as le
import generator as gg
import os, sys


oriPath = os.getcwd()
path = r"C:\Users\fxiao\Documents\Change request #\ACAa025"
file ="Flames Community Arena (Calgary) ACAa025 LTE MBI RF CIQ.xlsx"
dataset = le.load(path, file)
#print(dataset['LNCEL Name'])
os.chdir(oriPath)
sitecode="ACAa025 "
#gg.generate_unlock(dataset,sitecode)
#gg.generate_CM(dataset,sitecode)
gg.generate_SAS_info(dataset,sitecode)
#gg.generate_SCF(dataset,sitecode)


'''
#this is to test dataset avalibility
print(len(dataset["Data"]))
data = dataset["Data"][0][10]
#this is to indentify a 'none' data
print(data)
print(type(data))
print(data==None) 
'''
