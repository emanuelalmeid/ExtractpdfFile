# importa as bibliotecas necessárias
from asyncio.windows_events import NULL
from enum import unique
import string
import openpyxl
import os
import shutil
import pandas as pd
import pdfminer
from pdfminer.high_level import extract_text
import re
from openpyxl import Workbook, load_workbook
import matplotlib as plt 
# If you need to get the column letter, also import this
from openpyxl.utils import get_column_letter


pst_org = "PASTAORG"
variab_test = 1

df = pd.read_excel('C:\\Users\\User\\Python_projects\\LeitorNF\\BASENOTASNFE.xlsx', sheet_name= 'BASENOTAS', dtype= {'Notas': 'object','TP' : 'object','IC':'object','origem':'object'})
df = df.iloc[:,:4]
df.iloc[:,1] =  df.iloc[:,1].str.replace("NOME/RAZÃO","TP ")
df.iloc[:,1] =  df.iloc[:,1].str.replace("\n","")
df.iloc[:,1] =  df.iloc[:,1].str.replace(" ","")
df.iloc[:,3] =  df.iloc[:,3].str.replace("\n","")     
TP = df.iloc[:,1].drop_duplicates(keep= 'first')
Origem = df.iloc[:,3].drop_duplicates(keep= 'first')
print(df)

print(df.iloc[:,3])

dir = 'C:\\Users\\User\\Python_projects\\'+pst_org+'\\'
os.mkdir(dir)

i=0
for i in range (i, TP.shape[0]):
    transp = str(TP.iloc[i]) 
    print(transp)
    dir = 'C:\\Users\\User\\Python_projects\\'+pst_org+'\\' + transp
    print('Pasta '+ transp + ' criada' )
    os.mkdir(dir)

i=0
for i in range (i, Origem.shape[0]):
    NFEOrigem = str(Origem.iloc[i]) 
    print(NFEOrigem)
    dir = 'C:\\Users\\User\\Python_projects\\'+pst_org+'\\' + NFEOrigem
    print('Pasta '+ NFEOrigem + ' criada' )
    os.mkdir(dir)




i=0
for i in range(i, df.shape[0]):
    oldAdress = 'C:\\Users\\User\\Python_projects\\NOTAS\\' + str(df.iloc[i][0]) + "NFD.pdf" #pasta origem
    newAdress = 'C:\\Users\\User\\Python_projects\\'+pst_org+'\\'+ str(df.iloc[i][1]) +'\\'+ str(df.iloc[i][0]) + "NFD.pdf" #pasta destino
    print(oldAdress+ " para tp " +newAdress)
    shutil.copy2(oldAdress, newAdress)

    #organizar origem

    oldAdress = 'C:\\Users\\User\\Python_projects\\NOTAS\\' + str(df.iloc[i][0]) + "NFD.pdf" #pasta origem
    newAdress = 'C:\\Users\\User\\Python_projects\\'+pst_org+'\\'+ str(df.iloc[i][3]) +'\\'+ str(df.iloc[i][0]) + "NFD.pdf" #pasta destino
    print(oldAdress+ " para origem " +newAdress)
    shutil.copy2(oldAdress, newAdress)









