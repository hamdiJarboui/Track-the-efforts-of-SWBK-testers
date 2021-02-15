# -*- coding: utf-8 -*-
"""
Created on Wed Sep 16 06:12:07 2020

@author: $Hamtchi $
"""


import os
import re
import xlsxwriter


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('excelFile.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': 1})

def liste_fichier_repertoire(folder):
    file, rep= [], []
    for r, d, f in os.walk(folder):
        for a in d:
            rep.append(r + "\\" + a)
        for a in f:
            file.append(r + "\\" + a)
    return file, rep

folder = r"."
file, fold = liste_fichier_repertoire(folder)
fichiers=[]
folds=[]
for i, f in enumerate(file):
    fichiers.append(file)
    #print("fichier ", f)
    
for i, f in enumerate(fold):
    folds.append(fold)
    #print("répertoire ", f)
    
    
tcREady=0
tcdone=0
tmREady=0
tmdone=0
epic=''
tab=[["epic","tcREady","tcdone","tmREady","tmdone"]]
lignes=[]
for i, f in enumerate(file):

    filin = open(file[i], "r",encoding='utf-8')
    chaine=filin.readlines()
    for ligne in chaine:
        
        if re.search('^@series = TC_BCP43547\n',ligne):
            epic=ligne[:-1]
            epicList=epic.split("_")
            epic=epicList[-1]
            epicList=epic.split("P")
            epic=epicList[0]+"P-"+epicList[-1]
            
    
    if chaine.count("@series = TC_BCP43547\n")>=1 and chaine.count("@series = done\n")>=1 and chaine.count("@series = TC_A450\n"):
        tcdone=tcdone+1
    
    if chaine.count("@series = TC_BCP43547\n")>=1:
        tcREady=tcREady+1
    
    if chaine.count("@series = TM_BCP43547\n")>=1:
        tmready=tmdone+1
    
    if chaine.count("@series = TM_BCP43547\n")>=1 and chaine.count("@series = done\n")>=1 and chaine.count("@series = TM_A450\n"):
        tcdone=tcdone+1
    
    chaine=""
    filin.close()

lignes.append(epic)
lignes.append(tcREady)
lignes.append(tcdone)
lignes.append(tmREady)
lignes.append(tmdone)


if lignes!=["","","","",""]:
    tab.append(lignes)
    
#for i in tab :
#    print(i)
#print("the number of testCases with @series = TC_BCP43547=",epic)    
#print("tc ready = " ,tcREady)
#print("tcdone = " ,tcdone)
#print("tmREady = " ,tmREady)
#print("tmdone= " ,tmdone)


 # Start from the first cell below the headers.
row = 1
col = 0

for epic,tcREady,tcdone,tmREady,tmdone in (tab):
     worksheet.write_string(row, col,epic)
     worksheet.write(row, col + 1,tcREady)
     worksheet.write(row, col + 2,tcdone)
     worksheet.write(row, col + 3,tmREady)
     worksheet.write(row, col + 4,tmdone)
     row += 1
     
workbook.close()