# -*- coding: utf-8 -*-
"""
Created on Wed Sep 16 05:37:45 2020

@author: $Hamtchi $
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 16 05:31:13 2020

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
    print("fichier ", f)
    
for i, f in enumerate(fold):
    folds.append(fold)
    print("répertoire ", f)
tab=[["fileNAme","folderName","series","author","testCase"]]
for i, f in enumerate(file):
    series=""
    author=""
    testCase=""
    fileName=""
    folderName=""
    lignes=[]
    filin = open(file[i], "r")
    chaine=filin.readlines()
    if chaine.count("@series = TC_BCP43547\n")>=1:
        ##print("there is  ",chaine.count("@series = TC_BCP43547\n"), "file ")
       ## print("the path of the file is " ,file[i],"\n")  
        string=file[i]
        listeString=string.split("\\")
        fileName=listeString[-1]
        folderName=listeString[-2]
       
        for ligne in chaine:
            if re.search('^@series',ligne):
                series=series+" -  "+ligne[:-1]
               ## print(series)
            if re.search('^@author',ligne):
                author=author+" -  "+ligne[:-1]
               ## print(author)
            if re.search('^TestCase ',ligne):
                testCase=testCase+" -  "+ligne[:-1]
                ##print(testCase)
    lignes.append(fileName)
    lignes.append(folderName)
    lignes.append(series)
    lignes.append(author)
    lignes.append(testCase)
    if lignes!=["","","","",""]:
        tab.append(lignes)
    filin.close()
for i in tab :
    print(i)

 # Start from the first cell below the headers.
row = 1
col = 0

for fileNAme,folderName,series,author,testCase in (tab):
     worksheet.write_string(row, col,fileNAme)
     worksheet.write_string(row, col + 1,folderName)
     worksheet.write_string(row, col + 2,series)
     worksheet.write_string(row, col + 3,author)
     worksheet.write_string(row, col + 4,testCase)
     row += 1
     
workbook.close()
