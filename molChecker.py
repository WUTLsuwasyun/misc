#!/usr/bin/env python
# author SUWA Shunnosuke
# python scripts to extract existing Mol files 
import os
import pandas as pd
import glob
import shutil
import xlwings as xw
import openpyxl
import tkinter as tk
from tkinter import filedialog

root=tk.Tk()

# Choose new data file
fileName = filedialog.askopenfilename(filetypes=[('','*.xlsx')],title = "Choose new data file (xlsx)")
root.withdraw()
print(fileName)
crrtDir = fileName[:fileName.rfind('/')+1]
print(crrtDir)
crrtFile = fileName[fileName.rfind('/')+1:]
print(crrtFile)

# Choose reference folder
MolDir = filedialog.askdirectory(initialdir = crrtDir, title="Choose reference folder")

df_chemData = pd.read_excel(fileName)
size = len(df_chemData.index)

# makedir if no exist
dirName = crrtDir + fileName[fileName.rfind('_')+1:-5]
os.makedirs(dirName, exist_ok=True)


wb = xw.Book(fileName)
if "Mol Check" not in [sheet.name for sheet in wb.sheets]:
    wb.sheets['提供用'].copy(name='Mol Check') 
sht = wb.sheets["Mol Check"]

for i in range(size):
    tmpStr = df_chemData.loc[i, "Name"]
    tmpFilename =  glob.glob(MolDir+f'/*{tmpStr}*.mol')
    if (tmpFilename):
        print(str(df_chemData.loc[i,"No"])+": "+tmpStr + ": mol file found")
        sht.range(f'A{i+2}').color=(128,128,128)
        sht.range(f'B{i+2}').color=(128,128,128)
        shutil.copy(tmpFilename[0], f'{dirName}/{df_chemData.loc[i,"No"]}_{df_chemData.loc[i,"Name"]}.mol')
    else:
        print(str(df_chemData.loc[i,"No"])+": "+tmpStr + ": mol file not found ...")

wb.save(fileName)