# -*- coding: cp1251 -*-
import csv
import xlwt
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename

root = Tk()
root.withdraw()
name = askopenfilename(filetypes =(("CSV file", "*.csv"),
                                   ("All Files","*.*")),
                                   title = "Выберите файл...")
print (name)
input_file = open(name, "r")
rdr = csv.DictReader(input_file, fieldnames=['Description', 'Designator'])

font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.height = 320

alignment0 = xlwt.Alignment()
alignment0.shrink_to_fit = True

style0 = xlwt.XFStyle()
style0.font = font0
style0.alignment = alignment0

wb = xlwt.Workbook()
ws = wb.add_sheet('ПК6Ц.270.000 ВП')

gen_list = []
L1 = ['', '']

for rec in rdr:
    try:
        if str(L1[0]) == str(rec['Description']):
            L1 [1] = L1[1] + ', ' + str(rec['Designator'])
        else:
            gen_list.append(list(L1))
            L1 [0] = str(rec['Description'])
            L1 [1] = str(rec['Designator'])
    except: pass 

gen_list.append(list(L1))    

if (gen_list [0] [0] == ''):
    gen_list.pop(0) 

for i in gen_list:
    for j in i:
        if j.find ('Р1-8МП-0,1-') != -1:
           j.replace ('Р1-8МП-0,1-', 'ss')

for i in range(len(gen_list)):
    for j in range(len(gen_list[i])):
        ws.write(i, j, gen_list[i] [j], style0)        
        
input_file.close()
wb.save('ПК6Ц.270.000.xls')

print ('Done!')
