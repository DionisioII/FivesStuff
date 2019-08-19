import xlwings as xw

from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(filename)

#wb = xw.Book('8FOS_IO002_V4.9.xls')

wb = xw.Book(filename)

UPnum = {"1":"S1.UP09.","2":"S1.UP09."}
real_cabinets = ["1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26"]

mnm_column = 'R'
tag_filter_col = 'Y'

sht = wb.sheets[1]

rownum = sht.range('E1').current_region.last_cell.row + 1 #lenght of document(num of rows) on the bit column

print(rownum)

mnemonics = sht.range('R1:R'+str(rownum)).value  # Column where we put mnemonics
aka_mnemonics =  sht.range('S1:S'+str(rownum)).value  # Column where we put aka mnemonics
C_Column =   sht.range('C1:C'+str(rownum)).value # Node values Column  e.g. N0021
A_Column =   sht.range('A1:A'+str(rownum)).value # Name of modules Column e.g. Inverter Nord SK250E
M_Column =   sht.range('M1:M'+str(rownum)).value # Location Column e.g. +UC402.GF006
Z_Column =   sht.range('Z1:Z'+str(rownum)).value # Tag Filter Column e.g S1.UPXX.CABINETALYYZZ

#FOR LOOP to Add BELTAL Mnemonics
for i in range(1,rownum):

    
    # Add BELTAL 
    
    if C_Column[i]!= None and M_Column[i] != None and (M_Column[i][-5:-3] in ["GF","SK","TB","SP","RB","WB"] ):
        #print(C_Column[i])
        belt_number = "000"
        for y in range(i+2,i+5):
            if sht.range(mnm_column+str(y)).value != None and sht.range(mnm_column+str(y)).value[:4] =="MAIN":
                belt_number = sht.range(mnm_column+str(y)).value.split('_')[4]
                print(str(belt_number))
                break
            
        sht.range(mnm_column+str(i+1)).value = "BELTAL_1_"+belt_number+"_62"