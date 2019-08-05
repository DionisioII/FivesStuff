import xlwings as xw

from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(filename)

#wb = xw.Book('8FOS_IO002_V4.9.xls')

wb = xw.Book(filename)

UPnum = {"1":"S1.UP07."}
real_cabinets = ["1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26"]

mnm_column = 'R'
tag_filter_col = 'X'

sht = wb.sheets[4]

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
    
    if C_Column[i]!= None and M_Column[i] != None and (M_Column[i][-5:-3] in ["GF","SK","TB","SP","WB","RB"] ):
        #print(C_Column[i])
        belt_number = "000"
        for y in range(i+2,i+5):
            if sht.range(mnm_column+str(y)).value != None and sht.range(mnm_column+str(y)).value[:4] =="MAIN":
                system = sht.range(mnm_column+str(y)).value.split('_')[2]
                belt_number = sht.range(mnm_column+str(y)).value.split('_')[4]
                print(str(belt_number))
                break
            
        sht.range(mnm_column+str(i+1)).value = "BELTAL_"+ system+"_"+belt_number+"_62" 

mnemonics = sht.range('R1:R'+str(rownum)).value  # Column where we put mnemonics
aka_mnemonics =  sht.range('S1:S'+str(rownum)).value  # Column where we put aka mnemonics
C_Column =   sht.range('C1:C'+str(rownum)).value # Node values Column  e.g. N0021
A_Column =   sht.range('A1:A'+str(rownum)).value # Name of modules Column e.g. Inverter Nord SK250E
M_Column =   sht.range('M1:M'+str(rownum)).value # Location Column e.g. +UC402.GF006
Z_Column =   sht.range('Z1:Z'+str(rownum)).value # Tag Filter Column e.g S1.UPXX.CABINETALYYZZ   

#FOR LOOP to add Filter Tags for menmonics column
for i in range(1,rownum):

    if mnemonics[i] != None:

        #Add Zone ALARMS or CABINET Alarms: e.g. S1.UPXX.ZONEALYYZZ or S1.UPXX.CABINETALYYZZ
        if mnemonics[i][:4] == "ALLQ":
            _,system,cabinet,bit = mnemonics[i].split("_")
            if cabinet[-2:] in real_cabinets or cabinet[-1] in real_cabinets: # check if it is a real cabinet
                cabinet_int = int(cabinet) -1
                cabinet =  "{0:0=3d}".format(cabinet_int)
                sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "CABINETAL"+ cabinet + bit
            else:
                sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "ZONEAL"+ cabinet + bit

        if mnemonics[i][:8] == "CABINETAL":
            _,system,cabinet,bit = mnemonics[i].split("_")
            sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "CABINETAL0"+ cabinet + bit
        
        #Add Zone WARNINGS: e.g. S1.UPXX.CABINETWRYYZZ
        if mnemonics[i][:4] == "WARQ":
            _,system,cabinet,bit = mnemonics[i].split("_")
            if cabinet[-2:] in real_cabinets or cabinet[-1] in real_cabinets: # check if it is a real cabinet
                cabinet_int = int(cabinet) -1
                cabinet =  "{0:0=3d}".format(cabinet_int)
                sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "CABINETWR"+ cabinet + bit
            else:
                sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "ZONEWR"+ cabinet + bit

        if mnemonics[i][:8] == "CABINETWR":
            _,system,cabinet,bit = mnemonics[i].split("_")
            sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "CABINETWR0"+ cabinet + bit
        
        #Add BELTAL
        if mnemonics[i][:6] == "BELTAL":
            _,system,belt,bit = mnemonics[i].split("_")
            belt="{0:0=3d}".format(int(belt))
            sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "BELTAL"+ belt + bit +"\n"  +\
                                                       UPnum[system] + "BELTWR"+ belt + bit
            for t in range(1,5):
                if mnemonics[i+t] != None and mnemonics[i+t][:9] =="MAIN_JREV":
                    sht.range(tag_filter_col+str(i+1 + t)).value = UPnum[system] + "BELTAL"+ belt + "00" +"\n" +\
                                                                   UPnum[system] + "BELTAL"+ belt + "04" +"\n"
                    break 
        elif  len(mnemonics[i]) == 13 and mnemonics[i][-1] =="0" and mnemonics[i][:8] == "DMFBK_2_" and False :
            print(mnemonics[i])
            _,system,belt,bit = mnemonics[i].split("_")
            belt="{0:0=3d}".format(int(belt))
            sht.range(tag_filter_col+str(i+1)).value = UPnum[system] + "BELTAL"+ belt + "00" +"\n"  +\
                                                       UPnum[system] + "BELTWR"+ belt + "04"
            print("dmfKB: " +cabinet)




#FOR LOOP to add Filter Tags for aka menmonics column
for i in range(1,rownum):
    break

    if aka_mnemonics[i] != None:

        #Add Zone ALARMS or CABINET Alarms: e.g. S1.UPXX.ZONEALYYZZ or S1.UPXX.CABINETALYYZZ
        print(aka_mnemonics[i])
        if aka_mnemonics[i][:4] == "ALLQ":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            if cabinet[-2:] in real_cabinets or cabinet[-1] in real_cabinets: # check if it is a real cabinet
                cabinet_int = int(cabinet) -1
                cabinet =  "{0:0=3d}".format(cabinet_int)
                if sht.range(tag_filter_col+str(i+1)).value == None:
                    sht.range(tag_filter_col+str(i+1)).value =" "
                sht.range(tag_filter_col+str(i+1)).value +="\n"+ UPnum[system] + "CABINETAL"+ cabinet + bit
            else:
                sht.range(tag_filter_col+str(i+1)).value += "\n" + UPnum[system] + "ZONEAL"+ cabinet + bit

        if aka_mnemonics[i][:8] == "CABINETAL":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            sht.range(tag_filter_col+str(i+1)).value += "\n"+ UPnum[system] + "CABINETAL0"+ cabinet + bit
        
        #Add Zone WARNINGS: e.g. S1.UPXX.CABINETWRYYZZ
        if aka_mnemonics[i][:4] == "WARQ":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            if cabinet[-2:] in real_cabinets or cabinet[-1] in real_cabinets: # check if it is a real cabinet
                cabinet_int = int(cabinet) -1
                cabinet =  "{0:0=3d}".format(cabinet_int)
                sht.range(tag_filter_col+str(i+1)).value += "\n"+ UPnum[system] + "CABINETWR"+ cabinet + bit
            else:
                sht.range(tag_filter_col+str(i+1)).value +=  "\n" +UPnum[system] + "ZONEWR"+ cabinet + bit

        if aka_mnemonics[i][:8] == "CABINETWR":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            sht.range(tag_filter_col+str(i+1)).value += "\n" + UPnum[system] + "CABINETWR0"+ cabinet + bit
        
        #Add BELTAL
        if aka_mnemonics[i][:6] == "BELTAL":
            _,system,belt,bit = aka_mnemonics[i][:12].split("_")
            belt="{0:0=3d}".format(int(belt))
            sht.range(tag_filter_col+str(i+1)).value += "\n"+ UPnum[system] + "BELTAL"+ belt + bit
        
        



#wb.close()