import xlwings as xw

wb = xw.Book('8FOS_IO008_V4.1.xls')

System = "S1."
real_cabinets = ["1","2","3"]

mnm_column = 'R'
tag_filter_col = 'Z'

sht = wb.sheets[2]

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
                _,line_number,_,belt_number = sht.range(mnm_column+str(y)).value.split('_')[4][-2:-1] #belt number e.g. 11,01
                line_number =   line_number[-1]
                belt_number = belt_number[1:]
                print(str(belt_number))
                break
            
        sht.range(mnm_column+str(i+1)).value = System+"IL0"+line_number+"BELTAL"+belt_number+"63"


#FOR LOOP to add Filter Tags for menmonics column
for i in range(1,rownum):

    if mnemonics[i] != None:

        #addr CABINET Alarms: e.g. S1.ILXX.CABINETALYYZZ
        if mnemonics[i][:12] == "STROBE_FAULT":
            _,_,line,_ = mnemonics[i].split("_")
            

        #add CABINET Alarms: e.g. S1.ILXX.CABINETALYYZZ
        if mnemonics[i][:6] == "CABINET":
            sht.range(tag_filter_col+str(i+1)).value = Line_num +mnemonics[i][:8] + mnemonics[:9:]
        
        #Add Zone WARNINGS: e.g. S1.UPXX.CABINETWRYYZZ
        if mnemonics[i][:4] == "WARQ":
            _,system,cabinet,bit = mnemonics[i].split("_")
            if cabinet[-2:] in real_cabinets or cabinet[-1] in real_cabinets: # check if it is a real cabinet

                sht.range(tag_filter_col+str(i+1)).value = UPnum + "CABINETWR0"+ cabinet + bit
            else:
                sht.range(tag_filter_col+str(i+1)).value = UPnum + "ZONEWR"+ cabinet + bit

           
        #Add BELTAL
        if mnemonics[i][:6] == "BELTAL":
            _,system,cabinet,bit = mnemonics[i].split("_")
            sht.range(tag_filter_col+str(i+1)).value = UPnum + "BELTAL"+ cabinet + bit
    
    
#FOR LOOP to add Filter Tags for aka menmonics column
for i in range(1,rownum):

    if aka_mnemonics[i] != None:

        #Add Zone ALARMS or CABINET Alarms: e.g. S1.UPXX.ZONEALYYZZ or S1.UPXX.CABINETALYYZZ
        if aka_mnemonics[i][:4] == "ALLQ":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            if cabinet[-2:] in real_cabinets or cabinet[-1] in real_cabinets: # check if it is a real cabinet

                sht.range(tag_filter_col+str(i+1)).value += "\n"+ UPnum + "CABINETAL0"+ cabinet + bit
            else:
                sht.range(tag_filter_col+str(i+1)).value += "\n" + UPnum + "ZONEAL"+ cabinet + bit

        if aka_mnemonics[i][:8] == "CABINETAL":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            sht.range(tag_filter_col+str(i+1)).value += "\n"+ UPnum + "CABINETAL0"+ cabinet + bit
        
        #Add Zone WARNINGS: e.g. S1.UPXX.CABINETWRYYZZ
        if aka_mnemonics[i][:4] == "WARQ":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            if cabinet[-2:] in real_cabinets or cabinet[-1] in real_cabinets: # check if it is a real cabinet

                sht.range(tag_filter_col+str(i+1)).value += "\n"+ UPnum + "CABINETWR0"+ cabinet + bit
            else:
                sht.range(tag_filter_col+str(i+1)).value +=  "\n" +UPnum + "ZONEWR"+ cabinet + bit

        if aka_mnemonics[i][:8] == "CABINETWR":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            sht.range(tag_filter_col+str(i+1)).value += "\n" + UPnum + "CABINETAL0"+ cabinet + bit
        
        #Add BELTAL
        if aka_mnemonics[i][:6] == "BELTAL":
            _,system,cabinet,bit = aka_mnemonics[i][:12].split("_")
            sht.range(tag_filter_col+str(i+1)).value += "\n"+ UPnum + "BELTAL"+ cabinet + bit
        
        



#wb.close()