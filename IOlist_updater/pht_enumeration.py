import xlwings as xw

wb = xw.Book('8FOS_IO004_V4.3.xls')

mnm_column = 'R'
tag_filter_col = 'Z'

sht = wb.sheets[4]

rownum = sht.range('E1').current_region.last_cell.row + 1 #lenght of document(num of rows) on the bit column

print(rownum)

mnemonics = sht.range('R1:R'+str(rownum)).value  # Column where we put mnemonics
C_Column =   sht.range('C1:C'+str(rownum)).value # Node values Column  e.g. N0021
A_Column =   sht.range('A1:A'+str(rownum)).value # Name of modules Column e.g. Inverter Nord SK250E
M_Column =   sht.range('M1:M'+str(rownum)).value # Location Column e.g. +UC402.GF006



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
    
        
        



#wb.close()