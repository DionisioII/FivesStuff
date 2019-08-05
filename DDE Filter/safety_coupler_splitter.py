import xlwings as xw

wb = xw.Book('8FOS_SAFETY_IO_V2.6.xls')


sht = wb.sheets[1]
system="S1."
rownum_safety = sht.range('E1').current_region.last_cell.row + 1 #lenght of document(num of rows) on the bit column

print(rownum_safety)

mnemonics = sht.range('R1:R'+str(rownum_safety)).value  # Column where we put mnemonics
aka_mnemonics =  sht.range('S1:S'+str(rownum_safety)).value  # Column where we put aka mnemonics
J_Column =   sht.range('J1:J'+str(rownum_safety)).value # Node values Column  e.g. N0021
L_Column =   sht.range('L1:L'+str(rownum_safety)).value # Name of modules Column e.g. Inverter Nord SK250E
M_Column =   sht.range('M1:M'+str(rownum_safety)).value # Location Column e.g. +UC402.GF006
N_Column =   sht.range('N1:N'+str(rownum_safety)).value # Device Column e.g. .k3011
X_column =   sht.range('X1:X'+str(rownum_safety)).value # filter tree

dictionary = dict()
for i in range(1,rownum_safety):
        if L_Column[i] == "=SC01":
                dictionary[M_Column[i]] = X_column[i]
                if N_Column[i] != None:
                        dictionary[M_Column[i] + N_Column[i]] = X_column[i]

print(dictionary)

sht2 = wb.sheets[2]
rownum_coupler = sht2.range('E3').current_region.last_cell.row
rownum_coupler =413
G_Column = sht2.range('G1:G'+str(rownum_coupler)).value # Node values Column  e.g. N0021
F_Column = sht2.range('F1:F'+str(rownum_coupler)).value
T_Column2 = sht2.range('T1:T'+str(rownum_coupler)).value # tree
X_Column2 = sht2.range('X1:X'+str(rownum_coupler)).value

print(rownum_coupler)

for i in range(1,rownum_coupler):

        
        if G_Column[i] != None and G_Column[i][:9] == "EMERGENCY":
                try:
                        key = F_Column[i].split("+")[1]
                        key ="+"+key.replace(']','')
                        if key in dictionary.keys():
                                sht2.range('T'+str(i+1)).value = dictionary[key]
                                print("tag added:"+dictionary[key])
                                sht2.range('X'+str(i+1)).value = system + "CABINETAL" +G_Column[i][-5:]
                        else:
                                print(key +" not in dict")
                except:
                        print("not in list")
                
        if G_Column[i]!= None and G_Column[i][:7] == "CABINET":
                try:
                        key = F_Column[i].split("+")[1]
                        key ="+"+key.replace(']','')
                        if key in dictionary.keys():
                                sht2.range('T'+str(i+1)).value = dictionary[key]
                                print("tag added:"+dictionary[key])
                                sht2.range('X'+str(i+1)).value = system + G_Column[i]
                        else:
                                print(key + " not in dict")
                except:
                        print( " key not in list")
                


