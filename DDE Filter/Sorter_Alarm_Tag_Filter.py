import xlwings as xw

#wb = xw.Book('8FOS_IO008_V4.1.xls')
wb = xw.Book('8FOS_IO009_V4.3.xlsx')

Sorter = "S1."
real_cabinets = ["1","2","3"]

mnm_column = 'R'
tag_filter_col = 'Z'

modulator_count = 1
drive_unit_count = 1
chuteDP_count = 0
chuteDJ_count = 0

sht = wb.sheets[4]

rownum = sht.range('E1').current_region.last_cell.row + 1 #lenght of document(num of rows) calculated on the bit column

print(rownum)

mnemonics = sht.range('R1:R'+str(rownum)).value  # Column where we put mnemonics
aka_mnemonics =  sht.range('S1:S'+str(rownum)).value  # Column where we put aka mnemonics
C_Column =   sht.range('C1:C'+str(rownum)).value # Node values Column  e.g. N0021
A_Column =   sht.range('A1:A'+str(rownum)).value # Name of modules Column e.g. Inverter Nord SK250E
M_Column =   sht.range('M1:M'+str(rownum)).value # Location Column e.g. +UC402.GF006
Z_Column =   sht.range('Z1:Z'+str(rownum)).value # Tag Filter Column e.g S1.CABINETALYYZZ


#FOR LOOP to add Filter Tags for menmonics column
for i in range(1,rownum):

    if mnemonics[i] != None:

        #Add Filter Tags for Control Zone: e.g. S1.CONTROL01.ALARM1.01
        if mnemonics[i][:8] == "CTRLZONE":
            print(mnemonics[i].split("_"))
            ctrlNum = mnemonics[i].split("_")[2][0]
            sht.range(tag_filter_col+str(i+1)).value =            Sorter + "CONTROL0"+ ctrlNum + ".ALARM1.01" \
                                                         + "\n" + Sorter + "CONTROL0"+ ctrlNum + ".ALARM1.02" \
                                                         + "\n" + Sorter + "CONTROL0"+ ctrlNum + ".ALARM1.03" \
                                                         + "\n" + Sorter + "CONTROL0"+ ctrlNum + ".ALARM1.04" \
                                                         + "\n" + Sorter + "CONTROL0"+ ctrlNum + ".WARNING1.01" \
                                                         + "\n" + Sorter + "CONTROL0"+ ctrlNum + ".WARNING1.02" \
                                                         + "\n" + Sorter + "CONTROL0"+ ctrlNum + ".WARNING1.03" \
                                                         + "\n" + Sorter + "CONTROL0"+ ctrlNum + ".WARNING1.04" 
        #Add Filter Tags for Check Position: e.g. S1.CHECKPOS01.ALARM1.08
        if mnemonics[i][:6] == "CHKPOS":
            chklNum = mnemonics[i].split("_")[2][0]
            sht.range(tag_filter_col+str(i+1)).value =            Sorter + "CHECKPOS0"+ chklNum + ".ALARM1.08" \
                                                         + "\n" + Sorter + "CHECKPOS0"+ chklNum + ".ALARM1.09" \
                                                         + "\n" + Sorter + "CHECKPOS0"+ chklNum + ".ALARM1.10" \
                                                         + "\n" + Sorter + "CHECKPOS0"+ chklNum + ".ALARM1.11" \
                                                         + "\n" + Sorter + "CHECKPOS0"+ chklNum + ".ALARM1.12" \
                                                         + "\n" + Sorter + "CHECKPOS0"+ chklNum + ".WARNING1.00" \
                                                         
        
        #Add Filter Tags for Modulators: e.g. S1.MODULATOR001.ALARM1.00
        if mnemonics[i][:13] == "MODULATOR_DEI":
            modulatorNum = "{0:0=3d}".format(modulator_count)
            sht.range(tag_filter_col+str(i+1)).value =            Sorter + "MODULATOR"+ modulatorNum + ".ALARM1.00" \
                                                         + "\n" + Sorter + "MODULATOR"+ modulatorNum + ".ALARM1.01" \
                                                         + "\n" + Sorter + "MODULATOR"+ modulatorNum + ".ALARM1.02" \
                                                         + "\n" + Sorter + "MODULATOR"+ modulatorNum + ".ALARM1.03" \
                                                         + "\n" + Sorter + "MODULATOR"+ modulatorNum + ".ALARM1.04" \
                                                         + "\n" + Sorter + "MODULATOR"+ modulatorNum + ".ALARM1.05" \
                                                         + "\n" + Sorter + "MODULATOR"+ modulatorNum + ".ALARM1.06" \
                                                         + "\n" + Sorter + "MODULATOR"+ modulatorNum + ".WARNING1.07"    
            modulator_count +=1
        
        #Add Filter Tags for EBOX TEST: e.g. S1.EBOXTEST01.ALARM1.00
        if mnemonics[i][:9] == "EBOX_TEST":
            ebox_num = mnemonics[i][-1]
            sht.range(tag_filter_col+str(i+1)).value =            Sorter + "EBOXTEST0"+ ebox_num + ".ALARM1.00" \
                                                         + "\n" + Sorter + "EBOXTEST0"+ ebox_num + ".ALARM1.01"

        #Add Filter Tags for CELL TEST: e.g. S1.CELLTEST01.ALARM1.00
        if mnemonics[i][:21] == "DEI_GATEWAY_CELL_TEST":
            cellTest_num = mnemonics[i][-1]
            sht.range(tag_filter_col+str(i+1)).value =            Sorter + "CELLTEST0"+ cellTest_num + ".ALARM1.00" \
                                                         + "\n" + Sorter + "CELLTEST0"+ cellTest_num + ".ALARM1.01" \
                                                         + "\n" + Sorter + "CELLTEST0"+ cellTest_num + ".ALARM1.02" \
                                                         + "\n" + Sorter + "CELLTEST0"+ cellTest_num + ".ALARM1.03" \
                                                         + "\n" + Sorter + "CELLTEST0"+ cellTest_num + ".ALARM1.04"
        
        #Add Filter Tags for missing Blade: e.g. S1.BLADECTRL01.ALARM1.00
        if mnemonics[i][:15] == "MISSINGBLADE_PH" and mnemonics[i][-1] == "1":
            bladeCtrl_num = mnemonics[i][-2]
            if bladeCtrl_num == "0":
                bladeCtrl_num = mnemonics[i][-3:-1]
            else:
                bladeCtrl_num = "0" + bladeCtrl_num
            print(mnemonics[i] + str(bladeCtrl_num))
            sht.range(tag_filter_col+str(i+1)).value =  Sorter + "BLADECTRL"+bladeCtrl_num + ".ALARM1.00"

        #Add Filter Tags for ENCODER: e.g. S1.ENCODER01.ALARM1.03
        if mnemonics[i][:10] == "CELLZERO_M" and mnemonics[i][-1] == "1":
            print(mnemonics[i])            
            sht.range(tag_filter_col+str(i+1)).value =  Sorter + "ENCODER01.ALARM1.03" \
                                                        + "\n" + Sorter + "ENCODER01.ALARM1.04"
        
        #Add Filter Tags for Drive Units: e.g. S1.DRIVEUNIT007.ALARM1.01
        if mnemonics[i][:8] == "STATWORD":
            print(mnemonics[i])
            driveNum = "{0:0=3d}".format(drive_unit_count)
            sht.range(tag_filter_col+str(i+1)).value =            Sorter + "DRIVEUNIT"+ driveNum + ".ALARM1.00" \
                                                         + "\n" + Sorter + "DRIVEUNIT"+ ebox_num + ".ALARM1.01" \
                                                         + "\n" + Sorter + "DRIVEUNIT"+ ebox_num + ".ALARM1.02" \
                                                         + "\n" + Sorter + "DRIVEUNIT"+ ebox_num + ".ALARM1.03" \
                                                         + "\n" + Sorter + "DRIVEUNIT"+ ebox_num + ".ALARM1.04"
            drive_unit_count +=1


        #Add Filter Tags for CABINETAL or CABINETWR: e.g. S1.CABINETAL01563, S1.CABINETWR01563
        if mnemonics[i][:7] == "CABINET":
            
            sht.range(tag_filter_col+str(i+1)).value =  Sorter +mnemonics[i]
        
        #Add Filter Tags for CHUTES Full: e.g. S1.CHUTE0000.ALARM2.01
        if mnemonics[i][:2] == "DP":
            chuteNumber = "{0:0=4d}".format(chuteDP_count)
            sht.range(tag_filter_col+str(i+1)).value =  Sorter + "CHUTE"+ chuteNumber + ".ALARM2.01"
            chuteDP_count +=1
        
        #Add Filter Tags for CHUTES JAM: e.g. S1.CHUTE0000.ALARM2.01
        if mnemonics[i][:2] == "DJ":
            chuteNumber = "{0:0=4d}".format(chuteDJ_count)
            sht.range(tag_filter_col+str(i+1)).value =  Sorter + "CHUTE"+ chuteNumber + ".ALARM2.02"
            chuteDJ_count +=1

        
        



#wb.close()