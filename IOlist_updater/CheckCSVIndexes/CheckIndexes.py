import csv
import xlwings as xw

ip_address = '2.145.48.13'

directory_to_parse = '\\\\'+ ip_address +'\\C$\\PRODUCTION'

input_csv = directory_to_parse + "\\ConveyorSys\\SysGroup01\\Database\\DB_IOs\\Inputs.csv"
output_csv = directory_to_parse + "\\ConveyorSys\\SysGroup01\\Database\\DB_IOs\\Outputs.csv"
inputs ={}
outputs ={}

mnemonics ={}

with open(input_csv,mode='r') as inp:
    reader = csv.reader(inp, delimiter=";", dialect="excel")

    for row in reader:
        #print(row)
        try:
            if len(row) >10 and row[-6] is not None:
                "{0:0=4d}".format(int(cainet))
                inputs[row[1]] = "I."+"{0:0=4d}".format(int(row[11])) +"." + row[12]+"."+ row[13]
        except :
            print("except 1")
            pass
with open(output_csv,mode='r') as out:
    reader = csv.reader(out, delimiter=";", dialect="excel")

    for row in reader:
        #print(row)
        try:
            if len(row) >7 and row[-6] is not None:
                
                outputs[row[1]] = "O."+"{0:0=4d}".format(int(row[4])) +"." + row[5]+"."+ row[6]
        except :
            print("except 2")
            pass

#Apriamo la lista IO e prendiamo le colonne dei mnemonici e degli input/output

wb = xw.Book('8FOS_IO008_V5.2.xls')

hilscher_tabs = {1:1,2:2}  #e.g. {1:1,2:2}

for hilscher_card in hilscher_tabs.keys():

        print("elaborate hilscher card " +str(hilscher_card))

        sht = wb.sheets[hilscher_tabs[hilscher_card]]

        rownum = sht.range('E1').current_region.last_cell.row + 1 #lenght of document(num of rows) on the bit column
        
        R_Column =   sht.range('R1:R'+str(rownum)).value # Name of modules Column e.g. Inverter Nord SK250E
        E_Column =   sht.range('E1:E'+str(rownum)).value # 
        C_Column =   sht.range('C1:C'+str(rownum)).value # 

        for i in range(1,rownum):
            if R_Column[i] != "":
                mnemonics[R_Column[i]] = E_Column[i]


for i in inputs.keys():
    try:
        if inputs[i] != mnemonics[i]:
            print(i + " : " +inputs[i] + " not corresponding to io list:: " + inputs[i] + " != " + mnemonics[i])
    except:
        if i not in mnemonics.keys() and i[:3] not in ['INT','CHU','IOM','MOT','MST','1CT','ENM','MAN','RSM']:
            if i[:3] not in['PHT'] and i[-1] != '0': 
                print (i +" : "+ inputs[i] +" not present in menmonics")

for i in outputs.keys():
    try:
        if outputs[i] != mnemonics[i]:
            print(i + " : " +outputs[i] + " not corresponding to io list:: " + outputs[i] + " != " + mnemonics[i])
    except:
        if i not in mnemonics.keys() and i[:3] not in ['INT','CHU','IOM','MOT','MST','1CT','ENM','MAN','RSM']:
            #print(i[:3])
            if i[:3] not in['PHT'] and i[-1] != '0': 
                print (i +" : " + outputs[i]+ "  not present in menmonics")
            
#print(inputs)

