import xlwings as xw
import csv

wb = xw.Book("2018_11_30_FOS - Description_conveyor list.xls")

sht = wb.sheets[0]

rownum = sht.range('BX7').current_region.last_cell.row + 1 #lenght of document(num of rows) on the bit column

Linea_Column =   sht.range('BX7:BX'+str(rownum)).value # Node values Column  e.g. N0021
GF_Column = sht.range('BY7:BY'+str(rownum)).value
number_column = sht.range('BZ7:BZ'+str(rownum)).value
value_column = sht.range('CA7:CA'+str(rownum)).value
encoder_column = sht.range('BU7:BU'+str(rownum)).value

print(len(Linea_Column))
Belts ={}
for i in range(0,rownum):
    if i >= len(Linea_Column):
        print(i)
        break
    if Linea_Column[i] != None and GF_Column[i] != None and number_column[i] != None and value_column[i] != None:
        
        Belts[Linea_Column[i] + GF_Column[i] + number_column[i]] = [value_column[i],encoder_column[i]]

###### CSV SECTION###########
output_csv= "t_nord.csv"

def all_motors_with_encoder(writer):
    writer.writerow(['@',hilscher_card,nodo,301,0,0,5,1,'#']) # FOR ALL MOTORS with encoder
    writer.writerow(['@',hilscher_card,nodo,327,0,1,100,1,'#']) # FOR ALL MOTORS with encoder
    writer.writerow(['@',hilscher_card,nodo,328,0,1,5,1,'#']) # FOR ALL MOTORS with encoder
    writer.writerow(['@',hilscher_card,nodo,420,1,0,42,1,'#']) # FOR ALL MOTORS with encoder

def all_motors(writer):
    writer.writerow(['@',hilscher_card,nodo,481,9,0,35,1,'#']) # FOR ALL MOTORS
    writer.writerow(['@',hilscher_card,nodo,506,0,0,0,1,'#']) # FOR ALL MOTORS
    writer.writerow(['@',hilscher_card,nodo,509,0,0,3,1,'#']) # FOR ALL MOTORS

def rulliera(writer):
    writer.writerow(['@',hilscher_card,nodo,102,0,1,80,1,'#']) # for rulliere
    writer.writerow(['@',hilscher_card,nodo,209,0,1,10,1,'#']) # for rulliere

hilscher_tabs = {1:1}  #e.g. {1:1,2:2}

wb = xw.Book('8FOS_IO013_V4.0.xls')

with  open(output_csv, 'w', newline='') as out:
    writer = csv.writer(out, delimiter=",", dialect="excel")
    #writer.writerow(head)
    out.write('#,500,,,,,,,\n')
    out.write("#,Card ID,Node ID,Parameter ID,Parameter subindex,Parameter set,Parameter value,Write to EEPROM,# Comment\n")
    


    for hilscher_card in hilscher_tabs.keys():

        print("elaborate hilscher card " +str(hilscher_card))

        sht = wb.sheets[hilscher_tabs[hilscher_card]]

        rownum = sht.range('E1').current_region.last_cell.row + 1 #lenght of document(num of rows) on the bit column
        
        A_Column =   sht.range('A1:A'+str(rownum)).value # Name of modules Column e.g. Inverter Nord SK250E
        B_Column =   sht.range('B1:B'+str(rownum)).value # 
        C_Column =   sht.range('C1:C'+str(rownum)).value # 



        for i in range(0,rownum):
            if A_Column[i] =="Inverter Nord SK250E" :

                belt = B_Column[i][5:-3]
                
                nodo = C_Column[i-1][1:]
                print(nodo)
                Nodo= "N" + "{0:0=4d}".format(int(nodo))
                #check if belt has value::: if not go to next iteration
                if belt not in Belts.keys():
                    print(belt + " non parametrizzata ; Nodo: "+ Nodo )
                    continue
                motor_type = Belts[belt][0]
                enc_string = Belts[belt][1]
                nodo = str(int(Nodo[1:]))

                # 90T3/4 
                if motor_type in ["90T3/4 TF TI4 IG22","SK92372.1V-90T3/4 TF","SK92372.1AMH-90 T3/4 Y_10,33_no BRE_IG22"]:
                    if enc_string[-6:] in ["+HTL+R"] or enc_string[-4:] in ["+HTL"]: #if there is encoder

                        out.write("#" + Nodo +'  motor type = 90T3/4, encoder = yes   ### '+ belt +' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,105,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
                        #writer.writerow(['@',hilscher_card,nodo,301,0,0,5,1,'#'])
                        all_motors_with_encoder(writer)
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)
                    else:
                        out.write("#" + Nodo +'  motor type = 90T3/4, encoder = no    ### '+ belt +' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,105,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)
                
                #90T1/4    
                elif motor_type in ["SK92372.1AMH-90 T1/4 Y_10,33_no BRE_IG22","SK 92372.1 AMHD"]:

                    if enc_string[-6:] in ["+HTL+R"] or enc_string[-4:] in ["+HTL"]: #if there is encoder

                        out.write("#" + Nodo +'  motor type = 90T1/4, encoder = yes    ### '+ belt +' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,101,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
                        all_motors_with_encoder(writer)
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)
                    else:
                        out.write("#"+Nodo +'  motor type = 90T1/4, encoder = no    ### '+ belt + ' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,101,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)

                #80T1/4 
                elif motor_type in ["SK372.1F-80T1/4 TF TI4","SK92172.1AMH-80 T1/4 Y_10,83_no BRE_IG22","SK92372.1AMH-80 T1/4 Y_25,06_no BRE_IG22","SK 92172.1 AMHD","SK92172.1AMH-80 T1/4 Y_12,18_no BRE_no IG","190A_00054"]:

                    if enc_string != None and enc_string[-6:] in ["+HTL+R"] or enc_string[-4:] in ["+HTL"]: #if there is encoder

                        out.write("#" + Nodo +'  motor type = 80T1/4, encoder = yes    ### '+ belt +' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,98,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,0,1,'#'])
                        all_motors_with_encoder(writer)
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)
                    else:
                        out.write("#"+Nodo +'  motor type = 80T1/4, encoder = no    ### '+ belt + ' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,98,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,0,1,'#'])
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)
                
                #100T2/4
                elif motor_type in ["SK92372.1AMH-100 T2/4 Y_10,33_no BRE_IG22"]:

                    if enc_string[-6:] in ["+HTL+R"] or enc_string[-4:] in ["+HTL"]: #if there is encoder

                        out.write("#"+Nodo +'  motor type = 80T1/4, encoder = yes    ### '+ belt +' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,109,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,0,1,'#'])
                        all_motors_with_encoder(writer)
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)
                    else:
                        out.write("#"+Nodo +'  motor type = 80T1/4, encoder = no    ### '+ belt + ' \n')
                        writer.writerow(['@',hilscher_card,nodo,200,0,1,109,1,'#'])
                        writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
                        all_motors(writer)
                        if belt[-5:-3] == "RB":
                            rulliera(writer)
                else:
                    try:
                        print(belt+ " con tipologia " + Belts[belt]+ "  non parametrizzato. Nodo: " + nodo)
                    except:
                        print(belt+ " con tipologia " + " non chiara "+ "  non parametrizzato. Nodo: " + nodo)



out.close()

















