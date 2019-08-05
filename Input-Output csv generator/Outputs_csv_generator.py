import xlrd
import csv



def line(Output_Name,Node,offset,bit, Output_Type,NC_NO_Logic = 0, default_value=0,board_number=1,Description=''):
    return  ['@', Output_Name,	Output_Type,	NC_NO_Logic,Node,	offset,	bit,	default_value,	board_number,Description]

def get_profinet_address(IO_address):
    values = IO_address.split('.')
    Node = values[1][-3:]
    if Node[0] == '0' :
        Node = Node[-2:]
    offset = values[2]
    bit = values[3]
    #print(IO_address[2])
    return Node,offset,bit

num_chute = 2
head = ['@','300']
header = "#	Output Name	Output Type	NC/NO Logic	Node	Offset	Bit	Default Value (For Virtual I/Os)	BOARD NUMBER	Description"
header_list = header.split('\t' )


output_csv= "output.csv"


book = xlrd.open_workbook('8FOS_IO013_V5.1.xls')

sh1 = book.sheet_by_index(1)

with  open(output_csv, 'w', newline='') as out:
    writer = csv.writer(out, delimiter=";", dialect="excel")
    writer.writerow(head)
    writer.writerow(header_list)

#LAMP, SIREN,PEXCEPT_LAMP, OTK_LAMP
    for rx in range(1, sh1.nrows):


        if sh1.row(rx)[17].value.startswith('LAMP0') or sh1.row(rx)[17].value.startswith('LAMPA') or sh1.row(rx)[17].value.startswith('LAMPE') :
                    
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type = 2
            
            writer.writerow(line(mnemonic,node,offset,bit,output_type,Description = description))

        elif sh1.row(rx)[17].value.startswith('SIREN'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type= 2
            
            writer.writerow(line(mnemonic,node,offset,bit,output_type,Description = description))

        elif sh1.row(rx)[17].value.startswith('PEXCEPT_LAMP'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type= 2
            
            writer.writerow(line(mnemonic,node,offset,bit,output_type,Description = description))

        elif sh1.row(rx)[17].value.startswith('OTK_LAMP'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type= 2
            
            writer.writerow(line(mnemonic,node,offset,bit,output_type,Description = description))

        elif sh1.row(rx)[17].value.startswith('STROBE'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type= 2
            
            writer.writerow(line(mnemonic,node,offset,bit,output_type,Description = description))

        elif sh1.row(rx)[17].value.startswith('CANRECRUN'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type= 2
            
            writer.writerow(line(mnemonic,node,offset,bit,output_type,Description = description))

        elif sh1.row(rx)[17].value.startswith('IOM_PEXCEPT_ACK'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type= 2
            
            writer.writerow(line(mnemonic,node,offset,bit,output_type,Description = description))           



    #1CTRLDRV
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            

            mnemonic = sh1.row(rx)[17].value
            _,_,System,_,belt_number = mnemonic.split('_')

            ctrldvr= "1CTRLDRV_" + System + "_" +belt_number

            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type=4

            print(line(mnemonic,node,offset,bit,1))
            writer.writerow(line(ctrldvr,node,offset,bit,output_type))
            #print(node + " " + offset + " " + bit)
    
    #ENMO0
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            

            mnemonic = sh1.row(rx)[17].value
            _,_,System,_,belt_number = mnemonic.split('_')

            ctrldvr= "ENMO0_" + System + "_" +belt_number

            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type=2

            print(line(mnemonic,node,offset,bit,1))
            writer.writerow(line(ctrldvr,node,offset,bit,output_type))
            #print(node + " " + offset + " " + bit)

    #MANA
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            

            mnemonic = sh1.row(rx)[17].value
            _,_,System,_,belt_number = mnemonic.split('_')

            mana= "MANA_" + System + "_" +belt_number

            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type=4

            print(line(mnemonic,node,offset,bit,1))
            writer.writerow(line(mana,node,offset,bit,output_type))
            #print(node + " " + offset + " " + bit)

    #RSMO
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            

            mnemonic = sh1.row(rx)[17].value
            _,_,System,_,belt_number = mnemonic.split('_')

            rsmo= "RSMO_" + System + "_" +belt_number

            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            output_type=2

            print(line(mnemonic,node,offset,bit,1))
            writer.writerow(line(rsmo,node,offset,bit,output_type))
            #print(node + " " + offset + " " + bit)
    
    for chute in range(1,num_chute+1):
        
        writer.writerow(line("CHUTE_CANREC_1_00" + str(chute) ,8,14,0,1))
        writer.writerow(line("CHUTE_FULL_1_00" + str(chute) ,8,14,0,1))
        writer.writerow(line("CHUTE_UNLOAD_NOK_1_00" + str(chute) ,8,14,0,1))
        writer.writerow(line("CHUTE_UNLOAD_OK_1_00" + str(chute) ,8,14,0,1))


book.release_resources()

out.close()
print(header_list)