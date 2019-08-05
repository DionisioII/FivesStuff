import xlrd
import csv


#def line(Input_Name, Input_Type, NC_NO_Logic = 0, Rising_edge_delay=0,falling_edge_delay = 0, hold_on_delay=0,hold_off_delay=0,events=0,destination_queue=0,enabling_with_srter_running=0,Node,offset,bit, default_value=0,board_number=1,Description=''):
def line(Input_Name,Node,offset,bit, Input_Type, NC_NO_Logic = 0, Rising_edge_delay=0,falling_edge_delay = 0, hold_on_delay=0,hold_off_delay=0,events=0,destination_queue=0,enabling_with_sorter_running=0, default_value=0,board_number=1,Description=''):
    return  ['@', Input_Name,	Input_Type,	NC_NO_Logic,	Rising_edge_delay,	falling_edge_delay,	hold_on_delay,	hold_off_delay,	events,	destination_queue,  enabling_with_sorter_running,	Node,	offset,	bit,	default_value,	board_number,Description]

def get_profinet_address(IO_address):
    values = IO_address.split('.')
    Node = values[1][-3:]
    if Node[0] == '0' :
        Node = Node[-2:]
    offset = values[2]
    bit = values[3]
    #print(IO_address[2])
    return Node,offset,bit

num_chute = 0
head = ['@','300']
header = "#	Input Name	Input Type	NC/NO Logic	Rising Edge Delay	Falling Edge Delay	Hold ON Delay	Hold OFF Delay	Events	Destination Queue	Enabling With Sorter Running	Node	Offset	Bit	Default Value (For Virtual I/Os)	BOARD NUMBER	Description"
header_list = header.split('\t' )

output_csv= "input.csv"


book = xlrd.open_workbook('8FOS_IO013_V5.1.xls')

sh1 = book.sheet_by_index(1)


with  open(output_csv, 'w', newline='') as out:
    writer = csv.writer(out, delimiter=";", dialect="excel")
    writer.writerow(head)
    writer.writerow(header_list)




    #MAIN_JFOR
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            mnemonic = sh1.row(rx)[17].value
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            print(line(mnemonic,node,offset,bit,1))
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
            #print(node + " " + offset + " " + bit)

    #MAIN_JREV
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JREV'):
            mnemonic = sh1.row(rx)[17].value
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(mnemonic,node,offset,bit,1))
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
    
    #MOTCUR
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            mnemonic = sh1.row(rx)[17].value
            _,_,System,_,belt_number = mnemonic.split('_')

            motcur= "MOTCUR_" + System + "_" +belt_number
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(motcur,node,offset,bit,1))
            writer.writerow(line(motcur,node,offset,bit,3))
    
    #MOTORPOWER
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            mnemonic = sh1.row(rx)[17].value
            _,_,System,_,belt_number = mnemonic.split('_')

            motcur= "MOTORPOWER_" + System + "_" +belt_number
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(motcur,node,offset,bit,1))
            writer.writerow(line(motcur,node,offset,bit,1))
    
    #MSTATE
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('MAIN_JFOR'):
            mnemonic = sh1.row(rx)[17].value
            _,_,System,_,belt_number = mnemonic.split('_')

            mstate= "MSTATE_" + System + "_" +belt_number
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(motcur,node,offset,bit,1))
            writer.writerow(line(mstate,node,offset,bit,3))
    
    #ALLQ,WARQ
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('ALLQ') or sh1.row(rx)[17].value.startswith('WARQ') :
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(motcur,node,offset,bit,1))
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
    
    #ALLB,WARB
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('ALLB') or sh1.row(rx)[17].value.startswith('WARB') :
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(motcur,node,offset,bit,1))
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
    
    #JOGFOR
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('JOGFOR')  :
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(motcur,node,offset,bit,1))
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
    
    #START,STOP
    for rx in range(1, sh1.nrows):


        if sh1.row(rx)[17].value.startswith('START'):
                    
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))

        elif sh1.row(rx)[17].value.startswith('STOP'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description,NC_NO_Logic=1))

        elif sh1.row(rx)[17].value.startswith('RESET'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))

        elif sh1.row(rx)[17].value.startswith('LAMP_TEST'):
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))

        elif sh1.row(rx)[17].value.startswith('PEXCEPT_ ACK'):

            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
        
        elif sh1.row(rx)[17].value.startswith('PEXCEPT_ FWD'):

            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
        
        
        
      
        
        elif sh1.row(rx)[17].value.startswith('STROBE'):

            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
        
        elif sh1.row(rx)[17].value.startswith('OTK_LOADING'):

            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))

        elif sh1.row(rx)[17].value.startswith('MAIN_KEY'):

            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))

    #PHT
    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[17].value.startswith('PHT') :
            mnemonic = sh1.row(rx)[17].value
            
            description = sh1.row(rx)[9].value
            profinet_address = sh1.row(rx)[4].value
            node,offset,bit= get_profinet_address(profinet_address)
            
            #print(line(motcur,node,offset,bit,1))
            writer.writerow(line(mnemonic,node,offset,bit,1,Description = description))
            if description.startswith('BARRIERA'):
                for x in range(1,31):
                    if x< 10:
                        writer.writerow(line(mnemonic[:-3]+'00'+str(x),node,offset,bit,1,Description = description))
                    else:
                        writer.writerow(line(mnemonic[:-3]+'0'+str(x),node,offset,bit,1,Description = description))
    #CHUTE

    for chute in range(1,num_chute+1):
        
        writer.writerow(line("CHUTE_EXDATA_1_00" + str(chute) +"_0",8,14,0,1))
        writer.writerow(line("CHUTE_EXDATA_1_00" + str(chute) +"_4",8,14,0,1))
        writer.writerow(line("CHUTE_TRIGGER_1_00" + str(chute) ,8,14,0,1))
        writer.writerow(line("CHUTE_UNLOAD_REQ_1_00" + str(chute) ,8,14,0,1))
        






book.release_resources()

out.close()
print(header_list)