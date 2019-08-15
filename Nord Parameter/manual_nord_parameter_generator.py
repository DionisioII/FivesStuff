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
head = ['@','2000',""]
header = "#,Card ID,Node ID,Parameter ID,Parameter subindex,Parameter set,Parameter value,Write to EEPROM,# Comment"
header_list = header.split(',' )


output_csv= "t_nord.csv"

UPstream = "UPSTREAM 103 "
hilscher_card = 2
""" card 2 up108
# UC100
#dopo 44: UC103
t90_3_4_withEncoder =[31,33,34,38,39,40,41,43,44,64,65,32,60,51,52,53,54,55,57]
t90_3_4_noEncoder=[36,37]
t90_1_4_withEncoder = [30,59]
t90_1_4_noEncoder = []

t80_3_4_withEncoder =[]
t80_3_4_noEncoder =[]
t80_1_4_withEncoder=[35,50,61]
t80_1_4_noEncoder = [66,63]

t100_3_4_withEncoder =[]
t100_3_4_noEncoder =[]
t100_1_4_withEncoder=[]
t100_1_4_noEncoder = []

t100_2_4_withEncoder = [56,58]


rulliere = [66]
"""
"""#101
#dopo 42: 102
#dopo 62: DC 102
#dopo 69-75: DC 101
#dopo 90 :DC 103
# dopo 104  : DC 104
t90_3_4_withEncoder =[30,32,33,35,40,41,50,52,53,55,60,61,80,81,82,83,85,86,87,70,71,72,73,74,75,90,91,92,93,94,95,98,105,106,107,108,110,112]
t90_3_4_noEncoder=[]
t90_1_4_withEncoder = [84,99,109,113]
t90_1_4_noEncoder = []

t80_3_4_withEncoder =[]
t80_3_4_noEncoder =[]
t80_1_4_withEncoder=[34,54]
t80_1_4_noEncoder = [39,42,59,62,96]

t100_3_4_withEncoder =[]
t100_3_4_noEncoder =[]
t100_1_4_withEncoder=[]
t100_1_4_noEncoder = []

t100_2_4_withEncoder = []
rulliere =[42,62]"""
"""
UPSTREAM 104
t90_3_4_withEncoder =[35,37,38,41,42,46,47,48,49,50,53,55,57,59,60,61,62,35,65,66,67,68,69,80,81,82,84,85,30,31,32,33]
t90_3_4_noEncoder=[44]
t90_1_4_withEncoder = [40,64]
t90_1_4_noEncoder = [45]

t80_3_4_withEncoder =[]
t80_3_4_noEncoder =[]
t80_1_4_withEncoder=[34,36,43,]
t80_1_4_noEncoder = [51,52,54,58,83]

t100_3_4_withEncoder =[]
t100_3_4_noEncoder =[]
t100_1_4_withEncoder=[]
t100_1_4_noEncoder = []

t100_2_4_withEncoder = []
rulliere =[85]
"""
"""up 105
t90_3_4_withEncoder =[30,31,32,33,35,36,39,41,44,45,46,47,48,49,51,53,55,57,59,60,61,63,64,65,66,70,71,72,74,75]
t90_3_4_noEncoder=[42]
t90_1_4_withEncoder = [40]
t90_1_4_noEncoder = [43,50]

t80_3_4_withEncoder =[]
t80_3_4_noEncoder =[]
t80_1_4_withEncoder=[34,37,58,62]
t80_1_4_noEncoder = [38,52,56,73]

t100_3_4_withEncoder =[]
t100_3_4_noEncoder =[]
t100_1_4_withEncoder=[]
t100_1_4_noEncoder = []

t100_2_4_withEncoder = []
rulliere =[76]
"""
"""PUSTREAM 102
t90_3_4_withEncoder =[32,37,38,39,40,41,43,45,47,50,51,52,53,60,61,62,63,6570,71,72,73,75]
t90_3_4_noEncoder=[31,33,35]
t90_1_4_withEncoder = [30,34,48]
t90_1_4_noEncoder = [36,42,46,49,64,66,74,76]

t80_3_4_withEncoder =[]
t80_3_4_noEncoder =[]
t80_1_4_withEncoder=[]
t80_1_4_noEncoder = []

t100_3_4_withEncoder =[]
t100_3_4_noEncoder =[]
t100_1_4_withEncoder=[]
t100_1_4_noEncoder = []

t100_2_4_withEncoder = []
rulliere =[75]
"""
#upstream 103
t90_3_4_withEncoder =[31,32,37,38,39,40,41,42,44,46,48,50,51,53,54,55,60,61,62,64,70,71,72,74]
t90_3_4_noEncoder=[35]
t90_1_4_withEncoder = [30,34]
t90_1_4_noEncoder = [36]

t80_3_4_withEncoder =[]
t80_3_4_noEncoder =[]
t80_1_4_withEncoder=[49]
t80_1_4_noEncoder = [43,47,63,65,73,75]

t100_3_4_withEncoder =[]
t100_3_4_noEncoder =[]
t100_1_4_withEncoder=[]
t100_1_4_noEncoder = []

t100_2_4_withEncoder = [52]
t100_2_4_noEncoder = [33]
rulliere =[75]

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

with  open(output_csv, 'w', newline='') as out:
    writer = csv.writer(out, delimiter=",", dialect="excel")
    #writer.writerow(head)
    out.write('#,500,,,,,,,\n')
    writer.writerow(header_list)
    out.write('# ### '+ UPstream+' ###\n')

    for nodo in range(0,200):

        
        
        if nodo in t90_3_4_withEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 90T3/4, encoder = yes \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,105,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
            #writer.writerow(['@',hilscher_card,nodo,301,0,0,5,1,'#'])
            all_motors_with_encoder(writer)
            all_motors(writer)
            

        if nodo in t90_3_4_noEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 90T3/4, encoder = no \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,105,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,0,1,'#'])
            all_motors(writer)

        if nodo in t90_1_4_withEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 90T1/4, encoder = yes \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,101,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
            #writer.writerow(['@',hilscher_card,nodo,301,0,0,5,1,'#'])
            all_motors_with_encoder(writer)
            all_motors(writer)
        
        if nodo in t90_1_4_noEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 90T1/4, encoder = no \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,101,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,0,1,'#'])
           
            all_motors(writer)
        
        if nodo in t80_1_4_withEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 80T1/4, encoder = yes \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,98,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
            #writer.writerow(['@',hilscher_card,nodo,301,0,0,5,1,'#'])
            all_motors_with_encoder(writer)
            all_motors(writer)
            if nodo in rulliere:
                rulliera(writer)
        
        if nodo in t80_1_4_noEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 80T1/4, encoder = no \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,98,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,0,1,'#'])
           
            all_motors(writer)
            if nodo in rulliere:
                rulliera(writer)
        
        if nodo in t80_3_4_withEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 80T3/4, encoder = yes \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,98,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
            #writer.writerow(['@',hilscher_card,nodo,301,0,0,5,1,'#'])
            all_motors_with_encoder(writer)
            all_motors(writer)
            if nodo in rulliere:
                rulliera(writer)
        
        if nodo in t80_3_4_noEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 80T3/4, encoder = no \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,98,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,0,1,'#'])
           
            all_motors(writer)
        
        if nodo in t100_2_4_withEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 100T2/4, encoder = yes \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,109,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
            all_motors_with_encoder(writer)
            all_motors(writer)

        if nodo in t100_2_4_noEncoder:
            node_number = "{0:0=4d}".format(nodo)
            out.write('#n'+ node_number +' motor type = 100T2/4, encoder = yes \n')
            writer.writerow(['@',hilscher_card,nodo,200,0,1,109,1,'#'])
            writer.writerow(['@',hilscher_card,nodo,300,0,1,1,1,'#'])
            #all_motors_with_encoder(writer)
            all_motors(writer)
            
out.close()
print(header_list)