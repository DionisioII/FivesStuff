from tkinter import Tk
from tkinter.filedialog import askopenfilename,askdirectory
import xlwings as xw
import csv


PHTs = {} # PHTs[system][phtNum]
BELTs = {}  #BELTs[system][beltNum]

barriers = {} # barriers[system][phtNum] ------> here we put ID of photocells which are barriers

phts_belt_counter ={}

###### ASK FOR PRODUCTION FOLDER ########

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
directory_of_production = askdirectory(title="select PRODUCTION folder") # show an "Open" dialog box and return the path to the selected file
print(directory_of_production)

input_csv = directory_of_production + "\ConveyorSys\SysGroup01\Database\DB_IOs\Inputs.csv"
output_csv = directory_of_production + "\ConveyorSys\SysGroup01\Database\DB_IOs\Outputs.csv"
sensor_csv = directory_of_production + "\ConveyorSys\SysGroup01\Database\DB_Params\Sensor.csv"

###### FINISH ASK FOR PRODUCTION FOLDER ########

###### OPEN sensor csv to get list of PHTs ########

with  open (sensor_csv,'r') as inp:
    reader = csv.reader(inp,delimiter=";", dialect="excel")
    for row in reader:
        if row[2] == "PREV_BELT_ID":

            system = row[3]
            phtNum = row[5]
            belt = row[7]
            
            if system != "0" and phtNum != 0:
                if system not in PHTs.keys():
                    PHTs[system] = {}
                PHTs[system][phtNum] = belt
inp.close()
print(PHTs)
###### END OF OPEN sensor csv to get list of PHTs ########


#Ask for current SWP file

Swp = askopenfilename(title="Select SWP")
print(Swp)

###### PARSE SWP TO GET LIST OF BELTS #############
wb = xw.Book(Swp)

upstream_systems_tabs = {1:"UP07_BELT",2:"UP06_BELT"} # name of the swp tab for each system

for system in upstream_systems_tabs.keys():

    if system not in BELTs.keys():
        BELTs[system] = {}


    sht = wb.sheets[upstream_systems_tabs[system]]
    print(sht.range('A15').value)
    rownum = sht.range('A15').current_region.last_cell.row + 1 #lenght of document(num of rows) on the bit column

    for i in range(17,rownum,166):

            BeltNum = int(sht.range('C'+str(i)).value) + 1
            BeltName = sht.range('H'+str(i)).value + sht.range('I'+str(i)).value
            BeltName = BeltName[5:]
            
            
            BELTs[system][BeltNum] = BeltName
print(BELTs)
###### END OF PARSE SWP TO GET LIST OF BELTS #############

###### PARSE INPUT CSV TO GET BARRIERS ########
            
with  open (input_csv,'r') as inp:
    reader = csv.reader(inp,delimiter=";", dialect="excel")
    for row in reader:
        if row[1][0:3] == "PHT" and row[1][-1] == "5":
            _,system,phtNum,_ = row.split("_")
           
            
            if system not in barriers.keys():
                barriers[system] = []
            barriers[system].append(phtNum)
inp.close()
#print(barriers)
###### END OF PARSE INPUT CSV TO GET BARRIERS ########


out_wb = xw.Book()
out_wb.sheets.add('pht')
out_sheet= out_wb.sheets['pht']

out_sheet.range('A2:A2').value = "MachineNum int"
out_sheet.range('B2:B2').value = "SenderMsgIDint"
out_sheet.range('C2:C2').value = "SenderMsgWhoIDint"
out_sheet.range('D2:D2').value = "WhoID int"
out_sheet.range('E2:E2').value = "SensorsWho 50 char"

i = 3
for system in sorted(PHTs.keys()):
        for key in sorted(PHTs[system].keys() ):
            GroupId = 1
            machine_num = 2
            SenderMsgIDint = 6
            SenderMsgWhoIDint = 10000+(GroupId*1000) + (int(system) *100)

            out_sheet.range('A'+str(i)+':A'+str(i)).value = machine_num
            out_sheet.range('B'+str(i)+':B'+str(i)).value = SenderMsgIDint
            out_sheet.range('C'+str(i)+':C'+str(i)).value = SenderMsgWhoIDint
            out_sheet.range('D'+str(i)+':D'+str(i)).value = key
            belt_num =int(PHTs[system][key])
            
            pht = BELTs[int(system)][belt_num]
            if pht not in phts_belt_counter.keys():
                phts_belt_counter[pht] = 1
            else:
                phts_belt_counter[pht] += 1

            out_sheet.range('E'+str(i)+':E'+str(i)).value = pht + "-B" + str(phts_belt_counter[pht])
            i+=1

i = 3
out_wb.sheets.add('belt')
out_sheet= out_wb.sheets['belt']
for system in BELTs:
        for key in sorted(BELTs[system].keys() ):
            GroupId = 1
            machine_num = 2
            SenderMsgIDint = 6
            SenderMsgWhoIDint = 10000+(GroupId*1000) + (int(system) *100)

            out_sheet.range('B'+str(i)+':B'+str(i)).value = SenderMsgIDint
            out_sheet.range('C'+str(i)+':C'+str(i)).value = SenderMsgWhoIDint
            out_sheet.range('D'+str(i)+':D'+str(i)).value = key
            BeltName = BELTs[system][key] 
            if "RB" in BeltName:
                BeltName = BeltName + "-M1"
            elif BeltName[-2:] == "-2":
                BeltName = BeltName + "-T2-H2"
            else:
                BeltName = BeltName + "-T1-H2"
            out_sheet.range('E'+str(i)+':E'+str(i)).value = BeltName
            i+=1



   
