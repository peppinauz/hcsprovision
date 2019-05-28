## LINUX#!/usr/bin/python3
## OSX
#!/usr/local/bin/python3

import csv
import openpyxl
import time
import sys
import json

#INPUTS: SLC
siteslc="0000"

if len(sys.argv) != 6:
    print("ERROR: python3 <nombre-fichero.py> <static-data-file> <site-data-file> XXXX <ClusterPath> <Config Log File>")
    exit(1)
else:
    fmostaticdata=sys.argv[1]
    fmositedata=sys.argv[2]
    siteslc=str(sys.argv[3])   #INPUTS: SLC
    clusterpath=sys.argv[4]    #SLC Cluster path
    logconfigfile=sys.argv[5]   ## LOG config file

## LOG File
f = open(logconfigfile, 'a')

print("(II) #################################################", file=f)
print("(II) #################################################", file=f)
print("(II) INPUT CMO configuration      ::",fmostaticdata, file=f)
print("(II) INPUT datos de entorno       ::",fmostaticdata, file=f)
print("(II) INPUT SLC                    :: ",siteslc, file=f)
print("(II) INPUT LOG configuration file :: ",logconfigfile, file=f)
print("(II) INPUT Cluster Path           :: ",clusterpath, file=f)

# FMO datos de entorno:
fmoenvconfig={}
with open(fmostaticdata, "r") as fp:
        for line in fp.readlines():
            li = line.lstrip()
            if not li.startswith("#") and '=' in li:
                key, value = line.split('=', 1)
                fmoenvconfig[key] = value.strip()  ## variable de tipo diccionario
                #print("<<<<   ",fmoenvconfig[key])
fp.close()

# CMO datos del site
data = {}
#data['dp'] = []
#data['loc']=[]
#data['mrgl'] = []
#data['mrg'] = []
#data['srst'] = []
#data['gw'] = []
#data['rsc'] = []
#data['e164'] = []
#data['agencia'] = []

#e164={}
#agencia={}
cpg=[]

with open(fmositedata,'r') as infile:
    data = json.load(infile)
infile.close()


#INPUT: E164
e164=data['e164'][0]['head']
fmositename=data['fmosite']['name']
fmositeid=data['fmosite']['id']
cmg=data['fmosite']['cmg']

#FMO working path:
sitepath="../FMO/"+siteslc

# CMO
inputcpgfile = clusterpath+"/callpickupgroup.csv"
templateblkfile = "blk/02.cpg-template.xlsx"
outputblkfile = sitepath+"/02.cpg."+siteslc+".xlsx"

## FMO CUSTOMER INPUT DATA
hierarchynode=fmoenvconfig['hierarchynode']
customerid=fmoenvconfig['fmocustomerid']
fmositename=data['fmosite']['name']
fmositeid=data['fmosite']['id']
cucdmsite=fmoenvconfig['fmocustomerid']+"Si"+str(fmositeid)
preisrpt=customerid+"-PreISR-PT"
preisrcss=customerid+"-PreISR-CSS"
isrpt=customerid+"-ISR-PT"
isrcss=customerid+"-ISR-CSS"
dirnumpt=customerid+"-DirNum-PT"
dirnumcss=customerid+"-DirNum-CSS"
featurept=cucdmsite+"-Feature-PT"

## CMO pattern
range=siteslc+"-"

# CMO File INPUT DATA
fcpg = open(inputcpgfile,"r")
csv_f = csv.DictReader(fcpg)

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

# FMO commands:
action="add"
print("(II) Configurando CPG",file=f)

fila=7
for row in csv_f:
    # WR BLK OUTPUT DATA
    cpgnumber = row ['CPG NAME']
    #if cpgnumber.startswith(range):
    if cpgnumber.startswith(range):
        sheet = blk["CPG"]
        sheet['B'+str(fila)]=hierarchynode+"."+fmositename
        sheet['C'+str(fila)]=action
        sheet['D'+str(fila)]="name:"+row['CPG NAME']
        sheet['H'+str(fila)]=row['CPG NAME']
        sheet['I'+str(fila)]=fmositename+" "+row['CPG NAME']
        sheet['J'+str(fila)]=row['CPG NUMBER']
        sheet['K'+str(fila)]="false"
        sheet['L'+str(fila)]="true"
        sheet['M'+str(fila)]=row['CPG NOTIFICATION POLICY'] ## pickupNotification
        sheet['O'+str(fila)]=row['CPG NOTIFICATION TIMER']   ## pickupNotificationTimer
        sheet['P'+str(fila)]=featurept
        # Next row

        # Guardamos los números a abrir en PreISR/ISR
        cpg.append(row['CPG NUMBER'])
        print("(II) CPG ",row['CPG NAME'],"-",row['CPG NUMBER'],file=f)

        fila=fila+1

## CMO File INPUT DATA: Close
fcpg.close()
## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)

## Guardo los Hunt pilot patterns
with open(fmositedata,'w') as outfile:
    #data = json.load(outfile)
    data['cpgnumber']=cpg
    json.dump(data,outfile)
outfile.close()

## LOG de CONFIGURACION
f.close()

exit(0)
