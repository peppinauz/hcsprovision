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
print("(II) INPUT CMO configuration      ::",fmositedata, file=f)
print("(II) INPUT datos de entorno       ::",fmostaticdata, file=f)
print("(II) INPUT SLC                    :: ",siteslc, file=f)
print("(II) INPUT LOG configuration file :: ",logconfigfile, file=f)
print("(II) INPUT Cluster Path           :: ",clusterpath, file=f)

cl=clusterpath[14:16]       ## CL = dos digitos, 01, 02, 03,...

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

with open(fmositedata,'r') as infile:
    data = json.load(infile) 
infile.close()

#FMO working path:
sitepath="../FMO/"+siteslc

# FILES
templateblkfile = "../code/blk/90.siptest-template.xlsx" # SIN DATAINPUT
outputblkfile = sitepath+"/90.siptest."+siteslc+".xlsx"

## FMO CUSTOMER INPUT DATA
hierarchynode=fmoenvconfig['hierarchynode'] 
customerid=fmoenvconfig['fmocustomerid'] 
aargroup=fmoenvconfig['fmoaargroup']
fmoserviceurl=fmoenvconfig['fmoservice1url']
fmoservicename=fmoenvconfig['fmoservice1name']
##
fmositename=data['fmosite'][0]['name']
fmositeid=data['fmosite'][0]['id']  
cmg=data['fmosite'][0]['cmg']  

## FMO UserData
cucdmsite=fmoenvconfig['fmocustomerid']+"Si"+str(fmositeid)
cssfwd=customerid+"-DirNum-CSS"
linept=customerid+"-DirNum-PT"
linecss=cucdmsite+"-DBREnhIntl24HrsCLIPyFONnFACnCMC-CSS"
aarcss=customerid+"-AAR-CSS"
devicepool=cucdmsite+"-DevicePool"
location=cucdmsite+"-Location"
devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
subscribecss=cucdmsite+"-InternalOnly-CSS"
fwdnocss=customerid+"-DirnumEM-CSS"
#fwdallcss=cucdmsite+"-InternalOnly-CSS"
fwdallcss=cucdmsite+"-DBREnhIntl24HrsCLIPyFONnFACnCMC-CSS"

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

fila=5

# INPUT DATA
devicenumber="5"+siteslc+data['sipptest'][0]['extension']
devicename="SEP999"+devicenumber

print("(II) SIPP::",devicename,"(",devicenumber,") :: Ejecutar test para Area Code=",data['e164'][0]['ac'],file=f)
print("(II) SIPP::",devicename,"(",devicenumber,") :: Ejecutar test para Area Code=",data['e164'][0]['ac'])


sheet =  blk["QAS"]
sheet['B'+str(fila)]=hierarchynode+"."+fmositename 
sheet['C'+str(fila)]="add"
#sheet['D'+str(fila)]="name:"+row[3] # Search field
###################################################### 
sheet['i'+str(fila)]=devicenumber                               # firsname
sheet['j'+str(fila)]=devicenumber                               # lastname
sheet['k'+str(fila)]=devicenumber                               # username
#sheet['l'+str(fila)]="borrame@telefonica.com"                   # email
#sheet['m'+str(fila)]=""                                        # DeviceName (Info) Jabber Max 12 Char
sheet['n'+str(fila)]=devicenumber                               # password
sheet['o'+str(fila)]=devicenumber                               # pin
sheet['p'+str(fila)]=devicenumber                               # lines.0.directory_number
sheet['q'+str(fila)]="ProdubanSIPPTest"                         # qagroup_name
#sheet['r'+str(fila)]=""                                        # Entitlement profile
sheet['s'+str(fila)]="true"                                     # voice
#sheet['t'+str(fila)]=""                                        # phone_type
sheet['u'+str(fila)]=devicename                                 # phones.0.phone_name 
sheet['v'+str(fila)]="false"                                    # mobility
sheet['w'+str(fila)]="false"                                    # voicemail
sheet['x'+str(fila)]="false"                                    # snr
#sheet['y'+str(fila)]=""                                        # mobile_number
#sheet['z'+str(fila)]=""                                        # webex
sheet['aa'+str(fila)]="false"                                   # jabber
sheet['ak'+str(fila)]=data['e164'][0]['head']                                          # e164

sheet =  blk["PHONE.3rdparty"]
sheet['B'+str(fila)]=hierarchynode+"."+fmositename 
sheet['C'+str(fila)]="modify"
sheet['D'+str(fila)]="name:"+devicename                         # Search field
###################################################### 
sheet['Z'+str(fila)]=devicename                                 # name
sheet['BT'+str(fila)]=devicenumber                              # digest

## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)

###################################################### 
###################################################### 
###################################################### 
###################################################### 
templateblkfile = "../code/blk/90.delete-siptest-template.xlsx" 
outputblkfile = sitepath+"/90.delete-siptest."+siteslc+".xlsx"

blk = openpyxl.load_workbook(templateblkfile)

print("(II) Remove::",devicename,"(",devicenumber,")" ,file=f)

sheet =  blk["PHONE.RM"]
sheet['B'+str(fila)]=hierarchynode+"."+fmositename 
sheet['C'+str(fila)]="delete"
sheet['D'+str(fila)]="name:"+devicename                         # Search field
###################################################### 
sheet['Z'+str(fila)]=devicename                                 # name

sheet =  blk["LINE.RM"]
sheet['B'+str(fila)]=hierarchynode+"."+fmositename 
sheet['C'+str(fila)]="delete"
sheet['D'+str(fila)]="pattern:"+devicenumber+",routePartitionName:"+linept                         # Search field
###################################################### 
sheet['Z'+str(fila)]=devicenumber                                # name
sheet['Q'+str(fila)]=linept                                      # Partition


sheet =  blk["SUB.RM"]
sheet['B'+str(fila)]=hierarchynode+"."+fmositename 
sheet['C'+str(fila)]="delete"
sheet['D'+str(fila)]="userid:"+devicenumber                         # Search field
###################################################### 
sheet['GN'+str(fila)]=devicenumber                                 # name

## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)

## LOG de CONFIGURACION
f.close()

exit(0)