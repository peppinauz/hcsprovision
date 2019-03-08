## LINUX#!/usr/bin/python3
## OSX
#!/usr/local/bin/python3

import csv
import time
import sys
import json
import configparser

#INPUTS: SLC
siteslc="0000"
debug=1

if len(sys.argv) != 6:
    print("ERROR: python3 <nombre-fichero.py> <static-data-file> <site-data-file> XXXX <ClusterPath> <Config Log File>")
    exit(1)
else:
    fmostaticdata=sys.argv[1]
    fmositedata=sys.argv[2]
    siteslc=str(sys.argv[3])   #INPUTS: SLC
    clusterpath=sys.argv[4]    #SLC Cluster path
    logconfigfile=sys.argv[5]   ## LOG config file

cl=clusterpath[14:16]       ## CL = dos digitos, 01, 02, 03,...
clusterid="cl"+cl

## LOG File
f = open(logconfigfile, 'a')

print("(II) #################################################", file=f)
print("(II) #################################################", file=f)
print("(II) INPUT CMO configuration      ::",fmositedata, file=f)
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

with open(fmositedata,'r') as infile:
    data = json.load(infile)
infile.close()

#FMO working path:
sitepath="../FMO/"+siteslc

## LOG File
f = open(logconfigfile, 'a')

## OUTPUT FILE
cmoresetphone = open(sitepath+"/cmo-reset."+siteslc+".csv", 'w')
fmoresetphone = open(sitepath+"/fmo-reset."+siteslc+".csv", 'w')

print("entorno,clusterid,sub1,sub2,ip-phone-model,ip-phone-name,ip-phone-user,ip-phone-pass,ip-phone-web-access-enable,ip-phone-ssh-user,ip-phone-ssh-pass,deviceprofile,userid,ipaddress,registered", file=cmoresetphone)
print("entorno,clusterid,sub1,sub2,ip-phone-model,ip-phone-name,ip-phone-user,ip-phone-pass,ip-phone-web-access-enable,ip-phone-ssh-user,ip-phone-ssh-pass,deviceprofile,userid,ipaddress,registered", file=fmoresetphone)

## FMO CUSTOMER INPUT DATA
fmositeid=data['fmosite'][0]['id']
cmg=data['fmosite'][0]['cmg']

# CMO patterns
cmodevicepool=siteslc+"-DP"
cmolocation=siteslc+"-LOC"
cmoslc="5"+siteslc

## FMO UserData
cucdmsite=fmoenvconfig['fmocustomerid']+"Si"+str(fmositeid)
fmopass="vi123456"
fmopin="123456"

## fmo CMG
## IP PHONE API
fmocmg = configparser.ConfigParser()
fmocmg.read('doc/fmocmgroup.ini')


## migracion.phone.py: variables
cmocmg=""       ## CMGroup
cmocm=[]        ## CM members
cmocmserver=[]  ## CM server name/IP

# CMO File INPUT DATA: DevicePool
inputfile = clusterpath+"/devicepool.csv"
fin = open(inputfile,"r")
csv_f = csv.DictReader(fin)

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        if row['DEVICE POOL NAME'] == cmodevicepool:
            if debug:
                print("(DD) Device Pool:"+row['DEVICE POOL NAME']+", CMG:"+row['CISCO UNIFIED CALLMANAGER GROUP'], file=f)
            cmocmg=row['CISCO UNIFIED CALLMANAGER GROUP']
fin.close()

if cmocmg == "": ## No hemos encontrado el DevicePool
    print("(EE) Device Pool NO encontrado")
    exit(1)

# CMO File INPUT DATA: CUCMGroup servers
inputfile = clusterpath+"/callmanagergroup.csv"   ## Modificado
fin = open(inputfile,"r")
csv_f = csv.DictReader(fin)

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        if row['CALL MANAGER GROUP NAME'] == cmocmg:
            if "CALL MANAGER 3" in csv_f.fieldnames:
                if debug:
                    print("(DD) CMG:"+row['CALL MANAGER GROUP NAME']+", CM1:"+row['CALL MANAGER 1']+", CM2:"+row['CALL MANAGER 2']+", CM3:"+row['CALL MANAGER 3'], file=f)
                cmocm={row['CALL MANAGER 1'],row['CALL MANAGER 2'],row['CALL MANAGER 3']}
            else:
                if debug:
                    print("(DD) CMG:"+row['CALL MANAGER GROUP NAME']+", CM1:"+row['CALL MANAGER 1']+", CM2:"+row['CALL MANAGER 2'], file=f)
                cmocm={row['CALL MANAGER 1'],row['CALL MANAGER 2']}
fin.close()

if len(cmocm)==0:
    print("(EE) Call Manager NO encontrado", file=f)
    exit(1)

# CMO File INPUT DATA: CUCM servers
inputfile = clusterpath+"/callmanager.csv"   ## Modificado
fin = open(inputfile,"r")
csv_f = csv.DictReader(fin)

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        for cmname in cmocm:
            if row['CISCO UNIFIED CALLMANAGER NAME'] == cmname:
                if debug:
                    print("(DD) CM Name:"+cmname+", SERVER/IP:"+row['SERVER'], file=f)
                cmelement={'cmname':cmname,'cmserver':row['SERVER']}
                cmocmserver.append(cmelement)
fin.close()

if len(cmocmserver)==0:
    print("(EE) Call Manager NO encontrado")
    exit(1)

# CMO File INPUT DATA: PHONES
inputfile = clusterpath+"/phone.mod1.csv"   ## Modificado
fin = open(inputfile,"r")
csv_f = csv.DictReader(fin)

# CHECK: fieldname exists!!
if "Owner User ID" in csv_f.fieldnames:
    userid=1
else:
    userid=0

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        if row['Device Pool'].startswith(cmodevicepool) and (row['Device Name'].startswith('SEP') or row['Device Name'].startswith('ATA')):
            ## XML field
            ## webAccess
            webaccess="disable"
            #if row['XML'].split("<webAccess>")[1].split("</webAccess>")[0] == "0": ## LOGICA Inversa 0=ENABLE/1=DISABLE
            #    webaccess="enable"
            #else:
            #    webaccess="disable"

            ## DEBUG: FMO
            print("FMO,agencias-cl2,"+fmocmg[cmg]['primary']+","+fmocmg[cmg]['secondary']+","+row['Device Type']+","+row['Device Name']+","+"p"+row['Directory Number 1']+","+fmopass,",enable,,", file=fmoresetphone)
            #print("FMO,,"+fmocmg[cmg]['primary']+","+fmocmg[cmg]['secondary']+","+","+row['Device Type']+","+row['Device Name']+","+"p"+row['Directory Number 1']+","+fmopass,",enable")

            ## YAML??

            ## DEBUG: CMO
            if not userid:
                print("CMO,"+clusterid+","+cmocmserver[0]['cmserver']+","+cmocmserver[1]['cmserver']+","+row['Device Type']+","+row['Device Name']+",,,"+webaccess+","+row['Secure Shell User']+","+row['Secure Shell Password'], file=cmoresetphone)
            else:
                print("CMO,"+clusterid+","+cmocmserver[0]['cmserver']+","+cmocmserver[1]['cmserver']+","+row['Device Type']+","+row['Device Name']+","+row['Owner User ID']+","+fmopass+","+webaccess+","+row['Secure Shell User']+","+row['Secure Shell Password'], file=cmoresetphone)
            ## YAML??

fin.close()
fmoresetphone.close()
cmoresetphone.close()

# CMO File INPUT DATA: ENDUSER (Buscamos la pass)
#fin = open(inputfile,"r")
#csv_f = csv.DictReader(fin)
#inputfile = clusterpath+"/enduser.csv"   ## Modificado

#fin.close()

exit(0)
