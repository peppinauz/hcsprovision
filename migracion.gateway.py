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

#cl=clusterpath[14:16]       ## CL = dos digitos, 01, 02, 03,...
#clusterid="cl"+cl

## LOG File
f = open(logconfigfile, 'a')

print("(II) #################################################", file=f)
print("(II) Migracion GW ", file=f)
print("(II) #################################################", file=f)
print("(II) INPUT CMO configuration      ::",fmositedata, file=f)
print("(II) INPUT datos de entorno       ::",fmostaticdata, file=f)
print("(II) INPUT SLC                    :: ",siteslc, file=f)
print("(II) INPUT LOG configuration file :: ",logconfigfile, file=f)
print("(II) INPUT Cluster Path           :: ",clusterpath, file=f)

f.close()

# CMO datos del site
# Datos extraidos en filtra.dp
data = {}

with open(fmositedata,'r') as infile:
    data = json.load(infile)
infile.close()

#FMO working path:
sitepath="../FMO/"

## BUSCAMOS si ya hemos escrito el SLC:
ingateway = open(sitepath+"/dataInfo.csv",'r')
lines=ingateway.readlines()
ingateway.close()

newlines=[]

for l in lines:
    if not l.startswith(siteslc):
        newlines.append(l)

ingateway = open(sitepath+"/dataInfo.csv",'w')
ingateway.writelines(newlines)
ingateway.close()

## GENERAMOS OUTPUT FILE
csvgateway = open(sitepath+"/dataInfo.csv", 'a')

## FMO CUSTOMER INPUT DATA
cmg=data['fmosite']['cmg']
ipsip=data['srst'][0]['ipsccp']
areacode=data['e164'][0]['ac']

## fmo CMG
fmocmg = configparser.ConfigParser()
fmocmg.read('doc/fmocmgroup.ini')

## OUTPUT FILE
csvgateway = open(sitepath+"/dataInfo.csv", 'a')

## DEBUG: FMO
#print("id,areaCode,ipGerencia,hostname,sub1,sub2", file=csvgateway)
print(str(siteslc),",",areacode,",ipGerencia,hostname,",fmocmg[cmg]['primary'],",",fmocmg[cmg]['secondary'],",",ipsip, file=csvgateway)

csvgateway.close()

exit(0)
