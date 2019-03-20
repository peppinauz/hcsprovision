## LINUX#!/usr/bin/python3
## OSX
#!/usr/local/bin/python3

#import csv
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

templateblkfile = "blk/99.ON-template.xlsx" # SIN DATAINPUT
outputblkfile = sitepath+"/99.on."+siteslc+".xlsx"

## FMO CUSTOMER INPUT DATA
hierarchynode=fmoenvconfig['hierarchynode']
customerid=fmoenvconfig['fmocustomerid']
aargroup=fmoenvconfig['fmoaargroup']
##
fmositename=data['fmosite'][0]['name']
fmositeid=data['fmosite'][0]['id']
cmg=data['fmosite'][0]['cmg']

# CMO patterns
cmodevicepool=siteslc+"-DP"

## FMO UserData
cucdmsite=fmoenvconfig['fmocustomerid']+"Si"+str(fmositeid)
cssfwd=customerid+"-DirNum-CSS"
linept=customerid+"-DirNum-PT"
emlinept=customerid+"-DirNumEM-PT"
linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
aarcss=customerid+"-AAR-CSS"
devicepool=cucdmsite+"-DevicePool"
location=cucdmsite+"-Location"
devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
subscribecss=cucdmsite+"-InternalOnly-CSS"
preisrpt=customerid+"-PreISR-PT"
preisrcss=customerid+"-PreISR-CSS"
isrpt=customerid+"-ISR-PT"
isrcss=customerid+"-ISR-CSS"
dirnumpt=customerid+"-DirNum-PT"
dirnumcss=customerid+"-DirNum-CSS"
nullpt=customerid+"-NoMigrado-PT"
siterange="5"+siteslc+"XXXX"

# CMO File INPUT DATA
#fgw = open(inputfilehp,"r")
#csv_f = csv.reader(fgw)

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

fila=7

print("(II): Site activation: ",siteslc,file=f)

######################################################
## Habilitamos el routing del site en  PreISR
sheet = blk["TPpreISR"]
sheet['B'+str(fila)]=hierarchynode+"."+fmositename
sheet['C'+str(fila)]="modify"
sheet['D'+str(fila)]="routePartitionName:"+nullpt+",pattern:"+siterange
sheet['I'+str(fila)]="Cisco CallManager"           # callingPartyNumberingPlan
sheet['j'+str(fila)]="Default"                     # connectedLinePresentationBit
sheet['k'+str(fila)]=preisrpt                      # routePartitionName
sheet['l'+str(fila)]="No Error"                    # releaseClause
sheet['m'+str(fila)]="false"                       # blockEnable
sheet['n'+str(fila)]="Cisco CallManager"           # callingPartyNumberType
sheet['o'+str(fila)]="false"                       # provideOutsideDialtone
sheet['q'+str(fila)]=siterange                        # pattern
sheet['r'+str(fila)]="Default"                     # patternPrecedence
sheet['s'+str(fila)]=isrcss                        # callingSearchSpaceName
sheet['t'+str(fila)]=""                            # prefixDigitsOut
sheet['u'+str(fila)]="Translation"                 # usage
sheet['v'+str(fila)]="Cisco CallManager"           # calledPartyNumberingPlan
sheet['w'+str(fila)]="true"                        # dontWaitForIDTOnSubsequentHops
sheet['x'+str(fila)]="Default"                     # connectedNamePresentationBit
sheet['y'+str(fila)]="PreISR Site "+siterange      # Description
sheet['z'+str(fila)]="Default"                     # routeClass
sheet['aa'+str(fila)]="Default"                    # callingNamePresentationBit
sheet['ac'+str(fila)]="false"                      # routeNextHopByCgpn
sheet['ad'+str(fila)]="false"                      # useOriginatorCss
sheet['ae'+str(fila)]=""                           # digitDiscardInstructionName
sheet['ag'+str(fila)]="Off"                        # useCallingPartyPhoneMask
sheet['aj'+str(fila)]="Cisco CallManager"          # calledPartyNumberType
sheet['al'+str(fila)]="false"                      # patternUrgency
sheet['am'+str(fila)]="Default"                    # callingLinePresentationBit

######################################################
## Habilitamos el routing del site en  ILS
sheet = blk["ILS"]
sheet['B'+str(fila)]=hierarchynode+"."+fmositename
sheet['C'+str(fila)]="add"
sheet['i'+str(fila)]="0"                           # pstnFailStrip
sheet['j'+str(fila)]=siterange                     # pattern
sheet['k'+str(fila)]="Specify"                     # hostedRoutePSTNRule
sheet['l'+str(fila)]="Enterprise Number"           # patternType
sheet['m'+str(fila)]=fmositename                   # Description

fila+=1

## Habilitamos el routing del site en  PreISR
sheet = blk["TPpreISR"]
sheet['A'+str(fila)]="##"
## Habilitamos el routing del site en  ILS
sheet = blk["ILS"]
sheet['A'+str(fila)]="##"

fila+=1

## Cargamos los datos del HPilot
for hp in data['huntpilot']:
    #print(hp)
    ######################################################
    ## Habilitamos el routing del site en  PreISR
    sheet = blk["TPpreISR"]
    sheet['B'+str(fila)]=hierarchynode+"."+fmositename
    sheet['C'+str(fila)]="modify"
    sheet['D'+str(fila)]="routePartitionName:"+nullpt+",pattern:"+hp
    sheet['I'+str(fila)]="Cisco CallManager"           # callingPartyNumberingPlan
    sheet['j'+str(fila)]="Default"                     # connectedLinePresentationBit
    sheet['k'+str(fila)]=preisrpt                      # routePartitionName
    sheet['l'+str(fila)]="No Error"                    # releaseClause
    sheet['m'+str(fila)]="false"                       # blockEnable
    sheet['n'+str(fila)]="Cisco CallManager"           # callingPartyNumberType
    sheet['o'+str(fila)]="false"                       # provideOutsideDialtone
    sheet['q'+str(fila)]=hp                        # pattern
    sheet['r'+str(fila)]="Default"                     # patternPrecedence
    sheet['s'+str(fila)]=isrcss                        # callingSearchSpaceName
    sheet['t'+str(fila)]=""                            # prefixDigitsOut
    sheet['u'+str(fila)]="Translation"                 # usage
    sheet['v'+str(fila)]="Cisco CallManager"           # calledPartyNumberingPlan
    sheet['w'+str(fila)]="true"                        # dontWaitForIDTOnSubsequentHops
    sheet['x'+str(fila)]="Default"                     # connectedNamePresentationBit
    sheet['y'+str(fila)]="PreISR Site "+siterange      # Description
    sheet['z'+str(fila)]="Default"                     # routeClass
    sheet['aa'+str(fila)]="Default"                    # callingNamePresentationBit
    sheet['ac'+str(fila)]="false"                      # routeNextHopByCgpn
    sheet['ad'+str(fila)]="false"                      # useOriginatorCss
    sheet['ae'+str(fila)]=""                           # digitDiscardInstructionName
    sheet['ag'+str(fila)]="Off"                        # useCallingPartyPhoneMask
    sheet['aj'+str(fila)]="Cisco CallManager"          # calledPartyNumberType
    sheet['al'+str(fila)]="false"                      # patternUrgency
    sheet['am'+str(fila)]="Default"                    # callingLinePresentationBit

    ######################################################
    ## Habilitamos el routing del site en  ILS
    sheet = blk["ILS"]
    sheet['B'+str(fila)]=hierarchynode+"."+fmositename
    sheet['C'+str(fila)]="add"
    sheet['i'+str(fila)]="0"                           # pstnFailStrip
    sheet['j'+str(fila)]=hp                            # pattern
    sheet['k'+str(fila)]="Specify"                     # hostedRoutePSTNRule
    sheet['l'+str(fila)]="Enterprise Number"           # patternType
    sheet['m'+str(fila)]=fmositename                   # Description

    fila+=1

## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)

## LOG de CONFIGURACION
f.close()

exit(0)
