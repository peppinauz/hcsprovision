## LINUX#!/usr/bin/python3
## OSX
#!/usr/local/bin/python3

import csv
import openpyxl
import time
import sys
import json

## CUCM data Inputs
MaxDNInLineGroup=82
MaxLineGroupsInHuntList=3

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

with open(fmositedata,'r') as infile:
    data = json.load(infile)
infile.close()

#FMO working path:
sitepath="../FMO/"+siteslc

# FILES
inputfilehl = clusterpath+"/huntlist.csv"       ## ORIGINAL
inputfilehp = clusterpath+"/huntpilot.csv"      ## ORIGINAL
#inputfilelg = clusterpath+"/linegroup.csv"      ## ORIGINAL ->
inputfilelg = clusterpath+"/linegroup.mod2.csv"      ## MOD

tmpcsvblkfile=sitepath+"/XX.NOT_VALID_BLK."+siteslc+".csv"
outputblkfile = sitepath+"/06.newhunt."+siteslc+".xlsx"

## FMO CUSTOMER INPUT DATA
hierarchynode=fmoenvconfig['hierarchynode']
customerid=fmoenvconfig['fmocustomerid']
aargroup=fmoenvconfig['fmoaargroup']
##
fmositename=data['fmosite']['name']
fmositeid=data['fmosite']['id']
cmg=data['fmosite']['cmg']

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
fwdhuntcss=cucdmsite+"-HPFWD-CSS"

data=[]
datahp=[]
datahl=[]
lginhl=[]
uniquehl=[]
uniquelg=[]
tmphl=[]
tmplg=[]
datalg=[]
dninlg=[]
relhg=[]

######################################################
## CONSTRUYO OBJETO de CONFIGURACION
######################################################
print("(II): CMO Hunt Pilots site",siteslc,file=f)

########################
## HPilot info :: Search

fhp = open(inputfilehp,"r")  ## CMO: huntpilot.csv
csv_hp = csv.DictReader(fhp)  ## CMO: huntpilot.csv


for rowhp in csv_hp:
    if len(rowhp) != 0: # Me salto las líneas vacias
        if rowhp['HUNT PILOT'][1:].startswith(siteslc):
            #datahp['huntpilot'].append({'pattern':rowhp['HUNT PILOT'],'routePartitionName':"PT",'description':rowhp['DESCRIPTION'],'huntListName':rowhp['HUNT LIST 1']})
            print("(DD) "+rowhp['HUNT PILOT']+";"+rowhp['ROUTE PARTITION']+";"+rowhp['DESCRIPTION']+";"+rowhp['HUNT LIST 1']+";"+rowhp['FORWARD HUNT NO ANSWER DESTINATION']+";"+rowhp['FORWARD HUNT BUSY DESTINATION'],file=f) ## DEBUG
            datahp.append({'pattern':rowhp['HUNT PILOT'],'routePartitionName':rowhp['ROUTE PARTITION'],'description':rowhp['DESCRIPTION'],'huntListName':rowhp['HUNT LIST 1'],'cfwdnoan':rowhp['FORWARD HUNT NO ANSWER DESTINATION'],'cfwdbusy':rowhp['FORWARD HUNT BUSY DESTINATION']})

##print("(HuntPilot):",datahp)    ## DEBUG

fhp.close()                     ## CMO: huntpilot.csv

## HPilot info :: Printintg INFO
print("-----------------------------------",file=f)
print("(DD) Hunt Pilot len("+str(len(datahp))+")",file=f)

for hp in datahp:
    print("(DD) "+hp['pattern']+"::"+hp['huntListName'],file=f)
print("HP(raw)-----------------------------------",file=f)
print(datahp,file=f)
print("HP(raw)-----------------------------------",file=f)

########################
## HList info :: Search
fhl = open(inputfilehl,"r")     ## CMO: huntlist.csv
csv_hl = csv.DictReader(fhl)     ## CMO: huntlist.csv

cmohl=list(csv_hl)   ## Lo pasamos a una variable tipo lista para poder iterar varias veces sobre ella

for hl in datahp:
    print("(DD) Buscando...."+hl['huntListName']+"...."+inputfilehl,file=f)
    for rowhl in cmohl:
        if len(rowhl) != 0: # Me salto las líneas vacias
            if rowhl['NAME'] == hl['huntListName']:
                ## DEBUG
                print("(DD) Encontrado "+hl['huntListName']+"::"+rowhl['NAME']+";"+rowhl['DESCRIPTION'],file=f)
                dn=1
                while (dn < MaxLineGroupsInHuntList):
                    if rowhl['SELECTION ORDER '+str(dn)] != "":
                        lginhl.append({'lineSelectionOrder':rowhl['SELECTION ORDER '+str(dn)],'lineGroup':rowhl['LINE GROUP '+str(dn)]})
                        dn=dn+1
                    else:
                        break
                dn=1
                datahl.append({'name':rowhl['NAME'],'description':rowhl['DESCRIPTION'],'callManagerGroupName':cmg,'LineGroup':lginhl})
                lginhl=[]

print("HL(sin duplicados)----------------------------------",file=f)
print(datahl,file=f)
print("HL(sin duplicados)----------------------------------",file=f)

fhl.close() ## CMO: huntlist.csv

## HList info :: Printintg INFO
#print("-----------------------------------")
#print("(DD) Hunt List len("+str(len(datahl))+")")

#for hl in datahl:
#    for i in hl['LineGroup']:
#        print("(DD) "+hl['name']+"::"+hl['description']+"::"+hl['callManagerGroupName']+"::"+i['lineSelectionOrder']+"::"+i['lineGroup'])
#print("-----------------------------------")

## HL: Quitar duplicados
for hl in datahl:
    if hl not in uniquehl:
        uniquehl.append(hl)

print("(DD) Hunt List len("+str(len(uniquehl))+") ----- SIN DUPLICADOS",file=f)

for hl in uniquehl:
    print("(DD) HL "+hl['name']+"::"+hl['description']+"::"+hl['callManagerGroupName'],file=f)
    for i in hl['LineGroup']:
        print("(DD) LG       "+i['lineSelectionOrder']+":"+i['lineGroup'],file=f)
print("HL(sin duplicados)----------------------------------",file=f)
print(hl,file=f)
print("HL(sin duplicados)----------------------------------",file=f)

########################
## LineGroup info :: Search
flg = open(inputfilelg,"r")     ## CMO: linegroup.csv
csv_lg = csv.DictReader(flg)     ## CMO: linegroup.csv

cmolg=list(csv_lg)   ## Lo pasamos a una variable tipo lista para poder iterar varias veces sobre ella

for listalg in uniquehl:            ## HuntList
    for lg in listalg['LineGroup']: ## HuntList(LineGroup1, LineGroup2,...)
        #print("(DD) Buscando...."+lg['lineGroup']+"...."+inputfilelg)
        for rowlg in cmolg:
            if len(rowlg) != 0:     ## Me salto las líneas vacias
                if rowlg['NAME'] == lg['lineGroup']:
                    ## DEBUG
                    print("(DD) HL("+listalg['name']+") "+lg['lineGroup']+" == "+rowlg['NAME']+"  ;  "+rowlg['TYPE DISTRIBUTION ALGORITHM'],file=f)
                    ##
                    dn=1
                    while (dn <= MaxDNInLineGroup):  ## 82 Maximo número de DN en LineGroup
                        if rowlg['LINE SELECTION ORDER '+str(dn)] != "":
                            dninlg.append({'lineSelectionOrder':rowlg['LINE SELECTION ORDER '+str(dn)],'directoryNumber_pattern':rowlg['DN OR PATTERN '+str(dn)],'directoryNumber_routePartitionName':rowlg['ROUTE PARTITION '+str(dn)]})
                            dn=dn+1
                        else:
                            break
                    dn=1

                    ## Normalizo los parametros que cambio
                    if "Hunt" in rowlg['TYPE HUNT ALGORITHM RNA']:
                        tmphuntAlgorithmNoAnswer="Try next member; then, try next group in Hunt List"
                    else:
                        tmphuntAlgorithmNoAnswer="Try next member, but do not go to next group"
                    if "Hunt" in rowlg['TYPE HUNT ALGORITHM BUSY']:
                        tmphuntAlgorithmBusy="Try next member; then, try next group in Hunt List"
                    else:
                        tmphuntAlgorithmBusy="Try next member, but do not go to next group"
                    if "Hunt" in rowlg['TYPE HUNT ALGORITHM DOWN']:
                        tmphuntAlgorithmNotAvailable="Try next member; then, try next group in Hunt List"
                    else:
                        tmphuntAlgorithmNotAvailable="Try next member, but do not go to next group"
                        tmphuntAlgorithmBusy="Try next member, but do not go to next group"
                    ## Añadimos datos
                    datalg.append({'name':rowlg['NAME'],'rnaReversionTimeOut':rowlg['RNA REVERSION TIMEOUT'],'distributionAlgorithm':rowlg['TYPE DISTRIBUTION ALGORITHM'],'huntAlgorithmNoAnswer':tmphuntAlgorithmNoAnswer,'huntAlgorithmBusy':tmphuntAlgorithmBusy,'huntAlgorithmNotAvailable':tmphuntAlgorithmNotAvailable,'members':dninlg})
                    dninlg=[]
                    #print(dninlg)

##print("(LineGroup):",datalg)  ## DEBUG

flg.close() ## CMO: huntlist.csv

## LineGroup info :: Printintg INFO
print("-----------------------------------",file=f)
print("(DD) Line Group len("+str(len(datalg))+")",file=f)

for lg in datalg:
    if len(lg['members']) != 0:
        print("(DD) "+lg['name']+"::"+lg['distributionAlgorithm']+"::"+lg['huntAlgorithmNoAnswer']+"::"+lg['huntAlgorithmBusy']+"::"+lg['huntAlgorithmNotAvailable']+"::"+lg['members'][0]['directoryNumber_pattern'],file=f)
    else:
        print("(DD) "+lg['name']+"::"+lg['distributionAlgorithm']+"::"+lg['huntAlgorithmNoAnswer']+"::"+lg['huntAlgorithmBusy']+"::"+lg['huntAlgorithmNotAvailable'],file=f)

print("-----------------------------------",file=f)

## LG: Quitar duplicados
for lg in datalg:
    if lg not in uniquelg:
        uniquelg.append(lg)

## LineGroup info :: Printintg INFO
print("-----------------------------------",file=f)
print("(DD) Line Group len("+str(len(uniquelg))+") -- SIN DUPLICADOS",file=f)

for lg in uniquelg:
    if len(lg['members']) != 0:
        print("(DD) "+lg['name']+"::"+lg['distributionAlgorithm']+"::"+lg['huntAlgorithmNoAnswer']+"::"+lg['huntAlgorithmBusy']+"::"+lg['huntAlgorithmNotAvailable']+"::"+lg['members'][0]['directoryNumber_pattern'],file=f)
    else:
        print("(DD) "+lg['name']+"::"+lg['distributionAlgorithm']+"::"+lg['huntAlgorithmNoAnswer']+"::"+lg['huntAlgorithmBusy']+"::"+lg['huntAlgorithmNotAvailable'],file=f)
##    print("(DD) "+lg['name']+"::"+lg['distributionAlgorithm']+"::"+lg['huntAlgorithmNoAnswer']+"::"+lg['huntAlgorithmBusy']+"::"+lg['huntAlgorithmNotAvailable']+"::"+lg['members'][0]['directoryNumber_pattern'],file=f)
print("-----------------------------------",file=f)

##########################################################
## entity: relation/HuntGroupRelation
## Creamos los datos con toda la info del HP::HL(LG1,LG2,..)
##########################################################
## Construimos la relación HL<> LineGroup
print("-----------------------------------",file=f)
print("         relation/HuntGroupRelation",file=f)
print("-----------------------------------",file=f)
for hl in uniquehl:
    for i in hl['LineGroup']:
        #print("(II) HList "+hl['name']+" >>> "+i['lineGroup'])
        for lg00X in uniquelg:
            ## DEBUG
            #print("(DD) relation/HuntGroupRelation:"+lg00X['name']+" == "+i['lineGroup'])
            #print("(II) HList "+hl['name']+" >>> "+i['lineGroup'])
            if lg00X['name'] == i['lineGroup']:
                #print("(DD)                        Adding.."+lg00X['name'])
                print("(DD) HList("+hl['name']+") "+i['lineGroup']+" >>> ",file=f)
                #print(lg00X,file=f)
                tmplg.append(lg00X)

    print("(II) ------------------ (II)",file=f)
    #print("(II) "+hl['name'])
    print(tmplg,file=f)
    print("(II) ------------------ (II)",file=f)

    tmphl.append({'name':hl['name'],'description':hl['description'],'callManagerGroupName':cmg,'LineGroup':tmplg})
    tmplg=[]

for hp in datahp:
    for hl00X in tmphl:
        if hp['huntListName'] == hl00X['name']:
            relhg.append({'pattern':hp['pattern'],'routePartitionName':hp['routePartitionName'],'description':hp['description'],'cfwdnoan':hp['cfwdnoan'],'cfwdbusy':hp['cfwdbusy'],'huntListName':hl00X})
print("(TODO)-----------------------------------",file=f)
print(relhg,file=f)
print("(TODO)-----------------------------------",file=f)

## Recorremos la realcion HP <> HL+LineGroup: relhg tenemos la terna HP<>HL<LG01,LG02,..>
for rhg in relhg:
    print("(II) Pattern: "+rhg['pattern']+" Partition:"+rhg['routePartitionName']+" Description:"+hp['description'],file=f)
    print("     HuntListName: "+rhg['huntListName']['name']+" Description:"+rhg['huntListName']['description']+" CCMGroup:"+rhg['huntListName']['callManagerGroupName'],file=f)
    for lgm in rhg['huntListName']['LineGroup']:
        print("     >> LineGroup:"+lgm['name']+" rnaReversionTimeOut:"+lgm['rnaReversionTimeOut']+" distributionAlgorithm:"+lgm['distributionAlgorithm'],file=f)
        print("                   huntAlgorithmNoAnswer:"+lgm['huntAlgorithmNoAnswer']+" huntAlgorithmBusy:"+lgm['huntAlgorithmBusy']+" huntAlgorithmNotAvailable:"+lgm['huntAlgorithmNotAvailable'],file=f)
        print("                   members:"+str(len(lgm['members'])),file=f)
        for mm in lgm['members']:
            print("                   "+mm['lineSelectionOrder']+":"+mm['directoryNumber_pattern']+":"+mm['directoryNumber_routePartitionName'],file=f)

##########################################################
## entity: relation/HuntGroupRelation ::
## Controlamos si existen los DNs
##########################################################

##########################################################
## entity: relation/HuntGroupRelation ::
## Construimos el BLK
##########################################################

##
## Creamos CSV.intermedio con toda la info
## Cabeceras ESTATICAS
fieldnamestatic=['comments','hierarchy','action','search_fields','device','template','ndl','pkid']
## Cabeceras Hunt Pilot
## ALL
fieldnamehuntpilot=['pattern','routePartitionName','description','dialPlanName','routeFilterName','patternPrecedence','huntListName','callPickupGroupName','alertingName','asciiAlertingName','blockEnable','provideOutsideDialtone','patternUrgency','releaseClause','#maxHuntduration','forwardHuntNoAnswer.usePersonalPreferences','forwardHuntNoAnswer.destination','forwardHuntNoAnswer.callingSearchSpaceName','forwardHuntBusy.usePersonalPreferences','forwardHuntBusy.destination','forwardHuntBusy.callingSearchSpaceName','forwardHuntBusy.usePersonalPreferences','parkMonForwardNoRetrieve.destination','#parkMonForwardNoRetrieve.callingSearchSpaceName','queueCalls.networkHoldMohAudioSourceID','queueCalls.maxCallersInQueue','queueCalls.maxWaitTimeInQueue','queueCalls.queueFullDestination','queueCalls.callingSearchSpacePilotQueueFull','queueCalls.maxWaitTimeDestination','queueCalls.callingSearchSpaceMaxWaitTime','queueCalls.noAgentDestination','queueCalls.callingSearchSpaceNoAgent','useCallingPartyPhoneMask','callingPartyTransformationMask','callingPartyPrefixDigits','callingLinePresentationBit','callingNamePresentationBit','callingPartyNumberType','callingPartyNumberingPlan','connectedLinePresentationBit','displayConnectedNumber','connectedNamePresentationBit','digitDiscardInstructionName','calledPartyTransformationMask','prefixDigitsOut','calledPartyNumberingPlan','calledPartyNumberType','aarNeighborhoodName','e164Mask']
## Cabeceras Hunt List
fieldnamehuntlist=['HuntList.name','HuntList.description','HuntList.callManagerGroupName','HuntList.routeListEnabled','HuntList.voiceMailUsage']
## Cabeceras Line Group (se construyen dinamicamente)
## EJEMPLO
##fieldnamelinegroup=['HuntList.LineGroup.X.name','HuntList.members.member.X.lineGroupName','HuntList.members.member.X.selectionOrder','HuntList.LineGroup.X.rnaReversionTimeOut','HuntList.LineGroup.X.distributionAlgorithm','HuntList.LineGroup.X.huntAlgorithmNoAnswer','HuntList.LineGroup.X.huntAlgorithmBusy','HuntList.LineGroup.X.huntAlgorithmNotAvailable','HuntList.LineGroup.X.autoLogOffHunt']
##fieldnamelinegroupmember=['HuntList.LineGroup.X.members.member.Y.lineSelectionOrder','HuntList.LineGroup.X.members.member.Y.directoryNumber.pattern','HuntList.LineGroup.X.members.member.Y.directoryNumber.routePartitionName']

fieldnamelinegroup=[]

ii=0
while ii <= MaxLineGroupsInHuntList:
    fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".name")
    fieldnamelinegroup.append("HuntList.members.member."+str(ii)+".lineGroupName")
    fieldnamelinegroup.append("HuntList.members.member."+str(ii)+".selectionOrder")
    fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".rnaReversionTimeOut")
    fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".distributionAlgorithm")
    fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".huntAlgorithmNoAnswer")
    fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".huntAlgorithmBusy")
    fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".huntAlgorithmNotAvailable")
    fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".autoLogOffHunt")
    jj=0
    while jj <= MaxDNInLineGroup:
        fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".members.member."+str(jj)+".lineSelectionOrder")
        fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".members.member."+str(jj)+".directoryNumber.pattern")
        fieldnamelinegroup.append("HuntList.LineGroup."+str(ii)+".members.member."+str(jj)+".directoryNumber.routePartitionName")
        jj=jj+1
    ii=ii+1

fieldname=fieldnamestatic
print("(DD) Building headers...STATIC FIELDS: ",file=f)
#print(fieldnamestatic)
print("(DD) Building headers...HuntPilot FIELDS: ",file=f)
#print(fieldnamehuntpilot)
for x in fieldnamehuntpilot:
    fieldname.append(x)
print("(DD) Building headers...HuntList FIELDS: ",file=f)
#print(fieldnamehuntlist)
for x in fieldnamehuntlist:
    fieldname.append(x)
print("(DD) Building headers...LineGroup FIELDS: ",file=f)
#print(fieldnamelinegroup)
for x in fieldnamelinegroup:
    fieldname.append(x)

print("(DD) Final headers FIELDS: ",file=f)
#print(fieldname)

with open(tmpcsvblkfile,'w') as outcsv:
    writer = csv.DictWriter(outcsv,fieldname,delimiter=';', lineterminator='\n')
    writer.writeheader()
    ## Escribo 2 líneas comentadas
    row={'comments':"#"}
    writer.writerow(row)
    writer.writerow(row)
    ##
    ## Escribo los datos en formtao CSV == rows
    row={}
    for rhg in relhg:
        ## STATIC DATA
        row={'comments':"",'hierarchy':hierarchynode+"."+fmositename,'action':"add",'search_fields':"",'device':"",'template':"",'ndl':"",'pkid':""}
        ##
        ## HUNT PILOT
        row['pattern']=rhg['pattern']
        #row['routePartitionName']=rhg['routePartitionName']
        row['routePartitionName']=linept
        row['description']=rhg['description']
        row['huntListName']=rhg['huntListName']['name']
        row['forwardHuntNoAnswer.destination']=rhg['cfwdnoan']
        row['forwardHuntNoAnswer.callingSearchSpaceName']=fwdhuntcss
        row['forwardHuntBusy.destination']=rhg['cfwdbusy']
        row['forwardHuntBusy.callingSearchSpaceName']=fwdhuntcss
        row['useCallingPartyPhoneMask']="Off"
        ##
        ## HUNT LIST
        row['HuntList.name']=rhg['huntListName']['name']
        row['HuntList.description']=rhg['huntListName']['description']
        row['HuntList.callManagerGroupName']=rhg['huntListName']['callManagerGroupName']
        row['HuntList.routeListEnabled']="true"
        row['HuntList.voiceMailUsage']="false"

        ## LINEGROUP
        ##fieldnamelinegroup=['HuntList.LineGroup.X.name','HuntList.members.member.X.lineGroupName','HuntList.members.member.X.selectionOrder','HuntList.LineGroup.X.rnaReversionTimeOut','HuntList.LineGroup.X.distributionAlgorithm','HuntList.LineGroup.X.huntAlgorithmNoAnswer','HuntList.LineGroup.X.huntAlgorithmBusy','HuntList.LineGroup.X.huntAlgorithmNotAvailable','HuntList.LineGroup.X.autoLogOffHunt']
        ##fieldnamelinegroupmember=['HuntList.LineGroup.X.members.member.Y.lineSelectionOrder','HuntList.LineGroup.X.members.member.Y.directoryNumber.pattern','HuntList.LineGroup.X.members.member.Y.directoryNumber.routePartitionName']
        hllg=0
        for lgm in rhg['huntListName']['LineGroup']:
            row['HuntList.LineGroup.'+str(hllg)+'.name']=lgm['name']
            row['HuntList.members.member.'+str(hllg)+'.lineGroupName']=lgm['name']
            row['HuntList.members.member.'+str(hllg)+'.selectionOrder']=str(hllg)
            row['HuntList.LineGroup.'+str(hllg)+'.rnaReversionTimeOut']=lgm['rnaReversionTimeOut']
            row['HuntList.LineGroup.'+str(hllg)+'.distributionAlgorithm']=lgm['distributionAlgorithm']
            row['HuntList.LineGroup.'+str(hllg)+'.huntAlgorithmNoAnswer']=lgm['huntAlgorithmNoAnswer']
            row['HuntList.LineGroup.'+str(hllg)+'.huntAlgorithmBusy']=lgm['huntAlgorithmBusy']
            row['HuntList.LineGroup.'+str(hllg)+'.huntAlgorithmNotAvailable']=lgm['huntAlgorithmNotAvailable']
            row['HuntList.LineGroup.'+str(hllg)+'.autoLogOffHunt']=""

            ##
            lgdn=0
            for mm in lgm['members']:
                row['HuntList.LineGroup.'+str(hllg)+'.members.member.'+str(lgdn)+'.lineSelectionOrder']=mm['lineSelectionOrder']
                row['HuntList.LineGroup.'+str(hllg)+'.members.member.'+str(lgdn)+'.directoryNumber.pattern']=mm['directoryNumber_pattern']
                if mm['directoryNumber_routePartitionName'] == "Interna-PT": ## IP-PHONE
                    row['HuntList.LineGroup.'+str(hllg)+'.members.member.'+str(lgdn)+'.directoryNumber.routePartitionName']=emlinept
                elif mm['directoryNumber_routePartitionName'] == "Interna-EM-PT": ## EM
                    row['HuntList.LineGroup.'+str(hllg)+'.members.member.'+str(lgdn)+'.directoryNumber.routePartitionName']=linept
                elif mm['directoryNumber_routePartitionName'] == "PT-Migrated":  ## Hunt List
                    row['HuntList.LineGroup.'+str(hllg)+'.members.member.'+str(lgdn)+'.directoryNumber.routePartitionName']=linept
                else:
                    ## Condicion de error
                    print("(EE) Particion erronea: ",file=f)
                #row['HuntList.LineGroup.'+str(hllg)+'.members.member.'+str(lgdn)+'.directoryNumber.routePartitionName']=mm['directoryNumber_routePartitionName']
                lgdn=lgdn+1
            ## Volvemos a empezar
            hllg=hllg+1
            lgdn=0
        ## Escribimos la línea
        writer.writerow(row)
    ## Cerramos el fichero
    outcsv.close()

##
## Cargamos las "ROW" CSV -> Excel
blk = openpyxl.Workbook()
sheet =  blk.active
## WOrkFlow A ejecutar
sheet['A1']="entity: relation/HuntGroupRelation; parallel: False;"

ftmp = open(tmpcsvblkfile,"r")     ## CMO: huntlist.csv
csv_f = csv.reader(ftmp,delimiter=";")     ## CMO: huntlist.csv

for fila in csv_f:
    sheet.append(fila)

ftmp.close()
blk.save(outputblkfile)

exit(0)
