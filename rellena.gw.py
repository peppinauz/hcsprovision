## LINUX#!/usr/bin/python3
## OSX
#!/usr/local/bin/python3

import csv
import openpyxl
import time
import sys
import json
import os

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

# FMO datos de entorno:
fmoenvconfig={}
with open(fmostaticdata, "r") as fp:
        for line in fp.readlines():
            li = line.lstrip()
            if not li.startswith("#") and '=' in li:
                key, value = line.split('=', 1)
                fmoenvconfig[key] = value.strip()  ## variable de tipo diccionario
                #print("<<<<   ",fmoenvconfig[key], file=f)
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
inputgwfile = clusterpath+"/gateway.csv"       ## Fichero ORIGINAL
gatewaygw=clusterpath+"/gateway.gw.csv"          ## Fichero con los GW
gatewayslot=clusterpath+"/gateway.slot.csv"      ## Fichero con tarjetas del GW
gatewayanalog=clusterpath+"/gateway.analog.csv"  ## Fichero con puertos analogicos
##gatewayh323=clusterpath+"/gateway.h323.csv"      ## Fichero con Trunks H323 (no se usa)
##
templateblkfile = "../code/blk/03.gw-template.xlsx" # SIN DATAINPUT
outputblkfile = sitepath+"/03.gw."+siteslc+".xlsx"

## FMO CUSTOMER INPUT DATA
hierarchynode=fmoenvconfig['hierarchynode']
customerid=fmoenvconfig['fmocustomerid']
aargroup=fmoenvconfig['fmoaargroup']
##
fmositename=data['fmosite'][0]['name']
fmositeid=data['fmosite'][0]['id']
cmg=data['fmosite'][0]['cmg']

# CMO patterns
gwdomain="VACIO"
sitegwname=siteslc+"-GW"
cmodevicepool=siteslc+"-DP"
cmolocation=siteslc+"-LOC"

## FMO UserData
cucdmsite=fmoenvconfig['fmocustomerid']+"Si"+str(fmositeid)
cssfwd=customerid+"-DirNum-CSS"
linept=customerid+"-DirNum-PT"
#linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"  ## STD Nat
linecss=cucdmsite+"-DBREnhIntl24HrsCLIPyFONnFACnCMC-CSS"
linefaxcss=cucdmsite+"-DBREnhIntlFAX-CSS"
aarcss=customerid+"-AAR-CSS"
devicepool=cucdmsite+"-DevicePool"
location=cucdmsite+"-Location"
devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
subscribecss=cucdmsite+"-InternalOnly-CSS"
fmopass="vi123456"
fmopin="123456"

# CMO File INPUT DATA
fgw = open(gatewayanalog,"r")
csv_f = csv.DictReader(fgw)

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)
sheet = blk["GATEWAY"]

# FMO commands:
action="add"

fila=5
filagw=5
filadn=5
filaan=5

print("(II) Configurando GW",file=f)

for row in csv_f:
    if len(row) > 8: # Me salto las líneas vacias
        if row['DEVICE POOL'].startswith(cmodevicepool):

            ######################################################
            ## USER
            ######################################################
            sheet =  blk["SUBS"]
            sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
            sheet['C'+str(filadn)]=action
            #sheet['D'+str(filadn)]="name:"+row[3] # Search field
            ######################################################
            sheet['k'+str(filadn)]="false"                        # imAndPresenceEnable
            sheet['l'+str(filadn)]="false"                        # calendarPresence
            sheet['m'+str(filadn)]="false"                        # enableMobileVoiceAccess
            sheet['p'+str(filadn)]="true"                         # homeCluster
            sheet['q'+str(filadn)]="10000"                        # maxDeskPickupWaitTime
            sheet['r'+str(filadn)]=row['DIRECTORY NUMBER 1']      # primaryExtension
            sheet['u'+str(filadn)]="false"                        # pinCredentials.pinCredLockedByAdministrator
            sheet['v'+str(filadn)]=""                             # pinCredentials.pinCredTimeAdminLockout
            sheet['w'+str(filadn)]="true"                         # pinCredentials.pinCredDoesNotExpire
            sheet['x'+str(filadn)]="Default Credential Policy"    # pinCredentials.pinCredPolicyName
            sheet['y'+str(filadn)]="false"                        # pinCredentials.pinCredUserCantChange
            sheet['z'+str(filadn)]="false"                        # pinCredentials.pinCredTimeChanged
            sheet['aa'+str(filadn)]="true"                        # pinCredentials.pinCredUserMustChange
            sheet['ad'+str(filadn)]="false"                       # enableUserToHostConferenceNow
            sheet['af'+str(filadn)]="4"                           # remoteDestinationLimit
            sheet['ag'+str(filadn)]="1"                           # status
            sheet['ai'+str(filadn)]="true"                        # enableMobility
            sheet['aj'+str(filadn)]="false"                       # enableEmcc
            sheet['ak'+str(filadn)]=subscribecss                  # subscribeCallingSearchSpaceName
            sheet['al'+str(filadn)]="false"                       # passwordCredentials.pwdCredUserMustChange
            sheet['am'+str(filadn)]="false"                       # passwordCredentials.pwdCredUserCantChange
            sheet['ao'+str(filadn)]="false"                       # passwordCredentials.pwdCredTimeAdminLockout
            sheet['ap'+str(filadn)]="Default Credential Policy"   # passwordCredentials.pwdCredPolicyName
            sheet['aq'+str(filadn)]="false"                       # passwordCredentials.pwdCredLockedByAdministrator
            sheet['ar'+str(filadn)]="true"                        # passwordCredentials.pwdCredDoesNotExpire
            sheet['at'+str(filadn)]="p"+row['DIRECTORY NUMBER 1'] # HcsUserProvisioningStatusDAT.username
            sheet['ax'+str(filadn)]="true"                        # enableCti
            #sheet['be'+str(filadn)]=                             # mailid
            #sheet['ax'+str(filadn)]=                             # phoneProfiles
            sheet['be'+str(filadn)]="Standard Presence group"     # presenceGroupName
            #sheet['bf'+str(filadn)]=row['FIRST NAME']              # first name
            sheet['bg'+str(filadn)]=row['DESCRIPTION']              # lastname
            sheet['bi'+str(filadn)]="p"+row['DIRECTORY NUMBER 1'] # userid
            #sheet['bj'+str(filadn)]=""                           # ctiControlledDeviceProfiles
            sheet['bk'+str(filadn)]="p"+row['DIRECTORY NUMBER 1'] # NormalizedUser.username
            sheet['bl'+str(filadn)]="CUCM Local"                  # NormalizedUser.userType
            sheet['bm'+str(filadn)]="p"+row['DIRECTORY NUMBER 1'] # NormalizedUser.sn.0
            sheet['bq'+str(filadn)]="Standard CCM End Users"      # associatedGroups.userGroup.0.name
            sheet['br'+str(filadn)]="Standard CCM End Users"      # associatedGroups.userGroup.0.userRoles.userRole.0
            sheet['bs'+str(filadn)]="Standard CCMUSER Administration" # associatedGroups.userGroup.0.userRoles.userRole.1
            sheet['bv'+str(filadn)]=fmopass                       # NormalizedUser.userType
            #sheet['bw'+str(filadn)]=fmopin                       # NormalizedUser.userType

            ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
            ######################################################
            sheet = blk["LINE.PORT"]
            sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
            sheet['C'+str(filadn)]=action
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['I'+str(filadn)]=row['ASCII DISPLAY 1']                                   # asciiAlertingName
            sheet['N'+str(filadn)]=row['FORWARD UNREGISTERED INTERNAL DESTINATION 1']       # callForwardNotRegisteredInt.destination
            sheet['P'+str(filadn)]=cssfwd                                                   # callForwardNotRegisteredInt.callingSearchSpaceName
            sheet['Q'+str(filadn)]=linept                                                   # routePartitionName
            sheet['R'+str(filadn)]=row['FORWARD ON CTI FAILURE DESTINATION 1']              # callForwardOnFailure.destination
            sheet['T'+str(filadn)]=cssfwd                                                   # callForwardOnFailure.callingSearchSpaceName
            ##
            ## Verificamos la Device CSS para establecer el LineCSS
            ## DEBUG
            #print("(II) FAX:",row['CALLING SEARCH SPACE'], file=f)
            if row['CALLING SEARCH SPACE'] == "FAX-PSTN-CSS":
                ## DEBUG
                #print("(II) linea de FAX:",row ['DIRECTORY NUMBER 1'], file=f)
                sheet['W'+str(filadn)]=linefaxcss                                           # shareLineAppearanceCssName
            else:
                sheet['W'+str(filadn)]=linecss                                              # shareLineAppearanceCssName
            ##
            sheet['X'+str(filadn)]=data['e164'][0]['head'][:9]+row['DIRECTORY NUMBER 1'][5:]  # aarDestinationMask
            if (len(row) >= 117):
                sheet['Y'+str(filadn)]=row['CALL PICKUP GROUP 1']                           # callPickupGroupName
            sheet['Z'+str(filadn)]=row['DIRECTORY NUMBER 1']                                # pattern
            sheet['AC'+str(filadn)]=row['FORWARD NO ANSWER EXTERNAL DESTINATION 1']         # callForwardNoAnswer.destination
            sheet['AE'+str(filadn)]=cssfwd                                                  # callForwardNoAnswer.callingSearchSpaceName
            sheet['AG'+str(filadn)]=row['FORWARD NO COVERAGE EXTERNAL DESTINATION 1']       #callForwardNoCoverage.destination
            sheet['AI'+str(filadn)]=cssfwd                                                  #callForwardNoCoverage.callingSearchSpaceName
            sheet['AJ'+str(filadn)]=row['FORWARD UNREGISTERED EXTERNAL DESTINATION 1']      #callForwardNotRegistered.destination
            sheet['AL'+str(filadn)]=cssfwd                                                  #callForwardNotRegistered.callingSearchSpaceName
            sheet['AP'+str(filadn)]=row['ALERTING NAME 1']                                  # alertingName
            sheet['AQ'+str(filadn)]=row['LINE DESCRIPTION 1']                               # Description
            #sheet['BE'+str(filadn)]=row[]                                                  #callForwardAll.secondaryCallingSearchSpaceName
            sheet['BF'+str(filadn)]=row['FORWARD ALL DESTINATION 1']                        #callForwardAll.destination
            sheet['BH'+str(filadn)]=cssfwd                                                  #callForwardAll.callingSearchSpaceName
            sheet['BH'+str(filadn)]=row['FORWARD BUSY INTERNAL DESTINATION 1']              # callForwardBusyInt.destination
            sheet['BU'+str(filadn)]=cssfwd                                                  # callForwardBusyInt.callingSearchSpaceName
            sheet['BW'+str(filadn)]=row['FORWARD NO ANSWER INTERNAL DESTINATION 1']         # callForwardNoAnswerInt.destination
            sheet['BU'+str(filadn)]=cssfwd                                                  # callForwardNoAnswerInt.callingSearchSpaceName
            sheet['BZ'+str(filadn)]=row['FORWARD BUSY EXTERNAL DESTINATION 1']              # callForwardBusy.destination
            sheet['CB'+str(filadn)]=cssfwd                                                  # callForwardBusy.callingSearchSpaceName
            sheet['CG'+str(filadn)]=row['FORWARD NO COVERAGE INTERNAL DESTINATION 1']       # callForwardNoCoverageInt.destination
            sheet['CI'+str(filadn)]=cssfwd                                                  # callForwardNoCoverageInt.callingSearchSpaceName
            ## ILS
            sheet['CE'+str(filadn)]=aargroup                                                #aarNeighborhoodName
            sheet['BP'+str(filadn)]=aarcss                                                  #enterpriseAltNum.advertiseGloballyIls
            #sheet[''+str(filadn)]= #

            ######################################################
            sheet = blk["GATEWAY.PORT"]
            sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
            sheet['C'+str(filadn)]=action
            #sheet['D'+str(filadn)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['I'+str(filadn)]=row['SUBUNIT POSITION']                                  # subunit
            sheet['J'+str(filadn)]=row['PROTOCOL TYPE']                                     # endpoint.protocol
            sheet['O'+str(filadn)]="Default"                                                # endpoint.alwaysUsePrimeLine
            sheet['Q'+str(filadn)]=row['PHONE BUTTON TEMPLATE']                             # endpoint.phoneTemplateName
            sheet['R'+str(filadn)]=devicecss                                                # endpoint.callingSearchSpaceName
            sheet['S'+str(filadn)]=row['PORT NUMBER']                                       # endpoint.index
            sheet['V'+str(filadn)]="Default"                                                # cendpoint.useTrustedRelayPoint
            sheet['AA'+str(filadn)]="p"+row['DIRECTORY NUMBER 1']                           # endpoint.userid
            sheet['AF'+str(filadn)]=row['PRODUCT TYPE']                                     # endpoint.product
            sheet['AH'+str(filadn)]="Default"                                               # endpoint.alwaysUsePrimeLineForVM
            sheet['AI'+str(filadn)]=row['DESCRIPTION']                                      # endpoint.description
            sheet['AP'+str(filadn)]="Off"                                                   # endpoint.deviceMobilityMode
            sheet['AT'+str(filadn)]="Phone"                                                 # endpoint.class
            sheet['AV'+str(filadn)]=row['DEVICE SECURITY PROFILE']                          # endpoint.securityProfileName
            sheet['AW'+str(filadn)]="Standard Presence group"                               # endpoint.presenceGroupName
            sheet['AX'+str(filadn)]=row['ENDPOINT NAME']                                    # endpoint.name
            sheet['AZ'+str(filadn)]="User"                                                  # endpoint.protocolSide
            ##sheet['BA'+str(filadn)]=row['COMMON DEVICE CONFIGURATION']                      #endpoint.commonPhoneConfigName
            ## REQ UCC Delivery: 2018-11-29
            sheet['BA'+str(filadn)]="Standard Common Phone Profile"                      #endpoint.commonPhoneConfigName

            locationtmp=row['LOCATION'].replace('-NOSRST-','')
            devicepooltmp=row['DEVICE POOL'].replace('-NOSRST-','')
            #print("(II)",row['LOCATION'],">>>",locationtmp.replace(cmolocation,cucdmsite+"-Location"), file=f)
            #print("(II)",row['DEVICE POOL'],">>>",devicepooltmp.replace(cmodevicepool,cucdmsite+"-DevicePool"), file=f)
            sheet['BE'+str(filadn)]=locationtmp.replace(cmolocation,cucdmsite+"-Location")                  #endpoint.locationName
            sheet['BI'+str(filadn)]=devicepooltmp.replace(cmodevicepool,cucdmsite+"-DevicePool")                #endpoint.devicePoolName

            sheet['BL'+str(filadn)]=row['ASCII DISPLAY 1']                                  #endpoint.lines.line.0.displayAscii
            #sheet['BM'+str(filadn)]=row[]                                                  #userid
            #sheet['BS'+str(filadn)]=row[]                                                  # endpoint.lines.line.0.label
            sheet['BZ'+str(filadn)]=data['e164'][0]['head'][:9]+row['DIRECTORY NUMBER 1'][5:] # endpoint.lines.line.0.e164Mask
            sheet['CF'+str(filadn)]=row['DIRECTORY NUMBER 1']                               # endpoint.lines.line.0.dirn.pattern
            sheet['CG'+str(filadn)]=linept                                                  # endpoint.lines.line.0.dirn.routePartitionName
            sheet['CL'+str(filadn)]=row['DISPLAY 1']                                        # endpoint.lines.line.0.display
            sheet['CM'+str(filadn)]=row['SLOT POSITION']                                    # unit
            sheet['CN'+str(filadn)]=row['GATEWAY NAME']                                     # domainName

            ## ILS
            sheet['BD'+str(filadn)]=aargroup            #aarNeighborhoodName
            sheet['BH'+str(filadn)]=aarcss              #aarCSS
            #sheet[''+str(filadn)]= #
            ## DEBUG
            print("(II) GW PORT:",row['PORT NUMBER'],"::",row['DIRECTORY NUMBER 1'],"::",row['CALLING SEARCH SPACE'],"::",row['DEVICE POOL'],">>",devicepooltmp.replace(cmodevicepool,cucdmsite+"-DevicePool"),"::",row['LOCATION'],">>",locationtmp.replace(cmolocation,cucdmsite+"-Location"), file=f)

            # Next row
            filadn=filadn+1
            gwdomain=row['GATEWAY NAME']

## CMO File INPUT DATA: Close ## GATEWAY_ANALOG
fgw.close()


# CMO File INPUT DATA ## GATEWAY_GW
fgw = open(gatewaygw,"r")
csv_f = csv.DictReader(fgw)

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        if row['DOMAIN NAME'].startswith(gwdomain):
            gwproduct=row['PRODUCT']
            gwprotocol=row['PROTOCOL']
            #print("(II) Gateway:: ",gwdomain," :: ",gwproduct," :: ",gwprotocol, file=f)

fgw.close()  ## CMO File INPUT DATA: Close ## GATEWAY_GW

if gwdomain == "VACIO":
    print("(II) Gateway:: NO se ha encontrado Gateway", file=f)
    print("(II) Gateway:: NO se tiene que cargar el BLK de Gateway", file=f)
    print("(II) Gateway:: Salimos sin completar el BLK de Gateway", file=f)
    print("(II) Gateway:: Borramos el BLK de Gateway", file=f)
    ## Ceramos todos los ficheros
    blk.save(outputblkfile)     ## FMO File OUTPUT DATA: Close
    f.close()                   ## LOG de CONFIGURACION
    ## Borrar fichero outputblkfile
    os.remove(outputblkfile)
    exit(1)
else:
    print("(II) Gateway:: ",gwdomain," :: ",gwproduct," :: ",gwprotocol, file=f)


# CMO File INPUT DATA ## GATEWAY_SLOT
fgw = open(gatewayslot,"r")
csv_f = csv.DictReader(fgw)

print("(II) Buscando Gateway:: ",gwdomain," en ",gatewayslot, file=f)

for row in csv_f:
    # WR BLK OUTPUT DATA
    #print(row)
    if len(row) != 0: # Me salto las líneas vacias
        if row['GATEWAY NAME'].startswith(gwdomain) and row['SUBUNIT POSITION'] != "":
            sheet = blk["GATEWAY"]
            #print(row)
            ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
            sheet['B'+str(filagw)]=hierarchynode+"."+fmositename
            sheet['C'+str(filagw)]=action
            #sheet['D'+str(filagw)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['I'+str(filagw)]=row['SLOT POSITION']              # units.unit.0.index
            sheet['J'+str(filagw)]=row['SLOT MODULE']                # units.unit.0.product
            sheet['K'+str(filagw)]=row['SUBUNIT POSITION']           # units.unit.0.subunits.subunit.0.index
            sheet['L'+str(filagw)]=row['VIC']                        # units.unit.0.subunits.subunit.0.product
            sheet['M'+str(filagw)]=row['BEGINNING PORTNUMBER']       # units.unit.0.subunits.subunit.0.beginPort
            sheet['N'+str(filagw)]=gwproduct                         # Product
            sheet['O'+str(filagw)]=cmg                               # CMG
            sheet['P'+str(filagw)]=gwprotocol                        # Protocol
            sheet['q'+str(filagw)]=row['GATEWAY NAME']

fgw.close()                 ## CMO File INPUT DATA ## GATEWAY_SLOT
blk.save(outputblkfile)     ## FMO File OUTPUT DATA: Close
f.close()                   ## LOG de CONFIGURACION

exit(0)
