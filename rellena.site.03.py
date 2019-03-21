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
print("(II) GENERANDO SITE BLK de FMO", file=f)

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


#INPUT: E164
e164=data['e164'][0]['head']
fmotrunkipaddr=data['gw'][0]['trunk']
fmositename=data['fmosite'][0]['name']
fmositeid=data['fmosite'][0]['id']
cmg=data['fmosite'][0]['cmg']

#FMO working path:
sitepath="../FMO/"+siteslc

# CMO patrones
cmodevicepool=siteslc+"-DP"
cmolocation=siteslc+"-LOC"
sitegwname=siteslc+"-GW"
gwdomain=""
gwnametmp=""

##### FMO SITE INPUT DATA: Built-in
fmosite=fmoenvconfig['fmocustomerid']+"Si"+str(fmositeid)
cssfwd=fmoenvconfig['fmocustomerid']+"-DirNum-CSS"
linept=fmoenvconfig['fmocustomerid']+"-DirNum-PT"
linecss=fmosite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
aarcss=fmoenvconfig['fmocustomerid']+"-AAR-CSS"
devicepool=fmosite+"-DevicePool"
location=fmosite+"-Location"
devicecss=fmosite+"-BRADP-DBRDevice-CSS"
aintpt=fmosite+"-AInt-PT"
lbipt=fmosite+"-LBI-PT"
localpt=fmosite+"-BRADP-Local-PT"
ilscss=fmoenvconfig['fmocustomerid']+"-ILS-CSS"
tbsdddnacional="00[1-9][1-9].11[2-57]XXXXXXX"
tbsdddmobile="00[1-9][1-9].119XXXXXXXX"
preisrcss=fmoenvconfig['fmocustomerid']+"-PreISR-CSS"
#fmotrunkincomingcss=fmoenvconfig['fmocustomerid']+"-IngressFromCBO-CSS"
fmotrunkincomingcss=fmosite+"-LBI-CSS"
fmotrunkname=siteslc+"-trunk"

## FICHEROS
dpinputfile =clusterpath+"/devicepool.csv"              ## BBDD CMO
templateblkfile = "blk/01.site-template.xlsx"   ## FMO SITE TEMPLATE
outputblkfile = sitepath+"/01.site."+siteslc+".xlsx"    ## FMO BLK OUTPUT

# CMO File INPUT DATA
fin = open(dpinputfile,"r")
csv_f = csv.reader(fin)

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

# BLK index:
sitaddfila=5
addlocfila=5
addregfila=5
adddpfila=5
mtpaddfila=5
cnfaddfila=5
mrgaddfila=5
mrgladdfil=5
trkaddfila=5
rdaddfila=5
tpaddfila=5
srstaddfil=5
xcdaddfila=5
fila=5

## DEBUG
#print(len(data['dp']),">>>>",data['dp'])
#print(data['e164'][0]['head'])
#print(data['mtp'])
print("(II): Configurando SITE", file=f)

for dp in data['dp']:
    #print("(II) \n",dp,"\n(II)", file=f)
    if dp['devicepool'] == cmodevicepool:
        #################### DP.SITE ####################
        sheet = blk["DP.SITE"]
        ## datos estaticos: Hierarchy, site,...
        sheet['B'+str(sitaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(sitaddfila)]="add"
        #sheet['D'+str(fila)]="name:"+row[3] # Search field
        ## datos dinamicos
        sheet['I'+str(sitaddfila)]=4                      # extLen
        sheet['J'+str(sitaddfila)]=0                      # ext
        sheet['K'+str(sitaddfila)]=fmositename            # units.unit.0.subunits.subunit.0.index
        sheet['L'+str(sitaddfila)]="true"              # slcBased
        sheet['M'+str(sitaddfila)]="0"                    # emerNumber
        sheet['N'+str(sitaddfila)]=fmoenvconfig['fmocustomername']           # custName
        sheet['p'+str(sitaddfila)]=data['e164'][0]['ac']               # areaCodeArray.0.areaCode
        sheet['q'+str(sitaddfila)]=8                      # areaCodeArray.0.locNumLen
        sheet['r'+str(sitaddfila)]="false"              # extPrefixReq
        sheet['s'+str(sitaddfila)]=fmositeid              # siteId
        sheet['t'+str(sitaddfila)]="false"              # active
        sheet['U'+str(sitaddfila)]=data['e164'][0]['head']                   # pubNumber
        sheet['V'+str(sitaddfila)]="false"              # areaCodeInLocalDialing
        sheet['W'+str(sitaddfila)]=data['e164'][0]['slc']                # slc

        for rsc in data['mtp']:
            #print(rsc,"---  ")
            #################### RECURSOS ####################
            sheet = blk["MTP.ADD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(mtpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(mtpaddfila)]="add"
            #sheet['D'+str(mtpaddfila)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['I'+str(mtpaddfila)]=rsc['nombre']                                       # name
            sheet['j'+str(mtpaddfila)]="false"                               # callingLineIdPresentation
            sheet['l'+str(mtpaddfila)]=rsc['tipo']                                       # mtptype
            sheet['m'+str(mtpaddfila)]="Default"                                # devicePoolName
            #################### RECURSOS ####################
            sheet = blk["MTP.MOD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(mtpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(mtpaddfila)]="modify"
            sheet['D'+str(mtpaddfila)]="name:"+rsc['nombre'] # Search field
            ## datos dinamicos
            sheet['I'+str(mtpaddfila)]=rsc['nombre']                                       # name
            sheet['j'+str(mtpaddfila)]="false"                               # callingLineIdPresentation
            sheet['l'+str(mtpaddfila)]=rsc['tipo']                                       # mtptype
            sheet['m'+str(mtpaddfila)]=devicepool
            mtpaddfila=mtpaddfila+1

        for rsc in data['cnf']:
            #print(rsc,"---  ")
            #################### RECURSOS ####################
            sheet = blk["CONF.ADD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(cnfaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(cnfaddfila)]="add"
            #sheet['D'+str(cnfaddfila)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['i'+str(cnfaddfila)]=""                                       # normalizationscript
            sheet['j'+str(cnfaddfila)]=rsc['tipo']                                        # product
            sheet['k'+str(cnfaddfila)]="Non Secure Conference Bridge"           # securityProfileName
            sheet['n'+str(cnfaddfila)]="Default"                                # useTrustedRelayPoint
            sheet['r'+str(cnfaddfila)]="Hub_None"                               # LocationName
            sheet['u'+str(cnfaddfila)]="false"                               #allowCFBControlOfCallSecurityIcon
            sheet['v'+str(cnfaddfila)]="Default"                                # devicePoolName
            sheet['y'+str(cnfaddfila)]=rsc['nombre']                                        # name
            #################### RECURSOS ####################
            sheet = blk["CONF.MOD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(cnfaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(cnfaddfila)]="modify"
            sheet['D'+str(cnfaddfila)]="name:"+rsc['nombre']                     # Search field
            ## datos dinamicos
            sheet['i'+str(cnfaddfila)]=""                                       # normalizationscript
            sheet['j'+str(cnfaddfila)]=rsc['tipo']                                        # product
            sheet['k'+str(cnfaddfila)]="Non Secure Conference Bridge"           # securityProfileName
            sheet['n'+str(cnfaddfila)]="Default"                                # useTrustedRelayPoint
            sheet['r'+str(cnfaddfila)]=location                               # LocationName
            sheet['u'+str(cnfaddfila)]="false"                               #allowCFBControlOfCallSecurityIcon
            sheet['v'+str(cnfaddfila)]=devicepool                                # devicePoolName
            sheet['y'+str(cnfaddfila)]=rsc['nombre']                                        # name
            cnfaddfila=cnfaddfila+1

        for rsc in data['trans']:
            #print(rsc,"---  ")
            #################### RECURSOS ####################
            sheet = blk["XCODE.ADD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(xcdaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(xcdaddfila)]="add"
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['i'+str(xcdaddfila)]=""                                       # loadinfo
            sheet['j'+str(xcdaddfila)]=rsc['tipo']                                       # product
            sheet['k'+str(xcdaddfila)]=rsc['nombre']                                       # name
            sheet['m'+str(xcdaddfila)]="Default"                                # devicePoolName
            sheet['n'+str(xcdaddfila)]="false"                               # isTrustedRelayPoint
            #################### RECURSOS ####################
            sheet = blk["XCODE.ADD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(xcdaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(xcdaddfila)]="add"
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['i'+str(xcdaddfila)]=""                                       # loadinfo
            sheet['j'+str(xcdaddfila)]=rsc['tipo']                                       # product
            sheet['k'+str(xcdaddfila)]=rsc['nombre']                                       # name
            sheet['m'+str(xcdaddfila)]=devicepool                                # devicePoolName
            sheet['n'+str(xcdaddfila)]="false"                               # isTrustedRelayPoint
            xcdaddfila=xcdaddfila+1


        sumarsc=len(data['cnf'])+len(data['mtp'])+len(data['trans'])
        ## DEBUG
        #print("CNF >>>",len(data['cnf']),">>> ",cnfaddfila)
        #print("MTP >>>",len(data['mtp']),">>> ",mtpaddfila)
        #print("TRANS >>>",len(data['trans']),">>> ",xcdaddfila)
        #print("RSC >>>",sumarsc)
        #print(len(data['mrg']),data['mrg'],data['mrg'][0],data['mrg'][1])
        sumarsc=len(data['cnf'])+len(data['mtp'])+len(data['trans'])

        if sumarsc > 6: ## Maximos numero de recursos
            print("(WW): Maximo numero de recursos alcanzados solo configurarn 6",file=f)

        if len(data['mrg']) > 0:
            sheet = blk["MRG.ADD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(mrgaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(mrgaddfila)]="add"
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['i'+str(mrgaddfila)]="0"                                       # Multicast
            sheet['j'+str(mrgaddfila)]="mrg-"+data['e164'][0]['slc']                            # name
            sheet['k'+str(mrgaddfila)]=data['mrg'][0]                                     # members.member.0.deviceName
            if len(data['mrg']) > 1:
                sheet['l'+str(mrgaddfila)]=data['mrg'][1]                                     # members.member.1.deviceName
            if len(data['mrg']) > 2:
                sheet['m'+str(mrgaddfila)]=data['mrg'][2]                                        # members.member.2.deviceName
            if len(data['mrg']) > 3:
                sheet['n'+str(mrgaddfila)]=data['mrg'][3]                                        # members.member.3.deviceName
            if len(data['mrg']) > 4:
                sheet['o'+str(mrgaddfila)]=data['mrg'][4]                                        # members.member.4.deviceName
            if len(data['mrg']) > 5:
                sheet['p'+str(mrgaddfila)]=data['mrg'][5]                                        # members.member.5.deviceName

        sheet = blk["MRGL.ADD"]
        ## datos estaticos: Hierarchy, site,...
        sheet['B'+str(mrgaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(mrgaddfila)]="add"
        #sheet['D'+str(fila)]="name:"+row[3] # Search field
        ## datos dinamicos
        sheet['i'+str(mrgladdfil)]="mrgl-"+data['e164'][0]['slc']                            # name
        sheet['j'+str(mrgladdfil)]="0"                                        # members.member.0.order
        sheet['k'+str(mrgladdfil)]=fmoenvconfig['fmomrgmmoh']                                 # members.member.0.mediaResourceGroupName
        if len(data['mrg']) > 0:
            sheet['l'+str(mrgladdfil)]="1"                                        # members.member.1.order
            sheet['m'+str(mrgladdfil)]="mrg-"+data['e164'][0]['slc']                            # members.member.1.mediaResourceGroupName

        sheet = blk["TRUNK.ADD"]
        if data['gw'][0]['trunk'] != "0.0.0.0":
            ## SITE con GATEWAY configuramos el TRUNK
            print("(II) SITE con GW configuramos el GW",file=f)
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(trkaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(trkaddfila)]="add"
            sheet['G'+str(trkaddfila)]=fmoenvconfig['fmondl'] # $ndl
            #sheet['D'+str(trkaddfila)]="name:"+row[3] # Search field
            ## datos dinamicos
            sheet['I'+str(trkaddfila)]="SIP" # protocol
            sheet['k'+str(trkaddfila)]="Default" # callingLineIdPresentation
            sheet['l'+str(trkaddfila)]="false" # useCallerIdCallerNameinUriOutgoingRequest
            sheet['m'+str(trkaddfila)]="false" # acceptInboundRdnis
            sheet['n'+str(trkaddfila)]=fmotrunkincomingcss # callingSearchSpaceName
            sheet['o'+str(trkaddfila)]="Default" # useTrustedRelayPoint
            sheet['p'+str(trkaddfila)]="false" # enableQsigUtf8
            sheet['q'+str(trkaddfila)]="None(Default)" #
            sheet['r'+str(trkaddfila)]="false" # scriptTraceEnabled
            sheet['s'+str(trkaddfila)]="None" # tunneledProtocol
            sheet['t'+str(trkaddfila)]="When using both sRTP and TLS" # trunkTrafficSecure
            sheet['v'+str(trkaddfila)]=fmotrunkname ##### TRUNK NAME
            sheet['w'+str(trkaddfila)]="true" # retryVideoCallAsAudio
            sheet['x'+str(trkaddfila)]="Default" # sipPrivacy
            sheet['y'+str(trkaddfila)]="No Changes" # qsigVariant
            sheet['z'+str(trkaddfila)]="false" # mtpRequired
            sheet['ab'+str(trkaddfila)]="711ulaw" # tkSipCodec
            sheet['ac'+str(trkaddfila)]="Default" # calledPartyUnknownPrefix
            sheet['ae'+str(trkaddfila)]="99" # sigDigits
            sheet['af'+str(trkaddfila)]="true" # runOnEveryNode
            sheet['ag'+str(trkaddfila)]="true" # useDevicePoolCntdPnTransformationCss
            sheet['ai'+str(trkaddfila)]="true" # useDevicePoolCdpnTransformCss
            sheet['aj'+str(trkaddfila)]="Default" # connectedPartyIdPresentation
            sheet['al'+str(trkaddfila)]=fmoenvconfig['fmotrunksecurityprofile'] # securityProfileName
            sheet['am'+str(trkaddfila)]="true" # srtpFallbackAllowed
            sheet['an'+str(trkaddfila)]="Network" # protocolSide
            sheet['aq'+str(trkaddfila)]="Hub_None" # LocationName
            sheet['as'+str(trkaddfila)]="Default" # devicePoolName
            sheet['at'+str(trkaddfila)]="Default" # routeClassSignalling
            sheet['au'+str(trkaddfila)]="true" # useDevicePoolCalledCssUnkn
            sheet['aw'+str(trkaddfila)]="No Preference" # dtmfSignalingMethod
            sheet['ax'+str(trkaddfila)]=fmoenvconfig['fmotrunksipprofile'] # sipProfileName
            sheet['ay'+str(trkaddfila)]=data['gw'][0]['trunk'] # destinations.destination.0.addressIpv4
            sheet['az'+str(trkaddfila)]="5060" # destinations.destination.0.port
            sheet['ba'+str(trkaddfila)]="1" # destinations.destination.0.sortOrder
            sheet['bb'+str(trkaddfila)]="Disabled" # preemption
            sheet['bc'+str(trkaddfila)]="true" # pstnAccess
            sheet['bd'+str(trkaddfila)]="Originator" # callingPartySelection
            sheet['bf'+str(trkaddfila)]="Default" # unknownPrefix
            sheet['bg'+str(trkaddfila)]="false" # acceptOutboundRdnis
            sheet['bh'+str(trkaddfila)]="true" # useDevicePoolCgpnTransformCss
            sheet['bi'+str(trkaddfila)]="false" # pathReplacementSupport
            sheet['bj'+str(trkaddfila)]="false" # destAddrIsSrv
            sheet['bk'+str(trkaddfila)]="false" # traceFlag
            sheet['bn'+str(trkaddfila)]="Default" # callingname
            sheet['bp'+str(trkaddfila)]="Use System Default" # networkLocation
            sheet['bq'+str(trkaddfila)]="false" # unattendedPort
            sheet['bs'+str(trkaddfila)]="Default" # connectedNamePresentation
            sheet['bt'+str(trkaddfila)]="true" # useDevicePoolCgpnTransformCssUnkn
            sheet['bv'+str(trkaddfila)]="false" # useImePublicIpPort
            sheet['bw'+str(trkaddfila)]="Default" # sipAssertedType
            sheet['bz'+str(trkaddfila)]="0" # packetCaptureDuration
            sheet['ca'+str(trkaddfila)]="true" # isRpidEnabled
            sheet['cc'+str(trkaddfila)]="Off" # mlppIndicationStatus
            sheet['cd'+str(trkaddfila)]="None" # connectedNamePresentation
            sheet['ce'+str(trkaddfila)]="true" # isPaiEnabled
            sheet['cg'+str(trkaddfila)]="SIP Trunk" # product
            sheet['ci'+str(trkaddfila)]="false" # sendGeoLocation
            sheet['cj'+str(trkaddfila)]="No Changes" # asn1RoseOidEncoding
            sheet['ck'+str(trkaddfila)]="Trunk" # class
            sheet['cl'+str(trkaddfila)]="false" # srtpAllowed
            sheet['cm'+str(trkaddfila)]="Standard Presence group" # presenceGroupName
            sheet['cn'+str(trkaddfila)]="0" # recordingInformation
            sheet['cp'+str(trkaddfila)]="false" # transmitUtf8
            sheet['ct'+str(trkaddfila)]="Deliver DN only in connected party" # callingAndCalledPartyInfoFormat
            sheet['cu'+str(trkaddfila)]=fmoenvconfig['fmocucmmanagementip'] # networkDevice.nd

            sheet = blk["TRUNK.MOD"]
            ## datos estaticos: Hierarchy, site,...
            sheet['B'+str(trkaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(trkaddfila)]="modify"
            sheet['G'+str(trkaddfila)]=fmoenvconfig['fmondl']                    # $ndl
            sheet['D'+str(trkaddfila)]="name:"+fmotrunkname      # Search field
            ## datos dinamicos
            sheet['I'+str(trkaddfila)]="SIP"                     # protocol
            sheet['v'+str(trkaddfila)]=fmotrunkname              ##### TRUNK NAME
            sheet['aq'+str(trkaddfila)]=fmosite+"-Location"      # LocationName
            sheet['as'+str(trkaddfila)]=fmosite+"-DevicePool"    # devicePoolName

        sheet = blk["RG.ADD"]
        ## datos estaticos: Hierarchy, site,...
        sheet['B'+str(rdaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(rdaddfila)]="add"
        #sheet['D'+str(rdaddfila)]="name:"+row[3] # Search field
        ## datos dinamicos
        sheet['I'+str(rdaddfila)]="Circular"                 # distributionAlgorithm
        sheet['j'+str(rdaddfila)]="routegroup-"+siteslc      # name
        sheet['k'+str(rdaddfila)]="1"                        # members.member.0.deviceSelectionOrder
        if data['gw'][0]['trunk'] != "0.0.0.0":
            ## SITE con GATEWAY configuramos el TRUNK
            sheet['l'+str(rdaddfila)]=fmotrunkname           # members.member.0.deviceName
        else:
            print("(WW) SITE SIN GW configuramos el GW",file=f)
            print("(WW) Se configura fake-trunk",file=f)
            print("(WW) Es necesario configura Route Group con el Trunk correcto !!! ",file=f)
            sheet['l'+str(rdaddfila)]="fake-trunk"           # members.member.0.deviceName
        sheet['m'+str(rdaddfila)]="0"                        # members.member.0.port
        sheet['n'+str(rdaddfila)]=fmoenvconfig['fmocucmmanagementip']        # networkDevice.nd

        sheet = blk["TP"]
        ## Patron #1: Modificamos el patron intrasite [^50]XXX
        ## datos estaticos: Hierarchy, site,...
        #sheet['a'+str(tpaddfila)]="#"
        sheet['B'+str(tpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(tpaddfila)]="modify"
        sheet['D'+str(tpaddfila)]="routePartitionName:"+aintpt+","+"pattern:[^50]XXX"   # Search field
        ## datos dinamicos
        sheet['q'+str(tpaddfila)]=data['e164'][0]['patternintra']                       # pattern
        sheet['t'+str(tpaddfila)]="5"+data['e164'][0]['slc']                            # prefixDigitsOut
        sheet['y'+str(tpaddfila)]="XX CL2 type 2 Intrasite Routing AInt-PT"             # Description
        tpaddfila=tpaddfila+1

        sheet['a'+str(tpaddfila)]="##LBI"
        tpaddfila=tpaddfila+1

        ## TP: LBI
        ## datos estaticos: Hierarchy, site,...
        sheet['B'+str(tpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(tpaddfila)]="add"
        #sheet['D'+str(tpaddfila)]="name:"+row[3]                                       # Search field
        ## datos dinamicos
        sheet['I'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberingPlan
        sheet['j'+str(tpaddfila)]="Default"                                             # connectedLinePresentationBit
        sheet['k'+str(tpaddfila)]=lbipt                                                 # routePartitionName
        sheet['l'+str(tpaddfila)]="No Error"                                            # releaseClause
        sheet['m'+str(tpaddfila)]="false"                                               # blockEnable
        sheet['n'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberType
        sheet['o'+str(tpaddfila)]="false"                                               # provideOutsideDialtone
        sheet['q'+str(tpaddfila)]="5"+data['e164'][0]['slc']                            # pattern
        sheet['r'+str(tpaddfila)]="Default"                                             # patternPrecedence
        sheet['s'+str(tpaddfila)]=preisrcss                                             # callingSearchSpaceName
        sheet['t'+str(tpaddfila)]=""                                                    # prefixDigitsOut
        sheet['u'+str(tpaddfila)]="Translation"                                         # usage
        sheet['v'+str(tpaddfila)]="Cisco CallManager"                                   # calledPartyNumberingPlan
        sheet['w'+str(tpaddfila)]="true"                                                # dontWaitForIDTOnSubsequentHops
        sheet['x'+str(tpaddfila)]="Default"                                             # connectedNamePresentationBit
        if data['gw'][0]['trunk'] != "0.0.0.0":
            ## SITE con GATEWAY configuramos el TRUNK
            sheet['y'+str(tpaddfila)]="XX LBI 5XXXX"                                    # Description
        else:
            sheet['y'+str(tpaddfila)]="XX LBI 5XXXX -- NO-GW "
            print("(WW) Cambiar Partition CuXSiY-LBI-PT:","5"+data['e164'][0]['slc'] ,file=f)         # Description
        sheet['z'+str(tpaddfila)]="Default"                                             # routeClass
        sheet['aa'+str(tpaddfila)]="Default"                                            # callingNamePresentationBit
        sheet['ac'+str(tpaddfila)]="false"                                              # routeNextHopByCgpn
        sheet['ad'+str(tpaddfila)]="false"                                               # useOriginatorCss
        sheet['ae'+str(tpaddfila)]=""                                             # digitDiscardInstructionName
        sheet['ag'+str(tpaddfila)]="Off"                                                # useCallingPartyPhoneMask
        sheet['aj'+str(tpaddfila)]="Cisco CallManager"                                  # calledPartyNumberType
        sheet['al'+str(tpaddfila)]="false"                                              # patternUrgency
        sheet['am'+str(tpaddfila)]="Default"                                            # callingLinePresentationBit
        tpaddfila=tpaddfila+1

        ## datos estaticos: Hierarchy, site,...
        sheet['B'+str(tpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(tpaddfila)]="add"
        #sheet['D'+str(tpaddfila)]="name:"+row[3]                                       # Search field
        ## datos dinamicos
        sheet['I'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberingPlan
        sheet['j'+str(tpaddfila)]="Default"                                             # connectedLinePresentationBit
        sheet['k'+str(tpaddfila)]=lbipt                                                 # routePartitionName
        sheet['l'+str(tpaddfila)]="No Error"                                            # releaseClause
        sheet['m'+str(tpaddfila)]="false"                                               # blockEnable
        sheet['n'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberType
        sheet['o'+str(tpaddfila)]="false"                                               # provideOutsideDialtone
        sheet['q'+str(tpaddfila)]="5"+data['e164'][0]['slc']+"XXXX"                     # pattern
        sheet['r'+str(tpaddfila)]="Default"                                             # patternPrecedence
        sheet['s'+str(tpaddfila)]=preisrcss                                             # callingSearchSpaceName
        sheet['t'+str(tpaddfila)]=""                                                    # prefixDigitsOut
        sheet['u'+str(tpaddfila)]="Translation"                                         # usage
        sheet['v'+str(tpaddfila)]="Cisco CallManager"                                   # calledPartyNumberingPlan
        sheet['w'+str(tpaddfila)]="true"                                                # dontWaitForIDTOnSubsequentHops
        sheet['x'+str(tpaddfila)]="Default"                                             # connectedNamePresentationBit
        if data['gw'][0]['trunk'] != "0.0.0.0":
            ## SITE con GATEWAY configuramos el TRUNK
            sheet['y'+str(tpaddfila)]="XX LBI 5XXXXXXXX"                                    # Description
        else:
            sheet['y'+str(tpaddfila)]="XX LBI 5XXXXXXXX -- NO-GW "
            print("(WW) Cambiar Partition CuXSiY-LBI-PT:","5"+data['e164'][0]['slc']+"XXXX" ,file=f)         # Description
        sheet['z'+str(tpaddfila)]="Default"                                             # routeClass
        sheet['aa'+str(tpaddfila)]="Default"                                            # callingNamePresentationBit
        sheet['ac'+str(tpaddfila)]="false"                                              # routeNextHopByCgpn
        sheet['ad'+str(tpaddfila)]="false"                                               # useOriginatorCss
        sheet['ae'+str(tpaddfila)]=""                                             # digitDiscardInstructionName
        sheet['ag'+str(tpaddfila)]="Off"                                                # useCallingPartyPhoneMask
        sheet['aj'+str(tpaddfila)]="Cisco CallManager"                                  # calledPartyNumberType
        sheet['al'+str(tpaddfila)]="false"                                              # patternUrgency
        sheet['am'+str(tpaddfila)]="Default"                                            # callingLinePresentationBit
        tpaddfila=tpaddfila+1

        ## datos estaticos: Hierarchy, site,...
        sheet['B'+str(tpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(tpaddfila)]="add"
        #sheet['D'+str(tpaddfila)]="name:"+row[3]                                       # Search field
        ## datos dinamicos
        sheet['I'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberingPlan
        sheet['j'+str(tpaddfila)]="Default"                                             # connectedLinePresentationBit
        sheet['k'+str(tpaddfila)]=lbipt                                                 # routePartitionName
        sheet['l'+str(tpaddfila)]="No Error"                                            # releaseClause
        sheet['m'+str(tpaddfila)]="false"                                               # blockEnable
        sheet['n'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberType
        sheet['o'+str(tpaddfila)]="false"                                               # provideOutsideDialtone
        sheet['q'+str(tpaddfila)]="5"+data['e164'][0]['slc']+"*XXX"                     # pattern
        sheet['r'+str(tpaddfila)]="Default"                                             # patternPrecedence
        sheet['s'+str(tpaddfila)]=preisrcss                                             # callingSearchSpaceName
        sheet['t'+str(tpaddfila)]=""                                                    # prefixDigitsOut
        sheet['u'+str(tpaddfila)]="Translation"                                         # usage
        sheet['v'+str(tpaddfila)]="Cisco CallManager"                                   # calledPartyNumberingPlan
        sheet['w'+str(tpaddfila)]="true"                                                # dontWaitForIDTOnSubsequentHops
        sheet['x'+str(tpaddfila)]="Default"                                             # connectedNamePresentationBit
        sheet['y'+str(tpaddfila)]="XX LBI 5XXXX*XXX"                                    # Description
        if data['gw'][0]['trunk'] != "0.0.0.0":
            ## SITE con GATEWAY configuramos el TRUNK
            sheet['y'+str(tpaddfila)]="XX LBI 5XXXX*XXX"                                    # Description
        else:
            sheet['y'+str(tpaddfila)]="XX LBI 5XXXX*XXX -- NO-GW "
            print("(WW) Cambiar Partition CuXSiY-LBI-PT:","5"+data['e164'][0]['slc']+"*XXX" ,file=f)         # Description
        sheet['z'+str(tpaddfila)]="Default"                                             # routeClass
        sheet['aa'+str(tpaddfila)]="Default"                                            # callingNamePresentationBit
        sheet['ac'+str(tpaddfila)]="false"                                              # routeNextHopByCgpn
        sheet['ad'+str(tpaddfila)]="false"                                               # useOriginatorCss
        sheet['ae'+str(tpaddfila)]=""                                             # digitDiscardInstructionName
        sheet['ag'+str(tpaddfila)]="Off"                                                # useCallingPartyPhoneMask
        sheet['aj'+str(tpaddfila)]="Cisco CallManager"                                  # calledPartyNumberType
        sheet['al'+str(tpaddfila)]="false"                                              # patternUrgency
        sheet['am'+str(tpaddfila)]="Default"                                            # callingLinePresentationBit
        tpaddfila=tpaddfila+1

        #### TBP (AC distinto 11)
        ## datos estaticos: Hierarchy, site,...
        if data['e164'][0]['ac'] != '11':
            sheet['B'+str(tpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(tpaddfila)]="add"
            #sheet['D'+str(tpaddfila)]="name:"+row[3]                                       # Search field
            ## datos dinamicos
            sheet['I'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberingPlan
            sheet['j'+str(tpaddfila)]="Default"                                             # connectedLinePresentationBit
            sheet['k'+str(tpaddfila)]=localpt                                                 # routePartitionName
            sheet['l'+str(tpaddfila)]="No Error"                                            # releaseClause
            sheet['m'+str(tpaddfila)]="false"                                               # blockEnable
            sheet['n'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberType
            sheet['o'+str(tpaddfila)]="false"                                               # provideOutsideDialtone
            sheet['q'+str(tpaddfila)]=tbsdddnacional                     # pattern
            sheet['r'+str(tpaddfila)]="Default"                                             # patternPrecedence
            sheet['s'+str(tpaddfila)]=ilscss                                             # callingSearchSpaceName
            sheet['t'+str(tpaddfila)]="+55"                                                    # prefixDigitsOut
            sheet['u'+str(tpaddfila)]="Translation"                                         # usage
            sheet['v'+str(tpaddfila)]="Cisco CallManager"                                   # calledPartyNumberingPlan
            sheet['w'+str(tpaddfila)]="true"                                                # dontWaitForIDTOnSubsequentHops
            sheet['x'+str(tpaddfila)]="Default"                                             # connectedNamePresentationBit
            sheet['y'+str(tpaddfila)]="*** TBP *** DDD National (auto)"                                    # Description
            sheet['z'+str(tpaddfila)]="Default"                                             # routeClass
            sheet['aa'+str(tpaddfila)]="Default"                                            # callingNamePresentationBit
            sheet['ac'+str(tpaddfila)]="false"                                              # routeNextHopByCgpn
            sheet['ad'+str(tpaddfila)]="false"                                              # useOriginatorCss
            sheet['ae'+str(tpaddfila)]="PreDot"                                             # digitDiscardInstructionName
            sheet['ag'+str(tpaddfila)]="Off"                                                # useCallingPartyPhoneMask
            sheet['aj'+str(tpaddfila)]="Cisco CallManager"                                  # calledPartyNumberType
            sheet['al'+str(tpaddfila)]="false"                                              # patternUrgency
            sheet['am'+str(tpaddfila)]="Default"                                            # callingLinePresentationBit
            tpaddfila=tpaddfila+1

            sheet['B'+str(tpaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
            sheet['C'+str(tpaddfila)]="add"
            #sheet['D'+str(tpaddfila)]="name:"+row[3]                                       # Search field
            ## datos dinamicos
            sheet['I'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberingPlan
            sheet['j'+str(tpaddfila)]="Default"                                             # connectedLinePresentationBit
            sheet['k'+str(tpaddfila)]=localpt                                                 # routePartitionName
            sheet['l'+str(tpaddfila)]="No Error"                                            # releaseClause
            sheet['m'+str(tpaddfila)]="false"                                               # blockEnable
            sheet['n'+str(tpaddfila)]="Cisco CallManager"                                   # callingPartyNumberType
            sheet['o'+str(tpaddfila)]="false"                                               # provideOutsideDialtone
            sheet['q'+str(tpaddfila)]=tbsdddmobile                     # pattern
            sheet['r'+str(tpaddfila)]="Default"                                             # patternPrecedence
            sheet['s'+str(tpaddfila)]=ilscss                                             # callingSearchSpaceName
            sheet['t'+str(tpaddfila)]="+55"                                                    # prefixDigitsOut
            sheet['u'+str(tpaddfila)]="Translation"                                         # usage
            sheet['v'+str(tpaddfila)]="Cisco CallManager"                                   # calledPartyNumberingPlan
            sheet['w'+str(tpaddfila)]="true"                                                # dontWaitForIDTOnSubsequentHops
            sheet['x'+str(tpaddfila)]="Default"                                             # connectedNamePresentationBit
            sheet['y'+str(tpaddfila)]="*** TBP *** DDD Mobile (auto)"                                    # Description
            sheet['z'+str(tpaddfila)]="Default"                                             # routeClass
            sheet['aa'+str(tpaddfila)]="Default"                                            # callingNamePresentationBit
            sheet['ac'+str(tpaddfila)]="false"                                              # routeNextHopByCgpn
            sheet['ad'+str(tpaddfila)]="false"                                              # useOriginatorCss
            sheet['ae'+str(tpaddfila)]="PreDot"                                             # digitDiscardInstructionName
            sheet['ag'+str(tpaddfila)]="Off"                                                # useCallingPartyPhoneMask
            sheet['aj'+str(tpaddfila)]="Cisco CallManager"                                  # calledPartyNumberType
            sheet['al'+str(tpaddfila)]="false"                                              # patternUrgency
            sheet['am'+str(tpaddfila)]="Default"                                            # callingLinePresentationBit
            tpaddfila=tpaddfila+1

        sheet = blk["SRST.ADD"]
        ## datos estaticos: Hierarchy, site,...
        sheet['B'+str(srstaddfil)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(srstaddfil)]="add"
        #sheet['D'+str(srstaddfil)]="name:"+row[3] # Search field
        ## datos dinamicos
        sheet['i'+str(srstaddfil)]="false"                 # isSecure
        sheet['j'+str(srstaddfil)]="srst-"+siteslc            # name
        sheet['k'+str(srstaddfil)]=data['srst'][0]['ipsccp']             # SipNetwork
        sheet['l'+str(srstaddfil)]=data['srst'][0]['ipsccp']             # ipAddress
        sheet['m'+str(srstaddfil)]="2000"                     # port
        sheet['n'+str(srstaddfil)]="5060"                     # sipPort

        #######################################################
        ## Buscamos el location del DP
        n=0
        while data['loc'][n]['location'] != dp['location']:
            n=n+1
            #print(">>>",data['loc'][n]['audio'])
            #print(">>>",data['loc'][n]['video'])

        #print(">>>",n,">>>",dp['location'],">>> >>>",data['loc'][n]['location'])

        sheet = blk["LOCATION.MOD"]
        ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
        sheet['B'+str(addlocfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(addlocfila)]="modify"
        sheet['D'+str(addlocfila)]="name:"+fmosite+"-Location"           # Search field
        ## datos dinamicos
        sheet['N'+str(addlocfila)]=fmosite+"-Location"                    # Name
        sheet['I'+str(addlocfila)]="Hub_None"                             # betweenLocations.betweenLocation.0.locationName
        sheet['J'+str(addlocfila)]="-1"                                   # betweenLocations.betweenLocation.0.immersiveBandwidth
        sheet['K'+str(addlocfila)]=data['loc'][n]['video']                # betweenLocations.betweenLocation.0.videoBandwidth
        sheet['L'+str(addlocfila)]="50"                                   # betweenLocations.betweenLocation.0.weight
        sheet['M'+str(addlocfila)]=data['loc'][n]['audio']                # betweenLocations.betweenLocation.0.audioBandwidth
        sheet['O'+str(addlocfila)]=fmosite+"-Location"                    # betweenLocations.betweenLocation.0.locationName
        sheet['P'+str(addlocfila)]="No Reservation"                       # betweenLocations.betweenLocation.0.immersiveBandwidth
        sheet['Q'+str(addlocfila)]="0"                                    # betweenLocations.betweenLocation.0.videoBandwidth
        sheet['R'+str(addlocfila)]="0"                                    # betweenLocations.betweenLocation.0.weight
        sheet['S'+str(addlocfila)]="0"                                    # betweenLocations.betweenLocation.0.audioBandwidth

        sheet = blk["REGION.MOD"]
        ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
        sheet['B'+str(addregfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(addregfila)]="modify"
        sheet['D'+str(addregfila)]="name:"+fmosite+"-Region"               # Search field
        ## datos dinamicos
        sheet['I'+str(addregfila)]=fmosite+"-Region"                      # Name
        sheet['J'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.0.lossyNetwork
        sheet['K'+str(addregfila)]="-1"                                   # relatedRegions.relatedRegion.0.immersiveVideoBandwidth
        sheet['L'+str(addregfila)]="64"                                    # relatedRegions.relatedRegion.0.bandwidth
        sheet['M'+str(addregfila)]="384"                                   # relatedRegions.relatedRegion.0.videoBandwidth
        sheet['N'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.0.codecPreference
        sheet['o'+str(addregfila)]=fmosite+"-Region"                      # relatedRegions.relatedRegion.0.regionName
        sheet['p'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.1.lossyNetwork
        sheet['q'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.1.immersiveVideoBandwidth
        sheet['r'+str(addregfila)]="8"                                    # relatedRegions.relatedRegion.1.bandwidth
        sheet['s'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.1.videoBandwidth
        sheet['t'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.1.codecPreference
        sheet['u'+str(addregfila)]="Default"                              # relatedRegions.relatedRegion.2.regionName
        ##sheet['v'+str(addregfila)]=fmoenvconfig['fmocucmmanagementip']                  # networkDevice.nd
        sheet['v'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.2.lossyNetwork
        sheet['w'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.2.immersiveVideoBandwidth
        sheet['x'+str(addregfila)]="64"                                    # relatedRegions.relatedRegion.2.bandwidth
        sheet['y'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.2.videoBandwidth
        sheet['z'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.2.codecPreference
        sheet['aa'+str(addregfila)]=fmoenvconfig['fmomohregion']                           # relatedRegions.relatedRegion.2.regionName

        sheet = blk["DP.MOD"]
        ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
        sheet['B'+str(adddpfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(adddpfila)]="modify"
        sheet['D'+str(addlocfila)]="name:"+fmosite+"-DevicePool" # Search field
        ## datos dinamicos
        sheet['at'+str(adddpfila)]=fmosite+"-DevicePool"                  # regionName
        sheet['q'+str(adddpfila)]=fmosite+"-Region"                       # regionName
        sheet['az'+str(adddpfila)]=fmosite+"-Location"                    # regionName
        sheet['s'+str(adddpfila)]=cmg                                     # CMGName
        #sheet['t'+str(adddpfila)]="Default"                               # calledPartyUnknownPrefix
        #sheet['u'+str(adddpfila)]="Default"                               # singleButtonBarge
        #sheet['w'+str(adddpfila)]="Default"                               # calledPartyNationalPrefix
        #sheet['x'+str(adddpfila)]="Default"                               # calledPartySubscriberPrefix
        sheet['z'+str(adddpfila)]=fmoenvconfig['fmonetworklocale']                        # networkLocale
        sheet['ac'+str(adddpfila)]="srst-"+siteslc                   # srstName
        #sheet['ae'+str(adddpfila)]="Default"                               # callingPartySubscriberPrefix
        #sheet['ag'+str(adddpfila)]="Default"                              # revertPriority
        #sheet['aj'+str(adddpfila)]="Default"                             # calledPartyInternationalPrefix
        sheet['ak'+str(adddpfila)]=data['dp'][0]['uso']                        # dateTimeSettingName
        #sheet['am'+str(adddpfila)]="Default"                             # joinAcrossLines
        #sheet['aq'+str(adddpfila)]="Default"                             # callingPartyNationalPrefix
        #sheet['ar'+str(adddpfila)]="Default"                             # callingPartyInternationalPrefixh
        sheet['ay'+str(adddpfila)]=fmoenvconfig['fmoaargroup']                              # aarNeighborhoodName
        sheet['bb'+str(adddpfila)]="mrgl-"+siteslc                       # mediaResourceListName
        #sheet['bg'+str(adddpfila)]="Default"                             # callingPartyUnknownPrefix
        ##
        sheet['bi'+str(adddpfila)]="SLRG-Intl"                          # localRouteGroup.0.name
        sheet['bj'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bk'+str(adddpfila)]="SLRG-Natl"                          # localRouteGroup.0.name
        sheet['bl'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bm'+str(adddpfila)]="SLRG-Emer"                          # localRouteGroup.0.name
        sheet['bn'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bo'+str(adddpfila)]="SLRG-FPHN"                          # localRouteGroup.0.name
        sheet['bp'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bq'+str(adddpfila)]="SLRG-Mobl"                          # localRouteGroup.0.name
        sheet['br'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bs'+str(adddpfila)]="SLRG-Local"                         # localRouteGroup.0.name
        sheet['bt'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bu'+str(adddpfila)]="SLRG-Oper"                          # localRouteGroup.0.name
        sheet['bv'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bw'+str(adddpfila)]="SLRG-PCSN"                          # localRouteGroup.0.name
        sheet['bx'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['by'+str(adddpfila)]="SLRG-PRSN"                          # localRouteGroup.0.name
        sheet['bz'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['ca'+str(adddpfila)]="SLRG-SRSN"                          # localRouteGroup.0.name
        sheet['cb'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['cc'+str(adddpfila)]="SLRG-Serv"                          # localRouteGroup.0.name
        sheet['cd'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['ce'+str(adddpfila)]="Standard Local Route Group"         # localRouteGroup.0.name
        sheet['cf'+str(adddpfila)]="routegroup-"+siteslc

        sheet = blk["SITE.DEF"]
        ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
        sheet['B'+str(sitaddfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(sitaddfila)]="modify"
        sheet['D'+str(sitaddfila)]="name:"+fmositename # Search field
        ## datos dinamicos
        sheet['y'+str(sitaddfila)]=fmosite+"-Location"     # defaultLOC
        sheet['ag'+str(sitaddfila)]=fmosite+"-DevicePool"  # defaultDP
        sheet['aw'+str(sitaddfila)]=fmositename            # name
        sheet['ba'+str(sitaddfila)]=cmg                    # defaultcucmgroup


    #elif dp['devicepool'].startswith(cmodevicepool): ## Se ejecuta varias veces
    else: #dp['devicepool'].startswith(cmodevicepool): ## Se ejecuta varias veces
        ## BUSCAMOS Location
        n=0

        #print(">>>",n,">>>",dp['location'],">>> >>>",data['loc'][n]['location'])
        #print(">>>",len(data['loc']),"<<<")
        maxLocation=len(data['loc'])

        while data['loc'][n]['location'] != dp['location']:
            n=n+1
            #print(">>>",data['loc'][n]['audio'])
            #print(">>>",data['loc'][n]['video'])
            ##
            ## Si no se encuentra el LOCATION tenemos un error de configuraion
            if n >= maxLocation:
                print("(EE) Error LOCATION mal configurado en CMO:: ")
                print("(EE) CMO DP=",dp['devicepool'])
                print("(EE) -- No se encuentra en ")
                for locloc in data['loc']:
                    print("(EE) CMO Location=",locloc['location'])
                print("(EE) -- Se configura con Default Location ")
                break
        ## Se busca el indice del Default Location XXXX-LOC
        n=0
        while data['loc'][n]['location'] != cmolocation:
            n=n+1

        #print(">>>",n,">>>",dp['location'],">>> >>>",data['loc'][n]['location'])

        #print(dp['devicepool'].replace(cmodevicepool,fmosite))
        #################### LOCATION.ADD ####################
        sheet = blk["LOCATION.ADD"]
        ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
        sheet['B'+str(addlocfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(addlocfila)]="add"
        #sheet['D'+str(addlocfila)]="name:"+row[3] # Search field
        ## datos dinamicos
        sheet['N'+str(addlocfila)]=dp['devicepool'].replace(cmodevicepool,fmosite+"-Location")  # Name
        sheet['I'+str(addlocfila)]="Hub_None"                            # betweenLocations.betweenLocation.0.locationName
        sheet['J'+str(addlocfila)]="-1"                                   # betweenLocations.betweenLocation.0.immersiveBandwidth
        sheet['K'+str(addlocfila)]=data['loc'][n]['video']                # betweenLocations.betweenLocation.0.videoBandwidth
        sheet['L'+str(addlocfila)]="50"                                   # betweenLocations.betweenLocation.0.weight
        sheet['M'+str(addlocfila)]=data['loc'][n]['audio']                # betweenLocations.betweenLocation.0.audioBandwidth
        sheet['O'+str(addlocfila)]=dp['devicepool'].replace(cmodevicepool,fmosite+"-Location")  # betweenLocations.betweenLocation.0.locationName
        sheet['P'+str(addlocfila)]="No Reservation"                       # betweenLocations.betweenLocation.0.immersiveBandwidth
        sheet['Q'+str(addlocfila)]="0"                                    # betweenLocations.betweenLocation.0.videoBandwidth
        sheet['R'+str(addlocfila)]="0"                                    # betweenLocations.betweenLocation.0.weight
        sheet['S'+str(addlocfila)]="0"                                    # betweenLocations.betweenLocation.0.audioBandwidth
        addlocfila=addlocfila+1

        #################### REGION.ADD ####################
        sheet = blk["REGION.ADD"]
        ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
        sheet['B'+str(addregfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(addregfila)]="add"
        #sheet['D'+str(addregfila)]="name:"+row[3] # Search field
        ## datos dinamicos
        sheet['I'+str(addregfila)]=dp['devicepool'].replace(cmodevicepool,fmosite+"-Region")  # Name
        sheet['J'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.0.lossyNetwork
        sheet['K'+str(addregfila)]="-1"                                   # relatedRegions.relatedRegion.0.immersiveVideoBandwidth
        sheet['L'+str(addregfila)]="8"                                    # relatedRegions.relatedRegion.0.bandwidth
        sheet['M'+str(addregfila)]="-1"                                   # relatedRegions.relatedRegion.0.videoBandwidth
        sheet['N'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.0.codecPreference
        sheet['o'+str(addregfila)]=dp['devicepool'].replace(cmodevicepool,fmosite+"-Region")  # relatedRegions.relatedRegion.0.regionName
        sheet['p'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.1.lossyNetwork
        sheet['q'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.1.immersiveVideoBandwidth
        sheet['r'+str(addregfila)]="8"                                    # relatedRegions.relatedRegion.1.bandwidth
        sheet['s'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.1.videoBandwidth
        sheet['t'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.1.codecPreference
        sheet['u'+str(addregfila)]="Default"                              # relatedRegions.relatedRegion.1.regionName
        sheet['v'+str(addregfila)]=fmoenvconfig['fmocucmmanagementip']                                    # networkDevice.nd
        sheet['w'+str(addregfila)]="Use System Default"                                    # relatedRegions.relatedRegion.2.lossyNetwork
        sheet['x'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.2.immersiveVideoBandwidth
        sheet['y'+str(addregfila)]="64"                                    # relatedRegions.relatedRegion.2.bandwidth
        sheet['z'+str(addregfila)]="-1"                                    # relatedRegions.relatedRegion.2.videoBandwidth
        sheet['aa'+str(addregfila)]="Use System Default"                   # relatedRegions.relatedRegion.2.codecPreference
        sheet['ab'+str(addregfila)]=fmoenvconfig['fmomohregion']                           # relatedRegions.relatedRegion.2.regionName
        addregfila=addregfila+1

        #################### DP.ADD ####################
        sheet = blk["DP.ADD"]
        ## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
        sheet['B'+str(adddpfila)]=fmoenvconfig['hierarchynode']+"."+fmositename
        sheet['C'+str(adddpfila)]="add"
        #sheet['D'+str(addlocfila)]="name:"+row[3] # Search field
        ## datos dinamicos
        sheet['at'+str(adddpfila)]=dp['devicepool'].replace(cmodevicepool,fmosite+"-DevicePool")  # devicePool Name
        sheet['q'+str(adddpfila)]=dp['devicepool'].replace(cmodevicepool,fmosite+"-Region")  # regionName
        sheet['az'+str(adddpfila)]=dp['devicepool'].replace(cmodevicepool,fmosite+"-Location")  # regionName
        sheet['s'+str(adddpfila)]=cmg                                    # CMGName
        sheet['t'+str(adddpfila)]="Default"                              # calledPartyUnknownPrefix
        sheet['u'+str(adddpfila)]="Default"                              # singleButtonBarge
        sheet['w'+str(adddpfila)]="Default"                              # calledPartyNationalPrefix
        sheet['x'+str(adddpfila)]="Default"                              # calledPartySubscriberPrefix
        sheet['z'+str(adddpfila)]=fmoenvconfig['fmonetworklocale']                       # networkLocale
        sheet['ac'+str(adddpfila)]="srst-"+siteslc                        # srstName
        sheet['ae'+str(adddpfila)]="Default"                             # callingPartySubscriberPrefix
        sheet['ag'+str(adddpfila)]="Default"                             # revertPriority
        sheet['aj'+str(adddpfila)]="Default"                             # calledPartyInternationalPrefix
        sheet['ak'+str(adddpfila)]=dp['uso']                        # dateTimeSettingName
        sheet['am'+str(adddpfila)]="Default"                             # joinAcrossLines
        sheet['aq'+str(adddpfila)]="Default"                             # callingPartyNationalPrefix
        sheet['ar'+str(adddpfila)]="Default"                             # callingPartyInternationalPrefixh
        sheet['ay'+str(adddpfila)]=fmoenvconfig['fmoaargroup']                              # aarNeighborhoodName
        sheet['bb'+str(adddpfila)]="mrgl-"+siteslc                       # mediaResourceListName
        sheet['bg'+str(adddpfila)]="Default"                             # callingPartyUnknownPrefix
        ##
        sheet['bi'+str(adddpfila)]="SLRG-Intl"                          # localRouteGroup.0.name
        sheet['bj'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bk'+str(adddpfila)]="SLRG-Natl"                          # localRouteGroup.0.name
        sheet['bl'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bm'+str(adddpfila)]="SLRG-Emer"                          # localRouteGroup.0.name
        sheet['bn'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bo'+str(adddpfila)]="SLRG-FPHN"                          # localRouteGroup.0.name
        sheet['bp'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bq'+str(adddpfila)]="SLRG-Mobl"                          # localRouteGroup.0.name
        sheet['br'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bs'+str(adddpfila)]="SLRG-Local"                         # localRouteGroup.0.name
        sheet['bt'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bu'+str(adddpfila)]="SLRG-Oper"                          # localRouteGroup.0.name
        sheet['bv'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['bw'+str(adddpfila)]="SLRG-PCSN"                          # localRouteGroup.0.name
        sheet['bx'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['by'+str(adddpfila)]="SLRG-PRSN"                          # localRouteGroup.0.name
        sheet['bz'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['ca'+str(adddpfila)]="SLRG-SRSN"                          # localRouteGroup.0.name
        sheet['cb'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['cc'+str(adddpfila)]="SLRG-Serv"                          # localRouteGroup.0.name
        sheet['cd'+str(adddpfila)]="routegroup-"+siteslc                          #
        sheet['ce'+str(adddpfila)]="Standard Local Route Group"         # localRouteGroup.0.name
        sheet['cf'+str(adddpfila)]="routegroup-"+siteslc
        adddpfila=adddpfila+1

        # Next row
        fila=fila+1

## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)

## LOG de CONFIGURACION
f.close()

exit(0)
