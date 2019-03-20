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

if len(sys.argv) != 5:
    print("ERROR: python3 <nombre-fichero.py> <static-data-file> <site-data-file> XXXX")
    print(">>>>  [1] static-data-file:  nombre de fichero donde leemos los valores estáticos")
    print(">>>>  [2] site-data-file:    nombre de fichero donde leemos los datos del site")
    print(">>>>  [3] XXXX: SLC")
    print(">>>>  [4] Cluster-path")
    exit(1)
else:
    #print(">>>>  Datos de entorno: ",sys.argv[1])
    #print(">>>>  Datos del site:   ",sys.argv[2])
    #print(">>>>  SLC:              ",sys.argv[3])
    #print(">>>>  Area code:        ",sys.argv[4])
    fmostaticdata=sys.argv[1]
    fmositedata=sys.argv[2]
    siteslc=str(sys.argv[3])   #INPUTS: SLC
    clusterpath=sys.argv[4]    #SLC Cluster path
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
#print("<<<<>>>>>>   ",fmoenvconfig['fmocucmmanagementip'])

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
#inputfile = clusterpath+"/directorynumber.csv"  ## ORIGINAL
inputfile = clusterpath+"/phone.mod1.csv"   ## Modificado

templateblkfile = "blk/04.phone-template.xlsx" # SIN DATAINPUT
outputblkfile = sitepath+"/05.phone."+siteslc+".xlsx"

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
linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
aarcss=customerid+"-AAR-CSS"
devicepool=cucdmsite+"-DevicePool"
location=cucdmsite+"-Location"
devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
subscribecss=cucdmsite+"-InternalOnly-CSS"

# CMO File INPUT DATA
fgw = open(inputfile,"r")
csv_f = csv.reader(fgw)

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

# FMO commands:
action="add"

fila=7
filadn=7
delta=0

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        if row[4].startswith(cmodevicepool) and row[2].startswith('SEP'):
            dn=1 ## Consideramos que siempre hay al menos una línea
            delta=84*(dn-1) #L1 - 1
            ######################################################
            ## USER PHONE
            ######################################################
            sheet =  blk["USER"]
            sheet['B'+str(fila)]=hierarchynode+"."+fmositename
            sheet['C'+str(fila)]=action
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ######################################################
            sheet['i'+str(fila)]="p"+row[129]           # username
            sheet['j'+str(fila)]="p"+row[129]           # ps.username
            sheet['n'+str(fila)]="p"+row[129]           # rbac.username
            sheet['o'+str(fila)]="PHONE"            #rbac.first_name
            sheet['p'+str(fila)]="DUMMY"            #rbac.last_name
            sheet['q'+str(fila)]=""                 #rbac.language
            sheet['s'+str(fila)]=fmositename+"SelfService"                 #rbac.role
            sheet['u'+str(fila)]="CUCM Local"       #userType
            sheet['v'+str(fila)]="PHONE"            #sn.0
            sheet['w'+str(fila)]=""      # givenName.0
            sheet['x'+str(fila)]="Standard CCM End Users"       # memberOf.0
            sheet['y'+str(fila)]="false"       # rbac.account_information.disabled
            ## User must change the passowor on next loggin
            #sheet['z'+str(fila)]="true"       # rbac.account_information.change_password_on_login

            ## DEBUG
            print("PH#",fila,row[2]," ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164])

            ######################################################
            ## PHONE
            ######################################################
            sheet =  blk["PHONE"]
            sheet['B'+str(fila)]=hierarchynode+"."+fmositename
            sheet['C'+str(fila)]=action
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ######################################################
            sheet['i'+str(fila)]=""                         # directoryUrl
            sheet['j'+str(fila)]=row[54]                    # protocol
            sheet['k'+str(fila)]=""                         # secureInformationUrl
            sheet['l'+str(fila)]="false"                    # requireDtmfReception
            if row[5] != "":
                sheet['m'+str(fila)]=cl+"-"+row[5]          #phoneTemplateName
            #sheet['n'+str(fila)]=                          #callingSearchSpaceName (DEVICE CSS)
            sheet['p'+str(fila)]="Default"                  #useTrustedRelayPoint
            sheet['q'+str(fila)]="Brazil"                   #networkLocale
            sheet['r'+str(fila)]="Default"                  #ringSettingBusyBlfAudibleAlert
            sheet['t'+str(fila)]="Portuguese Brazil"        #userLocale
            sheet['u'+str(fila)]="Default"                  #deviceMobilityMode
            sheet['w'+str(fila)]="No Rollover"              # outboundCallRollover
            sheet['x'+str(fila)]=""                         # ip_address
            sheet['y'+str(fila)]=""                         # primaryPhoneName
            sheet['z'+str(fila)]=row[2]                     # name
            ####################
            sheet['aa'+str(fila)]="true"                    # retryVideoCallAsAudio
            sheet['ab'+str(fila)]="Default"                 # callInfoPrivacyStatus
            sheet['ad'+str(fila)]="Default"                 # phoneServiceDisplay
            sheet['af'+str(fila)]="Default"                 # alwaysUsePrimeLineForVoiceMessage
            sheet['ag'+str(fila)]="false"                   # mtprequired
            sheet['ak'+str(fila)]="false"                   # isProtected
            sheet['al'+str(fila)]=subscribecss              # subscribeCallingSearchSpaceName
            sheet['am'+str(fila)]="Default"                 # preemption
            sheet['an'+str(fila)]="false"                   #rfc2833Disabled
            sheet['ao'+str(fila)]="On"                      #hlogStatus
            sheet['ap'+str(fila)]="true"                    #isActive
            sheet['aq'+str(fila)]="false"                   #hotlineDevice
            sheet['at'+str(fila)]="User"                    #protocolSide
            sheet['av'+str(fila)]=aargroup                  #aarNeighborhoodName
            sheet['aw'+str(fila)]=location                  #locationName
            sheet['ay'+str(fila)]=devicepool                #devicepool
            sheet['az'+str(fila)]="false"                   #isDualMode
            ####################
            sheet['cl'+str(fila)]="No Pending Operation"    # certificateOperation
            sheet['co'+str(fila)]="false"                   #enableCallRoutingToRdWhenNoneIsActive
            #sheet['cp'+str(fila)]=row[63]                   #sipProfileName
            if row[54] == "SIP":
                sheet['cp'+str(fila)]="Standard SIP Profile"    #sipProfileName
            sheet['cq'+str(fila)]="false"                   #allowCtiControlFlag    @@@@ TRUE para grabación
            sheet['cr'+str(fila)]="Default"                 #singleButtonBarge
            sheet['ct'+str(fila)]="true"                    #enableExtensionMobility
            sheet['cy'+str(fila)]="true"                    #useDevicePoolCgpnTransformCss
            sheet['cz'+str(fila)]="false"                   #traceFlag
            sheet['da'+str(fila)]="Default"                 #phoneSuite
            sheet['dc'+str(fila)]=row[46]                   #securityProfileName
            sheet['dd'+str(fila)]="Off"                     #joinAcrossLines
            sheet['df'+str(fila)]="false"                   #dndStatus
            sheet['di'+str(fila)]="Use System Default"      # networkLocation
            sheet['dj'+str(fila)]="false"                   # unattendedPort
            ####################
            sheet['dl'+str(fila)]=row[129]                  #  mobilityUserIdName
            #sheet['cm'+str(fila)]="" # versionStamp
            sheet['dn'+str(fila)]="Not Trusted"             #deviceTrustMode
            ############################################
            ##
            ## PHONE: LINE #1
            ##
            sheet['do'+str(fila)]=row[180+delta]            #lines.line.0.displayAscii
            sheet['dp'+str(fila)]=""                        #lines.line.0.associatedEndusers.enduser.0.userId  @@@@@@ p+EXTENSION
            sheet['dq'+str(fila)]="Ring"                    #lines.line.0.ringSetting
            sheet['dr'+str(fila)]="Use System Default"      #lines.line.0.consecutiveRingSetting
            sheet['dr'+str(fila)]="Default"                 #lines.line.0.recordingProfileName
            sheet['dt'+str(fila)]=dn                        #lines.line.0.index
            sheet['du'+str(fila)]="Use System Default"      # lines.line.0.ringSettingActivePickupAlert
            sheet['dv'+str(fila)]=row[163+delta]            #lines.line.0.label
            sheet['dw'+str(fila)]="Gateway Preferred"       #lines.line.0.recordingMediaSource
            sheet['dx'+str(fila)]=row[165+delta]            #lines.line.0.maxNumCalls
            sheet['dy'+str(fila)]="General"                 #lines.line.0.partitionUsage
            sheet['dz'+str(fila)]="Call Recording Disabled" #lines.line.0.recordingMediaSource
            sheet['ec'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]            #lines.line.0.e164Mask
            sheet['ed'+str(fila)]="true"                    #lines.line.0.missedCallLogging
            sheet['ee'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.dialedNumber
            sheet['ef'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.redirectedNumber
            sheet['eg'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.callerName
            sheet['eh'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.callerNumber
            sheet['ei'+str(fila)]=row[129+delta]            #lines.line.0.dirn.pattern

            ## CSS depende de la PT
            if row[130] == "Interna-EM-PT": # Logica inversa
                ## PT + CSS: EM+Phones sin EM
                linept=customerid+"-DirNum-PT"
                linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                fwdnoregcss=customerid+"-DirnumEM-CSS"
                fwdnocss=customerid+"-DirnumEM-CSS"
                fwdallcss=customerid+"-InternalOnly-CSS"
            else:
                ## PT + CSS: Phones con EM
                linept=customerid+"-DirNumEM-PT"
                linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                fwdnoregcss=customerid+"-Dirnum-CSS"
                fwdnocss=customerid+"-InternalOnly-CSS"
                fwdallcss=customerid+"-InternalOnly-CSS"

            sheet['ej'+str(fila)]=linept                    #lines.line.0.dirn.routePartitionName
            sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
            #####################
            sheet['ek'+str(fila)]="Use System Policy"       #lines.line.0.mwlPolicy
            sheet['el'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
            sheet['em'+str(fila)]=row[166+delta]            #lines.line.0.busyTrigger
            sheet['en'+str(fila)]="Default"                 # lines.line.0.audibleMwi
            sheet['eo'+str(fila)]=row[179+delta]            # lines.line.0.display
            ######################
            sheet['ep'+str(fila)]="Default"                 # alwaysUsePrimeLine
            sheet['es'+str(fila)]="Default"                 # ringSettingIdleBlfAudibleAlert
            sheet['et'+str(fila)]="Default"                 # builtInBridgeStatus
            sheet['eu'+str(fila)]="0"                       # packetCaptureDuration
            sheet['ew'+str(fila)]="true"                    # useDevicePoolCgpnIngressDN
            sheet['ez'+str(fila)]="Off"                     # mlppIndicationStatus
            sheet['fb'+str(fila)]="None"                    # packetCaptureMode
            sheet['fc'+str(fila)]=""                        # status
            sheet['fd'+str(fila)]=""                        # loadInformation
            sheet['fe'+str(fila)]=row[42]                   # product
            sheet['ff'+str(fila)]=row[3]                    # description
            sheet['fg'+str(fila)]="false"                   # sendGeoLocation
            sheet['fh'+str(fila)]="p"+row[129]              # ownerUserName    @@@@@
            sheet['fi'+str(fila)]="false"                   # ignorePresentationIndicators
            ######################################################
            sheet['fm'+str(fila)]="Phone"                   # class
            sheet['fn'+str(fila)]="Use Common Phone Profile Setting"     # dndOption
            sheet['fo'+str(fila)]="Standard Presence group" # presenceGroupName
            sheet['fq'+str(fila)]="Standard Common Phone Profile"        # commonPhoneConfigName
            sheet['fs'+str(fila)]="711ulaw"                 # mtpPreferedCodec
            if row[21] != "":
                sheet['ft'+str(fila)]=cl+"-"+row[21]         # softkeyTemplateName
            sheet['fu'+str(fila)]="false"                   # remoteDevice
            sheet['fv'+str(fila)]=aarcss                    # automatedAlternateRoutingCssName
            ##
            ######################################################
            ######################################################
            ##
            ## LINE: LINE #1
            ##
            ######################################################
            ## DEBUG
            print("LN#",filadn,"                 ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164+delta])

            sheet = blk["LINE"]
            sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
            sheet['C'+str(filadn)]=action
            #sheet['D'+str(filadn)]="name:"+row[3] # Search field
            ######################################################
            sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
            sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
            sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
            sheet['m'+str(filadn)]=row[159+delta]                # CPG
            sheet['n'+str(filadn)]=row[188+delta]                #callForwardNotRegisteredInt.destination
            sheet['o'+str(filadn)]="false"                      #callForwardNotRegisteredInt.forwardToVoiceMail
            sheet['p'+str(filadn)]=fwdnocss                     #callForwardNotRegisteredInt.callingSearchSpaceName
            sheet['q'+str(filadn)]=linept                       #routePartitionName
            sheet['r'+str(filadn)]=row[182+delta]                #callForwardOnFailure.destination
            sheet['s'+str(filadn)]="false"                      #callForwardOnFailure.forwardToVoiceMail
            sheet['t'+str(filadn)]=fwdallcss                    #callForwardOnFailure.callingSearchSpaceName
            sheet['u'+str(filadn)]="false"                      #rejectAnonymousCall
            sheet['v'+str(filadn)]="true"                       #aarKeepCallHistory
            sheet['w'+str(filadn)]=linecss                      # LINE CSS
            if row[164] != "":
                sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]  # aarDestinationMask
            sheet['y'+str(filadn)]=row[176+delta]                # asciiAlertingName
            sheet['z'+str(filadn)]=row[129+delta]                # pattern
            sheet['aa'+str(filadn)]="Default"                   # patternPrecedence
            sheet['ab'+str(filadn)]=""                          # callForwardNoAnswer.duration
            sheet['ac'+str(filadn)]=row[150+delta]               # callForwardNoAnswer.destination
            sheet['ad'+str(filadn)]="false"                     # callForwardNoAnswer.forwardToVoiceMail
            sheet['ae'+str(filadn)]=fwdnocss                    # callForwardNoAnswer.callingSearchSpaceName
            sheet['ag'+str(filadn)]=row[156+delta]               # callForwardNoCoverage.destination
            sheet['ah'+str(filadn)]="false"                     # callForwardNoCoverage.forwardToVoiceMail
            sheet['ai'+str(filadn)]=fwdallcss                   # callForwardNoCoverage.callingSearchSpaceName
            sheet['aj'+str(filadn)]=row[191+delta]               # callForwardNotRegistered.destination
            sheet['ak'+str(filadn)]="false"                     # callForwardNotRegistered.forwardToVoiceMail
            sheet['al'+str(filadn)]=fwdnocss                    # callForwardNotRegistered.callingSearchSpaceName
            sheet['am'+str(filadn)]="Device"                    # usage
            sheet['ao'+str(filadn)]=row[174+delta]               #alertingName
            sheet['ap'+str(filadn)]=""                          #enterpriseAltNum.numMask
            sheet['aq'+str(filadn)]="false"                     #enterpriseAltNum.addLocalRoutePartition
            sheet['ar'+str(filadn)]="false"                     #enterpriseAltNum.advertiseGloballyIls
            sheet['as'+str(filadn)]=""                          #enterpriseAltNum.routePartition
            sheet['at'+str(filadn)]="false"                     #enterpriseAltNum.isUrgent
            sheet['au'+str(filadn)]=row[175+delta]              #description
            sheet['av'+str(filadn)]="false"                     #aarVoiceMailEnabled
            sheet['aw'+str(filadn)]="false"                     #useE164AltNum
            sheet['ba'+str(filadn)]="true"                      #allowCtiControlFlag
            sheet['bd'+str(filadn)]="No Error"                  #releaseClause
            sheet['be'+str(filadn)]=""                          #enterpriseAltNum.numMask
            sheet['bf'+str(filadn)]="false"                     #e164AltNum.addLocalRoutePartition
            sheet['bg'+str(filadn)]="true"                      #e164AltNum.advertiseGloballyIls
            sheet['bh'+str(filadn)]=""                          # e164AltNum.routePartition
            sheet['bi'+str(filadn)]="false"                     #e164AltNum.isUrgent
            sheet['bj'+str(filadn)]=devicecss                   # callForwardAll.secondaryCallingSearchSpaceName
            sheet['bk'+str(filadn)]=row[138+delta]               #callForwardAll.destination
            sheet['bl'+str(filadn)]="false"                     #callForwardAll.forwardToVoiceMail
            sheet['bm'+str(filadn)]=fwdallcss                   # callForwardAll.callingSearchSpaceName
            sheet['bn'+str(filadn)]="false"                     # parkMonForwardNoRetrieveVmEnabled
            sheet['bo'+str(filadn)]="true"                      # active
            sheet['bp'+str(filadn)]=""                          # VoiceMailProfileName
            sheet['bq'+str(filadn)]="false"                     # useEnterpriseAltNum
            sheet['bt'+str(filadn)]=row[141+delta]               # callForwardBusyInt.destination
            sheet['bu'+str(filadn)]="false"                     # callForwardBusyInt.forwardToVoiceMail
            sheet['bv'+str(filadn)]=fwdallcss                   # callForwardBusyInt.callingSearchSpaceName
            sheet['bw'+str(filadn)]=row[144+delta]               #  callForwardBusy.destination
            sheet['bx'+str(filadn)]="false"                     # callForwardBusy.forwardToVoiceMail
            sheet['by'+str(filadn)]=fwdallcss                   # callForwardBusy.callingSearchSpaceName
            sheet['ca'+str(filadn)]="false"                     #patternUrgency
            sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup'] #aarNeighborhoodName
            sheet['cc'+str(filadn)]="false"                     # parkMonForwardNoRetrieveIntVmEnabled
            sheet['cd'+str(filadn)]=row[153+delta]               # callForwardNoCoverageInt.destination
            sheet['ce'+str(filadn)]="false"                     # callForwardNoCoverageInt.forwardToVoiceMail
            sheet['cf'+str(filadn)]=fwdallcss                   #callForwardNoCoverageInt.callingSearchSpaceName
            sheet['cg'+str(filadn)]=""                          #callForwardAlternateParty.destination
            sheet['cj'+str(filadn)]=row[147+delta]               # callForwardNoAnswerInt.destination
            sheet['ck'+str(filadn)]="false"                     # callForwardNoAnswerInt.forwardToVoiceMail
            sheet['cl'+str(filadn)]=fwdnocss                    # callForwardNoAnswerInt.callingSearchSpaceName
            sheet['cn'+str(filadn)]="Standard Presence group"   #presenceGroupName
            filadn=filadn+1

            ##
            ## LINE #2
            ##
            if row[213][1:].startswith(siteslc):
                ##
                dn=dn+1
                delta=84*(dn-1) #L2 - 1
                ## DEBUG
                print("PH#",fila,row[2]," ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164])

                ######################################################
                ## PHONE
                ######################################################
                sheet =  blk["PHONE"]
                ############################################
                ##
                ## PHONE: LINE #2
                ##
                sheet['fx'+str(fila)]=row[180+delta]            #lines.line.0.displayAscii
                sheet['fy'+str(fila)]=""                        #lines.line.0.associatedEndusers.enduser.0.userId
                sheet['fz'+str(fila)]="Ring"                    #lines.line.0.ringSetting
                sheet['ga'+str(fila)]="Use System Default"      #lines.line.0.consecutiveRingSetting
                sheet['gb'+str(fila)]="Default"                 #lines.line.0.recordingProfileName
                sheet['gc'+str(fila)]=dn                        #lines.line.0.index
                sheet['gd'+str(fila)]="Use System Default"      # lines.line.0.ringSettingActivePickupAlert
                sheet['ge'+str(fila)]=row[163+delta]            #lines.line.0.label
                sheet['gf'+str(fila)]="Gateway Preferred"       #lines.line.0.recordingMediaSource
                sheet['gg'+str(fila)]=row[165+delta]            #lines.line.0.maxNumCalls
                sheet['gh'+str(fila)]="General"                 #lines.line.0.partitionUsage
                sheet['gi'+str(fila)]="Call Recording Disabled" #lines.line.0.recordingMediaSource
                sheet['gl'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]            #lines.line.0.e164Mask
                sheet['gm'+str(fila)]="true"                    #lines.line.0.missedCallLogging
                sheet['gn'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.dialedNumber
                sheet['go'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.redirectedNumber
                sheet['gp'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.callerName
                sheet['gq'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.callerNumber
                sheet['gr'+str(fila)]=row[129+delta]            #lines.line.0.dirn.pattern

                ## CSS depende de la PT
                #if row[130+delta] == "Interna-EM-PT": # Logica inversa
                #    ## PT + CSS: EM+Phones sin EM
                #    linept=customerid+"-DirNum-CSS"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-DirnumEM-CSS"
                #    fwdnocss=customerid+"-DirnumEM-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"
                #else:
                #    ## PT + CSS: Phones con EM
                #    linept=customerid+"-DirNumEM-PT"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-Dirnum-CSS"
                #    fwdnocss=customerid+"-InternalOnly-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"

                sheet['gs'+str(fila)]=linept                    #lines.line.0.dirn.routePartitionName
                #sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
                #####################
                sheet['gt'+str(fila)]="Use System Policy"       #lines.line.0.mwlPolicy
                sheet['gu'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                sheet['gv'+str(fila)]=row[166+delta]            #lines.line.0.busyTrigger
                sheet['gw'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                sheet['gx'+str(fila)]=row[179+delta]            # lines.line.0.display
                ######################################################
                ######################################################
                ##
                ## LINE: LINE #2
                ##
                ######################################################
                ## DEBUG
                print("LN#",filadn,"                 ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164+delta])

                sheet = blk["LINE"]
                sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                sheet['C'+str(filadn)]=action
                #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                ######################################################
                sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
                sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
                sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
                sheet['m'+str(filadn)]=row[159+delta]                # CPG
                sheet['n'+str(filadn)]=row[188+delta]                #callForwardNotRegisteredInt.destination
                sheet['o'+str(filadn)]="false"                      #callForwardNotRegisteredInt.forwardToVoiceMail
                sheet['p'+str(filadn)]=fwdnocss                     #callForwardNotRegisteredInt.callingSearchSpaceName
                sheet['q'+str(filadn)]=linept                       #routePartitionName
                sheet['r'+str(filadn)]=row[182+delta]                #callForwardOnFailure.destination
                sheet['s'+str(filadn)]="false"                      #callForwardOnFailure.forwardToVoiceMail
                sheet['t'+str(filadn)]=fwdallcss                    #callForwardOnFailure.callingSearchSpaceName
                sheet['u'+str(filadn)]="false"                      #rejectAnonymousCall
                sheet['v'+str(filadn)]="true"                       #aarKeepCallHistory
                sheet['w'+str(filadn)]=linecss                      # LINE CSS
                if row[164] != "":
                    sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]  # aarDestinationMask
                sheet['y'+str(filadn)]=row[176+delta]                # asciiAlertingName
                sheet['z'+str(filadn)]=row[129+delta]                # pattern
                sheet['aa'+str(filadn)]="Default"                   # patternPrecedence
                sheet['ab'+str(filadn)]=""                          # callForwardNoAnswer.duration
                sheet['ac'+str(filadn)]=row[150+delta]               # callForwardNoAnswer.destination
                sheet['ad'+str(filadn)]="false"                     # callForwardNoAnswer.forwardToVoiceMail
                sheet['ae'+str(filadn)]=fwdnocss                    # callForwardNoAnswer.callingSearchSpaceName
                sheet['ag'+str(filadn)]=row[156+delta]               # callForwardNoCoverage.destination
                sheet['ah'+str(filadn)]="false"                     # callForwardNoCoverage.forwardToVoiceMail
                sheet['ai'+str(filadn)]=fwdallcss                   # callForwardNoCoverage.callingSearchSpaceName
                sheet['aj'+str(filadn)]=row[191+delta]               # callForwardNotRegistered.destination
                sheet['ak'+str(filadn)]="false"                     # callForwardNotRegistered.forwardToVoiceMail
                sheet['al'+str(filadn)]=fwdnocss                    # callForwardNotRegistered.callingSearchSpaceName
                sheet['am'+str(filadn)]="Device"                    # usage
                sheet['ao'+str(filadn)]=row[174+delta]               #alertingName
                sheet['ap'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['aq'+str(filadn)]="false"                     #enterpriseAltNum.addLocalRoutePartition
                sheet['ar'+str(filadn)]="false"                     #enterpriseAltNum.advertiseGloballyIls
                sheet['as'+str(filadn)]=""                          #enterpriseAltNum.routePartition
                sheet['at'+str(filadn)]="false"                     #enterpriseAltNum.isUrgent
                sheet['au'+str(filadn)]=row[175+delta]              #description
                sheet['av'+str(filadn)]="false"                     #aarVoiceMailEnabled
                sheet['aw'+str(filadn)]="false"                     #useE164AltNum
                sheet['ba'+str(filadn)]="true"                      #allowCtiControlFlag
                sheet['bd'+str(filadn)]="No Error"                  #releaseClause
                sheet['be'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['bf'+str(filadn)]="false"                     #e164AltNum.addLocalRoutePartition
                sheet['bg'+str(filadn)]="true"                      #e164AltNum.advertiseGloballyIls
                sheet['bh'+str(filadn)]=""                          # e164AltNum.routePartition
                sheet['bi'+str(filadn)]="false"                     #e164AltNum.isUrgent
                sheet['bj'+str(filadn)]=devicecss                   # callForwardAll.secondaryCallingSearchSpaceName
                sheet['bk'+str(filadn)]=row[138+delta]               #callForwardAll.destination
                sheet['bl'+str(filadn)]="false"                     #callForwardAll.forwardToVoiceMail
                sheet['bm'+str(filadn)]=fwdallcss                   # callForwardAll.callingSearchSpaceName
                sheet['bn'+str(filadn)]="false"                     # parkMonForwardNoRetrieveVmEnabled
                sheet['bo'+str(filadn)]="true"                      # active
                sheet['bp'+str(filadn)]=""                          # VoiceMailProfileName
                sheet['bq'+str(filadn)]="false"                     # useEnterpriseAltNum
                sheet['bt'+str(filadn)]=row[141+delta]               # callForwardBusyInt.destination
                sheet['bu'+str(filadn)]="false"                     # callForwardBusyInt.forwardToVoiceMail
                sheet['bv'+str(filadn)]=fwdallcss                   # callForwardBusyInt.callingSearchSpaceName
                sheet['bw'+str(filadn)]=row[144+delta]               #  callForwardBusy.destination
                sheet['bx'+str(filadn)]="false"                     # callForwardBusy.forwardToVoiceMail
                sheet['by'+str(filadn)]=fwdallcss                   # callForwardBusy.callingSearchSpaceName
                sheet['ca'+str(filadn)]="false"                     #patternUrgency
                sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup'] #aarNeighborhoodName
                sheet['cc'+str(filadn)]="false"                     # parkMonForwardNoRetrieveIntVmEnabled
                sheet['cd'+str(filadn)]=row[153+delta]               # callForwardNoCoverageInt.destination
                sheet['ce'+str(filadn)]="false"                     # callForwardNoCoverageInt.forwardToVoiceMail
                sheet['cf'+str(filadn)]=fwdallcss                   #callForwardNoCoverageInt.callingSearchSpaceName
                sheet['cg'+str(filadn)]=""                          #callForwardAlternateParty.destination
                sheet['cj'+str(filadn)]=row[147+delta]               # callForwardNoAnswerInt.destination
                sheet['ck'+str(filadn)]="false"                     # callForwardNoAnswerInt.forwardToVoiceMail
                sheet['cl'+str(filadn)]=fwdnocss                    # callForwardNoAnswerInt.callingSearchSpaceName
                sheet['cn'+str(filadn)]="Standard Presence group"   #presenceGroupName
                filadn=filadn+1

            ##
            ## LINE #3
            ##
            if row[297][1:].startswith(siteslc):
                ##
                dn=dn+1
                delta=84*(dn-1) #L2 - 1
                ## DEBUG
                print("PH#",fila,row[2]," ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164])

                ######################################################
                ## PHONE
                ######################################################
                sheet =  blk["PHONE"]
                ############################################
                ##
                ## PHONE: LINE #3
                ##
                sheet['gy'+str(fila)]=row[180+delta]            #lines.line.0.displayAscii
                sheet['gz'+str(fila)]=""                        #lines.line.0.associatedEndusers.enduser.0.userId
                sheet['ha'+str(fila)]="Ring"                    #lines.line.0.ringSetting
                sheet['hb'+str(fila)]="Use System Default"      #lines.line.0.consecutiveRingSetting
                sheet['hc'+str(fila)]="Default"                 #lines.line.0.recordingProfileName
                sheet['hd'+str(fila)]=dn                        #lines.line.0.index
                sheet['he'+str(fila)]="Use System Default"      # lines.line.0.ringSettingActivePickupAlert
                sheet['hf'+str(fila)]=row[163+delta]            #lines.line.0.label
                sheet['hg'+str(fila)]="Gateway Preferred"       #lines.line.0.recordingMediaSource
                sheet['hh'+str(fila)]=row[165+delta]            #lines.line.0.maxNumCalls
                sheet['hi'+str(fila)]="General"                 #lines.line.0.partitionUsage
                sheet['hl'+str(fila)]="Call Recording Disabled" #lines.line.0.recordingMediaSource
                sheet['hm'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]            #lines.line.0.e164Mask
                sheet['hn'+str(fila)]="true"                    #lines.line.0.missedCallLogging
                sheet['ho'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.dialedNumber
                sheet['hp'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.redirectedNumber
                sheet['hq'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.callerName
                sheet['hr'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.callerNumber
                sheet['hs'+str(fila)]=row[129+delta]            #lines.line.0.dirn.pattern

                ## CSS depende de la PT
                #if row[130+delta] == "Interna-EM-PT": # Logica inversa
                #    ## PT + CSS: EM+Phones sin EM
                #    linept=customerid+"-DirNum-CSS"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-DirnumEM-CSS"
                #    fwdnocss=customerid+"-DirnumEM-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"
                #else:
                #    ## PT + CSS: Phones con EM
                #    linept=customerid+"-DirNumEM-PT"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-Dirnum-CSS"
                #    fwdnocss=customerid+"-InternalOnly-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"

                sheet['ht'+str(fila)]=linept                    #lines.line.0.dirn.routePartitionName
                #sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
                #####################
                sheet['hu'+str(fila)]="Use System Policy"       #lines.line.0.mwlPolicy
                sheet['hv'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                sheet['hw'+str(fila)]=row[166+delta]            #lines.line.0.busyTrigger
                sheet['hx'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                sheet['hy'+str(fila)]=row[179+delta]            # lines.line.0.display
                ######################################################
                ######################################################
                ##
                ## LINE: LINE #3
                ##
                ######################################################
                ## DEBUG
                print("LN#",filadn,"                 ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164+delta])

                sheet = blk["LINE"]
                sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                sheet['C'+str(filadn)]=action
                #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                ######################################################
                sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
                sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
                sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
                sheet['m'+str(filadn)]=row[159+delta]                # CPG
                sheet['n'+str(filadn)]=row[188+delta]                #callForwardNotRegisteredInt.destination
                sheet['o'+str(filadn)]="false"                      #callForwardNotRegisteredInt.forwardToVoiceMail
                sheet['p'+str(filadn)]=fwdnocss                     #callForwardNotRegisteredInt.callingSearchSpaceName
                sheet['q'+str(filadn)]=linept                       #routePartitionName
                sheet['r'+str(filadn)]=row[182+delta]                #callForwardOnFailure.destination
                sheet['s'+str(filadn)]="false"                      #callForwardOnFailure.forwardToVoiceMail
                sheet['t'+str(filadn)]=fwdallcss                    #callForwardOnFailure.callingSearchSpaceName
                sheet['u'+str(filadn)]="false"                      #rejectAnonymousCall
                sheet['v'+str(filadn)]="true"                       #aarKeepCallHistory
                sheet['w'+str(filadn)]=linecss                      # LINE CSS
                if row[164] != "":
                    sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]  # aarDestinationMask
                sheet['y'+str(filadn)]=row[176+delta]                # asciiAlertingName
                sheet['z'+str(filadn)]=row[129+delta]                # pattern
                sheet['aa'+str(filadn)]="Default"                   # patternPrecedence
                sheet['ab'+str(filadn)]=""                          # callForwardNoAnswer.duration
                sheet['ac'+str(filadn)]=row[150+delta]               # callForwardNoAnswer.destination
                sheet['ad'+str(filadn)]="false"                     # callForwardNoAnswer.forwardToVoiceMail
                sheet['ae'+str(filadn)]=fwdnocss                    # callForwardNoAnswer.callingSearchSpaceName
                sheet['ag'+str(filadn)]=row[156+delta]               # callForwardNoCoverage.destination
                sheet['ah'+str(filadn)]="false"                     # callForwardNoCoverage.forwardToVoiceMail
                sheet['ai'+str(filadn)]=fwdallcss                   # callForwardNoCoverage.callingSearchSpaceName
                sheet['aj'+str(filadn)]=row[191+delta]               # callForwardNotRegistered.destination
                sheet['ak'+str(filadn)]="false"                     # callForwardNotRegistered.forwardToVoiceMail
                sheet['al'+str(filadn)]=fwdnocss                    # callForwardNotRegistered.callingSearchSpaceName
                sheet['am'+str(filadn)]="Device"                    # usage
                sheet['ao'+str(filadn)]=row[174+delta]               #alertingName
                sheet['ap'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['aq'+str(filadn)]="false"                     #enterpriseAltNum.addLocalRoutePartition
                sheet['ar'+str(filadn)]="false"                     #enterpriseAltNum.advertiseGloballyIls
                sheet['as'+str(filadn)]=""                          #enterpriseAltNum.routePartition
                sheet['at'+str(filadn)]="false"                     #enterpriseAltNum.isUrgent
                sheet['au'+str(filadn)]=row[175+delta]              #description
                sheet['av'+str(filadn)]="false"                     #aarVoiceMailEnabled
                sheet['aw'+str(filadn)]="false"                     #useE164AltNum
                sheet['ba'+str(filadn)]="true"                      #allowCtiControlFlag
                sheet['bd'+str(filadn)]="No Error"                  #releaseClause
                sheet['be'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['bf'+str(filadn)]="false"                     #e164AltNum.addLocalRoutePartition
                sheet['bg'+str(filadn)]="true"                      #e164AltNum.advertiseGloballyIls
                sheet['bh'+str(filadn)]=""                          # e164AltNum.routePartition
                sheet['bi'+str(filadn)]="false"                     #e164AltNum.isUrgent
                sheet['bj'+str(filadn)]=devicecss                   # callForwardAll.secondaryCallingSearchSpaceName
                sheet['bk'+str(filadn)]=row[138+delta]               #callForwardAll.destination
                sheet['bl'+str(filadn)]="false"                     #callForwardAll.forwardToVoiceMail
                sheet['bm'+str(filadn)]=fwdallcss                   # callForwardAll.callingSearchSpaceName
                sheet['bn'+str(filadn)]="false"                     # parkMonForwardNoRetrieveVmEnabled
                sheet['bo'+str(filadn)]="true"                      # active
                sheet['bp'+str(filadn)]=""                          # VoiceMailProfileName
                sheet['bq'+str(filadn)]="false"                     # useEnterpriseAltNum
                sheet['bt'+str(filadn)]=row[141+delta]               # callForwardBusyInt.destination
                sheet['bu'+str(filadn)]="false"                     # callForwardBusyInt.forwardToVoiceMail
                sheet['bv'+str(filadn)]=fwdallcss                   # callForwardBusyInt.callingSearchSpaceName
                sheet['bw'+str(filadn)]=row[144+delta]               #  callForwardBusy.destination
                sheet['bx'+str(filadn)]="false"                     # callForwardBusy.forwardToVoiceMail
                sheet['by'+str(filadn)]=fwdallcss                   # callForwardBusy.callingSearchSpaceName
                sheet['ca'+str(filadn)]="false"                     #patternUrgency
                sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup'] #aarNeighborhoodName
                sheet['cc'+str(filadn)]="false"                     # parkMonForwardNoRetrieveIntVmEnabled
                sheet['cd'+str(filadn)]=row[153+delta]               # callForwardNoCoverageInt.destination
                sheet['ce'+str(filadn)]="false"                     # callForwardNoCoverageInt.forwardToVoiceMail
                sheet['cf'+str(filadn)]=fwdallcss                   #callForwardNoCoverageInt.callingSearchSpaceName
                sheet['cg'+str(filadn)]=""                          #callForwardAlternateParty.destination
                sheet['cj'+str(filadn)]=row[147+delta]               # callForwardNoAnswerInt.destination
                sheet['ck'+str(filadn)]="false"                     # callForwardNoAnswerInt.forwardToVoiceMail
                sheet['cl'+str(filadn)]=fwdnocss                    # callForwardNoAnswerInt.callingSearchSpaceName
                sheet['cn'+str(filadn)]="Standard Presence group"   #presenceGroupName
                filadn=filadn+1


            ##
            ## LINE #4
            ##
            if row[381][1:].startswith(siteslc):
                ##
                dn=dn+1
                delta=84*(dn-1) #L2 - 1
                ## DEBUG
                print("PH#",fila,row[2]," ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164])

                ######################################################
                ## PHONE
                ######################################################
                sheet =  blk["PHONE"]
                ############################################
                ##
                ## PHONE: LINE #4
                ##
                sheet['hz'+str(fila)]=row[180+delta]            #lines.line.0.displayAscii
                sheet['ia'+str(fila)]=""                        #lines.line.0.associatedEndusers.enduser.0.userId
                sheet['ib'+str(fila)]="Ring"                    #lines.line.0.ringSetting
                sheet['ic'+str(fila)]="Use System Default"      #lines.line.0.consecutiveRingSetting
                sheet['id'+str(fila)]="Default"                 #lines.line.0.recordingProfileName
                sheet['ie'+str(fila)]=dn                        #lines.line.0.index
                sheet['if'+str(fila)]="Use System Default"      # lines.line.0.ringSettingActivePickupAlert
                sheet['ig'+str(fila)]=row[163+delta]            #lines.line.0.label
                sheet['ih'+str(fila)]="Gateway Preferred"       #lines.line.0.recordingMediaSource
                sheet['ii'+str(fila)]=row[165+delta]            #lines.line.0.maxNumCalls
                sheet['il'+str(fila)]="General"                 #lines.line.0.partitionUsage
                sheet['im'+str(fila)]="Call Recording Disabled" #lines.line.0.recordingMediaSource
                sheet['in'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]            #lines.line.0.e164Mask
                sheet['io'+str(fila)]="true"                    #lines.line.0.missedCallLogging
                sheet['ip'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.dialedNumber
                sheet['iq'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.redirectedNumber
                sheet['ir'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.callerName
                sheet['is'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.callerNumber
                sheet['it'+str(fila)]=row[129+delta]            #lines.line.0.dirn.pattern

                ## CSS depende de la PT
                #if row[130+delta] == "Interna-EM-PT": # Logica inversa
                #    ## PT + CSS: EM+Phones sin EM
                #    linept=customerid+"-DirNum-CSS"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-DirnumEM-CSS"
                #    fwdnocss=customerid+"-DirnumEM-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"
                #else:
                #    ## PT + CSS: Phones con EM
                #    linept=customerid+"-DirNumEM-PT"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-Dirnum-CSS"
                #    fwdnocss=customerid+"-InternalOnly-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"

                sheet['iu'+str(fila)]=linept                    #lines.line.0.dirn.routePartitionName
                #sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
                #####################
                sheet['iv'+str(fila)]="Use System Policy"       #lines.line.0.mwlPolicy
                sheet['iw'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                sheet['ix'+str(fila)]=row[166+delta]            #lines.line.0.busyTrigger
                sheet['iy'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                sheet['iz'+str(fila)]=row[179+delta]            # lines.line.0.display
                ######################################################
                ######################################################
                ##
                ## LINE: LINE #4
                ##
                ######################################################
                ## DEBUG
                print("LN#",filadn,"                 ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164+delta])

                sheet = blk["LINE"]
                sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                sheet['C'+str(filadn)]=action
                #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                ######################################################
                sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
                sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
                sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
                sheet['m'+str(filadn)]=row[159+delta]                # CPG
                sheet['n'+str(filadn)]=row[188+delta]                #callForwardNotRegisteredInt.destination
                sheet['o'+str(filadn)]="false"                      #callForwardNotRegisteredInt.forwardToVoiceMail
                sheet['p'+str(filadn)]=fwdnocss                     #callForwardNotRegisteredInt.callingSearchSpaceName
                sheet['q'+str(filadn)]=linept                       #routePartitionName
                sheet['r'+str(filadn)]=row[182+delta]                #callForwardOnFailure.destination
                sheet['s'+str(filadn)]="false"                      #callForwardOnFailure.forwardToVoiceMail
                sheet['t'+str(filadn)]=fwdallcss                    #callForwardOnFailure.callingSearchSpaceName
                sheet['u'+str(filadn)]="false"                      #rejectAnonymousCall
                sheet['v'+str(filadn)]="true"                       #aarKeepCallHistory
                sheet['w'+str(filadn)]=linecss                      # LINE CSS
                if row[164] != "":
                    sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]  # aarDestinationMask
                sheet['y'+str(filadn)]=row[176+delta]                # asciiAlertingName
                sheet['z'+str(filadn)]=row[129+delta]                # pattern
                sheet['aa'+str(filadn)]="Default"                   # patternPrecedence
                sheet['ab'+str(filadn)]=""                          # callForwardNoAnswer.duration
                sheet['ac'+str(filadn)]=row[150+delta]               # callForwardNoAnswer.destination
                sheet['ad'+str(filadn)]="false"                     # callForwardNoAnswer.forwardToVoiceMail
                sheet['ae'+str(filadn)]=fwdnocss                    # callForwardNoAnswer.callingSearchSpaceName
                sheet['ag'+str(filadn)]=row[156+delta]               # callForwardNoCoverage.destination
                sheet['ah'+str(filadn)]="false"                     # callForwardNoCoverage.forwardToVoiceMail
                sheet['ai'+str(filadn)]=fwdallcss                   # callForwardNoCoverage.callingSearchSpaceName
                sheet['aj'+str(filadn)]=row[191+delta]               # callForwardNotRegistered.destination
                sheet['ak'+str(filadn)]="false"                     # callForwardNotRegistered.forwardToVoiceMail
                sheet['al'+str(filadn)]=fwdnocss                    # callForwardNotRegistered.callingSearchSpaceName
                sheet['am'+str(filadn)]="Device"                    # usage
                sheet['ao'+str(filadn)]=row[174+delta]               #alertingName
                sheet['ap'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['aq'+str(filadn)]="false"                     #enterpriseAltNum.addLocalRoutePartition
                sheet['ar'+str(filadn)]="false"                     #enterpriseAltNum.advertiseGloballyIls
                sheet['as'+str(filadn)]=""                          #enterpriseAltNum.routePartition
                sheet['at'+str(filadn)]="false"                     #enterpriseAltNum.isUrgent
                sheet['au'+str(filadn)]=row[175+delta]              #description
                sheet['av'+str(filadn)]="false"                     #aarVoiceMailEnabled
                sheet['aw'+str(filadn)]="false"                     #useE164AltNum
                sheet['ba'+str(filadn)]="true"                      #allowCtiControlFlag
                sheet['bd'+str(filadn)]="No Error"                  #releaseClause
                sheet['be'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['bf'+str(filadn)]="false"                     #e164AltNum.addLocalRoutePartition
                sheet['bg'+str(filadn)]="true"                      #e164AltNum.advertiseGloballyIls
                sheet['bh'+str(filadn)]=""                          # e164AltNum.routePartition
                sheet['bi'+str(filadn)]="false"                     #e164AltNum.isUrgent
                sheet['bj'+str(filadn)]=devicecss                   # callForwardAll.secondaryCallingSearchSpaceName
                sheet['bk'+str(filadn)]=row[138+delta]               #callForwardAll.destination
                sheet['bl'+str(filadn)]="false"                     #callForwardAll.forwardToVoiceMail
                sheet['bm'+str(filadn)]=fwdallcss                   # callForwardAll.callingSearchSpaceName
                sheet['bn'+str(filadn)]="false"                     # parkMonForwardNoRetrieveVmEnabled
                sheet['bo'+str(filadn)]="true"                      # active
                sheet['bp'+str(filadn)]=""                          # VoiceMailProfileName
                sheet['bq'+str(filadn)]="false"                     # useEnterpriseAltNum
                sheet['bt'+str(filadn)]=row[141+delta]               # callForwardBusyInt.destination
                sheet['bu'+str(filadn)]="false"                     # callForwardBusyInt.forwardToVoiceMail
                sheet['bv'+str(filadn)]=fwdallcss                   # callForwardBusyInt.callingSearchSpaceName
                sheet['bw'+str(filadn)]=row[144+delta]               #  callForwardBusy.destination
                sheet['bx'+str(filadn)]="false"                     # callForwardBusy.forwardToVoiceMail
                sheet['by'+str(filadn)]=fwdallcss                   # callForwardBusy.callingSearchSpaceName
                sheet['ca'+str(filadn)]="false"                     #patternUrgency
                sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup'] #aarNeighborhoodName
                sheet['cc'+str(filadn)]="false"                     # parkMonForwardNoRetrieveIntVmEnabled
                sheet['cd'+str(filadn)]=row[153+delta]               # callForwardNoCoverageInt.destination
                sheet['ce'+str(filadn)]="false"                     # callForwardNoCoverageInt.forwardToVoiceMail
                sheet['cf'+str(filadn)]=fwdallcss                   #callForwardNoCoverageInt.callingSearchSpaceName
                sheet['cg'+str(filadn)]=""                          #callForwardAlternateParty.destination
                sheet['cj'+str(filadn)]=row[147+delta]               # callForwardNoAnswerInt.destination
                sheet['ck'+str(filadn)]="false"                     # callForwardNoAnswerInt.forwardToVoiceMail
                sheet['cl'+str(filadn)]=fwdnocss                    # callForwardNoAnswerInt.callingSearchSpaceName
                sheet['cn'+str(filadn)]="Standard Presence group"   #presenceGroupName
                filadn=filadn+1


            ##
            ## LINE #5
            ##
            if row[465][1:].startswith(siteslc):
                ##
                dn=dn+1
                delta=84*(dn-1) #L2 - 1
                ## DEBUG
                print("PH#",fila,row[2]," ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164])

                ######################################################
                ## PHONE
                ######################################################
                sheet =  blk["PHONE"]
                ############################################
                ##
                ## PHONE: LINE #5
                ##
                sheet['ja'+str(fila)]=row[180+delta]            #lines.line.0.displayAscii
                sheet['jb'+str(fila)]=""                        #lines.line.0.associatedEndusers.enduser.0.userId
                sheet['jc'+str(fila)]="Ring"                    #lines.line.0.ringSetting
                sheet['jd'+str(fila)]="Use System Default"      #lines.line.0.consecutiveRingSetting
                sheet['je'+str(fila)]="Default"                 #lines.line.0.recordingProfileName
                sheet['jf'+str(fila)]=dn                        #lines.line.0.index
                sheet['jg'+str(fila)]="Use System Default"      # lines.line.0.ringSettingActivePickupAlert
                sheet['jh'+str(fila)]=row[163+delta]            #lines.line.0.label
                sheet['ji'+str(fila)]="Gateway Preferred"       #lines.line.0.recordingMediaSource
                sheet['jl'+str(fila)]=row[165+delta]            #lines.line.0.maxNumCalls
                sheet['jm'+str(fila)]="General"                 #lines.line.0.partitionUsage
                sheet['jn'+str(fila)]="Call Recording Disabled" #lines.line.0.recordingMediaSource
                sheet['jo'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]            #lines.line.0.e164Mask
                sheet['jp'+str(fila)]="true"                    #lines.line.0.missedCallLogging
                sheet['jq'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.dialedNumber
                sheet['jr'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.redirectedNumber
                sheet['js'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.callerName
                sheet['jt'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.callerNumber
                sheet['ju'+str(fila)]=row[129+delta]            #lines.line.0.dirn.pattern

                ## CSS depende de la PT
                #if row[130+delta] == "Interna-EM-PT": # Logica inversa
                #    ## PT + CSS: EM+Phones sin EM
                #    linept=customerid+"-DirNum-CSS"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-DirnumEM-CSS"
                #    fwdnocss=customerid+"-DirnumEM-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"
                #else:
                #    ## PT + CSS: Phones con EM
                #    linept=customerid+"-DirNumEM-PT"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-Dirnum-CSS"
                #    fwdnocss=customerid+"-InternalOnly-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"

                sheet['jv'+str(fila)]=linept                    #lines.line.0.dirn.routePartitionName
                #sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
                #####################
                sheet['jw'+str(fila)]="Use System Policy"       #lines.line.0.mwlPolicy
                sheet['jx'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                sheet['jy'+str(fila)]=row[166+delta]            #lines.line.0.busyTrigger
                sheet['jz'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                sheet['ka'+str(fila)]=row[179+delta]            # lines.line.0.display
                ######################################################
                ######################################################
                ##
                ## LINE: LINE #5
                ##
                ######################################################
                ## DEBUG
                print("LN#",filadn,"                 ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164+delta])

                sheet = blk["LINE"]
                sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                sheet['C'+str(filadn)]=action
                #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                ######################################################
                sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
                sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
                sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
                sheet['m'+str(filadn)]=row[159+delta]                # CPG
                sheet['n'+str(filadn)]=row[188+delta]                #callForwardNotRegisteredInt.destination
                sheet['o'+str(filadn)]="false"                      #callForwardNotRegisteredInt.forwardToVoiceMail
                sheet['p'+str(filadn)]=fwdnocss                     #callForwardNotRegisteredInt.callingSearchSpaceName
                sheet['q'+str(filadn)]=linept                       #routePartitionName
                sheet['r'+str(filadn)]=row[182+delta]                #callForwardOnFailure.destination
                sheet['s'+str(filadn)]="false"                      #callForwardOnFailure.forwardToVoiceMail
                sheet['t'+str(filadn)]=fwdallcss                    #callForwardOnFailure.callingSearchSpaceName
                sheet['u'+str(filadn)]="false"                      #rejectAnonymousCall
                sheet['v'+str(filadn)]="true"                       #aarKeepCallHistory
                sheet['w'+str(filadn)]=linecss                      # LINE CSS
                if row[164] != "":
                    sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]  # aarDestinationMask
                sheet['y'+str(filadn)]=row[176+delta]                # asciiAlertingName
                sheet['z'+str(filadn)]=row[129+delta]                # pattern
                sheet['aa'+str(filadn)]="Default"                   # patternPrecedence
                sheet['ab'+str(filadn)]=""                          # callForwardNoAnswer.duration
                sheet['ac'+str(filadn)]=row[150+delta]               # callForwardNoAnswer.destination
                sheet['ad'+str(filadn)]="false"                     # callForwardNoAnswer.forwardToVoiceMail
                sheet['ae'+str(filadn)]=fwdnocss                    # callForwardNoAnswer.callingSearchSpaceName
                sheet['ag'+str(filadn)]=row[156+delta]               # callForwardNoCoverage.destination
                sheet['ah'+str(filadn)]="false"                     # callForwardNoCoverage.forwardToVoiceMail
                sheet['ai'+str(filadn)]=fwdallcss                   # callForwardNoCoverage.callingSearchSpaceName
                sheet['aj'+str(filadn)]=row[191+delta]               # callForwardNotRegistered.destination
                sheet['ak'+str(filadn)]="false"                     # callForwardNotRegistered.forwardToVoiceMail
                sheet['al'+str(filadn)]=fwdnocss                    # callForwardNotRegistered.callingSearchSpaceName
                sheet['am'+str(filadn)]="Device"                    # usage
                sheet['ao'+str(filadn)]=row[174+delta]               #alertingName
                sheet['ap'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['aq'+str(filadn)]="false"                     #enterpriseAltNum.addLocalRoutePartition
                sheet['ar'+str(filadn)]="false"                     #enterpriseAltNum.advertiseGloballyIls
                sheet['as'+str(filadn)]=""                          #enterpriseAltNum.routePartition
                sheet['at'+str(filadn)]="false"                     #enterpriseAltNum.isUrgent
                sheet['au'+str(filadn)]=row[175+delta]              #description
                sheet['av'+str(filadn)]="false"                     #aarVoiceMailEnabled
                sheet['aw'+str(filadn)]="false"                     #useE164AltNum
                sheet['ba'+str(filadn)]="true"                      #allowCtiControlFlag
                sheet['bd'+str(filadn)]="No Error"                  #releaseClause
                sheet['be'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['bf'+str(filadn)]="false"                     #e164AltNum.addLocalRoutePartition
                sheet['bg'+str(filadn)]="true"                      #e164AltNum.advertiseGloballyIls
                sheet['bh'+str(filadn)]=""                          # e164AltNum.routePartition
                sheet['bi'+str(filadn)]="false"                     #e164AltNum.isUrgent
                sheet['bj'+str(filadn)]=devicecss                   # callForwardAll.secondaryCallingSearchSpaceName
                sheet['bk'+str(filadn)]=row[138+delta]               #callForwardAll.destination
                sheet['bl'+str(filadn)]="false"                     #callForwardAll.forwardToVoiceMail
                sheet['bm'+str(filadn)]=fwdallcss                   # callForwardAll.callingSearchSpaceName
                sheet['bn'+str(filadn)]="false"                     # parkMonForwardNoRetrieveVmEnabled
                sheet['bo'+str(filadn)]="true"                      # active
                sheet['bp'+str(filadn)]=""                          # VoiceMailProfileName
                sheet['bq'+str(filadn)]="false"                     # useEnterpriseAltNum
                sheet['bt'+str(filadn)]=row[141+delta]               # callForwardBusyInt.destination
                sheet['bu'+str(filadn)]="false"                     # callForwardBusyInt.forwardToVoiceMail
                sheet['bv'+str(filadn)]=fwdallcss                   # callForwardBusyInt.callingSearchSpaceName
                sheet['bw'+str(filadn)]=row[144+delta]               #  callForwardBusy.destination
                sheet['bx'+str(filadn)]="false"                     # callForwardBusy.forwardToVoiceMail
                sheet['by'+str(filadn)]=fwdallcss                   # callForwardBusy.callingSearchSpaceName
                sheet['ca'+str(filadn)]="false"                     #patternUrgency
                sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup'] #aarNeighborhoodName
                sheet['cc'+str(filadn)]="false"                     # parkMonForwardNoRetrieveIntVmEnabled
                sheet['cd'+str(filadn)]=row[153+delta]               # callForwardNoCoverageInt.destination
                sheet['ce'+str(filadn)]="false"                     # callForwardNoCoverageInt.forwardToVoiceMail
                sheet['cf'+str(filadn)]=fwdallcss                   #callForwardNoCoverageInt.callingSearchSpaceName
                sheet['cg'+str(filadn)]=""                          #callForwardAlternateParty.destination
                sheet['cj'+str(filadn)]=row[147+delta]               # callForwardNoAnswerInt.destination
                sheet['ck'+str(filadn)]="false"                     # callForwardNoAnswerInt.forwardToVoiceMail
                sheet['cl'+str(filadn)]=fwdnocss                    # callForwardNoAnswerInt.callingSearchSpaceName
                sheet['cn'+str(filadn)]="Standard Presence group"   #presenceGroupName
                filadn=filadn+1


            ##
            ## LINE #6
            ##
            if row[549][1:].startswith(siteslc):
                ##
                dn=dn+1
                delta=84*(dn-1) #L2 - 1
                ## DEBUG
                print("PH#",fila,row[2]," ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164])

                ######################################################
                ## PHONE
                ######################################################
                sheet =  blk["PHONE"]
                ############################################
                ##
                ## PHONE: LINE #6
                ##
                sheet['kb'+str(fila)]=row[180+delta]            #lines.line.0.displayAscii
                sheet['kc'+str(fila)]=""                        #lines.line.0.associatedEndusers.enduser.0.userId
                sheet['kd'+str(fila)]="Ring"                    #lines.line.0.ringSetting
                sheet['ke'+str(fila)]="Use System Default"      #lines.line.0.consecutiveRingSetting
                sheet['kf'+str(fila)]="Default"                 #lines.line.0.recordingProfileName
                sheet['kg'+str(fila)]=dn                        #lines.line.0.index
                sheet['kh'+str(fila)]="Use System Default"      # lines.line.0.ringSettingActivePickupAlert
                sheet['ki'+str(fila)]=row[163+delta]            #lines.line.0.label
                sheet['kj'+str(fila)]="Gateway Preferred"       #lines.line.0.recordingMediaSource
                sheet['kk'+str(fila)]=row[165+delta]            #lines.line.0.maxNumCalls
                sheet['kl'+str(fila)]="General"                 #lines.line.0.partitionUsage
                sheet['km'+str(fila)]="Call Recording Disabled" #lines.line.0.recordingMediaSource
                sheet['kp'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]            #lines.line.0.e164Mask
                sheet['kq'+str(fila)]="true"                    #lines.line.0.missedCallLogging
                sheet['kr'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.dialedNumber
                sheet['ks'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.redirectedNumber
                sheet['kt'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.callerName
                sheet['ku'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.callerNumber
                sheet['kv'+str(fila)]=row[129+delta]            #lines.line.0.dirn.pattern

                ## CSS depende de la PT
                #if row[130+delta] == "Interna-EM-PT": # Logica inversa
                #    ## PT + CSS: EM+Phones sin EM
                #    linept=customerid+"-DirNum-CSS"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-DirnumEM-CSS"
                #    fwdnocss=customerid+"-DirnumEM-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"
                #else:
                #    ## PT + CSS: Phones con EM
                #    linept=customerid+"-DirNumEM-PT"
                #    linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                #    devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                #    fwdnoregcss=customerid+"-Dirnum-CSS"
                #    fwdnocss=customerid+"-InternalOnly-CSS"
                #    fwdallcss=customerid+"-InternalOnly-CSS"

                sheet['kw'+str(fila)]=linept                    #lines.line.0.dirn.routePartitionName
                #sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
                #####################
                sheet['kx'+str(fila)]="Use System Policy"       #lines.line.0.mwlPolicy
                sheet['ky'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                sheet['kz'+str(fila)]=row[166+delta]            #lines.line.0.busyTrigger
                sheet['la'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                sheet['lb'+str(fila)]=row[179+delta]            # lines.line.0.display
                ######################################################
                ######################################################
                ##
                ## LINE: LINE #6
                ##
                ######################################################
                ## DEBUG
                print("LN#",filadn,"                 ##L",dn,"##: ",row[129+delta],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row[164+delta])

                sheet = blk["LINE"]
                sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                sheet['C'+str(filadn)]=action
                #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                ######################################################
                sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
                sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
                sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
                sheet['m'+str(filadn)]=row[159+delta]                # CPG
                sheet['n'+str(filadn)]=row[188+delta]                #callForwardNotRegisteredInt.destination
                sheet['o'+str(filadn)]="false"                      #callForwardNotRegisteredInt.forwardToVoiceMail
                sheet['p'+str(filadn)]=fwdnocss                     #callForwardNotRegisteredInt.callingSearchSpaceName
                sheet['q'+str(filadn)]=linept                       #routePartitionName
                sheet['r'+str(filadn)]=row[182+delta]                #callForwardOnFailure.destination
                sheet['s'+str(filadn)]="false"                      #callForwardOnFailure.forwardToVoiceMail
                sheet['t'+str(filadn)]=fwdallcss                    #callForwardOnFailure.callingSearchSpaceName
                sheet['u'+str(filadn)]="false"                      #rejectAnonymousCall
                sheet['v'+str(filadn)]="true"                       #aarKeepCallHistory
                sheet['w'+str(filadn)]=linecss                      # LINE CSS
                if row[164] != "":
                    sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row[164+delta]  # aarDestinationMask
                sheet['y'+str(filadn)]=row[176+delta]                # asciiAlertingName
                sheet['z'+str(filadn)]=row[129+delta]                # pattern
                sheet['aa'+str(filadn)]="Default"                   # patternPrecedence
                sheet['ab'+str(filadn)]=""                          # callForwardNoAnswer.duration
                sheet['ac'+str(filadn)]=row[150+delta]               # callForwardNoAnswer.destination
                sheet['ad'+str(filadn)]="false"                     # callForwardNoAnswer.forwardToVoiceMail
                sheet['ae'+str(filadn)]=fwdnocss                    # callForwardNoAnswer.callingSearchSpaceName
                sheet['ag'+str(filadn)]=row[156+delta]               # callForwardNoCoverage.destination
                sheet['ah'+str(filadn)]="false"                     # callForwardNoCoverage.forwardToVoiceMail
                sheet['ai'+str(filadn)]=fwdallcss                   # callForwardNoCoverage.callingSearchSpaceName
                sheet['aj'+str(filadn)]=row[191+delta]               # callForwardNotRegistered.destination
                sheet['ak'+str(filadn)]="false"                     # callForwardNotRegistered.forwardToVoiceMail
                sheet['al'+str(filadn)]=fwdnocss                    # callForwardNotRegistered.callingSearchSpaceName
                sheet['am'+str(filadn)]="Device"                    # usage
                sheet['ao'+str(filadn)]=row[174+delta]               #alertingName
                sheet['ap'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['aq'+str(filadn)]="false"                     #enterpriseAltNum.addLocalRoutePartition
                sheet['ar'+str(filadn)]="false"                     #enterpriseAltNum.advertiseGloballyIls
                sheet['as'+str(filadn)]=""                          #enterpriseAltNum.routePartition
                sheet['at'+str(filadn)]="false"                     #enterpriseAltNum.isUrgent
                sheet['au'+str(filadn)]=row[175+delta]              #description
                sheet['av'+str(filadn)]="false"                     #aarVoiceMailEnabled
                sheet['aw'+str(filadn)]="false"                     #useE164AltNum
                sheet['ba'+str(filadn)]="true"                      #allowCtiControlFlag
                sheet['bd'+str(filadn)]="No Error"                  #releaseClause
                sheet['be'+str(filadn)]=""                          #enterpriseAltNum.numMask
                sheet['bf'+str(filadn)]="false"                     #e164AltNum.addLocalRoutePartition
                sheet['bg'+str(filadn)]="true"                      #e164AltNum.advertiseGloballyIls
                sheet['bh'+str(filadn)]=""                          # e164AltNum.routePartition
                sheet['bi'+str(filadn)]="false"                     #e164AltNum.isUrgent
                sheet['bj'+str(filadn)]=devicecss                   # callForwardAll.secondaryCallingSearchSpaceName
                sheet['bk'+str(filadn)]=row[138+delta]               #callForwardAll.destination
                sheet['bl'+str(filadn)]="false"                     #callForwardAll.forwardToVoiceMail
                sheet['bm'+str(filadn)]=fwdallcss                   # callForwardAll.callingSearchSpaceName
                sheet['bn'+str(filadn)]="false"                     # parkMonForwardNoRetrieveVmEnabled
                sheet['bo'+str(filadn)]="true"                      # active
                sheet['bp'+str(filadn)]=""                          # VoiceMailProfileName
                sheet['bq'+str(filadn)]="false"                     # useEnterpriseAltNum
                sheet['bt'+str(filadn)]=row[141+delta]               # callForwardBusyInt.destination
                sheet['bu'+str(filadn)]="false"                     # callForwardBusyInt.forwardToVoiceMail
                sheet['bv'+str(filadn)]=fwdallcss                   # callForwardBusyInt.callingSearchSpaceName
                sheet['bw'+str(filadn)]=row[144+delta]               #  callForwardBusy.destination
                sheet['bx'+str(filadn)]="false"                     # callForwardBusy.forwardToVoiceMail
                sheet['by'+str(filadn)]=fwdallcss                   # callForwardBusy.callingSearchSpaceName
                sheet['ca'+str(filadn)]="false"                     #patternUrgency
                sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup'] #aarNeighborhoodName
                sheet['cc'+str(filadn)]="false"                     # parkMonForwardNoRetrieveIntVmEnabled
                sheet['cd'+str(filadn)]=row[153+delta]               # callForwardNoCoverageInt.destination
                sheet['ce'+str(filadn)]="false"                     # callForwardNoCoverageInt.forwardToVoiceMail
                sheet['cf'+str(filadn)]=fwdallcss                   #callForwardNoCoverageInt.callingSearchSpaceName
                sheet['cg'+str(filadn)]=""                          #callForwardAlternateParty.destination
                sheet['cj'+str(filadn)]=row[147+delta]               # callForwardNoAnswerInt.destination
                sheet['ck'+str(filadn)]="false"                     # callForwardNoAnswerInt.forwardToVoiceMail
                sheet['cl'+str(filadn)]=fwdnocss                    # callForwardNoAnswerInt.callingSearchSpaceName
                sheet['cn'+str(filadn)]="Standard Presence group"   #presenceGroupName
                filadn=filadn+1

            fila=fila+1

## CMO File INPUT DATA: Close
fgw.close()
## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)
exit(0)
