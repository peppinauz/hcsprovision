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

cl=clusterpath[14:16]       ## CL = dos digitos, 01, 02, 03,...

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
fmoserviceurl=fmoenvconfig['fmoservice0url']
fmoservicename=fmoenvconfig['fmoservice0name']
##
fmositename=data['fmosite'][0]['name']
fmositeid=data['fmosite'][0]['id']
cmg=data['fmosite'][0]['cmg']

# CMO patterns
cmodevicepool=siteslc+"-DP"
cmolocation=siteslc+"-LOC"
cmoslc="5"+siteslc

## FMO UserData
cucdmsite=fmoenvconfig['fmocustomerid']+"Si"+str(fmositeid)
cssfwd=customerid+"-DirNum-CSS"
linept=customerid+"-DirNum-PT"
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
fgw = open(inputfile,"r")
#csv_f = csv.reader(fgw)
csv_f = csv.DictReader(fgw)

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

# FMO commands:
action="add"

fila=7
filadn=7
delta=0

print("(II) IP-PHONE & ATA", file=f)

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        if row['Device Pool'].startswith(cmodevicepool) and (row['Device Name'].startswith('SEP') or row['Device Name'].startswith('ATA')):
            dn=1 ## Consideramos que siempre hay al menos una línea
            ######################################################
            ## USER MOBILITY
            ######################################################
            sheet =  blk["SUBS"]
            sheet['B'+str(fila)]=hierarchynode+"."+fmositename
            sheet['C'+str(fila)]=action
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ######################################################
            sheet['k'+str(fila)]="false"                        # imAndPresenceEnable
            sheet['l'+str(fila)]="false"                        # calendarPresence
            sheet['m'+str(fila)]="false"                        # enableMobileVoiceAccess
            sheet['p'+str(fila)]="true"                         # homeCluster
            sheet['q'+str(fila)]="10000"                        # maxDeskPickupWaitTime
            sheet['r'+str(fila)]=row['Directory Number 1']      # primaryExtension
            sheet['u'+str(fila)]="false"                        # pinCredentials.pinCredLockedByAdministrator
            sheet['v'+str(fila)]=""                             # pinCredentials.pinCredTimeAdminLockout
            sheet['w'+str(fila)]="true"                         # pinCredentials.pinCredDoesNotExpire
            sheet['x'+str(fila)]="Default Credential Policy"    # pinCredentials.pinCredPolicyName
            sheet['y'+str(fila)]="false"                        # pinCredentials.pinCredUserCantChange
            sheet['z'+str(fila)]="false"                        # pinCredentials.pinCredTimeChanged
            sheet['aa'+str(fila)]="true"                        # pinCredentials.pinCredUserMustChange
            sheet['ad'+str(fila)]="false"                       # enableUserToHostConferenceNow
            sheet['af'+str(fila)]="4"                           # remoteDestinationLimit
            sheet['ag'+str(fila)]="1"                           # status
            sheet['ai'+str(fila)]="true"                        # enableMobility
            sheet['aj'+str(fila)]="false"                       # enableEmcc
            sheet['ak'+str(fila)]=subscribecss                  # subscribeCallingSearchSpaceName
            sheet['al'+str(fila)]="false"                       # passwordCredentials.pwdCredUserMustChange
            sheet['am'+str(fila)]="false"                       # passwordCredentials.pwdCredUserCantChange
            sheet['ao'+str(fila)]="false"                       # passwordCredentials.pwdCredTimeAdminLockout
            sheet['ap'+str(fila)]="Default Credential Policy"   # passwordCredentials.pwdCredPolicyName
            sheet['aq'+str(fila)]="false"                       # passwordCredentials.pwdCredLockedByAdministrator
            sheet['ar'+str(fila)]="true"                        # passwordCredentials.pwdCredDoesNotExpire
            sheet['at'+str(fila)]="p"+row['Directory Number 1'] # HcsUserProvisioningStatusDAT.username
            sheet['ax'+str(fila)]="true"                        # enableCti
            #sheet['be'+str(fila)]=                             # mailid
            #sheet['ax'+str(fila)]=                             # phoneProfiles
            sheet['be'+str(fila)]="Standard Presence group"     # presenceGroupName
            #sheet['bf'+str(fila)]=row['FIRST NAME']              # first name
            sheet['bg'+str(fila)]=row['Description']              # lastname
            sheet['bi'+str(fila)]="p"+row['Directory Number 1'] # userid
            #sheet['bj'+str(fila)]=""                           # ctiControlledDeviceProfiles
            sheet['bk'+str(fila)]="p"+row['Directory Number 1'] # NormalizedUser.username
            sheet['bl'+str(fila)]="CUCM Local"                  # NormalizedUser.userType
            sheet['bm'+str(fila)]="p"+row['Directory Number 1'] # NormalizedUser.sn.0
            sheet['bq'+str(fila)]="Standard CCM End Users"      # associatedGroups.userGroup.0.name
            sheet['br'+str(fila)]="Standard CCM End Users"      # associatedGroups.userGroup.0.userRoles.userRole.0
            sheet['bs'+str(fila)]="Standard CCMUSER Administration" # associatedGroups.userGroup.0.userRoles.userRole.1
            sheet['bv'+str(fila)]=fmopass                       # NormalizedUser.userType
            #sheet['bw'+str(fila)]=fmopin                       # NormalizedUser.userType

            ## DEBUG
            ##print("PH#",fila,row['Device Name']," ##L",dn,"##: ",row['Directory Number 1'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 1'])

            ######################################################
            ## PHONE
            ######################################################
            sheet =  blk["PHONE"]
            sheet['B'+str(fila)]=hierarchynode+"."+fmositename
            sheet['C'+str(fila)]=action
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ######################################################
            sheet['i'+str(fila)]=""                         # directoryUrl
            sheet['j'+str(fila)]=row['Device Protocol']                    # protocol
            sheet['k'+str(fila)]=""                         # secureInformationUrl
            sheet['l'+str(fila)]="false"                    # requireDtmfReception
            if row['Phone Button Template'] != "":
                sheet['m'+str(fila)]=cl+"-"+row['Phone Button Template']          #phoneTemplateName
            #sheet['n'+str(fila)]=                          #callingSearchSpaceName (DEVICE CSS)
            sheet['p'+str(fila)]="Default"                  #useTrustedRelayPoint
            sheet['q'+str(fila)]="Brazil"                   #networkLocale
            #sheet['r'+str(fila)]="Default"                  #ringSettingBusyBlfAudibleAlert
            sheet['t'+str(fila)]="Portuguese Brazil"        #userLocale
            sheet['u'+str(fila)]="Default"                  #deviceMobilityMode
            sheet['w'+str(fila)]="No Rollover"              # outboundCallRollover
            sheet['x'+str(fila)]=""                         # ip_address
            sheet['y'+str(fila)]=""                         # primaryPhoneName
            sheet['z'+str(fila)]=row['Device Name']                     # name
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
            ####################
            ##sheet['aw'+str(fila)]=location                  #locationName
            ##sheet['ay'+str(fila)]=devicepool                #devicepool
            location=row['Location'].replace('-NOSRST-','')
            devicepool=row['Device Pool'].replace('-NOSRST-','')
            print("(II)",row['Location'],">>>",location.replace(cmolocation,cucdmsite+"-Location"), file=f)
            print("(II)",row['Device Pool'],">>>",devicepool.replace(cmodevicepool,cucdmsite+"-DevicePool"), file=f)
            sheet['aw'+str(fila)]=location.replace(cmolocation,cucdmsite+"-Location")
            sheet['ay'+str(fila)]=devicepool.replace(cmodevicepool,cucdmsite+"-DevicePool")                #devicepool
            ####################
            sheet['az'+str(fila)]="false"                   #isDualMode
            ####################
            sheet['ca'+str(fila)]="webAccess"                   #vendorConfig.13.key   ## V.11
            sheet['cb'+str(fila)]="0"                           #vendorConfig.13.value (logica inversa) ## V.11
            sheet['by'+str(fila)]="sshAccess"                   #vendorConfig.12.key   ## V.12
            sheet['bz'+str(fila)]="0"                           #vendorConfig.12.value (logica inversa) ## V.12
            sheet['fl'+str(fila)]="p"+row['Directory Number 1'] #sshUserId  ## V.12
            sheet['li'+str(fila)]=fmopass #sshUserId            ## V.12

            ####################
            sheet['cf'+str(fila)]=fmoserviceurl            # services.service.0.url
            sheet['cg'+str(fila)]=fmoservicename           # services.service.0.telecasterServiceName
            sheet['ch'+str(fila)]=fmoservicename           # services.service.0.urlLabel
            sheet['ci'+str(fila)]=fmoservicename           # services.service.0.urlLabel
            sheet['cj'+str(fila)]="1"                       # services.service.0.urlButtonIndex
            sheet['cl'+str(fila)]="No Pending Operation"    # certificateOperation
            sheet['co'+str(fila)]="false"                   #enableCallRoutingToRdWhenNoneIsActive
            #sheet['cp'+str(fila)]=row[63]                   #sipProfileName
            if row['Device Protocol'] == "SIP":
                sheet['cp'+str(fila)]="Standard SIP Profile"    #sipProfileName
            sheet['cq'+str(fila)]="false"                   #allowCtiControlFlag    @@@@ TRUE para grabación
            sheet['cr'+str(fila)]="Default"                 #singleButtonBarge
            # Depende del tipo de dispositivo
            #sheet['ct'+str(fila)]="true"                                   #enableExtensionMobility
            sheet['cy'+str(fila)]="true"                                    #useDevicePoolCgpnTransformCss
            sheet['cz'+str(fila)]="false"                                   #traceFlag
            sheet['da'+str(fila)]="Default"                                 #phoneSuite
            sheet['dc'+str(fila)]=row['Device Security Profile']            #securityProfileName
            sheet['dd'+str(fila)]="Off"                                     #joinAcrossLines
            sheet['df'+str(fila)]="false"                                   #dndStatus
            sheet['di'+str(fila)]="Use System Default"                      # networkLocation
            sheet['dj'+str(fila)]="false"                                   # unattendedPort
            ####################
            if row['Device Name'].startswith('SEP'):
                #sheet['dl'+str(fila)]=row['Directory Number 1']            #  mobilityUserIdName ## Depende de si tiene ADD-ON MODULE
                sheet['ct'+str(fila)]="true"                                #enableExtensionMobility
            ####################
            #sheet['cm'+str(fila)]="" # versionStamp
            sheet['dn'+str(fila)]="Not Trusted"                             #deviceTrustMode
            ############################################
            ##
            ## PHONE: LINE #1
            ##
            sheet['do'+str(fila)]=row['ASCII Display 1']                    #lines.line.0.displayAscii
            sheet['dp'+str(fila)]=""                                        #lines.line.0.associatedEndusers.enduser.0.userId  @@@@@@ p+EXTENSION
            sheet['dq'+str(fila)]="Ring"                                    #lines.line.0.ringSetting
            sheet['dr'+str(fila)]="Use System Default"                      #lines.line.02Setting
            #sheet['ds'+str(fila)]="Default"                                #lines.line.0.recordingProfileName
            sheet['dt'+str(fila)]=dn                                        #lines.line.0.index
            sheet['du'+str(fila)]="Use System Default"                      # lines.line.0.ringSettingActivePickupAlert
            sheet['dv'+str(fila)]=row['Line Text Label 1']                  #lines.line.0.label
            sheet['dw'+str(fila)]="Gateway Preferred"                       #lines.line.0.recordingMediaSource
            sheet['dx'+str(fila)]=row['Maximum Number of Calls 1']          #lines.line.0.maxNumCalls
            sheet['dy'+str(fila)]="General"                                 #lines.line.0.partitionUsage
            #sheet['dz'+str(fila)]="Call Recording Disabled"                 #lines.line.0.recordingMediaSource
            sheet['ec'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 1']            #lines.line.0.e164Mask
            sheet['ed'+str(fila)]="true"                                    #lines.line.0.missedCallLogging
            sheet['ee'+str(fila)]="true"                                    #lines.line.0.callInfoDisplay.dialedNumber
            sheet['ef'+str(fila)]="false"                                   #lines.line.0.callInfoDisplay.redirectedNumber
            sheet['eg'+str(fila)]="true"                                    #lines.line.0.callInfoDisplay.callerName
            sheet['eh'+str(fila)]="false"                                   #lines.line.0.callInfoDisplay.callerNumber
            sheet['ei'+str(fila)]=row['Directory Number 1']                 #lines.line.0.dirn.pattern

            ## CSS depende de la PT
            if row['CSS'] == "FAX-PSTN-CSS": ## FAXES
                linept=customerid+"-DirNum-PT"
                linecss=linefaxcss
                devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                fwdnoregcss=customerid+"-DirnumEM-CSS"
                fwdnocss=customerid+"-DirnumEM-CSS"
                fwdallcss=cucdmsite+"-InternalOnly-CSS"
            elif row['Route Partition 1'] == "Interna-EM-PT": # Logica inversa
                ## PT + CSS: EM+Phones sin EM
                linept=customerid+"-DirNum-PT"
                linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                fwdnoregcss=customerid+"-DirnumEM-CSS"
                fwdnocss=customerid+"-DirnumEM-CSS"
                fwdallcss=cucdmsite+"-InternalOnly-CSS"
            else:
                ## PT + CSS: Phones con EM
                linept=customerid+"-DirNumEM-PT"
                linecss=cucdmsite+"-DBRSTDNatl24HrsCLIPyFONnFACnCMC-CSS"
                devicecss=cucdmsite+"-BRADP-DBRDevice-CSS"
                fwdnoregcss=customerid+"-Dirnum-CSS"
                fwdnocss=cucdmsite+"-InternalOnly-CSS"
                fwdallcss=cucdmsite+"-InternalOnly-CSS"

            sheet['ej'+str(fila)]=linept                    #lines.line.0.dirn.routePartitionName
            sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
            #####################
            sheet['ek'+str(fila)]=""                        #lines.line.0.mwlPolicy
            sheet['el'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
            sheet['em'+str(fila)]=row['Busy Trigger 1']     #lines.line.0.busyTrigger
            sheet['en'+str(fila)]="Default"                 # lines.line.0.audibleMwi
            sheet['eo'+str(fila)]=row['Display 1']          # lines.line.0.display
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
            sheet['fe'+str(fila)]=row['Device Type']                   # product
            sheet['ff'+str(fila)]=row['Description']                    # description
            sheet['fg'+str(fila)]="false"                   # sendGeoLocation
            sheet['fh'+str(fila)]="p"+row['Directory Number 1']              # ownerUserName    @@@@@
            sheet['fi'+str(fila)]="false"                   # ignorePresentationIndicators
            ######################################################
            sheet['fm'+str(fila)]="Phone"                   # class
            sheet['fn'+str(fila)]="Use Common Phone Profile Setting"     # dndOption
            sheet['fo'+str(fila)]="Standard Presence group" # presenceGroupName
            sheet['fq'+str(fila)]="Standard Common Phone Profile"        # commonPhoneConfigName
            sheet['fs'+str(fila)]="711ulaw"                 # mtpPreferedCodec
            if row['Softkey Template'] != "":
                sheet['ft'+str(fila)]=cl+"-"+row['Softkey Template']         # softkeyTemplateName
            sheet['fu'+str(fila)]="false"                   # remoteDevice
            sheet['fv'+str(fila)]=aarcss                    # automatedAlternateRoutingCssName
            ######################################################
            # ADD ON MODULES
            # MODULE #1
            mobilityuserid=""
            if row['Module 1'] != "":
                print("PH#",row['Device Name'],"::ADD-ON MODULE #1::",row['Module 1'], file=f)
                sheet['lc'+str(fila)]=""                    #addOnModules.addOnModule.0.loadInformation
                sheet['ld'+str(fila)]=row['Module 1']       #addOnModules.addOnModule.0.model
                sheet['le'+str(fila)]="1"                   #addOnModules.addOnModule.0.index
                ## Si tiene ADD-ON Module -> Telefonista
                ## Chequeamos si el número termina "1111", mobilityUserIdName corresponde con el número de cabecera
                if row['Directory Number 1'].endswith("1111"): ## Si el DN termina en 1111, el mobilityID corresponde con el número de cabecera
                    ## Tenemos que sacar el número del campo Description
                    if row['Description'].startswith(cmoslc):
                        mobilityuserid=row['Description'][:9]        #  mobilityUserIdName
                    else:
                        mobilityuserid=row['Directory Number 1']     #  mobilityUserIdName
                        ## ERROR no podemos determinar el mobilityuserid de forma correcta
                        print("(EE) No podemos determinar el mobilityUserId de forma correcta::[Directory Number 1]=",row['Directory Number 1'],",[Descripton]=",row['Description'], file=f)
            else:
                if row['Device Name'].startswith('SEP'):
                    mobilityuserid=row['Directory Number 1']         #  mobilityUserIdName

            ## mobilityUserIdName
            sheet['dl'+str(fila)]=mobilityuserid         #  mobilityUserIdName
            ## DEBUG
            print("PH#",fila,row['Device Name']," ##L",dn,"##: ",row['Directory Number 1'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 1'],"::mobilityUserID::",mobilityuserid, file=f)


            # MODULE #2
            if row['Module 2'] != "":
                print("PH#",row['Device Name'],"::ADD-ON MODULE #1::",row['Module 2'], file=f)
                sheet['lf'+str(fila)]=""                    #addOnModules.addOnModule.1.loadInformation
                sheet['lg'+str(fila)]=row['Module 2']       #addOnModules.addOnModule.1.model
                sheet['lh'+str(fila)]="2"                   #addOnModules.addOnModule.1.index

            ##
            ######################################################
            ######################################################
            ##
            ## LINE: LINE #1
            ##
            ######################################################
            ## DEBUG
            print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 1'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 1'], file=f)

            sheet = blk["LINE"]
            sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
            sheet['C'+str(filadn)]=action
            #sheet['D'+str(filadn)]="name:"+row[3] # Search field
            ######################################################
            sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
            sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
            sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
            sheet['m'+str(filadn)]=row['Call Pickup Group 1']                # CPG
            sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 1']                #callForwardNotRegisteredInt.destination
            sheet['o'+str(filadn)]="false"                      #callForwardNotRegisteredInt.forwardToVoiceMail
            sheet['p'+str(filadn)]=fwdnocss                     #callForwardNotRegisteredInt.callingSearchSpaceName
            sheet['q'+str(filadn)]=linept                       #routePartitionName
            sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 1']                #callForwardOnFailure.destination
            sheet['s'+str(filadn)]="false"                      #callForwardOnFailure.forwardToVoiceMail
            sheet['t'+str(filadn)]=fwdallcss                    #callForwardOnFailure.callingSearchSpaceName
            sheet['u'+str(filadn)]="false"                      #rejectAnonymousCall
            sheet['v'+str(filadn)]="true"                       #aarKeepCallHistory
            sheet['w'+str(filadn)]=linecss                      # LINE CSS
            if row['External Phone Number Mask 1'] != "":
                sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 1']  # aarDestinationMask
            sheet['y'+str(filadn)]=row['ASCII Alerting Name 1']                # asciiAlertingName
            sheet['z'+str(filadn)]=row['Directory Number 1']                # pattern
            sheet['aa'+str(filadn)]="Default"                   # patternPrecedence
            sheet['ab'+str(filadn)]=""                          # callForwardNoAnswer.duration
            sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 1']               # callForwardNoAnswer.destination
            sheet['ad'+str(filadn)]="false"                     # callForwardNoAnswer.forwardToVoiceMail
            sheet['ae'+str(filadn)]=fwdnocss                    # callForwardNoAnswer.callingSearchSpaceName
            sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 1']               # callForwardNoCoverage.destination
            sheet['ah'+str(filadn)]="false"                     # callForwardNoCoverage.forwardToVoiceMail
            sheet['ai'+str(filadn)]=fwdallcss                   # callForwardNoCoverage.callingSearchSpaceName
            sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 1']               # callForwardNotRegistered.destination
            sheet['ak'+str(filadn)]="false"                     # callForwardNotRegistered.forwardToVoiceMail
            sheet['al'+str(filadn)]=fwdnocss                    # callForwardNotRegistered.callingSearchSpaceName
            sheet['am'+str(filadn)]="Device"                    # usage
            sheet['ao'+str(filadn)]=row['Alerting Name 1']               #alertingName
            sheet['ap'+str(filadn)]=""                          #enterpriseAltNum.numMask
            sheet['aq'+str(filadn)]="false"                     #enterpriseAltNum.addLocalRoutePartition
            sheet['ar'+str(filadn)]="false"                     #enterpriseAltNum.advertiseGloballyIls
            sheet['as'+str(filadn)]=""                          #enterpriseAltNum.routePartition
            sheet['at'+str(filadn)]="false"                     #enterpriseAltNum.isUrgent
            sheet['au'+str(filadn)]=row['Line Description 1']              #description
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
            sheet['bk'+str(filadn)]=row['Forward All Destination 1']               #callForwardAll.destination
            sheet['bl'+str(filadn)]="false"                     #callForwardAll.forwardToVoiceMail
            sheet['bm'+str(filadn)]=fwdallcss                   # callForwardAll.callingSearchSpaceName
            sheet['bn'+str(filadn)]="false"                     # parkMonForwardNoRetrieveVmEnabled
            sheet['bo'+str(filadn)]="true"                      # active
            sheet['bp'+str(filadn)]=""                          # VoiceMailProfileName
            sheet['bq'+str(filadn)]="false"                     # useEnterpriseAltNum
            sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 1']               # callForwardBusyInt.destination
            sheet['bu'+str(filadn)]="false"                     # callForwardBusyInt.forwardToVoiceMail
            sheet['bv'+str(filadn)]=fwdallcss                   # callForwardBusyInt.callingSearchSpaceName
            sheet['bw'+str(filadn)]=row['Forward Busy External Destination 1']               #  callForwardBusy.destination
            sheet['bx'+str(filadn)]="false"                     # callForwardBusy.forwardToVoiceMail
            sheet['by'+str(filadn)]=fwdallcss                   # callForwardBusy.callingSearchSpaceName
            sheet['ca'+str(filadn)]="false"                     #patternUrgency
            sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup'] #aarNeighborhoodName
            sheet['cc'+str(filadn)]="false"                     # parkMonForwardNoRetrieveIntVmEnabled
            sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 1']               # callForwardNoCoverageInt.destination
            sheet['ce'+str(filadn)]="false"                     # callForwardNoCoverageInt.forwardToVoiceMail
            sheet['cf'+str(filadn)]=fwdallcss                   #callForwardNoCoverageInt.callingSearchSpaceName
            sheet['cg'+str(filadn)]=""                          #callForwardAlternateParty.destination
            sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 1']               # callForwardNoAnswerInt.destination
            sheet['ck'+str(filadn)]="false"                     # callForwardNoAnswerInt.forwardToVoiceMail
            sheet['cl'+str(filadn)]=fwdnocss                    # callForwardNoAnswerInt.callingSearchSpaceName
            sheet['cn'+str(filadn)]="Standard Presence group"   #presenceGroupName
            filadn=filadn+1

            ##
            ## LINE #2
            ##
            if data['e164'][0]['phonedn'] >= 2:
                if row['Directory Number 2'][1:].startswith(siteslc):
                    ##
                    dn=dn+1
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 2'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 2'], file=f)

                    ######################################################
                    ## PHONE
                    ######################################################
                    sheet =  blk["PHONE"]
                    ############################################
                    ##
                    ## PHONE: LINE #2
                    ##
                    sheet['fx'+str(fila)]=row['ASCII Display 2']                    #lines.line.0.displayAscii
                    sheet['fy'+str(fila)]=""                                        #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['fz'+str(fila)]="Ring"                                    #lines.line.0.ringSetting
                    sheet['ga'+str(fila)]="Use System Default"                      #lines.line.0.consecutiveRingSetting
                    #sheet['gb'+str(fila)]="Default"                                #lines.line.0.recordingProfileName
                    sheet['gc'+str(fila)]=dn                                        #lines.line.0.index
                    sheet['gd'+str(fila)]="Use System Default"                      # lines.line.0.ringSettingActivePickupAlert
                    sheet['ge'+str(fila)]=row['Line Text Label 2']                  #lines.line.0.label
                    sheet['gf'+str(fila)]="Gateway Preferred"                       #lines.line.0.recordingMediaSource
                    sheet['gg'+str(fila)]=row['Maximum Number of Calls 2']          #lines.line.0.maxNumCalls
                    sheet['gh'+str(fila)]="General"                                 #lines.line.0.partitionUsage
                    #sheet['gi'+str(fila)]="Call Recording Disabled"                 #lines.line.0.recordingMediaSource
                    sheet['gl'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 2']            #lines.line.0.e164Mask
                    sheet['gm'+str(fila)]="true"                                    #lines.line.0.missedCallLogging
                    sheet['gn'+str(fila)]="true"                                    #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['go'+str(fila)]="false"                                   #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['gp'+str(fila)]="true"                                    #lines.line.0.callInfoDisplay.callerName
                    sheet['gq'+str(fila)]="false"                                   #lines.line.0.callInfoDisplay.callerNumber
                    sheet['gr'+str(fila)]=row['Directory Number 2']                 #lines.line.0.dirn.pattern

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
                    sheet['gt'+str(fila)]=""                        #lines.line.0.mwlPolicy
                    sheet['gu'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                    sheet['gv'+str(fila)]=row['Busy Trigger 2']            #lines.line.0.busyTrigger
                    sheet['gw'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                    sheet['gx'+str(fila)]=row['Display 2']            # lines.line.0.display
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #2
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 2'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 2'], file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                                        # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                                             # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                                                # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 2']                                       # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 2']               #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                                          #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                                         #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                                           #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 2']                      #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                                          #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                                        #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                                          #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                                           #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                                          # LINE CSS
                    if row['External Phone Number Mask 2'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 2']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 2']                                     # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 2']                                        # pattern
                    sheet['aa'+str(filadn)]="Default"                                                       # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                                              # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 2']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                                         # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 2']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                                         # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                                       # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 2']              # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                                         # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                                        # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                                        # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 2']                                          #alertingName
                    sheet['ap'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                                         #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                                         #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                                              #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                                         #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 2']                                       #description
                    sheet['av'+str(filadn)]="false"                                                         #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                                         #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                                          #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                                      #releaseClause
                    sheet['be'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                                         #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                                          #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                                              # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                                         #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                                       # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 2']                                #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                                         #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                                       # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                                          # active
                    sheet['bp'+str(filadn)]=""                                                              # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                                         # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 2']                      # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                                         # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                                       # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 2']                      #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                                         # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                                       # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                                         #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                                     #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 2']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                                         # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                                       #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                                              #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 2']                 # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                                         # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                                       #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #3
            ##
            if data['e164'][0]['phonedn'] >= 3:
                if row['Directory Number 3'][1:].startswith(siteslc):
                    ##
                    dn=dn+1
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 3'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 3'], file=f)

                    ######################################################
                    ## PHONE
                    ######################################################
                    sheet =  blk["PHONE"]
                    ############################################
                    ##
                    ## PHONE: LINE #3
                    ##
                    sheet['gy'+str(fila)]=row['ASCII Display 3']            #lines.line.2.displayAscii
                    sheet['gz'+str(fila)]=""                                #lines.line.2.associatedEndusers.enduser.0.userId
                    sheet['ha'+str(fila)]="Ring"                            #lines.line.0.ringSetting
                    sheet['hb'+str(fila)]="Use System Default"              #lines.line.0.consecutiveRingSetting
                    #sheet['hc'+str(fila)]=""                                #lines.line.0.recordingProfileName
                    sheet['hd'+str(fila)]=dn                                #lines.line.0.index
                    sheet['he'+str(fila)]="Use System Default"              # lines.line.0.ringSettingActivePickupAlert
                    sheet['hf'+str(fila)]=row['Line Text Label 3']          #lines.line.0.label
                    sheet['hg'+str(fila)]="Gateway Preferred"               #lines.line.0.recordingMediaSource
                    sheet['hh'+str(fila)]=row['Maximum Number of Calls 3']  #lines.line.0.maxNumCalls
                    sheet['hi'+str(fila)]="General"                         #lines.line.0.partitionUsage
                    #sheet['hj'+str(fila)]="Call Recording Disabled"         #lines.line.0.recordingMediaSource
                    sheet['hk'+str(fila)]=""
                    sheet['hl'+str(fila)]=""
                    sheet['hm'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 3']            #lines.line.0.e164Mask
                    sheet['hn'+str(fila)]="true"                            #lines.line.0.missedCallLogging
                    sheet['ho'+str(fila)]="true"                            #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['hp'+str(fila)]="false"                           #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['hq'+str(fila)]="true"                            #lines.line.0.callInfoDisplay.callerName
                    sheet['hr'+str(fila)]="false"                           #lines.line.0.callInfoDisplay.callerNumber
                    sheet['hs'+str(fila)]=row['Directory Number 3']            #lines.line.0.dirn.pattern

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

                    sheet['ht'+str(fila)]=linept                    #lines.line.2.dirn.routePartitionName
                    #sheet['n'+str(fila)]=devicecss                  #callingSearchSpaceName (DEVICE CSS)
                    #####################
                    sheet['hu'+str(fila)]=""       #lines.line.0.mwlPolicy
                    sheet['hv'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                    sheet['hw'+str(fila)]=row['Busy Trigger 3']            #lines.line.0.busyTrigger
                    sheet['hx'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                    sheet['hy'+str(fila)]=row['Display 3']            # lines.line.0.display
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #2
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 3'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 3'], file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                                        # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                                             # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                                                # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 3']                                       # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 3']               #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                                          #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                                         #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                                           #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 3']                      #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                                          #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                                        #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                                          #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                                           #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                                          # LINE CSS
                    if row['External Phone Number Mask 3'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 3']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 3']                                     # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 3']                                        # pattern
                    sheet['aa'+str(filadn)]="Default"                                                       # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                                              # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 3']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                                         # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 3']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                                         # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                                       # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 3']              # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                                         # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                                        # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                                        # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 3']                                          #alertingName
                    sheet['ap'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                                         #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                                         #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                                              #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                                         #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 3']                                       #description
                    sheet['av'+str(filadn)]="false"                                                         #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                                         #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                                          #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                                      #releaseClause
                    sheet['be'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                                         #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                                          #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                                              # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                                         #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                                       # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 3']                                #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                                         #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                                       # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                                          # active
                    sheet['bp'+str(filadn)]=""                                                              # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                                         # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 3']                      # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                                         # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                                       # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 3']                      #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                                         # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                                       # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                                         #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                                     #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 3']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                                         # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                                       #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                                              #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 3']                 # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                                         # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                                       #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #4
            ##
            if data['e164'][0]['phonedn'] >= 4:
                if row['Directory Number 4'][1:].startswith(siteslc):
                    ##
                    dn=dn+1
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 4'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 4'], file=f)

                    ######################################################
                    ## PHONE
                    ######################################################
                    sheet =  blk["PHONE"]
                    ############################################
                    ##
                    ## PHONE: LINE #2
                    ##
                    sheet['hz'+str(fila)]=row['ASCII Display 4']            #lines.line.0.displayAscii
                    sheet['ia'+str(fila)]=""                                #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['ib'+str(fila)]="Ring"                            #lines.line.0.ringSetting
                    sheet['ic'+str(fila)]="Use System Default"              #lines.line.0.consecutiveRingSetting
                    #sheet['id'+str(fila)]=""                                #lines.line.0.recordingProfileName
                    sheet['ie'+str(fila)]=dn                                #lines.line.0.index
                    sheet['if'+str(fila)]="Use System Default"              # lines.line.0.ringSettingActivePickupAlert
                    sheet['ig'+str(fila)]=row['Line Text Label 4']          #lines.line.0.label
                    sheet['ih'+str(fila)]="Gateway Preferred"               #lines.line.0.recordingMediaSource
                    sheet['ii'+str(fila)]=row['Maximum Number of Calls 4']  #lines.line.0.maxNumCalls
                    sheet['ij'+str(fila)]="General"                         #lines.line.0.partitionUsage
                    #sheet['ik'+str(fila)]="Call Recording Disabled"         #lines.line.0.recordingMediaSource
                    sheet['il'+str(fila)]=""
                    sheet['im'+str(fila)]=""
                    sheet['in'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 4']            #lines.line.0.e164Mask
                    sheet['io'+str(fila)]="true"                    #lines.line.0.missedCallLogging
                    sheet['ip'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['iq'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['ir'+str(fila)]="true"                    #lines.line.0.callInfoDisplay.callerName
                    sheet['is'+str(fila)]="false"                   #lines.line.0.callInfoDisplay.callerNumber
                    sheet['it'+str(fila)]=row['Directory Number 4']            #lines.line.0.dirn.pattern

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
                    sheet['iv'+str(fila)]=""       #lines.line.0.mwlPolicy
                    sheet['iw'+str(fila)]="Use System Default"      #lines.line.0.ringSettingIdlePickupAlert
                    sheet['ix'+str(fila)]=row['Busy Trigger 4']            #lines.line.0.busyTrigger
                    sheet['iy'+str(fila)]="Default"                 # lines.line.0.audibleMwi
                    sheet['iz'+str(fila)]=row['Display 4']            # lines.line.0.display
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #4
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 4'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 4'], file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                                        # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                                             # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                                                # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 4']                                       # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 4']               #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                                          #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                                         #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                                           #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 4']                      #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                                          #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                                        #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                                          #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                                           #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                                          # LINE CSS
                    if row['External Phone Number Mask 4'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 4']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 4']                                     # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 4']                                        # pattern
                    sheet['aa'+str(filadn)]="Default"                                                       # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                                              # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 4']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                                         # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 4']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                                         # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                                       # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 4']              # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                                         # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                                        # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                                        # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 4']                                          #alertingName
                    sheet['ap'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                                         #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                                         #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                                              #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                                         #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 4']                                       #description
                    sheet['av'+str(filadn)]="false"                                                         #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                                         #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                                          #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                                      #releaseClause
                    sheet['be'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                                         #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                                          #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                                              # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                                         #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                                       # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 4']                                #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                                         #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                                       # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                                          # active
                    sheet['bp'+str(filadn)]=""                                                              # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                                         # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 4']                      # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                                         # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                                       # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 4']                      #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                                         # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                                       # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                                         #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                                     #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 4']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                                         # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                                       #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                                              #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 4']                 # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                                         # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                                       #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #5
            ##
            if data['e164'][0]['phonedn'] >= 5:
                if row['Directory Number 5'][1:].startswith(siteslc):
                    ##
                    dn=dn+1
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 5'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 5'], file=f)

                    ######################################################
                    ## PHONE
                    ######################################################
                    sheet =  blk["PHONE"]
                    ############################################
                    ##
                    ## PHONE: LINE #2
                    ##
                    sheet['ja'+str(fila)]=row['ASCII Display 5']            #lines.line.0.displayAscii
                    sheet['jb'+str(fila)]=""                                #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['jc'+str(fila)]="Ring"                            #lines.line.0.ringSetting
                    sheet['jd'+str(fila)]="Use System Default"              #lines.line.0.consecutiveRingSetting
                    #sheet['je'+str(fila)]=""                                #lines.line.0.recordingProfileName
                    sheet['jf'+str(fila)]=dn                                #lines.line.0.index
                    sheet['jg'+str(fila)]="Use System Default"              # lines.line.0.ringSettingActivePickupAlert
                    sheet['jh'+str(fila)]=row['Line Text Label 5']          #lines.line.0.label
                    sheet['ji'+str(fila)]="Gateway Preferred"               #lines.line.0.recordingMediaSource
                    sheet['jj'+str(fila)]=row['Maximum Number of Calls 5']  #lines.line.0.maxNumCalls
                    sheet['jk'+str(fila)]="General"                         #lines.line.0.partitionUsage
                    sheet['jl'+str(fila)]=""
                    sheet['jm'+str(fila)]=""
                    #sheet['jn'+str(fila)]="Call Recording Disabled"         #lines.line.0.recordingMediaSource
                    sheet['jo'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 5']            #lines.line.0.e164Mask
                    sheet['jp'+str(fila)]="true"                            #lines.line.0.missedCallLogging
                    sheet['jq'+str(fila)]="true"                            #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['jr'+str(fila)]="false"                           #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['js'+str(fila)]="true"                            #lines.line.0.callInfoDisplay.callerName
                    sheet['jt'+str(fila)]="false"                           #lines.line.0.callInfoDisplay.callerNumber
                    sheet['ju'+str(fila)]=row['Directory Number 5']         #lines.line.0.dirn.pattern

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

                    sheet['jv'+str(fila)]=linept                            #lines.line.0.dirn.routePartitionName
                    #sheet['n'+str(fila)]=devicecss                         #callingSearchSpaceName (DEVICE CSS)
                    #####################
                    sheet['jw'+str(fila)]=""               #lines.line.0.mwlPolicy
                    sheet['jx'+str(fila)]="Use System Default"              #lines.line.0.ringSettingIdlePickupAlert
                    sheet['jy'+str(fila)]=row['Busy Trigger 5']             #lines.line.0.busyTrigger
                    sheet['jz'+str(fila)]="Default"                         # lines.line.0.audibleMwi
                    sheet['ka'+str(fila)]=row['Display 5']                  # lines.line.0.display
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #4
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 5'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 5'], file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                                        # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                                             # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                                                # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 5']                                       # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 5']               #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                                          #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                                         #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                                           #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 5']                      #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                                          #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                                        #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                                          #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                                           #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                                          # LINE CSS
                    if row['External Phone Number Mask 5'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 5']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 5']                                     # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 5']                                        # pattern
                    sheet['aa'+str(filadn)]="Default"                                                       # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                                              # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 5']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                                         # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 5']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                                         # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                                       # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 5']              # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                                         # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                                        # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                                        # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 5']                                          #alertingName
                    sheet['ap'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                                         #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                                         #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                                              #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                                         #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 5']                                       #description
                    sheet['av'+str(filadn)]="false"                                                         #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                                         #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                                          #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                                      #releaseClause
                    sheet['be'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                                         #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                                          #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                                              # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                                         #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                                       # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 5']                                #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                                         #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                                       # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                                          # active
                    sheet['bp'+str(filadn)]=""                                                              # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                                         # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 5']                      # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                                         # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                                       # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 5']                      #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                                         # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                                       # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                                         #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                                     #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 5']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                                         # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                                       #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                                              #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 5']                 # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                                         # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                                       #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #6
            ##
            if data['e164'][0]['phonedn'] >= 6:
                if row['Directory Number 6'][1:].startswith(siteslc):
                    ##
                    dn=dn+1
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 6'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 6'], file=f)

                    ######################################################
                    ## PHONE
                    ######################################################
                    sheet =  blk["PHONE"]
                    ############################################
                    ##
                    ## PHONE: LINE #6
                    ##
                    sheet['kb'+str(fila)]=row['ASCII Display 6']                    #lines.line.0.displayAscii
                    sheet['kc'+str(fila)]=""                                        #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['kd'+str(fila)]="Ring"                                    #lines.line.0.ringSetting
                    sheet['ke'+str(fila)]="Use System Default"                      #lines.line.0.consecutiveRingSetting
                    #sheet['kf'+str(fila)]=""                                       #lines.line.0.recordingProfileName
                    sheet['kg'+str(fila)]=dn                                        #lines.line.0.index
                    sheet['kh'+str(fila)]="Use System Default"                      # lines.line.0.ringSettingActivePickupAlert
                    sheet['ki'+str(fila)]=row['Line Text Label 6']                  #lines.line.0.label
                    sheet['kj'+str(fila)]="Gateway Preferred"                       #lines.line.0.recordingMediaSource
                    sheet['kk'+str(fila)]=row['Maximum Number of Calls 6']          #lines.line.0.maxNumCalls
                    sheet['kl'+str(fila)]="General"                                 #lines.line.0.partitionUsage
                    #sheet['km'+str(fila)]="Call Recording Disabled"                #lines.line.0.recordingMediaSource
                    sheet['kn'+str(fila)]=""
                    sheet['ko'+str(fila)]=""
                    sheet['kp'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 6']            #lines.line.0.e164Mask
                    sheet['kq'+str(fila)]="true"                                    #lines.line.0.missedCallLogging
                    sheet['kr'+str(fila)]="true"                                    #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['ks'+str(fila)]="false"                                   #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['kt'+str(fila)]="true"                                    #lines.line.0.callInfoDisplay.callerName
                    sheet['ku'+str(fila)]="false"                                   #lines.line.0.callInfoDisplay.callerNumber
                    sheet['kv'+str(fila)]=row['Directory Number 6']                 #lines.line.0.dirn.pattern

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

                    sheet['kw'+str(fila)]=linept                                    #lines.line.0.dirn.routePartitionName
                    #sheet['n'+str(fila)]=devicecss                                 #callingSearchSpaceName (DEVICE CSS)
                    #####################
                    sheet['kw'+str(fila)]=""                                        #lines.line.0.mwlPolicy
                    sheet['ky'+str(fila)]="Use System Default"                      #lines.line.0.ringSettingIdlePickupAlert
                    sheet['kz'+str(fila)]=row['Busy Trigger 6']                     #lines.line.0.busyTrigger
                    sheet['la'+str(fila)]="Default"                                 # lines.line.0.audibleMwi
                    sheet['lb'+str(fila)]=row['Display 6']                          # lines.line.0.display
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #4
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                 ##L",dn,"##: ",row['Directory Number 6'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 6'], file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                                        # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                                             # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                                                # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 6']                                       # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 6']               #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                                          #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                                         #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                                           #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 6']                      #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                                          #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                                        #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                                          #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                                           #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                                          # LINE CSS
                    if row['External Phone Number Mask 6'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 6']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 6']                                     # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 6']                                        # pattern
                    sheet['aa'+str(filadn)]="Default"                                                       # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                                              # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 6']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                                         # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 6']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                                         # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                                       # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 6']              # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                                         # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                                        # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                                        # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 6']                                          #alertingName
                    sheet['ap'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                                         #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                                         #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                                              #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                                         #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 6']                                       #description
                    sheet['av'+str(filadn)]="false"                                                         #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                                         #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                                          #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                                      #releaseClause
                    sheet['be'+str(filadn)]=""                                                              #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                                         #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                                          #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                                              # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                                         #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                                       # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 6']                                #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                                         #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                                       # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                                          # active
                    sheet['bp'+str(filadn)]=""                                                              # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                                         # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 6']                      # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                                         # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                                       # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 6']                      #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                                         # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                                       # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                                         #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                                     #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                                         # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 6']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                                         # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                                       #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                                              #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 6']                 # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                                         # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                                        # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                                       #presenceGroupName
                    filadn=filadn+1


            fila=fila+1

## CMO File INPUT DATA: Close
fgw.close()
## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)
## LOG de CONFIGURACION
f.close()

exit(0)
