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
#inputfile = clusterpath+"/directorynumber.csv"  ## ORIGINAL
inputfile = clusterpath+"/deviceprofile.csv"

templateblkfile = "../code/blk/04.devprof-template.xlsx" # SIN DATAINPUT
outputblkfile = sitepath+"/05.devprof."+siteslc+".xlsx"

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

# CMO patterns
cmodevprof="5"+siteslc

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
action="add"
fmopass="vi123456"
fmopin="123456"


# CMO File INPUT DATA
fgw = open(inputfile,"r")
csv_f = csv.DictReader(fgw)

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

fila=7
filadn=7
dn=1            ## Consideramos que siempre hay al menos una línea

print("(II) DEVICE PROFILE", file=f)

for row in csv_f:
    # WR BLK OUTPUT DATA
    if len(row) != 0: # Me salto las líneas vacias
        if row['Device Profile Name'].startswith(cmodevprof):
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
            sheet['aa'+str(fila)]="false"                       # pinCredentials.pinCredUserMustChange
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
            sheet['at'+str(fila)]=row['User ID 1']              # HcsUserProvisioningStatusDAT.username
            sheet['ax'+str(fila)]="true"                        # enableCti
            #sheet['be'+str(fila)]=                             # mailid
            #sheet['ax'+str(fila)]=                             # phoneProfiles
            sheet['be'+str(fila)]="Standard Presence group"     # presenceGroupName
            if row['Display 1'] != "":
                sheet['bg'+str(fila)]=row['Display 1']                       # lastname
            else:
                sheet['bg'+str(fila)]=row['Alerting Name 1']
            sheet['bi'+str(fila)]=row['User ID 1']              # userid
            #sheet['bj'+str(fila)]=""                           # ctiControlledDeviceProfiles
            sheet['bk'+str(fila)]=row['User ID 1']              # NormalizedUser.username
            sheet['bl'+str(fila)]="CUCM Local"                  # NormalizedUser.userType
            sheet['bm'+str(fila)]=row['User ID 1']              # NormalizedUser.sn.0
            sheet['bq'+str(fila)]="Standard CCM End Users"      # associatedGroups.userGroup.0.name
            sheet['br'+str(fila)]="Standard CCM End Users"      # associatedGroups.userGroup.0.userRoles.userRole.0
            sheet['bs'+str(fila)]="Standard CCMUSER Administration" # associatedGroups.userGroup.0.userRoles.userRole.1
            sheet['bv'+str(fila)]=fmopass                       # NormalizedUser.userType
            sheet['bw'+str(fila)]=fmopin                        # NormalizedUser.userType

            sheet =  blk["SUBS.ASOC"]
            sheet['B'+str(fila)]=hierarchynode+"."+fmositename
            sheet['C'+str(fila)]="modify"
            sheet['D'+str(fila)]="userid:"+row['User ID 1']     # Search field
            ######################################################
            sheet['bi'+str(fila)]=row['User ID 1']                       # userid
            sheet['bk'+str(fila)]=row['User ID 1']                       # NormalizedUser.username
            sheet['at'+str(fila)]=row['User ID 1']                       # HcsUserProvisioningStatusDAT.username
            sheet['bx'+str(fila)]=row['Device Profile Name']                        # phoneProfiles.profileName.0
            sheet['by'+str(fila)]=row['Directory Number 1']              # primaryExtension.pattern
            sheet['bz'+str(fila)]=linept                                # primaryExtension.routePartitionName


            ## DEBUG
            print("EM#",fila,row['Device Profile Name'] ," ##L",dn,"##: ",row['Directory Number 1'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],file=f)

            ######################################################
            ## EM
            ######################################################
            sheet =  blk["EM"]
            sheet['B'+str(fila)]=hierarchynode+"."+fmositename
            #sheet['C'+str(fila)]=action
            #sheet['D'+str(fila)]="name:"+row[3] # Search field
            ######################################################
            sheet['i'+str(fila)]=row['Device Protocol']                     # protocol
            sheet['j'+str(fila)]="Default"                                  # alwaysUsePrimeLineForVoiceMessage
            if row['Softkey Template'] != "":
                sheet['k'+str(fila)]=cl+"-"+row['Softkey Template']         # softkeyTemplateName
            if row['Phone Button Template'] != "":
                sheet['m'+str(fila)]=cl+"-"+row['Phone Button Template']    # phoneTemplateName
            sheet['o'+str(fila)]="Default"                                  # preemption
            sheet['q'+str(fila)]="Default"                                  # singleButtonBarge
            sheet['s'+str(fila)]="Off"                                      # mlppIndicationStatus
            sheet['u'+str(fila)]=row['Device Type']                         # Product
            sheet['v'+str(fila)]=row['Description']                         # Description
            sheet['w'+str(fila)]="false"                                    # traceFlag
            sheet['x'+str(fila)]="false"                                    # ignorePresentationIndicators
            sheet['y'+str(fila)]=row['Device User Locale']                  # userLocale
            sheet['z'+str(fila)]="Default"                                  # joinAcrossLines
            ####################
            #sheet['aa'+str(fila)]="Device Profile"                         # speeddials
            sheet['ab'+str(fila)]="Device Profile"                          # Class
            sheet['ac'+str(fila)]="Ringer Off"                              # dndOption
            sheet['ad'+str(fila)]="false"                                   # dndStatus
            sheet['ae'+str(fila)]=row['Device Profile Name']                # name
            sheet['af'+str(fila)]=""                                        # dndRingSetting
            sheet['ag'+str(fila)]="Default"                                 # alwaysUsePrimeLine
            sheet['ah'+str(fila)]="User"                                    # protocolSide
            sheet['ai'+str(fila)]="Default"                                 # callInfoPrivacyStatus
            ############################################
            # ADD ON MODULES
            # MODULE #1
            if row['Module 1'] != "":
                print("PH#",row['Device Profile Name'],"::ADD-ON MODULE #1::",row['Module 1'],file=f)
                sheet['gx'+str(fila)]=""                    #addOnModules.addOnModule.0.loadInformation
                sheet['gy'+str(fila)]=row['Module 1']       #addOnModules.addOnModule.0.model
                sheet['gz'+str(fila)]="1"                   #addOnModules.addOnModule.0.index
            # MODULE #2
            if row['Module 2'] != "":
                print("PH#",row['Device Profile Name'],"::ADD-ON MODULE #2::",row['Module 2'],file=f)
                sheet['ha'+str(fila)]=""                    #addOnModules.addOnModule.1.loadInformation
                sheet['hb'+str(fila)]=row['Module 2']       #addOnModules.addOnModule.1.model
                sheet['hc'+str(fila)]="2"                   #addOnModules.addOnModule.1.index
            ##
            ## EM: LINE #1
            ##
            sheet['aj'+str(fila)]=row['ASCII Display 1']                    #lines.line.0.displayAscii
            sheet['ak'+str(fila)]=row['User ID 1']                          #lines.line.0.associatedEndusers.enduser.0.userId
            sheet['al'+str(fila)]="Ring"                                    #lines.line.0.ringSetting
            sheet['am'+str(fila)]="Use System Default"                      #lines.line.0.consecutiveRingSetting
            #sheet['an'+str(fila)]="Default"                                # lines.line.0.recordingProfileName
            sheet['ao'+str(fila)]=dn                                        #lines.line.0.index
            #sheet['ap'+str(fila)]="Use System Default"                     # lines.line.0.ringSettingActivePickupAlert
            sheet['aq'+str(fila)]=row['Line Text Label 1']                  #lines.line.0.label
            sheet['ar'+str(fila)]="Gateway Preferred"                       #lines.line.0.recordingMediaSource
            sheet['as'+str(fila)]=row['Maximum Number of Calls 1']          #lines.line.0.maxNumCalls
            sheet['at'+str(fila)]="General"                                 #lines.line.0.partitionUsage
            sheet['au'+str(fila)]="Call Recording Disabled"                 #lines.line.0.recordingFlag
            sheet['av'+str(fila)]=""                                        #lines.line.0.speedDial
            sheet['aw'+str(fila)]=""                                        #lines.line.0.monitoringCssName
            if row['External Phone Number Mask 1'] != "":
                sheet['ax'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 1']  #lines.line.0.e164Mask
            sheet['ay'+str(fila)]="true"                                    #lines.line.0.missedCallLogging
            sheet['az'+str(fila)]="true"                                    #lines.line.0.callInfoDisplay.dialedNumber
            sheet['ba'+str(fila)]="false"                                   #lines.line.0.callInfoDisplay.redirectedNumber
            sheet['bb'+str(fila)]="false"                                   # lines.line.0.callInfoDisplay.callerName
            sheet['bc'+str(fila)]="true"                                    # lines.line.0.callInfoDisplay.callerNumber
            sheet['bd'+str(fila)]=row['Directory Number 1']                 #lines.line.0.dirn.pattern
            sheet['be'+str(fila)]=linept                                    #lines.line.0.dirn.routePartitionName
            sheet['bf'+str(fila)]="Use System Policy"                       #lines.line.0.mwlPolicy
            sheet['bg'+str(fila)]=""                                        #lines.line.0.ringSettingIdlePickupAlert
            sheet['bh'+str(fila)]=row['Busy Trigger 1']                     #lines.line.0.busyTrigger
            sheet['bi'+str(fila)]="Default"                                 # lines.line.0.audibleMwi
            sheet['bj'+str(fila)]=row['Display 1']                          # lines.line.0.display
            ######################
            sheet['bk'+str(fila)]=""                                        # services
            sheet['bl'+str(fila)]=""                                        # loginUserId
            sheet['bm'+str(fila)]=""                                        # emccCallingSearchSpace
            ##
            ######################################################
            sheet['gs'+str(fila)]=fmoserviceurl                             # services.service.0.url
            sheet['gt'+str(fila)]=fmoservicename                            # services.service.0.telecasterServiceName
            sheet['gu'+str(fila)]=fmoservicename                            # services.service.0.urlLabel
            sheet['gv'+str(fila)]=fmoservicename                            # services.service.0.urlLabel
            sheet['gw'+str(fila)]="1"                                       # services.service.0.urlButtonIndex
            ######################################################
            ######################################################
            ##
            ## LINE: LINE #1
            ##
            ######################################################
            ## DEBUG
            print("LN#",filadn,"                ##L",dn,"##: ",row['Directory Number 1'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 1'],file=f)

            sheet = blk["LINE"]
            sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
            sheet['C'+str(filadn)]=action
            #sheet['D'+str(filadn)]="name:"+row[3] # Search field
            ######################################################
            sheet['i'+str(filadn)]="Default"                    # partyEntranceTone
            sheet['j'+str(filadn)]="Use System Default"         # cfaCssPolicy
            sheet['k'+str(filadn)]="Auto Answer Off"            # autoAnswer
            sheet['m'+str(filadn)]=row['Call Pickup Group 1']   # CPG
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
            sheet['au'+str(filadn)]=row['Line Description 1']               #description
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
            if data['e164'][0]['emdn'] >= 2:
                if row['Directory Number 2']!="":
                    dn=dn+1
                    ## DEBUG
                    print("EM#",fila,"                ##L",dn,"##: ",row['Directory Number 2'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 2'],file=f)
                    ##
                    sheet =  blk["EM"]
                    ##
                    sheet['bn'+str(fila)]=row['ASCII Display 2']                        #lines.line.0.displayAscii
                    sheet['bo'+str(fila)]=row['User ID 1']                              #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['bp'+str(fila)]="Ring"                                        #lines.line.0.ringSetting
                    sheet['bq'+str(fila)]="Use System Default"                          #lines.line.0.consecutiveRingSetting
                    #sheet['br'+str(fila)]="Default"                                    # lines.line.0.recordingProfileName
                    sheet['bs'+str(fila)]=dn                                            #lines.line.0.index
                    #sheet['bt'+str(fila)]="Use System Default"                         # lines.line.0.ringSettingActivePickupAlert
                    sheet['bu'+str(fila)]=row['Line Text Label 2']                      #lines.line.0.label
                    sheet['bv'+str(fila)]="Gateway Preferred"                           #lines.line.0.recordingMediaSource
                    sheet['bw'+str(fila)]=row['Maximum Number of Calls 2']              #lines.line.0.maxNumCalls
                    sheet['bx'+str(fila)]="General"                                     #lines.line.0.partitionUsage
                    sheet['by'+str(fila)]="Call Recording Disabled"                     #lines.line.0.recordingFlag
                    sheet['bz'+str(fila)]=""                                            #lines.line.0.speedDial
                    sheet['ca'+str(fila)]=""                                            #lines.line.0.monitoringCssName
                    if row['External Phone Number Mask 2'] != "":
                        sheet['cb'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 2']  #lines.line.0.e164Mask
                    sheet['cc'+str(fila)]="true"                                        #lines.line.0.missedCallLogging
                    sheet['cd'+str(fila)]="true"                                        #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['ce'+str(fila)]="false"                                       #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['cf'+str(fila)]="false"                                       # lines.line.0.callInfoDisplay.callerName
                    sheet['cg'+str(fila)]="true"                                        # lines.line.0.callInfoDisplay.callerNumber
                    sheet['ch'+str(fila)]=row['Directory Number 2']                     #lines.line.0.dirn.pattern
                    sheet['ci'+str(fila)]=linept                                        #lines.line.0.dirn.routePartitionName
                    sheet['cj'+str(fila)]="Use System Policy"                           #lines.line.0.mwlPolicy
                    sheet['ck'+str(fila)]=""                                            #lines.line.0.ringSettingIdlePickupAlert
                    sheet['cl'+str(fila)]=row['Busy Trigger 2']                         #lines.line.0.busyTrigger
                    sheet['cm'+str(fila)]="Default"                                     # lines.line.0.audibleMwi
                    sheet['cn'+str(fila)]=row['Display 1']                              # lines.line.0.display
                    ##
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #2
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                ##L",dn,"##: ",row['Directory Number 2'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 2'],file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                    # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                         # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                            # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 2']                   # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 2']                #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                      #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                     #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                       #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 2']                #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                      #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                    #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                      #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                       #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                      # LINE CSS
                    if row['External Phone Number Mask 2'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 2']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 2']                 # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 2']                    # pattern
                    sheet['aa'+str(filadn)]="Default"                                   # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                          # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 2']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                     # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                    # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 2']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                     # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                   # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 2']               # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                     # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                    # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                    # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 2']                      #alertingName
                    sheet['ap'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                     #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                     #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                          #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                     #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 2']                   #description
                    sheet['av'+str(filadn)]="false"                                     #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                     #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                      #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                  #releaseClause
                    sheet['be'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                     #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                      #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                          # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                     #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                   # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 2']            #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                     #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                   # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                      # active
                    sheet['bp'+str(filadn)]=""                                          # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                     # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 2']  # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                     # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                   # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 2']  #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                     # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                   # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                     #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                 #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 2']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                     # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                   #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                          #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 2']               # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                     # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                    # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                   #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #3
            ##
            if data['e164'][0]['emdn'] >= 3:
                if row['Directory Number 3']!="":
                    dn=dn+1
                    ## DEBUG
                    print("EM#",fila,"                ##L",dn,"##: ",row['Directory Number 3'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 3'],file=f)
                    ##
                    sheet =  blk["EM"]
                    ##
                    sheet['co'+str(fila)]=row['ASCII Display 3']                        #lines.line.0.displayAscii
                    sheet['cp'+str(fila)]=row['User ID 1']                              #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['cq'+str(fila)]="Ring"                                        #lines.line.0.ringSetting
                    sheet['cr'+str(fila)]="Use System Default"                          #lines.line.0.consecutiveRingSetting
                    #sheet['cs'+str(fila)]="Default"                                    # lines.line.0.recordingProfileName
                    sheet['ct'+str(fila)]=dn                                            #lines.line.0.index
                    #sheet['cu'+str(fila)]="Use System Default"                         # lines.line.0.ringSettingActivePickupAlert
                    sheet['cv'+str(fila)]=row['Line Text Label 3']                      #lines.line.0.label
                    sheet['cw'+str(fila)]="Gateway Preferred"                           #lines.line.0.recordingMediaSource
                    sheet['cx'+str(fila)]=row['Maximum Number of Calls 3']              #lines.line.0.maxNumCalls
                    sheet['cy'+str(fila)]="General"                                     #lines.line.0.partitionUsage
                    sheet['cz'+str(fila)]="Call Recording Disabled"                     #lines.line.0.recordingFlag
                    sheet['da'+str(fila)]=""                                            #lines.line.0.speedDial
                    sheet['db'+str(fila)]=""                                            #lines.line.0.monitoringCssName
                    if row['External Phone Number Mask 3'] != "":
                        sheet['dc'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 3']  #lines.line.0.e164Mask
                    sheet['dd'+str(fila)]="true"                                        #lines.line.0.missedCallLogging
                    sheet['de'+str(fila)]="true"                                        #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['df'+str(fila)]="false"                                       #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['dg'+str(fila)]="false"                                       # lines.line.0.callInfoDisplay.callerName
                    sheet['dh'+str(fila)]="true"                                        # lines.line.0.callInfoDisplay.callerNumber
                    sheet['di'+str(fila)]=row['Directory Number 3']                     #lines.line.0.dirn.pattern
                    sheet['dj'+str(fila)]=linept                                        #lines.line.0.dirn.routePartitionName
                    sheet['dk'+str(fila)]="Use System Policy"                           #lines.line.0.mwlPolicy
                    sheet['dl'+str(fila)]=""                                            #lines.line.0.ringSettingIdlePickupAlert
                    sheet['dm'+str(fila)]=row['Busy Trigger 3']                         #lines.line.0.busyTrigger
                    sheet['dn'+str(fila)]="Default"                                     # lines.line.0.audibleMwi
                    sheet['do'+str(fila)]=row['Display 3']                              # lines.line.0.display
                    ##
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #3
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                ##L",dn,"##: ",row['Directory Number 3'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 3'],file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3]                              # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                    # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                         # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                            # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 3']                   # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 3']                #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                      #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                     #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                       #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 3']  #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                      #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                    #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                      #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                       #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                      # LINE CSS
                    if row['External Phone Number Mask 3'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 3']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 3']                 # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 3']                    # pattern
                    sheet['aa'+str(filadn)]="Default"                                   # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                          # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 3']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                     # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                    # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 3']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                     # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                   # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 3']               # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                     # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                    # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                    # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 3']                      #alertingName
                    sheet['ap'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                     #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                     #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                          #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                     #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 3']                   #description
                    sheet['av'+str(filadn)]="false"                                     #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                     #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                      #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                  #releaseClause
                    sheet['be'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                     #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                      #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                          # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                     #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                   # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 3']            #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                     #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                   # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                      # active
                    sheet['bp'+str(filadn)]=""                                          # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                     # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 3']  # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                     # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                   # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 3']  #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                     # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                   # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                     #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                 #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 3']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                     # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                   #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                          #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 3']               # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                     # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                    # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                   #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #4
            ##
            if data['e164'][0]['emdn'] >= 4:
                if row['Directory Number 4']!="":
                    dn=dn+1
                    ## DEBUG
                    print("EM#",fila,"                ##L",dn,"##: ",row['Directory Number 4'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 4'],file=f)
                    ##
                    sheet =  blk["EM"]
                    ##
                    sheet['dp'+str(fila)]=row['ASCII Display 4']                        #lines.line.0.displayAscii
                    sheet['dq'+str(fila)]=row['User ID 1']                              #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['dr'+str(fila)]="Ring"                                        #lines.line.0.ringSetting
                    sheet['ds'+str(fila)]="Use System Default"                          #lines.line.0.consecutiveRingSetting
                    #sheet['dt'+str(fila)]="Default"                                    # lines.line.0.recordingProfileName
                    sheet['du'+str(fila)]=dn                                            #lines.line.0.index
                    #sheet['dv'+str(fila)]="Use System Default"                         # lines.line.0.ringSettingActivePickupAlert
                    sheet['dw'+str(fila)]=row['Line Text Label 4']                      #lines.line.0.label
                    sheet['dx'+str(fila)]="Gateway Preferred"                           #lines.line.0.recordingMediaSource
                    sheet['dy'+str(fila)]=row['Maximum Number of Calls 4']              #lines.line.0.maxNumCalls
                    sheet['dz'+str(fila)]="General"                                     #lines.line.0.partitionUsage
                    sheet['ea'+str(fila)]="Call Recording Disabled"                     #lines.line.0.recordingFlag
                    sheet['eb'+str(fila)]=""                                            #lines.line.0.speedDial
                    sheet['ec'+str(fila)]=""                                            #lines.line.0.monitoringCssName
                    if row['External Phone Number Mask 4'] != "":
                        sheet['ed'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 4']  #lines.line.0.e164Mask
                    sheet['ee'+str(fila)]="true"                                        #lines.line.0.missedCallLogging
                    sheet['ef'+str(fila)]="true"                                        #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['eg'+str(fila)]="false"                                       #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['eh'+str(fila)]="false"                                       # lines.line.0.callInfoDisplay.callerName
                    sheet['ei'+str(fila)]="true"                                        # lines.line.0.callInfoDisplay.callerNumber
                    sheet['ej'+str(fila)]=row['Directory Number 4']                     #lines.line.0.dirn.pattern
                    sheet['ek'+str(fila)]=linept                                        #lines.line.0.dirn.routePartitionName
                    sheet['el'+str(fila)]="Use System Policy"                           #lines.line.0.mwlPolicy
                    sheet['em'+str(fila)]=""                                            #lines.line.0.ringSettingIdlePickupAlert
                    sheet['en'+str(fila)]=row['Busy Trigger 4']                         #lines.line.0.busyTrigger
                    sheet['eo'+str(fila)]="Default"                                     # lines.line.0.audibleMwi
                    sheet['ep'+str(fila)]=row['Display 4']                              # lines.line.0.display
                    ##
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #4
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                ##L",dn,"##: ",row['Directory Number 4'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 4'],file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3]                              # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                    # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                         # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                            # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 4']                   # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 4']                #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                      #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                     #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                       #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 4']  #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                      #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                    #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                      #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                       #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                      # LINE CSS
                    if row['External Phone Number Mask 4'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 4']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 4']                 # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 4']                    # pattern
                    sheet['aa'+str(filadn)]="Default"                                   # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                          # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 4']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                     # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                    # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 4']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                     # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                   # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 4']               # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                     # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                    # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                    # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 4']                      #alertingName
                    sheet['ap'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                     #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                     #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                          #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                     #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 4']                   #description
                    sheet['av'+str(filadn)]="false"                                     #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                     #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                      #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                  #releaseClause
                    sheet['be'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                     #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                      #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                          # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                     #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                   # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 4']               #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                     #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                   # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                      # active
                    sheet['bp'+str(filadn)]=""                                          # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                     # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 4']               # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                     # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                   # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 4']               #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                     # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                   # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                     #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                 #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 4']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                     # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                   #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                          #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 4']               # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                     # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                    # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                   #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #5
            ##
            if data['e164'][0]['emdn'] >= 5:
                if row['Directory Number 5']!="":
                    dn=dn+1
                    delta=84*(dn-1)
                    ## DEBUG
                    print("EM#",fila,"                ##L",dn,"##: ",row['Directory Number 5'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 5'],file=f)
                    ##
                    sheet =  blk["EM"]
                    ##
                    sheet['eq'+str(fila)]=row['ASCII Display 5']                        #lines.line.0.displayAscii
                    sheet['er'+str(fila)]=row['User ID 1']                              #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['es'+str(fila)]="Ring"                                        #lines.line.0.ringSetting
                    sheet['et'+str(fila)]="Use System Default"                          #lines.line.0.consecutiveRingSetting
                    #sheet['eu'+str(fila)]="Default"                                    # lines.line.0.recordingProfileName
                    sheet['ev'+str(fila)]=dn                                            #lines.line.0.index
                    #sheet['ew'+str(fila)]="Use System Default"                         # lines.line.0.ringSettingActivePickupAlert
                    sheet['ex'+str(fila)]=row['Line Text Label 5']                      #lines.line.0.label
                    sheet['ey'+str(fila)]="Gateway Preferred"                           #lines.line.0.recordingMediaSource
                    sheet['ez'+str(fila)]=row['Maximum Number of Calls 5']              #lines.line.0.maxNumCalls
                    sheet['fa'+str(fila)]="General"                                     #lines.line.0.partitionUsage
                    sheet['fb'+str(fila)]="Call Recording Disabled"                     #lines.line.0.recordingFlag
                    sheet['fc'+str(fila)]=""                                            #lines.line.0.speedDial
                    sheet['fd'+str(fila)]=""                                            #lines.line.0.monitoringCssName
                    if row['External Phone Number Mask 5'] != "":
                        sheet['fe'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 5']  #lines.line.0.e164Mask
                    sheet['ff'+str(fila)]="true"                                        #lines.line.0.missedCallLogging
                    sheet['fg'+str(fila)]="true"                                        #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['fh'+str(fila)]="false"                                       #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['fi'+str(fila)]="false"                                       # lines.line.0.callInfoDisplay.callerName
                    sheet['fj'+str(fila)]="true"                                        # lines.line.0.callInfoDisplay.callerNumber
                    sheet['fk'+str(fila)]=row['Directory Number 5']                     #lines.line.0.dirn.pattern
                    sheet['fl'+str(fila)]=linept                                        #lines.line.0.dirn.routePartitionName
                    sheet['fm'+str(fila)]="Use System Policy"                           #lines.line.0.mwlPolicy
                    sheet['fn'+str(fila)]=""                                            #lines.line.0.ringSettingIdlePickupAlert
                    sheet['fo'+str(fila)]=row['Busy Trigger 5']                         #lines.line.0.busyTrigger
                    sheet['fp'+str(fila)]="Default"                                     # lines.line.0.audibleMwi
                    sheet['fq'+str(fila)]=row['Display 5']                              # lines.line.0.display
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #5
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                ##L",dn,"##: ",row['Directory Number 5'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 5'],file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3]                              # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                    # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                         # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                            # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 5']                   # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 5']                #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                      #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                     #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                       #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 5']                #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                      #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                    #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                      #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                       #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                      # LINE CSS
                    if row['External Phone Number Mask 5'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 5']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 5']                 # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 5']                    # pattern
                    sheet['aa'+str(filadn)]="Default"                                   # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                          # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 5']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                     # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                    # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 5']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                     # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                   # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 5']               # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                     # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                    # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                    # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 5']                      #alertingName
                    sheet['ap'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                     #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                     #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                          #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                     #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 5']                   #description
                    sheet['av'+str(filadn)]="false"                                     #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                     #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                      #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                  #releaseClause
                    sheet['be'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                     #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                      #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                          # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                     #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                   # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 5']               #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                     #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                   # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                      # active
                    sheet['bp'+str(filadn)]=""                                          # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                     # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 5']               # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                     # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                   # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 5']               #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                     # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                   # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                     #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                 #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 5']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                     # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                   #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                          #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 5']               # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                     # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                    # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                   #presenceGroupName
                    filadn=filadn+1

            ##
            ## LINE #6
            ##
            if data['e164'][0]['emdn'] >= 6:
                if row['Directory Number 6']!="":
                    dn=dn+1
                    ## DEBUG
                    print("EM#",fila,"                ##L",dn,"##: ",row['Directory Number 6'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 6'],file=f)
                    ##
                    sheet =  blk["EM"]
                    ##
                    sheet['fr'+str(fila)]=row['ASCII Display 6']                        #lines.line.0.displayAscii
                    sheet['fs'+str(fila)]=row['User ID 1']                              #lines.line.0.associatedEndusers.enduser.0.userId
                    sheet['ft'+str(fila)]="Ring"                                        #lines.line.0.ringSetting
                    sheet['fu'+str(fila)]="Use System Default"                          #lines.line.0.consecutiveRingSetting
                    #sheet['fv'+str(fila)]="Default"                                    # lines.line.0.recordingProfileName
                    sheet['fw'+str(fila)]=dn                                            #lines.line.0.index
                    #sheet['fx'+str(fila)]="Use System Default"                         # lines.line.0.ringSettingActivePickupAlert
                    sheet['fy'+str(fila)]=row['Line Text Label 6']                      #lines.line.0.label
                    sheet['fz'+str(fila)]="Gateway Preferred"                           #lines.line.0.recordingMediaSource
                    sheet['ga'+str(fila)]=row['Maximum Number of Calls 6']              #lines.line.0.maxNumCalls
                    sheet['gb'+str(fila)]="General"                                     #lines.line.0.partitionUsage
                    sheet['gc'+str(fila)]="Call Recording Disabled"                     #lines.line.0.recordingFlag
                    sheet['gd'+str(fila)]=""                                            #lines.line.0.speedDial
                    sheet['ge'+str(fila)]=""                                            #lines.line.0.monitoringCssName
                    if row['External Phone Number Mask 6'] != "":
                        sheet['gf'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 6']  #lines.line.0.e164Mask
                    sheet['gg'+str(fila)]="true"                                        #lines.line.0.missedCallLogging
                    sheet['gh'+str(fila)]="true"                                        #lines.line.0.callInfoDisplay.dialedNumber
                    sheet['gi'+str(fila)]="false"                                       #lines.line.0.callInfoDisplay.redirectedNumber
                    sheet['gk'+str(fila)]="false"                                       # lines.line.0.callInfoDisplay.callerName
                    sheet['gj'+str(fila)]="true"                                        # lines.line.0.callInfoDisplay.callerNumber
                    sheet['gl'+str(fila)]=row['Directory Number 6']                     #lines.line.0.dirn.pattern
                    sheet['gm'+str(fila)]=linept                                        #lines.line.0.dirn.routePartitionName
                    sheet['gn'+str(fila)]="Use System Policy"                           #lines.line.0.mwlPolicy
                    sheet['go'+str(fila)]=""                                            #lines.line.0.ringSettingIdlePickupAlert
                    sheet['gp'+str(fila)]=row['Busy Trigger 6']                         #lines.line.0.busyTrigger
                    sheet['gq'+str(fila)]="Default"                                     # lines.line.0.audibleMwi
                    sheet['gr'+str(fila)]=row['Display 6']                              # lines.line.0.display
                    ######################################################
                    ######################################################
                    ##
                    ## LINE: LINE #6
                    ##
                    ######################################################
                    ## DEBUG
                    print("LN#",filadn,"                ##L",dn,"##: ",row['Directory Number 6'],"FMO-EPNM:",data['e164'][0]['cc'],data['e164'][0]['ac'],row['External Phone Number Mask 6'],file=f)

                    sheet = blk["LINE"]
                    sheet['B'+str(filadn)]=hierarchynode+"."+fmositename
                    sheet['C'+str(filadn)]=action
                    #sheet['D'+str(filadn)]="name:"+row[3] # Search field
                    ######################################################
                    sheet['i'+str(filadn)]="Default"                                    # partyEntranceTone
                    sheet['j'+str(filadn)]="Use System Default"                         # cfaCssPolicy
                    sheet['k'+str(filadn)]="Auto Answer Off"                            # autoAnswer
                    sheet['m'+str(filadn)]=row['Call Pickup Group 6']                # CPG
                    sheet['n'+str(filadn)]=row['Forward Unregistered Internal Destination 6']                #callForwardNotRegisteredInt.destination
                    sheet['o'+str(filadn)]="false"                                      #callForwardNotRegisteredInt.forwardToVoiceMail
                    sheet['p'+str(filadn)]=fwdnocss                                     #callForwardNotRegisteredInt.callingSearchSpaceName
                    sheet['q'+str(filadn)]=linept                                       #routePartitionName
                    sheet['r'+str(filadn)]=row['Forward on CTI Failure Destination 6']                #callForwardOnFailure.destination
                    sheet['s'+str(filadn)]="false"                                      #callForwardOnFailure.forwardToVoiceMail
                    sheet['t'+str(filadn)]=fwdallcss                                    #callForwardOnFailure.callingSearchSpaceName
                    sheet['u'+str(filadn)]="false"                                      #rejectAnonymousCall
                    sheet['v'+str(filadn)]="true"                                       #aarKeepCallHistory
                    sheet['w'+str(filadn)]=linecss                                      # LINE CSS
                    if row['External Phone Number Mask 1'] != "":
                        sheet['x'+str(fila)]=data['e164'][0]['cc']+data['e164'][0]['ac']+row['External Phone Number Mask 6']  # aarDestinationMask
                    sheet['y'+str(filadn)]=row['ASCII Alerting Name 6']                # asciiAlertingName
                    sheet['z'+str(filadn)]=row['Directory Number 6']                # pattern
                    sheet['aa'+str(filadn)]="Default"                                   # patternPrecedence
                    sheet['ab'+str(filadn)]=""                                          # callForwardNoAnswer.duration
                    sheet['ac'+str(filadn)]=row['Forward No Answer External Destination 6']               # callForwardNoAnswer.destination
                    sheet['ad'+str(filadn)]="false"                                     # callForwardNoAnswer.forwardToVoiceMail
                    sheet['ae'+str(filadn)]=fwdnocss                                    # callForwardNoAnswer.callingSearchSpaceName
                    sheet['ag'+str(filadn)]=row['Forward No Coverage External Destination 6']               # callForwardNoCoverage.destination
                    sheet['ah'+str(filadn)]="false"                                     # callForwardNoCoverage.forwardToVoiceMail
                    sheet['ai'+str(filadn)]=fwdallcss                                   # callForwardNoCoverage.callingSearchSpaceName
                    sheet['aj'+str(filadn)]=row['Forward Unregistered External Destination 6']               # callForwardNotRegistered.destination
                    sheet['ak'+str(filadn)]="false"                                     # callForwardNotRegistered.forwardToVoiceMail
                    sheet['al'+str(filadn)]=fwdnocss                                    # callForwardNotRegistered.callingSearchSpaceName
                    sheet['am'+str(filadn)]="Device"                                    # usage
                    sheet['ao'+str(filadn)]=row['Alerting Name 6']                      #alertingName
                    sheet['ap'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['aq'+str(filadn)]="false"                                     #enterpriseAltNum.addLocalRoutePartition
                    sheet['ar'+str(filadn)]="false"                                     #enterpriseAltNum.advertiseGloballyIls
                    sheet['as'+str(filadn)]=""                                          #enterpriseAltNum.routePartition
                    sheet['at'+str(filadn)]="false"                                     #enterpriseAltNum.isUrgent
                    sheet['au'+str(filadn)]=row['Line Description 6']                   #description
                    sheet['av'+str(filadn)]="false"                                     #aarVoiceMailEnabled
                    sheet['aw'+str(filadn)]="false"                                     #useE164AltNum
                    sheet['ba'+str(filadn)]="true"                                      #allowCtiControlFlag
                    sheet['bd'+str(filadn)]="No Error"                                  #releaseClause
                    sheet['be'+str(filadn)]=""                                          #enterpriseAltNum.numMask
                    sheet['bf'+str(filadn)]="false"                                     #e164AltNum.addLocalRoutePartition
                    sheet['bg'+str(filadn)]="true"                                      #e164AltNum.advertiseGloballyIls
                    sheet['bh'+str(filadn)]=""                                          # e164AltNum.routePartition
                    sheet['bi'+str(filadn)]="false"                                     #e164AltNum.isUrgent
                    sheet['bj'+str(filadn)]=devicecss                                   # callForwardAll.secondaryCallingSearchSpaceName
                    sheet['bk'+str(filadn)]=row['Forward All Destination 6']            #callForwardAll.destination
                    sheet['bl'+str(filadn)]="false"                                     #callForwardAll.forwardToVoiceMail
                    sheet['bm'+str(filadn)]=fwdallcss                                   # callForwardAll.callingSearchSpaceName
                    sheet['bn'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveVmEnabled
                    sheet['bo'+str(filadn)]="true"                                      # active
                    sheet['bp'+str(filadn)]=""                                          # VoiceMailProfileName
                    sheet['bq'+str(filadn)]="false"                                     # useEnterpriseAltNum
                    sheet['bt'+str(filadn)]=row['Forward Busy Internal Destination 6']               # callForwardBusyInt.destination
                    sheet['bu'+str(filadn)]="false"                                     # callForwardBusyInt.forwardToVoiceMail
                    sheet['bv'+str(filadn)]=fwdallcss                                   # callForwardBusyInt.callingSearchSpaceName
                    sheet['bw'+str(filadn)]=row['Forward Busy External Destination 6']               #  callForwardBusy.destination
                    sheet['bx'+str(filadn)]="false"                                     # callForwardBusy.forwardToVoiceMail
                    sheet['by'+str(filadn)]=fwdallcss                                   # callForwardBusy.callingSearchSpaceName
                    sheet['ca'+str(filadn)]="false"                                     #patternUrgency
                    sheet['cb'+str(filadn)]=fmoenvconfig['fmoaargroup']                 #aarNeighborhoodName
                    sheet['cc'+str(filadn)]="false"                                     # parkMonForwardNoRetrieveIntVmEnabled
                    sheet['cd'+str(filadn)]=row['Forward No Coverage Internal Destination 6']               # callForwardNoCoverageInt.destination
                    sheet['ce'+str(filadn)]="false"                                     # callForwardNoCoverageInt.forwardToVoiceMail
                    sheet['cf'+str(filadn)]=fwdallcss                                   #callForwardNoCoverageInt.callingSearchSpaceName
                    sheet['cg'+str(filadn)]=""                                          #callForwardAlternateParty.destination
                    sheet['cj'+str(filadn)]=row['Forward No Answer Internal Destination 6']               # callForwardNoAnswerInt.destination
                    sheet['ck'+str(filadn)]="false"                                     # callForwardNoAnswerInt.forwardToVoiceMail
                    sheet['cl'+str(filadn)]=fwdnocss                                    # callForwardNoAnswerInt.callingSearchSpaceName
                    sheet['cn'+str(filadn)]="Standard Presence group"                   #presenceGroupName
                    filadn=filadn+1
            fila=fila+1


## CMO File INPUT DATA: Close
fgw.close()
## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)
## LOG de CONFIGURACION
f.close()

exit(0)
