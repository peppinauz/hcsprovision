#!/bin/bash

## CONFIG ENTORNO
## STATIC DATA
entorno="doc/fmo-static-data.cfg"   ## Produccion
#entorno="doc/agencias-cl2.cfg"   ## agencias-cl2
#entorno="doc/agencias-clx.cfg"   ## agencias-clx
## CONFIG SO
#linux=0 # Desarrollo
linux=1 # Produccion

## Chequeamos en todos los clusters
base=".."
## Clusters
cl[1]="$base/CMO/Cluster01"
cl[2]="$base/CMO/Cluster02"
cl[3]="$base/CMO/Cluster03"
cl[4]="$base/CMO/Cluster04"
cl[5]="$base/CMO/Cluster05"
cl[6]="$base/CMO/Cluster06"
cl[7]="$base/CMO/Cluster07"
cl[8]="$base/CMO/Cluster08"
cl[9]="$base/CMO/Cluster09"
cl[10]="$base/CMO/Cluster10"

## Tools
if [ $linux == 1 ]
then
  ## PRODUCCION
  AWK=/usr/bin/awk
  GREP=/bin/grep
  WC=/usr/bin/wc
  PYTHON=/usr/bin/python3
  CSPLIT=/usr/bin/csplit
  RM=/bin/rm
  CAT=/bin/cat
  SED=/usr/bin/sed
else
  ## MIO
  AWK=/usr/bin/awk
  GREP=/usr/bin/grep
  WC=/usr/bin/wc
  PYTHON=/usr/local/bin/python3
  CSPLIT=/usr/bin/csplit
  RM=/bin/rm
  CAT=/bin/cat
  SED=/usr/bin/sed
fi

## Variables
slcencontrado=0
match=0

## FUNCIONES: Check SLC
function FN_CHECKSLC(){
local inputfile="$2"
local slc="$1"
local dpool
local slcp="$slc-DP"

if [ -f $inputfile ]
then
#echo "Buscando en:" $inputfile
	echo "  INFO:   Buscando   $slc   en   " $inputfile
else
	echo "  WARN:   El fichero $inputfile NO existe!!"
        exit 1
fi

dpool="$( $GREP $slcp $inputfile | $WC -l)"

## Controles adicionales
##
return $dpool
}

if [ "$#" -eq 1 ]
then
  ## Variables entrada
  slc=$1
  #areacode=$2
  ## BUSCA SLC en devicepool.csv
  ##
  for i in "${cl[@]}"
  do
    #echo "$i"
    devicepool="$i/devicepool.csv"
    #echo "$devicepool"
    FN_CHECKSLC $slc $devicepool
    slcencontrado=$?

    #echo "XXX $slcencontrado"
    if [ "$slcencontrado" -ne "0" ]
    then
      if [ "$match" -eq "1" ]
      then
        echo ""
        echo "ERROR: SLC duplicado"
        echo ""
        echo "ERROR: Encontrado en     $workdirectory"
        echo "ERROR: Encontrado en     $i"
        echo ""
        echo "ERROR: No puedo determinar el cluster origen"
        echo ""
        echo "ATENCION: No se va a generar ningun BLK"
        echo "ATENCION: Me paro hasta saber que hacer"
        echo ""
        exit 1
      fi
      workdirectory=$i
      match=1
      #echo "INFO:   SLC encontrado:    $workdirectory"
      #echo "@@@ Working directory: $workdirectory"
    fi
  done
else
  if [ "$#" -eq 3 ]
  then
    if [ $2 != "-forcecluster" ]
    then
      echo "  ERROR: Comando no reconocido: $3"
      echo "  El único comando soportado: -forcecluster"
      exit 1
    else
      forcedcl="$base/CMO/$3"
      if [ ! -d "$forcedcl" ]
      then
        echo "   ERROR: El directorio no existe: $forcedcl"
        echo "   Estos son los Cluster disponibles"
        for i in "${cl[@]}"
        do
          echo "  $i"
        done
        exit 1
      else
        # Si el directorio existe miramos si existe el SLC en el Cluster
          slc=$1
          #areacode=$2
          devicepool="$forcedcl/devicepool.csv"
          FN_CHECKSLC $slc $devicepool
          slcencontrado=$?
          if [ "$slcencontrado" -ne "0" ]
          then
            match=1
            workdirectory="$base/CMO/$3"
          fi
      fi
    fi
  else
    echo "  Si NO sabes el cluster tienes que invocar este comando:"
	  echo "  $0 SLC Area-Code" >&2
    echo "  --> SLC: XXXX, 4 dígitos que son el SLC del site a migrar " >&2
    echo "  --> Area-Code: YY, 2 dígitos que son el area code del site a migrar"
    echo "  -----------------------------------------------------------------"
    echo "  Si ya sabes el Cluster y no quieres ejecutar la búsqueda:"
	  echo "  $0 SLC Area-Code -forcecluster <Cluster-path>"
    echo "  --> SLC: XXXX, 4 dígitos que son el SLC del site a migrar " >&2
    echo "  --> Area-Code: YY, 2 dígitos que son el area code del site a migrar"
    echo "  --> -forcecluster: switch obligatorio"
    echo "  --> Cluster-path: directorio del cluster, ejemplo: Cluster01, Cluster02,...,Cluster10"
    exit 1
  fi
fi

### Sino se encuentra SLC
##
if [ "$match" -eq "0" ]
  then
    echo ""
    echo "ERROR:   SLC NO encontrado"
    echo ""
    echo "ATENCION: No se va a generar ningun BLK"
    echo "ATENCION: Me paro hasta saber que hacer"
    echo ""
    exit 1
  else
    echo ""
    echo "INFO:   SLC encontrado:    $workdirectory"
    echo ""
fi

#############################################################################################
#############################################################################################
#############################################################################################
##
## Directorio de trabajo
##
slcdir="$base/FMO/$slc"
now=$(date +"%Y-%m-%d-%H-%M")
slclogfile="$slcdir/output-$now.log"            ## Fichero LOG de configuracion
cmosite="$slcdir/cmo-site-config.json"          ## Datos de configuracion CMO
configlogfile="$slcdir/configlog-$now.txt"      ## LOG de CONFIGURACION

if [ ! -d $slcdir ]; then
  mkdir $slcdir
fi
touch $slclogfile
touch $cmosite

#############################################################################################
#############################################################################################
#############################################################################################
## CORRECCIONES
##
## LINEGROUP.csv, uno de los campos tiene comas lo cual hace que se cuenten los campos de forma incorrecta
lgnameFAKE=$workdirectory/linegroup.csv
lgnamemod1=$workdirectory/linegroup.mod1.csv
lgname=$workdirectory/linegroup.mod2.csv

if [ ! -f $lgname ] ## Si el fichero no existe
then
  ## Si el fichero NO existe lo creamos
  echo "WARN:  $lgname NO EXISTE ... lo creamos"
  sed "s/member;/member@/g" $lgnameFAKE > $lgnamemod1
  sed "s/then,/then@/g" $lgnamemod1 > $lgname
fi

##
## PHONE.csv, uno de los dispositivos tiene "," en el apartado XML

ipphonenameFAKE=$workdirectory/phone.csv
ipphonename=$workdirectory/phone.mod1.csv

if [ ! -f $ipphonename ] ## Si el fichero no existe
then
  ## Si el fichero NO existe lo creamos
  echo "WARN:  $ipphonename NO EXISTE ... lo creamos"
  sed "s/daysBacklightNotActive>1,/daysBacklightNotActive>1 /g" $ipphonenameFAKE > $ipphonename
fi

##
## GATEWAY.csv, el fichero GATEWAY aglutina 4 tipos de configuracion diferente
## - Entity:GATEWAY, definición de los gateways
## - Entity:SLOTCONFIGURATION, definición de la tarjeteria asociada al gateway
## - Entity:ANALOG, definición de los puertos analógicos asociados al gateway
## - Entity:H323, definición de los gateways h323
##
## Vamos a dividir el fichero gateway en 4 ficheros
gatewaygw=$workdirectory/gateway.gw.csv
gatewayslot=$workdirectory/gateway.slot.csv
gatewayanalog=$workdirectory/gateway.analog.csv
gatewayh323=$workdirectory/gateway.h323.csv
# ORIGINAL
gateway=$workdirectory/gateway.csv

if [ $linux == 1 ]
then
  ## PRODUCCION
  if [[ ! -f $gatewaygw || ! -f $gatewayslot || ! -f $gatewayanalog || ! -f $gatewayh323 ]] ## Si el fichero no existe
  then
    echo "(WW) Generando ficheros de gateway"
    $CSPLIT -s -k $gateway '/Entity:/' '{3}'
    ## Este comando divide el fichero GATEWAY en 4 partes
    $CAT xx01 | $GREP -v "Entity:" > $gatewaygw
    $CAT xx02 | $GREP -v "Entity:" > $gatewayslot
    $CAT xx03 | $GREP -v "Entity:" > $gatewayanalog
    $CAT xx04 | $GREP -v "Entity:" > $gatewayh323
    $RM xx00 xx01 xx02 xx03 xx04
  fi
else
    echo "(WW) Generando ficheros de gateway"
  ## MIO
  if [[ ! -f $gatewaygw || ! -f $gatewayslot || ! -f $gatewayanalog || ! -f $gatewayh323 ]] ## Si el fichero no existe
  then
    $CSPLIT -s -k $gateway '/Entity:/' '{2}'
    ## Este comando divide el fichero GATEWAY en 4 partes
    $CAT xx00 | $GREP -v "Entity:" | $SED 's/NULL//g' > $gatewaygw
    $CAT xx01 | $GREP -v "Entity:" | $SED 's/NULL//g' > $gatewayslot
    $CAT xx02 | $GREP -v "Entity:" | $SED 's/NULL//g' > $gatewayanalog
    $CAT xx03 | $GREP -v "Entity:" | $SED 's/NULL//g' > $gatewayh323
    $RM xx00 xx01 xx02 xx03
  fi
fi

#############################################################################################
#############################################################################################
#############################################################################################
##
## Empezamos con la generación de config:
## STEP 00: Generamos el blk para creación del SITE: --> 00.site.blk
echo "(II) Filtrando info del SITE:: $slc"
$PYTHON filtra.dp.01.py $cmosite $workdirectory $slc $configlogfile
echo "(II) Generando primer BLK del SITE"
$PYTHON config.fmosite.py $entorno $cmosite $slc $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## SITE
##
## STEP 01: Generamos el blk con el resto de elementos del SITE: --> 01.site.blk
echo "(II) Generando Site Info"
$PYTHON rellena.site.03.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## CPG
## Ficheros
## Salida --------------> 03.cpg.XXXX.xlsx
echo "(II) Generando Call Pickup Groups"
$PYTHON rellena.cpg.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## GATEWAYS + Lineas analogicas
## Ficheros
## Salida ---------------> 04.gw.XXXX.xlsx
nogw="$( $GREP "0.0.0.0" $cmosite | $WC -l )"

if [ $nogw == "0" ]
then
  echo "(II) Generando Gateways, lineas analogicas SCCP"
  $PYTHON rellena.gw.py $entorno $cmosite $slc $workdirectory $configlogfile
else
  echo "(WW) No hay Gateway asociado al SITE se trata de un SITE con Gateway compartido"
  echo "(WW) NO se va a generar BLK para Gateway"
fi

#############################################################################################
#############################################################################################
#############################################################################################
##
## IP-PHONES
## Ficheros
echo "(II) Generando IP phones y ATA"
$PYTHON rellena.new_phone.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## DEVICE PROFILE
## Ficheros
echo "(II) Generando Device Profile"
$PYTHON rellena.new_devprof.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## Hunt-List / Hunt-group / Line-group
## Ficheros
echo "(II) Generando Line Group, Hunt Group, Hunt Pilot"
$PYTHON rellena.hunt.py $entorno $cmosite $slc $workdirectory $configlogfile
##
echo "(II) Generando NEW Line Group, Hunt Group, Hunt Pilot"
$PYTHON rellena.new_hunt.py $entorno $cmosite $slc $workdirectory $configlogfile
echo "(II) Generando NEW Line Group, Hunt Group, Hunt Pilot..Routing"
$PYTHON rellena.new_hunt_routing.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## Site deactivation
## Actions: TP/Pre-ISR -> TP/Null
echo "(II) Generando BLK para desactivar el SITE"
$PYTHON rellena.off.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## Site activation
## Actions:
##   TP/Null -> TP/Pre-ISR
##   DN+HP Advertised pattern
echo "(II) Generando BLK para activar el SITE"
$PYTHON rellena.on.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## SIPP TEST
## Ficheros
echo "(II) Generando Dispositivo de Pruebas"
$PYTHON rellena.siptest.py $entorno $cmosite $slc $workdirectory $configlogfile

#############################################################################################
#############################################################################################
#############################################################################################
##
## MIGRACIONES
## Ficheros
echo "(II) Generando informacion para la migración::PHONES"
$PYTHON migracion.phone.py $entorno $cmosite $slc $workdirectory $configlogfile
## Envio de ficheros a Campinas:
echo "(II) Copiando ficheros para la migración::PHONES"
echo ".....$slcdir/*reset*.csv"

$SCP $slcdir/*reset*.csv delivery@vhcsubsaotb2:/home/delivery/source/inputfiles/

#############################################################################################
#############################################################################################
#############################################################################################
##
## MIGRACIONES
## Ficheros
echo "(II) Generando informacion para la migración::GATEWAY"
$PYTHON migracion.gateway.py $entorno $cmosite $slc $workdirectory $configlogfile
## Envio de ficheros a Campinas:
echo "(II) Copiando ficheros para la migración::GATEWAY"
echo ".....$base/FMO/dataInfo.csv"

#$SCP $base/FMO/dataInfo.csv gw@vhcsubsaotb2:/home/gw/inventario/


exit 0

#############################################################################################
#############################################################################################
#############################################################################################
##
## REPORT

#############################################################################################
#############################################################################################
#############################################################################################
## SITE
##
devicepoolname=$workdirectory/devicepool.csv
regionname=$workdirectory/region.csv
locationname=$workdirectory/location.csv
mrgname=$workdirectory/mediaresourcegroup.csv
mrglname=$workdirectory/mediaresourcegrouplist.csv
srstname=$workdirectory/srst.csv

## DEVICE POOL: Verificamos errores de config en CMO
##
devicepoolid="$slc-DP"
ndevicepoolid=0
regionid="$slc-RGN"
nregionid=0
srstid="$slc-SRST"
nsrstid=0
locationid="$slc-LOC"
nlocationid=0
mrglid="$slc-MRGL"
nmrglid=0
mrgid="$slc-MRG"
nmrgid=0

## DP
ndevicepoolid="$( $GREP $devicepoolid $devicepoolname | $AWK -F "," '{printf("%s \n",$1)}' | $WC -l )"
cmodevicepoolid=( $( $GREP $devicepoolid $devicepoolname | $AWK -F "," '{printf("%s \n",$1)}' ) )

echo "(II):   $devicepoolid >>>> $ndevicepoolid  [ ${cmodevicepoolid[@]} ]" >> $slclogfile

## REGION
nregionid="$( $GREP $devicepoolid $devicepoolname | $WC -l)"
cmoregionid=( $( $GREP $devicepoolid $devicepoolname | $AWK -F "," '{print $4}') )

if [ "$regionid" != "${cmoregionid[0]}" ]
then
  echo "(EE):   El nombre de la región difiere, el nombre de la región debería ser: $regionid  y hemos encontrado: $cmoregionid" >> $slclogfile
  echo "ATENCION: FMO se va a construir en base a la info >>>> $regionid" >> $slclogfile
else
  echo "(II):   $regionid >>> $nregionid  [ ${cmoregionid[@]} ]" >> $slclogfile
fi

## SRST
cmosrstid=( $( $GREP $devicepoolid $devicepoolname | $AWK -F "," '{print $5}') )
nsrstid="$( $GREP $srstid $srstname | $WC -l)"

if [ "$srstid" != "${cmosrstid[0]}" ]
then
  echo "(EE):   El nombre de la SRST difiere, el nombre de la región debería ser: $srstid  y hemos encontrado: $cmosrstid" >> $slclogfile
  echo "ATENCION: FMO se va a construir en base a la info >>>> $srstid" >> $slclogfile
  echo -n "ATENCION: Es necesario validar la IP SRST del GW >>>> " >> $slclogfile
  $GREP $srstid $srstname | $AWK -F "," '{printf("%s\n",$2)}' >> $slclogfile
else
  echo "(II):   $srstid >>> $nsrstid  [ ${cmosrstid[@]} ]" >> $slclogfile
fi

## LOCATION
nlocationid="$( $GREP $locationid $locationname | $WC -l)"
cmolocationid=( $( $GREP $devicepoolid $devicepoolname | $AWK -F "," '{print $10}') )

if [ "$locationid" != "${cmolocationid[0]}" ]
then
  echo "(WW):   El nombre de la LOCATION difiere, el nombre de la región debería ser: $locationid  y hemos encontrado: $cmolocationid" >> $slclogfile
else
  echo "(II):   $locationid >>> $nlocationid  [ ${cmolocationid[@]} ]" >> $slclogfile
fi

## MRGL
nmrglid="$( $GREP $mrglid $mrglname | $WC -l)"
cmomrglid=( $( $GREP $devicepoolid $devicepoolname | $AWK -F "," '{print $9}') )

if [ "$mrglid" != "${cmomrglid[0]}" ]
then
  echo "(WW):   El nombre de la MRGL difiere, el nombre de la región debería ser: $mrglid  y hemos encontrado: $cmomrglid" >> $slclogfile
else
  echo "(II):   $mrglid >>> $nmrglid  [ ${cmomrglid[@]} ]" >> $slclogfile
fi

## MRG
nmrgid="$( $GREP $mrgid $mrgname | $AWK -F "," '{print $3}' | $WC -l)"
cmomrgid=( $( $GREP $mrglid $mrglname | $AWK -F "," '{print $3}') )
echo "(II):   $mrgid >>> $nmrgid  [ ${cmomrgid[@]} ]" >> $slclogfile



#nipphones=0
#ipphone[0]=""
#ipphonenameFAKE=$workdirectory/phone.csv
#ipphonename=$workdirectory/phone.mod1.csv
#dnname=$workdirectory/directorynumber.csv

#if [ ! -f $ipphonename ] ## Si el fichero no existe
#then
#  ## Si el fichero NO existe lo creamos
#  echo "WARN:  $ipphonename NO EXISTE ... lo creamos"
#  sed "s/daysBacklightNotActive>1,/daysBacklightNotActive>1 /g" $ipphonenameFAKE > $ipphonename
#fi

## DPLIST
#dplist=( $($GREP $devicepoolid $devicepoolname | $AWK -F "," '{printf("%s ",$1)}') )
## echo "DEBUG:    "${dplist[@]}

#w=0
#for w in "${dplist[@]}"
#  do
#  nipphones="$( $GREP $w $ipphonename | $GREP SEP | $WC -l )"
#  echo "DEBUG:   IP-PHONES: $w    >>>>    $nipphones"

#  ipphoneMAC=( $( $GREP $w $ipphonename | $GREP SEP |$AWK -F "," '{printf("%s ",$3)}' ) )
#  for mac in "${ipphoneMAC[@]}"
#  do
#    ## DEVICE:
#    $GREP $mac $ipphonename | $AWK -F "," '{printf("@@IPPHONE@@,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, ",$3,$4,$5,$6,$9,$22,$23,$24,$25,$43,$47,$55)}'
#    ## LINE #1
#    ipphonel1="$( $GREP $mac $ipphonename | $AWK -F "," '{printf("%s ",$130)}' )"
#    ## LINE #2
#    ipphonel2="$( $GREP $mac $ipphonename | $AWK -F "," '{printf("%s ",$214)}' )"
#    ## LINE #3
#    ipphonel3="$( $GREP $mac $ipphonename | $AWK -F "," '{printf("%s ",$298)}' )"
#    ## LINE #4
#    ipphonel4="$( $GREP $mac $ipphonename | $AWK -F "," '{printf("%s ",$382)}' )"
#    ## LINE #5
#    ipphonel5="$( $GREP $mac $ipphonename | $AWK -F "," '{printf("%s ",$466)}' )"
#    ## LINE #6
#    ipphonel6="$( $GREP $mac $ipphonename | $AWK -F "," '{printf("%s ",$550)}' )"

#    nline=0
#    if [ "$ipphonel1" != " " ]
#    then
#      nline=1
#      $GREP $mac $ipphonename | $AWK -F "," -v line="$nline" '{printf("@@DN%s@@,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,",line,$130,$139,$142,$145,$148,$151,$154,$157,$183,$189,$192,$160,$164,$165,$166,$167,$175,$177,$180,$181,$185)}'
#    fi
#    if [ "$ipphonel2" != " " ]
#    then
#      nline=2
#      $GREP $mac $ipphonename | $AWK -F "," -v line="$nline" '{printf("@@DN%s@@,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,",line,$214,$223,$226,$229,$232,$235,$238,$241,$267,$273,$276,$244,$248,$249,$250,$251,$259,$261,$264,$265,$269)}'
#    fi
#    if [ "$ipphonel3" != " " ]
#    then
#      nline=3
#      $GREP $mac $ipphonename | $AWK -F "," -v line="$nline" '{printf("@@DN%s@@,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,",line,$298,$307,$310,$313,$316,$319,$322,$325,$351,$357,$360,$328,$332,$333,$334,$335,$343,$345,$348,$349,$353)}'
#    fi
#    if [ "$ipphonel4" != " " ]
#    then
#      nline=4
#      $GREP $mac $ipphonename | $AWK -F "," -v line="$nline" '{printf("@@DN%s@@,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,",line,$382,$391,$394,$397,$400,$403,$406,$409,$435,$441,$444,$412,$416,$417,$418,$419,$427,$429,$432,$433,$437)}'
#    fi
#    if [ "$ipphonel5" != " " ]
#    then
#      nline=5
#      $GREP $mac $ipphonename | $AWK -F "," -v line="$nline" '{printf("@@DN%s@@,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,",line,$466,$475,$478,$481,$484,$487,$490,$493,$519,$525,$528,$496,$500,$501,$502,$503,$511,$513,$516,$517,$521)}'
#    fi
#    if [ "$ipphonel6" != " " ]
#    then
#      nline=6
#      $GREP $mac $ipphonename | $AWK -F "," -v line="$nline" '{printf("@@DN%s@@,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,",line,$550,$559,$562,$565,$568,$571,$574,$577,$603,$609,$612,$580,$584,$585,$586,$587,$595,$597,$600,$601,$605)}'
#    fi
#    echo ""
#    echo "INFO: $mac DN#$nline[ L1: >$ipphonel1< L2: >$ipphonel2< L3: >$ipphonel3< L4: >$ipphonel4< L5: >$ipphonel5< L6: >$ipphonel6< ]"
#  done
#done


exit 0
