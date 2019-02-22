#!/bin/bash
#
# if [ $# -ne 1 ]
#    then
#	echo "$0 SLC " >&2
#	echo "  --> Necesito por lo menos el SLC " >&2
#  echo "  --> Formato: XXXX, 4 dígitos " >&2
#        exit 1    
#    fi
#
## Variables entrada
#slc=$1

## Variables
#slcencontrado=0
#match=0

## Tools
AWK=/usr/bin/awk
GREP=/usr/bin/grep
WC=/usr/bin/wc
#CAT=/usr/bin/cat

#############################################################################################
#############################################################################################
#############################################################################################
##
## Directorio de trabajo 
##
base=".."
finaldir="$base/FMO/phonekeys"
now=$(date +"%Y-%m-%d-%H-%M")
pbtlogfile="$finaldir/pbt-$now.log"
sktlogfile="$finaldir/skt-$now.log"
allsktfile="$finaldir/allskt-$now.log"
fmosktfile="$finaldir/fmoskt.csv"
fmopbtfile="$finaldir/fmopbt.csv"
fmopbtfileshort="$finaldir/fmopbtshort.csv"

if [ ! -d $finaldir ]; then
  mkdir $finaldir
fi
touch $pbtlogfile
touch $sktlogfile
touch $allsktfile
touch $fmosktfile
touch $fmopbtfile
touch $fmopbtfileshort

## Chequeamos en todos los clusters
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

## Phones
phonetype[1]="Cisco 6941"
phonetype[2]="Cisco 8945"
phonetype[3]="Cisco 7821"
phonetype[4]="Cisco 7942"
phonetype[5]="Cisco 7911"
phonetype[6]="Cisco 7941"
phonetype[7]="Cisco 7961"
phonetype[8]="Cisco 9971"
phonetype[9]="Cisco 9962"


## SOFTKEY TEMPLATE: REPORTS
echo ">>> SOFTKEY TEMPLATE"
for i in "${cl[@]}"
do
  skt="$i/softkeytemplate.csv"
  #phone="$i/phone.csv"
  #devprof="$i/deviceprofile.csv"

  $AWK -v file="$skt" -F "," '{print file"|" $1}' $skt  | grep -v "NAME" >> $sktlogfile

done

## PHONE BUTTON TEMPLATE: REPORTS
echo ">>> PHONE BUTTON TEMPLATE"
for j in "${phonetype[@]}" ## Para cada tipo de telefono
do
  for i in "${cl[@]}"
  do
    
    pbt="$i/phonebuttontemplate.csv"
    phone="$i/phone.csv"
    #echo "############################################################"

    #IFS=$'\n' pbtlist=( $( $AWK -F "," '/Cisco 6941/{print $1}' $pbt ) )
    IFS=$'\n' pbtlist=( $( $GREP "$j" $pbt | $AWK -F "," '{print $1}' ) )
  
    for z in "${pbtlist[@]}"
    do
      number="$( $GREP "$z" $phone | $WC -l )"
      #echo $pbt">>> $j >>> $z : $number <<<"
      #if [ "$number" != "0" ]
      #then
      #  echo $pbt">>> $j >>> $z : $number"
      #  echo $pbt",$j,$z,$number" >> $pbtlogfile
      #fi
      #echo $pbt">>> $j >>> $z : $number"
      echo $pbt",$j,$z,$number" >> $pbtlogfile

    done
  done
done

## SOFTKEY TEMPLATE & PHONE BUTTON TEMPLATE: renombrado + colapsado en unico file
echo ">>> SOFTKEY TEMPLATE & PHONE BUTTON TEMPLATE"

skt1="$base/CMO/Cluster01/softkeytemplate.csv"
pbt1="$base/CMO/Cluster01/phonebuttontemplate.csv"

head -1 $skt1 > $fmosktfile
head -1 $pbt1 > $fmopbtfile

## PBT list usados
pbtlist=( $($AWK -F "|" '{if ($4 !~ 0) print $3}' $pbtlogfile))

for i in "${cl[@]}"
do
  skt="$i/softkeytemplate.csv"
  pbt="$i/phonebuttontemplate.csv"
  
  ncluster="$( echo $i | sed 's/CMO//g' | sed 's/Cluster//g' | sed 's/\.//g'  | sed 's/\///g' )"
  #echo $ncluster
  $GREP -v "NAME" $skt | $AWK -v cl="$ncluster" '{printf("%s-%s\n",cl,$0)}' >> $fmosktfile
  ## P B T
  #$GREP -v "NAME" $pbt | $AWK -v cl="$ncluster" '{printf("%s%s\n",cl,$0)}' #>> $fmopbtfile
  #$GREP "Santander" $pbt | $AWK -v cl="$ncluster" '{printf("%s%s\n",cl,$0)}' #>> $fmopbtfile
  #$AWK -v cl="$ncluster" 'match($1,/Santander/){printf("%s-%s\n",cl,$0)}' $pbt >> $fmopbtfile
  #$AWK -v cl="$ncluster" 'match($1,/DevP/){printf("%s-%s\n",cl,$0)}' $pbt >> $fmopbtfile
  $GREP -v "NAME" $pbt | grep -v "Individual" | $AWK -v cl="$ncluster" '{printf("%s-%s\n",cl,$0)}' >> $fmopbtfile

done

#for i in "${pbtlist[@]}"
#do
#  #echo ">>>"$i
#  ## Sacamos el listado de los PBT usados
#  $grep "$i" $fmopbtfile >> $fmopbtfileshort
#done

## Quitamos los attributos: MAC OSX
## xattr -c <nombre-fichero>
xattr -c $fmosktfile
xattr -c $fmopbtfile

exit 0
