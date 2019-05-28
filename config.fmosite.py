## LINUX#!/usr/bin/python3
## OSX
#!/usr/local/bin/python3

import csv
import openpyxl
import time
import sys
import json

if len(sys.argv) != 5:
    print("ERROR: python3 <nombre-fichero.py> <datos de entorno> <cmo site data> XXXX <log-config-file>")
    print(">>>>  Datos de entorno: ",sys.argv[1])
    print(">>>>  [2] cmo site data: <nombre-fichero-con-datos de config CMO>")
    print(">>>>  [3] XXXX: SLC")
    print(">>>>  [4] config log file: ../FMO/XXXX/logconfigfilename-.txt")
    exit(1)
else:
    fmostaticdata=sys.argv[1]   ## Datos de entorno: Maqueta/Produccion/...
    siteconfig=sys.argv[2]      ## Datos de configuracion del SITE
    slc=sys.argv[3]             ## INPUTS: SLC
    logconfigfile=sys.argv[4]   ## LOG config file

## Ficheros de entrada: config CMO
cmgfile="doc/cmgroup.txt"

## LOG File
f = open(logconfigfile, 'a')

print("(II) #################################################", file=f)
print("(II) #################################################", file=f)
print("(II) INPUT CMO configuration      ::",siteconfig, file=f)
print("(II) INPUT datos de entorno       ::",fmostaticdata, file=f)
print("(II) INPUT SLC                    :: ",slc, file=f)
print("(II) INPUT LOG configuration file :: ",logconfigfile, file=f)
print("(II) GENERANDO PRIMER BLK de FMO", file=f)
print("(II) Preparando datos para FMO", file=f)

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

with open(siteconfig,'r') as infile:
    data = json.load(infile)
infile.close()

## DEBUG
#print(data)
#print("")

#data['dp'] = []
#data['loc']=[]
#data['mrgl'] = []
#data['mrg'] = []
#data['srst'] = []
#data['gw'] = []
#data['rsc'] = []
e164=data['e164']
agencia=data['fmosite']
devices=data['devices']
ndevices=devices[0]['phones']

print("(II) Buscando CallManager Group", file=f)
################################################################
######## CMG
with open(cmgfile, "r") as fp:
    cmg = json.loads(fp.read())  ## variable de tipo diccionario
    #print("<<<<   ",cmg[key])
fp.close()

maxdevice=int(cmg['max'])
#print("<<<<>>>>>>   ",maxdevice)

fmocmg="xxx"
total=0

if ndevices !=0:
    # Buscamos sitio
    for x in cmg.keys():
        #print("--->>>",x,cmg[x])
        num=int(cmg[x])
        if num < maxdevice:
            #print("--->>> ",num)
            if "max" not in x:
                total=num+ndevices
                if total <= maxdevice and "xxx" in fmocmg: ## total menor o igual
                    #print("--->>>--->>> ",x)
                    cmg[x]=str(total)
                    fmocmg=x
else:
    #exit(1)
    print("(EE): no hay phones en este site")
    print("Realmente quieres seguir adelante?? (S/n")

with open (cmgfile, "w") as fp:
    fp.write(json.dumps(cmg))
fp.close()



################################################################

fmonagencia=agencia['name']
cabecera=e164[0]['head']
nagencia=agencia['name'][5:]
nnum=cabecera
nac=e164[0]['ac']
areacode=e164[0]['ac']
xcambio=""
ycambio=""
scambio=""
running=True


print("\n\n")
print("Se va a hacer la provisión para SLC :: \t\t",slc)
print("En CMO se ha encontrado el nombre de agencia :: \t",agencia['name'])
print("\n")
print("En FMO el nuevo site se llamará :: \t\t\t",fmonagencia)
print("\nEn CMO se ha encontrado External Phone Number Mask :: \t",e164[0]['epnm'])
print("En FMO se va a configurar número de cabecera :: \t",cabecera)
################################################################
print("\n")
xcambio=input("(II) Quieres cambiar algo? (S/n) ")

if xcambio == 'S':
    while running:
        opcion=int(input("\n\n Menu de cambios: \n[1] Site name? \n[2] Número de cabecera \n[0] Salir \n"))
        if opcion == 1:     ## Cambiar nombre del site
            print("\n\nEste es el nombre que hemos detectado >>> ",agencia[0]['name'],"\n")
            print("Indicar el nuevo nombre de la agencia, por ejemplo:")
            print("\tSao Paolo, Morumbi,...")
            scambio=input("Nuevo nombre: ")
            print("\nEste es el NUEVO nombre de la agencia >>>",scambio)
            print("Este es el NUEVO nombre del site en FMO >>>",slc,"-",scambio)
            cinput=input("\t Quieres continuar? (S/n)")
            if cinput == 'S':
                nagencia=scambio
                fmonagencia=slc+"-"+scambio
                print(">>>> Datos cambiados:",fmonagencia)
            else:
                print(">>>> Descartando cambios")

        elif opcion == 2:   ## Cambiar numero de cabecera
            print("Este es el número que hemos detectado >>> ",cabecera)
            print("El dial-plan de HCS implica que el formato sea +E164: +57AAXXXXYYYY")
            print("Donde: \n\t AA = Area Code \n\t XXXX = rango publico \n\t YYYY = Extension")
            nnum=input("Nuevo número: ")
            if not nnum.startswith('+'):
                print(">>> Tiene que empezar por +")
            elif len(nnum) != 13:
                print(">>> El formato del número tiene tener 12 dígitos")
            else:
                print("Este es el NUEVO número de cabecera para FMO >>>",nnum)
                cinput=input("\t Quieres continuar? (S/n) ")
                print("El número de cabecerá que se configurará >>> \t",nnum)
                cabecera=nnum
        elif opcion == 0:
            running=False
            ##

    ## Si ha habido algún cambio
    if not running: ## Se ha ejecutado el menu de cambios y volvemos a imprimir la info
        print("\n\n")
        print("Se va a hacer la provisión para SLC >>> \t\t",slc)
        print("Este es el nombre que se ha asignado a la agencia >>> \t",nagencia)
        #print("\n")
        print("En FMO el nuevo site se llamará >>> \t\t\t",fmonagencia)
        print("\nEn CMO se ha encontrado External Phone Number Mask >>> \t",e164[0]['epnm'])
        print("Este es el número de cabecera >>> \t\t\t",cabecera)
        print("\n")
        input("Pulsa INTRO para continuar ")

################################################################
print("(II) Estado de asignaciones : CM groups : MAX=",cmg['max'], file=f)
for x in cmg.keys():
    if "max" not in x:
        print("(II) ",x," >> ",cmg[x], file=f)
print("(II)", file=f)
print("(II) En CMO se ha encontrado PHONES : \t\t\t",devices[0]['phones'], file=f)
print("(II) En CMO se ha encontrado Dev Prof : \t\t\t",devices[0]['udp'], file=f)
print("(II) CMG asignado : \t\t\t\t\t",fmocmg, file=f)
print("(II) CMG asignado : \t\t\t\t\t",fmocmg) ## A peticion de Delivery
print("(II)", file=f)

################################################################
## Rellenamos el BLK del SITE.00
#FMO working path:
sitepath="../FMO/"+slc
templateblkfile = "blk/00.site-template.xlsx"   ## FMO SITE TEMPLATE
outputblkfile = sitepath+"/00.site."+slc+".xlsx"    ## FMO BLK OUTPUT

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

# FMO File OUTPUT DATA
blk = openpyxl.load_workbook(templateblkfile)

## VARIABLES
fila=5

sheet = blk["SITE"]
## datos estaticos (SIN DATAINPUT) Hierarchy, site,...
sheet['B'+str(fila)]=fmoenvconfig['hierarchynode']
sheet['C'+str(fila)]="add"
#sheet['D'+str(fila)]="name:"+row[3] # Search field
## datos dinamicos
sheet['N'+str(fila)]="true"                             # hcsSite.isModifiable
sheet['p'+str(fila)]="CustomerLocation"                 # hcsSite.type
sheet['t'+str(fila)]="false"                            #hcsSite.isDefaultLocation
sheet['i'+str(fila)]=data['fmosite']['uniorg']          #hcsSite.ExternalID
sheet['u'+str(fila)]="PROVISIONADA"                     #hcsSite.Description
sheet['v'+str(fila)]=data['fmosite']['Address1']        #hcsSite.sho
sheet['w'+str(fila)]=fmonagencia                        #hcsSite.ExtendedName
#sheet['y'+str(fila)]="false"                           #hcsSite.isDeletable
sheet['z'+str(fila)]=fmoenvconfig['hcsSitecustomer']    #hcsSite.customer
sheet['aa'+str(fila)]=fmoenvconfig['fmonetworklocale']  #hcsSite.country
sheet['ae'+str(fila)]=fmoenvconfig['ndlr']              #ndlr
sheet['af'+str(fila)]=fmonagencia                       #name
sheet['ag'+str(fila)]=fmonagencia                       #hcsSiteDAT.name
sheet['ah'+str(fila)]="true"                            # hcsSiteDAT.push_cucm
sheet['ai'+str(fila)]="true"                            # hcsSiteDAT.create_admin @@@
sheet['al'+str(fila)]="false"                           #hcsSiteDAT.migrate
sheet['am'+str(fila)]=fmoenvconfig['fmonetworklocale']  #country
sheet['an'+str(fila)]=fmonagencia                       # siteNdlr.name
sheet['ao'+str(fila)]=fmoenvconfig['siteNdlrreference'] #siteNdlr.reference
sheet['ak'+str(fila)]=fmoenvconfig['hierarchynode']     #hcsSiteDAT.HierarchyPath
sheet['ap'+str(fila)]=fmoenvconfig['adminuserpass']     #data_user.password
sheet['aq'+str(fila)]=fmoenvconfig['hcsrole']           #hcs_role.clonedRole
sheet['ar'+str(fila)]=fmonagencia+"SiteAdmin"           #hcs_role.role
sheet['at'+str(fila)]="false"                           #hcs_role.role
sheet['as'+str(fila)]="en-us"                           #hcs_role.role

## FMO File OUTPUT DATA: Close
blk.save(outputblkfile)

###### Generado primer BLK para SITE: 00.SITE.SLC.xls
fmoidagencia=input("(II) Inserta el ID de la agencia? ")  ## Temporal

################################################################
## Guardamos datos FMO
## SITE
data['fmosite']['name']=fmonagencia
data['fmosite']['id']=fmoidagencia
data['fmosite']['cmg']=fmocmg
## E164
data['e164'][0]['head']=cabecera
data['e164'][0]['ac']=areacode

## DEBUG
print("CONFIG.FMOSITE@MAIN -----------------------------", file=f)
print(json.dumps(data,sort_keys=True,indent=2), file=f)
print("CONFIG.FMOSITE@MAIN -----------------------------", file=f)

with open(siteconfig,'w') as outfile:
    json.dump(data,outfile)
outfile.close()

## LOG de CONFIGURACION
f.close()

exit(0)
