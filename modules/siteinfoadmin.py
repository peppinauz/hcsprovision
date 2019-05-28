#!/usr/bin/env python3

import openpyxl
import json

## ####################
CMOSITEDATA="../CMO/AGENCIAS-CISCO_HCS.xlsx"

## ####################
## POSICIONES EN EXCEL
COD_POS=0
NOME_POS=1
JUNTORES_POS=4
RAMAIS_POS=5
REGIONAL_POS=6
CIDADE_POS=7
TIPO_POS=8
LOGRADOURO_POS=9
BAIRRO_POS=10
UF_POS=11
CEP_POS=12
PHONE_POS=13

## STATIC data
BRAZIL_CC="+55"

## SITE info para SI
MAPPING_SITEINFO_TO_SI=['uniorg','regional','cidade','tipo','logradouro','bairro','uf','cep']


## Description: Cojemos toda la info del SITE
## IN: JSON con la info admin del site
## OUT:
## 1) String con toda la info del site
def generate_si_siteadmin(info):
    hashinfo=""

    for map in MAPPING_SITEINFO_TO_SI:
        hashinfo=hashinfo+map[:2]+":"+info[map]+","
    return hashinfo[:-1]


## Description: Abre el fichero excel, extra la info del site
## IN: SLC + logfile
## OUT:
## 1) JSON: con la info del site
## 2) JSON: con numeraciones

def get_cmositedata(slc,logfile):
    blk = openpyxl.load_workbook(CMOSITEDATA,read_only=True)
    sheet = blk["Plan1"]

    ## Definicion de datos:
    siteinfo={}
    sitenum={}

    for row in sheet.rows:
        #print("<<>>",row[0].value)
        #DEBUG
        #print(sheet.cell(row=fila,column=col).value)
        rslc=str(row[0].value)
        if rslc[1:].startswith(slc):
            ## EXTRAEMOS INFO DEL SITE
            siteinfo['name']=slc+"-"+row[NOME_POS].value
            siteinfo['uniorg']=str(row[COD_POS].value)
            print(row[COD_POS].value,str(row[COD_POS].value),siteinfo['uniorg'])
            siteinfo['juntores']=row[JUNTORES_POS].value
            siteinfo['ramais']=row[RAMAIS_POS].value
            siteinfo['regional']=row[REGIONAL_POS].value
            siteinfo['cidade']=row[CIDADE_POS].value
            siteinfo['tipo']=row[TIPO_POS].value
            siteinfo['logradouro']=row[LOGRADOURO_POS].value
            siteinfo['bairro']=row[BAIRRO_POS].value
            siteinfo['uf']=row[UF_POS].value
            siteinfo['cep']=row[CEP_POS].value
            siteinfo['Address1']=generate_si_siteadmin(siteinfo)

            ## EXTRAEMOS INFO de NUMERACION
            sitenum['cc']=BRAZIL_CC
            sitenum['ac']=row[PHONE_POS].value[1:3]
            sitenum['head']=row[PHONE_POS].value[5:]
            sitenum['slc']=slc
            ## External Phone Number Mask: Country Code + Area Code + Numeracion nacional
            sitenum['epnm']=BRAZIL_CC+row[PHONE_POS].value[1:3]+row[PHONE_POS].value[5:]

            print("SITEINFOADMIN@GET_CMOSITEDATA -----------------------------", file=logfile)
            print(json.dumps(siteinfo,sort_keys=True,indent=2), file=logfile)
            print("SITEINFOADMIN@GET_CMOSITEDATA -----------------------------", file=logfile)
            print(json.dumps(sitenum,sort_keys=True,indent=2), file=logfile)
            print("SITEINFOADMIN@GET_CMOSITEDATA -----------------------------", file=logfile)

    return siteinfo,sitenum
