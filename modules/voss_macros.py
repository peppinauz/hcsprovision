

## ["coorporativo-cl8-ndl", "hcs.TGSOL.VIVO.Produban.Coorporativo"]
SITE_NDLR="{{ macro.SITE_NDLR }}"
SITE_LOCALE="{{ data.SiteDefaultsDoc.defaultNL }}"
SITE_NDLR_REFERENCE="{{ data.NetworkDeviceListReference.reference}}"

## ID Numerico del customer
CUSTOMER_ID="{{ macro.HcsDpCustomerId }}"
## ID Numerico del site
#SITE_NUM_ID="{{ data.BaseSiteDAT.InternalSiteID }}"
SITE_NUM_ID="{{ macro.HcsDpSiteId }}"
## CuXXXSiYYY
CUXSIY="{{ macro.HcsDpUniqueSitePrefixMCR }}"
## CuXXX
CUX="{{ macro.HcsDpUniqueCustomerPrefixMCR }}"

## SITE Default CMG
SITE_CMG="{{ data.SiteDefaultsDoc.defaultcucmgroup }}"

## Funcion de Fecha/hora actual
NOW="{{fn.now}}"
