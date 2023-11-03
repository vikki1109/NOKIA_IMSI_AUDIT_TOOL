import os 
import sys 
import pandas as pd
import numpy as np 
import zipfile
import re
from openpyxl import load_workbook
import openpyxl
import xlsxwriter
from pandas import ExcelWriter
from openpyxl.styles import PatternFill


script_dir = os.path.abspath('')

pd.set_option('display.max_columns',None) 
pd.set_option('display.max_rows',None)
pd.set_option('display.max_colwidth',None)

xls= pd.ExcelFile(os.path.join(script_dir , "IMSI_OUTPUT.xlsx"))
sheet_names= xls.sheet_names
sheet_names

df=pd.read_excel(xls,'sheet1')

MSRNHPLMN=20-30;IntMSRNVPLMN=30-90;NatMSRNVPLMN=20-90
NRE='Allow'; BLE='Block' 
GLE='Allow' or 'TRACE'
UIE='Block';CIPH='USED'
TAL=1 ;TLOC=5; TPER=5; TAIA=1 ;TMOC=10; TMOS=10; TMTC=10; TMTS=10; TAML=0
TAMU=0 ;TASO=0;  TCSF=0 ;AAL=1; ALOC=5; APER=5 ;AAIA=1; AMOC=0; AMOS=10; AMTC=10 ;AMTS=10; AAML=0; AAMU=0; AASO=0; ACSF=0; IAL=1; ILOC=5
IPER=1; IAIA=1; IMOC=10 ;IMOS=10; IMTC=10; IMTS=15; IAML=0; IAMU=0; IASO=0; ICSF=1; MTRF='USED'
data_dict = df.to_dict()
ciph = data_dict['CIPHERING']
nre1 = data_dict['No Response effect']
ble1 = data_dict['Black List Effect']
gle1 = data_dict['Grey List Effect']
uie1=data_dict['Unknown IMEI CHECK']
pnst1=data_dict['PNS TIME LIMIT']
msrn1=data_dict['MSRN LIFE TIME']
emcu1=data_dict['EXACT MS CATEGORY USAGE']
rcur1=data_dict['REJECT CAUSE FOR UDL REJECTION']
sob1=data_dict['SUPPORT OF BOR']
zcfr=data_dict['ZONE CODES FROM HLR']
rr=data_dict['REGIONAL ROAMING']
# import pdb; pdb.set_trace()

tloc1 = data_dict['TLOC']
tper1 = data_dict['TPER']
tmoc1 = data_dict['TMO CALL']
tmos1 = data_dict['TMO SMS']
tmtc1 = data_dict['TMT CALL']
tmts1 = data_dict['TMT SMS']
aloc1 = data_dict['ALOC']
aper1 = data_dict['APER']
amoc1 = data_dict['AMO CALL']
amos1 = data_dict['AMO SMS']
amtc1 = data_dict['AMT CALL']
amts1 = data_dict['AMT SMS']
iloc1 = data_dict['ILOC']
iper1= data_dict['IPER']
imoc1 = data_dict['IMO CALL']
imos1 = data_dict['IMO SMS']
imtc1 = data_dict['IMT CALL']
imts1 = data_dict['IMT SMS']
mtrf1 = data_dict['MOBILE TERMINATING ROAMING FORWARDING']
tmts1=data_dict['TMT SS OPER']
# import pdb; pdb.set_trace()

styler = df.style
for key,values in ciph.items():
    if 'USED' in  ciph[key]:
        x=ciph[key]
        print('o')
        styler = styler.applymap(lambda x : 'background-color: green' if x =='USED'  else '')
    
for key,values in nre1.items():
    if 'ALLOW'   in  nre1[key]:
        x1=nre1[key]
        styler = styler.applymap(lambda x1 : 'background-color: green' if x1 =='ALLOW'  else '')
        print('ok')
for key,values in ble1.items():
    if 'BLOCK' in  ble1[key]:
        x2=ble1[key]
        styler = styler.applymap(lambda x2 : 'background-color: #66CD00' if x2 =='BLOCK'  else '')
        print('ok')

for key,values in pnst1.items():
    if pnst1[key]==20 :
        x6=pnst1[key]
        styler = styler.applymap(lambda x6 : 'background-color: green' if x6 == 20  else '')
        print('ok')
    if pnst1[key]==30 :
        x32=pnst1[key]
        styler = styler.applymap(lambda x32: 'background-color: green' if x32== 30  else '')
        print('ok')
    

    if 30 == msrn1[key]:
        x7=msrn1[key]
        styler = styler.applymap(lambda x7 : 'background-color: green' if x7 ==30  else '')
        print('ok')
    if 90== msrn1[key]:
        x31=msrn1[key]
        styler = styler.applymap(lambda x31 : 'background-color: green' if x31 ==90 else '')
        print('ok')
    if 40== msrn1[key]:
        x32=msrn1[key]
        styler = styler.applymap(lambda x32 : 'background-color: green' if x32 ==40 else '')
        print('ok')
    if 60== msrn1[key]:
        x33=msrn1[key]
        styler = styler.applymap(lambda x33 :'background-color: green' if x33 ==60 else '')
        print('ok')

    if 5 == tloc1[key]:
        x11=tloc1[key]
        styler = styler.applymap(lambda x11: 'background-color: green' if x11 ==5  else '')
        print('ok')
    # if  5==  tper1[key]:
    #     x12=tper1[key]
    #     styler = styler.applymap(lambda x12: 'background-color: red' if x12 ==5  else '')
    #     print('ok')
    if  10==  tmoc1[key]:
        x13=tmoc1[key]
        styler = styler.applymap(lambda x13: 'background-color: #66CD00' if x13 ==10  else '')
        print('ok')

    if  1==  iper1[key]:
        x24=iper1[key]
        styler = styler.applymap(lambda x24: 'background-color:	#66CD00' if x24 ==1  else '')
        print('ok')
   
    if 15==  imts1[key]:
        x28=imts1[key]
        styler = styler.applymap(lambda x28: 'background-color: green' if x28 ==15  else '')
    if 0==  tmts1[key]:
        x29=tmts1[key]
        styler = styler.applymap(lambda x29: 'background-color: green' if x29 ==0 else '')

    

styler.to_excel('IMSI_HIGHLIGHT.xlsx', engine='openpyxl', index=False)
