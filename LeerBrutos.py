from operator import index
import os.path
import re
from os import listdir
from os.path import isfile, join
import pandas as pd
from pandas import ExcelWriter
df=pd.read_excel(r"C:\Users\daniel.chacon\Desktop\ReporteEfetividadLecturas.xlsx")
df=df[["medidor","tipo_suministro","suministro"]]
re_list=df.values.tolist()
for par in re_list:
    if par[0].isnumeric():
        pass
    else:
        re_list.remove(par)
def tipo_medidor(suministro):
    respuesta="sin valor"
    for par in re_list:
        if par[2]==suministro:
            respuesta=par[1]
            #print(par)
            break
    return respuesta
my_path = os.path.abspath(os.path.dirname("LeerBrutos.py"))
onlyfiles = [f for f in listdir('Datos Brutos') if isfile(join('Datos Brutos', f))]
#print(onlyfiles)
combined_csv_extend=pd.DataFrame()
combined_csv_reduc=pd.DataFrame()
campos_list_elster=[]
campos_list_emh=[]
for idx,file in enumerate(onlyfiles):
    #if idx<30:
    path = os.path.join(my_path,r"Datos Brutos",file)
    #print(file)
    with open(path,mode='r') as f:
        data=f.read()
        suministro=re.findall(r"_(.+)_",file)[0]
        tipo_med=tipo_medidor(suministro)
        if len(re.findall(r".255;",data))<10:
            obi_marca="EMH"
            try:
                medidor=re.findall(r"0\.0\.0;([^a-zA-Z]{1,});",data)[0]
            except:
                medidor="sin valor"
            try:
                fecha=re.findall(r"0\.9\.2;(.+?);",data)[0]
                hora=re.findall(r"0\.9\.1;(.+?);",data)[0]
                fecha_hora=f"{fecha} {hora}"
            except:
                fecha_hora='sin valor'
            try:
                eat=re.findall(r"1-1:1\.8\.0;([^a-zA-Z]{1,});",data)[0]
            except:
                eat="sin valor"
            try:    
                eafp=re.findall(r"1-1:1\.8\.1;([^a-zA-Z]{1,});",data)[0]
            except:
                eafp="sin valor"
            try:
                eahp=re.findall(r"1-1:1\.8\.2;([^a-zA-Z]{1,});",data)[0]
            except:
                eahp="sin valor"
            try:
                obi1=re.findall(r"\D1\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi1="sin valor"
            try:
                obi2=re.findall(r"\D21\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi2="sin valor"
            try:
                obi3=re.findall(r"\D41\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi3="sin valor"
            try:
                obi4=re.findall(r"\D61\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi4="sin valor"
            try:    
                obi5=re.findall(r"\D3\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi5="sin valor"
            try:
                obi6=re.findall(r"\D23\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi6="sin valor"
            try:
                obi7=re.findall(r"\D43\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi7="sin valor"
            try:
                obi8=re.findall(r"\D63\.25;([^a-zA-Z]{1,});",data)[0]
            except:
                obi8="sin valor"
            campos_emh={
                'suministro':suministro,
                'medidor':medidor,
                'tipo_med':tipo_med,
                'fecha_hora':fecha_hora,
                'EAT':eat,
                'EAFP':eafp,
                'EAHP':eahp,
                'obi(1.25)':obi1,
                'obi(21.25)':obi2,
                'obi(41.25)':obi3,
                'obi(61.25)':obi4,
                'obi(3.25)':obi5,
                'obi(23.25)':obi6,
                'obi(43.25)':obi7,
                'obi(63.25)':obi8,
            }
            campos_list_emh.append(campos_emh)
        else:
            obi_marca="Elster"
            try:
                medidor=re.findall(r"1\.0\.96\.1\.0\.255;([^a-zA-Z]{1,});",data)[0]
            except:
                medidor="sin valor"
            try:
                fecha_hora=re.findall(r"0\.0\.1\.0\.0\.255;(.+?);",data)[0]
            except:
                fecha_hora="sin valor"
            try:
                eat=re.findall(r"1\.1\.1\.8\.0\.255;([^a-zA-Z]{1,});",data)[0]
            except:
                eat="sin valor"
            try:
                eafp=re.findall(r"1\.1\.1\.8\.1\.255;([^a-zA-Z]{1,});",data)[0]
            except:
                eafp="sin valor"
            try:
                eahp=re.findall(r"1\.1\.1\.8\.2\.255;([^a-zA-Z]{1,});",data)[0]
            except:
                eahp="sin valor"
            try:    
                obi1=re.findall(r"1\.1\.2\.8\.0\.255;([^a-zA-Z]{1,});",data)[0]
            except:
                obi1="sin valor"
            try:
                obi2=re.findall(r"1\.1\.2\.8\.1\.255;([^a-zA-Z]{1,});",data)[0]
            except:
                obi2="sin valor"
            try:   
                obi3=re.findall(r"1\.1\.2\.8\.2\.255;([^a-zA-Z]{1,});",data)[0]
            except:
                obi3="sin valor"
            campos_eslter={
                'suministro':suministro,
                'medidor':medidor,
                'tipo_med':tipo_med,
                'fecha_hora':fecha_hora,
                'EAT':eat,
                'EAFP':eafp,
                'EAHP':eahp,
                'obi(1.1.2.8.0.255)':obi1,
                'obi(1.1.2.8.1.255)':obi2,
                'obi(1.1.2.8.2.255)':obi3,
            }
            campos_list_elster.append(campos_eslter)
        #print(f"{obi_marca} {file}")
df1=pd.DataFrame(campos_list_elster)
df2=pd.DataFrame(campos_list_emh)
writer = pd.ExcelWriter('brutos_ejm.xlsx', engine='xlsxwriter')
df1.to_excel(writer, sheet_name='ELSTER',index=False)
df2.to_excel(writer, sheet_name='EMH',index=False)
writer.save()
