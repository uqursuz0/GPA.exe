# -*- coding: utf-8 -*-
"""
Created on Mon Jun 27 02:07:52 2022

@author: U0UR0
"""

import xlsxwriter
import pandas as pd
url = input("Lisans ders planı url adresini giriniz")
dec=input("Harf notları transkripten aktarılsın mı? Bu özelliği kullanmak için transkriptinizin html formatında kaydedilmesi gerekmektedir.Evet için 1 hayır için 0 yazın.")
if dec=="1":
    tc=input("Eğer vermiş olduğunuz deersleri transkripten otomatik olarak doldurmak istiyorsanız okuyucu.exe'nin olduğu konuma transkriptinizi html formatında kaydedin ve ismini yazın.")
    if tc[-5:]!=".html":
        tc=tc+".html"
    dec=tc
def tcvar(dec):
    try:
        transcript=pd.read_html(tc,index_col=0,header=0)
        for j,k in transcript[6].iterrows():
            if i[0]==j:
                ws.write_string(rowa,3,k[2])
            else:
                ws.write_blank(rowa, 3,"")
                
                
    except:
        ws.write_blank(rowa, 3,"")
        
        
tablelist = pd.read_html(url, index_col=0, header=0)

courselist =[]


row_num=2


#OKUMA ÇÖZME
for i in tablelist:
    for j,k in i.iterrows():
        smallist=[j,k[0],k[1],""] 
        hnotu ='=IF(D'+str(row_num)+'="AA",4,IF(D'+str(row_num)+'="BA",3.5,IF(D'+str(row_num)+'="BB",3,IF(D'+str(row_num)+'="CB",2.5,IF(D'+str(row_num)+'="CC",2,IF(D'+str(row_num)+'="DC",1.5,IF(D'+str(row_num)+'="DD",1,-1)))))))'
        w_avg="=C"+str(row_num)+"*E"+str(row_num)
        ffornot="=IF(F"+str(row_num)+">0,F"+str(row_num)+",0)"
        givencredit="=IF(F"+str(row_num)+">-1,C"+str(row_num)+",0)"
        average="ORTALAMA AŞAĞIDA"
        smallist.append(hnotu)   
        smallist.append(w_avg)   
        smallist.append(ffornot)   
        smallist.append(givencredit)   
        smallist.append(average)   
        courselist.append(smallist)
        row_num+=1




wb=xlsxwriter.Workbook("GPA.xlsx")
ws=wb.add_worksheet("Harf Notları")

rowa=1
format=wb.add_format()
formata=wb.add_format()
bold=wb.add_format()
nann=wb.add_format()
format.set_font_color('white')
format.set_hidden(hidden=True)
formata.set_num_format('#.#0')
bold.set_bold()
nann.set_bg_color("yellow")
for i in courselist:
    if str(i[0])=="nan":
        ws.write_string(rowa,0,str(i[0]),nann)
        ws.write_string(rowa,1,str(i[1]),nann)
        ws.write_number(rowa,2,float(i[2]),nann)

    else:
        ws.write_string(rowa,0,str(i[0]),bold)
        ws.write_string(rowa,1,str(i[1]),bold)
        ws.write_number(rowa,2,float(i[2]),bold)
        
    tcvar(dec)
            
    
    # ws.write_formula(("J"+str(rowa+1)),'=IF(D'+str(row_num)+'="AA",1,IF(D5="BB",3,8))')
    ws.write_formula(("E"+str(rowa+1)),i[4],format)
    ws.write_formula(("F"+str(rowa+1)),i[5],format)
    ws.write_formula(("G"+str(rowa+1)),i[6],format)
    ws.write_formula(("H"+str(rowa+1)),i[7],format)
    ws.write_formula("I"+str(rowa+1),"G"+str(row_num)+"/H"+str(row_num),formata)
    rowa+=1

templist=["Toplam Kredi Sayısı=","","=SUM(C2:C"+str(row_num-1),"GPA*Credit Toplamı","ve","Verilen Kredi Sayısı","=SUM(G2:G"+str(row_num-1),"=SUM(H2,H"+str(row_num-1),"=G"+str(row_num-1)+"/G"+str(row_num-1)]
for i in templist:
    # ws.write(rowa,0,str(i[0]))
    # ws.write_string(rowa,1,str(i[1]))
    ws.write_array_formula("C"+str(rowa+1)+"C"+str(rowa+1),'=SUM(C2:C'+str(rowa)+')')
    ws.write_array_formula("G"+str(rowa+1)+"C"+str(rowa+1),'=SUM(G2:G'+str(rowa)+')')
    ws.write_array_formula("H"+str(rowa+1)+"H"+str(rowa+1),'=SUM(H2:H'+str(rowa)+')')
    ws.write_formula("I"+str(rowa+1),"G"+str(rowa+1)+"/H"+str(rowa+1),bold)



#DIŞA AKTARIM KISMI
columns=["Code","Course","Credit","Harf Notu Yazınız","Sayısal Harf Notu","Ağırlıklı Kredi","Alınan","Alınan Kredi","Ortalama"]
for c in range(len(columns)):
    ws.write_string(0,c,columns[c],bold)
wb.close()