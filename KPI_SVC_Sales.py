import pandas as pd
import numpy as np
import xlrd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date,timedelta
import msoffcrypto
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
from matplotlib.ticker import MaxNLocator
from dateutil.relativedelta import relativedelta
import calendar

server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To']='ethan.son@lge.com, jongseop.kim@lge.com, taehee.chang@lge.com, iggeun.kwon@lge.com, joonseok.ahn@lge.com, eunbi1.yoon@lge.com, soonan.park@lge.com, isaac.milad@lge.com, aaron1.garcia@lge.com, russell.wilson@lge.com, remoun.abdo@lge.com, cagan.paker@lge.com, ryan.parker@lge.com, dharmin.mistry@lge.com, matthew.sohn@lge.com,thomas.kenney@lge.com,jomar.acostatorres@lge.com, min1.park@lge.com, jiyoon1.heo@lge.com, seungjae.cho@lge.com,william.heidorn@lge.com,alfonza.hall@lge.com,timothy.knight@lge.com'

# Target 데이터 추출
fsvc_target=1855
fsvc_Ctarget=1390
fsales_target=41292

tsvc_target=1968
tsvc_Ctarget=1203
tsales_target=42004



# 달 이름 정하기
# 매달 1일에는 마지막 달 마지막 날이 추출되는 조건 추가
today=date.today()
k=today.strftime('%d')

if k=='01':
    print("first day of month")
    exp=today-relativedelta(days=1)
    startM0=exp.replace(day=1)
    pass_businessday=len(pd.bdate_range(startM0,exp))#이미 1뺐으니깐 더이상 뺄 필요 없음
    endM0=today-relativedelta(days=1)
    total_businessday=len(pd.bdate_range(startM0,endM0))
    today=today-relativedelta(days=1)
    

else:
    print("not the first day of month")
    exp=today
    startM0=exp.replace(day=1)
    pass_businessday=len(pd.bdate_range(startM0,today))-1#어제 기준으로 보기때문에 1 빼기
    endM0=startM0+relativedelta(months=1)-relativedelta(days=1)
    total_businessday=len(pd.bdate_range(startM0,endM0))

print("계산식 확인")
print(startM0)
print(endM0)
print("pass_businessday")
print(pass_businessday)
print("total_businessday")
print(total_businessday)
print("\n")

Data0M=today.strftime('%Y-%m')
date0M_name=today.strftime('%y.%m')

date1M_name=today-relativedelta(months=1)
Data1M=date1M_name.strftime('%Y-%m')
date1M_name=date1M_name.strftime('%y.%m')

date2M_name=today-relativedelta(months=2)
Data2M=date2M_name.strftime('%Y-%m')
date2M_name=date2M_name.strftime('%y.%m')

#Subject 꾸미기
today=date.today()
today=today.strftime('%m/%d')
msg['Subject']='[Daily KPI Report : closed on '+today+"] Services & Sales Status"

###########Body Data 추출
# file 열기
today=date.today()
today=today.strftime('%m%d')
data_loc='//us-so11-na08765/R&D Secrets/M+3 task/000 Sub KPI/Daily FFR Data/Daily FFR_'+date0M_name+'/Daily_'+today+'.xlsx'
#data_loc='//us-so11-na08765/R&D Secrets/M+3 task/000 Sub KPI/Daily FFR Data/Daily FFR_22.02/Daily_0201.xlsx'

file = msoffcrypto.OfficeFile(open(data_loc, "rb"))
file.load_key(password="11111") # Use password
decrypted = io.BytesIO()
file.decrypt(decrypted)
data = pd.read_excel(decrypted)


# 파일 정리하기
data=data.drop(['Unnamed: 0','Unnamed: 1', 'Unnamed: 2'],axis=1)
data=data.drop([0,1,2,3,4,5,6,7,11,12,13,14,15,18,19,20,21,22,23,24,25,26,27,28,29],axis=0)
data.index=['Date','FL SVC','FL Sales','TL SVC','TL Sales']
data2M=pd.DataFrame()
data=data.T

data2M=data[data['Date'].str.contains(Data2M)]
data2M=data2M.T
data2M.columns=data2M.loc['Date'].str[8:10]
data2M=data2M.drop(['Date'],axis=0)
data2M.index=['FL_SVC_'+date2M_name,'FL_Sales_'+date2M_name,'TL_SVC_'+date2M_name,'TL_Sales_'+date2M_name]

data1M=data[data['Date'].str.contains(Data1M)]
data1M=data1M.T
data1M.columns=data1M.loc['Date'].str[8:10]
data1M=data1M.drop(['Date'],axis=0)
data1M.index=['FL_SVC_'+date1M_name,'FL_Sales_'+date1M_name,'TL_SVC_'+date1M_name,'TL_Sales_'+date1M_name]


data0M=data[data['Date'].str.contains(Data0M)]
data0M=data0M.T
data0M.columns=data0M.loc['Date'].str[8:10]
data0M=data0M.drop(['Date'],axis=0)
data0M.index=['FL_SVC_'+date0M_name,'FL_Sales_'+date0M_name,'TL_SVC_'+date0M_name,'TL_Sales_'+date0M_name]

dataResult=data0M.sum(axis=1)
dataResult=dataResult.astype(int)

FSVC_Result=dataResult.loc['FL_SVC_'+date0M_name]
FSales_Result=dataResult.loc['FL_Sales_'+date0M_name]
TSVC_Result=dataResult.loc['TL_SVC_'+date0M_name]
TSales_Result=dataResult.loc['TL_Sales_'+date0M_name]

data2M=data2M.T
data2M.index=data2M.index.astype(int)
data2M=data2M.cumsum()
data1M=data1M.T
data1M.index=data1M.index.astype(int)
data1M=data1M.cumsum()
data0M=data0M.T
data0M.index=data0M.index.astype(int)
data0M=data0M.cumsum()


# FL TL Expected SVC data 추출
FSVC_Exp=int(round(FSVC_Result*total_businessday/pass_businessday,0))
TSVC_Exp=int(round(TSVC_Result*total_businessday/pass_businessday,0))
FSales_Exp=int(round(FSales_Result*total_businessday/pass_businessday,0))
TSales_Exp=int(round(TSales_Result*total_businessday/pass_businessday,0))


# Target trend 만들기
end_date=endM0.strftime('%d')
end_date=int(end_date)

FL_SVC_Target_Trend=np.arange(start = fsvc_target/end_date, stop = fsvc_target+fsvc_target/end_date, step = fsvc_target/end_date)
FL_SVC_Target_Trend=FL_SVC_Target_Trend.astype(int)

fsvc_Ctarget=928
FL_SVC_CTarget_Trend=np.arange(start = fsvc_Ctarget/end_date, stop = fsvc_Ctarget, step = fsvc_Ctarget/end_date)
FL_SVC_CTarget_Trend=FL_SVC_CTarget_Trend.astype(int)


FL_Sales_Target_Trend=np.arange(start = fsales_target/end_date, stop = fsales_target+fsales_target/end_date, step = fsales_target/end_date)
FL_Sales_Target_Trend=FL_Sales_Target_Trend.astype(int)

TL_SVC_Target_Trend=np.arange(start = tsvc_target/end_date, stop = tsvc_target+tsvc_target/end_date, step = tsvc_target/end_date)
TL_SVC_Target_Trend=TL_SVC_Target_Trend.astype(int)

TL_SVC_CTarget_Trend=np.arange(start = tsvc_Ctarget/end_date, stop = tsvc_Ctarget, step = tsvc_Ctarget/end_date)
TL_SVC_CTarget_Trend=TL_SVC_CTarget_Trend.astype(int)

TL_Sales_Target_Trend=np.arange(start = tsales_target/end_date, stop = tsales_target+tsales_target/end_date, step = tsales_target/end_date)
TL_Sales_Target_Trend=TL_Sales_Target_Trend.astype(int)



# Target data 추출
today=date.today()
yesterday=int(today.strftime('%d'))-1


FSVC_Target= int(round(fsvc_target*yesterday/end_date,0))   
FSVC_CTarget= int(round(fsvc_Ctarget*yesterday/end_date,0))
FSales_Target=int(round(fsales_target*yesterday/end_date,0))

TSVC_Target=int(round(tsvc_target*yesterday/end_date,0))
TSVC_CTarget=int(round(tsvc_Ctarget*yesterday/end_date,0))
TSales_Target=int(round(tsales_target*yesterday/end_date,0))

print("pgm check")
print("FSVC_Target: "+str(FSVC_Target))
print("FSVC_CTarget: "+str(FSVC_CTarget))
print("FSales_Target: "+str(FSales_Target))

print("TSVC_Target: "+str(TSVC_Target))
print("TSVC_CTarget: "+str(TSVC_CTarget))
print("TSales_Target: "+str(TSales_Target))

####  STATUS ####
# FL SVC Status data 추출
FSVC_Diff=int(round(FSVC_Result-FSVC_Target,0))
FSVC_DPer=int(round(FSVC_Diff*100/FSVC_Target,0))

if FSVC_Result-FSVC_Target>=0:
    FSVC_Status="Base: "+str(FSVC_Diff)+"EA ("+str(FSVC_DPer)+"%↑)"+" Challenge: "+str(FSVC_Diff)+"EA ("+str(FSVC_DPer)+"%↑)"
elif FSVC_Result-FSVC_Target<0:
    FSVC_Status=str(FSVC_Diff)+"EA ("+str(FSVC_DPer)+"%↓)"
else:
    print("Error")
print(FSVC_Status)

# FL SVC C Status data 추출
FSVC_CDiff=int(round(FSVC_Result-FSVC_CTarget,0))
FSVC_CDPer=int(round(FSVC_CDiff*100/FSVC_CTarget,0))
print("DD")
print(FSVC_CDiff)
print(FSVC_CDPer)

if FSVC_Result-FSVC_CTarget>=0:
    FSVC_CStatus=str(FSVC_CDiff)+"EA ("+str(FSVC_CDPer)+"%↑)"
elif FSVC_Result-FSVC_CTarget<0:
    FSVC_CStatus=str(FSVC_CDiff)+"EA ("+str(FSVC_CDPer)+"%↓)"
else:
    print("Error")
print(FSVC_CStatus)

# FL Sales Status data 추출
FSales_Diff=int(round(FSales_Result-FSales_Target,0))
FSales_DPer=int(round(FSales_Diff*100/FSales_Target,0))
if FSales_Diff >= 0:
    FSales_Status=str(FSales_Diff)+"EA ("+str(FSales_DPer)+"%↑)"
elif FSales_Diff < 0:
    FSales_Status=str(FSales_Diff)+"EA ("+str(FSales_DPer)+"%↓)"
else:
    print("Error")
print(FSales_Status)

# TL SVC Status data 추출
TSVC_Diff=int(round(TSVC_Result-TSVC_Target,0))
TSVC_DPer=int(round(TSVC_Diff*100/TSVC_Target,0))
if TSVC_Result-TSVC_Target>=0:
    TSVC_Status=str(TSVC_Diff)+"EA ("+str(TSVC_DPer)+"%↑)"
elif TSVC_Result-TSVC_Target<0:
    TSVC_Status=str(TSVC_Diff)+"EA ("+str(TSVC_DPer)+"%↓)"
else:
    print("Error")
print(TSVC_Status)

# TL SVC C Status data 추출
TSVC_CDiff=int(round(TSVC_Result-TSVC_CTarget,0))
TSVC_CDPer=int(round(TSVC_CDiff*100/TSVC_CTarget,0))
if TSVC_Result-TSVC_CTarget>=0:
    TSVC_CStatus=str(TSVC_CDiff)+"EA ("+str(TSVC_CDPer)+"%↑)"
elif TSVC_Result-TSVC_CTarget<0:
    TSVC_CStatus=str(TSVC_CDiff)+"EA ("+str(TSVC_CDPer)+"%↓)"
else:
    print("Error")
print(TSVC_CStatus)

# TL Sales Status data 추출
TSales_Diff=int(round(TSales_Result-TSales_Target,0))
TSales_DPer=int(round(TSales_Diff*100/TSales_Target,0))
if TSales_Result-TSales_Target>=0:
    TSales_Status=str(TSales_Diff)+"EA ("+str(TSales_DPer)+"%↑)"
elif TSales_Result-TSales_Target<0:
    TSales_Status=str(TSales_Diff)+"EA ("+str(TSales_DPer)+"%↓)"
else:
    print("Error")
print(TSales_Status)


################################################################################ 테이블 ################################################
############## plot의 요소들을 하나로 묶기
fig, ax = plt.subplots()
fig.set_size_inches(9, 2)
ax.set_axis_off()

# SVC 생성
table_vals=[[FSVC_Target,FSVC_CTarget,FSVC_Result,FSVC_Status,FSVC_CStatus,FSVC_Exp],[TSVC_Target, TSVC_CTarget,TSVC_Result, TSVC_Status, TSVC_CStatus,TSVC_Exp]]
col_labels=['Base Target',"Challenge Target",'Result','Base T.Status','Challenge T.Status',"Expected Closing"]
row_labels=["Front Loader","Top Loader"]
SVC_table=ax.table(cellText=table_vals, rowLabels=row_labels, colLabels=col_labels, loc='center', cellLoc='center')
SVC_table.auto_set_font_size(False)
SVC_table.set_fontsize(10)
SVC_table.auto_set_column_width(col=list(range(len(col_labels)-1)))
ax.set_title('*Service Overview',pad=0.1,x=0)

#그림 저장
plt.tight_layout()
plt.savefig('fig1.png')

############## plot의 요소들을 하나로 묶기
fig, ax = plt.subplots()
fig.set_size_inches(6.5, 2)
ax.set_axis_off()

# Sales 생성
table_vals=[[FSales_Target,FSales_Result,FSales_Status,FSales_Exp],[TSales_Target,TSales_Result, TSales_Status, TSales_Exp]]
col_labels=['Target','Result','Status',"Expected Closing"]
row_labels=["Front Loader","Top Loader"]
SVC_table=ax.table(cellText=table_vals, rowLabels=row_labels, colLabels=col_labels, loc='center', cellLoc='center')
SVC_table.auto_set_font_size(False)
SVC_table.set_fontsize(10)
SVC_table.auto_set_column_width(col=list(range(len(col_labels)-1)))
ax.set_title('*Sales Overview',pad=0.1,x=0.1)

#그림 저장
plt.tight_layout()
plt.savefig('fig2.png')


##############################################그래프 만들기##########################################################3
############## plot의 요소들을 하나로 묶기
fig, ax = plt.subplots(2,2)
fig.set_size_inches(18, 7)
ax[0,1].set_axis_off()
ax[1,1].set_axis_off()

############# 그래프 그리기 위해 데이터 합치기
result=pd.concat([data2M,data1M,data0M],axis=1)
result.columns=['FL_SVC_'+date2M_name,'FL_Sales_'+date2M_name,
              'TL_SVC_'+date2M_name,'TL_Sales_'+date2M_name,
                'FL_SVC_'+date1M_name,'FL_Sales_'+date1M_name,
                'TL_SVC_'+date1M_name,'TL_Sales_'+date1M_name,
                'FL_SVC_'+date0M_name,'FL_Sales_'+date0M_name,
                'TL_SVC_'+date0M_name,'TL_Sales_'+date0M_name]


result=result.reset_index()
############# FL 그래프 그리기
ax00T = ax[0,0].twinx()
ax00T.set_ylabel('Sales',color='gray')

#FL 데이터 넣기
Sdata=pd.DataFrame(result[['FL_Sales_'+date2M_name,'FL_Sales_'+date1M_name,'FL_Sales_'+date0M_name]])
Sdata.columns=["'"+date2M_name+' Sales',"'"+date1M_name+' Sales',"'"+date0M_name+' Sales']
Sdata[["'"+date2M_name+' Sales',"'"+date1M_name+' Sales',"'"+date0M_name+' Sales']].plot(kind='bar',color=['#C0B8CD','#B8DAFD','#F3BA0A'],ax=ax00T) # Sales 

ax[0,0].plot(result[['FL_SVC_'+date2M_name]],linestyle='-', linewidth=1.0,color='#CBCFC9',label="'"+date2M_name+' SVC') # SVC
ax[0,0].plot(result[['FL_SVC_'+date1M_name]],linestyle='-', linewidth=1.0,color='black',label="'"+date1M_name+' SVC') # SVC
ax[0,0].plot(result[['FL_SVC_'+date0M_name]], linestyle='-', marker='o', linewidth=2.0,color='red',label="'"+date0M_name+' SVC') # SVC


ax[0,0].plot(FL_SVC_Target_Trend,linestyle='-', linewidth=1.0,color='green',label="'"+date0M_name+' Base Target') # SVC # Twinx 만들기 위함
ax[0,0].plot(FL_SVC_CTarget_Trend,linestyle='--', linewidth=1.0,color='green',label="'"+date0M_name+' Challenge Target') # SVC # Twinx 만들기 위함
ax[0,0].set_ylabel('SVC',color='gray')

#FL 그래프 UI
ax00T.set_title("Front Loader Service & Sales Status",fontsize=13)
ax[0,0].set_xlabel("Date",color='gray')
ax[0,0].set_xticklabels(result["Date"])

ax00T.set_ylim(0,180000)
ax[0,0].set_ylim(0,1800)
#ax[0,0].set_xlim(-0.5,30.5)
ax00T.set_xlim=ax[0,0].set_xlim
ax[0,0].legend(loc='upper left')
ax00T.legend(loc='upper left',bbox_to_anchor=(0.3,1))


############# TL 그래프 그리기
ax10T = ax[1,0].twinx()
ax10T.set_ylabel('Sales',color='gray')

#TL 데이터 넣기
SSdata=pd.DataFrame(result[['TL_Sales_'+date2M_name,'TL_Sales_'+date1M_name,'TL_Sales_'+date0M_name]])
SSdata.columns=["'"+date2M_name+' Sales',"'"+date1M_name+' Sales',"'"+date0M_name+' Sales']
SSdata[["'"+date2M_name+' Sales',"'"+date1M_name+' Sales',"'"+date0M_name+' Sales']].plot(kind='bar',color=['#C0B8CD','#B8DAFD','#F3BA0A'],ax=ax10T) # Sales 

ax[1,0].plot(result[['TL_SVC_'+date2M_name]],linestyle='-', linewidth=1.0,color='#CBCFC9',label="'"+date2M_name+' SVC') # SVC
ax[1,0].plot(result[['TL_SVC_'+date1M_name]],linestyle='-', linewidth=1.0,color='black',label="'"+date1M_name+' SVC') # SVC
ax[1,0].plot(result[['TL_SVC_'+date0M_name]], linestyle='-', marker='o', linewidth=2.0,color='red',label="'"+date0M_name+' SVC') # SVC

ax[1,0].plot(TL_SVC_Target_Trend,linestyle='-', linewidth=1.0,color='green',label="'"+date0M_name+' Base Target') # SVC # Twinx 만들기 위함
ax[1,0].plot(TL_SVC_CTarget_Trend,linestyle='--', linewidth=1.0,color='green',label="'"+date0M_name+' Challenge Target') # SVC # Twinx 만들기 위함
ax[1,0].set_ylabel('SVC',color='gray')

#TL 그래프 UI
ax10T.set_title("Top Loader Service & Sales Status",fontsize=13)
ax[1,0].set_xlabel("Date",color='gray')
ax10T.set_xticklabels(result.index,rotation=0)

ax10T.set_ylim(0,220000)
ax[1,0].set_ylim(0,2000)
ax[1,0].set_xticklabels(result["Date"])
ax10T.set_xlim=ax[1,0].set_xlim
ax[1,0].legend(loc='upper left')
ax10T.legend(loc='upper left',bbox_to_anchor=(0.3,1))





#그래프 간격 띄우기
plt.tight_layout()
plt.savefig('fig3.png')


#Body 꾸미기
text1='Dear All,\n\nThis is the report of targets and results about daily services and sales.\n* It is based on what happened yesterday.\n\n'
msg.attach(MIMEText(text1,'plain'))


#첨부 파일1
with open('fig1.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('fig1.png'))
msg.attach(image)

msg.attach(MIMEText('\n','plain'))
#첨부 파일2
with open('fig2.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('fig2.png'))
msg.attach(image)

msg.attach(MIMEText('\n\n','plain'))

#첨부 파일2
with open('fig3.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('fig3.png'))
msg.attach(image)


#첨부 파일3
with open('sign.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('sign.png'))
msg.attach(image)


#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")

