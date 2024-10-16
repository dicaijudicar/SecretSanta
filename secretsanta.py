

import random
import pandas as pd
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
import numpy as np


        

df2=pd.read_excel(r"C:\gitclone\SecretSanta\emails.xlsx", engine='openpyxl')


list_ind=list(df2['Name'].unique())
df2['SecretSanta']=''

z=[]
[z.append(i) for i in range(0,len(list_ind))]

        

def shuffleAttach(df):
    random.shuffle(z)
    for row,rowS in df2.iterrows():
        df2.loc[row,'SecretSanta']=list_ind[z[row]]
    return df2


twoTruth=True
oneTruth=True
threeTruth=False
while ((oneTruth==True) or (twoTruth==True) or (threeTruth==True)):
    df2=shuffleAttach(df2)
    df2['testFilter1']=np.where(df2.Name==df2.SecretSanta,1,0)
    oneTruth= True if df2['testFilter1'].sum() > 0 else False
    df2['testFilter2']=np.where(df2.last_year==df2.SecretSanta,1,0)
    twoTruth= True if df2['testFilter2'].sum()>0 else False
    df2['testFilter3']=np.where((df2.SecretSanta.isin(['Rene','Carole']) & (df2.Name.isin(['Rene','Carole']))),1,0)
    threeTruth= True if df2['testFilter3'].sum()>0 else False
    
    



#loop over emails
for name in df2['Name'].unique():     
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    
    secretsanta=df2[df2.Name==name].iloc[0]['SecretSanta']
    
    message = f"Hello {name}. Are you ready for this? *Cue music* You are the Secret Santa for {secretsanta}. Reminder: Our budget is $50!"
    msg = MIMEMultipart()
    password = 'blork'
    msg['From'] = "juliendicaire@gmail.com"
    msg['To'] = df2[df2.Name==name].iloc[0]['Email']  
    msg['Subject'] = f"This message is for {name}. If you are not {name}, please disregard! Your Secret Santa Person awaits you!"
    msg.attach(MIMEText(message, 'plain'))
    server = smtplib.SMTP('smtp.gmail.com: 587')
    server.starttls()
    server.login(msg['From'], password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    

    
#export to excel the details 
df2.to_excel(r"C:\gitclone\SecretSanta\emails.xlsx", index=False)
