
# =============================================================================
# This is an alteration of a file I made last year to 
# email the secret santa of my family members anonymously
# =============================================================================

import random
import pandas as pd
import sys


L=['1','Julien','2','3','4','5','6']
L
df=pd.DataFrame(L)
df=df.rename(columns={0:"Name"})
df['SecretSanta']=0

z=[0,1,2,3,4,5,6]
random.shuffle(z)
for i in range(0,7):
    df['SecretSanta'].loc[i]=L[z[i]]


#testing
for i in range(0,7):
    if df.Name.loc[i]==df['SecretSanta'].loc[i]:
        sys.exit("Name match")
    else:
        print('All good')
        

df2=pd.read_excel("E:\\Programs\\Python\\emails.xlsx")
#df2=pd.DataFrame(emails).rename(columns={0:"Email"})
df3=pd.DataFrame(['Julien','etc']).rename(columns={0:"Name"})
df4=pd.concat([df2,df3],axis=1); del df2,df3
#creat finalized frame
final=pd.merge(df,df4, on='Name')

'''
#testing
for email in emails:
    dfred=final[final.Email==email]
    print('')
    print(dfred.iloc[0]['Name'])
    print(email)
    print(dfred.iloc[0]['SecretSanta'])
'''
ParentsWhoShareEmail='NameOfParents'
#loop over emails
for name in L:     
    dfred=final[final.Name==name]
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    
    name=dfred.iloc[0]['Name'];secretsanta=dfred.iloc[0]['SecretSanta']
    
    message = f"Hello {name}. I hope you've had a swell day. Mostly, I am making a long message so {ParentsWhoShareEmail} can't read eachother's people. You are the Secret Santa for {secretsanta}."
    msg = MIMEMultipart()
    password = "password" #insert password here 
    msg['From'] = "juliendicaire@gmail.com" #feel free to contact me with any questions
    msg['To'] = dfred.iloc[0]['Email']  
    msg['Subject'] = f"This message is for {name}. If you are not {name}, please disregard! Your Secret Santa Person!" # Type the subject of your message 
    msg.attach(MIMEText(message, 'plain'))
    server = smtplib.SMTP('smtp.gmail.com: 587')
    server.starttls()
    server.login(msg['From'], password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
