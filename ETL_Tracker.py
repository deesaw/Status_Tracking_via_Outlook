import win32com.client 
import datetime as dt
import pandas as pd
import glob
import re

def searchword(e,d):
     if re.search(str(d), e.lower()):
         if(str(d)=='8807'):
             print(str(d)+":"+str(e))
         return True
     return False
def searchword1(e,d):
     d=d+'$'
     if re.search(d, e.lower()):
         return True
     return False

myFiles = glob.glob('*.xlsx')
for file in myFiles:
    print(file)
    df=pd.read_excel(file,header=0,dtype=object)
day=int(input("Number of days to be considered:"))
y = (dt.date.today() - dt.timedelta(days=day))
print(y)
y = y.strftime('%m/%d/%Y %H:%M %p')
print(y)

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
namespace = outlook.Session
recipient = namespace.CreateRecipient("deesaw@deloitte.com")
inbox = outlook.GetDefaultFolder(6)#(recipient, 6)
messages = inbox.Items
messages = messages.Restrict("[ReceivedTime] >= '" + y +"'")
email_subject = []

for x in messages:
    sub = x.Subject
    PCOOO='Production Cutover Green-Light Task UID'
    if PCOOO in sub:
        email_subject.append(sub)
        
df['#start']=None
df['#issue']=None
df['#done']=None
df['There']=None

for d in df['UID']:
    print(d)
    for e in email_subject:
        s=searchword(e,d) 
#        print(str(d)+":"+str(e)+":"+str(s)+":"+str(start)+":"+str(issue)+":"+str(done))
        df.loc[df['UID']==d,'There']=s          
        if s is True:
            done=searchword1(e.lower(),'#done')#re.search('#done$', e.lower())
            issue=searchword1(e.lower(),'#issue')
            start=searchword1(e.lower(),'#start')
            df.loc[df['UID']==d,'#start']=start
            df.loc[df['UID']==d,'#issue']=issue
            df.loc[df['UID']==d,'#done']=done

df['Tag Status']=df.apply(lambda x :  '#done' if (x['#done']) else('#issue' if x['#issue'] else('#start' if x['#start'] else ('Received' if x['There'] else 'Yet to receive'))) ,axis=1)
#df=df.iloc[:,1:]
df.to_excel(file, sheet_name='ETL_Tracker', engine='xlsxwriter',index=False)




