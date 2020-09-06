import win32com.client 
import datetime as dt
import pandas as pd
import glob
import re

myFiles = glob.glob('*.xlsx')
print(myFiles)
for file in myFiles:
    print(file)
    df=pd.read_excel(file,header=0,dtype=object)
day=int(input("Nunmber of days to be considered:"))
y = (dt.date.today() - dt.timedelta(days=day))
print(y)
y = y.strftime('%m/%d/%Y %H:%M %p')
print(y)

def isWordPresent(sentence, word):
    sentence = re.sub('[^a-zA-Z0-9]',' ',sentence)    
    s = sentence.split(" ")  
    for i in s: 
        if (i.strip() == str(word).strip()):
            return True
    return False
def isWordPresent1(sentence, word):
    sentence = re.sub('[^a-zA-Z0-9#]',' ',sentence)    
    s = sentence.split("#")  
    for i in s: 
        if (i.strip() == str(word).strip()):
            return True
    return False

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
namespace = outlook.Session
recipient = namespace.CreateRecipient("deesaw@deloitte.com")
inbox = outlook.GetDefaultFolder(6)#(recipient, 6)
messages = inbox.Items
messages = messages.Restrict("[ReceivedTime] >= '" + y +"'")
email_subject = []

i=0
for x in messages:
    i=i+1
    sub = x.Subject
#    if isWordPresent(sub,'Task'):
    email_subject.append(sub)
df['#start']=None
df['#issue']=None
df['#done']=None
i=0
for d in df['UID']:
    
    for e in email_subject:
        s=isWordPresent(e,d)
#        print(str(d)+":"+str(e)+":"+str(s)+":"+str(start)+":"+str(issue)+":"+str(done))
        df['There']=s
        if s is True:
            done=isWordPresent1(e.lower(),'done')
            issue=isWordPresent1(e.lower(),'issue')
            start=isWordPresent1(e.lower(),'start')
            df['#start'][i]=start
            df['#issue'][i]=issue
            df['#done'][i]=done
    i=i+1
df['STATUS']=df.apply(lambda x :  '#done' if (x['#done']) else('#issue' if x['#issue'] else('#start' if x['#start'] else 'Task Received')) ,axis=1)
#df=df.iloc[:,1:]
df.to_excel(file, sheet_name='ETL_Tracker', engine='xlsxwriter',index=False)




