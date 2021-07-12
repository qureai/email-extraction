from O365 import Account
import pandas as pd
import argparse
import time
parser = argparse.ArgumentParser()
parser.add_argument("client_id", type=str,
                    help="Enter client-id here")
parser.add_argument("client_secret", type=str,
                    help="Enter client-secret here")
parser.add_argument("tenant_id", type=str,
                    help="Enter tenant-id here")

args = parser.parse_args()

client_id = args.client_id
client_secret = args.client_secret
tenant_id = args.tenant_id
#print(type(client_id),client_secret,tenant_id)
credentials = ( client_id ,client_secret )#These are the values you copied earlier
# Directory (tenant) ID copied earlier

account = Account(credentials, tenant_id = tenant_id) 


data=[] #this will store all usernames and email addresses
data_sent=[]
mailbox = account.mailbox()
x=0
y=0

print("\n \n \n ----------------Inbox Mail details----------------\n")


inbox = mailbox.inbox_folder()

seconds = time.time()
for message in inbox.get_messages(limit=100000000,batch=700):
    x+=1
    #print(message.subject)
    #print("To --",message.to)
    for recipient in message.to._recipients:
        data.append([str(recipient._name),str(recipient._address),'N'])

    #print("Cc --",message.cc)
    for recipient in message.cc._recipients:
        data.append([str(recipient._name),str(recipient._address),'N'])
    #print("Sender --")
    data.append([str(message.sender._name),str(message.sender._address),'N'])
print(time.time()-seconds)    
print(x)





print("\n \n \n ----------------Sent Mail details----------------\n")

sent_folder = mailbox.sent_folder()


for message in sent_folder.get_messages(limit=100000000,batch=700):
    y+=1
    #print(message.subject)
    #print("To --",message.to)
    
    for recipient in message.to._recipients:
        data.append([str(recipient._name),str(recipient._address),'Y'])

    #print("Cc --",message.cc)    
    for recipient in message.cc._recipients:
        data.append([str(recipient._name),str(recipient._address),'Y'])

    #print("Bcc --",message.bcc)
    for recipient in message.bcc._recipients:
        data.append([str(recipient._name),str(recipient._address),'Y'])
print(y)  



df = pd.DataFrame(data, columns = ['Username', 'Email','Only Sent']) #creating a pandas dataframe
df = df.drop_duplicates(subset=['Username','Email'],keep = 'first') #removing duplicate rows
df.to_csv('file_name.csv', index=False)
df.to_excel('Name and Emails.xls','Sheet1',index=False)
'''
for i in range(6000): 
    m = account.new_message() #creates a new mail draft
    m.to.add('arka.das@qure.ai') #takes the email id from our file
    m.subject = '.' #subject of the mail
    m.body = "." 
    m.send() # sending the mail and the loop repeats for all the selected customers #do uncomment the line for sending
'''
