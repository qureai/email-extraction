# Automated Extraction from Emails
This repo contains the scripts and the procedure for Qure users to extract all the Usernames and Email address from the CC, To , from in the inbox and To, cc,bcc from the sentbox in the MS-Outlook mailbox. 

These details shall be extracted in a .csv and .xlsx format and stored in the system. The files shall contain a column beside the username and email address having a parameter 'Y' or 'N' which says whether the person is only present in sent mails or in the inbox too respectively.


## User guide

### Giving authorizations to the api

The files ```giving_permissions.py``` file should be run only once. 
Run the following commands on the terminal of your machine.

These commands should be run only the first time.
```
pip install -r requirements.txt

python giving_permissions.py 'client_id' 'client_secret' 'tenant_id'
````
On running the above command the user will see a message like 
```
Visit the following url to give consent:
https://login.microsoftonline.com/...................
Paste the authenticated url here:

```

The user should do ```Ctrl+left click``` on the url to open the page ina new tab on their browser or copy and paste it  on a new tab on their browser and press enter . **The browser should be Google chrome.** In other browsers the link may not work as intended. After signing in through qure.ai mail and then granting the permissions;
After this copy the new url from the search bar and paste it in the terminal and press enter. 

If correctly done then this message shall be displayed 
```
Authentication Flow Completed. Oauth Access Token Stored. You can now use the API.
Authenticated!
```
If upto this correctly done the user can proceed below.





### Extraction of email usernames and addresses 
After that anytime we want to extract, the file *email_extract_data.py* will extract and store the unique usernames and emails in .xls and .csv format in *Name and Emails.xls* and *file_name.csv*.
```
python email_extract_data.py 'client_id' 'client_secret' 'tenant_id'
```

