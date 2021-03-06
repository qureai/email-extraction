# Automated Extraction from Emails
This repo contains the scripts and the procedure for outlook 365 users to extract all the Usernames and Email address from the CC, To , from in the inbox and To, cc,bcc from the sentbox in the MS-Outlook mailbox. 

These details shall be extracted in a .csv and .xlsx formats and separated in two files to differenciate email addresses internal and external to the organization. The files shall contain a column beside the username and email address having a parameter 'Y' or 'N' which says whether the person is only present in sent mails or in the inbox too respectively.


## User guide

### Giving authorizations to the api

The files ```giving_permissions.py``` file should be run only once. 
Run the following commands on the terminal of your machine.

These commands should be run only the first time.
```
python -m pip install -r requirements.txt

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
After that anytime we want to extract, the file *email_extract_data.py* will extract and store the unique usernames and emails in .xls and .csv format in *Name and Emails(all).xls* and *Name and Emails(all).csv*.

*Name and Emails(internal).csv* and *Name and Emails(internal).xls* contains the username and email-ids which are internal to the organization i.e. of the form @qure.ai.

*Name and Emails(external).csv* and *Name and Emails(external).xls* contains the username and email-ids which are external to the organization i.e.  **not** of the form @qure.ai.


```
python email_extract_data.py 'client_id' 'client_secret' 'tenant_id'
```
### Improvements that could be done

This can be transformed into a GUI interface instead of the current cli which is difficult to be used. 
The system can be improved by extracting messages from the various folders too along with the inbox and outbox.
The user should be given an option to select the folders from which he/she wants to extract messages.
