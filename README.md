## README
The files ```giving_permissions.py``` file should be run only once. 
On terminal 
run the following commands

These commands should be run only the first time
```
pip install xlwt
pip install O365
pip install pandas
python giving_permissions.py 'client_id' 'client_secret' 'tenant_id'
````

### Extraction of email usernames and addresses 
After that anytime we want to extract, the file *email_extract_data.py* will extract and store the unique usernames and emails in .xls and .csv format in *Name and Emails.xls* and *file_name.csv*.
```
python email_extract_data.py 'client_id' 'client_secret' 'tenant_id'
```

