from O365 import Account
import pandas as pd
import argparse
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
if account.authenticate(scopes=['basic', 'message_all']):
	print('Authenticated!')#Let's you know if the authentication was successful

