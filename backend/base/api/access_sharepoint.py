from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
from msal import ConfidentialClientApplication
import msal
import requests


####inputs########
# This will be the URL that points to your sharepoint site. 
# Make sure you change only the parts of the link that start with "Your"
url_shrpt = 'https://cittabasesolution-my.sharepoint.com/personal/it_helpdesk_cittabase_com'
site_url = 'https://cittabasesolution-my.sharepoint.com/personal/it_helpdesk_cittabase_com'
tenant_id = ''
client_id = ''
client_secret = ''
username_shrpt = ''
password_shrpt = ''
folder_url_shrpt = '/personal/it_helpdesk_cittabase_com/Documents/Shared_File_HRMS/Use%20For%20Local'
authority = f"https://login.microsoftonline.com/{tenant_id}"
#######################


# Authendicate with Auth Token
# Using azure.identity
from azure.identity import ClientSecretCredential
from azure.core.exceptions import AzureError

scope = ['https://graph.microsoft.com/.default']
# scope = ['https://graph.microsoft.com/Sites.Read.All']
sharepoint_site_url = 'https://cittabasesolution.sharepoint.com/sites/MySiteForDemo'
app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
credential = ClientSecretCredential(
       tenant_id=tenant_id,
       client_id=client_id,
       client_secret=client_secret,
   )

token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
# token_response = app.acquire_token_for_client(scopes=["https://cittabasesolution-my.sharepoint.com/.default"])

# Extract the access token
access_token = token_response['access_token']
print("access_token - 1", access_token)

try:
    # Get the access token
    # token = credential.get_token(*scope)
    # access_token = token.token
    # access_token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImgzamlfNnZ2Nm11ai1OQzg2Tlg1QW9Gay0tVWpzQWYzQ015bFRhSm9sbGciLCJhbGciOiJSUzI1NiIsIng1dCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyIsImtpZCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC80ODhkMGJhZS03NTVlLTRkMzAtOWZiOC1lMGM1NDk1NjU0YzQvIiwiaWF0IjoxNzM2NjE3ODczLCJuYmYiOjE3MzY2MTc4NzMsImV4cCI6MTczNjYyMTc3MywiYWlvIjoiazJCZ1lKaVVNTW4vMDllRW03S3lPcUtmdWt5MkFBQT0iLCJhcHBfZGlzcGxheW5hbWUiOiJUcmFrc2tvcmUiLCJhcHBpZCI6ImUyZWMyMmM2LWRkY2YtNDBhYi05MGM3LTUwOGIwYzczMGUwYyIsImFwcGlkYWNyIjoiMSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzQ4OGQwYmFlLTc1NWUtNGQzMC05ZmI4LWUwYzU0OTU2NTRjNC8iLCJpZHR5cCI6ImFwcCIsIm9pZCI6IjUxNDM3OWI2LTVlODItNGJmNi04NTkyLTI3YmJjOGM2ZWQ0MyIsInJoIjoiMS5BVlVBcmd1TlNGNTFNRTJmdU9ERlNWWlV4QU1BQUFBQUFBQUF3QUFBQUFBQUFBQ0lBQUJWQUEuIiwic3ViIjoiNTE0Mzc5YjYtNWU4Mi00YmY2LTg1OTItMjdiYmM4YzZlZDQzIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkFTIiwidGlkIjoiNDg4ZDBiYWUtNzU1ZS00ZDMwLTlmYjgtZTBjNTQ5NTY1NGM0IiwidXRpIjoiY3Jtb3FkbzRTVTZJcHU4MGtjWTdBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiMDk5N2ExZDAtMGQxZC00YWNiLWI0MDgtZDVjYTczMTIxZTkwIl0sInhtc19pZHJlbCI6IjI0IDciLCJ4bXNfdGNkdCI6MTYzOTk4MTE5MX0.ZmFyfBumigTcjy-6w3h1A4xt5-zpQ_jjv-E6SNnicstdrtd87NVaBQmPtorZjgYapvQO-1PSrbDjxulklyqtOVIA0l2dY2ZjaVKORf-MoJNu1eHc-z4Ah3AGs-weCqVoKBcpILjPgy3LF7gnVgJZOYwGMrjfGwsGLBBi9y9xICURSvCeHwvlkzbqGAU2pFxpxMc4vmvksKQVoew3kYCaJwkW7L3sleps4SYh6IO6Uzz881pzPYTxD5TnE1LWnFUdl4vBSV_-it-qX3WfJGmWxisiuFMpyQAI7SaY6wpfwysnrEESCln739szefDIrA8QnIzkQlz8pi4_SaLx8BDbWw"
    # print(credential.get_token_info(*scope).token)
    # access_token = credential.get_token_info(*scope).token
    # print(f"Access Token: {access_token}")
    
    # # Now, use the access token to make a request to the SharePoint REST API
    # headers = {
    #     'Authorization': f'Bearer {access_token}',
    #     'Accept': 'application/json'
    # }

    # # Make an API request to get the web properties
    # response = requests.get(f"{sharepoint_site_url}/_api/web", headers=headers)

    # if response.status_code == 200:
    #     print("Access Token is valid. Successfully connected to SharePoint.")
    #     print("Response:", response.json())
    # else:
    #     print(f"Error: {response.status_code}, {response.text}")
    
    ctx3 = ClientContext(sharepoint_site_url).with_access_token(access_token)
    # ctx3 = ClientContext(url_shrpt, auth_context=str(access_token))
    # print("ctx3", ctx3)
    # Get the Web object (site-level information)
    web3 = ctx3.web
    ctx3.load(web3)
    print("ctx3 loaded")
    # ctx3.execute_query()
    

except AzureError as e:
    print(f"Authentication failed: {str(e)}")



# Authendicate with Username / Password
ctx_auth = AuthenticationContext(url_shrpt)
if ctx_auth.acquire_token_for_user(username_shrpt, password_shrpt):
#   print("ctx_auth", ctx_auth.acquire_token_for_user(username_shrpt, password_shrpt))
  ctx = ClientContext(url_shrpt, ctx_auth)
  web = ctx.web
  ctx.load(web)
  print(f"user/pwd Web URL: {web.properties if ctx.web else 'No web loaded'}")
  ctx.execute_query()
  print('Authenticated using user/pwd into sharepoint as: ',web.properties['Title'])

else:
  print(ctx_auth.get_last_error())
  
  
# Get Toket using adal
import adal
import requests

# Azure AD App registration details
# sharepoint_site_url = 'https://yourtenantname.sharepoint.com/sites/yoursite'

# Azure AD authority URL
authority = f'https://login.microsoftonline.com/{tenant_id}'
msal_scope = ['https://cittabasesolution.sharepoint.com/.default']

app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)

# Get the access token
token_response = app.acquire_token_for_client(scopes=msal_scope)

print("token_response['access_token']", token_response['access_token'])

# Resource URL (SharePoint)
resource = 'https://cittabasesolution-my.sharepoint.com'

# Create AuthenticationContext
context = adal.AuthenticationContext(authority)

# Get the access token using client credentials
token_response = context.acquire_token_with_client_credentials(resource, client_id, client_secret)

# Extract the access token from the response
access_token = token_response['accessToken']

# print(f"Access Token Using Adal: {access_token}")

# Make an API request to SharePoint
headers = {
    'Authorization': f'Bearer {access_token}',
    'Accept': 'application/json'
}

# Send a GET request to check access (for example, check the site title)
response = requests.get(f"{sharepoint_site_url}/_api/web", headers=headers)

if response.status_code == 200:
    print("Access Token is valid. Successfully connected to SharePoint.")
    print("Response:", response.json())
else:
    print(f"Error: {response.status_code}, {response.text}")

  
  
  
  
global print_folder_contents
def print_folder_contents(ctx, folder_url):
    try:
        print("ctx", ctx, type(ctx))
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        print("folder---", folder)
        fold_names = []
        sub_folders = folder.files #Replace files with folders for getting list of folders
        print("sub_folders", sub_folders, type(sub_folders))
        ctx.load(sub_folders)
        ctx.execute_query()
        print("---execute Complete---")
        for s_folder in sub_folders:
            print("s_folder---", s_folder)
            fold_names.append(s_folder.properties["Name"])

        return fold_names

    except Exception as e:
        print('Problem printing out library contents: ', e)
        
        
# filelist_shrpt=print_folder_contents(ctx,folder_url_shrpt)
# print(filelist_shrpt)






# from azure.identity import ClientSecretCredential
# from azure.core.exceptions import AzureError

# scope = ['https://graph.microsoft.com/.default']
# sharepoint_site_url = 'https://cittabasesolution.sharepoint.com/sites/MySiteForDemo'
# app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
# credential = ClientSecretCredential(
#        tenant_id=tenant_id,
#        client_id=client_id,
#        client_secret=client_secret,
#    )
# # print("credential", credential)
# # token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
# # token_response = app.acquire_token_for_client(scopes=["https://cittabasesolution-my.sharepoint.com/.default"])

# # Extract the access token
# # access_token = token_response['access_token']

# try:
#     # Get the access token
#     token = credential.get_token(*scope)
#     access_token = token.token
#     print(f"Access Token: {access_token}")
    
#     # Now, use the access token to make a request to the SharePoint REST API
#     headers = {
#         'Authorization': f'Bearer {access_token}',
#         'Accept': 'application/json'
#     }

#     # Make an API request to get the web properties
#     response = requests.get(f"{sharepoint_site_url}/_api/web", headers=headers)

#     if response.status_code == 200:
#         print("Access Token is valid. Successfully connected to SharePoint.")
#         print("Response:", response.json())
#     else:
#         print(f"Error: {response.status_code}, {response.text}")
    
#     ctx3 = ClientContext(sharepoint_site_url).with_access_token(access_token)
#     print("ctx3", ctx3)
#     # Get the Web object (site-level information)
#     web3 = ctx3.web
#     ctx3.load(web3)
#     print("ctx3 loaded")
#     # ctx3.execute_query()
    

# except AzureError as e:
#     print(f"Authentication failed: {str(e)}")

# # Create the client context using the access token
# ctx = ClientContext(url_shrpt, auth_context=str(access_token))

# # Now you can interact with SharePoint
# try:
#     web = ctx.web
#     ctx.load(web)
#     print(f"ClientContext: {ctx}")
#     print(f"Web URL: {ctx.web.url if ctx.web else 'No web loaded'}")
#     # ctx.execute_query()  # This will execute the query

#     # If successful, print the web title
#     print(f"Web title: {web.properties['Title']}")
# except Exception as e:
#     print("Error during execute_query: %s", e)
#     # logging.error("Error during execute_query: %s", e)
#     # exit()

# # print(f"Web title: {web.properties['Title']}")


# # app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
# # token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
# # # print("token_response", token_response)

# # # Extract the access token
# # access_token = token_response['access_token']

# # # # Create the client context using the access token
# # # ctx1 = ClientContext(url_shrpt, auth_context=access_token)
# # ctx1 = ClientContext(url_shrpt).with_access_token(access_token)
# # web1 = ctx1.web
# # print("ctx.web1", type(web1), ctx1)
# # ctx1.load(web1)
# # # ctx1.execute_query()

# # # print(f"Web title: {web1.properties['Title']}")

# ###Authentication###For authenticating into your sharepoint site###
# ctx_auth = AuthenticationContext(url_shrpt)
# if ctx_auth.acquire_token_for_user(username_shrpt, password_shrpt):
#   ctx = ClientContext(url_shrpt, ctx_auth)
#   print("user/pwd ctx", ctx)
#   web = ctx.web
#   ctx.load(web)
#   print(f"user/pwd Web URL: {web.properties if ctx.web else 'No web loaded'}")
#   ctx.execute_query()
#   print('Authenticated using user/pwd into sharepoint as: ',web.properties['Title'])

# else:
#   print(ctx_auth.get_last_error())
# ############################
  
  
  
  
# ####Function for extracting the file names of a folder in sharepoint###
# ###If you want to extract the folder names instead of file names, you have to change "sub_folders = folder.files" to "sub_folders = folder.folders" in the below function
# global print_folder_contents
# def print_folder_contents(ctx, folder_url):
#     try:
#         print("ctx", ctx, type(ctx))
#         folder = ctx.web.get_folder_by_server_relative_url(folder_url)
#         print("folder---", folder)
#         fold_names = []
#         sub_folders = folder.files #Replace files with folders for getting list of folders
#         print("sub_folders", sub_folders, type(sub_folders))
#         ctx.load(sub_folders)
#         ctx.execute_query()
#         print("---execute Complete---")
#         for s_folder in sub_folders:
#             print("s_folder---", s_folder)
#             fold_names.append(s_folder.properties["Name"])

#         return fold_names

#     except Exception as e:
#         print('Problem printing out library contents: ', e)
# ######################################################
  
  
# # # Call the function by giving your folder URL as input  
# # filelist_shrpt=print_folder_contents(ctx,folder_url_shrpt) 

# # # #Print the list of files present in the folder
# # print(filelist_shrpt)