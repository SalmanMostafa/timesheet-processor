import os
import json
import msal
import requests
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")

AUTHORITY = "https://login.microsoftonline.com/consumers"

##For a personal Outlook mailbox, Azure Function cannot use its Managed Identity to impersonate your personal mailbox.
##Your Function needs to reuse a delegated user authorization, typically by obtaining an access token through a refresh-token-capable flow.
##Microsoft documents that requesting offline_access allows a refresh token to be issued so the application can obtain new access tokens
##without requiring you to sign in for every Graph call.
##https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow?utm_source=chatgpt.com

SCOPES = ["Mail.Read"]

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    raise Exception(f"Failed to create device flow: {flow}")

print(flow["message"])

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    print("Authentication failed:")
    print(result)
    raise SystemExit(1)

print("Authentication successful!")

headers = {
    "Authorization": f"Bearer {result['access_token']}"
}


url = "https://graph.microsoft.com/v1.0/me/messages?$top=10"

messages = requests.get(url, headers=headers).json()["value"]

unique_contentBytes = set()

for message in messages:
    print("Email:",message["subject"])
    print("Email lastmodified:",message["lastModifiedDateTime"])

    #attachments
    url = f"https://graph.microsoft.com/v1.0/me/messages/{message['id']}/attachments"
    attachments = requests.get(url, headers=headers).json()["value"]

    for attachment in attachments:
        content_bytes = attachment.get("contentBytes")
        if content_bytes:
            
            if content_bytes in unique_contentBytes:     # get the unique content bytes only in case there are replies or forwarded emails.
                continue
            unique_contentBytes.add(content_bytes)
            if attachment.get("isInline"):
                print("Pasted Image: ", attachment['name'])   # Inline or pasted images in the message body
            else:
                print("Attachment : ", attachment['name'])
            print(content_bytes)                      # output the unique content bytes for testing put content_bytes[-20:]
                

