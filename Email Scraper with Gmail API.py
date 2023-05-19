# Import packages
import os
import re
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email import message_from_bytes
import base64

# Function to authenticate access to the email address
def authenticate():
    creds = None

    # If there are no (valid) credentials available
    if not creds or not creds.valid:
        # If credentials is expired and needs a refresh, generate one
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        # Else, thre's a valid token start the OAuth flow
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret_867551012153-4f1rgr1umdj1olo3aeqo879ltm8jgae6.apps.googleusercontent.com.json', ['https://www.googleapis.com/auth/gmail.readonly']) 
            #Valid URI for the flow
            creds = flow.run_local_server(port=8080)
       
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

        print('\n\nAPI Connected\n\n')
    return creds

# Function to retrieve the readable email content
def get_email_content(msg):
    # If email payload contains not only text
    if 'parts' in msg['payload']:
        parts = msg['payload']['parts']
        # Loop through each part of the email 
        for part in parts:
            # Check for plain text and get encoded data, replace for - and + for proper decoding
            if part['mimeType'] == 'text/plain':
                data = part['body']['data']
                data = data.replace('-', '+').replace('_', '/')
                content = base64.b64decode(data).decode('utf-8')
                return content
    # If email is plain text, get data directly and decode
    else:
        data = msg['payload']['body']['data']
        data = data.replace('-', '+').replace('_', '/')
        content = base64.b64decode(data).decode('utf-8')
        return content
    

# 1 - Read in 20 emails from your personal email and pull out any emails that are from "john@venturebnb.io"
# and include "Traveler Housing Request" in the subject line
def ReadInFurnishedFinderHousingRequestsEmails():
    # Request Gmail API to fetch first 20 emails in the inbox
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=20).execute()
    # Save these emails in messages
    messages = results.get('messages', [])
    # List to save matching emails
    emails = []

    count = 1

    # Loop through each item in messages, 
    for message in messages:
        # Get all content from the emails and save to msg
        msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
        # Access headers of each email for the Subject and From fields
        subject = next(h for h in msg['payload']['headers'] if h['name'] == 'Subject')['value']
        sender = next(h for h in msg['payload']['headers'] if h['name'] == 'From')['value']
        # Check if subject contains 'Traveler Housing Request' and if it is from 'john@venturebnb.io'
        # Use get_email_content function to extract readable content from the email and save it to ms
        if 'john@venturebnb.io' in sender and 'Traveler Housing Request' in subject:
            msg = get_email_content(msg)

            print('\nMatched Email#', count , ' decoded')
            emails.append(msg)
            print('Matched Email #', count , ' extracted')
            count += 1
    return emails


# 2 - Loop through the emails and put the following information from EACH email into a new row of a pandas dataframe:
# Tenant, Email Address, Phone Number, Number of Travelers, and Dates
def PullInformationFromEmailsAndPutIntoDataframe(emails):
    # Data store for dataframe later
    data = []

    print('\n\nCreating Dataframe...')
    # Loop through each matched email
    for email in emails:

       # Define regex to search the trequired fields and their values, include all characters in search
        tenant_match = re.search(r'Tenant:\s*(.*?)(?=Email:)', email, re.DOTALL)
        email_match = re.search(r'Email:\s*(.*?)(?=Phone #)', email, re.DOTALL)
        phone_match = re.search(r'Phone #:\s*(.*?)(?=Travelers:)', email, re.DOTALL)
        travelers_match = re.search(r'Travelers:\s*(.*?)(?=Dates:)', email, re.DOTALL)
        dates_match = re.search(r'Dates:\s*(.*?)(?=Traveling To:)', email, re.DOTALL)

        # If there are matches for the regex above, remove trailing whitespace characters and newlines
        # And save, otherwise, None
        tenant = re.sub(r'\s*\n$', '', tenant_match.group(1).strip()) if tenant_match else None
        email_address = re.sub(r'\s*\n$', '', email_match.group(1).strip()) if email_match else None
        phone_number = re.sub(r'\s*\n$', '', phone_match.group(1).strip()) if phone_match else None
        num_travelers = re.sub(r'\s*\n$', '', travelers_match.group(1).strip()) if travelers_match else None
        dates = re.sub(r'\s*\n$', '', dates_match.group(1).strip()) if dates_match else None


        # Append the extracted values to data 
        data.append([tenant, email_address, phone_number, num_travelers, dates])

    # Create a pandas DataFrame with data
    dataframe = pd.DataFrame(data, columns=['Tenant', 'Email Address', 'Phone Number', 'Number of Travelers', 'Dates'])
    
    print('Dataframe Created')
    return dataframe

# Main Method
if __name__ == '__main__':

    # Authenticate and create gmail service using the api
    creds = authenticate()
    service = build('gmail', 'v1', credentials=creds)

    # 1 - Read in 20 emails from your personal email and pull out any emails that are from "john@venturebnb.io"
    # and include "Traveler Housing Request" in the subject line
    emails = ReadInFurnishedFinderHousingRequestsEmails()

    # 2 - Loop through the emails and put the following information from EACH email into a new row of a pandas dataframe:
    # Tenant, Email Address, Phone Number, Number of Travelers, and Dates
    dataframe = PullInformationFromEmailsAndPutIntoDataframe(emails)

    # Export the dataframe into an excel file
    dataframe.to_excel('Case Study.xlsx', index=False)

    print('Dataframe exported to Excel')
