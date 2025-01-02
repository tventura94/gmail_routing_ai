import os
import logging
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
from openai import OpenAI
import base64
import time
from datetime import datetime
import json
from dotenv import load_dotenv

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_processor.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
          'https://www.googleapis.com/auth/spreadsheets']

def get_google_credentials():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def extract_email_data(email_content):
    try:
        client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
        
        logger.info("Sending request to OpenAI")
        logger.debug(f"Email content: {email_content[:100]}...")
        
        prompt = f"""Extract the following information from this email:
        1. Email address
        2. City (if mentioned, otherwise extract from venue location)
        3. Venue name
        4. Requested dates (if multiple dates are given, include all)
        
        Email content:
        {email_content}
        
        Return the information in JSON format with keys: email, city, venue, dates
        Note: If there are multiple date ranges, combine them into a single string separated by ' & '"""
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",  # Changed from gpt-4o-mini which was incorrect
            messages=[
                {"role": "system", "content": "You are a helpful assistant that extracts specific information from emails."},
                {"role": "user", "content": prompt}
            ]
        )
        
        logger.info("Received response from OpenAI")
        logger.debug(f"OpenAI response: {response.choices[0].message.content}")
        
        parsed_response = json.loads(response.choices[0].message.content)
        return parsed_response
        
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse OpenAI response: {e}")
        logger.error(f"Raw response: {response.choices[0].message.content}")
        raise
    except Exception as e:
        logger.error(f"Error in extract_email_data: {str(e)}")
        raise

def update_spreadsheet(service, spreadsheet_id, values):
    try:
        # First, get all values to find the next empty row
        range_name = 'Hold Grid!A:F'
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        # Find the next empty row by checking email column (A)
        current_values = result.get('values', [])
        next_row = 1  # Start at 1 to account for header row
        
        for row in current_values:
            # Check if the email column (first column) has a value
            if row and row[0].strip():  # Check if first column (email) is non-empty
                next_row += 1
            else:
                break
        
        logger.info(f"Inserting data at row {next_row}")
        
        # Update specific range using the next empty row
        update_range = f'Hold Grid!A{next_row}:E{next_row}'
        body = {
            'values': values
        }
        
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=update_range,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        
        logger.info("Successfully updated spreadsheet")
        
    except Exception as e:
        logger.error(f"Error updating spreadsheet: {str(e)}")
        raise

def get_last_processed_id():
    try:
        with open('last_processed.txt', 'r') as f:
            return f.read().strip()
    except FileNotFoundError:
        return None

def save_last_processed_id(message_id):
    with open('last_processed.txt', 'w') as f:
        f.write(message_id)

def monitor_emails():
    try:
        creds = get_google_credentials()
        gmail_service = build('gmail', 'v1', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        
        SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
        logger.info(f"Starting email monitoring for spreadsheet: {SPREADSHEET_ID}")
        
        last_message_id = get_last_processed_id()
        logger.info(f"Last processed message ID: {last_message_id}")
        
        while True:
            try:
                results = gmail_service.users().messages().list(
                    userId='me',
                    labelIds=['SENT'],
                    maxResults=1
                ).execute()
                
                messages = results.get('messages', [])
                
                if messages and (last_message_id != messages[0]['id']):
                    last_message_id = messages[0]['id']
                    
                    logger.info("Processing new email")
                    
                    message = gmail_service.users().messages().get(
                        userId='me',
                        id=last_message_id,
                        format='full'
                    ).execute()
                    
                    # Get recipient email from headers
                    headers = message['payload']['headers']
                    to_email = next(
                        header['value'] for header in headers 
                        if header['name'].lower() == 'to'
                    )
                    logger.info(f"Email sent to: {to_email}")
                    
                    # Extract email body
                    if 'data' in message['payload']['body']:
                        email_content = base64.urlsafe_b64decode(
                            message['payload']['body']['data']
                        ).decode('utf-8')
                    else:
                        # Handle multipart messages
                        parts = message['payload']['parts']
                        email_content = base64.urlsafe_b64decode(
                            parts[0]['body']['data']
                        ).decode('utf-8')
                    
                    # Extract data using AI
                    extracted_data = extract_email_data(email_content)
                    
                    # Prepare data for spreadsheet (removed Date Confirmed)
                    row_data = [[
                        to_email,           # Email column (now using recipient's email)
                        extracted_data['city'],   # City column
                        extracted_data['venue'],  # Venue column
                        extracted_data['dates'],  # Dates requested column
                        'CONTACTED'               # Status column
                    ]]
                    
                    # Update spreadsheet
                    update_spreadsheet(sheets_service, SPREADSHEET_ID, row_data)
                    
                    logger.info("Successfully processed email")
                    
                    # Save the ID after successful processing
                    save_last_processed_id(last_message_id)
                    
                time.sleep(60)
                
            except Exception as e:
                logger.error(f"Error in monitor_emails loop: {str(e)}", exc_info=True)
                time.sleep(60)
                
    except Exception as e:
        logger.error(f"Fatal error in monitor_emails: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    logger.info("Starting application")
    try:
        monitor_emails()
    except Exception as e:
        logger.error("Application crashed", exc_info=True) 