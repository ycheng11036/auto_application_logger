import base64
import os 

from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import gspread
from email.utils import parsedate_to_datetime

import logging
import re
import json, openai 

SCOPES = ['https://www.googleapis.com/auth/gmail.modify',
          'https://www.googleapis.com/auth/spreadsheets']

MAX_RESULTS = 10

load_dotenv() 

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


client = openai.OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.getenv("OPEN_API_KEY")
)
## GMAIL AUTHENTICATION

def authenticate(): 
    creds = None
    ## check for current credentials
    logger.info("Authenticating User...")
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    ## if no valid credentials, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            ## refresh the current credential
            creds.refresh(Request())
        else: 
            ## run the oauth flow
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        ## save the credential for next time
        with open('token.json', 'w') as token: 
            token.write(creds.to_json())
    
    return creds

# --------------- GMAIL SERVICE ----------------#
def build_gmail(creds):
    service = build("gmail", "v1", credentials=creds)
    logger.info("✅ Successfully connected to Gmail!")
    return service


## GMAIL FECTHER FUNCTIONS
def fetch_recent_emails(service, max_results=MAX_RESULTS):
    ## try to get the most recent emails
    try: 
        logger.info(f"fetching {max_results} most recent emails...\n")

        ## get list of unread emails
        results = service.users().messages().list(userId='me', maxResults=max_results, labelIds=["INBOX", "UNREAD"]).execute() 

        messages = results.get('messages', [])

        if not messages: 
            print('No messages found')
            return
        
        return messages
    
    except Exception as error: 
        logger.error(f"❌ Error fetching emails: {error}")

def fetch_email_details(service, message_id): 
    ## Fetch and return details of a specific email
    try:
        message = service.users().messages().get(
            userId='me', 
            id=message_id,
            format='full'
        ).execute()

        headers = message['payload'].get('headers', [])

        email_data = {
            'subject': get_header(headers, 'subject'),
            'from': get_header(headers, 'from'),
            'date': get_header(headers, 'date'),
            'body': extract_body(message['payload'])
        }    
        return email_data
    except Exception as error: 
        logger.error(f"❌ Error getting email details: {error}")
        return None
    
def get_header(headers, name):
    """Extract a specific header value from the message headers."""
    for header in headers:
        if header['name'].lower() == name.lower():
            return header['value']
    return None

def extract_body(payload):
    """Extract plain text body from a message payload."""
    if 'parts' in payload:
        for part in payload['parts']:
            body = extract_body(part)
            if body:
                return body
    else: 
        mime_type = payload.get('mimeType', '')
        if mime_type == 'text/plain':
            data = payload['body'].get('data')
            if data:
                return decode_base64(data)
        elif mime_type == 'text/html':
            # fallback: extract html content (optional)
            data = payload['body'].get('data')
            if data:
                html = decode_base64(data)
                # You can strip html tags if you want plain text:
                text = re.sub('<[^<]+?>', '', html)
                return text
    return ""

def decode_base64(data):
    """Decode base64 email body and trim if too long."""
    text = base64.urlsafe_b64decode(data).decode('utf-8')
    return text

## Mark the Email as Read after fetching
def mark_email_as_read(service, message_id):
    try:
        service.users().messages().modify(
            userId='me',
            id=message_id,
            body={'removeLabelIds': ['UNREAD']}
        ).execute()
        logger.info(f"Marked email {message_id} as read.")
    except Exception as e:
        logger.error(f"Failed to mark email as read: {e}")


# ----------------------- Printing Helpers -------------------------#

def print_email(number, email_data): 
    ## """Print email information"""
    print(f"📬 Email #{number}")
    print(f"From: {email_data.get('from', 'Unknown')}")
    print(f"Subject: {email_data.get('subject', 'No Subject')}")
    print(f"Date: {email_data.get('date', 'Unknown')}")
    print(f"Preview: {email_data.get('body', 'No content')}")
    print()

# ----------------------- Google Sheets Functions ----------------------#

def log_data_to_sheets(sheet,messages):
    for subject, sender, date, snippet in messages:
        sheet.append_row([subject, sender, date, snippet])
    logger.info(f"Logged {len(messages)} emails to Google Sheets.")

# ----------------------- LLM Functions ---------------------------------#
def extract_job_info_llm(email_body):
    prompt = f"""
    You are analyzing an email to determine if it's related to a job application process.

    Job-related emails include:
    - Application confirmations
    - Interview invitations or scheduling
    - Rejection notifications
    - Offer letters
    - Application status updates
    - Requests for additional information
    - Any communication from a company's recruiting/HR team about a job application

    Email content:
    \"\"\"
    {email_body}
    \"\"\"

    Instructions:
    1. If this email is NOT related to a job application process, return exactly: {{"process": false}}. 
    Any promotions related email will also be classified as such. This includes emails such as we've found the right fit, your background might be a good fit and new match for jobs.  
    2. If this email IS related to a job application process, return valid JSON with these fields:
       - process: true
       - job_title: the position title mentioned (or "Not specified" if unclear)
       - company: the company name
       - applied_date: true for this field if this is an application notification
       - response_date: true for this field if this is a rejection or offer
       - status: one of "Applied", "Rejected", "Offered"

    Return only valid JSON, no additional text or formatting.
    """
    response = client.chat.completions.create(
        model="meta-llama/llama-3.3-70b-instruct:free",
        messages=[{"role":"user","content":prompt}],
        temperature=0
    )
    content = response.choices[0].message.content.strip()
    # print(f"LLM Response: {content}")  # Debug output
    # Handle markdown code blocks
    if "```" in content:
        json_match = re.search(r'```(?:json)?\s*({.*?})\s*```', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)

    try:
        return json.loads(content)
    except json.JSONDecodeError:
        return {"process": False}

def main(): 

    try: 
        creds = authenticate()
        gmail_service = build_gmail(creds)
        
        sheets_client = gspread.authorize(creds)
        sheet = sheets_client.open_by_key(SPREADSHEET_ID).sheet1
        headers = sheet.row_values(1)

        all_values = sheet.get_all_values()

        row_map = {}

        for idx, row in enumerate(all_values[1:], start=2):
            company = row[0].strip().lower()
            position = row[1].strip().lower()
            if company and position:
                row_map[(company,position)] = idx

        messages = fetch_recent_emails(gmail_service, max_results=MAX_RESULTS)

        if not messages: 
            logger.info("No Messages Found")
            return
          
        for i, msg in enumerate(messages, start=1):
            details = fetch_email_details(gmail_service, msg['id'])
            print(details)
            if not details:
                continue
            
            mark_email_as_read(gmail_service, msg['id'])
            llm_result = extract_job_info_llm(details["body"])
            if not llm_result.get("process"):
                continue

            email_date_str = details["date"]
            email_date = parsedate_to_datetime(email_date_str).strftime("%m-%d-%Y")

            application_date = email_date if llm_result.get("applied_date") else ""
            response_date = email_date if llm_result.get("response_date") else ""

            company = llm_result.get("company", "")
            position = llm_result.get("job_title", "")
            status = llm_result.get("status", "")

            company_key = company.strip().lower()
            position_key = position.strip().lower()

            key = (company_key, position_key)
            
            if key in row_map:
                row_index = row_map[key]
                if response_date:
                    sheet.update_cell(row_index, headers.index("Response Date") + 1, response_date)
                if status:
                    sheet.update_cell(row_index, headers.index("Status") + 1, status)
                logger.info(f"Updated existing row {row_index} for {company} - {position}")
            else: ## no existing position and company ... add a new one
                row_to_append = [
                    company,
                    position,
                    status,
                    application_date,
                    response_date
                ]
                sheet.append_row(row_to_append)
                row_map[key] = len(all_values) + 1
                logger.info(f"Added new row for {company} - {position}")

    except Exception as error: 
        logger.error("Error trying to fetch, process, or write.", error)


if __name__ == "__main__":
    main()


