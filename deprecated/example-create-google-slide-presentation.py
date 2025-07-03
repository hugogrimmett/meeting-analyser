import os
import sys
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'
SCOPES = ['https://www.googleapis.com/auth/presentations']

def check_and_help_credentials():
    if os.path.exists(CREDENTIALS_FILE):
        return True
    print(f"\nERROR: '{CREDENTIALS_FILE}' not found in the current directory.\n")
    print("To use the Google Slides API, you need OAuth 2.0 credentials from Google Cloud.")
    print("\nHow to get them:")
    print("1. Go to https://console.cloud.google.com/apis/credentials")
    print("2. Make sure you're in your project (or create a new one).")
    print("3. Click 'Create Credentials' > 'OAuth client ID'.")
    print("4. Choose 'Desktop app' and name it.")
    print("5. Click 'Create', then 'Download JSON'.")
    print("6. Rename the downloaded file to 'credentials.json' and put it in this directory.")
    print("\nOnce you have 'credentials.json', re-run this script.")
    sys.exit(1)

def get_slides_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return build('slides', 'v1', credentials=creds)

def main():
    check_and_help_credentials()
    service = get_slides_service()
    # Create a new Google Slides presentation
    presentation = service.presentations().create(body={
        'title': 'Automated Meeting Analysis Report'
    }).execute()
    presentation_id = presentation.get('presentationId')
    url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
    print("\n✅ Successfully created Google Slides presentation!")
    print(f"Presentation ID: {presentation_id}")
    print(f"View/Edit here: {url}")

    # Example: Add a second blank slide
    requests = [
        {
            "createSlide": {
                "slideLayoutReference": {
                    "predefinedLayout": "BLANK"
                }
            }
        }
    ]
    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}).execute()
    print("\nA blank slide was added. You can now use the API to add text, images, etc.")

if __name__ == "__main__":
    main()
