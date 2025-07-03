import os
import sys
import pickle
import io
import re
import datetime
from collections import Counter, defaultdict
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseDownload
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np

# ==== SETUP ====
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/calendar.readonly'
]

def check_and_help_credentials():
    if os.path.exists(CREDENTIALS_FILE):
        return True
    print(f"\nERROR: '{CREDENTIALS_FILE}' not found in the current directory.\n")
    print("To use this script, you need OAuth 2.0 credentials from Google Cloud.")
    print("\nHow to get them:")
    print("1. Go to https://console.cloud.google.com/apis/credentials")
    print("2. Make sure you're in your project (or create a new one).")
    print("3. Click 'Create Credentials' > 'OAuth client ID'.")
    print("4. Choose 'Desktop app' and name it.")
    print("5. Click 'Create', then 'Download JSON'.")
    print("6. Rename the downloaded file to 'credentials.json' and put it in this directory.")
    print("\nOnce you have 'credentials.json', re-run this script.")
    sys.exit(1)

def get_google_services():
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
    calendar_service = build('calendar', 'v3', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    slides_service = build('slides', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    return calendar_service, drive_service, slides_service, sheets_service

# ==== FIND CALENDAR EVENTS WITH GEMINI MEETING NOTES ====

def find_meetings_with_gemini_notes(calendar_service):
    events = []
    now = datetime.datetime.utcnow().isoformat() + 'Z'
    print('Getting upcoming 50 events...')
    events_result = calendar_service.events().list(
        calendarId='primary', timeMin=now, maxResults=50, singleEvents=True,
        orderBy='startTime').execute()
    for event in events_result.get('items', []):
        attachments = event.get('attachments', [])
        for att in attachments:
            if 'gemini' in att.get('title', '').lower():
                events.append({'event': event, 'attachment': att})
    return events

# ==== GET TRANSCRIPT CONTENT ====
def get_transcript_from_gemini_drive_file(drive_service, sheets_service, file_id):
    # Try as Google Sheets
    file_metadata = drive_service.files().get(fileId=file_id, fields="mimeType, name").execute()
    if file_metadata['mimeType'] == 'application/vnd.google-apps.spreadsheet':
        # Get sheet/tab names
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=file_id).execute()
        tabs = [s['properties']['title'] for s in sheet_metadata['sheets']]
        if 'Transcript' in tabs:
            # Read all rows from the "Transcript" tab
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=file_id, range='Transcript'
            ).execute()
            rows = result.get('values', [])
            if rows:
                return "\n".join([",".join(row) for row in rows])
    # TODO: Support docs/text if needed
    return None

# ==== ANALYSE WORDS SPOKEN ====
def analyze_transcript_and_generate_images(transcript_text, temp_dir, baseprefix):
    # Very basic parse for CSV-like content, can be replaced with your own robust logic
    lines = [l for l in transcript_text.split('\n') if l.strip()]
    speaker_turns = []
    for line in lines:
        # e.g. "Matthew Bostwick: All right."
        match = re.match(r"^([A-Za-z .'-]+):\s*(.*)$", line)
        if match:
            speaker_turns.append((match.group(1).strip(), match.group(2).strip()))
        else:
            # (Optional) append to previous
            if speaker_turns:
                speaker_turns[-1] = (speaker_turns[-1][0], speaker_turns[-1][1] + ' ' + line.strip())
    if not speaker_turns:
        print("No speaker turns found, skipping.")
        return []
    # Analysis (words spoken)
    def count_words(text):
        return len(re.findall(r'\b\w+\b', text))
    word_counts = Counter()
    cumulative_word_counts = defaultdict(list)
    running_totals = Counter()
    for speaker, utterance in speaker_turns:
        words = count_words(utterance)
        word_counts[speaker] += words
        running_totals[speaker] += words
        # For cumulative plot, snapshot all participant totals at each turn
        for p in word_counts:
            cumulative_word_counts[p].append(running_totals[p])
    participants = list(word_counts.keys())
    images = []

    # Bar Chart
    plt.figure(figsize=(8, 5))
    sorted_word_counts = dict(word_counts.most_common())
    plt.bar(sorted_word_counts.keys(), sorted_word_counts.values())
    plt.title("Total Words Spoken")
    plt.xlabel("Participant")
    plt.ylabel("Number of Words Spoken")
    plt.xticks(rotation=45)
    plt.tight_layout()
    fname = os.path.join(temp_dir, f"{baseprefix}_bar_chart.png")
    plt.savefig(fname, dpi=200)
    plt.close()
    images.append(fname)

    # Cumulative plot
    plt.figure(figsize=(10, 5))
    for name in participants:
        plt.plot(cumulative_word_counts[name], label=name)
    plt.title("Cumulative Words Spoken")
    plt.xlabel("Turn Number")
    plt.ylabel("Cumulative Words Spoken")
    if participants:
        plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
    plt.tight_layout()
    fname = os.path.join(temp_dir, f"{baseprefix}_cumulative.png")
    plt.savefig(fname, dpi=200)
    plt.close()
    images.append(fname)
    # (Optional: Network graph, omitted for brevity. See previous code)
    return images

# ==== UPLOAD IMAGE TO SLIDES ====
def insert_images_to_slide(slides_service, presentation_id, image_paths):
    # Add blank slide
    create_slide_req = [{
        "createSlide": {
            "slideLayoutReference": {
                "predefinedLayout": "BLANK"
            }
        }
    }]
    slide_resp = slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": create_slide_req}).execute()
    new_slide_id = slide_resp['replies'][0]['createSlide']['objectId']
    # Insert images (simple vertical stack)
    requests = []
    y = 50
    for img_path in image_paths:
        # First upload to Drive or use a public URL for images if needed
        with open(img_path, "rb") as imgfile:
            imgdata = imgfile.read()
        image_url = f"file://{os.path.abspath(img_path)}"  # For local preview only; to upload to Slides, need public URL or Drive image
        # Instead: Use the "createImage" request with image URL
        requests.append({
            "createImage": {
                "url": image_url,  # This won't work for local files; you need to upload to Drive and use a shareable link.
                "elementProperties": {
                    "pageObjectId": new_slide_id,
                    "size": {
                        "height": {"magnitude": 250, "unit": "PT"},
                        "width": {"magnitude": 400, "unit": "PT"},
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1, "translateX": 50, "translateY": y, "unit": "PT"
                    }
                }
            }
        })
        y += 270
    # Note: To insert images programmatically, you must upload images to Google Drive, set share permissions, and use the Drive URL. This is a known limitation.
    # So for a real script, upload the image to Drive, get its shareable link, and use that link in the createImage request.
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}).execute()
    print("Inserted images into slide", new_slide_id)

def main():
    check_and_help_credentials()
    calendar_service, drive_service, slides_service, sheets_service = get_google_services()

    # Create the presentation
    presentation = slides_service.presentations().create(body={
        'title': 'Automated Meeting Analysis Summary'
    }).execute()
    presentation_id = presentation.get('presentationId')
    url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
    print(f"\nCreated Slides: {url}")

    # TEMP DIR for images
    import tempfile
    with tempfile.TemporaryDirectory() as temp_dir:
        # 1. Find meetings with Gemini notes
        events = find_meetings_with_gemini_notes(calendar_service)
        if not events:
            print("No events with Gemini notes attachments found.")
            return
        for evt in events:
            event = evt['event']
            att = evt['attachment']
            print(f"Processing event '{event.get('summary')}', file: {att.get('title')}")
            file_id = att.get('fileId')
            if not file_id:
                print("Attachment missing fileId, skipping.")
                continue
            transcript_text = get_transcript_from_gemini_drive_file(drive_service, sheets_service, file_id)
            if not transcript_text:
                print(f"No transcript found in {att.get('title')}, skipping.")
                continue
            # 2. Analyse and create images
            baseprefix = re.sub(r'\W+', '_', event.get('summary', 'meeting')).lower()
            images = analyze_transcript_and_generate_images(transcript_text, temp_dir, baseprefix)
            # 3. Insert images (upload to Drive and use a shareable URL)
            if images:
                print(f"Inserting images for {event.get('summary')}")
                # You need to upload each image to Drive, get a shareable link, and use that link in the createImage request.
                # This part is complex due to Google Slides API limitations (local files not allowed).
                # See https://developers.google.com/slides/api/guides/presentations-images#image-urls
                print(f"Please upload images manually or implement Drive upload + permission sharing for full automation.")
                # insert_images_to_slide(slides_service, presentation_id, images)

    print(f"\nAll done. View your slides at: {url}")

if __name__ == "__main__":
    main()
