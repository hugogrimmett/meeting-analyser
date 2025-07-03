import os
import sys
import pickle
import tempfile
import io
import re
import datetime
import argparse
import webbrowser
from collections import Counter, defaultdict
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload
import matplotlib.pyplot as plt

CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/calendar.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

def parse_date(s):
    try:
        return datetime.datetime.strptime(s, "%Y-%m-%d")
    except Exception as e:
        print(f"Error parsing date '{s}': {e}")
        sys.exit(1)

def get_date_range_from_args_or_prompt():
    parser = argparse.ArgumentParser(description='Analyze Gemini meeting notes in calendar events.')
    parser.add_argument('--start', type=str, default=None,
                        help="Start date (YYYY-MM-DD), default is today.")
    parser.add_argument('--end', type=str, default=None,
                        help="End date (YYYY-MM-DD), default is 7 days from start.")
    args = parser.parse_args()

    if args.start:
        start = parse_date(args.start)
    else:
        default_start = datetime.datetime.utcnow() - datetime.timedelta(days=7)
        start_str = input(f"Enter start date (YYYY-MM-DD, default {default_start.date()}): ").strip()
        start = parse_date(start_str) if start_str else default_start

    if args.end:
        end = parse_date(args.end)
    else:
        end_str = input("Enter end date (YYYY-MM-DD, default +7 days): ").strip()
        end = parse_date(end_str) if end_str else start + datetime.timedelta(days=7)

    if end <= start:
        print("End date must be after start date.")
        sys.exit(1)

    print(f"Searching events from {start.date()} to {end.date()}")
    return start, end

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

def find_meetings_with_gemini_notes(calendar_service, time_min, time_max):
    events = []
    print('Getting events...')
    events_result = calendar_service.events().list(
        calendarId='primary', timeMin=time_min, timeMax=time_max, maxResults=100, singleEvents=True,
        orderBy='startTime').execute()
    for event in events_result.get('items', []):
        print("DEBUG: Event summary:", event.get('summary'))
        print("DEBUG: Attachments:", event.get('attachments'))
        print("DEBUG: Description:", event.get('description'))
        attachments = event.get('attachments', [])
        for att in attachments:
            if 'gemini' in att.get('title', '').lower():
                events.append({'event': event, 'attachment': att})
    return events

def get_transcript_from_gemini_drive_file(drive_service, sheets_service, file_id):
    file_metadata = drive_service.files().get(fileId=file_id, fields="mimeType, name").execute()
    print("DEBUG: File metadata:", file_metadata)
    if file_metadata['mimeType'] == 'application/vnd.google-apps.spreadsheet':
        pass  # Could add future Sheet logic here
    elif file_metadata['mimeType'] == 'application/vnd.google-apps.document':
        try:
            doc_content = drive_service.files().export(fileId=file_id, mimeType='text/plain').execute()
            text = doc_content.decode('utf-8') if isinstance(doc_content, bytes) else doc_content
            print("DEBUG: Got Google Doc text, first 200 chars:", text[:200])
            return text
        except Exception as e:
            print("DEBUG: Error fetching Google Doc content:", e)
            return None
    else:
        print("DEBUG: File is not a Google Sheet or Doc. MIME type is:", file_metadata['mimeType'])
    return None

def analyze_transcript_and_generate_images(transcript_text, baseprefix, override_date=None):
    os.makedirs("analysis", exist_ok=True)
    lines = [l for l in transcript_text.split('\n') if l.strip()]
    date_ymd = override_date if override_date else datetime.datetime.now().strftime("%Y-%m-%d")
    meeting_title = baseprefix
    if lines and re.match(r'^[A-Za-z]{3} \d{1,2}, \d{4}', lines[0]):
        # ... keep existing date/title logic, but don't overwrite date_ymd if override_date was provided
        if not override_date:
            date_line = lines[0].strip()
            date_match = re.search(r'([A-Za-z]{3}) (\d{1,2}), (\d{4})', date_line)
            if date_match:
                month_str, day_str, year_str = date_match.groups()
                try:
                    date_obj = datetime.datetime.strptime(f"{month_str} {day_str} {year_str}", "%b %d %Y")
                    date_ymd = date_obj.strftime("%Y-%m-%d")
                except:
                    pass
        meeting_title = lines[1].strip() if len(lines) > 1 else baseprefix
        body_lines = lines[2:]
    else:
        body_lines = lines

    # Find transcript start, skip preamble
    start_index = next((i for i, line in enumerate(body_lines) if '00:00:00' in line), None)
    if start_index is not None:
        transcript_lines = body_lines[start_index:]
    else:
        transcript_lines = body_lines

    # Parse transcript into speaker turns
    speaker_turns = []
    current_speaker = None
    current_utterance = []

    for line in transcript_lines:
        speaker_match = re.match(r"^([A-Za-z .'-]+):(.*)$", line)
        if speaker_match:
            if current_speaker is not None:
                joined = ' '.join(current_utterance).strip()
                if joined:
                    speaker_turns.append((current_speaker, joined))
            current_speaker = speaker_match.group(1).strip()
            current_utterance = [speaker_match.group(2).strip()]
        else:
            if current_speaker is not None:
                current_utterance.append(line.strip())

    if current_speaker is not None and current_utterance:
        joined = ' '.join(current_utterance).strip()
        if joined:
            speaker_turns.append((current_speaker, joined))

    if not speaker_turns:
        print("No speaker turns found after 00:00:00. Skipping.")
        return [], date_ymd, meeting_title

    def count_words(text):
        return len(re.findall(r'\b\w+\b', text))

    word_counts = Counter()
    cumulative_word_counts = defaultdict(list)
    running_totals = Counter()

    for speaker, utterance in speaker_turns:
        words = count_words(utterance)
        word_counts[speaker] += words
        running_totals[speaker] += words
        for p in word_counts:
            cumulative_word_counts[p].append(running_totals[p])

    participants = list(word_counts.keys())
    images = []

    # Bar Chart: Total Words Spoken
    plt.figure(figsize=(10, 6))
    sorted_word_counts = dict(word_counts.most_common())
    plt.bar(sorted_word_counts.keys(), sorted_word_counts.values())
    plt.title(meeting_title + " – Total Words Spoken")
    plt.xlabel("Participant")
    plt.ylabel("Number of Words Spoken")
    plt.xticks(rotation=45)
    plt.tight_layout()
    fname1 = os.path.join("analysis", f"{date_ymd}_{baseprefix}_bar_chart_total_words_spoken.png")
    plt.savefig(fname1, dpi=300)
    plt.close()
    images.append(fname1)

    # Cumulative Word Count Plot
    plt.figure(figsize=(12, 6))
    for name in participants:
        plt.plot(cumulative_word_counts[name], label=name)
    plt.title(meeting_title + " – Cumulative Words Spoken")
    plt.xlabel("Turn Number")
    plt.ylabel("Cumulative Words Spoken")
    if participants:
        plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
    plt.tight_layout()
    fname2 = os.path.join("analysis", f"{date_ymd}_{baseprefix}_cumulative_words_spoken.png")
    plt.savefig(fname2, dpi=300)
    plt.close()
    images.append(fname2)

    print("Plots saved using basename:", f"{date_ymd}_{baseprefix}")
    return images, date_ymd, meeting_title

def upload_image_to_drive_and_get_url(drive_service, image_path):
    file_metadata = {
        'name': os.path.basename(image_path),
        'mimeType': 'image/png'
    }
    media = MediaFileUpload(image_path, mimetype='image/png')
    uploaded = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    file_id = uploaded.get('id')
    drive_service.permissions().create(
        fileId=file_id,
        body={'type': 'anyone', 'role': 'reader'},
        fields='id'
    ).execute()
    public_url = f"https://drive.google.com/uc?id={file_id}"
    return public_url

def insert_images_to_slide(slides_service, presentation_id, image_urls, slide_title):
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
    requests = [{
        "createShape": {
            "objectId": f"title_{new_slide_id}",
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": new_slide_id,
                "size": {"height": {"magnitude": 60, "unit": "PT"}, "width": {"magnitude": 600, "unit": "PT"}},
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 60, "translateY": 20, "unit": "PT"}
            }
        }
    }, {
        "insertText": {
            "objectId": f"title_{new_slide_id}",
            "insertionIndex": 0,
            "text": slide_title
        }
    }]
    # Place up to 2 images side by side
    x_start = 0
    y = 100
    image_width = 340
    image_height = 230
    spacing = 20
    for idx, img_url in enumerate(image_urls):
        x = x_start + idx * (image_width + spacing)
        requests.append({
            "createImage": {
                "url": img_url,
                "elementProperties": {
                    "pageObjectId": new_slide_id,
                    "size": {"height": {"magnitude": image_height, "unit": "PT"}, "width": {"magnitude": image_width, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"}
                }
            }
        })
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}).execute()
    print("Inserted images side-by-side into slide", new_slide_id)

def main():
    check_and_help_credentials()
    start, end = get_date_range_from_args_or_prompt()
    time_min = start.isoformat() + 'Z'
    time_max = end.isoformat() + 'Z'
    calendar_service, drive_service, slides_service, sheets_service = get_google_services()

    presentation = slides_service.presentations().create(body={
        'title': 'Automated Meeting Analysis Summary'
    }).execute()
    presentation_id = presentation.get('presentationId')
    url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
    print(f"\nCreated Slides: {url}")

    with tempfile.TemporaryDirectory() as temp_dir:
        events = find_meetings_with_gemini_notes(calendar_service, time_min, time_max)
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
            baseprefix = re.sub(r'\W+', '_', event.get('summary', 'meeting')).lower()
            # Get event date in YYYY-MM-DD
            start_datetime = event.get('start', {}).get('dateTime') or event.get('start', {}).get('date')
            if start_datetime:
                if 'T' in start_datetime:
                    date_ymd = start_datetime.split('T')[0]
                else:
                    date_ymd = start_datetime
            else:
                date_ymd = datetime.datetime.now().strftime("%Y-%m-%d")
            images, _, meeting_title = analyze_transcript_and_generate_images(transcript_text, baseprefix, override_date=date_ymd)
            image_urls = [upload_image_to_drive_and_get_url(drive_service, img) for img in images]
            slide_title = f"{date_ymd} – {meeting_title}"
            if image_urls:
                insert_images_to_slide(slides_service, presentation_id, image_urls, slide_title)

    print(f"\nAll done. View your slides at:\n{url}\n")
    webbrowser.open(url)


if __name__ == "__main__":
    main()
