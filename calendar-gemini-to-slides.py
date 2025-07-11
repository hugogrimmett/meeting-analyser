import os
import sys
import pickle
import tempfile
import re
import datetime
import argparse
import webbrowser
import colorsys
from collections import Counter, defaultdict
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload
import matplotlib.pyplot as plt
import matplotlib as mpl
import numpy as np
from scipy.stats import t
import warnings
warnings.filterwarnings("ignore", message=".*NotOpenSSLWarning.*")


CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/calendar.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'openid',
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
                        help="Start date (YYYY-MM-DD), default is 7 days ago.")
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
    # Use the People API to get the user's email
    people_service = build('people', 'v1', credentials=creds)
    me = people_service.people().get(resourceName='people/me', personFields='emailAddresses').execute()
    email = None
    emails = me.get('emailAddresses', [])
    if emails:
        email = emails[0].get('value')
    calendar_service = build('calendar', 'v3', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    slides_service = build('slides', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    return calendar_service, drive_service, slides_service, sheets_service, email

def find_meetings_with_gemini_notes(calendar_service, time_min, time_max):
    events = []
    print('Getting events...')
    page_token = None
    total_fetched = 0
    batch_num = 0
    latest_event_time = None
    while True:
        batch_num += 1
        events_result = calendar_service.events().list(
            calendarId='primary',
            timeMin=time_min,
            timeMax=time_max,
            maxResults=250,
            singleEvents=True,
            orderBy='startTime',
            pageToken=page_token
        ).execute()
        items = events_result.get('items', [])
        total_fetched += len(items)
        if items:
            earliest = items[0].get('start', {}).get('dateTime') or items[0].get('start', {}).get('date')
            latest = items[-1].get('start', {}).get('dateTime') or items[-1].get('start', {}).get('date')
            latest_event_time = latest
            print(f"  Batch {batch_num}: {len(items)} events ({earliest} to {latest})")
        else:
            print(f"  Batch {batch_num}: No events in this batch.")

        for event in items:
            attachments = event.get('attachments', [])
            found_gemini = False
            for att in attachments:
                if 'gemini' in att.get('title', '').lower():
                    events.append({'event': event, 'attachment': att})
                    found_gemini = True
            # Fallback: Search description for Gemini Google Doc links
            if not found_gemini and event.get('description'):
                match = re.search(r'https://docs\.google\.com/document/d/([a-zA-Z0-9_-]+)', event['description'])
                if match:
                    doc_id = match.group(1)
                    fake_attachment = {'fileId': doc_id, 'title': 'Gemini Notes (from description link)'}
                    events.append({'event': event, 'attachment': fake_attachment})

        page_token = events_result.get('nextPageToken')
        if not page_token:
            print(f"Reached end of date range after fetching {total_fetched} events. Last event date: {latest_event_time}")
            break
        elif len(items) == 0:
            print(f"No more events to fetch, but received nextPageToken. Stopping.")
            break
        else:
            print(f"Fetched {total_fetched} events so far, continuing with next batch...")

    print(f"Found {len(events)} events with Gemini notes in the specified date range.")
    return events

def get_transcript_from_gemini_drive_file(drive_service, sheets_service, file_id):
    try:
        file_metadata = drive_service.files().get(fileId=file_id, fields="mimeType, name").execute()
    except Exception as e:
        print(f"  WARNING: Could not fetch file {file_id} from Drive: {e}")
        return None
    if file_metadata['mimeType'] == 'application/vnd.google-apps.spreadsheet':
        pass  # Could add future Sheet logic here
    elif file_metadata['mimeType'] == 'application/vnd.google-apps.document':
        try:
            doc_content = drive_service.files().export(fileId=file_id, mimeType='text/plain').execute()
            text = doc_content.decode('utf-8') if isinstance(doc_content, bytes) else doc_content
            return text
        except Exception as e:
            print(f"  WARNING: Error fetching Google Doc content for file {file_id}: {e}")
            return None
    else:
        print(f"  WARNING: File {file_id} is not a Google Sheet or Doc. MIME type is: {file_metadata['mimeType']}")
    return None

def collect_all_participants(events, drive_service, sheets_service):
    participant_order = []
    participant_set = set()
    for evt in events:
        event = evt['event']
        att = evt['attachment']
        file_id = att.get('fileId')
        if not file_id:
            continue
        transcript_text = get_transcript_from_gemini_drive_file(drive_service, sheets_service, file_id)
        if not transcript_text:
            continue
        lines = [l for l in transcript_text.split('\n') if l.strip()]
        start_index = next((i for i, line in enumerate(lines) if '00:00:00' in line), None)
        transcript_lines = lines[start_index+1:] if start_index is not None else lines
        for line in transcript_lines:
            match = re.match(r"^([A-Za-z .'-]+):(.*)$", line)
            if match:
                name = match.group(1).strip()
                if name not in participant_set:
                    participant_order.append(name)
                    participant_set.add(name)
    return participant_order

def distinct_color_grid(n):
    """
    Generate n visually distinct RGB colors by using a grid in HLS space.
    """
    import math
    # Choose grid size based on n (e.g. sqrt for hue, 2–3 levels for lightness)
    k = int(math.ceil(n ** 0.5))  # number of hues
    l_values = [0.45, 0.65, 0.8] if n > 20 else [0.5, 0.7]
    colors = []
    for l in l_values:
        for i in range(k):
            h = i / float(k)
            s = 0.7
            rgb = colorsys.hls_to_rgb(h, l, s)
            colors.append(rgb)
            if len(colors) >= n:
                break
        if len(colors) >= n:
            break
    return colors[:n]

def make_global_color_dict(participants):
    n = len(participants)
    colors = distinct_color_grid(n)
    return dict(zip(participants, colors))

def parse_timestamp(s):
    if not s: return None
    parts = s.strip().split(':')
    if len(parts) != 3: return None
    try:
        h, m, sec = [int(float(x)) for x in parts]
        return h * 3600 + m * 60 + sec
    except Exception:
        return None

def per_participant_wpm(transcript_lines):
    participant_times = defaultdict(list)
    participant_words = defaultdict(list)
    timestamp_pattern = re.compile(r"^(\d{1,2}:\d{2}:\d{2})")
    speaker_pattern = re.compile(r"^([A-Za-z .'-]+):(.*)$")
    current_timestamp = None
    for line in transcript_lines:
        line = line.strip()
        if not line:
            continue
        ts_match = timestamp_pattern.match(line)
        if ts_match:
            current_timestamp = ts_match.group(1)
            continue  # go to next line; this line isn't a speaker turn
        speaker_match = speaker_pattern.match(line)
        if speaker_match and current_timestamp:
            name = speaker_match.group(1).strip()
            utter = speaker_match.group(2).strip()
            participant_times[name].append(parse_timestamp(current_timestamp))
            participant_words[name].append(len(re.findall(r'\b\w+\b', utter)))
    # Calculate WPM per participant for this meeting
    wpm_dict = {}
    for name in participant_times:
        ts_list = [t for t in participant_times[name] if t is not None]
        if len(ts_list) < 2 or min(ts_list) == max(ts_list):
            continue  # skip: can't compute duration
        first, last = min(ts_list), max(ts_list)
        duration_min = (last - first) / 60.0
        total_words = sum(participant_words[name])
        wpm = total_words / duration_min if duration_min > 0 else None
        if wpm and wpm < 1000:  # filter out wild outliers
            wpm_dict[name] = wpm
    return wpm_dict


def analyze_transcript_and_generate_images(transcript_text, baseprefix, override_date=None, color_dict=None):
    os.makedirs("generated-files", exist_ok=True)
    lines = [l for l in transcript_text.split('\n') if l.strip()]
    date_ymd = override_date if override_date else datetime.datetime.now().strftime("%Y-%m-%d")
    meeting_title = baseprefix
    if lines and re.match(r'^[A-Za-z]{3} \d{1,2}, \d{4}', lines[0]):
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
        transcript_lines = body_lines[start_index+1:]
    else:
        transcript_lines = body_lines

    # Speaker turn parsing (for charts)
    speaker_turns = []
    current_speaker = None
    current_utterance = []

    for line in transcript_lines:
        line = line.strip()
        if not line:
            continue
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
                current_utterance.append(line)
    if current_speaker is not None and current_utterance:
        joined = ' '.join(current_utterance).strip()
        if joined:
            speaker_turns.append((current_speaker, joined))

    def count_words(text):
        return len(re.findall(r'\b\w+\b', text))

    word_counts = Counter()
    cumulative_word_counts = defaultdict(list)
    running_totals = Counter()
    total_words = 0
    for speaker, utterance in speaker_turns:
        words = count_words(utterance)
        word_counts[speaker] += words
        total_words += words
        running_totals[speaker] += words
        for p in word_counts:
            cumulative_word_counts[p].append(running_totals[p])

    participants = list(word_counts.keys())
    images = []

    if color_dict is None:
        cmap = plt.get_cmap('tab10')
        color_list = [cmap(i % 10) for i in range(len(participants))]
        color_dict = dict(zip(participants, color_list))

    base_size = 14
    mpl.rcParams.update({
        'axes.titlesize': base_size + 6,
        'axes.labelsize': base_size + 4,
        'xtick.labelsize': base_size + 2,
        'ytick.labelsize': base_size + 2,
        'legend.fontsize': base_size + 2,
        'figure.titlesize': base_size + 8
    })

    # Bar Chart: Total Words Spoken
    plt.figure(figsize=(10, 6))
    sorted_word_counts = dict(word_counts.most_common())
    bar_colors = [color_dict.get(p, (0.5,0.5,0.5,1)) for p in sorted_word_counts.keys()]
    plt.bar(sorted_word_counts.keys(), sorted_word_counts.values(), color=bar_colors)
    plt.title(meeting_title + " – Total Words Spoken")
    plt.xlabel("Participant")
    plt.ylabel("Number of Words Spoken")
    plt.xticks(rotation=45)
    plt.tight_layout()
    fname1 = os.path.join("generated-files", f"{date_ymd}_{baseprefix}_bar_chart_total_words_spoken.png")
    plt.savefig(fname1, dpi=300)
    plt.close()
    images.append(fname1)

    # Cumulative Word Count Plot
    plt.figure(figsize=(12, 6))
    for name in participants:
        plt.plot(cumulative_word_counts[name], label=name, color=color_dict.get(name, (0.5,0.5,0.5,1)), linewidth=2.5)
    plt.title(meeting_title + " – Cumulative Words Spoken")
    plt.xlabel("Turn Number")
    plt.ylabel("Cumulative Words Spoken")
    if participants:
        plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
    plt.tight_layout()
    fname2 = os.path.join("generated-files", f"{date_ymd}_{baseprefix}_cumulative_words_spoken.png")
    plt.savefig(fname2, dpi=300)
    plt.close()
    images.append(fname2)

    print("Plots saved using basename:", f"{date_ymd}_{baseprefix}")
    return images, date_ymd, meeting_title, total_words, transcript_lines

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

def insert_image_slide(slides_service, presentation_id, image_url, slide_title):
    create_slide_req = [{
        "createSlide": {
            "insertionIndex": 1,  # after title
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
                "size": {"height": {"magnitude": 60, "unit": "PT"}, "width": {"magnitude": 700, "unit": "PT"}},
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 40, "translateY": 20, "unit": "PT"}
            }
        }
    }, {
        "insertText": {
            "objectId": f"title_{new_slide_id}",
            "insertionIndex": 0,
            "text": slide_title
        }
    }, {
        "createImage": {
            "url": image_url,
            "elementProperties": {
                "pageObjectId": new_slide_id,
                "size": {"height": {"magnitude": 340, "unit": "PT"}, "width": {"magnitude": 620, "unit": "PT"}},
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 40, "translateY": 90, "unit": "PT"}
            }
        }
    }]
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}).execute()
    print("Inserted meta-analysis WPM bar chart slide.")

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

def insert_custom_title_slide(slides_service, presentation_id, date_range, email):
    # Insert BLANK slide at index 0
    requests = [{
        "createSlide": {
            "insertionIndex": 0,
            "slideLayoutReference": {"predefinedLayout": "BLANK"},
            "objectId": "custom_title_slide"
        }
    }]
    # Set black background
    requests.append({
        "updatePageProperties": {
            "objectId": "custom_title_slide",
            "pageProperties": {
                "pageBackgroundFill": {
                    "solidFill": {
                        "color": {"rgbColor": {"red": 0, "green": 0, "blue": 0}},
                        "alpha": 1.0
                    }
                }
            },
            "fields": "pageBackgroundFill.solidFill.color"
        }
    })
    # Big title text (center vertically, left aligned)
    requests.append({
        "createShape": {
            "objectId": "main_title_textbox",
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": "custom_title_slide",
                "size": {"height": {"magnitude": 120, "unit": "PT"}, "width": {"magnitude": 700, "unit": "PT"}},
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 60, "translateY": 120, "unit": "PT"}
            }
        }
    })
    requests.append({
        "insertText": {
            "objectId": "main_title_textbox",
            "insertionIndex": 0,
            "text": "Meeting Analyser Results"
        }
    })
    requests.append({
        "updateTextStyle": {
            "objectId": "main_title_textbox",
            "style": {
                "foregroundColor": {
                    "opaqueColor": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}
                },
                "fontFamily": "Arial",
                "fontSize": {"magnitude": 50, "unit": "PT"},
                "bold": True
            },
            "fields": "foregroundColor,fontFamily,fontSize,bold"
        }
    })
    # Date range, slightly smaller, left aligned
    requests.append({
        "createShape": {
            "objectId": "date_range_textbox",
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": "custom_title_slide",
                "size": {"height": {"magnitude": 60, "unit": "PT"}, "width": {"magnitude": 700, "unit": "PT"}},
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 65, "translateY": 190, "unit": "PT"}
            }
        }
    })
    requests.append({
        "insertText": {
            "objectId": "date_range_textbox",
            "insertionIndex": 0,
            "text": f"Date range: {date_range}"
        }
    })
    requests.append({
        "updateTextStyle": {
            "objectId": "date_range_textbox",
            "style": {
                "foregroundColor": {
                    "opaqueColor": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}
                },
                "fontFamily": "Arial",
                "fontSize": {"magnitude": 28, "unit": "PT"},
            },
            "fields": "foregroundColor,fontFamily,fontSize"
        }
    })
    # Email line, left aligned, smaller
    requests.append({
        "createShape": {
            "objectId": "email_textbox",
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": "custom_title_slide",
                "size": {"height": {"magnitude": 40, "unit": "PT"}, "width": {"magnitude": 700, "unit": "PT"}},
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": 65, "translateY": 240, "unit": "PT"}
            }
        }
    })
    requests.append({
        "insertText": {
            "objectId": "email_textbox",
            "insertionIndex": 0,
            "text": email
        }
    })
    requests.append({
        "updateTextStyle": {
            "objectId": "email_textbox",
            "style": {
                "foregroundColor": {
                    "opaqueColor": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}
                },
                "fontFamily": "Arial",
                "fontSize": {"magnitude": 18, "unit": "PT"},
            },
            "fields": "foregroundColor,fontFamily,fontSize"
        }
    })
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}
    ).execute()
    print("Inserted custom black title slide.")



def main():
    check_and_help_credentials()
    start, end = get_date_range_from_args_or_prompt()
    time_min = start.isoformat() + 'Z'
    time_max = end.isoformat() + 'Z'
    calendar_service, drive_service, slides_service, sheets_service, user_email = get_google_services()

    start_str = start.strftime('%Y-%m-%d')
    end_str = end.strftime('%Y-%m-%d')
    your_email = "your@email.com"  # <-- use your real email here

    # After creating the presentation
    presentation = slides_service.presentations().create(body={
        'title': f'Meeting Analyser Results - {user_email} - {start_str} - {end_str}'
    }).execute()
    presentation_id = presentation.get('presentationId')

    # --- DELETE THE DEFAULT SLIDE (index 0) ---
    slides_list = slides_service.presentations().get(presentationId=presentation_id).execute().get('slides', [])
    default_slide_id = slides_list[0]['objectId']
    slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": [{"deleteObject": {"objectId": default_slide_id}}]}
    ).execute()

    insert_custom_title_slide(slides_service, presentation_id, f"{start_str} to {end_str}", user_email)

    url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
    print(f"\nCreated Slides: {url}")

    all_participants = []
    all_wpm_by_participant = defaultdict(list)

    with tempfile.TemporaryDirectory() as temp_dir:
        events = find_meetings_with_gemini_notes(calendar_service, time_min, time_max)
        if not events:
            print("No events with Gemini notes attachments found.")
            return
        all_participants = collect_all_participants(events, drive_service, sheets_service)
        color_dict = make_global_color_dict(all_participants)

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

            # Optional: quick pre-check—does transcript_text contain at least two timestamps?
            timestamp_lines = [l for l in transcript_text.split('\n') if re.match(r'^\d{1,2}:\d{2}:\d{2}', l.strip())]
            if len(timestamp_lines) < 2:
                print(f"Transcript in {att.get('title')} is missing enough timestamps for WPM analysis, skipping meeting.")
                continue

            # Now, and ONLY now, generate images/plots and continue
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
            images, _, meeting_title, total_words, transcript_lines = analyze_transcript_and_generate_images(
                transcript_text, baseprefix, override_date=date_ymd, color_dict=color_dict)
            meeting_wpm = per_participant_wpm(transcript_lines)
            if not meeting_wpm:
                print(f"Transcript in {att.get('title')} is missing enough timestamps for WPM analysis, skipping meeting.")
                continue  # Skip this meeting entirely
            for name, wpm in meeting_wpm.items():
                all_wpm_by_participant[name].append(wpm)
            image_urls = [upload_image_to_drive_and_get_url(drive_service, img) for img in images]
            slide_title = f"{date_ymd} – {meeting_title}"
            if image_urls:
                insert_images_to_slide(slides_service, presentation_id, image_urls, slide_title)


        # Only keep participants with >5 WPM entries
        filtered_wpm_by_participant = {name: wpm_list for name, wpm_list in all_wpm_by_participant.items() if len(wpm_list) > 5}
        print(f"{len(filtered_wpm_by_participant)} participants spoke in >5 meetings and are included in the WPM analysis.")

        # Compute global per-participant WPM mean/CI and plot
        wpm_stats = {}
        for name, wpm_list in filtered_wpm_by_participant.items():
            arr = np.array(wpm_list)
            mean = arr.mean()
            n = len(arr)
            if n > 1:
                sem = arr.std(ddof=1)/np.sqrt(n)
                ci = t.ppf(0.975, n-1) * sem
            else:
                ci = 0  # Can't compute CI for n=1
            wpm_stats[name] = (mean, ci, n)
        # Plot and upload
        if wpm_stats:
            names = [k for k, (mean, ci, n) in sorted(wpm_stats.items(), key=lambda x: -x[1][0])]
            means = [wpm_stats[k][0] for k in names]
            cis = [wpm_stats[k][1] for k in names]
            bar_colors = [color_dict.get(name, (0.5,0.5,0.5,1)) for name in names]
            plt.figure(figsize=(max(7, len(names)*1.5),6))
            plt.bar(names, means, yerr=cis, capsize=7, color=bar_colors)
            plt.ylabel('Words Per Meeting-Minute (WPM)', fontsize=18)
            plt.xlabel('Participant', fontsize=18)
            plt.title('Participant Words Per Meeting-Minute (Mean ± 95% CI)\n(Across all meetings)', fontsize=20)
            plt.xticks(rotation=30, fontsize=16)
            plt.yticks(fontsize=16)
            plt.tight_layout()
            meta_bar_file = os.path.join("generated-files", "global_participant_wpm_bar.png")
            plt.savefig(meta_bar_file, dpi=300)
            plt.close()
            # Upload and insert as slide
            meta_bar_url = upload_image_to_drive_and_get_url(drive_service, meta_bar_file)
            insert_image_slide(slides_service, presentation_id, meta_bar_url, "Meta-Analysis: Participant Words Per Minute (WPM)")

    print(f"\nAll done. View your slides at:\n{url}\n")
    webbrowser.open(url)

if __name__ == "__main__":
    main()
