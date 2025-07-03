import re
import os
import datetime
from collections import Counter, defaultdict
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np

# Load the transcript from a text file
with open("transcript.txt", "r", encoding="utf-8") as f:
    raw = f.read()

lines = raw.splitlines()

# 0. Parse date from the first line (e.g., 'Jul 3, 2025') into YYYY-MM-DD
date_line = lines[0].strip()
date_match = re.search(r'([A-Za-z]{3}) (\d{1,2}), (\d{4})', date_line)
if not date_match:
    print("Could not parse date from the first line of transcript.")
    exit(1)
month_str, day_str, year_str = date_match.groups()
try:
    date_obj = datetime.datetime.strptime(f"{month_str} {day_str} {year_str}", "%b %d %Y")
    date_ymd = date_obj.strftime("%Y-%m-%d")
except Exception as e:
    print(f"Error parsing date: {e}")
    exit(1)

# 1. Find the first occurrence of '00:00:00'
start_index = next((i for i, line in enumerate(lines) if '00:00:00' in line), None)
if start_index is None:
    print("Error: Could not find a line with '00:00:00' in transcript.txt")
    exit(1)
body_lines = lines[start_index:]  # Meeting content only

# 2. Infer meeting name from second line
meeting_title = lines[1] if len(lines) > 1 else "meeting"
basename = meeting_title.strip().lower()
basename = re.sub(r'[^a-z0-9 ]', '', basename)
basename = re.sub(r'\s+', '_', basename)

# Prefix for all files
file_prefix = f"{date_ymd}_{basename}"

# 3. Parse transcript into speaker turns (list of (speaker, utterance) tuples)
speaker_turns = []
current_speaker = None
current_utterance = []

for line in body_lines:
    speaker_match = re.match(r"^([A-Za-z .'-]+):(.*)$", line)
    if speaker_match:
        # Save the previous speaker's utterance, if any
        if current_speaker is not None:
            joined = ' '.join(current_utterance).strip()
            if joined:
                speaker_turns.append((current_speaker, joined))
        current_speaker = speaker_match.group(1).strip()
        current_utterance = [speaker_match.group(2).strip()]
    else:
        if current_speaker is not None:
            current_utterance.append(line.strip())

# Append the last speaker's turn, if any
if current_speaker is not None and current_utterance:
    joined = ' '.join(current_utterance).strip()
    if joined:
        speaker_turns.append((current_speaker, joined))

if not speaker_turns:
    print("No speaker turns found after 00:00:00. Check your transcript format.")
    exit(1)

# 4. Count words spoken per participant
def count_words(text):
    return len(re.findall(r'\b\w+\b', text))

word_counts = Counter()
cumulative_word_counts = defaultdict(list)
running_totals = Counter()
ordered_speakers = []

for speaker, utterance in speaker_turns:
    words = count_words(utterance)
    word_counts[speaker] += words
    running_totals[speaker] += words
    ordered_speakers.append(speaker)
    # For cumulative plot, snapshot all participant totals at each turn
    for p in word_counts:
        cumulative_word_counts[p].append(running_totals[p])

participants = list(word_counts.keys())

# 5. Bar Chart: Total Words Spoken
plt.figure(figsize=(10, 6))
sorted_word_counts = dict(word_counts.most_common())
plt.bar(sorted_word_counts.keys(), sorted_word_counts.values())
plt.title(meeting_title + " – Total Words Spoken")
plt.xlabel("Participant")
plt.ylabel("Number of Words Spoken")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f"{file_prefix}_bar_chart_total_words_spoken.png", dpi=300)
plt.savefig(f"{file_prefix}_bar_chart_total_words_spoken.pdf")
plt.savefig(f"{file_prefix}_bar_chart_total_words_spoken.jpg", dpi=300)
plt.close()

# 6. Cumulative Word Count Plot
plt.figure(figsize=(12, 6))
for name in participants:
    plt.plot(cumulative_word_counts[name], label=name)
plt.title(meeting_title + " – Cumulative Words Spoken")
plt.xlabel("Turn Number")
plt.ylabel("Cumulative Words Spoken")
if participants:
    plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
plt.tight_layout()
plt.savefig(f"{file_prefix}_cumulative_words_spoken.png", dpi=300)
plt.savefig(f"{file_prefix}_cumulative_words_spoken.pdf")
plt.savefig(f"{file_prefix}_cumulative_words_spoken.jpg", dpi=300)
plt.close()

print("Plots saved using basename:", file_prefix)
