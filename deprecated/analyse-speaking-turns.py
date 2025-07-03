import re
import os
from collections import Counter
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np

# Load the transcript from a text file
with open("transcript.txt", "r", encoding="utf-8") as f:
    raw = f.read()

lines = raw.splitlines()

# 1. Find the line index of the first occurrence of '00:00:00'
start_index = next(i for i, line in enumerate(lines) if '00:00:00' in line)
body_lines = lines[start_index:]  # Transcript proper

# 2. Get meeting name from second line, use for filenames
meeting_title = lines[1] if len(lines) > 1 else "meeting"
# sanitize for filename
basename = meeting_title.strip().lower()
basename = re.sub(r'[^a-z0-9 ]', '', basename)   # keep alnum and space
basename = re.sub(r'\s+', '_', basename)         # underscores

# Recombine the transcript
transcript = "\n".join(body_lines)

# 3. Extract all speaker names (up to colon at start of line)
speaker_pattern = re.compile(r"^([A-Za-z .'-]+):", re.MULTILINE)
speaking_turns = speaker_pattern.findall(transcript)

# 4. Bar Chart of Total Speaking Turns
turn_counts = Counter(speaking_turns)
sorted_turns = dict(turn_counts.most_common())

plt.figure(figsize=(10, 6))
plt.bar(sorted_turns.keys(), sorted_turns.values())
plt.title(meeting_title + " – Total Speaking Turns")
plt.xlabel("Participant")
plt.ylabel("Number of Speaking Turns")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f"{basename}_bar_chart_total_speaking_turns.png", dpi=300)
plt.savefig(f"{basename}_bar_chart_total_speaking_turns.pdf")
plt.savefig(f"{basename}_bar_chart_total_speaking_turns.jpg", dpi=300)
plt.close()

# 5. Cumulative Frequency Plot
participants = list(sorted_turns.keys())
cumulative_counts = {name: [] for name in participants}
running_total = {name: 0 for name in participants}
for speaker in speaking_turns:
    for name in participants:
        if name == speaker:
            running_total[name] += 1
        cumulative_counts[name].append(running_total[name])

plt.figure(figsize=(12, 6))
for name in participants:
    plt.plot(cumulative_counts[name], label=name)
plt.title(meeting_title + " – Cumulative Speaking Turns")
plt.xlabel("Turn Number")
plt.ylabel("Cumulative Speaking Turns")
plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
plt.tight_layout()
plt.savefig(f"{basename}_cumulative_speaking_turns.png", dpi=300)
plt.savefig(f"{basename}_cumulative_speaking_turns.pdf")
plt.savefig(f"{basename}_cumulative_speaking_turns.jpg", dpi=300)
plt.close()

# 6. Speaker Response Network (with padded axis)
edges = []
prev_speaker = None
for speaker in speaking_turns:
    if prev_speaker and speaker != prev_speaker:
        edges.append((prev_speaker, speaker))
    prev_speaker = speaker

edge_counts = Counter(edges)
G = nx.DiGraph()
for (src, dst), weight in edge_counts.items():
    G.add_edge(src, dst, weight=weight)

num_nodes = G.number_of_nodes()
fig_size = max(10, num_nodes * 1.2)
pos = nx.circular_layout(G)
node_sizes = [500 + 300 * turn_counts[node] for node in G.nodes()]
edge_weights = [G[u][v]['weight'] * 1.2 for u, v in G.edges()]

plt.figure(figsize=(fig_size, fig_size))
nx.draw_networkx_nodes(G, pos, node_size=node_sizes, node_color='lightblue')
nx.draw_networkx_edges(G, pos, edgelist=G.edges(), width=edge_weights, arrowstyle='-|>', arrowsize=20)
nx.draw_networkx_labels(G, pos, font_size=10, font_weight='bold')
plt.title(meeting_title + " – Speaker Response Network")
plt.axis('off')

# --- Padding axes so large nodes don't get cropped ---
x_vals, y_vals = zip(*pos.values())
largest_radius = max(node_sizes) ** 0.5 / 72  # in inches
x_pad = (max(x_vals) - min(x_vals)) * 0.10 + largest_radius
y_pad = (max(y_vals) - min(y_vals)) * 0.10 + largest_radius
plt.xlim(min(x_vals) - x_pad, max(x_vals) + x_pad)
plt.ylim(min(y_vals) - y_pad, max(y_vals) + y_pad)

plt.tight_layout()
plt.savefig(f"{basename}_speaker_response_network.png", dpi=300)
plt.savefig(f"{basename}_speaker_response_network.pdf")
plt.savefig(f"{basename}_speaker_response_network.jpg", dpi=300)
plt.close()

print("Plots saved using basename:", basename)
