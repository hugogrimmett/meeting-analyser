# Meeting Analyser

Analyse Google Calendar meetings with Gemini notes, generate participation and communication graphs, and auto-create a Google Slides presentation of your results.

Written by Hugo Grimmett

July 2025

---

## Requirements

- Python 3.8 or newer
- Google Cloud Project with OAuth credentials (`credentials.json`)
- Access to your Google Calendar and Gemini meeting notes

---

## 1. Install Python (if not already installed)

**macOS (Homebrew):**
```sh
brew install python
```

**Ubuntu/Debian:**
```sh
sudo apt-get update
sudo apt-get install python3 python3-venv python3-pip
```

**Windows:**  
[Download from python.org](https://www.python.org/downloads/), and check “Add Python to PATH” during install.

---

## 2. Clone/download this project

```sh
git clone https://github.com/hugogrimmett/meeting-analyser
cd meeting-analyser
```

---

## 3. Create a virtual environment

```sh
python3 -m venv venv
source venv/bin/activate   # On Windows: venv\Scriptsctivate
```

---

## 4. Install dependencies

```sh
pip install -r requirements.txt
```

If you do not have a `requirements.txt`, create one using:
```sh
pip install pipreqs
pipreqs . --force
```

---

## 5. Add Google API Credentials

- Go to [Google Cloud Console](https://console.cloud.google.com/apis/credentials)
- Create OAuth 2.0 **Desktop App** credentials
- Download as `credentials.json` and place it in the project directory

---

## 6. Run the analyser

You can run the analyser interactively and it will prompt for a date range:
```sh
python3 calendar-gemini-to-slides.py
```

Or, to specify the date range on the command line:
```sh
python3 calendar-gemini-to-slides.py --start YYYY-MM-DD --end YYYY-MM-DD
```

- On first run, you’ll be prompted to authenticate with Google.

---

## 7. Output

- Your Google Slides presentation will be created and opened in your browser.
- Participation and analysis plots will be saved in the `generated-files/` directory.

---

## Tips

- To deactivate your virtual environment:  
  ```sh
  deactivate
  ```
- If you change scopes in your script, **delete `token.pickle`** and re-run to re-authenticate.

---

## Troubleshooting

- **SSL warnings on Mac:** Safe to ignore, unless you see actual HTTPS errors.
- **Permission errors:** Make sure your credentials and calendar access are correct.
- **No Gemini notes found:** Ensure your events have Gemini notes attached or linked in the description.

---