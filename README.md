# CCS Worksheet Update

Daily tool to enrich S4 Excel exports, find new lines, check discrepancies, and compare against existing worksheets.

## Free website (Streamlit Community Cloud)

1. Push this repo to GitHub (done if you followed the agent setup).
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app** → select this repo → Main file: `app.py` → Deploy.

**Note:** On the website, use **Upload files** (not laptop folder path). Folder path only works when you run the tool on your Mac.

## Run on your Mac

Double-click `Run CCS Tool.command`, or:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/streamlit run app.py
```
