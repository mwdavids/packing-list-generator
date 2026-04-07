# Packing List Generator

A local web app that generates tailored outdoor trip packing lists from natural language trip descriptions, informed by your personal library of past packing lists. Built for PNW outdoor trips — alpine climbing, ski touring, backpacking, rafting, snow camping, day hiking, and more.

## Setup

```bash
# Clone and enter the project
git clone https://github.com/mwdavids/packing-list-generator.git
cd packing-list-generator

# Install dependencies
pip install -r requirements.txt

# Configure
cp .env.example .env
# Edit .env and add your ANTHROPIC_API_KEY

# Run
uvicorn main:app --reload

# Open http://localhost:8000
```

## Adding past lists

Go to the **My Lists** tab and click **+ Add list**. You can:

- **Paste** any format — bullet lists, numbered lists, freeform text, tab-separated data from Google Sheets
- **Import a file** — `.csv`, `.tsv`, or `.txt` — the file contents populate the text area

Past lists are stored in `lists.json` and used to personalize future generated lists (reflecting your actual gear, flagging gaps, suggesting upgrades).

## Exporting to Google Sheets

After generating a list:

1. Click **Copy as CSV (for Sheets)**
2. Open a Google Sheet, click a cell, and paste (Ctrl+V / Cmd+V)
3. The data pastes as tab-separated columns: Category, Item, Essential, Note

For plain text sharing (messages, email), use **Copy text**.

## Swapping the model

Edit `MODEL_NAME` in your `.env` file:

```
MODEL_NAME=claude-sonnet-4-5
```

Change to any Anthropic model string (e.g. `claude-sonnet-4-20250514`). Swapping to a non-Anthropic provider (OpenAI, Ollama, etc.) requires updating the client initialization in `main.py`.

## Data

`lists.json` is your gear library — it stores all saved packing lists locally. **Back it up** if your lists are valuable. It's excluded from git by default via `.gitignore`. If you want to version it, remove `lists.json` from `.gitignore`.

## Future extensions

- **Google Sheets sync** — OAuth flow to pull sheets directly by URL
- **Multiple users / profiles** — profile field per list, profile selector in UI
- **Tagging and filtering** — filter library by trip type, season, year
- **List diffing** — compare two generated lists side by side
- **Alternative models** — swap to GPT-4o, Gemini, or local Ollama via config


