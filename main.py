import json
import os
import time
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import List

import anthropic
import openpyxl
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

load_dotenv()

API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
MODEL_NAME = os.getenv("MODEL_NAME", "claude-sonnet-4-5")
PORT = int(os.getenv("PORT", "8000"))
LISTS_FILE = os.getenv("LISTS_FILE", "lists.json")

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

def _read_lists() -> list[dict]:
    p = Path(LISTS_FILE)
    if not p.exists():
        return []
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_lists(data: list[dict]) -> None:
    with open(LISTS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

class ListCreate(BaseModel):
    name: str = Field(..., max_length=200)
    type: str = Field(..., max_length=100)
    content: str = Field(..., max_length=500_000)


class GenerateRequest(BaseModel):
    trip_type: str = ""
    location: str = ""
    duration: str = ""
    season: str = ""
    group_size: str = ""
    weight_priority: str = ""
    special_considerations: str = ""
    notes: str = ""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/")
async def root():
    return FileResponse("static/index.html")


@app.get("/api/lists")
async def get_lists():
    return _read_lists()


@app.post("/api/lists")
async def create_list(payload: ListCreate):
    lists = _read_lists()
    entry = {
        "id": str(int(time.time() * 1000)),
        "name": payload.name.strip(),
        "type": payload.type.strip(),
        "date_added": date.today().isoformat(),
        "content": payload.content,
    }
    lists.append(entry)
    _write_lists(lists)
    return entry


@app.delete("/api/lists/{list_id}")
async def delete_list(list_id: str):
    lists = _read_lists()
    filtered = [l for l in lists if l["id"] != list_id]
    if len(filtered) == len(lists):
        raise HTTPException(status_code=404, detail="List not found")
    _write_lists(filtered)
    return {"ok": True}


@app.post("/api/import-xlsx")
async def import_xlsx(files: List[UploadFile]):
    """Bulk import .xlsx files as gear lists. Extracts all sheets from each file."""
    imported = []
    lists = _read_lists()

    for upload in files:
        if not upload.filename:
            continue
        raw = await upload.read()
        try:
            wb = openpyxl.load_workbook(BytesIO(raw), read_only=True, data_only=True)
        except Exception:
            continue

        # Derive a name from the filename (strip extension)
        base_name = Path(upload.filename).stem

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = []
            for row in ws.iter_rows(values_only=True):
                cells = [str(c).strip() if c is not None else "" for c in row]
                line = "\t".join(cells).strip()
                if line and line != "\t" * len(cells):
                    rows.append(line)
            if not rows:
                continue

            # Use sheet name in the list name if the workbook has multiple sheets
            name = base_name if len(wb.sheetnames) == 1 else f"{base_name} — {sheet_name}"
            content = "\n".join(rows)

            # Guess trip type from filename keywords
            lower = base_name.lower()
            trip_type = ""
            for kw, tt in [("climb", "alpine climbing"), ("ski", "ski touring"),
                           ("backpack", "backpacking"), ("hik", "day hiking"),
                           ("raft", "rafting/paddling"), ("river", "rafting/paddling"),
                           ("camp", "car camping"), ("snow camp", "snow camping")]:
                if kw in lower:
                    trip_type = tt
                    break

            entry = {
                "id": str(int(time.time() * 1000) + len(imported)),
                "name": name[:200],
                "type": trip_type,
                "date_added": date.today().isoformat(),
                "content": content[:500_000],
            }
            lists.append(entry)
            imported.append({"name": entry["name"], "type": entry["type"], "items": len(rows)})

        wb.close()

    _write_lists(lists)
    return {"imported": imported, "count": len(imported)}


@app.post("/api/generate")
async def generate(req: GenerateRequest):
    if not API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured")

    past_lists = _read_lists()

    # Build the past-lists context block
    past_lists_block = ""
    if past_lists:
        entries = []
        for pl in past_lists:
            entries.append(f"--- {pl['name']} ({pl['type']}) ---\n{pl['content']}\n")
        past_lists_block = (
            f"\nThe user has provided {len(past_lists)} past packing lists. "
            "Use these to personalize recommendations — reflect what they typically bring, "
            "flag gaps, and note any upgrades worth considering:\n\n"
            + "\n".join(entries)
        )

    prompt = f"""You are an expert outdoor guide and gear specialist. Generate a comprehensive, \
well-organized packing list for the following trip:

Trip type: {req.trip_type}
Location/objective: {req.location}
Duration: {req.duration}
Season: {req.season}
Group size: {req.group_size}
Weight priority: {req.weight_priority}
Special considerations: {req.special_considerations}
Notes: {req.notes}
{past_lists_block}
Organize by category (Navigation, Shelter, Sleep system, Clothing layers, \
Food & water, Safety, Tools, Personal — use categories appropriate to this \
trip type).

Format each category as a markdown heading (## CATEGORY NAME) followed by a \
markdown table with these exact columns:

| Item | Priority | Notes |
|------|----------|-------|
| Item name | OPTIONAL | Only if noteworthy |

Rules for the table:
- Priority: leave blank for essential items. Only write OPTIONAL for non-essential items.
- Notes: keep very brief (a few words). Leave blank when the item is self-explanatory. \
No need to explain obvious gear like "pack" or "helmet". Only add a note when there is \
something specific to this trip, season, or group that the user should know.

Be specific to the trip type, season, terrain, and duration.

End with a "Key considerations" section with 2–3 specific tips for this \
trip type, location, and season combination. Write Key considerations as \
numbered prose paragraphs, not tables."""

    client = anthropic.Anthropic(api_key=API_KEY)

    def event_stream():
        with client.messages.stream(
            model=MODEL_NAME,
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            for text in stream.text_stream:
                yield f"data: {json.dumps(text)}\n\n"
        yield "data: [DONE]\n\n"

    return StreamingResponse(event_stream(), media_type="text/event-stream")
