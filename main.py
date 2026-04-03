import json
import os
import re
import time
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import List

import anthropic
import openpyxl
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
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


# ---------------------------------------------------------------------------
# xlsx export
# ---------------------------------------------------------------------------

class ExportRequest(BaseModel):
    markdown: str
    title: str = ""
    group_size: str = ""


def _parse_group_count(s: str) -> int:
    s = s.lower().strip()
    if "solo" in s or s == "1":
        return 1
    m = re.search(r"(\d+)", s)
    if m:
        val = int(m.group(1))
        # "3-4 people" → 4, "5+ people" → 5
        m2 = re.search(r"(\d+)\s*[-–]\s*(\d+)", s)
        if m2:
            return int(m2.group(2))
        return val
    return 1


def _parse_markdown_to_rows(text: str) -> list[dict]:
    """Parse the generated markdown into structured rows."""
    lines = text.split("\n")
    category = ""
    rows = []
    in_key_considerations = False

    for raw in lines:
        line = raw.strip()
        if not line:
            continue

        # Detect category headers
        header_match = (
            re.match(r"^#{1,3}\s+(.+)", line)
            or re.match(r"^\*\*(.+?)\*\*\s*$", line)
        )
        if header_match:
            h = header_match.group(1).strip()
            if re.search(r"key\s+considerations", h, re.I):
                in_key_considerations = True
                category = ""
                continue
            in_key_considerations = False
            category = h
            rows.append({"type": "header", "category": category})
            continue

        if in_key_considerations or not category:
            continue

        # Skip table separator rows
        if re.match(r"^\|[\s\-:]+\|", line):
            continue

        # Skip table header rows
        if re.match(r"^\|.*\bItem\b.*\|", line, re.I):
            continue

        # Table data row
        if line.startswith("|") and line.endswith("|"):
            cells = [c.strip() for c in line.split("|")[1:-1]]
            if len(cells) < 2:
                continue
            item = re.sub(r"\*\*", "", cells[0]).strip()
            priority = cells[1].strip() if len(cells) > 1 else ""
            note = cells[2].strip() if len(cells) > 2 else ""
            note = re.sub(r"[*_]", "", note)
            if not item or item == "-":
                continue
            rows.append({
                "type": "item",
                "category": category,
                "item": item,
                "priority": priority,
                "note": note,
            })
            continue

        # Bullet list fallback
        bullet_match = re.match(r"^[-*\d.)\s]+(.+)", line)
        if bullet_match:
            item_text = bullet_match.group(1).strip()
            priority = ""
            if re.search(r"\bOPTIONAL\b", item_text, re.I):
                priority = "OPTIONAL"
                item_text = re.sub(r"\bOPTIONAL\b", "", item_text, flags=re.I)
            item_text = re.sub(r"\*\*", "", item_text).strip()
            parts = re.split(r"\s*[—–\-:]\s+", item_text, maxsplit=1)
            item = parts[0].strip()
            note = parts[1].strip() if len(parts) > 1 else ""
            if item:
                rows.append({
                    "type": "item",
                    "category": category,
                    "item": item,
                    "priority": priority,
                    "note": note,
                })

    return rows


@app.post("/api/export-xlsx")
async def export_xlsx(req: ExportRequest):
    rows = _parse_markdown_to_rows(req.markdown)
    person_count = _parse_group_count(req.group_size)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Packing List"

    # Styles
    header_font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    header_fill = PatternFill(start_color="2E4057", end_color="2E4057", fill_type="solid")
    header_align = Alignment(horizontal="left", vertical="center")

    col_header_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    col_header_fill = PatternFill(start_color="4A6FA5", end_color="4A6FA5", fill_type="solid")

    item_font = Font(name="Calibri", size=10)
    optional_font = Font(name="Calibri", size=10, italic=True, color="888888")
    note_font = Font(name="Calibri", size=9, color="666666")
    check_font = Font(name="Calibri", size=10)

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    center_align = Alignment(horizontal="center", vertical="center")
    wrap_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Person column names
    person_names = []
    if person_count == 1:
        person_names = ["✓"]
    else:
        person_names = [f"Person {i+1}" for i in range(person_count)]

    # Title row
    if req.title:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3 + person_count)
        cell = ws.cell(row=1, column=1, value=req.title)
        cell.font = Font(name="Calibri", bold=True, size=16, color="2E4057")
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30
        current_row = 3
    else:
        current_row = 1

    total_cols = 3 + person_count  # Item, Priority, Notes, + person columns

    for row_data in rows:
        if row_data["type"] == "header":
            # Category header — full-width merged row
            ws.merge_cells(
                start_row=current_row, start_column=1,
                end_row=current_row, end_column=total_cols,
            )
            cell = ws.cell(row=current_row, column=1, value=row_data["category"])
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            for c in range(1, total_cols + 1):
                ws.cell(row=current_row, column=c).fill = header_fill
            ws.row_dimensions[current_row].height = 28
            current_row += 1

            # Column headers under each category
            col_headers = ["Item", "Priority", "Notes"] + person_names
            for ci, ch in enumerate(col_headers, 1):
                cell = ws.cell(row=current_row, column=ci, value=ch)
                cell.font = col_header_font
                cell.fill = col_header_fill
                cell.alignment = center_align
                cell.border = thin_border
            ws.row_dimensions[current_row].height = 22
            current_row += 1

        elif row_data["type"] == "item":
            # Item name
            cell = ws.cell(row=current_row, column=1, value=row_data["item"])
            cell.font = optional_font if row_data["priority"].upper() == "OPTIONAL" else item_font
            cell.alignment = wrap_align
            cell.border = thin_border

            # Priority
            cell = ws.cell(row=current_row, column=2, value=row_data["priority"])
            cell.font = Font(name="Calibri", size=9, italic=True, color="888888")
            cell.alignment = center_align
            cell.border = thin_border

            # Notes
            cell = ws.cell(row=current_row, column=3, value=row_data["note"])
            cell.font = note_font
            cell.alignment = wrap_align
            cell.border = thin_border

            # Person checkbox columns — empty with checkbox-style formatting
            for pi in range(person_count):
                cell = ws.cell(row=current_row, column=4 + pi, value="☐")
                cell.font = check_font
                cell.alignment = center_align
                cell.border = thin_border

            ws.row_dimensions[current_row].height = 20
            current_row += 1

    # Column widths
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 30
    for pi in range(person_count):
        col_letter = get_column_letter(4 + pi)
        ws.column_dimensions[col_letter].width = 12

    # Freeze panes (freeze below title)
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)

    # Write to bytes
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = re.sub(r"[^\w\s\-]", "", req.title or "packing-list").strip()[:80] or "packing-list"
    filename = filename.replace(" ", "_") + ".xlsx"

    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
