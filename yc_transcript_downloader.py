"""
YC YouTube Transcript Downloader — Excel Mode
==============================================
Reads URLs + folder names from an Excel file,
downloads transcripts, saves PDFs, and writes
status back to the Excel file.

SETUP:
    pip install yt-dlp youtube-transcript-api reportlab openpyxl

RUN:
    python yc_transcript_downloader.py
    → Enter path to your Excel file when prompted

EXCEL FORMAT (yc_urls_template.xlsx):
    A: YouTube URL
    B: Folder Name
    C: Status       ← script writes this
    D: Error Reason ← script writes this
    E: Downloaded At← script writes this
"""

import re
import sys
import logging
from pathlib import Path
from datetime import datetime

import yt_dlp
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from youtube_transcript_api import YouTubeTranscriptApi
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable

# ─── CONFIG ───────────────────────────────────────────────────────────────────

SLEEP_SECS     = 1.5
STATUS_COL     = 3   # C
ERROR_COL      = 4   # D
TIMESTAMP_COL  = 5   # E

# Status values
S_PENDING   = "Pending"
S_DONE      = "Done"
S_FAILED    = "Failed"
S_DUPLICATE = "Skipped (duplicate)"

# Cell fill colours
FILL_DONE  = PatternFill("solid", start_color="E2EFDA")
FILL_FAIL  = PatternFill("solid", start_color="FCE4D6")
FILL_SKIP  = PatternFill("solid", start_color="FFF2CC")
FILL_CLEAR = PatternFill(fill_type=None)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
log = logging.getLogger(__name__)

# ─── EXCEL HELPERS ────────────────────────────────────────────────────────────

def load_rows(xlsx_path: Path):
    """
    Read all data rows from the Transcripts sheet.
    Returns (workbook, worksheet, list_of_row_dicts).
    Row dict keys: row_num, url, folder, status
    """
    wb = load_workbook(xlsx_path)
    ws = wb["Transcripts"]

    rows = []
    for row in ws.iter_rows(min_row=2):
        url    = (row[0].value or "").strip()
        folder = (row[1].value or "General").strip()
        status = (row[2].value or S_PENDING).strip()
        if not url:
            continue
        rows.append({
            "row_num": row[0].row,
            "url":     url,
            "folder":  folder,
            "status":  status,
        })

    return wb, ws, rows


def update_row(ws, row_num: int, status: str, error: str = "", timestamp: str = ""):
    """Write status, error reason and timestamp back into the Excel row."""
    ws.cell(row=row_num, column=STATUS_COL).value    = status
    ws.cell(row=row_num, column=ERROR_COL).value     = error
    ws.cell(row=row_num, column=TIMESTAMP_COL).value = timestamp

    if status == S_DONE:
        fill = FILL_DONE
    elif status == S_FAILED:
        fill = FILL_FAIL
    elif status == S_DUPLICATE:
        fill = FILL_SKIP
    else:
        fill = FILL_CLEAR

    for col in range(1, 6):
        ws.cell(row=row_num, column=col).fill = fill


# ─── TRANSCRIPT HELPERS ───────────────────────────────────────────────────────

def extract_video_id(url: str) -> str:
    match = re.search(r"(?:v=|youtu\.be/|embed/|shorts/)([a-zA-Z0-9_-]{11})", url)
    if not match:
        raise ValueError("Invalid YouTube URL — could not extract video ID")
    return match.group(1)


def get_video_metadata(url: str) -> dict:
    ydl_opts = {"quiet": True, "skip_download": True, "ignoreerrors": True}
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
        return {
            "title":          info.get("title", "Unknown Title"),
            "channel":        info.get("channel", "Y Combinator"),
            "upload_date":    info.get("upload_date", ""),
            "duration":       info.get("duration", 0),
            "url":            url,
            "playlist_title": info.get("playlist_title", ""),
        }
    except Exception as e:
        log.warning(f"Metadata fetch failed: {e}")
        return {"title": "Unknown Title", "channel": "", "upload_date": "",
                "duration": 0, "url": url, "playlist_title": ""}


def fetch_transcript(video_id: str) -> tuple[str, str]:
    """Returns (text, source). source can be 'manual/auto', 'auto', or 'none'."""
    try:
        api = YouTubeTranscriptApi()
        fetched = api.fetch(video_id, languages=["en", "en-US", "en-GB"])
        return format_transcript(list(fetched)), "auto/manual"
    except Exception:
        pass
    try:
        api = YouTubeTranscriptApi()
        fetched = api.fetch(video_id)
        return format_transcript(list(fetched)), "auto"
    except Exception as e:
        return "", f"No transcript: {e}"


def format_transcript(entries) -> str:
    lines, buffer, buf_dur = [], [], 0.0
    for entry in entries:
        text     = entry.text     if hasattr(entry, "text")     else entry.get("text", "")
        duration = entry.duration if hasattr(entry, "duration") else entry.get("duration", 0)
        text = re.sub(r"\[.*?\]", "", text).replace("\n", " ").strip()
        if not text:
            continue
        buffer.append(text)
        buf_dur += duration
        if buf_dur >= 60:
            lines.append(" ".join(buffer))
            buffer, buf_dur = [], 0.0
    if buffer:
        lines.append(" ".join(buffer))
    return "\n\n".join(lines)


def sanitize_filename(name: str, max_len: int = 80) -> str:
    name = re.sub(r'[\\/*?:"<>|#%&{}\$!\'@+`=]', "", name)
    name = re.sub(r"\s+", "_", name.strip())
    return name[:max_len]


# ─── PDF BUILDER ──────────────────────────────────────────────────────────────

def save_as_pdf(output_path: Path, video: dict, transcript: str, source: str, folder: str):
    doc = SimpleDocTemplate(str(output_path), pagesize=letter,
                            rightMargin=inch, leftMargin=inch,
                            topMargin=inch, bottomMargin=inch)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("T", parent=styles["Title"], fontSize=18,
                                 spaceAfter=6, textColor=colors.HexColor("#FF6600"))
    meta_style  = ParagraphStyle("M", parent=styles["Normal"], fontSize=9,
                                 textColor=colors.HexColor("#888888"), spaceAfter=4)
    body_style  = ParagraphStyle("B", parent=styles["Normal"], fontSize=11,
                                 leading=16, spaceAfter=10)

    def safe(t):
        return (t or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

    story = [Paragraph(safe(video["title"]), title_style)]

    meta_items = [
        f"<b>URL:</b> {safe(video['url'])}",
        f"<b>Channel:</b> {safe(video.get('channel','Y Combinator'))}",
        f"<b>Folder:</b> {safe(folder)}",
    ]
    if video.get("playlist_title"):
        meta_items.append(f"<b>Playlist:</b> {safe(video['playlist_title'])}")
    if video.get("upload_date"):
        d = video["upload_date"]
        meta_items.append(f"<b>Published:</b> {d[:4]}-{d[4:6]}-{d[6:] if len(d)==8 else d}")
    if video.get("duration"):
        m, s = divmod(int(video["duration"]), 60)
        meta_items.append(f"<b>Duration:</b> {m}m {s}s")
    meta_items.append(f"<b>Transcript source:</b> {source}")
    meta_items.append(f"<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    for item in meta_items:
        story.append(Paragraph(item, meta_style))

    story += [Spacer(1, 8),
              HRFlowable(width="100%", thickness=1, color=colors.HexColor("#FF6600")),
              Spacer(1, 12)]

    if transcript:
        for para in transcript.split("\n\n"):
            para = para.strip()
            if para:
                story.append(Paragraph(safe(para), body_style))
    else:
        story.append(Paragraph("<i>No transcript available for this video.</i>", body_style))

    doc.build(story)


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  YC Transcript Downloader — Excel Mode")
    print("=" * 60)

    # Get Excel path
    if len(sys.argv) > 1:
        xlsx_path = Path(sys.argv[1])
    else:
        raw = input("\nPath to Excel file (e.g. yc_urls_template.xlsx): ").strip().strip('"')
        xlsx_path = Path(raw)

    if not xlsx_path.exists():
        print(f"❌  File not found: {xlsx_path}")
        return

    # Derive output dir from Excel location — works on local, Colab + Google Drive
    OUTPUT_DIR = xlsx_path.parent / "yc_transcripts"
    OUTPUT_DIR.mkdir(exist_ok=True)
    log.info(f"Output directory: {OUTPUT_DIR}")

    # Load rows
    wb, ws, rows = load_rows(xlsx_path)
    log.info(f"Loaded {len(rows)} rows from {xlsx_path.name}")

    # Detect duplicates up front
    seen_urls = {}
    for r in rows:
        url = r["url"].lower()
        if url in seen_urls:
            r["is_duplicate"] = True
        else:
            seen_urls[url] = True
            r["is_duplicate"] = False

    # Counters
    counts = {S_DONE: 0, S_FAILED: 0, S_DUPLICATE: 0, "skipped_done": 0}

    for idx, row in enumerate(rows, 1):
        url     = row["url"]
        folder  = row["folder"]
        status  = row["status"]
        row_num = row["row_num"]

        print(f"\n[{idx}/{len(rows)}] {url[:70]}")

        # Skip duplicates
        if row["is_duplicate"]:
            log.info("  → Duplicate URL, skipping")
            update_row(ws, row_num, S_DUPLICATE, "Same URL already in sheet")
            wb.save(xlsx_path)
            counts[S_DUPLICATE] += 1
            continue

        # Skip already done
        if status == S_DONE:
            log.info("  → Already done, skipping")
            counts["skipped_done"] += 1
            continue

        # ── Process this URL ──────────────────────────────────────────────────

        # 1. Extract video ID
        try:
            video_id = extract_video_id(url)
        except ValueError as e:
            log.warning(f"  → {e}")
            update_row(ws, row_num, S_FAILED, str(e), datetime.now().strftime("%Y-%m-%d %H:%M"))
            wb.save(xlsx_path)
            counts[S_FAILED] += 1
            continue

        # 2. Metadata
        log.info("  → Fetching metadata...")
        video = get_video_metadata(url)
        log.info(f"  → Title: {video['title']}")

        # 3. Transcript
        log.info("  → Fetching transcript...")
        transcript, source = fetch_transcript(video_id)

        if not transcript:
            log.warning(f"  → No transcript ({source})")
            update_row(ws, row_num, S_FAILED, source, datetime.now().strftime("%Y-%m-%d %H:%M"))
            wb.save(xlsx_path)
            counts[S_FAILED] += 1
            continue

        log.info(f"  → Transcript OK ({len(transcript)} chars, {source})")

        # 4. Save PDF
        folder_path = OUTPUT_DIR / sanitize_filename(folder)
        folder_path.mkdir(parents=True, exist_ok=True)
        pdf_name = sanitize_filename(video["title"]) + ".pdf"
        pdf_path = folder_path / pdf_name

        try:
            save_as_pdf(pdf_path, video, transcript, source, folder)
            log.info(f"  → PDF saved: {pdf_path}")
            update_row(ws, row_num, S_DONE, "", datetime.now().strftime("%Y-%m-%d %H:%M"))
            counts[S_DONE] += 1
        except Exception as e:
            log.error(f"  → PDF save failed: {e}")
            update_row(ws, row_num, S_FAILED, f"PDF error: {e}", datetime.now().strftime("%Y-%m-%d %H:%M"))
            counts[S_FAILED] += 1

        # Save Excel after every row — safe against interruptions
        wb.save(xlsx_path)

    # ── Summary ───────────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  COMPLETE")
    print(f"  ✅  Done          : {counts[S_DONE]}")
    print(f"  ❌  Failed        : {counts[S_FAILED]}")
    print(f"  ⏭️   Duplicates    : {counts[S_DUPLICATE]}")
    print(f"  ⏩  Already done  : {counts['skipped_done']}")
    print(f"  📁  Output folder : {OUTPUT_DIR.resolve()}")
    print("=" * 60)


if __name__ == "__main__":
    main()
