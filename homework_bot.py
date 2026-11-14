import docx
import os
import sys
import re
import io
import time
import json
import urllib.request
from datetime import datetime
from pypdf import PdfReader
import logging
logging.getLogger("pypdf").setLevel(logging.ERROR)
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from google import genai
from pydantic import BaseModel


# =====================================================================
# 1. CONFIGURATION (Keep your API Key and URLs here!)
# =====================================================================

# Reads Gemini key from local gemini_key.txt (on PC) or GitHub Secrets (in cloud)
GEMINI_KEY_FILE = "gemini_key.txt"
if os.path.exists(GEMINI_KEY_FILE):
    with open(GEMINI_KEY_FILE, "r", encoding="utf-8") as f:
        GEMINI_API_KEY = f.read().strip()
else:
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# You can mix regular course URLs, Canvas PDF links, and local .pdf filenames here:
COURSE_URLS = [
    "https://m3a.vhlcentral.com/courses/1582211/sections/2052076",          # Spanish VHL Vistas portal
    "https://franciscan.instructure.com/courses/14917/modules",             # Philosophy of God Modules
    "https://franciscan.instructure.com/courses/16213",                     # Trinity and Christology API
    "https://franciscan.instructure.com/courses/16213/files/1877796?module_item_id=900256", # T&C Word Syllabus
    "UpdatedReadingSchedule.pdf",                                           # Foundations of Ethics (PHIL 212) Syllabus
    "https://franciscan.instructure.com/courses/15459/files/1832962?module_item_id=889082", # Christian Moral Principles Syllabus PDF
    "https://franciscan.instructure.com/courses/15459",                     # Christian Moral Principles Course Page (for Exams/Quizzes)
    "https://franciscan.instructure.com/courses/15604/files/1832338?module_item_id=888658", # Spanish Canvas Syllabus PDF
    "https://franciscan.instructure.com",                                   # Canvas Dashboard
]
DISCORD_WEBHOOK_URL = ""
AUTH_FILE = "auth.json"
OUTPUT_FILE = "homework_list.md"

# =====================================================================
# 2. DATA STRUCTURE FOR THE AI
# =====================================================================

class Assignment(BaseModel):
    course: str
    title: str
    due_date: str
    status_or_details: str

class HomeworkList(BaseModel):
    assignments: list[Assignment]

# =====================================================================
# 3. LOGIN SAVER (Bypasses 2FA on future runs)
# =====================================================================

def save_logins():
    print("Opening browser... Log into all of your school tabs that open.")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        
        for url in COURSE_URLS:
            if not url.startswith("http"):
                continue  # Skip local PDF files during login
            page = context.new_page()
            try:
                page.goto(url)
            except Exception as e:
                print(f"Could not open {url}: {e}")
                
        input("\n>>> Once you are logged into ALL tabs and can see your courses, press ENTER here to save... ")
        context.storage_state(path=AUTH_FILE)
        browser.close()
        print(f"Success! Login session saved to {AUTH_FILE}.")
# =====================================================================
# 4. SCRAPER, PDF READER & AI PARSER
# =====================================================================

def extract_pdf_bytes(pdf_bytes: bytes) -> str:
    """Extracts text from raw PDF data page by page."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages_text = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if text:
            pages_text.append(f"[Syllabus PDF Page {i+1}]\n{text}")
    return "\n".join(pages_text)

def extract_docx_bytes(file_bytes: bytes) -> str:
    """Extracts text from Word .docx file bytes."""
    doc = docx.Document(io.BytesIO(file_bytes))
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_text:
                full_text.append(" | ".join(row_text))
    return "\n".join(full_text)
def fetch_canvas_course_via_api(base_url: str, course_id: str, context) -> tuple[str, int]:
    """Uses your logged-in session to pull modules, syllabus body, pages, & assignments from Canvas's API."""
    output = []
    
    # 1. Get Course Name & Syllabus Body (if the professor posted a reading schedule on the Syllabus tab)
    course_name = f"Course {course_id}"
    syllabus_text = ""
    c_resp = context.request.get(f"{base_url}/api/v1/courses/{course_id}?include[]=syllabus_body")
    if c_resp.ok:
        c_data = json.loads(c_resp.text().removeprefix("while(1);"))
        course_name = c_data.get("name", course_name)
        raw_syllabus = c_data.get("syllabus_body") or ""
        if raw_syllabus:
            s_soup = BeautifulSoup(raw_syllabus, "html.parser")
            syllabus_text = s_soup.get_text(separator="\n", strip=True)

    output.append(f"=== CANVAS COURSE: {course_name} (ID: {course_id}) ===")
    if syllabus_text:
        print(f"  -> Extracted {len(syllabus_text)} characters from Canvas Syllabus tab for '{course_name}'!")
        output.append(f"\n=== COURSE SYLLABUS TAB FOR {course_name} ===\n{syllabus_text[:30000]}")

    # 2. Paginate through ALL Modules & Module Items
    modules = []
    for page_num in range(1, 4):
        m_url = f"{base_url}/api/v1/courses/{course_id}/modules?include[]=items&per_page=100&page={page_num}"
        m_resp = context.request.get(m_url)
        if not m_resp.ok:
            break
        batch = json.loads(m_resp.text().removeprefix("while(1);"))
        if not batch or not isinstance(batch, list):
            break
        modules.extend(batch)
        if len(batch) < 100:
            break

    print(f"  -> Extracted {len(modules)} total modules directly from Canvas API for '{course_name}'!")
    for idx, mod in enumerate(modules, 1):
        mod_name = mod.get("name", "Unnamed Module")
        output.append(f"\n[MODULE #{idx} TITLE / HEADER]: {mod_name}")
        items = mod.get("items", [])
        # If Canvas omitted items inline, fetch this module's items directly
        if not items and mod.get("id"):
            it_resp = context.request.get(f"{base_url}/api/v1/courses/{course_id}/modules/{mod['id']}/items?per_page=100")
            if it_resp.ok:
                items = json.loads(it_resp.text().removeprefix("while(1);"))
        for item in items:
            item_title = item.get("title", "")
            item_type = item.get("type", "")
            due_at = item.get("content_details", {}).get("due_at", "")
            output.append(f"  - [MODULE SUB-ITEM ({item_type})]: {item_title} {f'(Due: {due_at})' if due_at else ''}")

    # 3. Also pull Upcoming/Dated Assignments from Canvas Assignments API
    a_url = f"{base_url}/api/v1/courses/{course_id}/assignments?per_page=100&order_by=due_at"
    a_resp = context.request.get(a_url)
    if a_resp.ok:
        assigns = json.loads(a_resp.text().removeprefix("while(1);"))
        if isinstance(assigns, list) and assigns:
            output.append(f"\n=== FORMAL ASSIGNMENTS LIST FOR {course_name} ===")
            for a in assigns:
                output.append(f"- [ASSIGNMENT]: {a.get('name')} | Due: {a.get('due_at')} | Points: {a.get('points_possible')}")

    return "\n".join(output), len(modules)


def fetch_clean_text(url: str, context) -> str:
    # 1. Local files in project folder
    if not url.startswith("http") and os.path.exists(url):
        if url.lower().endswith(".pdf"):
            print(f"Reading local syllabus PDF: {url} ...")
            try:
                with open(url, "rb") as f:
                    text = extract_pdf_bytes(f.read())
                    print(f"  -> Extracted {len(text)} characters from local PDF '{url}'!")
                    return f"=== SYLLABUS SCHEDULE PDF: Foundations of Ethics (PHIL 212) ===\n{text}"
            except Exception as e:
                print(f"Warning: Could not read local PDF {url} ({e})")
                return ""
        elif url.lower().endswith(".docx"):
            print(f"Reading local syllabus Word doc: {url} ...")
            try:
                with open(url, "rb") as f:
                    return extract_docx_bytes(f.read())
            except Exception as e:
                print(f"Warning: Could not read local docx {url} ({e})")
                return ""

    print(f"Checking: {url} ...")

    course_hint = ""
    if "14917" in url:
        course_hint = " (Philosophy of God)"
    elif "16213" in url:
        course_hint = " (Trinity and Christology)"
    elif "15459" in url:
        course_hint = " (Christian Moral Principles / THE-115)"
    elif "15604" in url or "vhlcentral" in url:
        course_hint = " (Elementary Spanish II / SPN 102)"

    # 2. Canvas file URLs or direct document links
    canvas_file_match = re.search(r"(https?://[^/]+/(?:courses/\d+/)?files/\d+)", url)
    is_direct_doc = url.lower().split("?")[0].endswith((".pdf", ".docx", ".doc"))

    if canvas_file_match or is_direct_doc:
        download_url = f"{canvas_file_match.group(1)}/download?download_frd=1" if canvas_file_match else url
        print(f"  -> Downloading course document from: {download_url} ...")
        try:
            response = context.request.get(download_url, timeout=20000)
            if response.ok:
                raw_bytes = response.body()
                if raw_bytes.startswith(b"PK\x03\x04"):
                    print("  -> Detected Word document (.docx). Extracting text...")
                    doc_text = extract_docx_bytes(raw_bytes)
                    print(f"  -> Extracted {len(doc_text)} characters from Word document{course_hint}!")
                    return f"=== COURSE SYLLABUS DOCUMENT{course_hint} ===\n{doc_text}"
                elif raw_bytes.startswith(b"%PDF"):
                    print("  -> Detected PDF document. Extracting text...")
                    pdf_text = extract_pdf_bytes(raw_bytes)
                    print(f"  -> Extracted {len(pdf_text)} characters from PDF{course_hint}!")
                    return f"=== COURSE SYLLABUS PDF{course_hint} ===\n{pdf_text}"
                else:
                    try:
                        return f"=== COURSE SYLLABUS PDF{course_hint} ===\n{extract_pdf_bytes(raw_bytes)}"
                    except Exception:
                        return f"=== COURSE SYLLABUS DOCUMENT{course_hint} ===\n{extract_docx_bytes(raw_bytes)}"
        except Exception as e:
            print(f"  -> Document download failed ({e}), falling back to browser view...")

    # 3. Direct Canvas Course API (also falls through to browser scrape if 0 modules were returned)
    api_text = ""
    canvas_course_match = re.match(r"^(https?://[^/]+instructure\.com)/courses/(\d+)(?:/modules|/assignments|/?)?$", url.split("?")[0])
    if canvas_course_match:
        try:
            api_text, mod_count = fetch_canvas_course_via_api(canvas_course_match.group(1), canvas_course_match.group(2), context)
            if mod_count > 0:
                return api_text
            print(f"  -> 0 API modules for course {canvas_course_match.group(2)}{course_hint}; also scraping browser page...")
        except Exception as e:
            print(f"  -> Canvas API check failed ({e}), falling back to browser scrape...")

    # 4. Standard webpage and module scraping (for VHL Central, Canvas Modules/Home, Dashboard)
    is_vhl = "vhlcentral.com" in url.lower()
    page = context.new_page()
    try:
        if is_vhl:
            try:
                page.goto(url, wait_until="networkidle", timeout=30000)
            except Exception:
                page.goto(url, wait_until="domcontentloaded", timeout=25000)
            page.wait_for_timeout(5000)
        else:
            try:
                page.goto(url, wait_until="networkidle", timeout=25000)
            except Exception:
                page.goto(url, wait_until="domcontentloaded", timeout=20000)
            page.wait_for_timeout(3500)

        print("  -> Scrolling down through page...")
        scroll_height = page.evaluate("document.body.scrollHeight")
        current_position = 0
        while current_position < scroll_height:
            current_position += 800
            page.evaluate(f"window.scrollTo(0, {current_position})")
            page.wait_for_timeout(400)
            scroll_height = page.evaluate("document.body.scrollHeight")

        page.wait_for_timeout(1500)
        final_url = page.url.lower()
        html = page.content()
    except Exception as e:
        print(f"Warning: Issue loading {url} ({e})")
        final_url = ""
        html = page.content() if page else ""
    finally:
        page.close()

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "svg", "noscript", "footer", "nav"]):
        tag.decompose()
    clean_text = soup.get_text(separator="\n", strip=True)

    if is_vhl:
        if "sso.vhlcentral.com" in final_url or "login" in final_url or "sign_in" in final_url or len(clean_text) < 300:
            print("  -> ⚠️ WARNING: VHL Central session expired! Run `python homework_bot.py --login` to refresh auth.json.")
        else:
            print(f"  -> Extracted {len(clean_text)} characters from VHL Central!")
        return f"=== VHL CENTRAL ONLINE ASSIGNMENTS (Elementary Spanish II - Vistas) ===\n{clean_text}"

    print(f"  -> Extracted {len(clean_text)} characters from webpage{course_hint}!")
    if api_text:
        return f"{api_text}\n\n=== BROWSER PAGE SCRAPE{course_hint} ===\n{clean_text}"
    return f"=== BROWSER PAGE SCRAPE{course_hint} ===\n{clean_text}"


def parse_homework_with_ai(combined_text: str) -> HomeworkList | None:
    client = genai.Client(
        api_key=GEMINI_API_KEY,
        http_options={"timeout": 60000}
    )
    today_str = datetime.now().strftime("%A, %B %d, %Y")

    trimmed_text = combined_text[:350000]
    print(f"Payload size: {len(trimmed_text)} characters. Sending request to Gemini...")

    prompt = f"""
    Today's date is {today_str}.
    Below is text scraped from 5 college courses across Canvas API module/assignment lists, VHL Central, and course syllabi (PDF and Word docs).
    
    CRITICAL INSTRUCTIONS:
    1. EXTRACT ASSIGNMENTS & DAILY READINGS FOR ALL 5 COURSES (DO NOT SKIP ANY COURSE):
       You MUST include upcoming items for TODAY ({today_str}) and the NEXT 14 DAYS from ALL FIVE of these courses:
       - **"Philosophy of God"** (Course 14917): The professor puts class readings and homework topics directly inside the `[MODULE #... TITLE / HEADER]` titles, module sub-items, or course page scrape. Extract the upcoming modules/readings for today and the next 14 days! (If a module header has a date or class number corresponding to the current/upcoming two weeks, include it as an assignment!)
       - **"Christian Moral Principles"** (Course 15459 / THE-115): Look inside `=== COURSE SYLLABUS PDF (Christian Moral Principles / THE-115) ===` AND the course schedule. You MUST extract every daily class reading assignment scheduled for today and the next 14 days (set `due_date` to the class date the reading is assigned for!), in addition to any online Canvas assessments/exams!
       - **"Foundations of Ethics"** (PHIL 212): Extract all daily readings (e.g., John Duns Scotus) and essays scheduled for today and the next 14 days from `UpdatedReadingSchedule.pdf`.
       - **"Trinity and Christology"** (Course 16213): Merge the Word syllabus reading details with the exact Canvas `Journal for Class #` due dates.
       - **"Elementary Spanish II"** (Course 15604 / VHL): Include BOTH the online VHL Central Vistas activities AND the in-class syllabus assignments (`Escritura`, `Pruebas`, etc.).
    2. CLASS READINGS ARE REQUIRED HOMEWORK: Even if a syllabus reading does not have an online submission box on Canvas, you MUST include it as a row in `assignments` with its class date as `due_date`!
    3. ACTUAL DUE DATES ONLY: If a syllabus row on an earlier date mentions a future paper or exam (e.g. 'Essay 2 due 11/19'), set its `due_date` to the actual submission date (2026-11-19), NOT the earlier reminder date!
    4. STRICT DATE FORMAT: Always format `due_date` starting with `YYYY-MM-DD` followed by the time if known (for example: `2026-11-19, 08:00 AM` or `2026-11-15, 11:59 PM` or `2026-11-17`).
    5. Ignore readings or assignments whose dates have already passed before {today_str}, or items marked 'Submitted'/'Completed'.
    6. Sort the final list chronologically by due date.
    
    Return ONLY a valid JSON object matching this structure:
    {{
      "assignments": [
        {{
          "course": "Exact Course Name",
          "title": "Assignment, Reading, or Module Title",
          "due_date": "YYYY-MM-DD, HH:MM AM/PM (or YYYY-MM-DD if no time)",
          "status_or_details": "Pages to read, module sub-items, or syllabus notes"
        }}
      ]
    }}

    SCRAPED CONTENT:
    {trimmed_text}
    """
    
    models_to_try = [
        "gemini-3.5-flash-lite",
        "gemini-3.8-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
    ]
    
    for attempt_round in range(1, 3):
        for model_name in models_to_try:
            try:
                print(f"Trying model: {model_name} (Pass {attempt_round}/2)...")
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config={
                        "response_mime_type": "application/json",
                    },
                )
                print(f"Response received from {model_name}! Formatting output...")
                
                raw_json = response.text.strip()
                if raw_json.startswith("```"):
                    raw_json = raw_json.strip("`").replace("json\n", "", 1)

                data = json.loads(raw_json)
                return HomeworkList(**data)

            except Exception as e:
                print(f"Warning: {model_name} encountered an issue ({e}). Retrying in 4s...")
                time.sleep(4)
        
        if attempt_round == 1:
            print("Models are busy (503). Waiting 10 seconds before second pass...")
            time.sleep(10)
            
    print("\nAll models were temporarily unavailable due to high demand.")
    return None


# =====================================================================
# 5. OUTPUT & NOTIFICATIONS
# =====================================================================

def format_markdown(hw_list: HomeworkList) -> str:
    today_str = datetime.now().strftime("%B %d, %Y")
    lines = [f"# Upcoming Homework ({today_str})\n"]
    
    if not hw_list.assignments:
        lines.append("🎉 No upcoming assignments found! (Or your login cookies in `auth.json` expired—run `python homework_bot.py --login` to refresh).")
        return "\n".join(lines)
        
    lines.append("| Course | Assignment | Due Date | Details |")
    lines.append("| :--- | :--- | :--- | :--- |")
    for item in hw_list.assignments:
        lines.append(f"| **{item.course}** | {item.title} | {item.due_date} | {item.status_or_details} |")
    return "\n".join(lines)

def send_to_discord(markdown_text: str):
    if not DISCORD_WEBHOOK_URL:
        return
    payload = json.dumps({"content": markdown_text[:1990]}).encode("utf-8")
    req = urllib.request.Request(
        DISCORD_WEBHOOK_URL, 
        data=payload, 
        headers={"Content-Type": "application/json", "User-Agent": "HomeworkBot/1.0"}
    )
    try:
        urllib.request.urlopen(req)
        print("Sent summary to Discord!")
    except Exception as e:
        print(f"Failed to send Discord notification: {e}")

def main():
    if "--login" in sys.argv or not os.path.exists(AUTH_FILE):
        if not os.path.exists(AUTH_FILE):
            print("No saved login session found. Starting first-time login setup...")
        save_logins()
        return

    combined_text = ""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(storage_state=AUTH_FILE)

        # Warm up Canvas session first (constructed via concatenation so copy-paste never adds markdown brackets)
        try:
            print("Warming up Canvas session...")
            warmup_page = context.new_page()
            canvas_base = "https://" + "franciscan.instructure.com"
            warmup_page.goto(canvas_base, wait_until="domcontentloaded", timeout=20000)
            warmup_page.wait_for_timeout(3000)
            warmup_page.close()
        except Exception as e:
            print(f"Canvas warmup note: {e}")

        for url in COURSE_URLS:
            page_text = fetch_clean_text(url, context)
            # Put syllabus PDFs and specific course pages at full length, cap generic dashboard at 20k
            limit = 20000 if url.rstrip("/").endswith("instructure.com") else 65000
            combined_text += f"\n\n--- SOURCE: {url} ---\n{page_text[:limit]}"
            
        browser.close()

    hw_data = parse_homework_with_ai(combined_text)
    
    # SAFEGUARD: Never overwrite homework_list.md if the AI failed or returned 0 items when a file already exists!
    if hw_data is None or (not hw_data.assignments and os.path.exists(OUTPUT_FILE)):
        print(f"\nKeeping existing '{OUTPUT_FILE}' intact so your app doesn't lose its assignments!")
    else:
        report = format_markdown(hw_data)
        print("\n" + report + "\n")
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(report)
        print(f"Saved checklist to {OUTPUT_FILE}")
        send_to_discord(report)

    import launch_app
    launch_app.build_and_open_app(launch_app.parse_markdown_file(OUTPUT_FILE))

if __name__ == "__main__":
    main()
