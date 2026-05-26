"""
generate_excel.py  —  Job Hunting Machine (OpenRouter Edition)
--------------------------------------------------------------
Split into TWO API calls to avoid free-model timeouts:
  Call 1 -> candidate analysis + job listings (JSON)
  Call 2 -> 30-day resume tips (JSON)
Then merges both into one Excel workbook.
"""

import os, json, datetime, glob, sys, base64
import urllib.request, urllib.error
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── styling helpers ───────────────────────────────────────────────────────────

def hfill(hex_color):
    return PatternFill("solid", fgColor=hex_color.lstrip("#"))

def bfont(size=11, color="FFFFFF", bold=True, italic=False):
    return Font(bold=bold, italic=italic, size=size, color=color.lstrip("#"))

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def left(wrap=True):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

PROB_COLORS = {
    "High":    ("1B5E20", "C8E6C9"),
    "Medium":  ("E65100", "FFE0B2"),
    "Stretch": ("B71C1C", "FFCDD2"),
}

SCORE_FILL = {
    (75, 101): "C8E6C9",
    (50,  75): "FFE0B2",
    (0,   50): "FFCDD2",
}

def score_fill(score):
    for (lo, hi), color in SCORE_FILL.items():
        if lo <= score < hi:
            return hfill(color)
    return hfill("FFFFFF")

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def header_row(ws, headers, fill_hex, font_color="FFFFFF"):
    ws.append(headers)
    row = ws.max_row
    for col in range(1, len(headers)+1):
        c = ws.cell(row=row, column=col)
        c.fill      = hfill(fill_hex)
        c.font      = bfont(10, font_color)
        c.alignment = center(wrap=True)
        c.border    = thin_border()
    ws.row_dimensions[row].height = 22

# ── resume reader ─────────────────────────────────────────────────────────────

def read_resume():
    for pattern in ["resume.pdf", "resume.PDF", "resume.txt", "cv.pdf", "cv.PDF", "cv.txt"]:
        matches = glob.glob(pattern)
        if matches:
            path = matches[0]
            if path.lower().endswith(".pdf"):
                with open(path, "rb") as f:
                    data = f.read()
                return None, "application/pdf", base64.b64encode(data).decode()
            else:
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    return f.read(), "text/plain", None
    raise FileNotFoundError("No resume found! Add resume.pdf or resume.txt to the repo root.")

# ── PDF text extractor ────────────────────────────────────────────────────────

def extract_pdf_text(resume_b64):
    try:
        from pdfminer.high_level import extract_text_to_fp
        from pdfminer.layout import LAParams
        import io
        pdf_bytes = base64.b64decode(resume_b64)
        output = io.StringIO()
        extract_text_to_fp(io.BytesIO(pdf_bytes), output, laparams=LAParams())
        text = output.getvalue().strip()
        if text:
            print(f"✅ PDF extracted ({len(text)} chars)")
            return text
    except Exception as e:
        print(f"⚠️  pdfminer failed: {e}")
    return None

# ── OpenRouter API call ───────────────────────────────────────────────────────

FREE_MODELS = [
    "openrouter/free",
    "deepseek/deepseek-v4-flash:free",
    "nvidia/nemotron-3-super-120b-a12b:free",
    "meta-llama/llama-3.3-70b-instruct:free",
    "google/gemma-4-31b-it:free",
]

def call_openrouter(prompt_text, call_label=""):
    api_key = os.environ["OPENROUTER_API_KEY"]
    url     = "https://openrouter.ai/api/v1/chat/completions"

    for model in FREE_MODELS:
        print(f"📡 [{call_label}] Trying {model} ...")
        payload = json.dumps({
            "model": model,
            "messages": [{"role": "user", "content": prompt_text}],
            "temperature": 0.3,
            "max_tokens": 8000,
            "response_format": {"type": "json_object"}
        }).encode("utf-8")

        req = urllib.request.Request(
            url, data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
                "HTTP-Referer": "https://github.com/job-hunt-bot",
                "X-Title": "Job Hunt Bot"
            },
            method="POST"
        )

        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                result = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8")
            print(f"⚠️  {model} HTTP {e.code}: {body[:200]} — trying next...")
            continue
        except Exception as e:
            print(f"⚠️  {model} error: {e} — trying next...")
            continue

        if "error" in result:
            print(f"⚠️  {model} upstream error: {str(result['error'])[:200]} — trying next...")
            continue

        raw = result["choices"][0]["message"]["content"].strip()

        # Print first 500 chars of raw response for debugging
        print(f"🔍 Raw response preview: {raw[:500]}")

        # Strip markdown fences
        if "```" in raw:
            parts = raw.split("```")
            for part in parts:
                cleaned = part.strip()
                if cleaned.startswith("json"):
                    cleaned = cleaned[4:].strip()
                if cleaned.startswith("{"):
                    raw = cleaned
                    break

        try:
            parsed = json.loads(raw.strip())
            # Validate it has actual data — not an empty shell
            jobs = parsed.get("jobs", [])
            tips = parsed.get("daily_resume_tips", [])
            if call_label == "JOBS" and len(jobs) == 0:
                print(f"⚠️  {model} returned 0 jobs — trying next model...")
                continue
            if call_label == "TIPS" and len(tips) == 0:
                print(f"⚠️  {model} returned 0 tips — trying next model...")
                continue
            print(f"✅ [{call_label}] Success with {model} — jobs:{len(jobs)} tips:{len(tips)}")
            return parsed
        except json.JSONDecodeError as e:
            print(f"⚠️  {model} JSON parse error: {e} — trying next...")
            continue

    print(f"❌ All models failed for [{call_label}].")
    sys.exit(1)

# ── prompt builders ───────────────────────────────────────────────────────────

def build_jobs_prompt(resume_text, today, target_date):
    return f"""You are an expert AI recruiter. Today: {today}. Target date: {target_date}.

CANDIDATE RESUME:
{resume_text}

Find 25 real job opportunities in India for Associate/Mid-level DevOps Engineer (3-5 years experience).
Target: startups, scale-ups, MNCs, consulting firms, tech and non-tech sectors.
For each job assign a fit_score out of 100 and category: High (>=75), Medium (50-74), Stretch (<50).
Never list the same company+role twice.

Return ONLY a JSON object in this exact schema (no markdown, no explanation):
{{
  "analysis_date": "{today}",
  "target_date": "{target_date}",
  "candidate_summary": "<2-3 sentence profile summary>",
  "resume_strengths": ["strength1", "strength2", "strength3"],
  "resume_gaps": ["gap1", "gap2", "gap3"],
  "jobs": [
    {{
      "rank": 1,
      "company": "Company Name",
      "role": "Exact Job Title",
      "sector": "Tech/Non-Tech/Consulting/Startup/MNC",
      "company_size": "Startup/Scale-up/MNC/Consulting",
      "location": "City, India or Remote",
      "experience_required": "3-5 years",
      "key_skills_required": ["skill1", "skill2", "skill3"],
      "fit_score": 85,
      "probability_category": "High",
      "match_reason": "Why this is a good match in 1-2 sentences",
      "gap_reason": "What is missing in 1 sentence",
      "apply_url": "https://direct-application-url.com",
      "source": "LinkedIn/Company Careers/Naukri",
      "estimated_ctc_lpa": "8-14 LPA"
    }}
  ]
}}"""

def build_tips_prompt(resume_text, today, target_date):
    return f"""You are a professional resume coach. Today: {today}.

CANDIDATE RESUME:
{resume_text}

Generate exactly 30 day-by-day resume improvement tips for a DevOps Engineer job hunt.
Each tip must be specific to THIS candidate's resume, actionable, and include a before/after example.

Return ONLY a JSON object (no markdown, no explanation):
{{
  "daily_resume_tips": [
    {{
      "day": 1,
      "date": "{today}",
      "focus_area": "Summary Section",
      "tip": "Specific actionable tip for this candidate",
      "example_before": "What their resume currently says",
      "example_after": "Improved version"
    }}
  ]
}}

Generate all 30 days. Vary the focus areas across: Summary, Skills, Experience bullets, Quantification, Keywords, Certifications, Projects, LinkedIn, Cover letter, GitHub profile, etc."""

# ── Excel builder ─────────────────────────────────────────────────────────────

JOB_HEADERS = [
    "Rank", "Company", "Role", "Sector", "Size",
    "Location", "Exp. Required", "Key Skills",
    "Fit Score", "Probability", "Match Reason",
    "Gap / Stretch Reason", "Apply URL", "Source", "Est. CTC (LPA)"
]
JOB_WIDTHS = [6, 22, 28, 14, 12, 16, 13, 34, 10, 11, 36, 30, 40, 14, 14]

def write_job_row(ws, job):
    skills = ", ".join(job.get("key_skills_required", []))
    prob   = job.get("probability_category", "Medium")
    score  = int(job.get("fit_score", 0))
    row = [
        job.get("rank", ""), job.get("company", ""), job.get("role", ""),
        job.get("sector", ""), job.get("company_size", ""), job.get("location", ""),
        job.get("experience_required", ""), skills, score, prob,
        job.get("match_reason", ""), job.get("gap_reason", ""),
        job.get("apply_url", ""), job.get("source", ""),
        job.get("estimated_ctc_lpa", "Unknown"),
    ]
    ws.append(row)
    r = ws.max_row
    ws.row_dimensions[r].height = 42
    prob_text_color, prob_bg = PROB_COLORS.get(prob, ("000000", "FFFFFF"))
    for col_idx in range(1, len(row)+1):
        cell = ws.cell(row=r, column=col_idx)
        cell.border    = thin_border()
        cell.alignment = left() if col_idx > 2 else center()
        if col_idx == 9:
            cell.fill = score_fill(score)
            cell.font = bfont(11, "1B1B1B"); cell.alignment = center()
        elif col_idx == 10:
            cell.fill = hfill(prob_bg)
            cell.font = bfont(10, prob_text_color); cell.alignment = center()
        elif col_idx == 13:
            cell.font = Font(color="1565C0", underline="single", size=10)
        else:
            cell.font = Font(size=10, color="1B1B1B")

def build_excel(jobs_data, tips_data, output_path, used_model):
    wb = openpyxl.Workbook()

    today     = jobs_data.get("analysis_date", str(datetime.date.today()))
    target    = jobs_data.get("target_date", "")
    candidate = jobs_data.get("candidate_summary", "")
    strengths = jobs_data.get("resume_strengths", [])
    gaps      = jobs_data.get("resume_gaps", [])
    jobs      = jobs_data.get("jobs", [])
    tips      = tips_data.get("daily_resume_tips", [])

    # Deduplicate by (company, role) — keep highest fit_score
    seen = {}
    for j in jobs:
        key = (j.get("company","").lower(), j.get("role","").lower())
        if key not in seen or j.get("fit_score", 0) > seen[key].get("fit_score", 0):
            seen[key] = j
    jobs = sorted(seen.values(), key=lambda x: -x.get("fit_score", 0))
    for i, j in enumerate(jobs, 1):
        j["rank"] = i

    high    = [j for j in jobs if j.get("probability_category") == "High"]
    medium  = [j for j in jobs if j.get("probability_category") == "Medium"]
    stretch = [j for j in jobs if j.get("probability_category") == "Stretch"]

    # ── Sheet 1: Dashboard ────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 62

    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value = f"Job Hunt Dashboard  —  {today}  to  {target}"
    c.font = bfont(15, "FFFFFF"); c.fill = hfill("1A237E"); c.alignment = center()
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:B2")
    c = ws["A2"]
    c.value = candidate
    c.font = Font(italic=True, size=11, color="37474F")
    c.fill = hfill("E8EAF6"); c.alignment = left()
    ws.row_dimensions[2].height = 30

    ws.append([])
    for label, val, fill_hex in [
        ("Total Opportunities", len(jobs),   "283593"),
        ("High Probability",    len(high),   "1B5E20"),
        ("Medium Probability",  len(medium), "E65100"),
        ("Stretch Roles",       len(stretch),"B71C1C"),
        ("Target Date",         target,      "4A148C"),
        ("AI Model",            model_name,  "546E7A"),
    ]:
        ws.append([label, val])
        r = ws.max_row
        for col in [1, 2]:
            cell = ws.cell(row=r, column=col)
            cell.fill = hfill(fill_hex); cell.font = bfont(11, "FFFFFF")
            cell.alignment = left() if col == 2 else center()
            cell.border = thin_border()
        ws.row_dimensions[r].height = 24

    ws.append([])
    ws.append(["Resume Strengths", ""])
    ws.merge_cells(f"A{ws.max_row}:B{ws.max_row}")
    c = ws.cell(row=ws.max_row, column=1)
    c.font = bfont(12, "1B5E20"); c.fill = hfill("C8E6C9"); c.border = thin_border()
    for s in strengths:
        ws.append(["", f"- {s}"])
        ws.cell(row=ws.max_row, column=2).font = Font(size=11, color="1B1B1B")
        ws.row_dimensions[ws.max_row].height = 20

    ws.append([])
    ws.append(["Resume Gaps", ""])
    ws.merge_cells(f"A{ws.max_row}:B{ws.max_row}")
    c = ws.cell(row=ws.max_row, column=1)
    c.font = bfont(12, "B71C1C"); c.fill = hfill("FFCDD2"); c.border = thin_border()
    for g in gaps:
        ws.append(["", f"- {g}"])
        ws.cell(row=ws.max_row, column=2).font = Font(size=11, color="1B1B1B")
        ws.row_dimensions[ws.max_row].height = 20

    # ── Sheet 2: All Jobs ─────────────────────────────────────────────────────
    ws2 = wb.create_sheet("All Jobs")
    ws2.sheet_view.showGridLines = False; ws2.freeze_panes = "A3"
    ws2.merge_cells(f"A1:{get_column_letter(len(JOB_HEADERS))}1")
    c = ws2["A1"]
    c.value = f"All Opportunities — Sorted by Fit Score  |  {today}"
    c.font = bfont(13); c.fill = hfill("1A237E"); c.alignment = center()
    ws2.row_dimensions[1].height = 30
    header_row(ws2, JOB_HEADERS, "283593")
    for j in jobs: write_job_row(ws2, j)
    set_col_widths(ws2, JOB_WIDTHS)

    # ── Sheets 3-5: Categorised ───────────────────────────────────────────────
    for sheet_title, job_list, fill_title, fill_hdr in [
        ("High Probability",   high,    "1B5E20", "2E7D32"),
        ("Medium Probability", medium,  "E65100", "EF6C00"),
        ("Stretch Roles",      stretch, "B71C1C", "C62828"),
    ]:
        wsc = wb.create_sheet(sheet_title)
        wsc.sheet_view.showGridLines = False; wsc.freeze_panes = "A3"
        wsc.merge_cells(f"A1:{get_column_letter(len(JOB_HEADERS))}1")
        c = wsc["A1"]
        c.value = f"{sheet_title}  |  {len(job_list)} roles  |  {today}"
        c.font = bfont(13); c.fill = hfill(fill_title); c.alignment = center()
        wsc.row_dimensions[1].height = 30
        header_row(wsc, JOB_HEADERS, fill_hdr)
        for j in job_list: write_job_row(wsc, j)
        set_col_widths(wsc, JOB_WIDTHS)

    # ── Sheet 6: 30-Day Tips ──────────────────────────────────────────────────
    ws6 = wb.create_sheet("30-Day Resume Tips")
    ws6.sheet_view.showGridLines = False; ws6.freeze_panes = "A3"
    ws6.merge_cells("A1:F1")
    c = ws6["A1"]
    c.value = f"30-Day Resume Improvement Plan  |  {today} to {target}"
    c.font = bfont(13); c.fill = hfill("4A148C"); c.alignment = center()
    ws6.row_dimensions[1].height = 30
    header_row(ws6, ["Day", "Date", "Focus Area", "Tip (Actionable)", "Before", "After"], "6A1B9A")
    for tip in tips:
        ws6.append([
            tip.get("day", ""), tip.get("date", ""), tip.get("focus_area", ""),
            tip.get("tip", ""), tip.get("example_before", ""), tip.get("example_after", ""),
        ])
        r = ws6.max_row; ws6.row_dimensions[r].height = 44
        fill_hex = "F3E5F5" if tip.get("day", 0) % 2 == 0 else "FFFFFF"
        for col in range(1, 7):
            cell = ws6.cell(row=r, column=col)
            cell.fill = hfill(fill_hex); cell.alignment = left()
            cell.border = thin_border(); cell.font = Font(size=10, color="1B1B1B")
        ws6.cell(row=r, column=1).alignment = center()
    set_col_widths(ws6, [5, 12, 20, 52, 36, 36])

    wb.save(output_path)
    print(f"✅ Excel saved -> {output_path}  ({len(jobs)} jobs, {len(tips)} tips)")

# ── entrypoint ────────────────────────────────────────────────────────────────

def main():
    today       = str(datetime.date.today())
    target_date = str(datetime.date.today() + datetime.timedelta(days=30))

    print("📄 Reading resume...")
    resume_text, media_type, resume_b64 = read_resume()
    if media_type == "application/pdf":
        resume_content = extract_pdf_text(resume_b64) or "[PDF unreadable]"
    else:
        resume_content = resume_text or ""

    print(f"📝 Resume length: {len(resume_content)} chars")

    # ── Call 1: Jobs ──────────────────────────────────────────────────────────
    print("\n🤖 CALL 1: Fetching job listings...")
    jobs_prompt = build_jobs_prompt(resume_content, today, target_date)
    jobs_data   = call_openrouter(jobs_prompt, call_label="JOBS")

    # ── Call 2: Tips ──────────────────────────────────────────────────────────
    print("\n🤖 CALL 2: Generating 30-day resume tips...")
    tips_prompt = build_tips_prompt(resume_content, today, target_date)
    tips_data   = call_openrouter(tips_prompt, call_label="TIPS")

    # ── Build Excel ───────────────────────────────────────────────────────────
    output_path = f"jobs_{today}.xlsx"
    print(f"\n📊 Building Excel workbook -> {output_path}")
    build_excel(jobs_data, tips_data, output_path, used_model)

if __name__ == "__main__":
    main()
