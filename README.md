# 🎯 AI Job Hunting Machine — GitHub Actions (Zero Budget)

Runs every day at **10 AM**, reads your resume, searches for **real live DevOps
jobs in India**, scores them against your profile, and emails you a rich Excel
workbook — automatically, for free.

---

## 📁 Repo Structure

```
.
├── .github/
│   └── workflows/
│       └── daily_report.yml   ← Scheduler + orchestration
├── scripts/
│   ├── generate_excel.py      ← AI job search + Excel builder
│   └── send_email.py          ← Gmail SMTP sender
├── resume.pdf                 ← ⭐ YOUR RESUME — add this!
├── prompt.txt                 ← The AI prompt (edit to customise)
└── README.md
```

---

## ⚡ What You Get (Daily Email)

An Excel workbook with **6 sheets**:

| Sheet | Contents |
|-------|----------|
| 📊 Dashboard | Candidate summary, strengths, gaps, stats |
| 🗂 All Jobs | All roles, sorted by fit score, deduplicated |
| 🟢 High Probability | Fit score ≥ 75 — apply here first |
| 🟡 Medium Probability | Fit score 50–74 |
| 🔴 Stretch Roles | Fit score < 50 — aspirational targets |
| 📅 30-Day Resume Tips | Day-by-day tailored resume improvement plan |

---

## 🛠 One-Time Setup (~10 minutes)

### Step 1 — Add Your Resume

Copy your resume into the repo root and name it `resume.pdf` (or `resume.txt`):

```bash
cp /path/to/your/CV.pdf resume.pdf
git add resume.pdf
git commit -m "Add resume"
git push
```

### Step 2 — Free Anthropic API Key

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign up → API Keys → **Create Key**
3. Copy the key (`sk-ant-...`)

> Free tier is enough for one run per day.

### Step 3 — Gmail App Password

> ⚠️ Use an **App Password**, NOT your real Gmail password.

1. Google Account → **Security**
2. Enable **2-Step Verification**
3. Search **"App passwords"** → Create → Name it "Job Hunt Bot"
4. Copy the 16-character code (e.g. `abcd efgh ijkl mnop`)

### Step 4 — GitHub Secrets

Repo → **Settings → Secrets and variables → Actions → New repository secret**

| Secret | Value |
|--------|-------|
| `ANTHROPIC_API_KEY` | Your Anthropic key (`sk-ant-...`) |
| `GMAIL_USER` | `you@gmail.com` |
| `GMAIL_APP_PASS` | 16-char App Password (no spaces) |
| `EMAIL_TO` | `you@gmail.com` (comma-separate for multiple) |

### Step 5 — Adjust Timezone (Optional)

The default cron fires at **10:00 AM IST**. Change in `daily_report.yml`:

| Target | Timezone | Cron |
|--------|----------|------|
| 10:00 AM IST | UTC+5:30 | `30 4 * * *` |
| 10:00 AM GMT | UTC+0 | `0 10 * * *` |
| 10:00 AM EST | UTC-5 | `0 15 * * *` |
| 10:00 AM PST | UTC-8 | `0 18 * * *` |

### Step 6 — Test It Now

GitHub → **Actions → 🎯 Daily Job Hunt → Run workflow**

Don't wait for 10 AM — test immediately with the manual trigger.

---

## 💡 Cost Breakdown — Truly Zero

| Component | Service | Cost |
|-----------|---------|------|
| Scheduler | GitHub Actions cron | Free |
| Runner | `ubuntu-latest` (2000 min/month free for public) | Free |
| AI model | Anthropic Claude API (free tier) | Free* |
| Web search | Built into Claude API | Free* |
| Excel | `openpyxl` Python library | Free |
| Email | Gmail SMTP | Free |
| Backup storage | GitHub Actions Artifacts (500MB free) | Free |

*Free tier credits. One daily run uses ~2000–4000 tokens, well within limits.

---

## 🔧 Troubleshooting

**Workflow doesn't trigger at 10 AM?**
GitHub Actions cron can be up to 15 min late under load. Also: repos with no pushes for 60 days have scheduled workflows disabled — just trigger manually once to re-activate.

**`FileNotFoundError: No resume found`**
Make sure `resume.pdf` (or `resume.txt`) is committed to the repo root.

**`json.JSONDecodeError`**
Claude returned non-JSON. Usually happens if the prompt is ambiguous. Try running manually — the error log shows Claude's raw response.

**Gmail authentication failed**
Re-generate the App Password — they expire if 2FA settings change.

**Public repo — is my resume safe?**
If the repo is public, your resume will be visible. Use a **private repo** instead. GitHub gives unlimited private repos for free.
