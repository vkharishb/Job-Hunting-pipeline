# 🎯 AI Job Hunting Machine — 100% Free (Gemini Edition)

Runs every day at **10 AM IST**, reads your resume, finds real DevOps jobs in
India, scores them, and emails you a rich Excel workbook — **completely free**.

---

## 📁 Repo Structure

```
.
├── .github/
│   └── workflows/
│       └── daily_report.yml   ← Scheduler
├── scripts/
│   ├── generate_excel.py      ← Gemini AI + Excel builder
│   └── send_email.py          ← Gmail SMTP sender
├── resume.pdf                 ← ⭐ Add your resume here
├── prompt.txt                 ← The AI prompt
└── README.md
```

---

## ⚡ Excel Output (6 Sheets Daily)

| Sheet | Contents |
|-------|----------|
| 📊 Dashboard | Summary, strengths, gaps, counts |
| 🗂 All Jobs | All roles sorted by fit score, deduplicated |
| 🟢 High Probability | Fit ≥ 75 — apply here first |
| 🟡 Medium Probability | Fit 50–74 |
| 🔴 Stretch Roles | Fit < 50 — aspirational |
| 📅 30-Day Resume Tips | One tailored tip per day |

---

## 🛠 One-Time Setup

### Step 1 — Add your resume
```bash
cp your-cv.pdf resume.pdf
git add resume.pdf
git commit -m "Add resume"
git push
```
> Use a **private repo** to keep your resume safe.

### Step 2 — Get free Gemini API key
1. Go to → [aistudio.google.com](https://aistudio.google.com)
2. Sign in with Google
3. Click **"Get API Key"** → **"Create API key"**
4. Copy the key (looks like `AIzaSy...`)
5. **Free tier: 1,500 requests/day — no credit card needed**

### Step 3 — Get Gmail App Password
1. [myaccount.google.com/security](https://myaccount.google.com/security)
2. Enable **2-Step Verification**
3. Search **"App passwords"** → Create → name it `Job Hunt Bot`
4. Copy the 16-character code → remove spaces

### Step 4 — Add GitHub Secrets
Repo → **Settings → Secrets and variables → Actions → New repository secret**

| Secret | Where to get it | Example |
|--------|----------------|---------|
| `GEMINI_API_KEY` | aistudio.google.com | `AIzaSyXXXX...` |
| `GMAIL_USER` | Your Gmail address | `you@gmail.com` |
| `GMAIL_APP_PASS` | Google App Passwords | `abcdefghijklmnop` |
| `EMAIL_TO` | Where to receive report | `you@gmail.com` |

### Step 5 — Test immediately
GitHub → **Actions → 🎯 Daily Job Hunt → Run workflow**

---

## 💰 Cost — Truly Zero

| Component | Service | Cost |
|-----------|---------|------|
| Scheduler | GitHub Actions cron | Free |
| Runner | ubuntu-latest | Free |
| AI model | Gemini 1.5 Flash | **Free (1500 req/day)** |
| Excel | openpyxl | Free |
| Email | Gmail SMTP | Free |
| Backup | GitHub Artifacts | Free |

---

## 🔧 Troubleshooting

**`FileNotFoundError: No resume found`** → commit `resume.pdf` to repo root

**`GEMINI_API_KEY` error** → double-check the secret name matches exactly

**Workflow not triggering** → repos inactive 60+ days need a manual trigger first

**Gmail auth failed** → regenerate App Password, ensure 2FA is on
