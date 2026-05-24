# Job Hunting Pipeline 🤖

An automated job hunting assistant that reads your resume, discovers relevant DevOps jobs in India using free AI models via **OpenRouter**, scores each listing against your profile, and delivers a structured Excel report to your inbox — every day at 10 AM IST, for free.

---

## How It Works

1. **GitHub Actions** triggers the workflow daily at 10 AM IST (or on demand).
2. Your `resume.pdf` is extracted and passed to **OpenRouter's free model tier** via two sequential API calls:
   - **Call 1** — candidate analysis + 25 scored job listings (JSON)
   - **Call 2** — 30-day personalised resume improvement tips (JSON)
3. The script tries multiple free models in order, falling back automatically if one times out or returns empty results.
4. Both outputs are merged into an **Excel workbook** with 6 sheets.
5. The report is **emailed to you via Gmail SMTP** and saved as a GitHub Actions artifact.

---

## Repository Structure

```
Job-Hunting-pipeline/
├── .github/
│   └── workflows/
│       └── daily_report.yml    # Cron scheduler (10 AM IST daily)
├── scripts/
│   ├── generate_excel.py       # OpenRouter AI calls + Excel builder
│   └── send_email.py           # Gmail SMTP sender
├── resume.pdf                  # ⭐ Your resume goes here
├── prompt.txt                  # AI prompt for job search customisation
└── README.md
```

---

## Excel Report — 6 Sheets Delivered Daily

| Sheet | Contents |
|---|---|
| 📊 Dashboard | Summary stats, skill strengths, gaps, and role counts |
| 🗂️ All Jobs | Full list of roles sorted by fit score, deduplicated |
| 🟢 High Probability | Fit score ≥ 75 — prioritise these |
| 🟡 Medium Probability | Fit score 50–74 |
| 🔴 Stretch Roles | Fit score < 50 — aspirational targets |
| 📅 30-Day Resume Tips | One personalised improvement tip per day with before/after examples |

---

## AI Models Used

The script cycles through OpenRouter's free-tier models in order, moving to the next if a model fails or returns empty results:

1. `deepseek/deepseek-v4-flash:free`
2. `nvidia/nemotron-3-super-120b-a12b:free`
3. `meta-llama/llama-3.3-70b-instruct:free`
4. `google/gemma-4-31b-it:free`

No credit card or paid plan required — OpenRouter's free tier covers all usage.

---

## Setup (One-Time)

### Step 1 — Fork and add your resume

Fork this repository (use a **private fork** to protect your resume), then add your resume:

```bash
cp your-cv.pdf resume.pdf
git add resume.pdf
git commit -m "Add resume"
git push
```

### Step 2 — Get a free OpenRouter API key

1. Visit [openrouter.ai](https://openrouter.ai)
2. Sign up and go to **Keys → Create Key**
3. Copy the key (starts with `sk-or-...`)

> Free-tier models are available with no billing setup required.

### Step 3 — Get a Gmail App Password

1. Go to [myaccount.google.com/security](https://myaccount.google.com/security)
2. Enable **2-Step Verification**
3. Search for **App passwords** → Create one → name it `Job Hunt Bot`
4. Copy the 16-character password (no spaces)

### Step 4 — Add GitHub Secrets

Go to your repo → **Settings → Secrets and variables → Actions → New repository secret** and add:

| Secret Name | Description | Example |
|---|---|---|
| `OPENROUTER_API_KEY` | Your OpenRouter API key | `sk-or-v1-XXXXXXXX` |
| `GMAIL_USER` | Gmail address used to send | `you@gmail.com` |
| `GMAIL_APP_PASS` | 16-character app password | `abcdefghijklmnop` |
| `EMAIL_TO` | Address to receive the report | `you@gmail.com` |

### Step 5 — Run your first report

Go to **Actions → 🎯 Daily Job Hunt → Run workflow** to trigger it immediately without waiting for the 10 AM schedule.

---

## Cost Breakdown

Everything used in this pipeline is free:

| Component | Service | Cost |
|---|---|---|
| Scheduler | GitHub Actions cron | Free |
| CI Runner | ubuntu-latest | Free |
| AI Models | OpenRouter free tier | Free |
| Excel Generation | openpyxl | Free |
| Email Delivery | Gmail SMTP | Free |
| Report Backup | GitHub Actions artifacts | Free |

---

## Updating Your Secrets

If you need to rotate an API key, change your Gmail account, or add more recipients, here's how to update each secret.

### How to edit a secret on GitHub

1. Go to your forked repo on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Find the secret you want to change and click the **pencil icon** next to it
4. Paste the new value and click **Update secret**

> GitHub does not let you view an existing secret's value — you can only overwrite it. If you're unsure what's saved, just replace it with a fresh value.

---

### `OPENROUTER_API_KEY`

Used by `generate_excel.py` to call the AI models via OpenRouter.

**When to update:** Your key was revoked, you hit usage limits, or you want to use a different account.

1. Log in to [openrouter.ai](https://openrouter.ai)
2. Go to **Keys** in the left sidebar
3. Click **Create Key** (or delete the old one first)
4. Copy the new key (starts with `sk-or-...`)
5. Update the `OPENROUTER_API_KEY` secret in GitHub

---

### `GMAIL_USER`

The Gmail address the workflow sends the report **from**.

**When to update:** You want to send from a different Gmail account.

1. Decide which Gmail address you want to use as the sender
2. Update the `GMAIL_USER` secret to that address (e.g. `yourname@gmail.com`)
3. Make sure you also update `GMAIL_APP_PASS` to match the new account (app passwords are account-specific)

---

### `GMAIL_APP_PASS`

The 16-character app password that authenticates Gmail SMTP. This is **not** your Gmail login password.

**When to update:** Authentication starts failing, you changed Gmail accounts, or you revoked the old app password.

1. Go to [myaccount.google.com/security](https://myaccount.google.com/security)
2. Make sure **2-Step Verification** is enabled
3. Search for **App passwords**
4. Delete the old `Job Hunt Bot` entry if it exists
5. Create a new one → select **Other** → name it `Job Hunt Bot`
6. Copy the 16-character code shown (remove spaces)
7. Update the `GMAIL_APP_PASS` secret in GitHub

---

### `EMAIL_TO`

The address (or addresses) that receive the daily Excel report.

**When to update:** You want to receive reports at a different address, or share them with multiple people.

- Single recipient: `you@gmail.com`
- Multiple recipients (comma-separated): `you@gmail.com,colleague@gmail.com`

Update the `EMAIL_TO` secret with the new value. No changes to code required — `send_email.py` handles comma-separated lists automatically.

---

### Quick Reference

| Secret | Where to get it | Looks like |
|---|---|---|
| `OPENROUTER_API_KEY` | [openrouter.ai](https://openrouter.ai) → Keys | `sk-or-v1-XXXXXXXX` |
| `GMAIL_USER` | Your Gmail address | `you@gmail.com` |
| `GMAIL_APP_PASS` | Google Account → Security → App passwords | `abcdefghijklmnop` |
| `EMAIL_TO` | Any email address(es) | `you@gmail.com` |

After updating any secret, trigger a manual run via **Actions → Daily Job Hunt → Run workflow** to verify everything works.

---

## Troubleshooting

**`FileNotFoundError: No resume found`**
Ensure `resume.pdf` is committed to the root of your repository.

**`OPENROUTER_API_KEY` error**
Double-check the secret name matches exactly — it is case-sensitive.

**All models failed for [JOBS] or [TIPS]**
Free-tier models can be rate-limited during peak hours. Try triggering the workflow manually at a different time, or add more fallback models to the `FREE_MODELS` list in `generate_excel.py`.

**Workflow not triggering automatically**
GitHub pauses scheduled workflows on repos inactive for 60+ days. Trigger it manually once to reactivate.

**Gmail authentication failed**
Regenerate your App Password and make sure 2-Step Verification is still enabled on your account.

---

## Customisation

Edit `prompt.txt` to change the job search behaviour — for example, targeting a different role (e.g., Cloud Engineer, SRE), location, experience level, or tech stack. The prompt is passed directly to the AI model along with your resume, so the more specific you are, the better the results.

To change the fallback model order, edit the `FREE_MODELS` list in `scripts/generate_excel.py`.

---

## Author

Built by **V K Harish Bodapati**
