"""
send_failure_email.py  —  Job Hunt Edition
------------------------------------------
Sends a failure alert email when the workflow crashes.
Called automatically by daily_report.yml on any step failure.

Uses the same Gmail secrets as send_email.py — no extra setup needed.
"""

import os, smtplib, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText


def main():
    today       = datetime.date.today().isoformat()
    run_id      = os.environ.get("RUN_ID",      "unknown")
    run_url     = os.environ.get("RUN_URL",      "#")
    failed_step = os.environ.get("FAILED_STEP",  "Unknown step")

    gmail_user  = os.environ["GMAIL_USER"]
    gmail_pass  = os.environ["GMAIL_APP_PASS"]
    recipients  = [e.strip() for e in os.environ["EMAIL_TO"].split(",") if e.strip()]

    html_body = f"""
<html><body style="font-family:Arial,sans-serif;max-width:640px;margin:auto">

  <div style="background:#B71C1C;color:#fff;padding:24px;border-radius:8px 8px 0 0">
    <h1 style="margin:0;font-size:22px">🚨 Job Hunt Pipeline Failed</h1>
    <p style="margin:6px 0 0;opacity:.85">{today} — Your daily job report was not generated</p>
  </div>

  <div style="background:#f5f5f5;padding:20px">
    <p>Hi,</p>
    <p>Your AI job-hunting pipeline failed today and the Excel report was <strong>not</strong> sent.</p>

    <table style="width:100%;border-collapse:collapse;margin:16px 0;background:#fff;border-radius:6px">
      <tr>
        <td style="padding:12px 16px;border-bottom:1px solid #eee;width:38%">
          <strong style="color:#555">❌ Failed Step</strong>
        </td>
        <td style="padding:12px 16px;border-bottom:1px solid #eee;color:#B71C1C;font-weight:bold">
          {failed_step}
        </td>
      </tr>
      <tr>
        <td style="padding:12px 16px;border-bottom:1px solid #eee">
          <strong style="color:#555">🔢 Run ID</strong>
        </td>
        <td style="padding:12px 16px;border-bottom:1px solid #eee;font-family:monospace">
          {run_id}
        </td>
      </tr>
      <tr>
        <td style="padding:12px 16px">
          <strong style="color:#555">📅 Date</strong>
        </td>
        <td style="padding:12px 16px">{today}</td>
      </tr>
    </table>

    <div style="text-align:center;margin:24px 0">
      <a href="{run_url}"
         style="background:#1A237E;color:#fff;padding:12px 28px;border-radius:6px;
                text-decoration:none;font-weight:bold;font-size:14px">
        🔍 View Full Logs on GitHub
      </a>
    </div>

    <div style="background:#FFF8E1;border-left:4px solid #F9A825;padding:14px;border-radius:4px;margin:16px 0">
      <strong style="color:#F57F17">Common causes &amp; quick fixes:</strong>
      <ul style="color:#555;margin:8px 0;padding-left:20px;line-height:1.7">
        <li><strong>All LLMs rate-limited</strong> — wait an hour and re-run manually</li>
        <li><strong>OPENROUTER_API_KEY missing/expired</strong> — check repo Secrets</li>
        <li><strong>GMAIL_APP_PASS expired</strong> — regenerate in Google Account settings</li>
        <li><strong>resume.pdf not committed</strong> — make sure it's in the repo root</li>
      </ul>
    </div>

    <p style="font-size:12px;color:#888;margin-top:24px">
      To re-run manually: GitHub → Actions → 🎯 Daily Job Hunt → Run workflow
    </p>
  </div>

  <div style="background:#B71C1C;color:#fff;padding:12px;text-align:center;
              border-radius:0 0 8px 8px;font-size:12px">
    Zero-budget job hunting automation · Powered by OpenRouter + GitHub Actions
  </div>

</body></html>
"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🚨 Job Hunt Pipeline Failed — {today}"
    msg["From"]    = gmail_user
    msg["To"]      = ", ".join(recipients)
    msg.attach(MIMEText(html_body, "html"))

    print(f"📧 Sending failure alert to: {recipients}")
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(gmail_user, gmail_pass)
        smtp.send_message(msg)

    print("✅ Failure alert sent!")


if __name__ == "__main__":
    main()
