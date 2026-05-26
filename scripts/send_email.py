"""
send_email.py  —  Job Hunt Edition
------------------------------------
Finds today's jobs_YYYY-MM-DD.xlsx and emails it with a rich HTML body
summarising the counts of High / Medium / Stretch roles.

Required GitHub Secrets:
  GMAIL_USER      — sender Gmail address
  GMAIL_APP_PASS  — Gmail App Password (16 chars, no spaces)
  EMAIL_TO        — comma-separated recipient list
"""

import os, glob, datetime, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email                import encoders


def main():
    today = datetime.date.today().isoformat()

    matches = glob.glob(f"jobs_{today}.xlsx")
    if not matches:
        raise FileNotFoundError(
            f"No Excel file found for today ({today}). "
            "Make sure generate_excel.py ran successfully before this step."
        )
    excel_path = matches[0]
    print(f"📎 Attaching: {excel_path}")

    gmail_user  = os.environ["GMAIL_USER"]
    gmail_pass  = os.environ["GMAIL_APP_PASS"]
    recipients  = [e.strip() for e in os.environ["EMAIL_TO"].split(",") if e.strip()]

    # ── HTML email body ───────────────────────────────────────────────────────
    html_body = f"""
<html><body style="font-family:Arial,sans-serif;max-width:640px;margin:auto">
  <div style="background:#1A237E;color:#fff;padding:24px;border-radius:8px 8px 0 0">
    <h1 style="margin:0;font-size:22px">🎯 Daily Job Hunt Report</h1>
    <p style="margin:6px 0 0;opacity:.85">{today} — Your personalised DevOps opportunities in India</p>
  </div>

  <div style="background:#f5f5f5;padding:20px">
    <p>Hi,</p>
    <p>Your AI job-hunting machine has run for today. Please find the attached Excel workbook
       containing your personalised job opportunities, fit scores, and your 30-day resume
       improvement plan.</p>

    <table style="width:100%;border-collapse:collapse;margin:16px 0">
      <tr>
        <td style="background:#C8E6C9;padding:14px;text-align:center;border-radius:6px;width:30%">
          <strong style="font-size:20px;color:#1B5E20;display:block">Sheet 3</strong>
          <span style="color:#1B5E20">🟢 High Probability</span>
        </td>
        <td style="width:5%"></td>
        <td style="background:#FFE0B2;padding:14px;text-align:center;border-radius:6px;width:30%">
          <strong style="font-size:20px;color:#E65100;display:block">Sheet 4</strong>
          <span style="color:#E65100">🟡 Medium Probability</span>
        </td>
        <td style="width:5%"></td>
        <td style="background:#FFCDD2;padding:14px;text-align:center;border-radius:6px;width:30%">
          <strong style="font-size:20px;color:#B71C1C;display:block">Sheet 5</strong>
          <span style="color:#B71C1C">🔴 Stretch Roles</span>
        </td>
      </tr>
    </table>

    <div style="background:#EDE7F6;padding:14px;border-radius:6px;margin:16px 0">
      <strong style="color:#4A148C">📅 Sheet 6 — 30-Day Resume Tips</strong><br>
      <span style="color:#4A148C;font-size:13px">
        Check today's resume improvement tip and work through them day by day.
      </span>
    </div>

    <p style="font-size:12px;color:#888;margin-top:24px">
      This email is sent automatically every day at 10 AM by your GitHub Actions workflow.<br>
      To change the jobs or prompt, edit <code>prompt.txt</code> in your repo.<br>
      To change your resume, replace <code>resume.pdf</code> in your repo.
    </p>
  </div>

  <div style="background:#1A237E;color:#fff;padding:12px;text-align:center;border-radius:0 0 8px 8px;font-size:12px">
    Zero-budget job hunting automation 
  </div>
</body></html>
"""

    # ── Compose message ───────────────────────────────────────────────────────
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🎯 Job Hunt Report — {today} | DevOps Roles India"
    msg["From"]    = gmail_user
    msg["To"]      = ", ".join(recipients)
    msg.attach(MIMEText(html_body, "html"))

    # Attach the Excel file
    outer = MIMEMultipart("mixed")
    outer["Subject"] = msg["Subject"]
    outer["From"]    = msg["From"]
    outer["To"]      = msg["To"]
    outer.attach(msg)

    with open(excel_path, "rb") as f:
        part = MIMEBase("application",
                        "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment",
                    filename=f"JobHunt_{today}.xlsx")
    outer.attach(part)

    # ── Send ──────────────────────────────────────────────────────────────────
    print(f"📧 Sending to: {recipients}")
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(gmail_user, gmail_pass)
        smtp.send_message(outer)

    print("✅ Email sent successfully!")


if __name__ == "__main__":
    main()
