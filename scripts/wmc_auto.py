#!/usr/bin/env python3
"""
WMC Auto Pipeline — ดึง Excel จาก Gmail อัตโนมัติ แล้ว run dashboard
สำหรับ Scheduled Task: credentials อ่านจาก Environment Variables
"""

import imaplib
import email
import email.header
import os
import sys
import re
import subprocess
import json
import base64
import urllib.request
from datetime import datetime, timedelta

# =================== CONFIG ===================
# อ่านจาก env var (set โดย scheduled task prompt)
GMAIL_USER         = os.getenv("WMC_GMAIL_USER", "patradul.a@gmail.com")
GMAIL_APP_PASSWORD = os.getenv("WMC_GMAIL_PASSWORD", "")
CALLCENTER_FROM    = "WMC-Callcenter@wmchospital.com"
EXCEL_PREFIX       = "WMCDailyManagementType16"
GITHUB_TOKEN       = os.getenv("WMC_GH_TOKEN", "")
GITHUB_REPO        = "Adisornpatradul/wmc-dashboard"
COWORK_DIR         = os.path.dirname(os.path.abspath(__file__))
# =============================================


def decode_filename(raw):
    """Decode MIME-encoded filename"""
    parts = email.header.decode_header(raw)
    result = ""
    for part, enc in parts:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="replace")
        else:
            result += part
    return result


def download_excel_from_gmail():
    """ดาวน์โหลด Excel ล่าสุดจาก Gmail (IMAP)"""
    print("📧 กำลังเชื่อมต่อ Gmail...")

    if not GMAIL_APP_PASSWORD:
        print("❌ ยังไม่ได้ตั้ง WMC_GMAIL_PASSWORD environment variable")
        return None, None

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        mail.select("inbox")
    except Exception as e:
        print(f"❌ Login ไม่ได้: {e}")
        return None, None

    date_since = (datetime.now() - timedelta(days=2)).strftime("%d-%b-%Y")
    _, msg_ids = mail.search(None, f'FROM "{CALLCENTER_FROM}" SINCE "{date_since}"')

    if not msg_ids[0]:
        print("❌ ไม่พบ email จาก callcenter วันนี้")
        mail.close(); mail.logout()
        return None, None

    all_ids = msg_ids[0].split()
    print(f"   พบ {len(all_ids)} email(s) จาก callcenter")

    for msg_id in reversed(all_ids):
        _, msg_data = mail.fetch(msg_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if not part.get("Content-Disposition"):
                continue
            raw_fn = part.get_filename()
            if not raw_fn:
                continue
            filename = decode_filename(raw_fn)
            if EXCEL_PREFIX in filename and filename.endswith(".xlsx"):
                filepath = os.path.join(COWORK_DIR, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"✅ ดาวน์โหลดแล้ว: {filename}")
                mail.close(); mail.logout()
                return filepath, filename

    mail.close(); mail.logout()
    print("❌ ไม่พบ Excel attachment ใน email")
    return None, None


def parse_date_from_filename(filename):
    """แปลงวันที่จากชื่อไฟล์ → YYYY-MM-DD"""
    fname = re.sub(r"\.\w+$", "", os.path.basename(filename))
    fname = re.sub(r"WMCDailyManagementType\d+", "", fname)
    patterns = [
        r"(\d{1,2})[\s\-_()/]+(\d{1,2})[\s\-_()/]+(\d{2,4})",
        r"(\d{4})[\s\-_()/]+(\d{1,2})[\s\-_()/]+(\d{1,2})",
    ]
    for pat in patterns:
        m = re.search(pat, fname)
        if m:
            a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if a > 1000:
                y, mo, d = a, b, c
            else:
                d, mo, y = a, b, c
                if y < 100:
                    y = y + 1957
            return f"{y:04d}-{mo:02d}-{d:02d}"
    return datetime.now().strftime("%Y-%m-%d")


def download_html_from_github(date_str):
    """ดาวน์โหลด HTML dashboard ปัจจุบันจาก GitHub API"""
    months_en = ["", "January", "February", "March", "April", "May",
                 "June", "July", "August", "September", "October", "November", "December"]
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        html_name = f"WMC_Monthly_Dashboard_{months_en[dt.month]}{dt.year}.html"
    except Exception:
        return False

    local_path = os.path.join(COWORK_DIR, html_name)
    if os.path.exists(local_path):
        return True

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{html_name}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    try:
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req) as r:
            data = json.loads(r.read())
        html_bytes = base64.b64decode(data["content"])
        with open(local_path, "wb") as f:
            f.write(html_bytes)
        print(f"✅ ดาวน์โหลด {html_name} จาก GitHub ({len(html_bytes)//1024} KB)")
        return True
    except Exception as e:
        print(f"⚠️  ดาวน์โหลด HTML ไม่ได้: {e}")
        return False


def find_latest_local_excel():
    """fallback: หา Excel ล่าสุดใน working dir"""
    files = [
        f for f in os.listdir(COWORK_DIR)
        if f.startswith(EXCEL_PREFIX) and f.endswith(".xlsx")
    ]
    if not files:
        return None, None
    files.sort(key=lambda f: os.path.getmtime(os.path.join(COWORK_DIR, f)), reverse=True)
    filename = files[0]
    return os.path.join(COWORK_DIR, filename), filename


def run_pipeline(filename, date_str):
    """เรียก wmc_excel_processor.py"""
    proc_script = os.path.join(COWORK_DIR, "wmc_excel_processor.py")
    if not os.path.exists(proc_script):
        print(f"❌ ไม่พบ {proc_script}")
        return 1
    result = subprocess.run(
        ["python3", proc_script, filename, date_str],
        cwd=COWORK_DIR
    )
    return result.returncode


# =================== MAIN ===================
if __name__ == "__main__":
    print("=" * 50)
    print("🤖 WMC Auto Pipeline")
    print(f"⏰ {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)
    print()

    manual_date = None
    if "--date" in sys.argv:
        idx = sys.argv.index("--date")
        if idx + 1 < len(sys.argv):
            manual_date = sys.argv[idx + 1]

    # 1. ดึงจาก Gmail
    filepath, filename = download_excel_from_gmail()

    # 2. fallback: ใช้ไฟล์ใน working dir
    if not filepath:
        print("⚠️  fallback → ใช้ไฟล์ล่าสุดใน working dir แทน")
        filepath, filename = find_latest_local_excel()

    if not filepath:
        print("❌ ไม่พบข้อมูลเลย — หยุดทำงาน")
        sys.exit(1)

    # 3. แปลงวันที่
    date_str = manual_date or parse_date_from_filename(filename)
    print(f"📅 วันที่: {date_str}")
    print()

    # 3.5 ดาวน์โหลด HTML dashboard จาก GitHub ถ้ายังไม่มี local
    download_html_from_github(date_str)

    # 4. รัน pipeline
    rc = run_pipeline(filename, date_str)
    sys.exit(rc)
