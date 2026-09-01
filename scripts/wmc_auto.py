#!/usr/bin/env python3
"""
WMC Auto Pipeline — ดึง Excel จาก Gmail อัตโนมัติ แล้ว run dashboard
ไม่ต้องวางไฟล์เอง ไม่ต้องพิม run dashboard

วิธีใช้:
  python3 wmc_auto.py                    # ดึง Gmail อัตโนมัติ
  python3 wmc_auto.py --date 2026-06-28  # ระบุวันที่เอง (ถ้าต้องการ)
"""

import imaplib
import email
import email.header
import email.utils
import calendar
import os
import sys
import re
import subprocess
import json
import base64
import urllib.request
import urllib.error
import time
from datetime import datetime, timedelta

# =================== CONFIG ===================
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
        print("❌ ยังไม่ได้ตั้ง GMAIL_APP_PASSWORD ใน wmc_auto.py")
        print("   กรุณาไปที่ https://myaccount.google.com/apppasswords")
        print("   แล้วใส่รหัสในตัวแปร GMAIL_APP_PASSWORD")
        return None, None

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        mail.select("inbox")
    except Exception as e:
        print(f"❌ Login ไม่ได้: {e}")
        return None, None

    # ค้นหา email จาก callcenter ย้อนหลัง 2 วัน
    date_since = (datetime.now() - timedelta(days=2)).strftime("%d-%b-%Y")
    _, msg_ids = mail.search(None, f'FROM "{CALLCENTER_FROM}" SINCE "{date_since}"')

    if not msg_ids[0]:
        print("❌ ไม่พบ email จาก callcenter วันนี้")
        mail.close(); mail.logout()
        return None, None

    all_ids = msg_ids[0].split()
    print(f"   พบ {len(all_ids)} email(s) จาก callcenter")

    # ดึง Excel ทุกฉบับจากทุก email แล้ว sort by วันที่ในชื่อไฟล์ (แม่นที่สุด)
    candidates = []  # list of (filepath, filename, date_str)
    for msg_id in all_ids:
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
                date_str = parse_date_from_filename(filename)
                candidates.append((filepath, filename, date_str))
                print(f"   พบ Excel: {filename} → {date_str}")

    mail.close(); mail.logout()

    if not candidates:
        print("❌ ไม่พบ Excel attachment ใน email")
        return None, None

    # เลือก Excel ที่มีวันที่ล่าสุดในชื่อไฟล์
    candidates.sort(key=lambda x: x[2], reverse=True)
    filepath, filename, date_str = candidates[0]
    print(f"✅ เลือก: {filename} (ล่าสุด: {date_str})")
    return filepath, filename


def parse_date_from_filename(filename):
    """แปลงวันที่จากชื่อไฟล์ → YYYY-MM-DD
    ตัด prefix WMCDailyManagementType16 ออกก่อน เพื่อไม่ให้ตัวเลข 16 ไปรบกวน
    """
    fname = re.sub(r"\.\w+$", "", os.path.basename(filename))
    # ตัด prefix ที่รู้จัก
    fname = re.sub(r"WMCDailyManagementType\d+", "", fname)
    patterns = [
        r"(\d{1,2})[\s\-_()/]+(\d{1,2})[\s\-_()/]+(\d{2,4})",
        r"(\d{4})[\s\-_()/]+(\d{1,2})[\s\-_()/]+(\d{1,2})",
    ]
    for pat in patterns:
        m = re.search(pat, fname)
        if m:
            a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if a > 1000:  # YYYY-MM-DD
                y, mo, d = a, b, c
            else:         # DD-MM-YY(YY)
                d, mo, y = a, b, c
                if y < 100:
                    y = y + 1957  # Thai short BE → CE
            return f"{y:04d}-{mo:02d}-{d:02d}"
    return datetime.now().strftime("%Y-%m-%d")


def download_html_from_github(date_str):
    """ดาวน์โหลด HTML dashboard ปัจจุบันจาก GitHub API
    ใช้เมื่อรันจาก /tmp (scheduled task) และไม่มีไฟล์ local

    Return:
      "exists"     - มีไฟล์ local อยู่แล้ว
      "downloaded" - ดาวน์โหลดสำเร็จ
      "new_month"  - ไม่พบไฟล์บน GitHub จริง ๆ (404) ปลอดภัยที่จะสร้างใหม่
      "error"      - ดาวน์โหลดล้มเหลวหลังพยายามหลายครั้ง (เช่น network/rate-limit)
                     ห้าม fallback ไปสร้างไฟล์ใหม่ทับของเดิม เพราะจะทำข้อมูลหาย
    """
    months_en = ["", "January", "February", "March", "April", "May",
                 "June", "July", "August", "September", "October", "November", "December"]
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        html_name = f"WMC_Monthly_Dashboard_{months_en[dt.month]}{dt.year}.html"
    except Exception:
        return "error"

    local_path = os.path.join(COWORK_DIR, html_name)
    if os.path.exists(local_path):
        return "exists"  # มีอยู่แล้ว ไม่ต้องโหลดซ้ำ

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{html_name}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        try:
            req = urllib.request.Request(api_url, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as r:
                data = json.loads(r.read())
            html_bytes = base64.b64decode(data["content"])
            with open(local_path, "wb") as f:
                f.write(html_bytes)
            print(f"✅ ดาวน์โหลด {html_name} จาก GitHub สำเร็จ ({len(html_bytes)//1024} KB)")
            return "downloaded"
        except urllib.error.HTTPError as e:
            if e.code == 404:
                print(f"ℹ️  ไม่พบ {html_name} บน GitHub (404) — เป็นเดือนใหม่ สร้างไฟล์ใหม่ได้อย่างปลอดภัย")
                return "new_month"
            print(f"⚠️  HTTP error {e.code} ตอนดาวน์โหลด {html_name} (ครั้งที่ {attempt}/{max_attempts}): {e}")
        except Exception as e:
            print(f"⚠️  ดาวน์โหลด {html_name} ล้มเหลว (ครั้งที่ {attempt}/{max_attempts}): {e}")
        if attempt < max_attempts:
            time.sleep(3)

    print(f"❌ ไม่สามารถดาวน์โหลด {html_name} จาก GitHub ได้หลังพยายาม {max_attempts} ครั้ง")
    print("   หยุดก่อนดำเนินการต่อ เพื่อป้องกันไม่ให้ระบบสร้างไฟล์ใหม่ทับข้อมูลเดือนนี้ที่มีอยู่แล้วโดยไม่ตั้งใจ")
    return "error"


def find_latest_local_excel():
    """fallback: หา Excel ล่าสุดใน Cowork folder"""
    files = [
        f for f in os.listdir(COWORK_DIR)
        if f.startswith(EXCEL_PREFIX) and f.endswith(".xlsx")
    ]
    if not files:
        return None, None
    files.sort(key=lambda f: os.path.getmtime(os.path.join(COWORK_DIR, f)),
               reverse=True)
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

    # รับ argument --date ถ้ามี
    manual_date = None
    if "--date" in sys.argv:
        idx = sys.argv.index("--date")
        if idx + 1 < len(sys.argv):
            manual_date = sys.argv[idx + 1]

    # 1. ลองดึงจาก Gmail ก่อน
    filepath, filename = download_excel_from_gmail()

    # 2. fallback: ใช้ไฟล์ใน Cowork folder
    if not filepath:
        print("⚠️  fallback → ใช้ไฟล์ล่าสุดใน Cowork folder แทน")
        filepath, filename = find_latest_local_excel()

    if not filepath:
        print("❌ ไม่พบข้อมูลเลย — หยุดทำงาน")
        sys.exit(1)

    # 3. แปลงวันที่
    date_str = manual_date or parse_date_from_filename(filename)
    print(f"📅 วันที่: {date_str}")
    print()

    # 3.5 ดาวน์โหลด HTML dashboard จาก GitHub ถ้ายังไม่มี local
    #     (จำเป็นเมื่อรันจาก /tmp ใน scheduled task — ไม่มี Cowork folder)
    html_status = download_html_from_github(date_str)
    if html_status == "error":
        print("❌ หยุด pipeline: ไม่สามารถยืนยันสถานะ Dashboard HTML บน GitHub ได้")
        print("   ไม่ดำเนินการต่อ เพื่อป้องกันข้อมูลเดือนนี้เสียหาย — กรุณารันใหม่อีกครั้ง")
        sys.exit(1)

    # 4. รัน pipeline
    rc = run_pipeline(filename, date_str)
    sys.exit(rc)
