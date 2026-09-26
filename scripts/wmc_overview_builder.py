#!/usr/bin/env python3
"""
WMC Hospital Overview — Daily Data Builder
สร้าง data/overview/{date}.json สำหรับ dashboard แบบเลือกวันที่/ช่วงวันที่ได้
ต้องรันบนคอมพิวเตอร์ของผู้ใช้ (ผ่าน device_bash) เพราะต้องอ่านไฟล์ CH1288/CH1335/CH1434
จากโฟลเดอร์ Cowork ในเครื่อง (ไฟล์เหล่านี้ไม่ได้มาทาง Gmail)

วิธีใช้:
  python3 wmc_overview_builder.py --date 2026-09-25 \
      --cowork-dir "/Users/dradisorn/Desktop/COWORK/Cowork NEW daily dashboard" \
      --gmail-user patradul.a@gmail.com --gmail-password "xxxx xxxx xxxx xxxx" \
      --gh-token ghp_xxx --gh-repo Adisornpatradul/wmc-dashboard

ถ้าไม่ระบุ --date จะใช้ "เมื่อวาน" (ตามเวลาที่รัน) โดยอัตโนมัติ
เพราะรายงานของวันที่ X จะถูกส่งอีเมลตอนราวตี 4 ของวันที่ X+1
"""

import argparse
import base64
import imaplib
import email
import email.header
import json
import os
import re
import sys
import urllib.request
import urllib.error
from datetime import datetime, timedelta

EXCEL_PREFIX = "WMCDailyManagementType16"


# --------------------------------------------------------------------------
# Gmail: ดึง Report 16 ของวันที่ target_date
# --------------------------------------------------------------------------
def decode_filename(raw):
    parts = email.header.decode_header(raw)
    result = ""
    for part, enc in parts:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="replace")
        else:
            result += part
    return result


def thai_be_short(dt):
    """9/25/69-style short Thai BE year suffix used in subject/filename"""
    return f"{dt.day:02d}/{dt.month:02d}/{(dt.year + 543) % 100:02d}"


def fetch_report16(gmail_user, gmail_password, target_date, out_dir):
    """คืน path ของไฟล์ Report16 ที่ดาวน์โหลดมา หรือ None บ้าไม่พบ"""
    dd_mm_yy = thai_be_short(target_date)  # e.g. 25/09/69
    subject_needle_1 = f"{dd_mm_yy.replace('/', '/')}"  # 25/09/69
    filename_needle = f"({target_date.day:02d}-{target_date.month:02d}-{(target_date.year + 543) % 100:02d})"

    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(gmail_user, gmail_password)
    M.select("INBOX")
    status, data = M.search(None, '(FROM "WMC-Callcenter@wmchospital.com")')
    ids = data[0].split()

    target_id = None
    # เช็คจากใหม่ไปเก่า ย้อนหลังไม่เกิน 20 ฉบับ
    # หมายเหตุ: Subject เป็น MIME encoded-word (=?utf-8?B?...?=) เสมอ เพราะมีอักษรไทย
    # ต้อง decode ก่อนเทียบ ห้ามเทียบ substring กับ raw header ตรงๆ (จะไม่เจอ)
    for msgid in reversed(ids[-20:]):
        status, msgdata = M.fetch(msgid, "(BODY[HEADER.FIELDS (SUBJECT DATE)])")
        header_bytes = msgdata[0][1]
        msg_hdr = email.message_from_bytes(header_bytes)
        subject_decoded = decode_filename(msg_hdr.get("Subject", "") or "")
        if subject_needle_1 in subject_decoded:
            target_id = msgid
            break

    if target_id is None:
        M.logout()
        return None

    status, msgdata = M.fetch(target_id, "(RFC822)")
    raw = msgdata[0][1]
    msg = email.message_from_bytes(raw)
    saved_path = None
    for part in msg.walk():
        fn = part.get_filename()
        if not fn:
            continue
        fn_decoded = decode_filename(fn)
        if fn_decoded.lower().endswith(".xlsx") and EXCEL_PREFIX in fn_decoded:
            payload = part.get_payload(decode=True)
            out_name = f"{EXCEL_PREFIX} ({target_date.day}-{target_date.month}-{(target_date.year + 543) % 100}).xlsx"
            saved_path = os.path.join(out_dir, out_name)
            with open(saved_path, "wb") as f:
                f.write(payload)
    M.logout()
    return saved_path


# --------------------------------------------------------------------------
# หาไฟล์ CH1288 / CH1335 / CH1434 ล่าสุดในโฟลเดอร์ Cowork ที่ตรงกับวันที่เป้าหมาย
# --------------------------------------------------------------------------
def find_ch_file(cowork_dir, ch_code, target_date):
    """
    6ต้องเปิดดูข้างในเพื่อเช็ควันที่
    (แถวแรกของ sheet Document_CH#### เป็น datetime ของวันที่ข้อมูล)
    คืน path ของไฟล์ที่ตรงวันที่ ถ้าไม่เจอคืน None
    """
    import openpyxl

    candidates = []
    for fn in os.listdir(cowork_dir):
        if not fn.lower().endswith(".xlsx"):
            continue
        full = os.path.join(cowork_dir, fn)
        try:
            wb = openpyxl.load_workbook(full, data_only=True, read_only=True)
            ws = wb.worksheets[0]
            if ws.title != f"Document_{ch_code}":
                wb.close()
                continue
            first_cell = ws.cell(row=1, column=1).value
            wb.close()
        except Exception:
            continue
        if isinstance(first_cell, datetime) and first_cell.date() == target_date.date():
            candidates.append((full, os.path.getmtime(full)))

    if not candidates:
        return None
    # ถ้ามีหลายไฟล์ตรงวันที่ (ไม่ควรเกิด) เลือกไฟล์ที่ใหม่สุด
    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0]


# --------------------------------------------------------------------------
# คำนวณ D (ข้อมูล Hospital Overview ของวันเดียว) จาก Report16 + 3 ไฟล์ CH
# --------------------------------------------------------------------------
def pct(s):
    if s is None:
        return None
    if isinstance(s, str):
        s = s.replace("%", "").strip()
        try:
            return round(float(s))
        except ValueError:
            return None
    return s


def clean_num(v):
    if isinstance(v, str):
        v = v.replace(",", "").strip()
        if v in ("", "-"):
            return None
        try:
            return int(v)
        except ValueError:
            return float(v)
    return v


def build_overview_data(report16_path, ch1288_path, ch1335_path, ch1434_path, target_date):
    import openpyxl

    wb = openpyxl.load_workbook(report16_path, data_only=True)
    ws = wb.worksheets[0]
    rows = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), start=2):
        rows[i] = row

    r11 = rows[11]
    summary = {
        "total_visit": r11[1], "ipd_admit": r11[4], "opd_new": r11[9], "opd_old": r11[10],
        "opd_total": r11[8], "counter_visit": r11[12], "total_bill_amount": r11[13],
        "avg_per_bill": r11[18], "avg_per_counter": r11[23],
    }

    ward = []
    r = 14
    while True:
        row = rows.get(r)
        if row is None or row[5] in (None, ""):
            break
        name = row[5].strip()
        short = name.split(" (")[0].strip()
        ward.append({"name": short, "rev": row[14]})
        r += 1
    ward_total_row = rows.get(r)
    ward_total = ward_total_row[14] if ward_total_row else sum(w["rev"] for w in ward)

    # หมายเหตุสำคัญ: จำนวนแถว "วอร์ด" (ward) ด้านบนไม่คงที่ — วันไหนมีวอร์ดที่มีคนไข้มากกว่ากัน
    # แถวของทุก section ที่อยู่ถัดลงไป (OR, OPD Revenue, Checkup Revenue, Cashier Billed ฯลฯ)
    # จะเลื่อนตามจำนวนวอร์ดนั้นด้วย ห้าม hardcode เลขแถวเด็ดขาด ต้องค้นด้วย label เสมอ

    or_total = None
    for rr, row in rows.items():
        if row and len(row) > 6 and row[6] == "Operating Room":
            or_total = row[11]
            break

    # "OPD Revenue" section: header row (col1=="OPD Revenue") แล้วข้อมูลอยู่ header+2..+4 (Thai/Inter/Mobile)
    # ระวังสับสนกับ "OPD Revenue Location" (อีก section ที่อยู่ไกลลงไป) — ต้องเทียบเท่ากันเป๊ะๆ
    opd_rev_header = None
    for rr, row in rows.items():
        if row and row[1] == "OPD Revenue":
            opd_rev_header = rr
            break
    opd_thai = opd_inter = opd_mobile = opd_total_amt = None
    opd_thai_visits = opd_inter_visits = opd_mobile_visits = opd_total_visits = None
    if opd_rev_header:
        r_thai = rows.get(opd_rev_header + 2)
        r_inter = rows.get(opd_rev_header + 3)
        r_mobile = rows.get(opd_rev_header + 4)
        r_total = rows.get(opd_rev_header + 5)
        if r_thai: opd_thai = clean_num(r_thai[8]); opd_thai_visits = clean_num(r_thai[10])
        if r_inter: opd_inter = clean_num(r_inter[8]); opd_inter_visits = clean_num(r_inter[10])
        if r_mobile: opd_mobile = clean_num(r_mobile[8]); opd_mobile_visits = clean_num(r_mobile[10])
        if r_total: opd_total_amt = clean_num(r_total[8]); opd_total_visits = clean_num(r_total[10])

    # "Checkup Revenue" section: header (col1=="Checkup Revenue") แล้ว Thai=header+2, Inter=header+3
    checkup_header = None
    for rr, row in rows.items():
        if row and row[1] == "Checkup Revenue":
            checkup_header = rr
            break
    checkup_total = 0
    checkup_visits = 0
    if checkup_header:
        r_thai = rows.get(checkup_header + 2)
        r_inter = rows.get(checkup_header + 3)
        if r_thai:
            checkup_total += clean_num(r_thai[8]) or 0
            checkup_visits += r_thai[10] or 0
        if r_inter:
            checkup_total += clean_num(r_inter[8]) or 0
            checkup_visits += r_inter[10] or 0

    # "Cashier Billed" section: header (มีคำว่า "Cashier Billed" อยู่ในแถว)
    # IPD Total billed by Type = header+3, OPD Total billed by Type = header+6, Total billed by Day = header+7
    cashier_header = None
    for rr, row in rows.items():
        if row and any(isinstance(c, str) and "Cashier Billed" in c for c in row):
            cashier_header = rr
            break
    cashier_ipd = cashier_ipd_count = None
    cashier_opd = cashier_opd_count = None
    cashier_total = cashier_total_count = None
    if cashier_header:
        ipd_row = rows.get(cashier_header + 3)
        opd_row = rows.get(cashier_header + 6)
        day_row = rows.get(cashier_header + 7)
        if ipd_row: cashier_ipd = ipd_row[10]; cashier_ipd_count = ipd_row[13]
        if opd_row: cashier_opd = opd_row[10]; cashier_opd_count = opd_row[13]
        if day_row: cashier_total = day_row[10]; cashier_total_count = day_row[13]

    revenue = {
        "cashier_ipd": cashier_ipd, "cashier_ipd_count": cashier_ipd_count,
        "cashier_opd": cashier_opd, "cashier_opd_count": cashier_opd_count,
        "cashier_total": cashier_total, "cashier_total_count": cashier_total_count,
        "ward_total": ward_total, "or_total": or_total,
        "opd_thai": opd_thai, "opd_thai_visits": opd_thai_visits,
        "opd_inter": opd_inter, "opd_inter_visits": opd_inter_visits,
        "opd_mobile": opd_mobile, "opd_mobile_visits": opd_mobile_visits,
        "opd_total": opd_total_amt, "opd_total_visits": opd_total_visits,
        "checkup_total": checkup_total, "checkup_visits": checkup_visits,
    }

    # OPD Revenue Location table — find header row dynamically ("OPD Revenue Location")
    header_row = None
    for rr, row in rows.items():
        if row and row[1] == "OPD Revenue Location":
            header_row = rr
            break
    depts = []
    if header_row:
        r = header_row + 2
        while True:
            row = rows.get(r)
            if row is None or row[1] in (None, ""):
                break
            depts.append({
                "name": row[1], "visits": row[9], "forecast": row[13], "rev": row[16],
                "achieve": pct(row[22]), "mtd_forecast": row[26], "mtd_rev": row[29],
                "mtd_achieve": pct(row[34]),
            })
            r += 1

    dept_rev_total = sum(d["rev"] for d in depts if d["rev"] is not None)
    dept_forecast_total = sum(d["forecast"] for d in depts if d["forecast"] is not None)
    dept_mtd_rev_total = sum(d["mtd_rev"] for d in depts if d["mtd_rev"] is not None)
    dept_mtd_forecast_total = sum(d["mtd_forecast"] for d in depts if d["mtd_forecast"] is not None)
    dept_achieve = round(dept_rev_total / dept_forecast_total * 100) if dept_forecast_total else None
    dept_mtd_achieve = round(dept_mtd_rev_total / dept_mtd_forecast_total * 100) if dept_mtd_forecast_total else None

    bu_opd_top5 = sorted(depts, key=lambda d: d["rev"] or 0, reverse=True)[:5]
    bu_opd_top5 = [{"name": d["name"], "rev": d["rev"]} for d in bu_opd_top5]

    patient_by_dept_top6 = sorted(depts, key=lambda d: d["visits"] or 0, reverse=True)[:6]
    patient_by_dept_top6 = [{"name": d["name"], "count": d["visits"]} for d in patient_by_dept_top6]

    # New Patient Admission (summary row)
    new_admission = None
    for rr, row in rows.items():
        if row and row[1] == "Admission Date":
            data_row = rows.get(rr + 1)
            if data_row:
                new_admission = {"total": data_row[8], "thai": data_row[9], "inter": data_row[10],
                                  "new": data_row[13], "old": data_row[18]}
            break

    # IPD Discharge (summary row) — row ที่มี label "Discharge" เป็นหัวคอลัมน์ย่อย
    # ข้อมูลจริงอยู่แถวถัดไป (เหมือน Admission Date ที่ข้อมูลอยู่แถว rr+1)
    discharge = None
    for rr, row in rows.items():
        if row and row[1] == "Discharge":
            data_row = rows.get(rr + 1)
            if data_row:
                discharge = {"count": data_row[1], "total_bill": data_row[10], "avg": data_row[13]}
            break

    # Current on ward — Total row
    bed_flow = None
    for rr, row in rows.items():
        if row and row[1] == "Total" and rr > 130:
            bed_flow = {"yesterday": row[10], "discharge": row[13], "new_admission": row[16],
                        "transfer": row[18], "today": row[23]}
            break

    # ---- CH1335: hourly registration ----
    # ไฟล์นี้ export มาจากเครื่อง ไม่ได้มาทาง Gmail — ถ้าหาไม่เจอ (เช่น export ไม่ทันเวลา)
    # ให้ปล่อยเป็น None ทั้งชุด ไม่ crash ทั้ง pipeline เพราะไฟล์เดียวหาย
    hourly = None
    if ch1335_path:
        wb2 = openpyxl.load_workbook(ch1335_path, data_only=True)
        ws2 = wb2.worksheets[0]
        hourly = [0] * 24
        for row in ws2.iter_rows(min_row=3, values_only=True):
            if row[0] is None:
                continue
            h = int(str(row[0]).split(":")[0])
            hourly[h] = row[1] or 0

    # ---- CH1434: revenue per visit (Total row) ----
    revenue_per_visit = None
    if ch1434_path:
        wb3 = openpyxl.load_workbook(ch1434_path, data_only=True)
        ws3 = wb3.worksheets[0]
        for row in ws3.iter_rows(min_row=3, values_only=True):
            if row[3] == "Total":
                revenue_per_visit = {"today": row[0], "last_year": row[1] if row[1] != "-" else None,
                                      "growth": row[2] if row[2] != "-" else None}
                break

    # ---- CH1288: wait time flow ----
    flow = {}
    if ch1288_path:
        wb4 = openpyxl.load_workbook(ch1288_path, data_only=True)
        ws4 = wb4.worksheets[0]
        keys = ["reg", "queue", "wait_dr", "doctor", "nursing", "cashier", "pharmacy", "total"]
        group_map = {"Non-ER,With Lab&X-ray": "non_er_lab", "Non-ER,No Lab&X-ray": "non_er_nolab"}
        for row in ws4.iter_rows(min_row=3, values_only=True):
            label = row[0]
            if label not in group_map:
                continue
            key = group_map[label]
            vals = [row[1], row[3], row[5], row[7], row[9], row[11], row[13], row[15]]
            g = {}
            for k, val in zip(keys, vals):
                if val is None:
                    g[k] = None
                    g[k + "_t"] = None
                    continue
                v_str, t_str = val.split("/")
                g[k] = float(v_str) if v_str.strip() != "" else None
                g[k + "_t"] = float(t_str) if t_str.strip() != "" else None
            flow[key] = g

    opd_billing = summary["total_bill_amount"]
    ipd_ward = ward_total
    total_revenue = opd_billing + ipd_ward
    opd_share = round(opd_billing / total_revenue * 100) if total_revenue else 0
    ipd_share = 100 - opd_share
    overview = {"opd_billing": opd_billing, "ipd_ward": ipd_ward, "total_revenue": total_revenue,
                "opd_share": opd_share, "ipd_share": ipd_share, "or_revenue": or_total}

    D = {
        "date": target_date.strftime("%Y-%m-%d"),
        "hourly": hourly,
        "depts": depts,
        "dept_totals": {"rev": dept_rev_total, "forecast": dept_forecast_total, "achieve": dept_achieve,
                         "mtd_rev": dept_mtd_rev_total, "mtd_forecast": dept_mtd_forecast_total,
                         "mtd_achieve": dept_mtd_achieve},
        "bu_opd_top5": bu_opd_top5,
        "bu_opd_total": dept_rev_total,
        "summary": summary,
        "revenue": revenue,
        "ward": ward,
        "bed_flow": bed_flow,
        "new_admission": new_admission,
        "discharge": discharge,
        "revenue_per_visit": revenue_per_visit,
        "flow": flow,
        "patient_by_dept_top6": patient_by_dept_top6,
        "overview": overview,
    }
    return D


# --------------------------------------------------------------------------
# GitHub: อ่าน/เขียนไฟล์ผ่าน Contents API
# --------------------------------------------------------------------------
def gh_get_file(repo, path, token):
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
    })
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            data = json.loads(r.read())
        content = base64.b64decode(data["content"]).decode("utf-8")
        return content, data["sha"]
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None, None
        raise


def gh_put_file(repo, path, content_str, token, message, sha=None):
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    body = {
        "message": message,
        "content": base64.b64encode(content_str.encode("utf-8")).decode("ascii"),
    }
    if sha:
        body["sha"] = sha
    req = urllib.request.Request(
        url, data=json.dumps(body).encode("utf-8"), method="PUT",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())


def push_day_data(D, repo, token):
    date_str = D["date"]
    path = f"data/overview/{date_str}.json"
    content_str = json.dumps(D, ensure_ascii=False)
    _, sha = gh_get_file(repo, path, token)
    gh_put_file(repo, path, content_str, token, f"Add overview data for {date_str}", sha)
    print(f"✅ pushed {path}")

    # update index.json
    index_path = "data/overview/index.json"
    idx_content, idx_sha = gh_get_file(repo, index_path, token)
    dates = json.loads(idx_content) if idx_content else []
    if date_str not in dates:
        dates.append(date_str)
    dates = sorted(set(dates))
    gh_put_file(repo, index_path, json.dumps(dates), token, f"Update index with {date_str}", idx_sha)
    print(f"✅ updated {index_path} ({len(dates)} dates)")


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="YYYY-MM-DD (default: yesterday)")
    ap.add_argument("--cowork-dir", required=True)
    ap.add_argument("--gmail-user", required=True)
    ap.add_argument("--gmail-password", required=True)
    ap.add_argument("--gh-token", required=True)
    ap.add_argument("--gh-repo", default="Adisornpatradul/wmc-dashboard")
    ap.add_argument("--no-push", action="store_true", help="คำนวณอย่างเดียว ไม่ push GitHub (สำหรับทดสอบ)")
    ap.add_argument("--out", help="เขียน JSON ผลลัෞธ์ลงไฟล์นี้ด้วย (สำหรับทดสอบ)")
    args = ap.parse_args()

    if args.date:
        target_date = datetime.strptime(args.date, "%Y-%m-%d")
    else:
        target_date = datetime.now() - timedelta(days=1)

    print(f"📅 target date: {target_date.strftime('%Y-%m-%d')}")

    report16_path = fetch_report16(args.gmail_user, args.gmail_password, target_date, args.cowork_dir)
    if not report16_path:
        print("❌ ไม่พบอีเมลรายงานประจำวันของวันที่นี้ใน Gmail")
        sys.exit(1)
    print(f"✅ Report16: {report16_path}")

    ch1288 = find_ch_file(args.cowork_dir, "CH1288", target_date)
    ch1335 = find_ch_file(args.cowork_dir, "CH1335", target_date)
    ch1434 = find_ch_file(args.cowork_dir, "CH1434", target_date)
    print(f"   CH1288 (เวลารอคอฆ): {ch1288}")
    print(f"   CH1335 (ผู้ป่วยตามช่วงเวลา): {ch1335}")
    print(f"   CH1434 (รายได้ต่อ Visit): {ch1434}")

    if not (ch1288 and ch1335 and ch1434):
        print("⚠️  ไม่พบไฟล์ CH1288/CH1335/CH1434 ที่ตรงวันที่ในโฟลเดอร์ Cowork — จะสร้างเฉพาะข้อมูลจาก Report16")

    D = build_overview_data(report16_path, ch1288, ch1335, ch1434, target_date)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(D, f, ensure_ascii=False, indent=1)
        print(f"📝 wrote {args.out}")

    if not args.no_push:
        push_day_data(D, args.gh_repo, args.gh_token)


if __name__ == "__main__":
    main()
