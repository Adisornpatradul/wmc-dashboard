#!/usr/bin/env python3
"""
WMC BU Dashboard — Daily Data Builder (แบบ interactive เลือกวันที่/ช่วงวันที่)
สร้าง data/bu/{bu_slug}/{date}.json ต่อ Business Unit จาก Report16 (OPD Revenue Location +
New Patient Admission on Ward) และ CH1434 (รายได้เฉลี่ยต่อ Visit แยกตามแผนก สำหรับ cross-check)

รันบนคอมพิวเตอร์ผู้ใช้ (device_bash) เหมือน wmc_overview_builder.py เพราะต้องอ่าน CH1434 จากโฟลเดอร์ Cowork

วิธีใช้:
  python3 wmc_bu_builder.py --report16 "/path/WMCDailyManagementType16 (25-9-69).xlsx" \
      --ch1434 "/path/xxxx.xlsx" --date 2026-09-25 \
      --gh-token ghp_xxx --gh-repo Adisornpatradul/wmc-dashboard

ปกติจะถูกเรียกต่อจาก wmc_overview_builder.py ในสคริปต์เดียวกัน (ใช้ไฟล์ Report16/CH1434 ที่หามาแล้ว)
"""

import argparse
import base64
import json
import os
import re
import sys
import urllib.request
import urllib.error
from datetime import datetime, timedelta

EXCEL_PREFIX = "WMCDailyManagementType16"

# --------------------------------------------------------------------------
# นิยาม BU ที่ต้องทำ dashboard แบบ interactive (เริ่มจาก 3 BU รายได้สูงสุด)
# name_report16   = ชื่อที่ใช้ใน "OPD Revenue Location" table ของ Report16 (ต้องตรงเป๊ะ)
# match_keywords  = คำที่ใช้จับคู่กับชื่อใน CH1434 / Admission Department (ตัด "Department"/เลขนำหน้า/เว้นวรรคออกก่อนเทียบ)
# --------------------------------------------------------------------------
BU_CONFIG = [
    {
        "slug": "checkup",
        "name_report16": "Check up Center",
        "name_th": "เช็คอัพเซ็นเตอร์",
        "name_en": "Check up Center",
        "match_keywords": ["check up center", "checkup center"],
    },
    {
        "slug": "emergency",
        "name_report16": "Emergency",
        "name_th": "แผนกฉุกเฉิน",
        "name_en": "Emergency",
        "match_keywords": ["emergency"],
    },
    {
        "slug": "advanced-surgery",
        "name_report16": "Advanced Surgery",
        "name_th": "ศูนย์ศัลยกรรมขั้นสูง",
        "name_en": "Advanced Surgery",
        "match_keywords": ["advanced surgery"],
    },
]


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


def g(row, idx):
    """เข้าถึง row[idx] อย่างปลอดภัย — บางวัน openpyxl ตัดคอลัมน์ท้ายที่ว่างออก ทำให้ length ของแต่ละแถวไม่เท่ากัน
    ห้าม index ตรงๆ เด็ดขาด ต้องผ่านฟังก์ชันนี้เสมอ"""
    if row is None or idx >= len(row):
        return None
    return row[idx]


def norm_label(s):
    """ตัด 'Department'/'Deparment', เลขนำหน้า, วงเล็บ, เว้นวรรคซ้ำ ออก แล้ว lowercase เพื่อจับคู่ชื่อ"""
    if not s:
        return ""
    s = re.sub(r"^\d+\s*", "", s)  # ตัดเลขนำหน้า เช่น "01 "
    s = re.sub(r"\(.*?\)", "", s)  # ตัดวงเล็บ
    s = re.sub(r"depar?tment\.?", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def find_dept_table(rows):
    """คืน list ของ {name, visits, forecast, rev, achieve, mtd_forecast, mtd_rev, mtd_achieve}
    จาก 'OPD Revenue Location' table ของ Report16 (เหมือน wmc_overview_builder.py)"""
    header_row = None
    for rr, row in rows.items():
        if row and g(row, 1) == "OPD Revenue Location":
            header_row = rr
            break
    depts = []
    if header_row:
        r = header_row + 2
        while True:
            row = rows.get(r)
            if row is None or g(row, 1) in (None, ""):
                break
            depts.append({
                "name": g(row, 1), "visits": g(row, 9), "forecast": g(row, 13), "rev": g(row, 16),
                "achieve": pct(g(row, 22)), "mtd_forecast": g(row, 26), "mtd_rev": g(row, 29),
                "mtd_achieve": pct(g(row, 34)),
            })
            r += 1
    return depts


def find_admission_items(rows, match_keywords):
    """เดินหา block 'New Patient Admission on Ward' แล้วดึงรายการที่ Admission Department
    ตรงกับ BU (จับคู่ผ่าน match_keywords) กลุ่มตาม ward -> คืน list [{ward, department, count}]"""
    section_start = None
    for rr, row in rows.items():
        if row and g(row, 8) == "New Patient Admission on  Ward":
            section_start = rr
            break
    if section_start is None:
        return []

    items = []
    current_ward = None
    r = section_start + 1
    blank_streak = 0
    while True:
        row = rows.get(r)
        if row is None:
            blank_streak += 1
            if blank_streak > 2:
                break
            r += 1
            continue
        # แถวหัว ward: col1 มีชื่อ ward, col12 = "Total New Patient Admission on Ward >>"
        if g(row, 1) not in (None, ""):
            current_ward = str(g(row, 1)).strip()
            blank_streak = 0
            r += 1
            continue
        # แถวรายละเอียดแผนก: col12 = ชื่อแผนก, col18 = จำนวน
        dept_label = g(row, 12)
        if dept_label and dept_label != "Total New patient admission >>":
            norm = norm_label(dept_label)
            if any(kw in norm for kw in match_keywords):
                count = g(row, 18)
                items.append({"ward": current_ward, "department": str(dept_label).strip(),
                              "count": clean_num(count) or 0})
            blank_streak = 0
        else:
            blank_streak += 1
            if blank_streak > 2:
                break
        r += 1

    return items


def find_ch1434_row(ch1434_path, match_keywords):
    import openpyxl
    if not ch1434_path:
        return None
    wb = openpyxl.load_workbook(ch1434_path, data_only=True)
    ws = wb.worksheets[0]
    for row in ws.iter_rows(min_row=3, values_only=True):
        label = g(row, 3)
        if not label:
            continue
        norm = norm_label(label)
        if any(kw in norm for kw in match_keywords):
            today = g(row, 0) if g(row, 0) != "-" else None
            last_year = g(row, 1) if g(row, 1) != "-" else None
            growth = g(row, 2) if g(row, 2) != "-" else None
            return {"per_visit": today, "last_year": last_year, "yoy": growth,
                    "matched_label": str(label).strip()}
    return None


def build_bu_data(report16_path, ch1434_path, target_date, bu_conf):
    import openpyxl

    wb = openpyxl.load_workbook(report16_path, data_only=True)
    ws = wb.worksheets[0]
    rows = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), start=2):
        rows[i] = row

    depts = find_dept_table(rows)
    dept = next((d for d in depts if d["name"] == bu_conf["name_report16"]), None)

    daily = None
    if dept:
        target = dept["forecast"]
        actual = dept["rev"]
        gap = (actual - target) if (actual is not None and target is not None) else None
        # หมายเหตุ: ห้ามใช้ dept["achieve"] ตรงๆ — cell นั้นใน Report16 คือ "gap%" (ส่วนต่างจากเป้า)
        # ไม่ใช่ "achieve%" (ผลงานเทียบเป้า) ต้องคำนวณเองจาก actual/target*100 เสมอ
        achieve = round(actual / target * 100) if (actual is not None and target) else None
        daily = {
            "visits": dept["visits"], "target": target, "actual": actual,
            "achieve": achieve, "gap": gap,
        }

    ch1434_row = find_ch1434_row(ch1434_path, bu_conf["match_keywords"])
    own_per_visit = None
    if dept and dept.get("rev") is not None and dept.get("visits"):
        own_per_visit = round(dept["rev"] / dept["visits"], 2) if dept["visits"] else None
    billing = {
        "per_visit_own": own_per_visit,  # rev/visits คำนวณเองจาก OPD Revenue Location
        "ch1434": ch1434_row,  # None ถ้าไม่พบใน CH1434 (บาง BU ไม่ถูก track แบบต่อ Visit)
    }

    admission_items = find_admission_items(rows, bu_conf["match_keywords"])
    admission_total = sum(i["count"] for i in admission_items)

    D = {
        "date": target_date.strftime("%Y-%m-%d"),
        "meta": {"slug": bu_conf["slug"], "name_th": bu_conf["name_th"], "name_en": bu_conf["name_en"]},
        "daily": daily,
        "billing": billing,
        "admission": {"total": admission_total, "items": admission_items},
    }
    return D


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
    body = {"message": message, "content": base64.b64encode(content_str.encode("utf-8")).decode("ascii")}
    if sha:
        body["sha"] = sha
    req = urllib.request.Request(
        url, data=json.dumps(body).encode("utf-8"), method="PUT",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json",
                 "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())


def push_bu_day_data(D, slug, repo, token):
    date_str = D["date"]
    path = f"data/bu/{slug}/{date_str}.json"
    content_str = json.dumps(D, ensure_ascii=False)
    _, sha = gh_get_file(repo, path, token)
    gh_put_file(repo, path, content_str, token, f"Add BU {slug} data for {date_str}", sha)
    print(f"✅ pushed {path}")

    index_path = f"data/bu/{slug}/index.json"
    idx_content, idx_sha = gh_get_file(repo, index_path, token)
    dates = json.loads(idx_content) if idx_content else []
    if date_str not in dates:
        dates.append(date_str)
    dates = sorted(set(dates))
    gh_put_file(repo, index_path, json.dumps(dates), token, f"Update BU {slug} index with {date_str}", idx_sha)
    print(f"✅ updated {index_path} ({len(dates)} dates)")


def resolve_report16_path(cowork_dir, target_date):
    """path เดียวกันเป๊ะๆ กับที่ wmc_overview_builder.py เซฟไว้ตอน fetch จาก Gmail
    (เรียก bu_builder ต่อจาก overview_builder ในการรันเดียวกัน ไฟล์นี้จะมีอยู่แล้วเสมอ)"""
    name = f"{EXCEL_PREFIX} ({target_date.day}-{target_date.month}-{(target_date.year + 543) % 100}).xlsx"
    return os.path.join(cowork_dir, name)


def find_ch_file(cowork_dir, ch_code, target_date):
    """สำเนาของ find_ch_file ใน wmc_overview_builder.py (คัดลอกมาให้ในตัว ไม่พึ่ง import ข้ามสคริปต์
    เพราะ wmc_overview_builder.py ถูกดาวน์โหลดไปไว้ที่ /tmp ตอนรันจริง ไม่ใช่ใน cowork_dir)
    ต้องเปิดดูข้างในเพื่อเช็ควันที่ (แถวแรกของ sheet Document_CH#### เป็น datetime ของวันที่ข้อมูล)"""
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
    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="YYYY-MM-DD (default: yesterday)")
    ap.add_argument("--cowork-dir", default=None,
                     help="ถ้าระบุ จะหา Report16 (path เดียวกับที่ wmc_overview_builder.py เซฟไว้) และ CH1434 ให้เอง")
    ap.add_argument("--report16", default=None, help="ระบุตรงๆ ถ้าไม่ใช้ --cowork-dir")
    ap.add_argument("--ch1434", default=None, help="ระบุตรงๆ ถ้าไม่ใช้ --cowork-dir")
    ap.add_argument("--gh-token", required=True)
    ap.add_argument("--gh-repo", default="Adisornpatradul/wmc-dashboard")
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--out-dir", default=None)
    args = ap.parse_args()

    if args.date:
        target_date = datetime.strptime(args.date, "%Y-%m-%d")
    else:
        target_date = datetime.now() - timedelta(days=1)

    print(f"📅 BU builder target date: {target_date.strftime('%Y-%m-%d')}")

    report16_path = args.report16
    ch1434_path = args.ch1434
    if args.cowork_dir:
        report16_path = report16_path or resolve_report16_path(args.cowork_dir, target_date)
        if not os.path.exists(report16_path):
            print(f"❌ ไม่พบ Report16 ที่ {report16_path} — ต้องรัน wmc_overview_builder.py ก่อนเสมอ (เพื่อ fetch จาก Gmail)")
            sys.exit(1)
        if not ch1434_path:
            ch1434_path = find_ch_file(args.cowork_dir, "CH1434", target_date)
        print(f"✅ Report16: {report16_path}")
        print(f"   CH1434: {ch1434_path}")

    args.report16 = report16_path
    args.ch1434 = ch1434_path

    for bu_conf in BU_CONFIG:
        D = build_bu_data(args.report16, args.ch1434, target_date, bu_conf)
        print(f"   [{bu_conf['slug']}] daily={D['daily']} admission_total={D['admission']['total']}")
        if args.out_dir:
            with open(os.path.join(args.out_dir, f"{bu_conf['slug']}_{D['date']}.json"), "w", encoding="utf-8") as f:
                json.dump(D, f, ensure_ascii=False, indent=1)
        if not args.no_push:
            push_bu_day_data(D, bu_conf["slug"], args.gh_repo, args.gh_token)


if __name__ == "__main__":
    main()
