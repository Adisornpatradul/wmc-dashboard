#!/usr/bin/env python3
"""
WMC Excel Processor — แปลง Excel รายวันเป็น HTML Dashboard แล้ว Deploy GitHub Pages
ใช้: python3 wmc_excel_processor.py <excel_file.xlsx> [YYYY-MM-DD]

🔑 ต้องใส่ GitHub Personal Access Token ที่บรรทัด GITHUB_TOKEN ด้านล่าง
   สร้าง token ที่: https://github.com/settings/tokens/new
   ✅ เลือก scope: repo (Full control of private repositories)
   แล้ว copy token มาแทน YOUR_GITHUB_PAT_HERE
"""
import sys, os, json, time, re, base64
import urllib.request, urllib.error, urllib.parse

# ─── CONFIG ───────────────────────────────────────────────────────────────────
GITHUB_TOKEN  = os.getenv("WMC_GH_TOKEN", "")
GITHUB_REPO   = "Adisornpatradul/wmc-dashboard"
SITE_URL      = "https://adisornpatradul.github.io/wmc-dashboard/"
LINE_TOKEN    = os.getenv("WMC_LINE_TOKEN", "")
LINE_GROUPS   = [
    ("HOD IPD", "C5db527a782afe0656e0a56c7c300d4e0"),
    ("HEC WMC", "C3f444ad88df282e1fd0853c63442d097"),
    ("Head BU", "C27d447816a22cb3b8248feb1bf684dc9"),
]
OUTPUT_DIR    = os.path.dirname(os.path.abspath(__file__))
def get_dashboard_html(date_str):
    """คืนชื่อไฟล์ HTML ที่ถูกต้องตามเดือน/ปีของ date_str"""
    from datetime import datetime
    th_months_en = ["", "January", "February", "March", "April", "May",
                    "June", "July", "August", "September",
                    "October", "November", "December"]
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        name = f"WMC_Monthly_Dashboard_{th_months_en[dt.month]}{dt.year}.html"
    except Exception:
        name = "WMC_Monthly_Dashboard_June2026.html"
    return os.path.join(OUTPUT_DIR, name)

DASHBOARD_HTML = os.path.join(OUTPUT_DIR, "WMC_Monthly_Dashboard_June2026.html")
# ──────────────────────────────────────────────────────────────────────────────


def install_openpyxl():
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl",
                    "--break-system-packages", "-q"], check=False)


def _num(v):
    """แปลง cell value เป็น float"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def extract_data_from_workbook(wb, date_str):
    """
    Dynamic extractor — ค้นหา section headers ด้วย text search
    แล้วอ่าน column positions จาก header row ของแต่ละ section
    รองรับทั้ง layout June-16 และ June-17 (และ layout อื่นในอนาคต)
    """
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    n = len(rows)
    g = _num

    def safe(r, c):
        try:
            return g(rows[r][c])
        except Exception:
            return 0.0

    def cell_str(r, c):
        try:
            v = rows[r][c]
            return str(v).strip() if v is not None else ""
        except Exception:
            return ""

    def find_row(text, start=0, end=None, ci=None):
        """หา row index แรกที่มี text (case-insensitive) อยู่ใน cell ใดก็ได้ หรือใน col ci"""
        end = end or n
        tl = text.lower()
        for i in range(start, min(end, n)):
            cols = [ci] if ci is not None else range(len(rows[i]))
            for j in cols:
                try:
                    if tl in str(rows[i][j] or "").lower():
                        return i
                except Exception:
                    pass
        return -1

    def find_col(row_idx, text):
        """หา column index แรกใน row ที่มี text"""
        tl = text.lower()
        try:
            for j, v in enumerate(rows[row_idx]):
                if tl in str(v or "").lower():
                    return j
        except Exception:
            pass
        return -1

    # ══════════════════════════════════════════════════════════════════════
    # 1. VISIT SUMMARY (Row 9=header, Row 10=data)
    # ══════════════════════════════════════════════════════════════════════
    hdr = find_row("Total Visit", start=5, end=15)   # header row index
    if hdr < 0:
        hdr = 8
    dat = hdr + 1                                      # data row index

    total_visit   = int(safe(dat, 1))                 # col B always
    col_ipd   = find_col(hdr, "IPD")
    col_opd   = find_col(hdr, "OPD")
    col_new   = find_col(hdr, "New")
    col_old   = find_col(hdr, "Old")
    col_bill  = find_col(hdr, "Bill Amount")
    col_avg   = find_col(hdr, "AVG Per Bill")
    # Counter Visit header may be in row above data header
    col_cnt = find_col(hdr, "Total")          # "Total" under Counter Visit group
    if col_cnt < 0 or col_cnt == 1:           # avoid mistaking "Total Visit" col B
        col_cnt = find_col(max(0, hdr - 1), "Counter")

    ipd_admit     = int(safe(dat, col_ipd))  if col_ipd  >= 0 else 0
    opd_visit     = int(safe(dat, col_opd))  if col_opd  >= 0 else 0
    new_pt        = int(safe(dat, col_new))  if col_new  >= 0 else 0
    old_pt        = int(safe(dat, col_old))  if col_old  >= 0 else 0
    bill_amt      = safe(dat, col_bill)      if col_bill >= 0 else 0.0
    avg_bill      = safe(dat, col_avg)       if col_avg  >= 0 else 0.0
    counter_visit = int(safe(dat, col_cnt))  if col_cnt  >= 0 else 0
    opd_total     = bill_amt

    # ══════════════════════════════════════════════════════════════════════
    # 2 & 2b. IPD REVENUE + WARD REVENUE (Location/Amount table)
    #   Layout Aug-2026+: the ward Location/Amount table appears directly
    #   after the Visit Summary, WITHOUT an "IPD Revenue" text header above
    #   it. Older layouts (June/July-2026) still have that label. We search
    #   for the Location+Amount header row directly so it works either way,
    #   and derive ipdRev as the sum of all ward rows (matches the sheet's
    #   own unlabeled total row, which we skip since it has no location text).
    # ══════════════════════════════════════════════════════════════════════
    ipd_rev = 0.0
    ward_rev_map = {}   # {ward_name_raw: revenue_float}

    ipd_sec = find_row("IPD Revenue", start=8, end=25)
    search_start = (ipd_sec + 1) if ipd_sec >= 0 else (dat + 1)
    loc_hdr_row = -1
    for ri in range(search_start, min(search_start + 25, 40, n)):
        if find_col(ri, "Location") >= 0 and find_col(ri, "Amount") >= 0:
            loc_hdr_row = ri
            break

    if loc_hdr_row >= 0:
        c_loc = find_col(loc_hdr_row, "Location")
        c_rev = find_col(loc_hdr_row, "Amount")
        if c_loc >= 0 and c_rev >= 0:
            for ri in range(loc_hdr_row + 1, min(loc_hdr_row + 12, n)):
                loc_name = cell_str(ri, c_loc)
                if not loc_name:
                    continue
                if "total" in loc_name.lower():
                    break
                rev_val = g(rows[ri][c_rev]) if len(rows[ri]) > c_rev else 0.0
                if rev_val > 0:
                    ward_rev_map[loc_name.strip()] = rev_val

    if ward_rev_map:
        ipd_rev = sum(ward_rev_map.values())

    # Fallback for old layouts where IPD Revenue was a single number (no ward table)
    if not ipd_rev and ipd_sec >= 0:
        date_hdr = find_row("Date", start=ipd_sec + 1, end=ipd_sec + 4)
        if date_hdr >= 0:
            r = rows[date_hdr + 1] if date_hdr + 1 < n else []
            for ci in range(2, min(15, len(r))):
                v = r[ci]
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 50000:
                    ipd_rev = float(v)
                    break
        if not ipd_rev:
            for ri in range(ipd_sec + 1, min(ipd_sec + 25, n)):
                r = rows[ri]
                if not r:
                    continue
                if any(hasattr(v, "year") for v in r[:8]):
                    continue
                if any(isinstance(v, str) and v.strip() for v in r[:8]):
                    continue
                for ci in range(2, min(22, len(r))):
                    v = r[ci]
                    if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 50000:
                        ipd_rev = float(v)
                        break
                if ipd_rev:
                    break

    # ══════════════════════════════════════════════════════════════════════
    # 2c. OPERATING ROOM REVENUE (OT) — new section added Aug-2026
    #     Table: "Order Date | Order From Location | Total Net Amount"
    # ══════════════════════════════════════════════════════════════════════
    ot_rev = 0.0
    ot_hdr_row = find_row("Order From Location", start=10, end=45)
    if ot_hdr_row >= 0:
        c_amt = find_col(ot_hdr_row, "Total Net Amount")
        if c_amt < 0:
            c_amt = find_col(ot_hdr_row, "Net Amount")
        c_loc = find_col(ot_hdr_row, "Order From Location")
        if c_amt >= 0:
            for ri in range(ot_hdr_row + 1, min(ot_hdr_row + 10, n)):
                loc = cell_str(ri, c_loc) if c_loc >= 0 else ""
                if not loc:
                    break
                if "total" in loc.lower():
                    ot_rev = g(rows[ri][c_amt]) if len(rows[ri]) > c_amt else ot_rev
                    break
                ot_rev += g(rows[ri][c_amt]) if len(rows[ri]) > c_amt else 0.0

    # ══════════════════════════════════════════════════════════════════════
    # 3. OPD REVENUE BY NATIONALITY
    # ══════════════════════════════════════════════════════════════════════
    opd_thai = opd_inter = 0.0
    opd_nat_sec = find_row("OPD Revenue", start=20, end=40)
    if opd_nat_sec >= 0:
        for ri in range(opd_nat_sec, min(opd_nat_sec + 15, n)):
            # Thai/Inter label อาจอยู่ col 0 หรือ col 1 ขึ้นกับ layout
            label = (cell_str(ri, 0) or cell_str(ri, 1) or cell_str(ri, 2)).lower()
            # find the amount column: first col 1-13 with large numeric
            if "thai" in label:
                for ci in range(1, 14):
                    v = g(rows[ri][ci]) if len(rows[ri]) > ci else 0
                    if v > 1000:
                        opd_thai = v; break
            elif "inter" in label:
                for ci in range(1, 14):
                    v = g(rows[ri][ci]) if len(rows[ri]) > ci else 0
                    if v > 1000:
                        opd_inter = v; break

    # ══════════════════════════════════════════════════════════════════════
    # 4. CHECKUP TOTAL
    # ══════════════════════════════════════════════════════════════════════
    chk_total = 0.0
    chk_sec = find_row("Checkup Revenue", start=30, end=55)
    if chk_sec >= 0:
        for ri in range(chk_sec, min(chk_sec + 15, n)):
            if cell_str(ri, 0).lower() == "total":
                for ci in range(5, 13):
                    v = g(rows[ri][ci]) if len(rows[ri]) > ci else 0
                    if v > 1000:
                        chk_total = v; break
                break

    # ══════════════════════════════════════════════════════════════════════
    # 5. BU OPD REVENUE LOCATION
    # ══════════════════════════════════════════════════════════════════════
    bu_list = []
    bu_sec = find_row("OPD Revenue Location", start=45, end=80)
    # Find where the BU section ends (next section starts)
    sup_sec_approx = find_row("Revenue by Support location", start=bu_sec + 1 if bu_sec >= 0 else 75, end=115)
    bu_end = sup_sec_approx if sup_sec_approx > 0 else (bu_sec + 35 if bu_sec >= 0 else 0)

    if bu_sec >= 0:
        bu_hdr = rows[bu_sec + 1]   # header row with column labels
        # find column positions
        c_cnt  = -1; c_mfc = -1; c_dfc = -1; c_drev = -1; c_mtd = -1
        for ci, v in enumerate(bu_hdr):
            s = str(v or "").lower().replace("\n", " ")
            if ("counter" in s or "counter visit" in s) and c_cnt < 0:
                c_cnt = ci
            elif "monthly forecast" in s and c_mfc < 0:
                c_mfc = ci
            elif "daily forecast" in s and c_dfc < 0:
                c_dfc = ci
            elif "daily revenue" in s and c_drev < 0:
                c_drev = ci
            elif "mtd" in s and "rev" in s and c_mtd < 0:
                c_mtd = ci
        # fallback to known June-17 positions
        if c_cnt  < 0: c_cnt  = 9
        if c_mfc  < 0: c_mfc  = 11
        if c_dfc  < 0: c_dfc  = 12
        if c_drev < 0: c_drev = 14
        if c_mtd  < 0: c_mtd  = 27

        for ri in range(bu_sec + 2, min(bu_end, n)):
            row = rows[ri]
            # BU name is in col 0 or col 1
            name = str(row[0] or "").strip()
            if not name or name == "None":
                name = str(row[1] or "").strip()
            if not name or name.lower() in ("none", "", "location"):
                continue
            if any(k in name.lower() for k in ("total", "grand", "รวม", "header")):
                break
            # Skip support-section-like rows (Visit Type, InPatient, etc.)
            if any(k in name.lower() for k in ("visit type", "inpatient", "outpatient", "in patient", "out patient", "revenue by")):
                break
            try:
                bu_list.append({
                    "name":      name,
                    "visits":    int(g(row[c_cnt])),
                    "dailyRev":  g(row[c_drev]),
                    "dailyFc":   g(row[c_dfc]),
                    "mtdRev":    g(row[c_mtd]),
                    "monthlyFc": g(row[c_mfc]),
                })
            except Exception:
                continue

    # ══════════════════════════════════════════════════════════════════════
    # 6. SUPPORT LOCATION REVENUE
    # ══════════════════════════════════════════════════════════════════════
    sup_ip_kidney = sup_ip_imaging = sup_ip_rehab = 0.0
    sup_op_kidney = sup_op_imaging = sup_op_rehab = 0.0
    sup_mtd_ip_kidney = sup_mtd_ip_imaging = sup_mtd_ip_rehab = 0.0
    sup_mtd_op_kidney = sup_mtd_op_imaging = sup_mtd_op_rehab = 0.0

    sup_sec = find_row("Revenue by Support location", start=75, end=105)
    if sup_sec >= 0:
        sup_hdr = rows[sup_sec + 1]
        c_amt = -1; c_mtd_s = -1
        for ci, v in enumerate(sup_hdr):
            s = str(v or "").lower()
            if "amount" in s and c_amt < 0:
                c_amt = ci
            elif "mtd" in s and c_mtd_s < 0:
                c_mtd_s = ci
        if c_amt   < 0: c_amt   = 9
        if c_mtd_s < 0: c_mtd_s = 21

        in_ip = False; in_op = False
        for ri in range(sup_sec + 2, min(sup_sec + 15, n)):
            row = rows[ri]
            l0 = cell_str(ri, 0).lower()
            # location label is in col 7 (June-17) or col 6 (June-16)
            l_loc = ""
            for lci in [7, 6]:
                v = cell_str(ri, lci)
                if v:
                    l_loc = v.lower(); break

            if "inpatient" in l0 or "in patient" in l0:
                in_ip = True; in_op = False; continue
            elif "outpatient" in l0 or "out patient" in l0:
                in_ip = False; in_op = True; continue

            if not l_loc:
                continue
            amt = g(row[c_amt]) if len(row) > c_amt else 0
            mtd = g(row[c_mtd_s]) if len(row) > c_mtd_s else 0

            if "kidney" in l_loc or "dialysis" in l_loc or "11 " in l_loc:
                if in_ip:
                    sup_ip_kidney = amt; sup_mtd_ip_kidney = mtd
                else:
                    sup_op_kidney = amt; sup_mtd_op_kidney = mtd
            elif "imaging" in l_loc or "13 " in l_loc:
                if in_ip:
                    sup_ip_imaging = amt; sup_mtd_ip_imaging = mtd
                else:
                    sup_op_imaging = amt; sup_mtd_op_imaging = mtd
            elif "rehab" in l_loc or "16 " in l_loc:
                if in_ip:
                    sup_ip_rehab = amt; sup_mtd_ip_rehab = mtd
                else:
                    sup_op_rehab = amt; sup_mtd_op_rehab = mtd

    # Override BU revenue for 3 support BUs
    for bu in bu_list:
        nm = bu["name"].lower()
        if "kidney dialysis" in nm:
            bu["dailyRev"] = sup_ip_kidney + sup_op_kidney
            bu["mtdRev"]   = sup_mtd_ip_kidney + sup_mtd_op_kidney
        elif "imaging center" in nm:
            bu["dailyRev"] += sup_ip_imaging + sup_op_imaging
            bu["mtdRev"]   += sup_mtd_ip_imaging + sup_mtd_op_imaging
        elif "medical rehabilitation" in nm or "medical rehab" in nm:
            bu["dailyRev"] = sup_ip_rehab + sup_op_rehab
            bu["mtdRev"]   = sup_mtd_ip_rehab + sup_mtd_op_rehab

    # ══════════════════════════════════════════════════════════════════════
    # 7. CASHIER BILLED
    # ══════════════════════════════════════════════════════════════════════
    ipd_inter_bill = ipd_thai_bill = opd_inter_bill = opd_thai_bill = 0.0
    total_billed = 0.0; total_bills_n = 0

    cash_sec = find_row("Cashier Billed", start=85, end=115)
    if cash_sec >= 0:
        # detect column layout from header rows
        c_type = -1; c_nat = -1; c_camt = -1; c_ccnt = -1
        for ri in range(cash_sec, min(cash_sec + 4, n)):
            row = rows[ri]
            for ci, v in enumerate(row):
                s = str(v or "").lower()
                if "type" in s and c_type < 0:
                    c_type = ci
                elif "nationality" in s and c_nat < 0:
                    c_nat = ci
                elif "amount" in s and c_camt < 0:
                    c_camt = ci
                elif ("number of billed" in s or "number" in s) and c_ccnt < 0:
                    c_ccnt = ci
        if c_type < 0: c_type = 0
        if c_nat  < 0: c_nat  = 7
        if c_camt < 0: c_camt = 9
        if c_ccnt < 0: c_ccnt = 12

        in_ipd = False; in_opd = False
        for ri in range(cash_sec + 1, min(cash_sec + 15, n)):
            row = rows[ri]
            t0 = cell_str(ri, 0).lower()
            ttype = cell_str(ri, c_type).lower()
            tnat  = cell_str(ri, c_nat).lower()
            if t0 == "ipd":
                in_ipd = True; in_opd = False; continue
            elif t0 == "opd":
                in_ipd = False; in_opd = True; continue
            amt = g(row[c_camt]) if len(row) > c_camt else 0
            cnt = int(g(row[c_ccnt])) if len(row) > c_ccnt else 0
            if "inter" in tnat:
                if in_ipd: ipd_inter_bill = amt
                elif in_opd: opd_inter_bill = amt
            elif "thai" in tnat:
                if in_ipd: ipd_thai_bill = amt
                elif in_opd: opd_thai_bill = amt
            elif "total billed" in tnat and "day" in tnat:
                total_billed = amt; total_bills_n = cnt; break

    # ══════════════════════════════════════════════════════════════════════
    # 8. IPD DISCHARGE
    # ══════════════════════════════════════════════════════════════════════
    discharge_total = 0; discharge_bill = 0.0
    dc_sec = find_row("IPD Discharge", start=95, end=120)
    if dc_sec >= 0:
        for ri in range(dc_sec + 1, min(dc_sec + 8, n)):
            row = rows[ri]
            # check col 0 or col 1 for numeric discharge count
            for check_col in [0, 1]:
                v = row[check_col] if len(row) > check_col else None
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0:
                    discharge_total = int(v)
                    # bill amount: first large numeric in cols 6-15
                    for bci in range(6, 16):
                        bv = g(row[bci]) if len(row) > bci else 0
                        if bv > 10000:
                            discharge_bill = bv; break
                    break
            if discharge_total:
                break

    # ══════════════════════════════════════════════════════════════════════
    # 9. WARD ON WARD (Current On Ward)
    # ══════════════════════════════════════════════════════════════════════
    ward_on_ward = []; ward_total_on = 0
    ward_sec = find_row("Current On Ward", start=125, end=160)
    if ward_sec >= 0:
        # find header row with "Yesterday"
        ward_hdr = -1
        for ri in range(ward_sec, min(ward_sec + 5, n)):
            if find_col(ri, "Yesterday") >= 0:
                ward_hdr = ri; break

        if ward_hdr >= 0:
            c_yest  = find_col(ward_hdr, "Yesterday");  c_yest  = c_yest  if c_yest  >= 0 else 9
            c_dc    = find_col(ward_hdr, "Discharge");  c_dc    = c_dc    if c_dc    >= 0 else 12
            c_nadm  = find_col(ward_hdr, "Admission");  c_nadm  = c_nadm  if c_nadm  >= 0 else 14
            c_tr    = find_col(ward_hdr, "Transfer");   c_tr    = c_tr    if c_tr    >= 0 else 16
            c_today = find_col(ward_hdr, "Onward");
            if c_today < 0: c_today = find_col(ward_hdr, "Today")
            if c_today < 0: c_today = 21

            for ri in range(ward_hdr + 1, min(ward_hdr + 12, n)):
                row = rows[ri]
                name = cell_str(ri, 0) or cell_str(ri, 1)
                if not name: continue
                if "total" in name.lower():
                    ward_total_on = int(g(row[c_today]) if len(row) > c_today else 0)
                    break
                # Match ward revenue: compare first 2 words of name, with ICU special-case
                ward_rev = 0.0
                nm_lo = name.lower()
                nm_words = nm_lo.split()
                for rev_raw, rev_val in ward_rev_map.items():
                    rv_lo = rev_raw.lower()
                    rv_words = rv_lo.split()
                    if len(nm_words) >= 2 and len(rv_words) >= 2 and nm_words[:2] == rv_words[:2]:
                        ward_rev = rev_val; break
                    if ("intensive" in nm_lo or "icu" in nm_lo) and \
                       ("intensive" in rv_lo or "icu" in rv_lo):
                        ward_rev = rev_val; break
                ward_on_ward.append({
                    "name":      name,
                    "yesterday": int(g(row[c_yest])  if len(row) > c_yest  else 0),
                    "discharge": int(g(row[c_dc])    if len(row) > c_dc    else 0),
                    "newAdmit":  int(g(row[c_nadm])  if len(row) > c_nadm  else 0),
                    "transfer":  int(g(row[c_tr])    if len(row) > c_tr    else 0),
                    "today":     int(g(row[c_today]) if len(row) > c_today else 0),
                    "wardRev":   ward_rev,
                })
            if not ward_total_on:
                ward_total_on = sum(w["today"] for w in ward_on_ward)

    # ══════════════════════════════════════════════════════════════════════
    # 10. NEW PATIENT ADMISSION
    # ══════════════════════════════════════════════════════════════════════
    admit_thai = admit_inter = admit_new = admit_old = 0
    adm_sec = find_row("New Patient Admission", start=110, end=135)
    if adm_sec >= 0:
        # find sub-header row with Thai/Inter/New/Old
        for ri in range(adm_sec, min(adm_sec + 6, n)):
            row = rows[ri]
            has_thai  = any("thai"  in str(v or "").lower() for v in row)
            has_inter = any("inter" in str(v or "").lower() for v in row)
            if has_thai and has_inter:
                c_athai  = find_col(ri, "Thai")
                c_ainter = find_col(ri, "Inter")
                c_anew   = find_col(ri, "New")
                c_aold   = find_col(ri, "OLD") if find_col(ri, "OLD") >= 0 else find_col(ri, "Old")
                dr = rows[ri + 1]
                if c_athai  >= 0: admit_thai  = int(g(dr[c_athai]))
                if c_ainter >= 0: admit_inter = int(g(dr[c_ainter]))
                if c_anew   >= 0: admit_new   = int(g(dr[c_anew]))
                if c_aold   >= 0: admit_old   = int(g(dr[c_aold]))
                break

    return {
        "date":          date_str,
        "totalVisit":    total_visit,
        "opdVisit":      opd_visit,
        "ipdAdmit":      ipd_admit,
        "newPt":         new_pt,
        "oldPt":         old_pt,
        "counterVisit":  counter_visit,
        "billAmt":       bill_amt,
        "avgBill":       avg_bill,
        "opdTotal":      opd_total,
        "opdThai":       opd_thai,
        "opdInter":      opd_inter,
        "chkTotal":      chk_total,
        "ipdRev":        ipd_rev,
        "otRev":         ot_rev,
        "ipdInterBill":  ipd_inter_bill,
        "ipdThaiBill":   ipd_thai_bill,
        "opdInterBill":  opd_inter_bill,
        "opdThaiBill":   opd_thai_bill,
        "totalBilled":   total_billed,
        "totalBillsN":   total_bills_n,
        "wardTotal":     ward_total_on,
        "wardOnWard":    ward_on_ward,
        "dischargeTotal": discharge_total,
        "dischargeBill": discharge_bill,
        "admitTotal":    ipd_admit,
        "admitThai":     admit_thai,
        "admitInter":    admit_inter,
        "admitNew":      admit_new,
        "admitOld":      admit_old,
        "bu":            bu_list,
    }


def _create_new_month_html(date_str, new_path):
    """สร้าง Dashboard HTML ใหม่สำหรับเดือนนี้ โดย clone จากเดือนก่อนหน้า"""
    import calendar as _cal
    from datetime import datetime as _dt, date as _date
    th_months    = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม",
                    "มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    th_months_en = ["","January","February","March","April","May",
                    "June","July","August","September","October","November","December"]
    th_abbr      = ["","ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.",
                    "มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."]

    d = _dt.strptime(date_str, "%Y-%m-%d")
    yr, mo = d.year, d.month
    be = yr + 543
    days_in_month = _cal.monthrange(yr, mo)[1]

    # หา JS start_dow (0=Sun) ของวันที่ 1 ของเดือนนี้
    first_dow_py = _date(yr, mo, 1).weekday()   # 0=Mon
    start_dow_js = (first_dow_py + 1) % 7       # 0=Sun

    # หาไฟล์เดือนก่อนเป็น template
    prev_mo = mo - 1 if mo > 1 else 12
    prev_yr = yr if mo > 1 else yr - 1
    template_path = os.path.join(OUTPUT_DIR,
        f"WMC_Monthly_Dashboard_{th_months_en[prev_mo]}{prev_yr}.html")

    # fallback: ใช้ไฟล์ล่าสุดที่มีใน folder
    if not os.path.exists(template_path):
        candidates = sorted(
            [f for f in os.listdir(OUTPUT_DIR)
             if f.startswith("WMC_Monthly_Dashboard_") and f.endswith(".html")],
            reverse=True
        )
        if not candidates:
            print("❌ ไม่พบ template HTML สำหรับสร้างเดือนใหม่")
            return
        template_path = os.path.join(OUTPUT_DIR, candidates[0])

    html = open(template_path, encoding="utf-8").read()

    # เดือนเก่า (template)
    t = _dt.strptime(date_str, "%Y-%m-%d")
    prev_th  = th_months[prev_mo];    curr_th  = th_months[mo]
    prev_en  = th_months_en[prev_mo]; curr_en  = th_months_en[mo]
    prev_abbr= th_abbr[prev_mo];      curr_abbr= th_abbr[mo]
    prev_be  = prev_yr + 543

    html = html.replace(f"{prev_th} {prev_be}", f"{curr_th} {be}")
    html = html.replace(f"{prev_en} {prev_yr}", f"{curr_en} {yr}")
    html = html.replace(f"{prev_en}{prev_yr}",  f"{curr_en}{yr}")
    html = html.replace(prev_abbr, curr_abbr)

    # ล้าง const RAW
    import re as _re
    html = _re.sub(r"const RAW\s*=\s*\{.*?\};", "const RAW = {};", html, flags=_re.DOTALL)

    # อัปเดตปฏิทิน: จำนวนวัน, START_DOW, date prefix
    old_days = _cal.monthrange(prev_yr, prev_mo)[1]
    html = html.replace(f"d<={old_days}", f"d<={days_in_month}")

    old_start = ((_date(prev_yr, prev_mo, 1).weekday() + 1) % 7)
    html = html.replace(f"START_DOW={old_start}", f"START_DOW={start_dow_js}")

    old_prefix = f"{prev_yr:04d}-{prev_mo:02d}-"
    new_prefix = f"{yr:04d}-{mo:02d}-"
    html = html.replace(f"`{old_prefix}${{", f"`{new_prefix}${{"  )

    # แก้ WARD snapshot reference
    old_ward_ref = f"{prev_yr:04d}-{prev_mo:02d}-09"
    new_ward_ref = f"{yr:04d}-{mo:02d}-09"
    html = html.replace(f"dk==='{old_ward_ref}'", f"dk==='{new_ward_ref}'")

    with open(new_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"   ✅ สร้าง {os.path.basename(new_path)} สำเร็จ")


def update_dashboard_html(date_str, new_data):
    """
    อัปเดต const RAW ใน HTML โดยตรง — ไม่ rebuild ทั้งไฟล์
    ทำให้ style / structure ไม่เปลี่ยนแปลง
    รองรับหลายเดือนอัตโนมัติ
    """
    from datetime import datetime as _dt
    th_months_abbr = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.",
                      "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    try:
        d = _dt.strptime(date_str, "%Y-%m-%d")
        month_abbr = th_months_abbr[d.month]
        be_year = d.year + 543
    except Exception:
        month_abbr = "มิ.ย."
        be_year = 2569

    html_path = get_dashboard_html(date_str)
    global DASHBOARD_HTML
    DASHBOARD_HTML = html_path

    # ถ้าไฟล์เดือนนี้ยังไม่มี → สร้างใหม่จากเดือนก่อนหน้าอัตโนมัติ
    if not os.path.exists(html_path):
        from datetime import datetime as _dt2
        _day = _dt2.strptime(date_str, "%Y-%m-%d").day
        if _day != 1:
            print(f"   ⚠️  คำเตือน: ไม่พบ {os.path.basename(html_path)} แต่วันที่ {date_str} ไม่ใช่วันที่ 1 ของเดือน")
            print(f"   ⚠️  นี่อาจไม่ใช่เดือนใหม่จริง — อาจเกิดจากดาวน์โหลดไฟล์เดิมล้มเหลว การสร้างไฟล์ใหม่จะทำให้ข้อมูลวันก่อนหน้าหายไป")
        print(f"   ℹ️  ไม่พบ {os.path.basename(html_path)} — สร้างใหม่สำหรับเดือนนี้...")
        _create_new_month_html(date_str, html_path)

    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    m = re.search(r"const RAW\s*=\s*(\{.*?\});", html, re.DOTALL)
    if not m:
        print("❌ ไม่พบ const RAW ใน HTML")
        return False

    raw = json.loads(m.group(1))
    raw[date_str] = new_data

    new_raw_js = json.dumps(raw, ensure_ascii=False, separators=(",", ":"))
    new_html = html[: m.start()] + f"const RAW = {new_raw_js};" + html[m.end() :]

    # อัปเดต day-count label ใน header (รองรับทุกเดือน)
    n = len(raw)
    last_day = max(int(d.split("-")[2]) for d in raw.keys())
    new_html = re.sub(
        r"ข้อมูล \d+ วัน \([^)]+\)",
        f"ข้อมูล {n} วัน (1–{last_day} {month_abbr} {be_year})",
        new_html,
    )

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    print(f"✅ อัปเดต Dashboard HTML สำเร็จ ({len(new_html):,} bytes, {n} วัน)")
    return True


def _gh_upload(filename, local_path, commit_msg):
    """Helper: upload a single file to GitHub via Contents API"""
    from datetime import datetime
    gh_headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filename}"

    with open(local_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    # ดึง SHA (จำเป็นสำหรับ update ไฟล์ที่มีอยู่แล้ว)
    req = urllib.request.Request(api_url, headers=gh_headers)
    try:
        with urllib.request.urlopen(req) as r:
            sha = json.loads(r.read())["sha"]
    except urllib.error.HTTPError as e:
        sha = None if e.code == 404 else (_ for _ in ()).throw(e)

    payload = {"message": commit_msg, "content": content_b64}
    if sha:
        payload["sha"] = sha

    req2 = urllib.request.Request(
        api_url,
        data=json.dumps(payload).encode(),
        headers={**gh_headers, "Content-Type": "application/json"},
        method="PUT",
    )
    with urllib.request.urlopen(req2) as r:
        result = json.loads(r.read())
    return result.get("commit", {}).get("sha", "")[:10]


def deploy_to_github_pages():
    """Deploy ไป GitHub Pages:
       - อัปโหลด monthly dashboard เป็น WMC_Monthly_Dashboard_<Month><Year>.html
         (Launcher index.html อ่านชื่อนี้โดยอัตโนมัติ — cache-busting ?v=YYYYMMDD)
    """
    from datetime import datetime

    monthly_filename = os.path.basename(DASHBOARD_HTML)   # e.g. WMC_Monthly_Dashboard_July2026.html
    commit_msg = f"Update WMC Dashboard {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    print(f"🔍 กำลัง Push {monthly_filename} ไป GitHub Pages...")
    commit = _gh_upload(monthly_filename, DASHBOARD_HTML, commit_msg)
    print(f"✅ Deploy สำเร็จ → {SITE_URL}")
    print(f"   Commit: {commit}")
    return SITE_URL


def send_line_messages(date_str, site_url):
    """ส่ง Flex Message ไปทุก LINE group"""
    from datetime import datetime

    th_months = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม",
                 "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน",
                 "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        date_th = f"{dt.day} {th_months[dt.month]} {dt.year + 543}"
    except Exception:
        date_th = date_str

    flex = {
        "type": "flex",
        "altText": f"🏥 WMC Daily Report — {date_th}",
        "contents": {
            "type": "bubble",
            "size": "giga",
            "header": {
                "type": "box", "layout": "vertical",
                "backgroundColor": "#1E7E34",
                "paddingAll": "16px",
                "contents": [
                    {"type": "text", "text": "🏥 World Medical Hospital",
                     "color": "#FFFFFF", "weight": "bold", "size": "md"},
                    {"type": "text", "text": f"📊 Daily Report — {date_th}",
                     "color": "#E8F5E9", "size": "sm", "margin": "xs"},
                ],
            },
            "body": {
                "type": "box", "layout": "vertical", "paddingAll": "16px",
                "contents": [
                    {"type": "text",
                     "text": "ดูข้อมูลผู้ป่วย OPD/IPD รายวัน รายได้แต่ละ BU "
                             "และ Monthly Calendar แบบ Interactive",
                     "size": "sm", "color": "#444444", "wrap": True},
                    {"type": "button",
                     "action": {"type": "uri", "label": "📊 เปิด Dashboard",
                                "uri": site_url},
                     "style": "primary", "color": "#1E7E34", "margin": "lg"},
                ],
            },
            "footer": {
                "type": "box", "layout": "vertical",
                "backgroundColor": "#F5F5F5", "paddingAll": "10px",
                "contents": [
                    {"type": "text",
                     "text": "อัปเดตอัตโนมัติทุกเช้า 07:30 น. | WMC Smart Hospital",
                     "size": "xs", "color": "#888888", "align": "center"},
                ],
            },
        },
    }

    for group_name, group_id in LINE_GROUPS:
        body = json.dumps({"to": group_id, "messages": [flex]}).encode("utf-8")
        req = urllib.request.Request(
            "https://api.line.me/v2/bot/message/push",
            data=body,
            headers={"Authorization": f"Bearer {LINE_TOKEN}",
                     "Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req):
                print(f"✅ LINE → {group_name}")
        except Exception as e:
            print(f"❌ LINE → {group_name}: {e}")


def parse_date_from_filename(filename):
    """
    ดึงวันที่จากชื่อไฟล์ รองรับหลายรูปแบบ:
      DD-MM-YY, DD-MM-YYYY, (DD-MM-YY), DD MM YY, etc.
    YY < 100 ถือว่าเป็น Thai BE สั้น: YY + 2500 แล้วลบ 543 → CE
    """
    fname = os.path.basename(filename)
    # ลบ extension
    fname = re.sub(r"\.\w+$", "", fname)

    patterns = [
        r"(\d{1,2})[\s\-_/](\d{1,2})[\s\-_/](\d{2,4})",  # D-M-Y
        r"(\d{4})[\s\-_/](\d{1,2})[\s\-_/](\d{1,2})",     # Y-M-D
    ]
    for pat in patterns:
        m = re.search(pat, fname)
        if not m:
            continue
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        # ตรวจว่าเป็น YYYY-MM-DD หรือ DD-MM-YY
        if a > 31:  # YYYY-MM-DD
            y, mo, d = a, b, c
        else:       # DD-MM-YY / DD-MM-YYYY
            d, mo, y = a, b, c
        # แปลง Thai BE สั้น → CE
        if y < 100:
            y = (2500 + y) - 543  # e.g. 69 → 2569 → 2026
        elif y > 2400:
            y = y - 543           # Thai BE full → CE
        if 1 <= d <= 31 and 1 <= mo <= 12:
            return f"{y:04d}-{mo:02d}-{d:02d}"

    from datetime import datetime
    return datetime.today().strftime("%Y-%m-%d")


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 wmc_excel_processor.py <excel_file.xlsx> [YYYY-MM-DD]")
        sys.exit(1)

    excel_path = sys.argv[1]
    date_str   = sys.argv[2] if len(sys.argv) > 2 else parse_date_from_filename(excel_path)

    print(f"📅 วันที่: {date_str}")
    print(f"📂 Excel:  {excel_path}")

    # ─── 1. อ่าน Excel ───────────────────────────────────────────────────────
    print("\n1️⃣  อ่าน Excel...")
    try:
        import openpyxl
    except ImportError:
        install_openpyxl()
        import openpyxl

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    day_data = extract_data_from_workbook(wb, date_str)
    wb.close()
    print(f"   OPD Revenue: {day_data['opdTotal']:,.0f} บาท")
    print(f"   IPD Revenue: {day_data['ipdRev']:,.0f} บาท")
    print(f"   OT Revenue:  {day_data['otRev']:,.0f} บาท")
    print(f"   OPD Visit:   {day_data['opdVisit']:,} ราย")
    print(f"   IPD Admit:   {day_data['ipdAdmit']:,} ราย")
    print(f"   BU rows:     {len(day_data['bu'])}")

    # ─── 2. อัปเดต Dashboard HTML ────────────────────────────────────────────
    print("\n2️⃣  อัปเดต Dashboard HTML...")
    if not update_dashboard_html(date_str, day_data):
        sys.exit(1)

    # ─── 3. Deploy ไป GitHub Pages ──────────────────────────────────────────────
    print("\n3️⃣  Deploy ไป GitHub Pages...")
    site_url = deploy_to_github_pages()

    # ─── 4. LINE (ปิดไว้ — ใช้ Pin URL ใน LINE Group แทน เพื่อประหยัด quota) ───
    # send_line_messages(date_str, site_url)
    print("\n4️⃣  LINE: ใช้ Pin URL แทนการส่งอัตโนมัติ (ไม่ใช้ quota)")

    print(f"\n🎉 Pipeline เสร็จสมบูรณ์!")
    print(f"   Dashboard: {site_url}")


if __name__ == "__main__":
    main()
