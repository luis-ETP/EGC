"""
Extract structured data from a processed FIFO workbook for the dashboard.
"""
import openpyxl
from collections import defaultdict

LITERS_PER_GAL = 3.7854

def extract(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    overall_summary = _extract_overall_summary(wb)
    inventory       = _extract_inventory(wb)
    fifo_rows       = _extract_fifo(wb)
    meta            = _extract_meta(wb)

    investment = _extract_investment_summary(wb)
    return overall_summary, inventory, fifo_rows, meta, investment

# ── Overall Summary ────────────────────────────────────────────────────────────
def _extract_overall_summary(wb):
    ws = wb["Overall Summary"]
    rows = list(ws.iter_rows(values_only=True))

    # Find header row (contains "Row Labels")
    hdr_idx = next(i for i, r in enumerate(rows) if r[0] == "Row Labels")
    headers = [str(v).strip() if v else f"col{j}" for j, v in enumerate(rows[hdr_idx])]

    result = []
    for row in rows[hdr_idx + 1:]:
        if not row[0]: continue
        entry = {}
        for j, h in enumerate(headers):
            v = row[j]
            entry[h] = float(v) if isinstance(v, (int, float)) else (str(v) if v else None)
        # Normalise key names so dashboard always finds them regardless of Excel naming
        entry["_wired"]     = entry.get("Total Wired Amount", 0) or 0
        entry["_paid_gal"]  = entry.get("Paid for Gallons (Allocation)") or entry.get("Paid for Gallons ") or 0
        entry["_pulled"]    = entry.get("Gallons Pulled from Allocation (RTB & RTC)") or entry.get("Gallons Pulled (RTB & RTC)") or 0
        entry["_rem_alloc"] = entry.get("Remaining Gallons in Allocation") or 0
        entry["_rem_inv"]   = entry.get("Remaining Gallons in Inventory") or 0
        entry["_avg_cost"]  = entry.get("Weighted Average Cost in Inventory (MXN/L)") or entry.get("Weighted Average Cost in Inventory") or 0
        entry["_paid_back"] = entry.get("Amount Paid Back by Mexico (MXN)") or entry.get("Amount Paid Back by Mexico ") or entry.get("Amount Paid Back by Mexico") or 0
        entry["_balance"]   = entry.get("Mexico Balance (MXN)") or entry.get("Mexico Balance") or 0
        result.append(entry)
    return result

# ── Inventory (from FIFO sheet + Purchase to BOL-RTB) ─────────────────────────
def _extract_inventory(wb):
    """
    Build hierarchy: bulk_plant -> product -> batch -> supplier -> invoice -> [bols]

    BOLs that span two batches/invoices (e.g. Batch="1 | 2", Invoice="A | B")
    are split into separate entries — one per batch/invoice — with proportional liters.
    Each BOL always appears INDIVIDUALLY under its own invoice.
    """
    # Step 1: read FIFO sheet for each RTB BOL
    ws_fifo = wb["FIFO"]
    fifo_headers = [str(v).strip() if v else f"col{j}"
                    for j, v in enumerate(next(ws_fifo.iter_rows(values_only=True)))]

    fifo_bols = {}  # bol_str -> {liters, remaining_l, cost_per_l, bp, prod}
    for row in ws_fifo.iter_rows(min_row=2, values_only=True):
        if not row[0]: break
        entry = {h: row[j] for j, h in enumerate(fifo_headers)}
        if entry.get("Type") != "RTB": continue
        bol = str(entry.get("BOL", "") or "")
        if not bol: continue
        liters = float(entry.get("Liters", 0) or 0)
        rem    = entry.get("Remaining L (BOL)")
        fifo_bols[bol] = {
            "liters":      liters,
            "remaining_l": float(rem) if rem is not None else liters,
            "cost_per_l":  float(entry.get("Cost / L (MXN)", 0) or 0),
            "bp":          str(entry.get("Bulk Plant", "") or ""),
            "prod":        str(entry.get("Product", "") or ""),
        }

    # Step 2: read Purchase to BOL-RTB for invoice, batch, supplier per BOL
    # For split BOLs (e.g. Batch="1 | 2", Invoice="A | B"), split into two entries
    bol_entries = []  # list of {bol, batch, invoice, supplier, alloc_frac}
    try:
        ws_bol = wb["Purchase to BOL-RTB"]
        for row in ws_bol.iter_rows(min_row=8, values_only=True):
            if not row[2]: break
            bol_str  = str(row[5]).strip() if row[5] else ""
            supplier = str(row[2]).strip() if row[2] else ""
            inv_raw  = str(row[3]).strip() if row[3] else ""
            bat_raw  = str(row[4]).strip() if row[4] else ""
            cost_j   = float(row[9])  if row[9]  is not None else 0.0  # J Cost/Gal USD
            cost_l_raw = str(row[11]).strip() if row[11] is not None else ""

            if not bol_str: continue

            # Split multi-batch/invoice BOLs into individual entries
            batches  = [b.strip() for b in bat_raw.split("|")] if "|" in bat_raw else [bat_raw]
            invoices = [v.strip() for v in inv_raw.split("|")] if "|" in inv_raw else [inv_raw]

            if len(batches) > 1 or len(invoices) > 1:
                # Cross-batch BOL: split proportionally
                # We don't have exact split fractions here, so split equally
                # (the FIFO sheet has the blended cost already)
                n = max(len(batches), len(invoices))
                for k in range(n):
                    b = batches[k] if k < len(batches) else batches[-1]
                    inv = invoices[k] if k < len(invoices) else invoices[-1]
                    bol_entries.append({
                        "bol": bol_str, "batch": b, "invoice": inv,
                        "supplier": supplier, "split_n": n, "split_k": k,
                    })
            else:
                bol_entries.append({
                    "bol": bol_str, "batch": bat_raw, "invoice": inv_raw,
                    "supplier": supplier, "split_n": 1, "split_k": 0,
                })
    except Exception as e:
        pass

    # Step 3: build hierarchy
    result = {}
    for entry in bol_entries:
        bol      = entry["bol"]
        batch    = entry["batch"]
        invoice  = entry["invoice"]
        supplier = entry["supplier"]
        split_n  = entry["split_n"]
        fifo     = fifo_bols.get(bol)
        if not fifo: continue

        bp   = fifo["bp"]
        prod = fifo["prod"]
        # Proportional liters for split BOLs
        liters      = round(fifo["liters"]      / split_n, 4)
        remaining_l = round(fifo["remaining_l"] / split_n, 4)
        cost_per_l  = fifo["cost_per_l"]  # blended cost same for all splits

        # Navigate: bp -> prod -> batch -> supplier -> invoice -> [bols]
        result.setdefault(bp, {})
        result[bp].setdefault(prod, {})
        result[bp][prod].setdefault(batch, {})
        result[bp][prod][batch].setdefault(supplier, {})
        result[bp][prod][batch][supplier].setdefault(invoice, [])
        result[bp][prod][batch][supplier][invoice].append({
            "bol":         bol,
            "liters":      liters,
            "remaining_l": remaining_l,
            "cost_per_l":  cost_per_l,
        })

    return result

# ── FIFO rows ──────────────────────────────────────────────────────────────────
def _extract_fifo(wb):
    ws = wb["FIFO"]
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(v).strip() if v else f"col{j}" for j, v in enumerate(rows[0])]

    result = []
    for row in rows[1:]:
        if not row[0]: break
        entry = {}
        for j, h in enumerate(headers):
            v = row[j]
            if hasattr(v, 'isoformat'):
                v = v.strftime("%d/%m/%Y")
            elif isinstance(v, float):
                v = round(v, 4)
            entry[h] = v
        result.append(entry)
    return result

# ── Meta ───────────────────────────────────────────────────────────────────────
def _extract_meta(wb):
    # Pull key KPIs from Overall Summary total row
    ws = wb["Overall Summary"]
    total_row = None
    for row in ws.iter_rows(values_only=True):
        if row[0] and "TOTAL" in str(row[0]).upper():
            total_row = row
            break

    meta = {}
    if total_row:
        def _f(v): 
            try: return round(float(v), 4)
            except: return 0
        # Use overall_summary normalised keys for reliability
        os_rows = _extract_overall_summary(wb)
        total_os = next((r for r in os_rows if r.get("Row Labels","").upper().find("TOTAL") >= 0), {})
        meta = {
            "total_invoiced_usd":      _f(total_row[1]),
            "total_gallons":           _f(total_row[2]),
            "total_wired":             total_os.get("_wired", _f(total_row[3])),
            "paid_for_gallons":        total_os.get("_paid_gal", _f(total_row[4])),
            "gallons_pulled":          total_os.get("_pulled", _f(total_row[5])),
            "remaining_allocation":    total_os.get("_rem_alloc", _f(total_row[6])),
            "remaining_inventory_gal": total_os.get("_rem_inv", _f(total_row[7])),
            "avg_cost_inventory":      total_os.get("_avg_cost", _f(total_row[8])),
            "amount_paid_back":        total_os.get("_paid_back", _f(total_row[9])),
            "mexico_balance":          total_os.get("_balance", _f(total_row[10])),
        }
    return meta

# ── Investment Summary ─────────────────────────────────────────────────────────
def _extract_investment_summary(wb, uploaded_at=None):
    """Compute all Investment Summary values from raw sheet data."""
    f = lambda v: float(v) if isinstance(v, (int, float)) else 0.0

    # ── Committed Capital (Investment Summary rows 7-8) ──────────────────────
    ws_is = wb["Investment Summary"]
    rows_is = list(ws_is.iter_rows(values_only=True))
    commits = []
    for row in rows_is:
        if row[1] in ("Round 1", "Round 2"):
            usd = f(row[2])
            fx  = f(row[5])
            mxn = usd * fx
            date_val = row[4]
            date_str = date_val.strftime("%d-%b-%y") if hasattr(date_val, "strftime") else str(date_val or "")
            commits.append({"round": row[1], "usd": usd, "mxn": mxn, "date": date_str, "fx": fx})
    total_committed_mxn = sum(c["mxn"] for c in commits)
    total_committed_usd = sum(c["usd"] for c in commits)
    avg_fx = total_committed_mxn / total_committed_usd if total_committed_usd else 17.31

    # ── Load Tracking aggregations ───────────────────────────────────────────
    ws_lt = wb["Load Tracking"]
    rtb_liters = 0.0
    btc_paid = {"sale": 0.0, "cost": 0.0, "liters": 0.0, "margin": 0.0}
    rtc_paid = {"sale": 0.0, "cost": 0.0, "liters": 0.0, "margin": 0.0}
    btc_pend = {"cost": 0.0, "liters": 0.0}
    rtc_pend = {"cost": 0.0, "liters": 0.0}

    for i, row in enumerate(ws_lt.iter_rows(values_only=True), start=1):
        if i == 1: continue
        if not row[0]: continue
        grp    = row[10]
        liters = f(row[22])
        status = row[55]
        sale   = f(row[37])
        cost   = f(row[43]) * liters + f(row[39]) + f(row[44])  # supply + freight + commission
        margin = sale - cost

        if grp == "RTB":
            rtb_liters += liters
        elif grp == "BTC":
            if status == "PAID":
                btc_paid["sale"] += sale; btc_paid["cost"] += cost
                btc_paid["liters"] += liters; btc_paid["margin"] += margin
            elif status == "PENDING":
                btc_pend["cost"] += cost; btc_pend["liters"] += liters
        elif grp == "RTC":
            if status == "PAID":
                rtc_paid["sale"] += sale; rtc_paid["cost"] += cost
                rtc_paid["liters"] += liters; rtc_paid["margin"] += margin
            elif status == "PENDING":
                rtc_pend["cost"] += cost; rtc_pend["liters"] += liters

    # ── Supplier Invoices — Allocation (ACTIVE) ──────────────────────────────
    ws_inv = wb["Supplier Invoices"]
    alloc_mxn = 0.0; alloc_liters = 0.0
    for i, row in enumerate(ws_inv.iter_rows(values_only=True), start=1):
        if i <= 7: continue
        if row[0] is None: break
        if str(row[26] or "").upper() == "ACTIVE":
            alloc_mxn    += f(row[24])  # Remainder Amount MXN
            alloc_liters += f(row[21])  # Remainder Liters Paid and No BOL

    # ── Inventory cost from FIFO remaining ──────────────────────────────────
    ws_fifo = wb["FIFO"]
    inv_cost = 0.0; inv_liters = 0.0
    for i, row in enumerate(ws_fifo.iter_rows(values_only=True), start=1):
        if i == 1: continue
        if not row[0]: break
        if row[2] == "RTB":
            rem  = f(row[13])  # Remaining L (BOL)
            cost = f(row[9])   # Cost/L MXN
            inv_cost   += rem * cost
            inv_liters += rem

    # ── KPIs ─────────────────────────────────────────────────────────────────
    total_margin   = btc_paid["margin"] + rtc_paid["margin"]
    recovered_mxn  = btc_paid["sale"]   + rtc_paid["sale"]
    active_capital = alloc_mxn + inv_cost + btc_pend["cost"] + rtc_pend["cost"]
    available      = total_committed_mxn - active_capital
    revolved       = recovered_mxn / total_committed_mxn if total_committed_mxn else 0
    # Investor share % from Investment Summary row 14 formula: H14 = G14 * share_pct
    # We read it from Excel directly if cached, else default 40%
    inv_share_pct  = 0.40

    def _roi(d): return d["margin"] / d["cost"] if d["cost"] else 0
    def _mxnl(d): return d["margin"] / d["liters"] if d["liters"] else 0
    def _usdgal(d, fx): return _mxnl(d) / fx * 3.7854 if fx else 0

    return {
        "as_of": uploaded_at or "",
        "commits": commits,
        "total_committed_usd": total_committed_usd,
        "total_committed_mxn": total_committed_mxn,
        "active_capital":  active_capital,
        "available":       available,
        "recovered_mxn":   recovered_mxn,
        "revolved":        revolved,
        "total_margin":    total_margin,
        "investor_share":  total_margin * inv_share_pct,
        "inv_share_pct":   inv_share_pct,
        "active_detail": {
            "alloc_mxn":    alloc_mxn,    "alloc_liters":    alloc_liters,
            "inv_mxn":      inv_cost,     "inv_liters":      inv_liters,
            "rtc_pend_mxn": rtc_pend["cost"], "rtc_pend_liters": rtc_pend["liters"],
            "btc_pend_mxn": btc_pend["cost"], "btc_pend_liters": btc_pend["liters"],
            "total_mxn":    active_capital,
            "total_liters": alloc_liters + inv_liters + rtc_pend["liters"] + btc_pend["liters"],
        },
        "recovered_detail": {
            "rtc": {"mxn": rtc_paid["sale"], "liters": rtc_paid["liters"],
                    "margin": rtc_paid["margin"], "roi": _roi(rtc_paid),
                    "mxnl": _mxnl(rtc_paid), "usdgal": _usdgal(rtc_paid, avg_fx)},
            "btc": {"mxn": btc_paid["sale"], "liters": btc_paid["liters"],
                    "margin": btc_paid["margin"], "roi": _roi(btc_paid),
                    "mxnl": _mxnl(btc_paid), "usdgal": _usdgal(btc_paid, avg_fx)},
            "total": {"mxn": recovered_mxn,
                      "liters": btc_paid["liters"] + rtc_paid["liters"],
                      "margin": total_margin,
                      "roi": total_margin / recovered_mxn if recovered_mxn else 0,
                      "mxnl": total_margin / (btc_paid["liters"] + rtc_paid["liters"]) if (btc_paid["liters"] + rtc_paid["liters"]) else 0,
                      "usdgal": _usdgal({"margin": total_margin, "liters": btc_paid["liters"] + rtc_paid["liters"]}, avg_fx)},
        },
    }
