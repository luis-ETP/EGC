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

    return overall_summary, inventory, fifo_rows, meta

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

# ── Inventory (from FIFO sheet) ────────────────────────────────────────────────
def _extract_inventory(wb):
    ws = wb["FIFO"]
    rows = list(ws.iter_rows(values_only=True))

    # Header is row 1 (no title row anymore)
    headers = [str(v).strip() if v else f"col{j}" for j, v in enumerate(rows[0])]

    # Build hierarchy: bulk_plant -> product -> batch -> invoice -> [bols]
    hierarchy = {}

    for row in rows[1:]:
        if not row[0]: break
        entry = {h: row[j] for j, h in enumerate(headers)}
        if entry.get("Type") != "RTB": continue

        bp      = str(entry.get("Bulk Plant", "") or "")
        prod    = str(entry.get("Product", "") or "")
        batch   = str(entry.get("Batch", "") or "")
        bol     = str(entry.get("BOL", "") or "")
        liters  = float(entry.get("Liters", 0) or 0)
        rem_l   = entry.get("Remaining L (BOL)")
        rem_l   = float(rem_l) if rem_l is not None else liters
        cost    = float(entry.get("Cost / L (MXN)", 0) or 0)

        # supplier invoice from BOL source col (we need it — get from Supplier Invoice col)
        inv_str = str(entry.get("Source BOLs", "") or "")

        # Navigate hierarchy
        if bp not in hierarchy:
            hierarchy[bp] = {}
        if prod not in hierarchy[bp]:
            hierarchy[bp][prod] = {}
        if batch not in hierarchy[bp][prod]:
            hierarchy[bp][prod][batch] = {}

        # We need invoice — it's not in FIFO sheet directly, use batch as key for now
        inv_key = batch  # will group by batch since invoice isn't in FIFO sheet
        if inv_key not in hierarchy[bp][prod][batch]:
            hierarchy[bp][prod][batch][inv_key] = []

        hierarchy[bp][prod][batch][inv_key].append({
            "bol":         bol,
            "liters":      liters,
            "remaining_l": rem_l,
            "cost_per_l":  cost,
        })

    # Now get invoice info from Purchase to BOL-RTB
    try:
        ws_bol = wb["Purchase to BOL-RTB"]
        bol_to_inv = {}
        for row in ws_bol.iter_rows(min_row=8, values_only=True):
                if not row[2]: break  # col C (Supplier) = reliable non-formula col
                bol_val = str(row[5]).strip() if row[5] else ""
                inv_val = str(row[3]).strip() if row[3] else ""
                if bol_val:
                    bol_to_inv[bol_val] = inv_val
    except:
        bol_to_inv = {}

    # Rebuild with invoice level
    result = {}
    for bp, prods in hierarchy.items():
        result[bp] = {}
        for prod, batches in prods.items():
            result[bp][prod] = {}
            for batch, inv_groups in batches.items():
                # Re-group by actual invoice
                inv_map = defaultdict(list)
                for _, bols in inv_groups.items():
                    for b in bols:
                        inv = bol_to_inv.get(b["bol"], "Unknown")
                        inv_map[inv].append(b)

                result[bp][prod][batch] = {}
                for inv, bols in inv_map.items():
                    result[bp][prod][batch][inv] = bols

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
