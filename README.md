# ETPMEX — FIFO Engine & Dashboard

Full-stack web app for FIFO cost allocation and inventory tracking for fuel distribution operations.

---

## What it does

### FIFO Engine
Processes the Investor Summary Excel workbook in two stages:

**Stage 1 — Supplier Invoice → BOL**
- Allocates supplier invoice gallons to each BOL in Load Tracking sheet order (FIFO per supplier)
- Writes to Purchase to BOL-RTB: Supplier Invoice (D), Batch (E), Cost/Gal USD (I), Cost/Gal + Adder (K)
- Writes to Supplier Invoices: Net RTB Gallons (W), Remainder Gallons (X), liter formulas (U, V)

**Stage 2 — RTB → BTC FIFO**
- Builds inventory queue per (product, bulk plant) from RTB Total Cost/L
- BTCs draw from queue in sheet order; weighted average when spanning multiple RTBs
- Writes to Load Tracking: Supply Cost (AR), Batch (BE), Supplier Invoice (BF), BOL Source (BG), Batch Source (BH)
- Creates a FIFO sheet with full chronological view

---

## Dashboard

### Roles
| User | Password | Access |
|------|----------|--------|
| `admin` | `ETP@admin2024` | All tabs: Overview, Inventory, FIFO Log, Investor Summary |
| `investor` | `ETP@inv2024` | Investor Summary only |

> Change credentials in app.py under the USERS dict before going to production.

### Tabs
- **Overview** — KPI cards, 3 charts, supplier summary table, file upload (admin only)
- **Inventory** — Bulk Plant → Product → Batch → Invoice → BOL tank visualization
- **FIFO Log** — Full chronological RTB/BTC allocation table
- **Investor Summary** — Investor-facing view with KPIs, supplier breakdown, and charts

---

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask app — routes, auth, API endpoints |
| `fifo_engine.py` | Core FIFO allocation engine |
| `extract.py` | Extracts structured data from processed workbook |
| `db.py` | Supabase/PostgreSQL read/write |
| `dashboard.html` | Full dashboard UI |
| `requirements.txt` | Python dependencies |

---

## Deploy to Render

1. Push all 6 files to the root of your GitHub repo
2. Create Web Service on Render, connect your repo
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
3. Add Environment Variables (Render → Settings → Environment):

```
DATABASE_URL = postgresql://postgres.mgozbmyvrwcsnnwgqpvy:[password]@aws-0-us-west-2.pooler.supabase.com:6543/postgres
SECRET_KEY   = [choose a strong random string]
```

The Supabase table is created automatically on first startup — no manual DB setup needed.

---

## Run locally

```bash
pip install -r requirements.txt
export DATABASE_URL="postgresql://..."
export SECRET_KEY="any-local-secret"
python app.py
# Open http://localhost:5000
```

---

## Security notes
- Change default credentials before sharing the URL
- Rotate your Supabase password after initial setup
- Set a strong SECRET_KEY in production
- Only processed numerical data is stored in Supabase — no raw Excel files persisted
