ETP FIFO Engine — Web App
Upload the Investor Summary Excel file, the engine processes it, and you download the result.
What it does
Stage 1 — Supplier Invoice → BOL

Allocates supplier invoice gallons to each BOL in Load Tracking sheet order (FIFO per supplier)
Writes to Purchase to BOL-RTB: Supplier Invoice (col D), Batch (col O), Cost/Gal USD (col J), Total/Liter MXN (col R)
Writes to Supplier Invoices: Net RTB Gallons (col W), Remainder Gallons (col X), liter formulas (col U, V)

Stage 2 — RTB → BTC FIFO

Builds inventory queue per (product, bulk plant) from RTB Total Cost/L (col AV)
BTCs draw from queue in sheet order, weighted average when spanning multiple RTBs
Writes to Load Tracking: Supply Cost (col AR), Batch (col BE), Supplier Invoice (col BF), BOL Source (col BG), Batch Source (col BH)

FIFO sheet — Full chronological view with running inventory balance and summary block
Overall Summary — Remaining gallons and weighted average cost in inventory per supplier
Run locally
bashpip install -r requirements.txt
python app.py
# Open http://localhost:5000
Deploy to Render

Push this folder to a GitHub repo
Go to https://render.com → New → Web Service
Connect your repo
Settings:

Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
Plan: Free


Deploy → get a public URL to share with your team
