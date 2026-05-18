import os, json, psycopg2
from psycopg2.extras import RealDictCursor

DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://postgres.mgozbmyvrwcsnnwgqpvy:[password]@aws-0-us-west-2.pooler.supabase.com:5432/postgres"
)

def get_conn():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

def init_db():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS dashboard_data (
                    id SERIAL PRIMARY KEY,
                    uploaded_at TIMESTAMPTZ DEFAULT NOW(),
                    filename TEXT,
                    overall_summary JSONB,
                    inventory JSONB,
                    fifo_rows JSONB,
                    meta JSONB,
                    investment JSONB,
                    bol_tab JSONB
                );
            """)
            # Add investment column if missing (idempotent)
            cur.execute("""
                ALTER TABLE dashboard_data ADD COLUMN IF NOT EXISTS investment JSONB;
            """)
            cur.execute("""
                ALTER TABLE dashboard_data ADD COLUMN IF NOT EXISTS bol_tab JSONB;
            """)
        conn.commit()

def save_data(filename, overall_summary, inventory, fifo_rows, meta, investment=None, bol_tab=None):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM dashboard_data")
            cur.execute("""
                INSERT INTO dashboard_data
                    (filename, overall_summary, inventory, fifo_rows, meta, investment, bol_tab)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (
                filename,
                json.dumps(overall_summary),
                json.dumps(inventory),
                json.dumps(fifo_rows),
                json.dumps(meta),
                json.dumps(investment or {}),
                json.dumps(bol_tab or {}),
            ))
        conn.commit()

def load_data():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM dashboard_data ORDER BY uploaded_at DESC LIMIT 1")
            row = cur.fetchone()
            if not row:
                return None
            return {
                "uploaded_at":     row["uploaded_at"].isoformat() if row["uploaded_at"] else None,
                "filename":        row["filename"],
                "overall_summary": row["overall_summary"],
                "inventory":       row["inventory"],
                "fifo_rows":       row["fifo_rows"],
                "meta":            row["meta"],
                "investment":      row.get("investment") or {},
                "bol_tab":         row.get("bol_tab") or {},
            }
