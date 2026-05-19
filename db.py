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
                    bol_tab JSONB,
                    overview_exp JSONB
                );
            """)
            # Add investment column if missing (idempotent)
            cur.execute("""
                ALTER TABLE dashboard_data ADD COLUMN IF NOT EXISTS investment JSONB;
            """)
            cur.execute("""
                ALTER TABLE dashboard_data ADD COLUMN IF NOT EXISTS bol_tab JSONB;
            """)
            cur.execute("""
                ALTER TABLE dashboard_data ADD COLUMN IF NOT EXISTS overview_exp JSONB;
            """)
        conn.commit()

def save_data(filename, overall_summary, inventory, fifo_rows, meta, investment=None, bol_tab=None, overview_exp=None):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM dashboard_data")
            cur.execute("""
                INSERT INTO dashboard_data
                    (filename, overall_summary, inventory, fifo_rows, meta, investment, bol_tab, overview_exp)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                filename,
                json.dumps(overall_summary),
                json.dumps(inventory),
                json.dumps(fifo_rows),
                json.dumps(meta),
                json.dumps(investment or {}),
                json.dumps(bol_tab or {}),
                json.dumps(overview_exp or {}),
            ))
        conn.commit()

def load_data():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM dashboard_data ORDER BY uploaded_at DESC LIMIT 1")
            row = cur.fetchone()
            if not row:
                return None
            def _get(key, default=None):
                try: return row[key] if row[key] is not None else default
                except: return default
            return {
                "uploaded_at":     row["uploaded_at"].isoformat() if row["uploaded_at"] else None,
                "filename":        row["filename"],
                "overall_summary": _get("overall_summary", []),
                "inventory":       _get("inventory", {}),
                "fifo_rows":       _get("fifo_rows", []),
                "meta":            _get("meta", {}),
                "investment":      _get("investment", {}),
                "bol_tab":         _get("bol_tab", {}),
                "overview_exp":    _get("overview_exp", {}),
            }

# ── Settings ───────────────────────────────────────────────────────────────────
def init_settings():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
            """)
            # Default markup
            cur.execute("""
                INSERT INTO settings (key, value) VALUES ('markup_usd', '0.02')
                ON CONFLICT (key) DO NOTHING;
            """)
        conn.commit()

def get_setting(key, default=None):
    try:
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT value FROM settings WHERE key = %s", (key,))
                row = cur.fetchone()
                return row['value'] if row else default
    except:
        return default

def set_setting(key, value):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO settings (key, value) VALUES (%s, %s)
                ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
            """, (key, str(value)))
        conn.commit()
