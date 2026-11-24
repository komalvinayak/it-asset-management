import streamlit as st
import pandas as pd
import psycopg2
import os
import io
from datetime import date

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Akal University – IT Asset Management",
    page_icon="akal_logo.png" if os.path.exists("akal_logo.png") else "🖥️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Theme: Akal University gold + navy ────────────────────────────────────────
GOLD  = "#C8830A"
NAVY  = "#1B3A6B"
CREAM = "#FFF8EE"

st.markdown(f"""
<style>
/* ---- Base ---- */
.main, [data-testid="stAppViewContainer"] {{
    background-color: #ffffff;
}}
/* ---- Sidebar ---- */
[data-testid="stSidebar"] {{
    background-color: {CREAM};
    border-right: 2px solid {GOLD};
}}
[data-testid="stSidebar"] .stCaption {{ color: {NAVY}; font-weight:600; }}

/* ---- Metric cards ---- */
div[data-testid="metric-container"] {{
    background: #ffffff;
    border: 1px solid #e8d8b8;
    border-top: 4px solid {GOLD};
    border-radius: 10px;
    padding: 16px 20px;
    box-shadow: 0 2px 6px rgba(200,131,10,0.10);
}}
div[data-testid="metric-container"] label {{
    color: {NAVY}; font-size: 0.82rem; font-weight: 600; letter-spacing:0.03em;
}}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {{
    color: {GOLD}; font-size: 1.9rem; font-weight: 800;
}}

/* ---- Sidebar nav buttons ---- */
[data-testid="stSidebar"] .stButton > button {{
    border-radius: 8px;
    font-weight: 500;
    border: 1px solid #e0c88a;
    background: #ffffff;
    color: {NAVY};
    text-align: left;
    transition: all 0.15s;
}}
[data-testid="stSidebar"] .stButton > button:hover {{
    background: {GOLD};
    color: #ffffff;
    border-color: {GOLD};
}}
/* Active nav button (primary type = current page) */
[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
    background: {NAVY};
    color: #ffffff;
    border-color: {NAVY};
    font-weight: 700;
}}
[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {{
    background: {GOLD};
    border-color: {GOLD};
}}

/* ---- Main area buttons ---- */
.main .stButton > button {{
    border-radius: 7px;
    font-weight: 500;
    border: 1px solid #c0ad8a;
    background: #ffffff;
    color: {NAVY};
    transition: all 0.15s;
}}
.main .stButton > button:hover {{
    background: {GOLD};
    color: #ffffff;
    border-color: {GOLD};
}}
.main .stButton > button[kind="primary"] {{
    background: {GOLD}; color: #fff; border-color: {GOLD};
}}
.main .stButton > button[kind="primary"]:hover {{
    background: #a86a06;
}}

/* ---- Page headings ---- */
h1 {{ color: {NAVY}; border-bottom: 3px solid {GOLD}; padding-bottom: 6px; }}
h2, h3 {{ color: {NAVY}; }}

/* ---- Tab bar ---- */
[data-testid="stTabs"] [data-baseweb="tab-list"] {{
    border-bottom: 2px solid {GOLD};
}}
[data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {{
    color: {GOLD}; border-bottom: 3px solid {GOLD}; font-weight: 700;
}}

/* ---- Dataframe header ---- */
[data-testid="stDataFrame"] thead tr th {{
    background-color: {NAVY} !important;
    color: #ffffff !important;
}}

/* ---- Divider ---- */
hr {{ border-color: #e8d8b8; }}
</style>
""", unsafe_allow_html=True)

# ── DB helpers ─────────────────────────────────────────────────────────────────
@st.cache_resource
def _get_conn():
    """Single persistent connection with autocommit — avoids transaction state issues."""
    conn = psycopg2.connect(st.secrets["DATABASE_URL"])
    conn.autocommit = True
    return conn


def _conn():
    conn = _get_conn()
    try:
        conn.cursor().execute("SELECT 1")
    except Exception:
        st.cache_resource.clear()
        conn = _get_conn()
    return conn


def db_query(sql, params=None) -> pd.DataFrame:
    conn = _conn()
    try:
        return pd.read_sql_query(sql, conn, params=params)
    except Exception:
        st.cache_resource.clear()
        raise


def db_exec(sql, params=None):
    conn = _conn()
    with conn.cursor() as cur:
        cur.execute(sql, params)


import tempfile
BACKUP_DIR = st.secrets.get("BACKUP_DIR", os.path.join(tempfile.gettempdir(), "it_asset_backups"))
os.makedirs(BACKUP_DIR, exist_ok=True)

def take_backup(trigger_action: str, trigger_summary: str) -> str:
    """Create both a plain-SQL dump and an Excel snapshot of the full database."""
    import subprocess
    from datetime import datetime
    from openpyxl.styles import PatternFill, Font, Alignment
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = trigger_action.upper().replace(" ", "_")[:20]

    sql_fname  = f"backup_{ts}_{safe}.sql"
    xlsx_fname = f"backup_{ts}_{safe}.xlsx"
    sql_fpath  = os.path.join(BACKUP_DIR, sql_fname)
    xlsx_fpath = os.path.join(BACKUP_DIR, xlsx_fname)

    PG_DUMP = st.secrets.get("PG_DUMP_PATH", "")

    # ── 1. SQL dump (local only — skipped on cloud) ───────────────────────────
    sql_ok = False
    if PG_DUMP and os.path.exists(PG_DUMP):
        try:
            db_url = st.secrets["DATABASE_URL"]
            import urllib.parse
            parsed = urllib.parse.urlparse(db_url)
            env = os.environ.copy()
            env["PGPASSWORD"] = parsed.password or ""
            result = subprocess.run(
                [PG_DUMP, "-h", parsed.hostname, "-p", str(parsed.port or 5432),
                 "-U", parsed.username, "-d", parsed.path.lstrip("/"),
                 "-F", "p", "-f", sql_fpath],
                env=env, capture_output=True, text=True, timeout=120,
            )
            sql_ok = result.returncode == 0
        except Exception:
            pass

    # ── 2. Excel snapshot (all key tables) ───────────────────────────────────
    xlsx_ok = False
    sheets = {
        "Assets": """
            SELECT a.asset_id, t.type_name, a.serial_no, m.make_name, a.model_name,
                   a.purpose, a.status, a.asset_category_code,
                   COALESCE(d.dept_name,'Unallocated') AS department,
                   COALESCE(l.location_name,'') AS location
            FROM assets a
            JOIN asset_types t ON a.type_id=t.type_id
            LEFT JOIN makes m ON a.make_id=m.make_id
            LEFT JOIN allocations al ON a.asset_id=al.asset_id AND al.is_current=TRUE
            LEFT JOIN departments d ON al.dept_id=d.dept_id
            LEFT JOIN locations l ON al.location_id=l.location_id
            ORDER BY a.asset_id""",
        "Procurement":    "SELECT * FROM procurement ORDER BY procurement_id",
        "Allocations":    "SELECT * FROM allocations ORDER BY allocation_id",
        "Computer_Specs": "SELECT * FROM computer_specs ORDER BY asset_id",
        "Peripherals":    "SELECT * FROM peripherals ORDER BY peripheral_id",
        "Network_Config": "SELECT * FROM network_config ORDER BY asset_id",
        "Intercom":       "SELECT * FROM intercom_directory ORDER BY directory_id",
        "Departments":    "SELECT * FROM departments ORDER BY dept_id",
        "Locations":      "SELECT * FROM locations ORDER BY location_id",
        "Asset_Types":    "SELECT * FROM asset_types ORDER BY type_id",
        "Makes":          "SELECT * FROM makes ORDER BY make_id",
    }
    try:
        gold_fill = PatternFill("solid", fgColor="C8830A")
        wfont     = Font(bold=True, color="FFFFFF")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for sheet, sql in sheets.items():
                try:
                    df = db_query(sql)
                except Exception:
                    df = pd.DataFrame()
                df.to_excel(writer, sheet_name=sheet, index=False)
                ws = writer.sheets[sheet]
                for cell in ws[1]:
                    cell.fill = gold_fill; cell.font = wfont
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = min(
                        max((len(str(c.value or "")) for c in col), default=10) + 4, 50)
        with open(xlsx_fpath, "wb") as f:
            f.write(buf.getvalue())
        xlsx_ok = True
    except Exception:
        pass

    if not sql_ok and not xlsx_ok:
        return ""

    sql_kb  = (os.path.getsize(sql_fpath)  // 1024) if sql_ok  else 0
    xlsx_kb = (os.path.getsize(xlsx_fpath) // 1024) if xlsx_ok else 0
    total_kb = sql_kb + xlsx_kb

    db_exec(
        "INSERT INTO app_backups "
        "(trigger_action, trigger_summary, sql_filename, xlsx_filename, file_size_kb) "
        "VALUES (%s, %s, %s, %s, %s)",
        (trigger_action, trigger_summary,
         sql_fname if sql_ok else None,
         xlsx_fname if xlsx_ok else None,
         total_kb),
    )
    return sql_fpath if sql_ok else xlsx_fpath


def fmt_currency(val):
    v = float(val)
    if v >= 10_000_000:
        return f"₹{v/10_000_000:.2f} Cr"
    elif v >= 100_000:
        return f"₹{v/100_000:.2f} L"
    else:
        return f"₹{v:,.0f}"


# ── Confirmation dialog helper ─────────────────────────────────────────────────
def confirm_dialog(pending_key: str) -> str:
    """
    Render a styled confirmation dialog for a pending action stored in
    st.session_state[pending_key].

    The pending value must be a dict with keys:
        action   : "DELETE" | "SAVE" | "ADD"
        summary  : one-line description shown in bold
        details  : list of "Label: Value" strings shown in the info box
        fn       : callable() that executes the DB operation

    Returns:
        "confirmed"  if the user clicked Yes (fn() was called, state cleared)
        "cancelled"  if the user clicked Cancel (state cleared)
        "waiting"    if the dialog is on screen waiting for user input
        "none"       if there is no pending action
    """
    if pending_key not in st.session_state:
        return "none"

    p = st.session_state[pending_key]
    action  = p["action"]
    summary = p["summary"]
    details = p["details"]
    fn      = p["fn"]

    is_delete = action == "DELETE"
    border    = "#dc3545" if is_delete else GOLD
    bg        = "#fff5f5" if is_delete else "#fffbf0"
    icon      = "🗑️" if is_delete else ("💾" if action == "SAVE" else "➕")
    warn_line = "⚠️ This action <strong>cannot be undone</strong>." if is_delete else \
                "This will <strong>write to the database</strong>."

    rows_html = "".join(
        f"<tr><td style='color:#888;padding:3px 10px 3px 0;white-space:nowrap'>"
        f"{k}</td><td style='font-weight:600;color:#1a1a2e;padding:3px 0'>"
        f"{v}</td></tr>"
        for k, v in details
    )

    st.markdown(f"""
    <div style="border:2px solid {border};border-radius:12px;padding:22px 26px;
                background:{bg};margin:12px 0 6px 0;">
      <p style="font-size:1.1rem;font-weight:700;color:{border};margin:0 0 6px 0;">
        {icon}&nbsp;&nbsp;Confirm {action} — Are you sure?
      </p>
      <p style="color:#444;margin:0 0 12px 0;font-size:0.9rem;">
        You are about to <strong>{action.lower()}</strong> the following record
        {'from' if is_delete else 'in'} the database:
      </p>
      <div style="background:#fff;border:1px solid #ddd;border-radius:8px;
                  padding:12px 16px;margin-bottom:14px;">
        <p style="font-weight:700;color:{border};margin:0 0 8px 0;font-size:1rem;">
          {summary}
        </p>
        <table style="border-collapse:collapse;font-size:0.88rem">{rows_html}</table>
      </div>
      <p style="color:{border};font-weight:600;margin:0;font-size:0.9rem;">{warn_line}</p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns([2, 2, 5])
    with c1:
        if st.button(f"✅  Yes, {action.title()}", key=f"_yes_{pending_key}", type="primary"):
            try:
                take_backup(action, summary)
                fn()
                del st.session_state[pending_key]
                st.session_state[f"_success_{pending_key}"] = True
                st.rerun()
            except Exception as e:
                st.error(f"Database error: {e}")
            return "confirmed"
    with c2:
        if st.button("❌  Cancel", key=f"_no_{pending_key}"):
            del st.session_state[pending_key]
            st.rerun()
            return "cancelled"
    return "waiting"


# ── Login ──────────────────────────────────────────────────────────────────────
def show_login():
    st.markdown("<div style='height:60px'></div>", unsafe_allow_html=True)
    col_l, col_c, col_r = st.columns([1, 1.2, 1])
    with col_c:
        st.markdown(
            f"<h2 style='text-align:center;color:{NAVY};margin-bottom:2px;'>"
            "Akal University</h2>"
            f"<p style='text-align:center;color:{GOLD};font-weight:600;"
            "font-size:0.95rem;margin-top:0 0 20px 0;'>IT Asset Management System</p>",
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter username")
            password = st.text_input("Password", type="password",
                                     placeholder="Password")
            login_btn = st.form_submit_button("Sign In", use_container_width=True)

        if login_btn:
            if username == "admin" and password == "akal@123":
                st.session_state.logged_in = True
                st.session_state.page = "dashboard"
                st.rerun()
            else:
                st.error("Incorrect username or password.")


# ── Sidebar nav ────────────────────────────────────────────────────────────────
NAV = [
    ("Dashboard",          "dashboard"),
    ("Assets",             "assets"),
    ("Procurement",        "procurement"),
    ("Intercom Directory", "intercom"),
    ("Network Devices",    "network"),
    ("Analytics",          "analytics"),
    ("Download Data",      "download"),
]


def show_sidebar():
    with st.sidebar:
        st.markdown(
            f"<div style='padding:18px 8px 4px 8px;'>"
            f"<p style='font-size:1.15rem;font-weight:800;color:{NAVY};"
            f"margin:0;letter-spacing:0.02em;'>Akal University</p>"
            f"<p style='font-size:0.78rem;font-weight:600;color:{GOLD};"
            f"margin:2px 0 0 0;'>IT Asset Management System</p>"
            f"</div>",
            unsafe_allow_html=True,
        )
        st.divider()

        for label, key in NAV:
            active = st.session_state.get("page") == key
            if st.button(label, use_container_width=True, key=f"nav_{key}",
                         type="primary" if active else "secondary"):
                st.session_state.page = key
                st.rerun()

        st.divider()
        if st.button("Sign Out", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.page = "dashboard"
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 – Dashboard
# ══════════════════════════════════════════════════════════════════════════════
def _card(label: str, value: str) -> str:
    return (
        f"<div style='background:#fff;border:1px solid #e8d8b8;border-top:4px solid {GOLD};"
        f"border-radius:10px;padding:18px 12px 14px 12px;"
        f"box-shadow:0 2px 8px rgba(200,131,10,0.10);text-align:center;'>"
        f"<div style='font-size:0.72rem;font-weight:600;color:{NAVY};"
        f"letter-spacing:0.03em;text-transform:uppercase;margin-bottom:6px;"
        f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>{label}</div>"
        f"<div style='font-size:1.6rem;font-weight:800;color:{GOLD};"
        f"white-space:nowrap;'>{value}</div>"
        f"</div>"
    )


def page_dashboard():
    st.title("Dashboard")
    st.caption("Live summary of IT assets across Akal University")
    st.divider()

    # ── Metric cards ──────────────────────────────────────────────────────────
    try:
        total_assets      = int(db_query("SELECT COUNT(*) AS n FROM assets").iloc[0]["n"])
        proc_bills        = int(db_query("SELECT COUNT(*) AS n FROM procurement").iloc[0]["n"])
        unallocated       = int(db_query("""
            SELECT COUNT(*) AS n FROM assets a
            WHERE NOT EXISTS (
                SELECT 1 FROM allocations al
                WHERE al.asset_id=a.asset_id AND al.is_current=TRUE)
        """).iloc[0]["n"])
        network_devices   = int(db_query("SELECT COUNT(*) AS n FROM network_config").iloc[0]["n"])
        acad_desktops     = int(db_query(
            "SELECT COUNT(*) AS n FROM assets a JOIN asset_types t ON a.type_id=t.type_id "
            "WHERE LOWER(t.type_name)='desktop' AND LOWER(a.purpose)='academic'").iloc[0]["n"])
        offl_desktops     = int(db_query(
            "SELECT COUNT(*) AS n FROM assets a JOIN asset_types t ON a.type_id=t.type_id "
            "WHERE LOWER(t.type_name)='desktop' AND LOWER(a.purpose)='official'").iloc[0]["n"])
        total_laptops     = int(db_query(
            "SELECT COUNT(*) AS n FROM assets a JOIN asset_types t ON a.type_id=t.type_id "
            "WHERE LOWER(t.type_name)='laptop'").iloc[0]["n"])
        total_printers    = int(db_query(
            "SELECT COUNT(*) AS n FROM assets a JOIN asset_types t ON a.type_id=t.type_id "
            "WHERE LOWER(t.type_name)='printer'").iloc[0]["n"])

        # Row 1
        cols = st.columns(6)
        cols[0].markdown(_card("Total Assets",      f"{total_assets:,}"),   unsafe_allow_html=True)
        cols[1].markdown(_card("Academic Desktops", f"{acad_desktops:,}"),  unsafe_allow_html=True)
        cols[2].markdown(_card("Official Desktops", f"{offl_desktops:,}"),  unsafe_allow_html=True)
        cols[3].markdown(_card("Laptops",           f"{total_laptops:,}"),  unsafe_allow_html=True)
        cols[4].markdown(_card("Printers",          f"{total_printers:,}"), unsafe_allow_html=True)
        cols[5].markdown(_card("Unallocated",       f"{unallocated:,}"),    unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Could not load summary cards: {e}")

    st.divider()

    # ── Procurement Trend | Top Departments ───────────────────────────────────
    try:
        import plotly.express as px
        HAS_PLOTLY = True
    except ImportError:
        HAS_PLOTLY = False

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Procurement Trend by Year")
        try:
            df_trend = db_query("""
                SELECT EXTRACT(YEAR FROM purchase_date)::int AS "Year",
                       COUNT(*) AS "Bills",
                       COALESCE(SUM(total_amount),0)::bigint AS "Amount (₹)"
                FROM procurement
                WHERE purchase_date IS NOT NULL
                GROUP BY 1 ORDER BY 1
            """)
            if not df_trend.empty:
                if HAS_PLOTLY:
                    fig = px.bar(df_trend, x="Year", y="Bills",
                                 color_discrete_sequence=[GOLD],
                                 text="Bills",
                                 hover_data=["Amount (₹)"])
                    fig.update_layout(plot_bgcolor="#fff", paper_bgcolor="#fff",
                                      margin=dict(t=10, b=10, l=10, r=10),
                                      xaxis=dict(tickmode="linear"),
                                      yaxis_title="No. of Bills",
                                      showlegend=False)
                    fig.update_traces(textposition="outside")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.bar_chart(df_trend.set_index("Year")["Bills"])
            else:
                st.info("No procurement data with dates available.")
        except Exception as e:
            st.error(str(e))

    with col2:
        st.subheader("Top 10 Departments by Assets")
        try:
            df_dept = db_query("""
                SELECT d.dept_name AS "Department", COUNT(*) AS "Assets"
                FROM allocations al
                JOIN departments d ON al.dept_id = d.dept_id
                WHERE al.is_current = TRUE
                GROUP BY d.dept_name ORDER BY COUNT(*) DESC
                LIMIT 10
            """)
            if not df_dept.empty and HAS_PLOTLY:
                fig2 = px.bar(df_dept, x="Assets", y="Department",
                              orientation="h",
                              color_discrete_sequence=[NAVY],
                              text="Assets")
                fig2.update_layout(plot_bgcolor="#fff", paper_bgcolor="#fff",
                                   margin=dict(t=10, b=10, l=10, r=10),
                                   yaxis=dict(autorange="reversed"),
                                   xaxis_title="Assets",
                                   showlegend=False)
                fig2.update_traces(textposition="outside")
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.dataframe(df_dept, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(str(e))

    st.divider()

    # ── Assets by Brand | Recent Procurement ─────────────────────────────────
    col3, col4 = st.columns(2)

    with col3:
        st.subheader("Assets by Brand")
        try:
            df_make = db_query("""
                SELECT m.make_name AS "Brand", COUNT(*) AS "Assets"
                FROM assets a LEFT JOIN makes m ON a.make_id=m.make_id
                GROUP BY m.make_name ORDER BY COUNT(*) DESC
                LIMIT 10
            """)
            if not df_make.empty and HAS_PLOTLY:
                fig3 = px.pie(df_make, names="Brand", values="Assets",
                              color_discrete_sequence=px.colors.sequential.Oranges_r)
                fig3.update_layout(margin=dict(t=10, b=10, l=10, r=10),
                                   paper_bgcolor="#fff")
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.dataframe(df_make, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(str(e))

    with col4:
        st.subheader("Recent Procurement (Last 10)")
        try:
            df_recent = db_query("""
                SELECT p.bill_no       AS "Bill No.",
                       p.purchase_date AS "Date",
                       m.make_name     AS "Brand",
                       p.model_name    AS "Model",
                       p.qty           AS "Qty",
                       p.total_amount  AS "Amount (₹)"
                FROM procurement p
                LEFT JOIN makes m ON p.make_id = m.make_id
                WHERE p.purchase_date IS NOT NULL
                ORDER BY p.purchase_date DESC
                LIMIT 10
            """)
            st.dataframe(df_recent, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 – Assets
# ══════════════════════════════════════════════════════════════════════════════
def page_assets():
    st.title("Assets")
    st.divider()

    # ── Asset Inventory (always shown at top) ─────────────────────────────────
    st.subheader("Asset Inventory")

    types_df = db_query("SELECT type_id, type_name FROM asset_types ORDER BY type_name")
    depts_df = db_query("SELECT dept_id, dept_name FROM departments ORDER BY dept_name")
    makes_df = db_query("SELECT make_id, make_name FROM makes ORDER BY make_name")

    fc1, fc2, fc3 = st.columns([2, 2, 3])
    with fc1:
        sel_type = st.selectbox("Asset Type", ["All Types"] + types_df["type_name"].tolist())
    with fc2:
        sel_dept = st.selectbox("Department", ["All Departments"] + depts_df["dept_name"].tolist())
    with fc3:
        search = st.text_input("Search by Serial Number or Model",
                               placeholder="e.g.  SN12345  or  OptiPlex 7050")

    sql = """
        SELECT
            a.asset_id          AS "Asset ID",
            t.type_name         AS "Type",
            a.serial_no         AS "Serial No.",
            m.make_name         AS "Make",
            a.model_name        AS "Model",
            a.purpose           AS "Purpose",
            a.status            AS "Status",
            d.dept_name         AS "Department",
            l.location_name     AS "Location"
        FROM assets a
        LEFT JOIN asset_types  t  ON a.type_id  = t.type_id
        LEFT JOIN makes        m  ON a.make_id  = m.make_id
        LEFT JOIN allocations al ON a.asset_id  = al.asset_id AND al.is_current = TRUE
        LEFT JOIN departments  d  ON al.dept_id = d.dept_id
        LEFT JOIN locations    l  ON al.location_id = l.location_id
        WHERE 1=1
    """
    params = []
    if sel_type != "All Types":
        sql += " AND t.type_name = %s"
        params.append(sel_type)
    if sel_dept != "All Departments":
        sql += " AND d.dept_name = %s"
        params.append(sel_dept)
    if search.strip():
        sql += " AND (a.serial_no ILIKE %s OR a.model_name ILIKE %s)"
        params += [f"%{search.strip()}%", f"%{search.strip()}%"]
    sql += " ORDER BY a.asset_id"

    try:
        df = db_query(sql, params or None)
        st.caption(f"{len(df):,} asset(s) found")
        st.dataframe(df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Could not load assets: {e}")
        df = pd.DataFrame()

    # ── Add / Edit / Delete ───────────────────────────────────────────────────
    st.divider()
    st.subheader("Manage Assets")
    tab_add, tab_edit, tab_del = st.tabs(
        ["Add New Asset", "Edit Asset", "Delete Asset"]
    )

    with tab_add:
        st.subheader("Add New Asset")
        if st.session_state.get("_success_padd"):
            st.success("✅  Asset added successfully to the database!")
            del st.session_state["_success_padd"]
        with st.form("form_add_asset"):
            r1, r2, r3 = st.columns(3)
            with r1:
                na_type   = st.selectbox("Asset Type *",   types_df["type_name"].tolist())
                na_serial = st.text_input("Serial Number *")
            with r2:
                na_make   = st.selectbox("Make / Brand *", makes_df["make_name"].tolist())
                na_model  = st.text_input("Model *")
            with r3:
                na_purpose = st.selectbox("Purpose *",  ["academic", "official"])
                na_status  = st.selectbox("Status *",   ["active", "inactive", "under repair", "disposed"])

            if st.form_submit_button("🔍  Review & Add Asset", use_container_width=True):
                if not na_serial.strip() or not na_model.strip():
                    st.error("Serial Number and Model are required.")
                else:
                    tid = int(types_df.loc[types_df["type_name"] == na_type, "type_id"].iloc[0])
                    mid = int(makes_df.loc[makes_df["make_name"] == na_make, "make_id"].iloc[0])
                    st.session_state["padd"] = {
                        "action": "ADD",
                        "summary": f"{na_type}  |  {na_make}  {na_model.strip()}",
                        "details": [
                            ("Serial Number", na_serial.strip()),
                            ("Asset Type",    na_type),
                            ("Make / Brand",  na_make),
                            ("Model",         na_model.strip()),
                            ("Purpose",       na_purpose),
                            ("Status",        na_status),
                        ],
                        "fn": lambda tid=tid, mid=mid: db_exec(
                            "INSERT INTO assets (type_id, serial_no, make_id, model_name, purpose, status) "
                            "VALUES (%s,%s,%s,%s,%s,%s)",
                            (tid, na_serial.strip(), mid, na_model.strip(), na_purpose, na_status),
                        ),
                    }
        confirm_dialog("padd")

    with tab_edit:
        st.subheader("Edit Asset")
        st.caption("Use the filters above (Type, Department, Search) to narrow the list, then pick the asset.")
        if df.empty:
            st.info("No assets match the current filters. Clear the filters to find the asset you want to edit.")
        else:
            # Readable label: "Serial – Make Model (Type)"
            label_map = {
                row["Asset ID"]: f"{row['Serial No.']}  —  {row['Make']} {row['Model']}  ({row['Type']})"
                for _, row in df.iterrows()
            }
            sel_label = st.selectbox(
                "Select asset to edit",
                options=list(label_map.values()),
                key="edit_asset_label",
            )
            sel_id = [k for k, v in label_map.items() if v == sel_label][0]
            row = df[df["Asset ID"] == sel_id].iloc[0]

            with st.form("form_edit_asset"):
                r1, r2, r3 = st.columns(3)
                tlist = types_df["type_name"].tolist()
                mlist = makes_df["make_name"].tolist()
                with r1:
                    ea_type   = st.selectbox("Asset Type", tlist,
                                             index=tlist.index(row["Type"]) if row["Type"] in tlist else 0)
                    ea_serial = st.text_input("Serial Number", value=str(row["Serial No."]))
                with r2:
                    ea_make   = st.selectbox("Make / Brand", mlist,
                                             index=mlist.index(row["Make"]) if row["Make"] in mlist else 0)
                    ea_model  = st.text_input("Model", value=str(row["Model"]))
                with r3:
                    purposes = ["academic", "official"]
                    statuses = ["active", "inactive", "under repair", "disposed"]
                    ea_purpose = st.selectbox("Purpose", purposes,
                                              index=purposes.index(row["Purpose"]) if row["Purpose"] in purposes else 0)
                    ea_status  = st.selectbox("Status", statuses,
                                              index=statuses.index(row["Status"]) if row["Status"] in statuses else 0)

                if st.form_submit_button("🔍  Review & Save Changes", use_container_width=True):
                    tid = int(types_df.loc[types_df["type_name"] == ea_type, "type_id"].iloc[0])
                    mid = int(makes_df.loc[makes_df["make_name"] == ea_make, "make_id"].iloc[0])
                    st.session_state["pedit_asset"] = {
                        "action": "SAVE",
                        "summary": f"Asset ID {sel_id}  |  {ea_type}  |  {ea_make}  {ea_model.strip()}",
                        "details": [
                            ("Asset ID",      str(sel_id)),
                            ("Serial Number", ea_serial.strip()),
                            ("Asset Type",    ea_type),
                            ("Make / Brand",  ea_make),
                            ("Model",         ea_model.strip()),
                            ("Purpose",       ea_purpose),
                            ("Status",        ea_status),
                        ],
                        "fn": lambda tid=tid, mid=mid, sid=sel_id: db_exec(
                            "UPDATE assets SET type_id=%s, serial_no=%s, make_id=%s, model_name=%s, "
                            "purpose=%s, status=%s WHERE asset_id=%s",
                            (tid, ea_serial.strip(), mid, ea_model.strip(), ea_purpose, ea_status, sid),
                        ),
                    }
        if st.session_state.get("_success_pedit_asset"):
            st.success("✅  Asset updated successfully in the database!")
            del st.session_state["_success_pedit_asset"]
        confirm_dialog("pedit_asset")

    with tab_del:
        st.subheader("Delete Asset")
        st.warning("⚠️  Deleting an asset is permanent and cannot be undone.")
        st.caption("Step 1 — Use the filters below to find the asset. Step 2 — Select it. Step 3 — Confirm.")

        # Independent filters just for the delete tab
        dc1, dc2, dc3 = st.columns([2, 2, 3])
        with dc1:
            del_type = st.selectbox("Filter by Type",
                                    ["All Types"] + types_df["type_name"].tolist(),
                                    key="del_type_filter")
        with dc2:
            del_dept = st.selectbox("Filter by Department",
                                    ["All Departments"] + depts_df["dept_name"].tolist(),
                                    key="del_dept_filter")
        with dc3:
            del_search = st.text_input("Search Serial No. or Model",
                                       placeholder="e.g.  SN12345  or  OptiPlex",
                                       key="del_search_filter")

        # Build filtered query for delete tab
        del_sql = """
            SELECT a.asset_id, t.type_name, a.serial_no, m.make_name, a.model_name,
                   a.status, d.dept_name, l.location_name
            FROM assets a
            JOIN asset_types t ON a.type_id = t.type_id
            LEFT JOIN makes       m ON a.make_id  = m.make_id
            LEFT JOIN allocations al ON a.asset_id = al.asset_id AND al.is_current = TRUE
            LEFT JOIN departments  d ON al.dept_id  = d.dept_id
            LEFT JOIN locations    l ON al.location_id = l.location_id
            WHERE 1=1
        """
        del_params = []
        if del_type != "All Types":
            del_sql += " AND t.type_name = %s"
            del_params.append(del_type)
        if del_dept != "All Departments":
            del_sql += " AND d.dept_name = %s"
            del_params.append(del_dept)
        if del_search.strip():
            del_sql += " AND (a.serial_no ILIKE %s OR a.model_name ILIKE %s)"
            del_params += [f"%{del_search.strip()}%", f"%{del_search.strip()}%"]
        del_sql += " ORDER BY t.type_name, a.serial_no LIMIT 200"

        try:
            del_df = db_query(del_sql, del_params or None)
        except Exception as e:
            st.error(str(e))
            del_df = pd.DataFrame()

        if del_df.empty:
            st.info("No assets found. Try different filters.")
        else:
            st.markdown(f"**{len(del_df):,} asset(s) found — select one to delete:**")

            # Readable label
            del_label_map = {
                row["asset_id"]: (
                    f"{row['serial_no']}  —  {row['make_name']} {row['model_name']}"
                    f"  ({row['type_name']})"
                    + (f"  |  {row['dept_name']}" if row["dept_name"] else "")
                )
                for _, row in del_df.iterrows()
            }
            del_sel_label = st.selectbox(
                "Select asset to delete",
                options=list(del_label_map.values()),
                key="del_asset_label",
            )
            del_id = [k for k, v in del_label_map.items() if v == del_sel_label][0]
            drow = del_df[del_df["asset_id"] == del_id].iloc[0]

            if st.button("🗑️  Prepare to Delete", key="btn_del_asset_prep", type="primary"):
                st.session_state["pdel_asset"] = {
                    "action": "DELETE",
                    "summary": f"{drow['type_name']}  |  {drow['make_name']} {drow['model_name']}",
                    "details": [
                        ("Asset ID",    str(del_id)),
                        ("Serial No.",  drow["serial_no"]),
                        ("Type",        drow["type_name"]),
                        ("Make",        drow["make_name"]),
                        ("Model",       drow["model_name"]),
                        ("Status",      drow["status"]),
                        ("Department",  str(drow["dept_name"] or "—")),
                        ("Location",    str(drow["location_name"] or "—")),
                    ],
                    "fn": lambda did=del_id: [
                        db_exec("DELETE FROM allocations    WHERE asset_id        = %s", (did,)),
                        db_exec("DELETE FROM computer_specs WHERE asset_id        = %s", (did,)),
                        db_exec("DELETE FROM network_config WHERE asset_id        = %s", (did,)),
                        db_exec("DELETE FROM peripherals    WHERE parent_asset_id = %s", (did,)),
                        db_exec("DELETE FROM server_details WHERE asset_id        = %s", (did,)),
                        db_exec("DELETE FROM assets         WHERE asset_id        = %s", (did,)),
                    ],
                }
            if st.session_state.get("_success_pdel_asset"):
                st.success("✅  Asset permanently deleted from the database!")
                del st.session_state["_success_pdel_asset"]
            confirm_dialog("pdel_asset")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 – Procurement
# ══════════════════════════════════════════════════════════════════════════════
def page_procurement():
    st.title("Procurement")
    st.caption("All purchase bills for IT equipment")
    st.divider()

    try:
        df = db_query("""
            SELECT
                p.procurement_id    AS "ID",
                m.make_name         AS "Make / Brand",
                p.model_name        AS "Model",
                p.configuration     AS "Configuration (CPU / RAM / HDD)",
                p.bill_no           AS "Bill Number",
                p.purchase_date     AS "Purchase Date",
                p.purchase_year     AS "Year",
                p.qty               AS "Quantity",
                p.unit_price        AS "Unit Price (₹)",
                p.total_amount      AS "Total Amount (₹)",
                p.folio_number      AS "Folio No.",
                p.register_book_no  AS "Register Book"
            FROM procurement p
            LEFT JOIN makes m ON p.make_id = m.make_id
            ORDER BY p.purchase_date DESC NULLS LAST
        """)
        st.markdown(f"**{len(df):,} procurement record(s)**")
        st.dataframe(df, use_container_width=True, hide_index=True)
        if not df.empty:
            total = float(df["Total Amount (₹)"].sum())
            st.markdown(
                f"<div style='text-align:right;font-size:1.1rem;font-weight:700;"
                f"color:{GOLD};padding:8px 4px;'>"
                f"Total Procurement Value: {fmt_currency(total)}&nbsp;&nbsp;"
                f"<span style='font-size:0.85rem;color:#888;'>"
                f"(₹{total:,.2f})</span></div>",
                unsafe_allow_html=True,
            )
    except Exception as e:
        st.error(f"Could not load procurement records: {e}")

    st.divider()
    st.subheader("Add New Procurement Record")
    makes_df = db_query("SELECT make_id, make_name FROM makes ORDER BY make_name")
    types_df = db_query("SELECT type_id, type_name FROM asset_types ORDER BY type_name")

    with st.form("form_add_procurement"):
        r1, r2, r3 = st.columns(3)
        with r1:
            pm_make  = st.selectbox("Make / Brand *", makes_df["make_name"].tolist())
            pm_type  = st.selectbox("Asset Type *",   types_df["type_name"].tolist())
            pm_model = st.text_input("Model *")
        with r2:
            pm_bill  = st.text_input("Bill Number *")
            pm_date  = st.date_input("Purchase Date *", value=date.today())
        with r3:
            pm_qty   = st.number_input("Quantity *",         min_value=1,   step=1,     value=1)
            pm_unit  = st.number_input("Unit Price (₹) *",   min_value=0.0, step=100.0, value=0.0)
            pm_total = st.number_input("Total Amount (₹) *", min_value=0.0, step=100.0, value=0.0)

        if st.form_submit_button("🔍  Review & Add Record", use_container_width=True):
            if not pm_model.strip() or not pm_bill.strip():
                st.error("Model and Bill Number are required.")
            else:
                mid = int(makes_df.loc[makes_df["make_name"] == pm_make, "make_id"].iloc[0])
                tid = int(types_df.loc[types_df["type_name"] == pm_type, "type_id"].iloc[0])
                st.session_state["pproc_add"] = {
                    "action": "ADD",
                    "summary": f"{pm_make}  —  {pm_model.strip()}  |  Bill: {pm_bill.strip()}",
                    "details": [
                        ("Make / Brand",     pm_make),
                        ("Asset Type",       pm_type),
                        ("Model",            pm_model.strip()),
                        ("Bill Number",      pm_bill.strip()),
                        ("Purchase Date",    str(pm_date)),
                        ("Quantity",         str(pm_qty)),
                        ("Unit Price",       f"₹{pm_unit:,.2f}"),
                        ("Total Amount",     f"₹{pm_total:,.2f}"),
                    ],
                    "fn": lambda mid=mid, tid=tid: db_exec(
                        "INSERT INTO procurement (make_id, type_id, model_name, bill_no, "
                        "purchase_date, qty, unit_price, total_amount) "
                        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                        (mid, tid, pm_model.strip(), pm_bill.strip(),
                         pm_date, pm_qty, pm_unit, pm_total),
                    ),
                }
    if st.session_state.get("_success_pproc_add"):
        st.success("✅  Procurement record added to the database!")
        del st.session_state["_success_pproc_add"]
    confirm_dialog("pproc_add")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 – Intercom Directory
# ══════════════════════════════════════════════════════════════════════════════
def page_intercom():
    st.title("Intercom Directory")
    st.caption("University internal telephone extension list")
    st.divider()

    try:
        df = db_query("""
            SELECT intercom_id  AS "ID",
                   extension    AS "Extension No.",
                   contact_name AS "Name",
                   dept_name_raw AS "Department"
            FROM intercom_directory
            ORDER BY extension
        """)
        st.markdown(f"**{len(df):,} extension(s) listed**")
        st.dataframe(df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Could not load intercom directory: {e}")
        df = pd.DataFrame()

    st.divider()
    tab_add, tab_edit, tab_del = st.tabs(
        ["Add Extension", "Edit Extension", "Delete Extension"]
    )

    with tab_add:
        st.subheader("Add New Extension")
        if st.session_state.get("_success_pic_add"):
            st.success("✅  Extension added to the database!")
            del st.session_state["_success_pic_add"]
        with st.form("form_add_intercom"):
            ic1, ic2, ic3 = st.columns(3)
            with ic1: ia_ext  = st.text_input("Extension Number *")
            with ic2: ia_name = st.text_input("Person Name *")
            with ic3: ia_dept = st.text_input("Department", placeholder="e.g. Computer Science")
            if st.form_submit_button("🔍  Review & Add Extension", use_container_width=True):
                if not ia_ext.strip() or not ia_name.strip():
                    st.error("Extension number and person name are required.")
                else:
                    st.session_state["pic_add"] = {
                        "action": "ADD",
                        "summary": f"Ext {ia_ext.strip()}  —  {ia_name.strip()}",
                        "details": [
                            ("Extension No.", ia_ext.strip()),
                            ("Person Name",   ia_name.strip()),
                            ("Department",    ia_dept.strip() or "—"),
                        ],
                        "fn": lambda: db_exec(
                            "INSERT INTO intercom_directory (extension, contact_name, dept_name_raw) "
                            "VALUES (%s,%s,%s)",
                            (ia_ext.strip(), ia_name.strip(), ia_dept.strip()),
                        ),
                    }
        confirm_dialog("pic_add")

    with tab_edit:
        st.subheader("Edit Extension")
        if df.empty:
            st.info("No entries found.")
        else:
            ic_edit_label_map = {
                row["ID"]: f"Ext {row['Extension No.']}  —  {row['Name']}  ({row['Department']})"
                for _, row in df.iterrows()
            }
            edit_label = st.selectbox(
                "Select extension to edit",
                options=list(ic_edit_label_map.values()),
                key="ic_edit_label",
            )
            sel_id = [k for k, v in ic_edit_label_map.items() if v == edit_label][0]
            row = df[df["ID"] == sel_id].iloc[0]
            with st.form("form_edit_intercom"):
                ic1, ic2, ic3 = st.columns(3)
                with ic1: ie_ext  = st.text_input("Extension Number", value=str(row["Extension No."]))
                with ic2: ie_name = st.text_input("Person Name",      value=str(row["Name"]))
                with ic3: ie_dept = st.text_input("Department",        value=str(row["Department"] or ""))
                if st.form_submit_button("🔍  Review & Save Changes", use_container_width=True):
                    st.session_state["pic_edit"] = {
                        "action": "SAVE",
                        "summary": f"Ext {ie_ext.strip()}  —  {ie_name.strip()}",
                        "details": [
                            ("Intercom ID",   str(sel_id)),
                            ("Extension No.", ie_ext.strip()),
                            ("Person Name",   ie_name.strip()),
                            ("Department",    ie_dept.strip() or "—"),
                        ],
                        "fn": lambda sid=sel_id: db_exec(
                            "UPDATE intercom_directory SET extension=%s, contact_name=%s, "
                            "dept_name_raw=%s WHERE intercom_id=%s",
                            (ie_ext.strip(), ie_name.strip(), ie_dept.strip(), sid),
                        ),
                    }
        if st.session_state.get("_success_pic_edit"):
            st.success("✅  Extension updated in the database!")
            del st.session_state["_success_pic_edit"]
        confirm_dialog("pic_edit")

    with tab_del:
        st.subheader("Delete Extension")
        st.warning("⚠️  This will permanently remove the extension entry.")
        if df.empty:
            st.info("No entries found.")
        else:
            ic_label_map = {
                row["ID"]: f"Ext {row['Extension No.']}  —  {row['Name']}  ({row['Department']})"
                for _, row in df.iterrows()
            }
            del_label = st.selectbox(
                "Select extension to delete",
                options=list(ic_label_map.values()),
                key="ic_del_label",
            )
            del_id = [k for k, v in ic_label_map.items() if v == del_label][0]
            drow = df[df["ID"] == del_id].iloc[0]
            st.error(
                f"⚠️  Permanently deleting — Ext **{drow['Extension No.']}** "
                f"| {drow['Name']} ({drow['Department']})"
            )
            if st.button("🗑️  Prepare to Delete", key="btn_del_ic_prep", type="primary"):
                st.session_state["pic_del"] = {
                    "action": "DELETE",
                    "summary": f"Ext {drow['Extension No.']}  —  {drow['Name']}",
                    "details": [
                        ("Intercom ID",   str(del_id)),
                        ("Extension No.", str(drow["Extension No."])),
                        ("Person Name",   drow["Name"]),
                        ("Department",    str(drow["Department"] or "—")),
                    ],
                    "fn": lambda did=del_id: db_exec(
                        "DELETE FROM intercom_directory WHERE intercom_id = %s", (did,)
                    ),
                }
            if st.session_state.get("_success_pic_del"):
                st.success("✅  Extension permanently deleted from the database!")
                del st.session_state["_success_pic_del"]
            confirm_dialog("pic_del")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5 – Network Devices
# ══════════════════════════════════════════════════════════════════════════════
def page_network():
    st.title("Network Devices")
    st.caption("IP addresses and network details for printers, switches, and WiFi access points")
    st.divider()

    try:
        df = db_query("""
            SELECT
                a.asset_id      AS "Asset ID",
                t.type_name     AS "Device Type",
                m.make_name     AS "Make",
                a.model_name    AS "Model",
                a.serial_no     AS "Serial No.",
                n.ip_address    AS "IP Address",
                n.subnet_mask   AS "Subnet Mask",
                n.mac_address   AS "MAC Address",
                n.hostname      AS "Hostname",
                n.network_type  AS "Network Type",
                a.status        AS "Status",
                d.dept_name     AS "Department",
                l.location_name AS "Location"
            FROM network_config n
            JOIN assets       a  ON n.asset_id       = a.asset_id
            JOIN asset_types  t  ON a.type_id         = t.type_id
            LEFT JOIN makes        m  ON a.make_id          = m.make_id
            LEFT JOIN allocations al ON a.asset_id    = al.asset_id AND al.is_current = TRUE
            LEFT JOIN departments  d  ON al.dept_id   = d.dept_id
            LEFT JOIN locations    l  ON al.location_id = l.location_id
            ORDER BY t.type_name, n.ip_address
        """)
        if not df.empty:
            dt_opts = ["All Device Types"] + sorted(df["Device Type"].dropna().unique().tolist())
            sel_dt = st.selectbox("Filter by Device Type", dt_opts)
            if sel_dt != "All Device Types":
                df = df[df["Device Type"] == sel_dt]
        st.markdown(f"**{len(df):,} device(s) found**")
        st.dataframe(df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Could not load network devices: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 6 – Download Data
# ══════════════════════════════════════════════════════════════════════════════

# Each entry: (Sheet name, SQL query)
DOWNLOAD_SHEETS = [
    ("Assets", """
        SELECT a.asset_id, t.type_name AS asset_type, a.serial_no, m.make_name AS make,
               a.model_name AS model, a.purpose, a.status,
               d.dept_name AS department, l.location_name AS location
        FROM assets a
        JOIN asset_types t ON a.type_id = t.type_id
        LEFT JOIN makes       m ON a.make_id  = m.make_id
        LEFT JOIN allocations al ON a.asset_id = al.asset_id AND al.is_current = TRUE
        LEFT JOIN departments  d ON al.dept_id  = d.dept_id
        LEFT JOIN locations    l ON al.location_id = l.location_id
        ORDER BY a.asset_id
    """),
    ("Procurement", """
        SELECT p.procurement_id, m.make_name AS make, t.type_name AS asset_type,
               p.model_name AS model, p.bill_no, p.purchase_date,
               p.qty AS quantity, p.unit_price, p.total_amount
        FROM procurement p
        LEFT JOIN makes m ON p.make_id = m.make_id
        LEFT JOIN asset_types t ON p.type_id = t.type_id
        ORDER BY p.purchase_date DESC
    """),
    ("Allocations", """
        SELECT al.allocation_id, a.asset_id, t.type_name AS asset_type,
               a.serial_no, d.dept_name AS department, l.location_name AS location,
               al.assigned_user, al.is_current, al.allocated_at, al.deallocated_at, al.notes
        FROM allocations al
        JOIN assets       a ON al.asset_id    = a.asset_id
        JOIN asset_types  t ON a.type_id       = t.type_id
        LEFT JOIN departments d ON al.dept_id  = d.dept_id
        LEFT JOIN locations   l ON al.location_id = l.location_id
        ORDER BY al.allocation_id
    """),
    ("Computer Specs", """
        SELECT cs.spec_id, a.asset_id, t.type_name AS asset_type, a.serial_no,
               m.make_name AS make, a.model_name AS model,
               cs.cpu_model, cs.processor_speed, cs.ram, cs.hdd,
               cs.bit_type, cs.configuration
        FROM computer_specs cs
        JOIN assets      a ON cs.asset_id = a.asset_id
        JOIN asset_types t ON a.type_id    = t.type_id
        LEFT JOIN makes       m ON a.make_id    = m.make_id
        ORDER BY cs.spec_id
    """),
    ("Network Config", """
        SELECT n.config_id, a.asset_id, t.type_name AS asset_type,
               m.make_name AS make, a.model_name AS model, a.serial_no,
               n.ip_address, n.subnet_mask, n.mac_address, n.hostname,
               n.ssid, n.network_type, n.port_capacity, n.ports_in_use
        FROM network_config n
        JOIN assets      a ON n.asset_id = a.asset_id
        JOIN asset_types t ON a.type_id   = t.type_id
        LEFT JOIN makes       m ON a.make_id   = m.make_id
        ORDER BY n.config_id
    """),
    ("Peripherals", """
        SELECT p.peripheral_id, a.asset_id AS parent_asset_id,
               t.type_name AS parent_type, a.serial_no AS parent_serial,
               p.peripheral_type, mk.make_name AS make,
               p.model_name AS model, p.serial_no, p.size
        FROM peripherals p
        JOIN assets      a  ON p.parent_asset_id = a.asset_id
        JOIN asset_types t  ON a.type_id          = t.type_id
        LEFT JOIN makes  mk ON p.make_id           = mk.make_id
        ORDER BY p.peripheral_id
    """),
    ("Server Details", """
        SELECT sd.server_id, a.asset_id, m.make_name AS make,
               a.model_name AS model, a.serial_no,
               sd.server_type, sd.server_name, sd.ip_address, sd.access_url
        FROM server_details sd
        JOIN assets a ON sd.asset_id = a.asset_id
        LEFT JOIN makes  m ON a.make_id   = m.make_id
        ORDER BY sd.server_id
    """),
    ("Intercom Directory", """
        SELECT intercom_id, extension, contact_name AS person_name, dept_name_raw AS department
        FROM intercom_directory
        ORDER BY extension
    """),
    ("Departments", """
        SELECT dept_id, dept_name FROM departments ORDER BY dept_name
    """),
    ("Locations", """
        SELECT location_id, location_name, block, floor, room_no, location_type
        FROM locations ORDER BY location_name
    """),
    ("Asset Types", """
        SELECT type_id, type_name FROM asset_types ORDER BY type_name
    """),
    ("Makes", """
        SELECT make_id, make_name FROM makes ORDER BY make_name
    """),
]


def build_excel() -> bytes:
    """Build an in-memory Excel workbook with one sheet per table."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, sql in DOWNLOAD_SHEETS:
            try:
                df = db_query(sql)
            except Exception:
                df = pd.DataFrame({"Error": ["Could not load this sheet"]})
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Basic header styling
            ws = writer.sheets[sheet_name]
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            gold_fill  = PatternFill("solid", fgColor="C8830A")
            white_font = Font(bold=True, color="FFFFFF")
            thin_border = Border(
                bottom=Side(style="thin", color="C8830A")
            )
            for cell in ws[1]:
                cell.fill      = gold_fill
                cell.font      = white_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border    = thin_border

            # Auto-width columns
            for col in ws.columns:
                max_len = max(
                    (len(str(c.value)) for c in col if c.value is not None),
                    default=10,
                )
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    buf.seek(0)
    return buf.read()


ORIGINAL_EXCEL_PATH = r"D:\MBA\MBA Term 6\IT PRPJECT\IT Assests Data.xlsx"


def page_download():
    st.title("Download Data")
    st.caption("Choose your preferred download format")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["Database Export (Structured)", "Original Format (Legacy Excel)", "Backup Manager"])

    # ── Tab 1: DB Export ──────────────────────────────────────────────────────
    with tab1:
        st.markdown("#### Export from live database")
        st.markdown(
            "Downloads the latest data from the database as a clean Excel workbook. "
            "One sheet per table with gold-styled headers."
        )
        st.markdown("""
        | Sheet | Contents |
        |---|---|
        | Assets | All physical assets with type, make, model, department and location |
        | Procurement | All purchase bills with bill number, date, quantity and prices |
        | Allocations | Full allocation history — which asset went to which department and when |
        | Computer Specs | CPU, RAM, HDD and OS details for desktops and laptops |
        | Network Config | IP address, subnet, MAC address and hostname for network devices |
        | Peripherals | Monitors, keyboards and mice linked to desktops |
        | Server Details | Server type, name and access URL for servers |
        | Intercom Directory | Extension numbers and person names |
        | Departments | All university departments |
        | Locations | All physical locations and rooms |
        | Asset Types | Lookup list of asset categories |
        | Makes | Lookup list of brands / manufacturers |
        """)

        sel_sheets = st.multiselect(
            "Select sheets to include (leave blank = all sheets)",
            options=[s[0] for s in DOWNLOAD_SHEETS],
            default=[],
            placeholder="All sheets included by default",
        )

        generate = st.button("⚙️  Prepare Excel File", use_container_width=True, key="gen_db")

        if generate or st.session_state.get("excel_ready"):
            with st.spinner("Building Excel file from database…"):
                sheets_to_use = (
                    [(n, s) for n, s in DOWNLOAD_SHEETS if n in sel_sheets]
                    if sel_sheets else DOWNLOAD_SHEETS
                )
                buf = io.BytesIO()
                try:
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
                        gold_fill   = PatternFill("solid", fgColor="C8830A")
                        white_font  = Font(bold=True, color="FFFFFF")
                        thin_border = Border(bottom=Side(style="thin", color="C8830A"))

                        for sheet_name, sql in sheets_to_use:
                            try:
                                df = db_query(sql)
                            except Exception as ex:
                                df = pd.DataFrame({"Error": [str(ex)]})
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            ws = writer.sheets[sheet_name]
                            for cell in ws[1]:
                                cell.fill      = gold_fill
                                cell.font      = white_font
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                                cell.border    = thin_border
                            for col in ws.columns:
                                max_len = max(
                                    (len(str(c.value)) for c in col if c.value is not None),
                                    default=10,
                                )
                                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

                    buf.seek(0)
                    st.session_state.excel_ready = True
                    st.success(f"✅  Excel file ready — {len(sheets_to_use)} sheet(s) included.")
                    st.download_button(
                        label="📥  Download Database Export",
                        data=buf.read(),
                        file_name="Akal_University_IT_Assets_DB.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Could not build Excel file: {e}")
        else:
            st.session_state.excel_ready = False

    # ── Tab 2: Original Excel ─────────────────────────────────────────────────
    with tab2:
        st.markdown("#### Download original IT Assets file")
        st.markdown(
            "This is the original Excel workbook used before the system was set up — "
            "same sheets, same column names, same layout."
        )

        orig_sheets = [
            "Desktop(A)", "Desktop(O)", "Laptops", "Servers",
            "Printers", "Printers1", "Projectors", "Scanners",
            "WebCam", "SmartClasses", "Wifi", "Switches",
            "Intercom", "NAAC ",
        ]
        st.markdown(
            "**Sheets included:** " + ", ".join(f"`{s.strip()}`" for s in orig_sheets)
        )

        if os.path.exists(ORIGINAL_EXCEL_PATH):
            with open(ORIGINAL_EXCEL_PATH, "rb") as f:
                orig_bytes = f.read()
            st.download_button(
                label="📋  Download Original Excel File",
                data=orig_bytes,
                file_name="IT_Assets_Original.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            size_kb = len(orig_bytes) // 1024
            st.caption(f"File size: {size_kb:,} KB")
        else:
            st.warning(
                f"Original file not found at:\n`{ORIGINAL_EXCEL_PATH}`\n\n"
                "Please ensure the file is accessible on this machine."
            )

    # ── Tab 3: Backup Manager ─────────────────────────────────────────────────
    with tab3:
        st.markdown("#### Backup Manager")
        st.markdown(
            "A full **SQL backup** (`.sql` file) is created automatically "
            "**before every edit or delete** operation using `pg_dump`. "
            "Download it and run it in pgAdmin's Query Tool to restore. "
            "Enter the admin password to unlock downloads."
        )
        st.divider()

        # Load backup list
        try:
            bk_df = db_query("""
                SELECT backup_id, backup_time, trigger_action, trigger_summary,
                       sql_filename, xlsx_filename, file_size_kb
                FROM app_backups
                ORDER BY backup_time DESC
                LIMIT 100
            """)
        except Exception:
            bk_df = pd.DataFrame()

        if bk_df.empty:
            st.info("No backups yet. Backups are created automatically when data is edited or deleted.")
        else:
            bk_df["backup_time"] = pd.to_datetime(bk_df["backup_time"]).dt.strftime("%Y-%m-%d  %H:%M:%S")
            bk_df["file_size_kb"] = bk_df["file_size_kb"].fillna(0).astype(int).astype(str) + " KB"
            st.markdown(f"**{len(bk_df)} backup(s) stored**")
            st.dataframe(
                bk_df[["backup_id", "backup_time", "trigger_action", "trigger_summary", "file_size_kb"]]
                .rename(columns={
                    "backup_id":       "ID",
                    "backup_time":     "Date / Time",
                    "trigger_action":  "Action",
                    "trigger_summary": "What Changed",
                    "file_size_kb":    "Total Size",
                }),
                use_container_width=True, hide_index=True,
            )

            st.divider()
            st.markdown("##### Download a Backup")

            # Date filter
            dates = sorted(
                pd.to_datetime(bk_df["backup_time"]).dt.date.unique().tolist(), reverse=True)
            sel_date = st.selectbox("Filter by date", ["All dates"] + [str(d) for d in dates],
                                    key="bk_date_filter")
            filtered = bk_df[bk_df["backup_time"].str.startswith(sel_date)] \
                       if sel_date != "All dates" else bk_df

            if filtered.empty:
                st.info("No backups on that date.")
            else:
                options = {
                    f"[{row['backup_id']}]  {row['backup_time']}  —  {row['trigger_action']}  |  "
                    f"{str(row['trigger_summary'])[:55]}": row
                    for _, row in filtered.iterrows()
                }
                sel_label = st.selectbox("Select backup", list(options.keys()), key="bk_sel")
                sel_row   = options[sel_label]

                bk_pwd = st.text_input("Admin password to unlock downloads",
                                       type="password", key="bk_pwd")
                if bk_pwd:
                    if bk_pwd == "akal@123":
                        st.success("Password correct.")
                        dc1, dc2 = st.columns(2)

                        # SQL download
                        with dc1:
                            st.markdown("**SQL Backup** — restore in pgAdmin")
                            st.caption("Query Tool → File → Open → F5")
                            sql_f = sel_row.get("sql_filename")
                            if sql_f and str(sql_f) not in ("", "None", "nan"):
                                sql_path = os.path.join(BACKUP_DIR, str(sql_f))
                                if os.path.exists(sql_path):
                                    with open(sql_path, "rb") as f:
                                        st.download_button(
                                            "Download .sql",
                                            data=f.read(),
                                            file_name=str(sql_f),
                                            mime="text/plain",
                                            use_container_width=True,
                                            key="bk_dl_sql",
                                        )
                                else:
                                    st.warning("SQL file not found on disk.")
                            else:
                                st.info("No SQL file for this backup.")

                        # Excel download
                        with dc2:
                            st.markdown("**Excel Backup** — open in Excel")
                            st.caption("All tables as separate sheets")
                            xlsx_f = sel_row.get("xlsx_filename")
                            if xlsx_f and str(xlsx_f) not in ("", "None", "nan"):
                                xlsx_path = os.path.join(BACKUP_DIR, str(xlsx_f))
                                if os.path.exists(xlsx_path):
                                    with open(xlsx_path, "rb") as f:
                                        st.download_button(
                                            "Download .xlsx",
                                            data=f.read(),
                                            file_name=str(xlsx_f),
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            use_container_width=True,
                                            key="bk_dl_xlsx",
                                        )
                                else:
                                    st.warning("Excel file not found on disk.")
                            else:
                                st.info("No Excel file for this backup.")
                    else:
                        st.error("Incorrect password.")

            st.divider()
            st.markdown("##### Create a Manual Backup Now")
            st.caption("Saves both a `.sql` (for pgAdmin) and `.xlsx` (for Excel) of the full database.")
            if st.button("Create Manual Backup", key="bk_manual_btn", type="secondary"):
                with st.spinner("Running backup (pg_dump + Excel export)…"):
                    path = take_backup("MANUAL", "Manual backup triggered from Backup Manager")
                if path:
                    st.success("Backup created — both .sql and .xlsx files saved.")
                    st.rerun()
                else:
                    st.error("Backup failed. Check that PostgreSQL 18 bin is accessible.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE – Analytics / Data Explorer
# ══════════════════════════════════════════════════════════════════════════════
def _excel_download_btn(df: "pd.DataFrame", filename: str, label: str = "Download Excel", key: str = "") -> None:
    """Render a download button that saves df as a styled Excel file."""
    from openpyxl.styles import PatternFill, Font, Alignment
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Data", index=False)
        ws = writer.sheets["Data"]
        gold_fill = PatternFill("solid", fgColor="C8830A")
        wfont     = Font(bold=True, color="FFFFFF")
        navy_fill = PatternFill("solid", fgColor="1B3A6B")
        sub_fill  = PatternFill("solid", fgColor="FFF3DC")
        navy_font = Font(bold=True, color="1B3A6B")
        for cell in ws[1]:
            cell.fill = gold_fill; cell.font = wfont
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max((len(str(c.value or "")) for c in col), default=10) + 4, 50)
    buf.seek(0)
    st.download_button(label, buf.read(), file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=key or filename)


def page_analytics():
    st.title("Analytics")
    st.caption("Build pivot tables and charts from live database data")
    st.divider()

    try:
        import plotly.express as px
        HAS_PLOTLY = True
    except ImportError:
        HAS_PLOTLY = False

    # ── Rich flat data (includes peripherals) ────────────────────────────────
    @st.cache_data(ttl=120)
    def load_flat():
        return db_query("""
            SELECT
                t.type_name                              AS "Asset Type",
                m.make_name                              AS "Brand",
                a.model_name                             AS "Model",
                COALESCE(a.purpose,'')                   AS "Purpose",
                COALESCE(a.status,'')                    AS "Status",
                COALESCE(d.dept_name,'Unallocated')      AS "Department",
                COALESCE(l.location_name,'')             AS "Location",
                COALESCE(a.asset_category_code,'')       AS "Short Code",
                EXTRACT(YEAR FROM p.purchase_date)::int  AS "Purchase Year",
                COALESCE(p.total_amount,0)               AS "Total Amount",
                COALESCE((SELECT per.model_name FROM peripherals per
                           WHERE per.parent_asset_id=a.asset_id
                           AND per.peripheral_type='LED Monitor' LIMIT 1),'') AS "LED Monitor",
                CASE WHEN EXISTS(SELECT 1 FROM peripherals per
                                 WHERE per.parent_asset_id=a.asset_id
                                 AND per.peripheral_type='Keyboard')
                     THEN 'Keyboard' ELSE '' END          AS "Keyboard",
                CASE WHEN EXISTS(SELECT 1 FROM peripherals per
                                 WHERE per.parent_asset_id=a.asset_id
                                 AND per.peripheral_type='Mouse')
                     THEN 'Mouse' ELSE '' END              AS "Mouse"
            FROM assets a
            JOIN asset_types t  ON a.type_id = t.type_id
            LEFT JOIN makes m        ON a.make_id  = m.make_id
            LEFT JOIN allocations al ON a.asset_id = al.asset_id AND al.is_current = TRUE
            LEFT JOIN departments d  ON al.dept_id  = d.dept_id
            LEFT JOIN locations l    ON al.location_id = l.location_id
            LEFT JOIN procurement p  ON a.procurement_id = p.procurement_id
        """)

    df_base = load_flat()
    for c in df_base.select_dtypes(include="object").columns:
        df_base[c] = df_base[c].fillna("")

    ALL_DIMS = ["Asset Type", "Brand", "Model", "Purpose", "Status",
                "Department", "Location", "Short Code", "Purchase Year",
                "LED Monitor", "Keyboard", "Mouse"]

    t_excel, t_chart, t_presets = st.tabs(["Pivot Table Builder", "Charts", "Quick Presets"])

    # ══ PIVOT TABLE BUILDER ════════════════════════════════════════════════════
    with t_excel:
        st.markdown("#### Excel-style Pivot Table")

        # ── Layout: field pickers ────────────────────────────────────────────
        lc, rc = st.columns([3, 2])
        with lc:
            row_fields = st.multiselect(
                "Row Fields  *(up to 4, order matters)*",
                ALL_DIMS, default=["Asset Type"],
                placeholder="Pick row fields…", key="xpiv_rows")
        with rc:
            remaining = ["(None)"] + [f for f in ALL_DIMS if f not in row_fields]
            col_field = st.selectbox("Column Field  *(optional)*", remaining, key="xpiv_col")
            col_field = None if col_field == "(None)" else col_field

        vc1, vc2, vc3, vc4 = st.columns(4)
        with vc1:
            val_choice = st.selectbox("Value", ["Count of Assets", "Sum of Amount"], key="xpiv_val")
        with vc2:
            if val_choice == "Count of Assets":
                agg_choice = st.selectbox("Function", ["Count"], key="xpiv_agg", disabled=True)
                agg_choice = "Count"
            else:
                agg_choice = st.selectbox("Function", ["Sum", "Average", "Max", "Min"], key="xpiv_agg")
        with vc3:
            show_sub = st.checkbox("Subtotals", value=True, key="xpiv_sub")
        with vc4:
            show_gt  = st.checkbox("Grand Total", value=True, key="xpiv_gt")

        # ── Filters ──────────────────────────────────────────────────────────
        with st.expander("Filters", expanded=False):
            df_work = df_base.copy()
            fc = st.columns(3)
            for i, fld in enumerate(["Asset Type", "Brand", "Department", "Purpose", "Status", "Location"]):
                with fc[i % 3]:
                    opts = sorted(df_work[fld].dropna().unique().tolist())
                    sel  = st.multiselect(fld, opts, default=[], key=f"xf_{fld}")
                    if sel:
                        df_work = df_work[df_work[fld].isin(sel)]
        if "df_work" not in dir():
            df_work = df_base.copy()

        if not row_fields:
            st.info("Select at least one Row field above.")
        else:
            try:
                # Prepare value column
                df_work = df_work.copy()
                if val_choice == "Count of Assets":
                    df_work["__v"] = 1
                    vcol = "__v"
                else:
                    df_work["__v"] = pd.to_numeric(df_work["Total Amount"], errors="coerce").fillna(0)
                    vcol = "__v"

                fn_map = {"Count": "sum", "Sum": "sum", "Average": "mean",
                          "Max": "max", "Min": "min"}
                pfn = fn_map[agg_choice]
                val_label = agg_choice

                # ── Cross-tab (with column field) ─────────────────────────
                if col_field:
                    pivot = pd.pivot_table(
                        df_work, values=vcol,
                        index=row_fields, columns=col_field,
                        aggfunc=pfn, fill_value=0
                    )
                    pivot.columns = [str(c) for c in pivot.columns]
                    pivot = pivot.round(0).astype(int)
                    # Reset index FIRST to avoid MultiIndex → tuple collapse
                    result_df = pivot.reset_index()
                    val_cols = [c for c in result_df.columns if c not in row_fields]
                    if show_gt:
                        result_df["Grand Total"] = result_df[val_cols].sum(axis=1)
                        gt = {f: "" for f in row_fields}
                        gt[row_fields[0]] = "Grand Total"
                        for c in val_cols:
                            gt[c] = int(result_df[c].sum())
                        gt["Grand Total"] = int(result_df["Grand Total"].sum())
                        result_df = pd.concat(
                            [result_df, pd.DataFrame([gt])], ignore_index=True)
                    row_kinds = [
                        "grandtotal" if str(r.get(row_fields[0], "")) == "Grand Total"
                        else "detail"
                        for _, r in result_df.iterrows()
                    ]

                # ── Hierarchical (no column field) ────────────────────────
                else:
                    rows_out = []

                    def _agg(sub):
                        v = sub[vcol].agg(pfn)
                        return int(round(v)) if not pd.isna(v) else 0

                    def _add(label_dict, kind):
                        rows_out.append((kind, label_dict))

                    def recurse(sub, level, path):
                        if level >= len(row_fields):
                            return
                        field = row_fields[level]
                        for val in sorted(sub[field].dropna().unique()):
                            grp  = sub[sub[field] == val]
                            here = path + [(field, val)]
                            if level == len(row_fields) - 1:
                                row = {f: "" for f in row_fields}
                                for f, v in here: row[f] = v
                                row[val_label] = _agg(grp)
                                _add(row, "detail")
                            else:
                                recurse(grp, level + 1, here)
                                if show_sub:
                                    row = {f: "" for f in row_fields}
                                    for f, v in here: row[f] = v
                                    row[field] = f"{val}  Total"
                                    row[val_label] = _agg(grp)
                                    _add(row, "subtotal")

                    recurse(df_work, 0, [])

                    if show_gt:
                        gt = {f: "" for f in row_fields}
                        gt[row_fields[0]] = "Grand Total"
                        gt[val_label] = _agg(df_work)
                        rows_out.append(("grandtotal", gt))

                    all_cols = row_fields + [val_label]
                    result_df  = pd.DataFrame([r[1] for r in rows_out], columns=all_cols).fillna("")
                    row_kinds  = [r[0] for r in rows_out]

                # ── Style & show ─────────────────────────────────────────
                def _style(row):
                    idx = row.name
                    if idx < len(row_kinds):
                        k = row_kinds[idx]
                        if k == "grandtotal":
                            return [f"background:{GOLD};color:#fff;font-weight:700"] * len(row)
                        if k == "subtotal":
                            return [f"background:#FFF3DC;color:{NAVY};font-weight:700"] * len(row)
                    return [""] * len(row)

                st.caption(f"{len(result_df):,} rows")
                st.dataframe(result_df.style.apply(_style, axis=1),
                             use_container_width=True, hide_index=True)

                _excel_download_btn(result_df, "pivot_table.xlsx",
                                    "Download Pivot as Excel", key="dl_pivot")

            except Exception as e:
                st.error(str(e))
                import traceback; st.code(traceback.format_exc())

    # ══ CHARTS ═════════════════════════════════════════════════════════════════
    with t_chart:
        st.markdown("#### Interactive Chart Builder")
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            chart_dim  = st.selectbox("Group by", ALL_DIMS, index=0, key="ch_dim")
        with cc2:
            chart_type = st.selectbox("Chart type",
                                      ["Bar", "Horizontal Bar", "Pie", "Line"], key="ch_type")
        with cc3:
            top_n = st.slider("Top N", 5, 50, 15, key="ch_topn")

        try:
            grp = df_base.groupby(chart_dim).size().reset_index(name="Count")
            grp = grp.sort_values("Count", ascending=False).head(top_n)

            if HAS_PLOTLY:
                if chart_type == "Bar":
                    fig = px.bar(grp, x=chart_dim, y="Count",
                                 color_discrete_sequence=[GOLD], text="Count")
                    fig.update_traces(textposition="outside")
                elif chart_type == "Horizontal Bar":
                    fig = px.bar(grp, x="Count", y=chart_dim, orientation="h",
                                 color_discrete_sequence=[NAVY], text="Count")
                    fig.update_traces(textposition="outside")
                    fig.update_layout(yaxis=dict(autorange="reversed"))
                elif chart_type == "Pie":
                    fig = px.pie(grp, names=chart_dim, values="Count",
                                 color_discrete_sequence=px.colors.sequential.Oranges_r)
                else:
                    fig = px.line(grp, x=chart_dim, y="Count",
                                  color_discrete_sequence=[GOLD], markers=True)
                fig.update_layout(plot_bgcolor="#fff", paper_bgcolor="#fff",
                                  margin=dict(t=20, b=20), showlegend=(chart_type == "Pie"))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.bar_chart(grp.set_index(chart_dim)["Count"])

            _excel_download_btn(grp, "chart_data.xlsx", "Download Chart Data as Excel", key="dl_chart")
        except Exception as e:
            st.error(str(e))

    # ══ QUICK PRESETS ══════════════════════════════════════════════════════════
    with t_presets:
        st.markdown("#### Quick Preset Views")

        preset = st.radio("", [
            "Assets by Type and Purpose",
            "Assets by Department and Type",
            "Assets by Brand and Type",
            "Department × Location × Type",
            "Procurement by Year",
            "Status Summary",
        ], horizontal=False, key="preset_sel")

        try:
            df_base["__n"] = 1

            if preset == "Assets by Type and Purpose":
                piv = pd.pivot_table(df_base, values="__n", index="Asset Type",
                                     columns="Purpose", aggfunc="sum", fill_value=0).astype(int)
                piv["Grand Total"] = piv.sum(axis=1)
                piv = piv.sort_values("Grand Total", ascending=False)
                st.dataframe(piv, use_container_width=True)
                _excel_download_btn(piv.reset_index(), "type_purpose.xlsx",
                                    "Download as Excel", key="dl_p1")
                if HAS_PLOTLY:
                    cols_to_plot = [c for c in piv.columns if c != "Grand Total"]
                    fig = px.bar(piv.reset_index(), x="Asset Type", y=cols_to_plot,
                                 barmode="group", color_discrete_sequence=[GOLD, NAVY, "#888"])
                    fig.update_layout(plot_bgcolor="#fff", paper_bgcolor="#fff")
                    st.plotly_chart(fig, use_container_width=True)

            elif preset == "Assets by Department and Type":
                piv = pd.pivot_table(df_base, values="__n", index="Department",
                                     columns="Asset Type", aggfunc="sum", fill_value=0).astype(int)
                piv["Grand Total"] = piv.sum(axis=1)
                st.dataframe(piv.sort_values("Grand Total", ascending=False), use_container_width=True)
                _excel_download_btn(piv.reset_index(), "dept_type.xlsx",
                                    "Download as Excel", key="dl_p2")

            elif preset == "Assets by Brand and Type":
                piv = pd.pivot_table(df_base, values="__n", index="Brand",
                                     columns="Asset Type", aggfunc="sum", fill_value=0).astype(int)
                piv["Grand Total"] = piv.sum(axis=1)
                st.dataframe(piv.sort_values("Grand Total", ascending=False), use_container_width=True)
                _excel_download_btn(piv.reset_index(), "brand_type.xlsx",
                                    "Download as Excel", key="dl_p3")

            elif preset == "Department × Location × Type":
                piv = pd.pivot_table(
                    df_base[df_base["Department"] != "Unallocated"],
                    values="__n",
                    index=["Department", "Location"],
                    columns="Asset Type",
                    aggfunc="sum", fill_value=0
                ).astype(int)
                piv["Grand Total"] = piv.sum(axis=1)
                st.dataframe(piv.sort_values("Grand Total", ascending=False), use_container_width=True)
                _excel_download_btn(piv.reset_index(), "dept_loc_type.xlsx",
                                    "Download as Excel", key="dl_p4")

            elif preset == "Procurement by Year":
                yr_df = db_query("""
                    SELECT purchase_year AS "Year",
                           COUNT(*) AS "Bills",
                           SUM(qty) AS "Total Qty",
                           SUM(total_amount)::bigint AS "Total Amount (Rs)"
                    FROM procurement
                    WHERE purchase_year IS NOT NULL
                    GROUP BY purchase_year ORDER BY purchase_year
                """)
                st.dataframe(yr_df, use_container_width=True, hide_index=True)
                _excel_download_btn(yr_df, "procurement_by_year.xlsx",
                                    "Download as Excel", key="dl_p5")
                if HAS_PLOTLY:
                    fig = px.bar(yr_df, x="Year", y="Total Amount (Rs)",
                                 color_discrete_sequence=[GOLD], text="Bills")
                    fig.update_layout(plot_bgcolor="#fff", paper_bgcolor="#fff")
                    st.plotly_chart(fig, use_container_width=True)

            elif preset == "Status Summary":
                piv = pd.pivot_table(df_base, values="__n", index="Status",
                                     columns="Asset Type", aggfunc="sum", fill_value=0).astype(int)
                piv["Grand Total"] = piv.sum(axis=1)
                st.dataframe(piv.sort_values("Grand Total", ascending=False), use_container_width=True)
                _excel_download_btn(piv.reset_index(), "status_summary.xlsx",
                                    "Download as Excel", key="dl_p6")
                if HAS_PLOTLY:
                    grp = df_base.groupby("Status").size().reset_index(name="Count")
                    fig = px.pie(grp, names="Status", values="Count",
                                 color_discrete_sequence=[GOLD, NAVY, "#aaa", "#ccc"])
                    fig.update_layout(paper_bgcolor="#fff")
                    st.plotly_chart(fig, use_container_width=True)

        except Exception as e:
            st.error(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# Main router
# ══════════════════════════════════════════════════════════════════════════════
def main():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "page" not in st.session_state:
        st.session_state.page = "dashboard"

    if not st.session_state.logged_in:
        show_login()
        return

    show_sidebar()

    page = st.session_state.get("page", "dashboard")
    if   page == "dashboard":   page_dashboard()
    elif page == "assets":      page_assets()
    elif page == "procurement":  page_procurement()
    elif page == "intercom":    page_intercom()
    elif page == "network":     page_network()
    elif page == "analytics":   page_analytics()
    elif page == "download":    page_download()


if __name__ == "__main__":
    main()

