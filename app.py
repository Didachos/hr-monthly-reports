import base64
import calendar as _calendar
import datetime
import html as _html
import io
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

import onedrive as od
from main import (
    ERGANI_CODE_TO_DESCRIPTION,
    ERGANI_LEAVE_TYPES,
    PLANNED_LEAVES_COLUMNS,
    build_alerts_report,
    build_classified_absence_template,
    build_classified_template_excel_bytes,
    build_ergani_export_df,
    build_leave_summary,
    build_validation_report,
    calculate_overtime,
    calculate_work_days,
    clean_attendance,
    find_absences,
    force_text_column,
    format_dates_for_excel,
    get_holidays,
    load_attendance,
    load_classified_absences,
    load_employees,
    load_planned_leaves,
    working_days_in_range,
)

MONTHS = {
    1: "Ιανουάριος", 2: "Φεβρουάριος", 3: "Μάρτιος",
    4: "Απρίλιος", 5: "Μάιος", 6: "Ιούνιος",
    7: "Ιούλιος", 8: "Αύγουστος", 9: "Σεπτέμβριος",
    10: "Οκτώβριος", 11: "Νοέμβριος", 12: "Δεκέμβριος",
}

# Ελληνικά ονόματα ημερών (Δευτέρα=0) και μηνών σε γενική (για "3 Οκτωβρίου")
GREEK_DAYS = ["Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή", "Σάββατο", "Κυριακή"]
GREEK_MONTHS_GEN = {
    1: "Ιανουαρίου", 2: "Φεβρουαρίου", 3: "Μαρτίου", 4: "Απριλίου",
    5: "Μαΐου", 6: "Ιουνίου", 7: "Ιουλίου", 8: "Αυγούστου",
    9: "Σεπτεμβρίου", 10: "Οκτωβρίου", 11: "Νοεμβρίου", 12: "Δεκεμβρίου",
}

# Εικονίδια ανά τύπο άδειας (για πιο ευανάγνωστη λίστα)
LEAVE_TYPE_ICON = {
    "Κανονική άδεια": "🏖️",
    "Άδεια ασθένειας (ανυπαίτιο κώλυμα παροχής εργασίας)": "🤒",
    "Άδεια ασθενείας": "🤒",
    "Άδεια άνευ αποδοχών": "💸",
    "Άνευ αποδοχών άδεια": "💸",
}


def format_greek_date(ts) -> str:
    """Μετατρέπει ημερομηνία σε 'Δευτέρα 3 Οκτωβρίου 2026'."""
    ts = pd.to_datetime(ts)
    return f"{GREEK_DAYS[ts.weekday()]} {ts.day} {GREEK_MONTHS_GEN[ts.month]} {ts.year}"

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "data/output"


def save_upload_to_temp(uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    return Path(tmp.name)


def excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            format_dates_for_excel(df).to_excel(writer, sheet_name=sheet_name, index=False)
            force_text_column(writer.sheets[sheet_name], "ΑΦΜ")
    return buf.getvalue()


# =========================
# PLANNED LEAVES HELPERS
# =========================

PLANNED_FILENAME = "planned_leaves.xlsx"
LOCAL_CONFIG_DIR = Path(__file__).resolve().parent / "data/config"


def get_employees_df():
    """Επιστρέφει το employees DataFrame από OneDrive bytes (ή None)."""
    emp_bytes = st.session_state.get("employees_od_bytes")
    od_token = st.session_state.get("od_token")
    if emp_bytes is None and od_token:
        try:
            cfg = od.list_files(od_token, subfolder="config")
            if any(f["name"] == "employees.xlsx" for f in cfg):
                emp_bytes = od.download_file(od_token, "employees.xlsx", subfolder="config")
                st.session_state["employees_od_bytes"] = emp_bytes
        except Exception:
            pass
    if not emp_bytes:
        return None
    try:
        t = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        t.write(emp_bytes); t.close()
        return load_employees(Path(t.name))
    except Exception:
        return None


def load_planned_df(force: bool = False):
    """Φορτώνει planned leaves από OneDrive (config) ή τοπικά, με cache."""
    if not force and "planned_df" in st.session_state:
        return st.session_state["planned_df"]
    od_token = st.session_state.get("od_token")
    df = pd.DataFrame(columns=PLANNED_LEAVES_COLUMNS)
    try:
        if od_token:
            files = od.list_files(od_token, subfolder="config")
            if any(f["name"] == PLANNED_FILENAME for f in files):
                b = od.download_file(od_token, PLANNED_FILENAME, subfolder="config")
                t = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                t.write(b); t.close()
                df = load_planned_leaves(Path(t.name))
        else:
            df = load_planned_leaves(LOCAL_CONFIG_DIR / PLANNED_FILENAME)
    except Exception:
        pass
    st.session_state["planned_df"] = df
    return df


def save_planned_df(df: pd.DataFrame):
    """Αποθηκεύει planned leaves (OneDrive config ή τοπικά) και ενημερώνει cache."""
    df = df.copy()
    for col in PLANNED_LEAVES_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
    df = df[PLANNED_LEAVES_COLUMNS]
    data = excel_bytes({"Προγραμματισμένες": df})
    od_token = st.session_state.get("od_token")
    if od_token:
        od.upload_file(od_token, PLANNED_FILENAME, data, subfolder="config")
    else:
        LOCAL_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        (LOCAL_CONFIG_DIR / PLANNED_FILENAME).write_bytes(data)
    st.session_state["planned_df"] = df


# Παλέτα χρωμάτων ανά υποκατάστημα (λειτουργεί σε light & dark)
BRANCH_COLORS = ["#3b82f6", "#22c55e", "#ef4444", "#a855f7",
                 "#f59e0b", "#06b6d4", "#ec4899", "#84cc16"]


def _branch_color(branch) -> str:
    try:
        return BRANCH_COLORS[int(branch) % len(BRANCH_COLORS)]
    except (TypeError, ValueError):
        return "#9ca3af"


def build_calendar_html(planned_df: pd.DataFrame, year: int, month: int,
                        today: datetime.date | None = None) -> str:
    """Φτιάχνει HTML ημερολόγιο μήνα με τις προγραμματισμένες απουσίες ανά ημέρα."""
    today = today or datetime.date.today()
    holidays_set = get_holidays(year)

    # Χάρτης ημέρα -> λίστα (name, branch, type, status)
    by_day: dict[int, list] = {}
    if planned_df is not None and not planned_df.empty:
        _p = planned_df.copy()
        _p["Ημ/νία"] = pd.to_datetime(_p["Ημ/νία"], errors="coerce")
        _p = _p[(_p["Ημ/νία"].dt.year == year) & (_p["Ημ/νία"].dt.month == month)]
        for _, r in _p.iterrows():
            by_day.setdefault(int(r["Ημ/νία"].day), []).append({
                "name": f"{r['Επώνυμο']} {r['Όνομα']}",
                "branch": r["ΑΑ Παραρτηματος"],
                "type": r.get("Τύπος Απουσίας", ""),
                "status": r.get("Κατάσταση", ""),
            })

    css = """
    <style>
    .cal { width:100%; border-collapse:collapse; table-layout:fixed; font-size:0.82rem; }
    .cal th { padding:6px 4px; text-align:center; font-weight:600; opacity:0.7; }
    .cal td { border:1px solid rgba(128,128,128,0.25); vertical-align:top;
              height:96px; padding:4px; overflow:hidden; }
    .cal .daynum { font-weight:700; opacity:0.65; font-size:0.78rem; }
    .cal .off { display:block; background:rgba(128,128,128,0.10); }
    .cal .today { outline:2px solid #3b82f6; outline-offset:-2px; }
    .cal .person { display:flex; align-items:center; gap:4px; margin-top:2px;
                   white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .cal .dot { width:8px; height:8px; border-radius:50%; flex:0 0 auto; }
    .cal .pend { opacity:0.55; font-style:italic; }
    </style>
    """

    head = "".join(f"<th>{d}</th>" for d in ["Δε", "Τρ", "Τε", "Πε", "Πα", "Σα", "Κυ"])
    rows_html = ""
    for week in _calendar.monthcalendar(year, month):
        cells = ""
        for i, day in enumerate(week):
            if day == 0:
                cells += "<td></td>"
                continue
            d = datetime.date(year, month, day)
            is_weekend = i >= 5
            is_holiday = pd.Timestamp(d).normalize() in holidays_set
            is_today = d == today
            people = by_day.get(day, [])

            cls = "off" if people else ""
            if is_today:
                cls += " today"
            cell_style = "opacity:0.55;" if (is_weekend or is_holiday) else ""

            inner = f'<span class="daynum">{day}</span>'
            shown = people[:4]
            for p in shown:
                col = _branch_color(p["branch"])
                pend = " pend" if p["status"] != "εγκρίθηκε" else ""
                nm = _html.escape(p["name"])
                inner += (f'<span class="person{pend}" title="{nm} · {_html.escape(str(p["type"]))}">'
                          f'<span class="dot" style="background:{col}"></span>{nm}</span>')
            if len(people) > 4:
                inner += f'<span class="person">+{len(people) - 4} ακόμη</span>'

            cells += f'<td class="{cls.strip()}" style="{cell_style}">{inner}</td>'
        rows_html += f"<tr>{cells}</tr>"

    return css + f'<table class="cal"><thead><tr>{head}</tr></thead><tbody>{rows_html}</tbody></table>'


def reconstruct_classified_from_ergani(
    ergani_files: list[tuple[int, bytes]]
) -> pd.DataFrame:
    """
    Ανακατασκευάζει classified DataFrame από Ergani export αρχεία.

    ergani_files: list of (branch_aa, excel_bytes)
        Το branch_aa είναι ο αριθμός παραρτήματος (ΑΑ Παραρτηματος),
        εξαγόμενος από το όνομα αρχείου ergani_export_parartima_{aa}_{year}_{month}.xlsx
    """
    _CODE_TO_DESC = ERGANI_CODE_TO_DESCRIPTION

    frames = []
    for aa, raw_bytes in ergani_files:
        try:
            xf = pd.read_excel(io.BytesIO(raw_bytes), dtype={"ΑΦΜ": str})
            xf.columns = [str(c).strip() for c in xf.columns]

            if "ΤΥΠΟΣ" not in xf.columns or "ΗΜΕΡΑ" not in xf.columns:
                continue

            xf["Τύπος Απουσίας"] = xf["ΤΥΠΟΣ"].map(_CODE_TO_DESC)
            xf["Ημ/νία"] = pd.to_datetime(xf["ΗΜΕΡΑ"], errors="coerce").dt.normalize()
            xf["Έτος Άδειας"] = pd.to_numeric(xf.get("ΕΤΟΣ ΑΝΑΦΟΡΑΣ"), errors="coerce")
            xf["ΑΑ Παραρτηματος"] = aa
            xf["ΑΦΜ"] = xf["ΑΦΜ"].astype(str).str.strip()
            xf["Επώνυμο"] = xf.get("ΕΠΩΝΥΜΟ", "").astype(str).str.strip()
            xf["Όνομα"] = xf.get("ΟΝΟΜΑ", "").astype(str).str.strip()

            # Για μη-κανονική άδεια, Έτος Άδειας = NaN
            xf.loc[xf["Τύπος Απουσίας"] != "Κανονική άδεια", "Έτος Άδειας"] = pd.NA

            cls = xf[["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία",
                       "Τύπος Απουσίας", "Έτος Άδειας"]].copy()
            cls = cls.dropna(subset=["ΑΦΜ", "Ημ/νία", "Τύπος Απουσίας"])
            frames.append(cls)
        except Exception:
            continue

    if not frames:
        return pd.DataFrame()

    result = pd.concat(frames, ignore_index=True).drop_duplicates().reset_index(drop=True)
    result["ΑΦΜ"] = result["ΑΦΜ"].astype(str)
    return result


def ergani_excel_bytes(df: pd.DataFrame, sheet_name: str = "DAILY") -> bytes:
    """Γράφει το Ergani export ακριβώς όπως το πρότυπο EXCEL_PROTOTYPE_DAILY_LEAVES."""
    df = df.copy()

    # ΩΡΑ ΑΠΌ - ΩΡΑ ΕΩΣ: κενό string → None (αληθινό άδειο κελί, dtype='n' όπως πρότυπο)
    if "ΩΡΑ ΑΠΌ - ΩΡΑ ΕΩΣ" in df.columns:
        df["ΩΡΑ ΑΠΌ - ΩΡΑ ΕΩΣ"] = df["ΩΡΑ ΑΠΌ - ΩΡΑ ΕΩΣ"].replace("", None)

    # ΕΤΟΣ ΑΝΑΦΟΡΑΣ / ΔΙΚ. ΗΜΕΡΕΣ: χρήση float (όχι Int64) ώστε τα None να γράφονται
    # ως αληθινά άδεια κελιά (dtype='n'), όχι ως inlineStr
    for col in ["ΕΤΟΣ ΑΝΑΦΟΡΑΣ", "ΔΙΚ. ΗΜΕΡΕΣ"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")  # NaN = αληθινό κενό

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        force_text_column(ws, "ΑΦΜ")
        force_text_column(ws, "ΩΡΑ ΑΠΌ - ΩΡΑ ΕΩΣ")  # fmt='@' όπως πρότυπο

        # Headers ΑΦΜ + ΩΡΑ: fmt='@' και στο header (όπως πρότυπο)
        for cell in ws[1]:
            if cell.value in ("ΑΦΜ", "ΩΡΑ ΑΠΌ - ΩΡΑ ΕΩΣ"):
                cell.number_format = "@"

        # ΗΜΕΡΑ: format mm-dd-yy (ακριβώς όπως πρότυπο)
        for cell in ws[1]:
            if cell.value == "ΗΜΕΡΑ":
                col_idx = cell.column
                for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    c = row[0]
                    if c.value is not None:
                        c.number_format = "mm-dd-yy"
                break

        # ΕΤΟΣ ΑΝΑΦΟΡΑΣ / ΔΙΚ. ΗΜΕΡΕΣ: float 2025.0 → int 2025 στα κελιά
        for header_name in ["ΕΤΟΣ ΑΝΑΦΟΡΑΣ", "ΔΙΚ. ΗΜΕΡΕΣ"]:
            for cell in ws[1]:
                if cell.value == header_name:
                    col_idx = cell.column
                    for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                        c = row[0]
                        try:
                            if c.value is not None:
                                c.value = int(float(c.value))
                        except (TypeError, ValueError):
                            pass
                    break

        # Καθαρισμός κενών strings → αληθινά άδεια κελιά (dtype='n' όπως πρότυπο)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if c.value == "":
                    c.value = None

    return buf.getvalue()


def _branch_val(raw):
    """Επιστρέφει int αν είναι αριθμός, αλλιώς pd.NA (ώστε sort_values να δουλεύει)."""
    try:
        v = float(raw)
        return int(v) if not pd.isna(v) else pd.NA
    except (TypeError, ValueError):
        return pd.NA


def leave_balance_table_current(leaves: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in leaves.iterrows():
        curr_taken = int(r["Κανονική Άδεια από Τρέχον Έτος"])
        curr_total = int(r["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"])
        balance = int(r["Υπόλοιπο Τρέχοντος Έτους Μετά"])
        rows.append({
            "Υποκατ.": _branch_val(r.get("ΑΑ Παραρτηματος")),
            "ΑΦΜ": str(r["ΑΦΜ"]),
            "Επώνυμο": r["Επώνυμο"],
            "Όνομα": r["Όνομα"],
            "Δικαιούμενες": curr_total,
            "Ληφθείσες": curr_taken,
            "Υπόλοιπο": balance,
        })
    df = pd.DataFrame(rows)
    df["Υποκατ."] = pd.array(df["Υποκατ."], dtype="Int64")
    df = df.sort_values("Υποκατ.", na_position="last", kind="stable").reset_index(drop=True)
    return df


def leave_balance_table_prev(leaves: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in leaves.iterrows():
        prev_taken = int(r["Κανονική Άδεια από Προηγούμενο Έτος"])
        prev_available = int(r["Υπόλοιπο Προηγούμενου Έτους"])
        balance = int(r["Υπόλοιπο Προηγούμενου Έτους Μετά"])
        rows.append({
            "Υποκατ.": _branch_val(r.get("ΑΑ Παραρτηματος")),
            "ΑΦΜ": str(r["ΑΦΜ"]),
            "Επώνυμο": r["Επώνυμο"],
            "Όνομα": r["Όνομα"],
            "Διαθέσιμες": prev_available,
            "Ληφθείσες": prev_taken,
            "Υπόλοιπο": balance,
        })
    df = pd.DataFrame(rows)
    df["Υποκατ."] = pd.array(df["Υποκατ."], dtype="Int64")
    df = df.sort_values("Υποκατ.", na_position="last", kind="stable").reset_index(drop=True)
    return df


# =========================
# ONEDRIVE AUTH
# =========================

def init_onedrive():
    """Αρχικοποιεί OneDrive auth από secrets. Επιστρέφει token ή None."""
    try:
        cfg = st.secrets["onedrive"]
        client_id = cfg["client_id"]
        tenant_id = cfg["tenant_id"]
        token_cache_raw = cfg.get("token_cache", "")

        # Αποκωδικοποίηση base64 αν χρειάζεται
        if token_cache_raw:
            try:
                token_cache_str = base64.b64decode(token_cache_raw.strip()).decode("utf-8")
            except Exception:
                token_cache_str = token_cache_raw  # fallback: χρησιμοποίησε raw
        else:
            token_cache_str = ""

        # Debug: τι διαβάζουμε από τα secrets
        st.session_state["od_debug_cache_len"] = len(token_cache_str) if token_cache_str else 0

        app, cache = od.build_app(client_id, tenant_id, token_cache_str or None)

        accounts = app.get_accounts()

        if accounts:
            token, _ = od.get_token_silent(app, cache)
            if token:
                st.session_state["od_token"] = token
                st.session_state["od_app"] = app
                st.session_state["od_cache"] = cache
                return token

        # Δεν υπάρχει token — ξεκίνα device flow μια φορά
        if "od_flow" not in st.session_state:
            st.session_state["od_app"] = app
            st.session_state["od_cache"] = cache
            flow = od.start_device_flow(app)
            st.session_state["od_flow"] = flow

        return None
    except Exception as e:
        st.session_state["od_init_error"] = str(e)
        return None


# =========================
# UI
# =========================

st.set_page_config(page_title="Εργάνη - Απουσίες", page_icon="📋", layout="wide")

# =========================
# PASSWORD PROTECTION
# =========================

def check_password() -> bool:
    try:
        correct = st.secrets["app"]["password"]
    except Exception:
        return True  # Αν δεν υπάρχει password στα secrets, επέτρεψε πρόσβαση

    if st.session_state.get("authenticated"):
        return True

    with st.form("login"):
        st.subheader("🔐 Σύνδεση")
        pwd = st.text_input("Κωδικός", type="password")
        submitted = st.form_submit_button("Είσοδος")
        if submitted:
            if pwd == correct:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Λάθος κωδικός.")
    return False

if not check_password():
    st.stop()

st.title("📋 Εργάνη — Διαχείριση Απουσιών")

# --- OneDrive sidebar ---
with st.sidebar:
    st.subheader("☁️ OneDrive")
    od_token = st.session_state.get("od_token") or init_onedrive()

    if od_token:
        st.success("Συνδεδεμένο ✅")

        # Ανανέωση token cache (για χρήση όταν πλησιάζει λήξη ~90 μέρες)
        if st.button("🔄 Ανανέωση σύνδεσης"):
            st.session_state.pop("od_token", None)
            st.session_state.pop("od_flow", None)
            st.session_state.pop("od_app", None)
            st.session_state.pop("od_cache", None)
            st.rerun()

        # Αν μόλις συνδέθηκε, δείξε το token_cache για αποθήκευση στα secrets
        new_cache_str = st.session_state.get("od_new_cache_str")
        if new_cache_str:
            st.info("📋 Αντέγραψε το παρακάτω και πρόσθεσέ το στα Streamlit Secrets ως `token_cache`:")
            st.code(new_cache_str)
            st.caption("Μετά από αυτό η σύνδεση θα γίνεται αυτόματα κάθε φορά.")
            if st.button("✔ Το αντέγραψα"):
                del st.session_state["od_new_cache_str"]
                st.rerun()
    else:
        init_err = st.session_state.get("od_init_error")
        if init_err:
            st.error(f"Σφάλμα αρχικοποίησης: {init_err}")
        flow = st.session_state.get("od_flow")
        if flow:
            if "error" in flow:
                st.error(f"Σφάλμα Azure: {flow.get('error_description', flow.get('error'))}")
                st.caption("Βεβαιώσου ότι το Azure app έχει Files.ReadWrite permission και έχεις κάνει 'Grant admin consent'.")
            elif "user_code" in flow:
                st.warning("Απαιτείται σύνδεση")
                st.markdown("1. Πήγαινε στο [microsoft.com/devicelogin](https://microsoft.com/devicelogin)")
                st.code(flow["user_code"], language=None)
                st.caption("Εισήγαγε τον κωδικό παραπάνω και συνδέσου με τον Microsoft λογαριασμό σου.")
                if st.button("✅ Έγινε σύνδεση"):
                    with st.spinner("Αναμονή επιβεβαίωσης από Microsoft..."):
                        result = od.complete_device_flow(
                            st.session_state["od_app"],
                            flow,
                        )
                    if "access_token" in result:
                        st.session_state["od_token"] = result["access_token"]
                        app_obj = st.session_state["od_app"]
                        cache_json = od.get_cache_str(app_obj)
                        # Κωδικοποίηση σε base64 για ασφαλή αποθήκευση στο TOML
                        cache_b64 = base64.b64encode(cache_json.encode("utf-8")).decode("ascii")
                        accounts_now = app_obj.get_accounts()
                        st.session_state["od_new_cache_str"] = cache_b64
                        st.session_state["od_debug_cache_accounts"] = len(accounts_now)
                        st.rerun()
                    else:
                        err = result.get("error_description") or result.get("error") or str(result)
                        st.error(f"Αποτυχία σύνδεσης: {err}")
                        st.caption("Δοκίμασε να ανανεώσεις τη σελίδα για νέο κωδικό.")
        else:
            st.info("Δεν έχουν οριστεί OneDrive credentials.")

tab_dashboard, tab_run, tab_planned, tab_history, tab_balances = st.tabs(
    ["🏠 Επισκόπηση", "▶ Εκτέλεση", "📆 Προγραμματισμένες Άδειες", "📁 Ιστορικό", "📊 Υπόλοιπα Αδειών"]
)


# =========================
# TAB: ΕΠΙΣΚΟΠΗΣΗ (DASHBOARD)
# =========================

with tab_dashboard:
    _dash_today = datetime.date.today()
    _gr_full_date = format_greek_date(pd.Timestamp(_dash_today))
    st.subheader(f"🏠 Επισκόπηση — {_gr_full_date}")

    _dash_emp = get_employees_df()
    _dash_planned = load_planned_df()
    _dash_leaves = st.session_state.get("leaves")

    # Ενεργοί υπάλληλοι (χωρίς αποχώρηση)
    _dash_active_emp = None
    if _dash_emp is not None and not _dash_emp.empty:
        _dep = pd.to_datetime(_dash_emp.get("Ημερομηνία Αποχώρησης"), errors="coerce")
        _dash_active_emp = _dash_emp[_dep.isna() | (_dep.dt.date > _dash_today)]

    # Εβδομάδα (Δευτ–Κυρ)
    _week_start = _dash_today - datetime.timedelta(days=_dash_today.weekday())
    _week_end = _week_start + datetime.timedelta(days=6)

    # --- Προγραμματισμένες άδειες: σήμερα / εβδομάδα ---
    _off_today = pd.DataFrame()
    _off_week = pd.DataFrame()
    if _dash_planned is not None and not _dash_planned.empty:
        _p = _dash_planned.copy()
        _p["Ημ/νία"] = pd.to_datetime(_p["Ημ/νία"], errors="coerce")
        _off_today = _p[_p["Ημ/νία"].dt.date == _dash_today]
        _off_week = _p[
            (_p["Ημ/νία"].dt.date >= _week_start) &
            (_p["Ημ/νία"].dt.date <= _week_end)
        ]

    # --- Χαμηλό υπόλοιπο (ενεργοί) ---
    _low_balance = pd.DataFrame()
    if _dash_leaves is not None and not _dash_leaves.empty:
        _lb = _dash_leaves.copy()
        if "Ημερομηνία Αποχώρησης" in _lb.columns:
            _lbdep = pd.to_datetime(_lb["Ημερομηνία Αποχώρησης"], dayfirst=True, errors="coerce")
            _lb = _lb[_lbdep.isna() | (_lbdep.dt.date > _dash_today)]
        if "Υπόλοιπο Τρέχοντος Έτους Μετά" in _lb.columns:
            _lb["_bal"] = pd.to_numeric(_lb["Υπόλοιπο Τρέχοντος Έτους Μετά"], errors="coerce").fillna(0)
            _low_balance = _lb[_lb["_bal"] <= 3]

    # --- Εκκρεμείς εγκρίσεις ---
    _pending = pd.DataFrame()
    if _dash_planned is not None and not _dash_planned.empty:
        _pending = _dash_planned[_dash_planned["Κατάσταση"] == "εκκρεμεί"]

    # --- Metrics row ---
    _m1, _m2, _m3, _m4 = st.columns(4)
    _m1.metric("👥 Ενεργοί υπάλληλοι",
               len(_dash_active_emp) if _dash_active_emp is not None else "—")
    _m2.metric("🌴 Λείπουν σήμερα", _off_today["ΑΦΜ"].nunique() if not _off_today.empty else 0)
    _m3.metric("📅 Λείπουν αυτή την εβδομάδα", _off_week["ΑΦΜ"].nunique() if not _off_week.empty else 0)
    _m4.metric("⚠️ Χαμηλό υπόλοιπο (≤3)", len(_low_balance) if not _low_balance.empty else 0)

    st.divider()

    def _dash_person_line(row):
        _br = row.get("ΑΑ Παραρτηματος")
        _br_s = f"[Υποκ. {int(_br)}] " if pd.notna(_br) else ""
        _icon = LEAVE_TYPE_ICON.get(row.get("Τύπος Απουσίας", ""), "📌")
        _st = row.get("Κατάσταση", "")
        _st_s = "" if _st == "εγκρίθηκε" else f" *({_st})*"
        return f"- {_br_s}**{row['Επώνυμο']} {row['Όνομα']}** · {_icon} {row.get('Τύπος Απουσίας','')}{_st_s}"

    _dc1, _dc2 = st.columns(2)

    with _dc1:
        st.markdown("### 🌴 Λείπουν σήμερα")
        if _off_today.empty:
            st.caption("Κανείς δεν λείπει σήμερα. 🎉")
        else:
            _rows = _off_today.sort_values(["ΑΑ Παραρτηματος", "Επώνυμο"])
            st.markdown("\n".join(_dash_person_line(r) for _, r in _rows.iterrows()))

    with _dc2:
        st.markdown("### 📅 Αυτή την εβδομάδα")
        if _off_week.empty:
            st.caption("Καμία προγραμματισμένη άδεια αυτή την εβδομάδα.")
        else:
            _wk = _off_week.copy()
            _wk["Ημ/νία"] = pd.to_datetime(_wk["Ημ/νία"], errors="coerce")
            for _day, _grp in _wk.groupby(_wk["Ημ/νία"].dt.date):
                _names = ", ".join(
                    f"{r['Επώνυμο']} {r['Όνομα']}"
                    for _, r in _grp.sort_values(["ΑΑ Παραρτηματος", "Επώνυμο"]).iterrows()
                )
                st.markdown(f"**{format_greek_date(pd.Timestamp(_day))}** — {_names}")

    st.divider()

    _dc3, _dc4 = st.columns(2)

    with _dc3:
        st.markdown("### ⚠️ Χαμηλό υπόλοιπο")
        if _low_balance.empty:
            if _dash_leaves is None:
                st.caption("Άνοιξε την καρτέλα «Υπόλοιπα Αδειών» για να φορτωθούν τα δεδομένα.")
            else:
                st.caption("Κανείς με χαμηλό υπόλοιπο. ✅")
        else:
            _lbv = _low_balance.sort_values("_bal")
            for _, r in _lbv.iterrows():
                _br = r.get("ΑΑ Παραρτηματος")
                _br_s = f"[Υποκ. {int(_br)}] " if pd.notna(_br) else ""
                _bal = int(r["_bal"])
                _color = "🔴" if _bal <= 1 else "🟠"
                st.markdown(f"- {_color} {_br_s}**{r['Επώνυμο']} {r['Όνομα']}** · υπόλοιπο **{_bal}** ημέρες")

    with _dc4:
        st.markdown("### ✅ Εκκρεμείς εγκρίσεις")
        if _pending.empty:
            st.caption("Καμία εκκρεμής έγκριση.")
        else:
            _pv = _pending.copy()
            _pv["Ημ/νία"] = pd.to_datetime(_pv["Ημ/νία"], errors="coerce")
            _grouped = (
                _pv.groupby(["ΑΦΜ", "Επώνυμο", "Όνομα", "Τύπος Απουσίας"])
                .agg(_from=("Ημ/νία", "min"), _to=("Ημ/νία", "max"), _n=("Ημ/νία", "size"))
                .reset_index()
            )
            for _, r in _grouped.iterrows():
                _rng = (format_greek_date(r["_from"]) if r["_n"] == 1
                        else f"{format_greek_date(r['_from'])} → {format_greek_date(r['_to'])}")
                st.markdown(f"- **{r['Επώνυμο']} {r['Όνομα']}** · {r['Τύπος Απουσίας']} · {_rng} ({int(r['_n'])} ημ.)")
            st.caption("Η έγκριση γίνεται από την καρτέλα «Προγραμματισμένες Άδειες».")


# =========================
# TAB: ΕΚΤΕΛΕΣΗ
# =========================

with tab_run:
    st.subheader("Περίοδος")
    _today = datetime.date.today()
    # Default: προηγούμενος μήνας (γιατί συνήθως επεξεργαζόμαστε τον περασμένο μήνα)
    _default_month = _today.month - 1 if _today.month > 1 else 12
    _default_year = _today.year if _today.month > 1 else _today.year - 1

    col1, col2 = st.columns(2)
    with col1:
        year = st.number_input("Έτος", min_value=2020, max_value=2100, value=_default_year, step=1)
    with col2:
        month = st.selectbox("Μήνας", options=list(MONTHS.keys()), format_func=lambda m: MONTHS[m], index=_default_month - 1)

    st.subheader("Αρχεία Εισόδου")

    # Raw attendance: upload ή αυτόματη φόρτωση από OneDrive
    _raw_od_token = st.session_state.get("od_token")
    _raw_filename = f"raw_attendance_{year}_{month:02d}.xlsx"
    _raw_cache_key = f"raw_od_{year}_{month:02d}"

    # Καθαρισμός cache αν άλλαξε μήνας/έτος
    if st.session_state.get("_raw_cache_key") != _raw_cache_key:
        st.session_state.pop("raw_od_bytes", None)
        st.session_state["_raw_cache_key"] = _raw_cache_key

    # Αυτόματη φόρτωση από OneDrive αν δεν έχει ήδη φορτωθεί
    if _raw_od_token and "raw_od_bytes" not in st.session_state:
        try:
            _raw_od_files = od.list_files(_raw_od_token, subfolder="raw")
            if any(f["name"] == _raw_filename for f in _raw_od_files):
                st.session_state["raw_od_bytes"] = od.download_file(
                    _raw_od_token, _raw_filename, subfolder="raw"
                )
        except Exception:
            pass

    raw_file = st.file_uploader("Αρχείο παρουσίας (.xlsx)", type=["xlsx"])

    # Αν δεν ανέβηκε αρχείο, χρησιμοποίησε το cached από OneDrive
    if not raw_file and st.session_state.get("raw_od_bytes"):
        class _FakeFile:
            def __init__(self, data, name):
                self._data = data
                self.name = name
            def getvalue(self):
                return self._data
            def read(self):
                return self._data
        raw_file = _FakeFile(st.session_state["raw_od_bytes"], _raw_filename)
        st.success(f"☁️ Χρησιμοποιείται το `{_raw_filename}` από OneDrive.")

    # --- employees.xlsx: αυτόματη φόρτωση από OneDrive (subfolder "config") ---
    _od_emp_token = st.session_state.get("od_token")
    _emp_od_bytes = None
    if _od_emp_token:
        try:
            if "employees_od_bytes" not in st.session_state:
                _emp_files = od.list_files(_od_emp_token, subfolder="config")
                if any(f["name"] == "employees.xlsx" for f in _emp_files):
                    st.session_state["employees_od_bytes"] = od.download_file(
                        _od_emp_token, "employees.xlsx", subfolder="config"
                    )
            _emp_od_bytes = st.session_state.get("employees_od_bytes")
        except Exception:
            pass

    if _emp_od_bytes:
        st.success("✅ employees.xlsx φορτώθηκε αυτόματα από OneDrive")
        _emp_cols = st.columns([3, 1])
        with _emp_cols[0]:
            employees_file = st.file_uploader(
                "Αντικατάσταση employees.xlsx (προαιρετικό)", type=["xlsx"], key="emp_override"
            )
        with _emp_cols[1]:
            st.write("")
            st.write("")
            if employees_file and st.button("⬆ Αποθήκευση στο OneDrive", key="emp_save_btn"):
                od.upload_file(_od_emp_token, "employees.xlsx", employees_file.getvalue(), subfolder="config")
                st.session_state["employees_od_bytes"] = employees_file.getvalue()
                st.success("Αποθηκεύτηκε!")
        _employees_ready = True
    else:
        employees_file = st.file_uploader("employees.xlsx", type=["xlsx"])
        if employees_file and _od_emp_token:
            if st.button("⬆ Αποθήκευση στο OneDrive", key="emp_save_btn"):
                od.upload_file(_od_emp_token, "employees.xlsx", employees_file.getvalue(), subfolder="config")
                st.session_state["employees_od_bytes"] = employees_file.getvalue()
                st.success("Αποθηκεύτηκε στο OneDrive!")
        _employees_ready = bool(employees_file)

    st.subheader("Ταξινόμηση Απουσιών (προαιρετικό)")
    st.caption("Αν έχεις ήδη συμπληρώσει το classified_absences, ανέβασέ το εδώ για να παραχθεί το πλήρες report.")

    classified_file = None
    classified_bytes = None

    od_token_cls = st.session_state.get("od_token")
    cls_filename = f"classified_absences_{year}_{month:02d}.xlsx"

    # Καθαρισμός cache αν άλλαξε μήνας/έτος
    _cls_cache_key = f"classified_od_{year}_{month:02d}"
    if st.session_state.get("_cls_cache_key") != _cls_cache_key:
        st.session_state.pop("classified_od_bytes", None)
        st.session_state["_cls_cache_key"] = _cls_cache_key

    # Αυτόματη φόρτωση classified από OneDrive (χωρίς κουμπί)
    if od_token_cls and "classified_od_bytes" not in st.session_state:
        try:
            od_files = od.list_files(od_token_cls, subfolder="output")
            if any(f["name"] == cls_filename for f in od_files):
                st.session_state["classified_od_bytes"] = od.download_file(
                    od_token_cls, cls_filename, subfolder="output"
                )
        except Exception:
            pass

    classified_file = st.file_uploader(
        f"Classified absences — {cls_filename} (ή ανέβασε νέο)",
        type=["xlsx"], key="cls_upload"
    )

    if not classified_file:
        classified_bytes = st.session_state.get("classified_od_bytes")
        if classified_bytes:
            st.success(f"☁️ Φορτώθηκε αυτόματα: `{cls_filename}`")

    st.divider()

    # Έλεγχος αν υπάρχει ήδη report για τον μήνα
    od_token_check = st.session_state.get("od_token")
    if od_token_check:
        try:
            existing = od.list_files(od_token_check, subfolder="output")
            exists = any(f["name"] == f"monthly_report_{year}_{month:02d}.xlsx" for f in existing)
            if exists:
                st.warning(f"⚠️ Υπάρχει ήδη report για {MONTHS[month]} {year} στο OneDrive. Αν συνεχίσεις θα αντικατασταθεί.")
        except Exception:
            pass

    run = st.button("▶ Εκτέλεση", type="primary", disabled=not (raw_file and _employees_ready))

    if run:
        try:
            with st.spinner("Επεξεργασία..."):
                raw_path = save_upload_to_temp(raw_file)
                # employees: από OneDrive bytes ή από upload
                _emp_bytes = employees_file.getvalue() if employees_file else _emp_od_bytes
                _emp_tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                _emp_tmp.write(_emp_bytes)
                _emp_tmp.close()
                emp_path = Path(_emp_tmp.name)

                raw_df = load_attendance(raw_path)
                df = clean_attendance(raw_df)
                employees = load_employees(emp_path)

                absences = find_absences(df, employees, year, month)
                workdays = calculate_work_days(df, year, month)
                overtime_d, overtime_s = calculate_overtime(df.copy(), year, month)

                if classified_bytes:
                    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    tmp.write(classified_bytes)
                    tmp.close()
                    classified = load_classified_absences(Path(tmp.name))
                elif classified_file:
                    cls_path = save_upload_to_temp(classified_file)
                    classified = load_classified_absences(cls_path)
                else:
                    classified = pd.DataFrame()

                # ── Year-to-date classified: συνδυάζει ΟΛΟΥΣ τους μήνες του έτους ──
                # Χρησιμοποιείται μόνο για το leave_summary ώστε τα υπόλοιπα να είναι σωστά
                classified_ytd = classified.copy()
                _od_ytd = st.session_state.get("od_token")
                if _od_ytd and month > 1:
                    try:
                        _all_od = od.list_files(_od_ytd, subfolder="output")
                        for _m in range(1, month):
                            _prev_name = f"classified_absences_{year}_{_m:02d}.xlsx"
                            if any(f["name"] == _prev_name for f in _all_od):
                                _pb = od.download_file(_od_ytd, _prev_name, subfolder="output")
                                _pt = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                                _pt.write(_pb); _pt.close()
                                _prev_cls = load_classified_absences(Path(_pt.name))
                                if not _prev_cls.empty:
                                    classified_ytd = pd.concat([_prev_cls, classified_ytd], ignore_index=True)
                    except Exception:
                        pass  # fallback: μόνο ο τρέχων μήνας

                leaves = build_leave_summary(classified_ytd, employees, year, month)
                validation = build_validation_report(raw_df, df, employees, absences, classified, year, month)
                alerts = build_alerts_report(employees, absences, classified, workdays, overtime_s, leaves, year)
                ergani_df = build_ergani_export_df(classified, employees, year)

            st.success("Η επεξεργασία ολοκληρώθηκε!")

            errors = validation[validation["Επίπεδο"] == "ERROR"]
            warnings = validation[validation["Επίπεδο"] == "WARNING"]

            col1, col2, col3 = st.columns(3)
            col1.metric("Απουσίες", len(absences))
            col2.metric("Σφάλματα", len(errors), delta=None if errors.empty else f"{len(errors)}", delta_color="inverse")
            col3.metric("Προειδοποιήσεις", len(warnings))

            if not errors.empty:
                with st.expander("🔴 Σφάλματα Validation", expanded=True):
                    st.dataframe(errors[["Κατηγορία", "Μήνυμα", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία", "Τιμή"]], use_container_width=True)

            if not warnings.empty:
                with st.expander("🟡 Προειδοποιήσεις Validation"):
                    st.dataframe(warnings[["Κατηγορία", "Μήνυμα", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία", "Τιμή"]], use_container_width=True)

            st.subheader("Λήψη Αρχείων")

            # Η κρίσιμη συνθήκη: έχει φορτωθεί classified ΜΕ πραγματικά δεδομένα;
            _has_classified = not classified.empty

            if not _has_classified:
                # Δεν υπάρχουν ταξινομημένες απουσίες — δείξε template
                # Auto-fill από προγραμματισμένες άδειες του μήνα
                _prefill = None
                try:
                    _planned_all = load_planned_df()
                    if not _planned_all.empty:
                        _pa = _planned_all.copy()
                        _pa["Ημ/νία"] = pd.to_datetime(_pa["Ημ/νία"], errors="coerce")
                        _prefill = _pa[
                            (_pa["Ημ/νία"].dt.year == year) &
                            (_pa["Ημ/νία"].dt.month == month)
                        ][["ΑΦΜ", "Ημ/νία", "Τύπος Απουσίας", "Έτος Άδειας"]]
                except Exception:
                    _prefill = None

                template_bytes = build_classified_template_excel_bytes(absences, prefill=_prefill)
                if classified_file or classified_bytes:
                    st.warning("⚠️ Το classified αρχείο δεν περιέχει συμπληρωμένες απουσίες (όλοι οι τύποι είναι κενοί). Συμπλήρωσε τη στήλη 'Τύπος Απουσίας' και ανέβασέ το ξανά.")
                    # Καθάρισε το λάθος cached classified από OneDrive
                    st.session_state.pop("classified_od_bytes", None)
                elif _prefill is not None and not _prefill.empty:
                    _matched = _prefill["ΑΦΜ"].astype(str).isin(absences["ΑΦΜ"].astype(str)).sum()
                    st.success(f"✨ Το template προσυμπληρώθηκε αυτόματα από {len(_prefill)} προγραμματισμένες άδειες. "
                               "Έλεγξε/συμπλήρωσε τα υπόλοιπα και ανέβασέ το ξανά.")
                else:
                    st.info("Κατέβασε το template, συμπλήρωσε τις στήλες 'Τύπος Απουσίας' και 'Έτος Άδειας', και ανέβασέ το ξανά.")
                st.download_button(
                    label="⬇ Κατέβασε classified_absences template",
                    data=template_bytes,
                    file_name=f"classified_absences_{year}_{month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                report_bytes = excel_bytes({
                    "Απουσίες": absences,
                    "Ημέρες": workdays,
                    "Υπερωρίες": overtime_d,
                    "Σύνολο Extra": overtime_s,
                    "Άδειες": leaves,
                    "Validation": validation,
                    "Alerts": alerts,
                })
                st.download_button(
                    label=f"⬇ monthly_report_{year}_{month:02d}.xlsx",
                    data=report_bytes,
                    file_name=f"monthly_report_{year}_{month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                if not ergani_df.empty:
                    st.markdown("**Ergani exports ανά παράρτημα:**")
                    for branch_value, branch_df in ergani_df.groupby("ΑΑ Παραρτηματος", dropna=False):
                        branch_out = branch_df.drop(columns=["ΑΑ Παραρτηματος"]).copy()
                        branch_label = int(branch_value) if pd.notna(branch_value) else "unknown"
                        st.download_button(
                            label=f"⬇ Ergani — Παράρτημα {branch_label}",
                            data=ergani_excel_bytes(branch_out),
                            file_name=f"ergani_export_parartima_{branch_label}_{year}_{month:02d}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                else:
                    st.warning("⚠️ Δεν παράχθηκαν Ergani exports — έλεγξε αν οι τύποι απουσίας στο classified αντιστοιχούν σε κωδικούς Εργάνης.")

                # Αποθήκευση αποτελεσμάτων για tab υπολοίπων
                st.session_state["leaves"] = leaves
                st.session_state["leaves_month"] = month
                st.session_state["leaves_year"] = year
                # Αναλυτικά δεδομένα (ανά ημέρα) για το per-employee dropdown
                st.session_state["classified_detail"] = classified_ytd.copy()

                # Auto-save στο OneDrive — μόνο αν το classified έχει πραγματικά δεδομένα
                od_token = st.session_state.get("od_token")
                if od_token:
                    try:
                        with st.spinner("Αποθήκευση στο OneDrive..."):
                            od.upload_file(od_token, f"raw_attendance_{year}_{month:02d}.xlsx", raw_file.getvalue(), subfolder="raw")
                            _cls_save_bytes = classified_file.getvalue() if classified_file else classified_bytes
                            if _cls_save_bytes:
                                od.upload_file(od_token, f"classified_absences_{year}_{month:02d}.xlsx", _cls_save_bytes, subfolder="output")
                            od.upload_file(od_token, f"monthly_report_{year}_{month:02d}.xlsx", report_bytes)
                            if not ergani_df.empty:
                                for branch_value, branch_df in ergani_df.groupby("ΑΑ Παραρτηματος", dropna=False):
                                    branch_out = branch_df.drop(columns=["ΑΑ Παραρτηματος"]).copy()
                                    branch_label = int(branch_value) if pd.notna(branch_value) else "unknown"
                                    od.upload_file(od_token, f"ergani_export_parartima_{branch_label}_{year}_{month:02d}.xlsx", ergani_excel_bytes(branch_out))
                        st.success("✅ Αποθηκεύτηκε στο OneDrive!")
                    except Exception as e:
                        st.warning(f"⚠️ Δεν ήταν δυνατή η αποθήκευση στο OneDrive: {e}")

        except Exception as e:
            st.error(f"Σφάλμα: {e}")


# =========================
# TAB: ΠΡΟΓΡΑΜΜΑΤΙΣΜΕΝΕΣ ΑΔΕΙΕΣ
# =========================

with tab_planned:
    st.subheader("📆 Προγραμματισμένες Άδειες")
    st.caption("Δήλωσε εκ των προτέρων τις άδειες ώστε να ξέρεις ποιοι θα λείπουν. "
               "Όταν έρθει η ώρα επεξεργασίας του μήνα, προσυμπληρώνουν αυτόματα το classified.")

    _pl_employees = get_employees_df()
    _pl_df = load_planned_df()
    _pl_leaves = st.session_state.get("leaves")  # για έλεγχο υπολοίπου (αν υπάρχει)
    _pl_today = datetime.date.today()

    st.text_input(
        "Όνομα εγκρίνοντος (καταγράφεται στις εγκρίσεις)",
        key="pl_approver",
        placeholder="π.χ. Χ. Διδάχος",
    )

    if _pl_employees is None or _pl_employees.empty:
        st.warning("⚠️ Δεν βρέθηκε το employees.xlsx. Σύνδεσε OneDrive ή ανέβασέ το από την καρτέλα Εκτέλεση.")
    else:
        # ── Φόρμα καταχώρησης ──────────────────────────────────────────
        st.markdown("### ➕ Νέα δήλωση άδειας")

        _emp_opts = {}
        for _, _e in _pl_employees.sort_values(["ΑΑ Παραρτηματος", "Επώνυμο", "Όνομα"]).iterrows():
            _br = _e["ΑΑ Παραρτηματος"]
            _br_s = f"[Υποκ. {int(_br)}] " if pd.notna(_br) else ""
            _label = f"{_br_s}{_e['Επώνυμο']} {_e['Όνομα']} — {_e['ΑΦΜ']}"
            _emp_opts[_label] = str(_e["ΑΦΜ"])

        _sel_emp_label = st.selectbox(
            "Υπάλληλος",
            options=list(_emp_opts.keys()),
            index=None,
            placeholder="Επίλεξε υπάλληλο…",
            key="pl_emp",
        )
        if not _sel_emp_label:
            st.caption("Επίλεξε υπάλληλο για να δηλώσεις νέα άδεια.")
        else:
            _sel_afm = _emp_opts.get(_sel_emp_label)
            _sel_emp_row = _pl_employees[_pl_employees["ΑΦΜ"].astype(str) == str(_sel_afm)].iloc[0]

            _types_list = ["Κανονική άδεια"] + sorted(t for t in ERGANI_LEAVE_TYPES if t != "Κανονική άδεια")
            _c1, _c2 = st.columns([2, 1])
            with _c1:
                _sel_type = st.selectbox("Τύπος άδειας", options=_types_list, key="pl_type")
            with _c2:
                _is_annual = _sel_type == "Κανονική άδεια"
                if _is_annual:
                    _sel_leave_year = st.selectbox(
                        "Έτος άδειας",
                        options=[_pl_today.year, _pl_today.year - 1],
                        key="pl_year",
                    )
                else:
                    _sel_leave_year = None
                    st.caption("—")

            _mode = st.radio("Τρόπος δήλωσης", options=["Εύρος ημερομηνιών", "Μεμονωμένες ημέρες"],
                             horizontal=True, key="pl_mode")

            _selected_days = []
            if _mode == "Εύρος ημερομηνιών":
                _dc1, _dc2 = st.columns(2)
                with _dc1:
                    _from = st.date_input("Από", value=_pl_today, key="pl_from")
                with _dc2:
                    _to = st.date_input("Έως", value=_pl_today, key="pl_to")
                _selected_days = working_days_in_range(_from, _to)
                if _selected_days:
                    st.caption(f"📅 {len(_selected_days)} εργάσιμες ημέρες (εξαιρούνται ΣΚ + αργίες)")
            else:
                _stage = st.session_state.setdefault("pl_staged_days", [])
                _sc1, _sc2 = st.columns([2, 1])
                with _sc1:
                    _new_day = st.date_input("Πρόσθεσε ημέρα", value=_pl_today, key="pl_single")
                with _sc2:
                    st.write("")
                    st.write("")
                    if st.button("➕ Πρόσθεσε", key="pl_add_day"):
                        _iso = _new_day.isoformat()
                        if _iso not in _stage:
                            _stage.append(_iso)
                            _stage.sort()
                if _stage:
                    st.caption("Επιλεγμένες ημέρες: " + ", ".join(
                        format_greek_date(d) for d in _stage
                    ))
                    if st.button("🗑️ Καθαρισμός ημερών", key="pl_clear_days"):
                        st.session_state["pl_staged_days"] = []
                        st.rerun()
                _selected_days = [pd.to_datetime(d).normalize() for d in _stage]

            _sel_status = st.selectbox("Κατάσταση", options=["εγκρίθηκε", "εκκρεμεί"], key="pl_status")

            # ── Έλεγχος υπολοίπου (μόνο για κανονική άδεια) ────────────────
            if _is_annual and _selected_days:
                _req = len(_selected_days)
                _avail = None
                if _pl_leaves is not None:
                    _lrow = _pl_leaves[_pl_leaves["ΑΦΜ"].astype(str) == str(_sel_afm)]
                    if not _lrow.empty:
                        _lrow = _lrow.iloc[0]
                        if _sel_leave_year == _pl_today.year:
                            _avail = int(_lrow.get("Υπόλοιπο Τρέχοντος Έτους Μετά", 0))
                        elif _sel_leave_year == _pl_today.year - 1:
                            _avail = int(_lrow.get("Υπόλοιπο Προηγούμενου Έτους Μετά", 0))
                if _avail is not None:
                    # Αφαίρεσε ήδη προγραμματισμένες κανονικές άδειες ίδιου έτους
                    _already = _pl_df[
                        (_pl_df["ΑΦΜ"].astype(str) == str(_sel_afm)) &
                        (_pl_df["Τύπος Απουσίας"] == "Κανονική άδεια") &
                        (_pl_df["Έτος Άδειας"] == _sel_leave_year)
                    ]
                    _net_avail = _avail - len(_already)
                    if _req > _net_avail:
                        st.error(f"⚠️ Ζητούνται **{_req}** ημέρες αλλά διαθέσιμο υπόλοιπο **{_net_avail}** "
                                 f"(υπόλοιπο {_avail} − {len(_already)} ήδη προγραμματισμένες).")
                    else:
                        st.success(f"✅ Διαθέσιμο υπόλοιπο: {_net_avail} ημέρες (ζητούνται {_req}).")
                else:
                    st.info("ℹ️ Δεν υπάρχουν δεδομένα υπολοίπου — τρέξε πρώτα την καρτέλα Υπόλοιπα για έλεγχο.")

            # ── Κουμπί καταχώρησης ─────────────────────────────────────────
            if st.button("💾 Καταχώρηση άδειας", type="primary", disabled=not _selected_days):
                try:
                    _now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    _approver = st.session_state.get("pl_approver", "").strip()
                    _new_rows = []
                    for _d in _selected_days:
                        _new_rows.append({
                            "ΑΑ Παραρτηματος": _sel_emp_row["ΑΑ Παραρτηματος"],
                            "ΑΦΜ": str(_sel_afm),
                            "Επώνυμο": _sel_emp_row["Επώνυμο"],
                            "Όνομα": _sel_emp_row["Όνομα"],
                            "Ημ/νία": pd.to_datetime(_d).normalize(),
                            "Τύπος Απουσίας": _sel_type,
                            "Έτος Άδειας": int(_sel_leave_year) if _is_annual else pd.NA,
                            "Κατάσταση": _sel_status,
                            "Καταχωρήθηκε": _now_str,
                            "Εγκρίθηκε από": _approver if _sel_status == "εγκρίθηκε" else "",
                            "Ημ. Έγκρισης": _now_str if _sel_status == "εγκρίθηκε" else "",
                        })
                    _new_df = pd.DataFrame(_new_rows)
                    _merged = pd.concat([_pl_df, _new_df], ignore_index=True)
                    # Νέα δήλωση υπερισχύει σε διπλότυπα (ΑΦΜ + Ημ/νία)
                    _merged = _merged.drop_duplicates(subset=["ΑΦΜ", "Ημ/νία"], keep="last").reset_index(drop=True)
                    save_planned_df(_merged)
                    st.session_state["pl_staged_days"] = []
                    st.success(f"✅ Καταχωρήθηκαν {len(_new_rows)} ημέρες για {_sel_emp_row['Επώνυμο']} {_sel_emp_row['Όνομα']}.")
                    st.rerun()
                except Exception as _e:
                    st.error(f"Σφάλμα καταχώρησης: {_e}")

        st.divider()

        # ── Εγκρίσεις εκκρεμών δηλώσεων ─────────────────────────────────
        st.markdown("### ✅ Εγκρίσεις")
        _appr_df = load_planned_df()
        _pending = (_appr_df[_appr_df["Κατάσταση"] == "εκκρεμεί"].copy()
                    if not _appr_df.empty else pd.DataFrame())
        if _pending.empty:
            st.caption("Δεν υπάρχουν εκκρεμείς δηλώσεις προς έγκριση.")
        else:
            _pending["Ημ/νία"] = pd.to_datetime(_pending["Ημ/νία"], errors="coerce")
            _grp = (
                _pending.groupby(["ΑΦΜ", "Επώνυμο", "Όνομα", "Τύπος Απουσίας"])
                .agg(_from=("Ημ/νία", "min"), _to=("Ημ/νία", "max"), _n=("Ημ/νία", "size"))
                .reset_index()
                .sort_values(["Επώνυμο", "Όνομα"])
            )
            for _i, _r in _grp.iterrows():
                _rng = (format_greek_date(_r["_from"]) if _r["_n"] == 1
                        else f"{format_greek_date(_r['_from'])} → {format_greek_date(_r['_to'])}")
                _ac1, _ac2, _ac3 = st.columns([5, 1, 1])
                with _ac1:
                    _icon = LEAVE_TYPE_ICON.get(_r["Τύπος Απουσίας"], "📌")
                    st.markdown(f"{_icon} **{_r['Επώνυμο']} {_r['Όνομα']}** · {_r['Τύπος Απουσίας']} · "
                                f"{_rng} ({int(_r['_n'])} ημ.)")
                _mask = (
                    (_appr_df["ΑΦΜ"].astype(str) == str(_r["ΑΦΜ"])) &
                    (_appr_df["Τύπος Απουσίας"] == _r["Τύπος Απουσίας"]) &
                    (_appr_df["Κατάσταση"] == "εκκρεμεί")
                )
                with _ac2:
                    if st.button("✅ Έγκριση", key=f"appr_ok_{_i}"):
                        _approver = st.session_state.get("pl_approver", "").strip()
                        if not _approver:
                            st.warning("Συμπλήρωσε πρώτα το «Όνομα εγκρίνοντος» στην κορυφή.")
                        else:
                            _appr_df.loc[_mask, "Κατάσταση"] = "εγκρίθηκε"
                            _appr_df.loc[_mask, "Εγκρίθηκε από"] = _approver
                            _appr_df.loc[_mask, "Ημ. Έγκρισης"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                            save_planned_df(_appr_df)
                            st.success(f"✅ Εγκρίθηκε: {_r['Επώνυμο']} {_r['Όνομα']}")
                            st.rerun()
                with _ac3:
                    if st.button("🗑️ Απόρριψη", key=f"appr_no_{_i}"):
                        _kept = _appr_df[~_mask].reset_index(drop=True)
                        save_planned_df(_kept)
                        st.success(f"🗑️ Απορρίφθηκε: {_r['Επώνυμο']} {_r['Όνομα']}")
                        st.rerun()

        st.divider()

        # ── Οπτικό ημερολόγιο μήνα ─────────────────────────────────────
        st.markdown("### 📅 Ημερολόγιο μήνα")
        _cal_df = load_planned_df()
        _cc1, _cc2 = st.columns(2)
        with _cc1:
            _cal_year = st.number_input("Έτος", min_value=2020, max_value=2100,
                                        value=_pl_today.year, step=1, key="cal_year")
        with _cc2:
            _cal_month = st.selectbox("Μήνας", options=list(MONTHS.keys()),
                                      format_func=lambda m: MONTHS[m],
                                      index=_pl_today.month - 1, key="cal_month")

        # Υπόμνημα χρωμάτων υποκαταστημάτων
        _branches = sorted(
            {int(b) for b in _cal_df["ΑΑ Παραρτηματος"].dropna().unique()}
        ) if not _cal_df.empty else []
        if _branches:
            _legend = " &nbsp; ".join(
                f'<span style="display:inline-flex;align-items:center;gap:4px">'
                f'<span style="width:10px;height:10px;border-radius:50%;'
                f'background:{_branch_color(b)};display:inline-block"></span>Υποκ. {b}</span>'
                for b in _branches
            )
            st.markdown(_legend + " &nbsp;·&nbsp; *πλάγια = εκκρεμεί*", unsafe_allow_html=True)

        st.markdown(
            build_calendar_html(_cal_df, int(_cal_year), int(_cal_month), _pl_today),
            unsafe_allow_html=True,
        )

        st.divider()

        # ── Επερχόμενες άδειες + κάλυψη ────────────────────────────────
        st.markdown("### 👥 Ποιοι λείπουν")
        _pl_df = load_planned_df()  # φρέσκο μετά από πιθανή αποθήκευση

        if _pl_df.empty:
            st.info("Δεν υπάρχουν προγραμματισμένες άδειες ακόμα.")
        else:
            _view = _pl_df.copy()
            _view["Ημ/νία"] = pd.to_datetime(_view["Ημ/νία"], errors="coerce")

            _fc1, _fc2 = st.columns(2)
            with _fc1:
                _range_from = st.date_input("Προβολή από", value=_pl_today, key="pl_view_from")
            with _fc2:
                _range_to = st.date_input("Προβολή έως", value=_pl_today + datetime.timedelta(days=30), key="pl_view_to")

            _mask = (
                (_view["Ημ/νία"].dt.date >= _range_from) &
                (_view["Ημ/νία"].dt.date <= _range_to)
            )
            _period = _view[_mask].sort_values(["Ημ/νία", "ΑΑ Παραρτηματος", "Επώνυμο"])

            if _period.empty:
                st.caption("Καμία άδεια στο επιλεγμένο διάστημα.")
            else:
                # Προειδοποίηση επικάλυψης ανά ημέρα/υποκατάστημα
                _overlap = (
                    _period.groupby([_period["Ημ/νία"].dt.date, "ΑΑ Παραρτηματος"])
                    .size().reset_index(name="Πλήθος")
                )
                _hot = _overlap[_overlap["Πλήθος"] >= 3]
                if not _hot.empty:
                    _msgs = []
                    for _, _h in _hot.iterrows():
                        _br = _h["ΑΑ Παραρτηματος"]
                        _br_s = f"Υποκ. {int(_br)}" if pd.notna(_br) else "—"
                        _msgs.append(f"{format_greek_date(_h.iloc[0])} · {_br_s}: {int(_h['Πλήθος'])} άτομα")
                    st.warning("⚠️ Ημέρες με πολλές ταυτόχρονες απουσίες:\n\n" + "\n\n".join(f"- {m}" for m in _msgs))

                # Λίστα ανά ημέρα
                for _day, _grp in _period.groupby(_period["Ημ/νία"].dt.date):
                    _names = []
                    for _, _r in _grp.iterrows():
                        _br = _r["ΑΑ Παραρτηματος"]
                        _br_s = f"[Υποκ. {int(_br)}] " if pd.notna(_br) else ""
                        _icon = LEAVE_TYPE_ICON.get(_r["Τύπος Απουσίας"], "📌")
                        _st_s = "" if _r["Κατάσταση"] == "εγκρίθηκε" else " *(εκκρεμεί)*"
                        _names.append(f"- {_br_s}{_r['Επώνυμο']} {_r['Όνομα']} · {_icon} {_r['Τύπος Απουσίας']}{_st_s}")
                    with st.expander(f"📅 {format_greek_date(_day)} — {len(_grp)} άτομα"):
                        st.markdown("\n".join(_names))

            st.divider()

            # ── Διαγραφή δηλώσεων ──────────────────────────────────────
            st.markdown("### 🗑️ Διαγραφή δηλώσεων")
            _del_view = _pl_df.copy()
            _del_view["Ημ/νία"] = pd.to_datetime(_del_view["Ημ/νία"], errors="coerce")
            _del_view = _del_view.sort_values(["ΑΑ Παραρτηματος", "Επώνυμο", "Ημ/νία"])
            _del_opts = {}
            for _idx, _r in _del_view.iterrows():
                _key = f"{_r['Επώνυμο']} {_r['Όνομα']} · {format_greek_date(_r['Ημ/νία'])} · {_r['Τύπος Απουσίας']}"
                _del_opts[_key] = (str(_r["ΑΦΜ"]), pd.to_datetime(_r["Ημ/νία"]).normalize())
            _to_delete = st.multiselect("Επίλεξε δηλώσεις προς διαγραφή", options=list(_del_opts.keys()), key="pl_del")
            if _to_delete and st.button("🗑️ Διαγραφή επιλεγμένων", key="pl_del_btn"):
                _del_keys = {_del_opts[k] for k in _to_delete}
                _kept = _pl_df[~_pl_df.apply(
                    lambda r: (str(r["ΑΦΜ"]), pd.to_datetime(r["Ημ/νία"]).normalize()) in _del_keys, axis=1
                )].reset_index(drop=True)
                save_planned_df(_kept)
                st.success(f"✅ Διαγράφηκαν {len(_to_delete)} δηλώσεις.")
                st.rerun()

            # Λήψη όλων σε Excel
            st.download_button(
                "⬇ Λήψη όλων των προγραμματισμένων (Excel)",
                data=excel_bytes({"Προγραμματισμένες": _pl_df}),
                file_name="planned_leaves.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pl_download",
            )


# =========================
# TAB: ΙΣΤΟΡΙΚΟ
# =========================

with tab_history:
    st.subheader("Παλαιότερα Αρχεία")

    od_token = st.session_state.get("od_token")

    # --- Upload υπαρχόντων αρχείων στο OneDrive ---
    if od_token:
        with st.expander("⬆ Ανέβασε υπάρχοντα αρχεία στο OneDrive"):
            st.caption("Ανέβασε παλιά monthly reports ή raw attendance αρχεία για να τα αποθηκεύσεις στο OneDrive.")
            upload_col1, upload_col2 = st.columns(2)
            with upload_col1:
                files_to_upload = st.file_uploader(
                    "Monthly reports / Ergani exports (.xlsx)",
                    type=["xlsx"],
                    accept_multiple_files=True,
                    key="manual_upload_output",
                )
                if files_to_upload and st.button("⬆ Ανέβασμα reports"):
                    with st.spinner("Ανέβασμα..."):
                        for f in files_to_upload:
                            try:
                                od.upload_file(od_token, f.name, f.getvalue(), subfolder="output")
                                st.success(f"✅ {f.name}")
                            except Exception as e:
                                st.error(f"❌ {f.name}: {e}")
            with upload_col2:
                raw_files_to_upload = st.file_uploader(
                    "Raw attendance αρχεία (.xlsx)",
                    type=["xlsx"],
                    accept_multiple_files=True,
                    key="manual_upload_raw",
                )
                if raw_files_to_upload and st.button("⬆ Ανέβασμα raw"):
                    with st.spinner("Ανέβασμα..."):
                        for f in raw_files_to_upload:
                            try:
                                od.upload_file(od_token, f.name, f.getvalue(), subfolder="raw")
                                st.success(f"✅ {f.name}")
                            except Exception as e:
                                st.error(f"❌ {f.name}: {e}")

            st.caption("Classified absences — ονόμασε τα αρχεία ως `classified_absences_YYYY_MM.xlsx`")
            cls_files_to_upload = st.file_uploader(
                "Classified absences πολλών μηνών (.xlsx)",
                type=["xlsx"],
                accept_multiple_files=True,
                key="manual_upload_cls",
            )
            if cls_files_to_upload and st.button("⬆ Ανέβασμα classified"):
                with st.spinner("Ανέβασμα..."):
                    for f in cls_files_to_upload:
                        try:
                            od.upload_file(od_token, f.name, f.getvalue(), subfolder="output")
                            st.success(f"✅ {f.name}")
                        except Exception as e:
                            st.error(f"❌ {f.name}: {e}")
        st.divider()

    # --- Αναδημιουργία Αναφορών ---
    if od_token:
        with st.expander("🔄 Αναδημιουργία Αναφορών (YTD fix)"):
            st.caption(
                "Βρίσκει μήνες που έχουν **raw** στο OneDrive και αναδημιουργεί το monthly report "
                "με σωστό YTD υπόλοιπο αδειών. Αν δεν υπάρχει classified, το ανακατασκευάζει "
                "αυτόματα από τα Ergani export αρχεία."
            )
            try:
                import re as _re
                _regen_raw_files = od.list_files(od_token, subfolder="raw")
                _regen_out_files = od.list_files(od_token, subfolder="output")

                # Βρες μήνες που έχουν raw_attendance_{year}_{MM}.xlsx στο "raw"
                _regen_months = []
                for _rf in _regen_raw_files:
                    _m = _re.match(r"raw_attendance_(\d{4})_(\d{2})\.xlsx", _rf["name"])
                    if _m:
                        _ry, _rm = int(_m.group(1)), int(_m.group(2))
                        _cls_name = f"classified_absences_{_ry}_{_rm:02d}.xlsx"
                        _has_cls = any(f["name"] == _cls_name for f in _regen_out_files)
                        # Έλεγχος αν υπάρχουν Ergani exports για αυτόν τον μήνα
                        _ergani_pat = _re.compile(
                            rf"ergani_export_parartima_(\d+)_{_ry}_{_rm:02d}\.xlsx"
                        )
                        _ergani_matches = [
                            (int(_ergani_pat.match(f["name"]).group(1)), f["name"])
                            for f in _regen_out_files
                            if _ergani_pat.match(f["name"])
                        ]
                        _regen_months.append((_ry, _rm, _has_cls, _ergani_matches))

                _regen_months.sort()

                if not _regen_months:
                    st.info("Δεν βρέθηκαν raw αρχεία στο OneDrive (subfolder: raw).")
                else:
                    # Κατάσταση κάθε μήνα
                    _ready = []
                    _no_source = []
                    for _ry, _rm, _has_cls, _ergani_m in _regen_months:
                        if _has_cls or _ergani_m:
                            _ready.append((_ry, _rm, _has_cls, _ergani_m))
                        else:
                            _no_source.append((_ry, _rm))

                    if _no_source:
                        st.warning(
                            "Οι παρακάτω μήνες δεν έχουν **ούτε classified ούτε Ergani exports** — "
                            "ανέβασε τα classified χειροκίνητα: "
                            + ", ".join(f"{MONTHS[_m]} {_y}" for _y, _m in _no_source)
                        )

                    if _ready:
                        for _ry, _rm, _has_cls, _ergani_m in _ready:
                            _src = "classified" if _has_cls else f"Ergani ({len(_ergani_m)} παρ.)"
                            st.write(f"✅ **{MONTHS[_rm]} {_ry}** — πηγή: {_src}")

                        if st.button("🔄 Αναδημιουργία Όλων", type="primary"):
                            try:
                                # Φόρτωση employees
                                _regen_emp_bytes = od.download_file(od_token, "employees.xlsx", subfolder="config")
                                _regen_emp_tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                                _regen_emp_tmp.write(_regen_emp_bytes); _regen_emp_tmp.close()
                                _regen_employees = load_employees(Path(_regen_emp_tmp.name))

                                # Φόρτωση / ανακατασκευή ΟΛΩΝ των classified για YTD (μία φορά)
                                _regen_all_cls = {}
                                for _cy, _cm, _cok, _cerg in _regen_months:
                                    if not (_cok or _cerg):
                                        continue
                                    if _cok:
                                        _cn = f"classified_absences_{_cy}_{_cm:02d}.xlsx"
                                        _cb = od.download_file(od_token, _cn, subfolder="output")
                                        _ct = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                                        _ct.write(_cb); _ct.close()
                                        _regen_all_cls[(_cy, _cm)] = load_classified_absences(Path(_ct.name))
                                    else:
                                        # Ανακατασκευή από Ergani exports
                                        _erg_pairs = []
                                        for _aa, _ename in _cerg:
                                            _eb = od.download_file(od_token, _ename, subfolder="output")
                                            _erg_pairs.append((_aa, _eb))
                                        _reconstructed = reconstruct_classified_from_ergani(_erg_pairs)
                                        if not _reconstructed.empty:
                                            # Αποθήκευση στο OneDrive για μελλοντική χρήση
                                            _rcls_name = f"classified_absences_{_cy}_{_cm:02d}.xlsx"
                                            _rcls_buf = io.BytesIO()
                                            with pd.ExcelWriter(_rcls_buf, engine="openpyxl") as _rw:
                                                format_dates_for_excel(_reconstructed).to_excel(_rw, index=False)
                                                force_text_column(_rw.sheets["Sheet1"], "ΑΦΜ")
                                            od.upload_file(od_token, _rcls_name, _rcls_buf.getvalue(), subfolder="output")
                                            _regen_all_cls[(_cy, _cm)] = _reconstructed

                                _regen_progress = st.progress(0)
                                _regen_status = st.empty()

                                for _idx, (_ry, _rm, _rhas_cls, _rergani) in enumerate(_ready):
                                    _regen_status.info(f"Επεξεργασία {MONTHS[_rm]} {_ry}...")

                                    # Raw
                                    _regen_raw_b = od.download_file(od_token, f"raw_attendance_{_ry}_{_rm:02d}.xlsx", subfolder="raw")
                                    _regen_raw_t = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                                    _regen_raw_t.write(_regen_raw_b); _regen_raw_t.close()

                                    _regen_raw_df = load_attendance(Path(_regen_raw_t.name))
                                    _regen_df = clean_attendance(_regen_raw_df)

                                    _regen_cls = _regen_all_cls.get((_ry, _rm), pd.DataFrame())

                                    # YTD: όλοι οι προηγούμενοι μήνες του ίδιου έτους
                                    _regen_ytd = _regen_cls.copy()
                                    for _pm in range(1, _rm):
                                        _prev = _regen_all_cls.get((_ry, _pm), pd.DataFrame())
                                        if not _prev.empty:
                                            _regen_ytd = pd.concat([_prev, _regen_ytd], ignore_index=True)

                                    _regen_abs = find_absences(_regen_df, _regen_employees, _ry, _rm)
                                    _regen_wd = calculate_work_days(_regen_df, _ry, _rm)
                                    _regen_ot_d, _regen_ot_s = calculate_overtime(_regen_df.copy(), _ry, _rm)
                                    _regen_leaves = build_leave_summary(_regen_ytd, _regen_employees, _ry, _rm)
                                    _regen_val = build_validation_report(_regen_raw_df, _regen_df, _regen_employees, _regen_abs, _regen_cls, _ry, _rm)
                                    _regen_alerts = build_alerts_report(_regen_employees, _regen_abs, _regen_cls, _regen_wd, _regen_ot_s, _regen_leaves, _ry)

                                    _regen_report = excel_bytes({
                                        "Απουσίες": _regen_abs,
                                        "Ημέρες": _regen_wd,
                                        "Υπερωρίες": _regen_ot_d,
                                        "Σύνολο Extra": _regen_ot_s,
                                        "Άδειες": _regen_leaves,
                                        "Validation": _regen_val,
                                        "Alerts": _regen_alerts,
                                    })

                                    od.upload_file(od_token, f"monthly_report_{_ry}_{_rm:02d}.xlsx", _regen_report)

                                    _regen_progress.progress((_idx + 1) / len(_ready))

                                _regen_status.success(f"✅ Αναδημιουργήθηκαν {len(_ready)} αναφορές!")
                            except Exception as _regen_err:
                                st.error(f"❌ Σφάλμα: {_regen_err}")
                    else:
                        st.info("Δεν υπάρχουν μήνες με raw + classified.")
            except Exception as _e:
                st.error(f"Σφάλμα φόρτωσης λίστας OneDrive: {_e}")
        st.divider()

    # --- employees.xlsx από config ---
    if od_token:
        try:
            cfg_files = od.list_files(od_token, subfolder="config")
            emp_cfg = next((f for f in cfg_files if f["name"] == "employees.xlsx"), None)
            if emp_cfg:
                with st.expander("👥 employees.xlsx"):
                    emp_content = od.download_file(od_token, "employees.xlsx", subfolder="config")
                    st.download_button(
                        label="⬇ employees.xlsx",
                        data=emp_content,
                        file_name="employees.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_employees_history",
                    )
                    new_emp = st.file_uploader("Αντικατάσταση employees.xlsx", type=["xlsx"], key="hist_emp_upload")
                    if new_emp and st.button("⬆ Αποθήκευση στο OneDrive", key="hist_emp_save"):
                        od.upload_file(od_token, "employees.xlsx", new_emp.getvalue(), subfolder="config")
                        st.session_state.pop("employees_od_bytes", None)  # invalidate cache
                        st.success("✅ Αποθηκεύτηκε!")
                st.divider()
        except Exception:
            pass

    if od_token:
        # Φόρτωσε από OneDrive
        try:
            files = od.list_files(od_token, subfolder="output")
            if not files:
                st.info("Δεν υπάρχουν αρχεία στο OneDrive ακόμα.")
            else:
                # Φόρτωσε και raw αρχεία
                raw_files = od.list_files(od_token, subfolder="raw")

                # Ομαδοποίηση ανά περίοδο (output + raw)
                all_od_files = files + raw_files
                periods = sorted(
                    set(
                        "_".join(f["name"].replace(".xlsx", "").split("_")[-2:])
                        for f in all_od_files
                        if f["name"].endswith(".xlsx")
                    ),
                    reverse=True,
                )
                for period in periods:
                    try:
                        y, m = period.split("_")
                        label = f"{MONTHS[int(m)]} {y}"
                    except Exception:
                        label = period

                    period_output = [f for f in files if period in f["name"]]
                    period_raw = [f for f in raw_files if period in f["name"]]

                    with st.expander(f"📅 {label}"):
                        if period_output:
                            st.caption("📊 Reports")
                            for f in period_output:
                                try:
                                    content = od.download_file(od_token, f["name"], subfolder="output")
                                    st.download_button(
                                        label=f"⬇ {f['name']}",
                                        data=content,
                                        file_name=f["name"],
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"od_{f['name']}",
                                    )
                                except Exception as e:
                                    st.warning(f"Δεν ήταν δυνατή η λήψη του {f['name']}: {e}")
                        if period_raw:
                            st.caption("📁 Raw Attendance")
                            for f in period_raw:
                                try:
                                    content = od.download_file(od_token, f["name"], subfolder="raw")
                                    st.download_button(
                                        label=f"⬇ {f['name']}",
                                        data=content,
                                        file_name=f["name"],
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"od_raw_{f['name']}",
                                    )
                                except Exception as e:
                                    st.warning(f"Δεν ήταν δυνατή η λήψη του {f['name']}: {e}")

                        # Ανέβασμα classified για αυτόν τον μήνα (για YTD υπολογισμό)
                        try:
                            _py, _pm = int(y), int(m)
                            _cls_name = f"classified_absences_{_py}_{_pm:02d}.xlsx"
                            _cls_exists = any(f["name"] == _cls_name for f in files)
                            st.caption("📋 Classified Absences")
                            if _cls_exists:
                                _cls_content = od.download_file(od_token, _cls_name, subfolder="output")
                                st.download_button(
                                    label=f"⬇ {_cls_name}",
                                    data=_cls_content,
                                    file_name=_cls_name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"od_cls_{_cls_name}",
                                )
                            _cls_upload = st.file_uploader(
                                f"{'Αντικατάσταση' if _cls_exists else '⬆ Ανέβασε'} {_cls_name}",
                                type=["xlsx"],
                                key=f"hist_cls_{period}",
                            )
                            if _cls_upload:
                                if st.button("💾 Αποθήκευση στο OneDrive", key=f"hist_cls_save_{period}"):
                                    od.upload_file(od_token, _cls_name, _cls_upload.getvalue(), subfolder="output")
                                    st.success(f"✅ {_cls_name} αποθηκεύτηκε!")
                                    st.rerun()
                        except Exception:
                            pass
        except Exception as e:
            st.error(f"Σφάλμα φόρτωσης από OneDrive: {e}")
    else:
        # Fallback: τοπικά αρχεία
        if not OUTPUT_DIR.exists():
            st.info("Δεν βρέθηκε φάκελος output. Σύνδεσε το OneDrive ή τρέξε τοπικά.")
        else:
            reports = sorted(OUTPUT_DIR.glob("monthly_report_*.xlsx"), reverse=True)
            ergani_files = sorted(OUTPUT_DIR.glob("ergani_export_*.xlsx"), reverse=True)
            all_files = [*reports, *ergani_files]

            if not all_files:
                st.info("Δεν υπάρχουν αρχεία ακόμα.")
            else:
                periods = sorted(
                    set(
                        "_".join(f.stem.split("_")[-2:])
                        for f in all_files
                        if len(f.stem.split("_")) >= 2
                    ),
                    reverse=True,
                )
                for period in periods:
                    try:
                        y, m = period.split("_")
                        label = f"{MONTHS[int(m)]} {y}"
                    except Exception:
                        label = period

                    with st.expander(f"📅 {label}"):
                        for f in [*reports, *ergani_files]:
                            if period in f.stem:
                                st.download_button(
                                    label=f"⬇ {f.name}",
                                    data=f.read_bytes(),
                                    file_name=f.name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=str(f),
                                )


# =========================
# TAB: ΥΠΟΛΟΙΠΑ ΑΔΕΙΩΝ
# =========================

with tab_balances:
    st.subheader("Υπόλοιπα Κανονικής Άδειας")

    leaves_df = st.session_state.get("leaves")
    leaves_month = st.session_state.get("leaves_month", _today.month)
    leaves_year = st.session_state.get("leaves_year", _today.year)

    od_token = st.session_state.get("od_token")

    # Αυτόματος υπολογισμός YTD από OneDrive
    if od_token and leaves_df is None:
        try:
            _bal_files = od.list_files(od_token, subfolder="output")
            _bal_year = _today.year

            # Βρες τον τελευταίο μήνα classified και το τελευταίο monthly report
            _bal_last_cls_month = max(
                (int(f["name"].replace(".xlsx","").split("_")[-1])
                 for f in _bal_files
                 if f["name"].startswith(f"classified_absences_{_bal_year}_")),
                default=0
            )
            _bal_report_files = sorted(
                [f["name"] for f in _bal_files
                 if f["name"].startswith(f"monthly_report_{_bal_year}_") and f["name"].endswith(".xlsx")],
                reverse=True,
            )
            _bal_last_report_month = (
                int(_bal_report_files[0].replace(".xlsx","").split("_")[-1])
                if _bal_report_files else 0
            )

            _bal_emp_bytes = st.session_state.get("employees_od_bytes")
            if _bal_emp_bytes is None:
                try:
                    _cfg = od.list_files(od_token, subfolder="config")
                    if any(f["name"] == "employees.xlsx" for f in _cfg):
                        _bal_emp_bytes = od.download_file(od_token, "employees.xlsx", subfolder="config")
                except Exception:
                    pass

            if _bal_emp_bytes and (_bal_last_cls_month > 0 or _bal_last_report_month > 0):
                _bal_emp_tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                _bal_emp_tmp.write(_bal_emp_bytes); _bal_emp_tmp.close()
                _bal_employees = load_employees(Path(_bal_emp_tmp.name))

                if _bal_last_cls_month >= _bal_last_report_month:
                    # Classified καλύπτουν τον πιο πρόσφατο μήνα → YTD από classified
                    _bal_classified_list = []
                    for _bm in range(1, _bal_last_cls_month + 1):
                        _bcls = f"classified_absences_{_bal_year}_{_bm:02d}.xlsx"
                        if any(f["name"] == _bcls for f in _bal_files):
                            _bb = od.download_file(od_token, _bcls, subfolder="output")
                            _bt = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                            _bt.write(_bb); _bt.close()
                            _bc = load_classified_absences(Path(_bt.name))
                            if not _bc.empty:
                                _bal_classified_list.append(_bc)
                    if _bal_classified_list:
                        _bal_classified_all = pd.concat(_bal_classified_list, ignore_index=True)
                        leaves_df = build_leave_summary(_bal_classified_all, _bal_employees, _bal_year, _bal_last_cls_month)
                        leaves_month = _bal_last_cls_month
                        leaves_year = _bal_year
                        # Αναλυτικά δεδομένα για το per-employee dropdown
                        st.session_state["classified_detail"] = _bal_classified_all.copy()
                        _missing = [m for m in range(1, _bal_last_cls_month + 1)
                                    if not any(f["name"] == f"classified_absences_{_bal_year}_{m:02d}.xlsx" for f in _bal_files)]
                        _note = f"📂 YTD classified — {_bal_last_cls_month} μήνες ({_bal_year})"
                        if _missing:
                            _note += f" ⚠️ Λείπουν: {', '.join(MONTHS[m] for m in _missing)}"
                        st.caption(_note)
                else:
                    # Monthly report είναι νεότερο → χρησιμοποίησε το ως πιο πλήρες snapshot
                    _content = od.download_file(od_token, _bal_report_files[0], subfolder="output")
                    leaves_df = pd.read_excel(io.BytesIO(_content), sheet_name="Άδειες")
                    leaves_month = _bal_last_report_month
                    leaves_year = _bal_year
                    _cls_note = f"μέχρι {MONTHS[_bal_last_cls_month]}" if _bal_last_cls_month else "κανένα"
                    st.caption(f"📂 Monthly report {MONTHS[_bal_last_report_month]} {_bal_year} "
                               f"(classified: {_cls_note})")
        except Exception:
            pass

    if leaves_df is None and not od_token and OUTPUT_DIR.exists():
        reports = sorted(OUTPUT_DIR.glob("monthly_report_*.xlsx"), reverse=True)
        if reports:
            try:
                leaves_df = pd.read_excel(reports[0], sheet_name="Άδειες")
                parts = reports[0].stem.split("_")
                leaves_month = int(parts[-1])
                leaves_year = int(parts[-2])
                st.caption(f"Από: {reports[0].name}")
            except Exception:
                pass

    if leaves_df is None and OUTPUT_DIR.exists():
        reports = sorted(OUTPUT_DIR.glob("monthly_report_*.xlsx"), reverse=True)
        if reports:
            try:
                leaves_df = pd.read_excel(reports[0], sheet_name="Άδειες")
                parts = reports[0].stem.split("_")
                leaves_month = int(parts[-1])
                leaves_year = int(parts[-2])
                st.caption(f"Από: {reports[0].name}")
            except Exception:
                pass

    if leaves_df is None:
        st.info("Δεν υπάρχουν δεδομένα. Τρέξε πρώτα μια εκτέλεση.")
    else:
        # ── Διαχωρισμός ενεργών / αποχωρούντων (βάσει Ημερομηνίας Αποχώρησης) ──
        _bal_today = datetime.date.today()
        if "Ημερομηνία Αποχώρησης" in leaves_df.columns:
            _dep_dates = pd.to_datetime(
                leaves_df["Ημερομηνία Αποχώρησης"], dayfirst=True, errors="coerce"
            )
            _departed_mask = _dep_dates.notna() & (_dep_dates.dt.date <= _bal_today)
        else:
            _departed_mask = pd.Series(False, index=leaves_df.index)

        leaves_active = leaves_df[~_departed_mask].reset_index(drop=True)
        leaves_departed = leaves_df[_departed_mask].reset_index(drop=True)

        def _balance_color(val):
            try:
                v = float(val)
            except (TypeError, ValueError):
                return ""
            if v <= 3:
                return "color: #c0392b; font-weight: bold"
            elif v <= 7:
                return "color: #b7770d; font-weight: bold"
            return "color: #1a7a3c; font-weight: bold"

        def _to_display(df: pd.DataFrame) -> pd.DataFrame:
            """Μετατρέπει nullable / exotic dtypes σε standard Python types
            ώστε το pandas Styler να μην κολλάει."""
            out = df.copy()
            for col in out.columns:
                dtype_str = str(out[col].dtype)
                if dtype_str in ("Int64", "Int32", "Int16", "Int8",
                                 "UInt64", "UInt32", "UInt16", "UInt8"):
                    out[col] = out[col].astype(object).where(out[col].notna(), other=None)
            return out

        def _show_table(df, balance_col, col_cfg):
            display = _to_display(df)
            try:
                styled = display.style.map(_balance_color, subset=[balance_col])
                st.dataframe(styled, use_container_width=True,
                             hide_index=True, column_config=col_cfg)
            except Exception:
                # Fallback χωρίς styling αν κάτι πάει στραβά
                st.dataframe(display, use_container_width=True,
                             hide_index=True, column_config=col_cfg)

        _col_cfg_curr = {
            "Υποκατ.":      st.column_config.NumberColumn("Υποκατ.", width=80, format="%d"),
            "ΑΦΜ":          st.column_config.TextColumn("ΑΦΜ", width=110),
            "Επώνυμο":      st.column_config.TextColumn("Επώνυμο", width=130),
            "Όνομα":        st.column_config.TextColumn("Όνομα", width=110),
            "Δικαιούμενες": st.column_config.NumberColumn("Δικαιούμενες", width=110, format="%d ημ."),
            "Ληφθείσες":    st.column_config.NumberColumn("Ληφθείσες", width=100, format="%d ημ."),
            "Υπόλοιπο":     st.column_config.NumberColumn("Υπόλοιπο", width=100, format="%d ημ."),
        }
        _col_cfg_prev = {
            "Υποκατ.":      st.column_config.NumberColumn("Υποκατ.", width=80, format="%d"),
            "ΑΦΜ":          st.column_config.TextColumn("ΑΦΜ", width=110),
            "Επώνυμο":      st.column_config.TextColumn("Επώνυμο", width=130),
            "Όνομα":        st.column_config.TextColumn("Όνομα", width=110),
            "Διαθέσιμες":   st.column_config.NumberColumn("Διαθέσιμες", width=110, format="%d ημ."),
            "Ληφθείσες":    st.column_config.NumberColumn("Ληφθείσες", width=100, format="%d ημ."),
            "Υπόλοιπο":     st.column_config.NumberColumn("Υπόλοιπο", width=100, format="%d ημ."),
        }

        # Τρέχον έτος (μόνο ενεργοί)
        st.subheader(f"📅 Τρέχον Έτος{f' {leaves_year}' if leaves_year else ''}")
        curr_table = leave_balance_table_current(leaves_active)
        _show_table(curr_table, "Υπόλοιπο", _col_cfg_curr)

        # Προηγούμενο έτος — Ιανουάριος έως Απρίλιος
        if 1 <= leaves_month <= 4:
            prev_table = leave_balance_table_prev(leaves_active)
            has_prev_balance = prev_table["Υπόλοιπο"].sum() > 0
            if has_prev_balance:
                prev_year = int(leaves_year) - 1 if leaves_year else ""
                st.subheader(f"📅 Προηγούμενο Έτος{f' {prev_year}' if prev_year else ''}")
                st.caption("⚠️ Το υπόλοιπο λήγει στο τέλος Μαρτίου.")
                _show_table(prev_table, "Υπόλοιπο", _col_cfg_prev)

        # Δεκέμβριος — προεπισκόπηση υπολοίπου που μεταφέρεται στο επόμενο έτος
        if leaves_month == 12:
            curr_table_dec = leave_balance_table_current(leaves_active)
            carryover = curr_table_dec[curr_table_dec["Υπόλοιπο"] > 0]
            if not carryover.empty:
                next_year = int(leaves_year) + 1 if leaves_year else ""
                st.subheader(f"📅 Μεταφορά υπολοίπου στο {next_year}")
                st.caption("Οι παρακάτω εργαζόμενοι έχουν υπόλοιπο που μεταφέρεται στο νέο έτος (λήγει τέλος Μαρτίου).")
                _show_table(carryover, "Υπόλοιπο", _col_cfg_curr)

        # ── Αναλυτικά ανά υπάλληλο: dropdown με τις ημέρες απουσίας ──
        _detail = st.session_state.get("classified_detail")
        st.divider()
        st.subheader("🔍 Αναλυτικά ανά υπάλληλο")

        if _detail is None or _detail.empty:
            st.caption("Δεν υπάρχουν αναλυτικά δεδομένα απουσιών (φόρτωσε classified με συμπληρωμένους τύπους).")
        else:
            _det = _detail.copy()
            _det["ΑΦΜ"] = _det["ΑΦΜ"].astype(str)
            _det["Ημ/νία"] = pd.to_datetime(_det["Ημ/νία"], errors="coerce")
            _det = _det.dropna(subset=["Ημ/νία"])

            # Φίλτρο τύπου άδειας
            _all_types = sorted(t for t in _det["Τύπος Απουσίας"].dropna().unique() if t)
            _sel_types = st.multiselect(
                "Φίλτρο τύπου άδειας (κενό = όλοι)",
                options=_all_types,
                default=[],
                key="detail_type_filter",
            )
            _det_view = _det[_det["Τύπος Απουσίας"].isin(_sel_types)] if _sel_types else _det

            def _emp_order_of(_src: pd.DataFrame) -> pd.DataFrame:
                _cols = ["ΑΦΜ", "Επώνυμο", "Όνομα", "ΑΑ Παραρτηματος"]
                if "Ημερομηνία Αποχώρησης" in _src.columns:
                    _cols.append("Ημερομηνία Αποχώρησης")
                _o = _src[_cols].copy()
                _o["ΑΦΜ"] = _o["ΑΦΜ"].astype(str)
                _o["_br"] = pd.to_numeric(_o["ΑΑ Παραρτηματος"], errors="coerce")
                return _o.sort_values(["_br", "Επώνυμο", "Όνομα"], na_position="last")

            _emp_order = _emp_order_of(leaves_df)  # όλοι (για λήψη Excel)

            # Κουμπί λήψης αναλυτικής λίστας σε Excel (σέβεται το φίλτρο, όλοι)
            if not _det_view.empty:
                _dl = _det_view.merge(
                    _emp_order[["ΑΦΜ", "_br"]], on="ΑΦΜ", how="left"
                ).sort_values(["_br", "Επώνυμο", "Όνομα", "Ημ/νία"], na_position="last")
                _dl_out = pd.DataFrame({
                    "Υποκατάστημα": _dl["_br"].apply(lambda x: int(x) if pd.notna(x) else ""),
                    "ΑΦΜ": _dl["ΑΦΜ"].astype(str),
                    "Επώνυμο": _dl["Επώνυμο"],
                    "Όνομα": _dl["Όνομα"],
                    "Ημέρα": _dl["Ημ/νία"].apply(format_greek_date),
                    "Τύπος Άδειας": _dl["Τύπος Απουσίας"],
                    "Έτος Άδειας": _dl.apply(
                        lambda r: int(r["Έτος Άδειας"])
                        if (r["Τύπος Απουσίας"] == "Κανονική άδεια" and pd.notna(r.get("Έτος Άδειας")))
                        else "",
                        axis=1,
                    ),
                })
                st.download_button(
                    label="⬇ Λήψη αναλυτικής λίστας (Excel)",
                    data=excel_bytes({"Αναλυτικές Απουσίες": _dl_out}),
                    file_name=f"analytika_apousies_{leaves_year}_{leaves_month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="detail_excel_dl",
                )

            def _render_detail(emp_rows: pd.DataFrame) -> int:
                shown = 0
                for _, _emp in emp_rows.iterrows():
                    _afm = str(_emp["ΑΦΜ"])
                    _emp_days = _det_view[_det_view["ΑΦΜ"] == _afm].sort_values("Ημ/νία")
                    if _emp_days.empty:
                        continue
                    shown += 1
                    _n = len(_emp_days)
                    _br = _emp["ΑΑ Παραρτηματος"]
                    _br_str = f" · Υποκ. {int(_br)}" if pd.notna(_br) else ""
                    _dep = _emp.get("Ημερομηνία Αποχώρησης", "")
                    _dep_str = f" · 🚪 αποχώρηση {_dep}" if isinstance(_dep, str) and _dep.strip() else ""
                    _label = (f"👤 {_emp['Επώνυμο']} {_emp['Όνομα']}{_br_str}{_dep_str} · "
                              f"{_n} ημέρ{'α' if _n == 1 else 'ες'} απουσίας")
                    with st.expander(_label):
                        for _lt, _grp in _emp_days.groupby("Τύπος Απουσίας", sort=False):
                            _icon = LEAVE_TYPE_ICON.get(_lt, "📌")
                            _cnt = len(_grp)
                            _is_annual = _lt == "Κανονική άδεια"
                            _has_year_col = "Έτος Άδειας" in _grp.columns
                            st.markdown(f"**{_icon} {_lt}** — {_cnt} ημέρ{'α' if _cnt == 1 else 'ες'}")
                            _grp_lines = []
                            for _, _row in _grp.iterrows():
                                _date_str = format_greek_date(_row["Ημ/νία"])
                                _yr = _row["Έτος Άδειας"] if _has_year_col else None
                                if _is_annual and pd.notna(_yr):
                                    _grp_lines.append(f"- {_date_str}  ·  *άδεια έτους {int(_yr)}*")
                                else:
                                    _grp_lines.append(f"- {_date_str}")
                            st.markdown("\n".join(_grp_lines))
                return shown

            # Ενεργοί
            _shown = _render_detail(_emp_order_of(leaves_active))
            if _shown == 0:
                st.caption("Κανένας ενεργός υπάλληλος δεν έχει απουσίες για το επιλεγμένο φίλτρο.")

            # Αποχωρούντες — διατηρείται το ιστορικό τους
            _dep_rows = _emp_order_of(leaves_departed) if not leaves_departed.empty else pd.DataFrame()

            # Δίχτυ ασφαλείας: υπάλληλοι που αφαιρέθηκαν τελείως από το employees.xlsx
            # αλλά έχουν ιστορικό απουσιών στο classified
            _known_afms = set(leaves_df["ΑΦΜ"].astype(str))
            _orphan = _det[~_det["ΑΦΜ"].astype(str).isin(_known_afms)]
            if not _orphan.empty:
                _orphan_rows = (
                    _orphan[["ΑΦΜ", "Επώνυμο", "Όνομα", "ΑΑ Παραρτηματος"]]
                    .drop_duplicates(subset=["ΑΦΜ"]).copy()
                )
                _orphan_rows["ΑΦΜ"] = _orphan_rows["ΑΦΜ"].astype(str)
                _orphan_rows["Ημερομηνία Αποχώρησης"] = ""
                _orphan_rows["_br"] = pd.to_numeric(_orphan_rows["ΑΑ Παραρτηματος"], errors="coerce")
                _dep_rows = pd.concat([_dep_rows, _orphan_rows], ignore_index=True) if not _dep_rows.empty else _orphan_rows

            if not _dep_rows.empty:
                _dep_rows = _dep_rows.sort_values(["_br", "Επώνυμο", "Όνομα"], na_position="last")
                st.divider()
                st.subheader("🚪 Αποχωρούντες")
                st.caption("Το ιστορικό απουσιών όσων έχουν αποχωρήσει διατηρείται εδώ "
                           "(δεν προσμετρώνται στα τρέχοντα υπόλοιπα).")
                _shown_dep = _render_detail(_dep_rows)
                if _shown_dep == 0:
                    st.caption("Οι αποχωρούντες δεν έχουν καταγεγραμμένες απουσίες για το επιλεγμένο φίλτρο.")

        # ── Τέλος έτους: μεταφορά υπολοίπων στο επόμενο έτος ────────────
        st.divider()
        _next_year = int(leaves_year) + 1 if leaves_year else _bal_today.year + 1
        with st.expander(f"🔄 Τέλος έτους — Δημιουργία employees.xlsx για το {_next_year}"):
            st.caption(
                f"Υπολογίζει αυτόματα το νέο αρχείο εργαζομένων για το **{_next_year}**: "
                "το υπόλοιπο του τρέχοντος έτους μεταφέρεται ως «Υπόλοιπο Προηγούμενου Έτους» "
                "(λήγει τέλος Μαρτίου). Μπορείς να διορθώσεις τη «Δικαιούμενη Τρέχοντος» πριν τη λήψη."
            )

            _cy_src = leaves_active.copy()  # μόνο ενεργοί (οι αποχωρούντες δεν μεταφέρονται)
            if _cy_src.empty:
                st.info("Δεν υπάρχουν ενεργοί εργαζόμενοι για μεταφορά.")
            else:
                _new_emp = pd.DataFrame({
                    "ΑΑ Παραρτηματος": _cy_src["ΑΑ Παραρτηματος"],
                    "ΑΦΜ": _cy_src["ΑΦΜ"].astype(str),
                    "Επώνυμο": _cy_src["Επώνυμο"],
                    "Όνομα": _cy_src["Όνομα"],
                    "Ημερομηνία Πρόσληψης": _cy_src["Ημερομηνία Πρόσληψης"],
                    "Ημερομηνία Αποχώρησης": _cy_src.get("Ημερομηνία Αποχώρησης", ""),
                    "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους":
                        pd.to_numeric(_cy_src["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"], errors="coerce").fillna(0).astype(int),
                    "Υπόλοιπο Προηγούμενου Έτους":
                        pd.to_numeric(_cy_src["Υπόλοιπο Τρέχοντος Έτους Μετά"], errors="coerce").fillna(0).astype(int),
                    "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους":
                        pd.to_numeric(_cy_src["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"], errors="coerce").fillna(0).astype(int),
                }).sort_values(["ΑΑ Παραρτηματος", "Επώνυμο", "Όνομα"]).reset_index(drop=True)

                st.markdown("**Προεπισκόπηση — μπορείς να διορθώσεις τη «Δικαιούμενη Τρέχοντος»:**")
                _edited = st.data_editor(
                    _new_emp,
                    use_container_width=True,
                    hide_index=True,
                    disabled=[c for c in _new_emp.columns if c != "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"],
                    column_config={
                        "ΑΦΜ": st.column_config.TextColumn("ΑΦΜ"),
                        "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους":
                            st.column_config.NumberColumn("Δικ. Προηγ.", format="%d"),
                        "Υπόλοιπο Προηγούμενου Έτους":
                            st.column_config.NumberColumn("Υπόλ. Προηγ. (μεταφορά)", format="%d"),
                        "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους":
                            st.column_config.NumberColumn(f"Δικ. Τρέχοντος ({_next_year})", format="%d"),
                    },
                    key="carryover_editor",
                )

                _new_emp_bytes = excel_bytes({"employees": _edited})
                _cc1, _cc2 = st.columns(2)
                with _cc1:
                    st.download_button(
                        f"⬇ Λήψη employees_{_next_year}.xlsx",
                        data=_new_emp_bytes,
                        file_name=f"employees_{_next_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="carryover_dl",
                    )
                with _cc2:
                    if st.session_state.get("od_token"):
                        if st.button(f"☁️ Αποθήκευση στο OneDrive ως employees_{_next_year}.xlsx", key="carryover_save"):
                            try:
                                od.upload_file(
                                    st.session_state["od_token"],
                                    f"employees_{_next_year}.xlsx",
                                    _new_emp_bytes,
                                    subfolder="config",
                                )
                                st.success(f"✅ Αποθηκεύτηκε ως employees_{_next_year}.xlsx στο OneDrive (config).")
                            except Exception as _e:
                                st.error(f"Σφάλμα: {_e}")

                st.info(
                    f"ℹ️ Όταν ξεκινήσει το {_next_year}, μετονόμασε το `employees_{_next_year}.xlsx` "
                    "σε `employees.xlsx` (από την καρτέλα Ιστορικό → employees) για να γίνει το ενεργό αρχείο. "
                    "Έτσι δεν επηρεάζεται το τρέχον έτος μέχρι να είσαι έτοιμος."
                )
