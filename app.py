import base64
import datetime
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
    load_attendance,
    load_classified_absences,
    load_employees,
)

MONTHS = {
    1: "Ιανουάριος", 2: "Φεβρουάριος", 3: "Μάρτιος",
    4: "Απρίλιος", 5: "Μάιος", 6: "Ιούνιος",
    7: "Ιούλιος", 8: "Αύγουστος", 9: "Σεπτέμβριος",
    10: "Οκτώβριος", 11: "Νοέμβριος", 12: "Δεκέμβριος",
}

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

tab_run, tab_history, tab_balances = st.tabs(["▶ Εκτέλεση", "📁 Ιστορικό", "📊 Υπόλοιπα Αδειών"])


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

            if not classified_file and not classified_bytes:
                template = build_classified_absence_template(absences)
                template_bytes = build_classified_template_excel_bytes(absences)
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

                # Αποθήκευση αποτελεσμάτων για tab υπολοίπων
                st.session_state["leaves"] = leaves
                st.session_state["leaves_month"] = month
                st.session_state["leaves_year"] = year

                # Auto-save στο OneDrive αν είναι συνδεδεμένο
                od_token = st.session_state.get("od_token")
                if od_token and (classified_file or classified_bytes):
                    try:
                        with st.spinner("Αποθήκευση στο OneDrive..."):
                            # Αποθήκευση raw attendance στο OneDrive (subfolder: raw)
                            od.upload_file(od_token, f"raw_attendance_{year}_{month:02d}.xlsx", raw_file.getvalue(), subfolder="raw")
                            # Αποθήκευση classified (ώστε μελλοντικά μήνες να έχουν YTD δεδομένα)
                            _cls_save_bytes = classified_file.getvalue() if classified_file else classified_bytes
                            if _cls_save_bytes:
                                od.upload_file(od_token, f"classified_absences_{year}_{month:02d}.xlsx", _cls_save_bytes, subfolder="output")
                            # Αποθήκευση monthly report
                            od.upload_file(od_token, f"monthly_report_{year}_{month:02d}.xlsx", report_bytes)
                            if not ergani_df.empty:
                                for branch_value, branch_df in ergani_df.groupby("ΑΑ Παραρτηματος", dropna=False):
                                    branch_out = branch_df.drop(columns=["ΑΑ Παραρτηματος"]).copy()
                                    branch_label = int(branch_value) if pd.notna(branch_value) else "unknown"
                                    od.upload_file(od_token, f"ergani_export_parartima_{branch_label}_{year}_{month:02d}.xlsx", ergani_excel_bytes(branch_out))
                        st.success("✅ Αποθηκεύτηκε στο OneDrive! (output + raw)")
                    except Exception as e:
                        st.warning(f"⚠️ Δεν ήταν δυνατή η αποθήκευση στο OneDrive: {e}")

        except Exception as e:
            st.error(f"Σφάλμα: {e}")


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

        # Τρέχον έτος
        st.subheader(f"📅 Τρέχον Έτος{f' {leaves_year}' if leaves_year else ''}")
        curr_table = leave_balance_table_current(leaves_df)
        _show_table(curr_table, "Υπόλοιπο", _col_cfg_curr)

        # Προηγούμενο έτος — Ιανουάριος έως Απρίλιος
        if 1 <= leaves_month <= 4:
            prev_table = leave_balance_table_prev(leaves_df)
            has_prev_balance = prev_table["Υπόλοιπο"].sum() > 0
            if has_prev_balance:
                prev_year = int(leaves_year) - 1 if leaves_year else ""
                st.subheader(f"📅 Προηγούμενο Έτος{f' {prev_year}' if prev_year else ''}")
                st.caption("⚠️ Το υπόλοιπο λήγει στο τέλος Μαρτίου.")
                _show_table(prev_table, "Υπόλοιπο", _col_cfg_prev)

        # Δεκέμβριος — προεπισκόπηση υπολοίπου που μεταφέρεται στο επόμενο έτος
        if leaves_month == 12:
            curr_table_dec = leave_balance_table_current(leaves_df)
            carryover = curr_table_dec[curr_table_dec["Υπόλοιπο"] > 0]
            if not carryover.empty:
                next_year = int(leaves_year) + 1 if leaves_year else ""
                st.subheader(f"📅 Μεταφορά υπολοίπου στο {next_year}")
                st.caption("Οι παρακάτω εργαζόμενοι έχουν υπόλοιπο που μεταφέρεται στο νέο έτος (λήγει τέλος Μαρτίου).")
                _show_table(carryover, "Υπόλοιπο", _col_cfg_curr)
