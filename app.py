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
    build_alerts_report,
    build_classified_absence_template,
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


def leave_balance_table_current(leaves: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in leaves.iterrows():
        curr_taken = int(r["Κανονική Άδεια από Τρέχον Έτος"])
        curr_total = int(r["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"])
        rows.append({
            "Επώνυμο": r["Επώνυμο"],
            "Όνομα": r["Όνομα"],
            "Χρησιμοποιήθηκαν/Σύνολο": f"{curr_taken}/{curr_total}",
            "Υπόλοιπο": int(r["Υπόλοιπο Τρέχοντος Έτους Μετά"]),
        })
    return pd.DataFrame(rows)


def leave_balance_table_prev(leaves: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in leaves.iterrows():
        prev_taken = int(r["Κανονική Άδεια από Προηγούμενο Έτος"])
        prev_available = int(r["Υπόλοιπο Προηγούμενου Έτους"])
        rows.append({
            "Επώνυμο": r["Επώνυμο"],
            "Όνομα": r["Όνομα"],
            "Χρησιμοποιήθηκαν/Διαθέσιμο": f"{prev_taken}/{prev_available}",
            "Υπόλοιπο": int(r["Υπόλοιπο Προηγούμενου Έτους Μετά"]),
        })
    return pd.DataFrame(rows)


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
    raw_file = st.file_uploader("Αρχείο παρουσίας (.xlsx)", type=["xlsx"])
    employees_file = st.file_uploader("employees.xlsx", type=["xlsx"])

    st.subheader("Ταξινόμηση Απουσιών (προαιρετικό)")
    st.caption("Αν έχεις ήδη συμπληρώσει το classified_absences, ανέβασέ το εδώ για να παραχθεί το πλήρες report.")

    classified_file = None
    classified_bytes = None

    # Φόρτωση classified από OneDrive αν υπάρχει
    od_token_cls = st.session_state.get("od_token")
    cls_filename = f"classified_absences_{year}_{month:02d}.xlsx"
    if od_token_cls:
        try:
            od_files = od.list_files(od_token_cls, subfolder="output")
            if any(f["name"] == cls_filename for f in od_files):
                st.info(f"☁️ Βρέθηκε `{cls_filename}` στο OneDrive.")
                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button("✅ Χρησιμοποίησε από OneDrive"):
                        classified_bytes = od.download_file(od_token_cls, cls_filename, subfolder="output")
                        st.success("Φορτώθηκε από OneDrive!")
                with col_b:
                    classified_file = st.file_uploader("Ή ανέβασε νέο αρχείο", type=["xlsx"], key="cls_upload")
            else:
                classified_file = st.file_uploader(cls_filename, type=["xlsx"], key="cls_upload")
        except Exception:
            classified_file = st.file_uploader(cls_filename, type=["xlsx"], key="cls_upload")
    else:
        classified_file = st.file_uploader(cls_filename, type=["xlsx"], key="cls_upload")

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

    run = st.button("▶ Εκτέλεση", type="primary", disabled=not (raw_file and employees_file))

    if run:
        try:
            with st.spinner("Επεξεργασία..."):
                raw_path = save_upload_to_temp(raw_file)
                emp_path = save_upload_to_temp(employees_file)

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

                leaves = build_leave_summary(classified, employees, year, month)
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
                template_bytes = excel_bytes({"Sheet1": template})
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
                            data=excel_bytes({"Άδειες": branch_out}),
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
                            # Αποθήκευση monthly report
                            od.upload_file(od_token, f"monthly_report_{year}_{month:02d}.xlsx", report_bytes)
                            if not ergani_df.empty:
                                for branch_value, branch_df in ergani_df.groupby("ΑΑ Παραρτηματος", dropna=False):
                                    branch_out = branch_df.drop(columns=["ΑΑ Παραρτηματος"]).copy()
                                    branch_label = int(branch_value) if pd.notna(branch_value) else "unknown"
                                    od.upload_file(od_token, f"ergani_export_parartima_{branch_label}_{year}_{month:02d}.xlsx", excel_bytes({"Άδειες": branch_out}))
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
        st.divider()

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

    # Από την τελευταία εκτέλεση (session) ή από το τελευταίο αποθηκευμένο report
    leaves_df = st.session_state.get("leaves")
    leaves_month = st.session_state.get("leaves_month", 12)
    leaves_year = st.session_state.get("leaves_year", "")

    if leaves_df is None:
        od_token = st.session_state.get("od_token")
        if od_token:
            # Φόρτωσε το πιο πρόσφατο monthly report από OneDrive
            try:
                files = od.list_files(od_token, subfolder="output")
                report_files = sorted(
                    [f["name"] for f in files if f["name"].startswith("monthly_report_") and f["name"].endswith(".xlsx")],
                    reverse=True,
                )
                if report_files:
                    latest_name = report_files[0]
                    content = od.download_file(od_token, latest_name, subfolder="output")
                    leaves_df = pd.read_excel(io.BytesIO(content), sheet_name="Άδειες")
                    parts = latest_name.replace(".xlsx", "").split("_")
                    leaves_month = int(parts[-1])
                    leaves_year = int(parts[-2])
                    st.caption(f"📂 Από OneDrive: {latest_name}")
            except Exception:
                pass

    if leaves_df is None:
        # Fallback: τοπικά αρχεία
        if OUTPUT_DIR.exists():
            reports = sorted(OUTPUT_DIR.glob("monthly_report_*.xlsx"), reverse=True)
            if reports:
                latest = reports[0]
                try:
                    leaves_df = pd.read_excel(latest, sheet_name="Άδειες")
                    parts = latest.stem.split("_")
                    leaves_month = int(parts[-1])
                    leaves_year = int(parts[-2])
                    st.caption(f"Από: {latest.name}")
                except Exception:
                    pass

    if leaves_df is None:
        st.info("Δεν υπάρχουν δεδομένα. Τρέξε πρώτα μια εκτέλεση.")
    else:
        def color_balance(val):
            if isinstance(val, int):
                if val <= 3:
                    return "color: red"
                elif val <= 7:
                    return "color: orange"
                return "color: green"
            return ""

        # Τρέχον έτος
        st.subheader(f"📅 Τρέχον Έτος{f' {leaves_year}' if leaves_year else ''}")
        curr_table = leave_balance_table_current(leaves_df)
        st.dataframe(
            curr_table.style.map(color_balance, subset=["Υπόλοιπο"]),
            use_container_width=True,
            hide_index=True,
        )

        # Προηγούμενο έτος — μόνο αν είναι Q1 και υπάρχει υπόλοιπο
        if leaves_month <= 3:
            prev_table = leave_balance_table_prev(leaves_df)
            has_prev_balance = prev_table["Υπόλοιπο"].sum() > 0
            if has_prev_balance:
                prev_year = int(leaves_year) - 1 if leaves_year else ""
                st.subheader(f"📅 Προηγούμενο Έτος{f' {prev_year}' if prev_year else ''}")
                st.caption("⚠️ Το υπόλοιπο λήγει στο τέλος Μαρτίου.")
                st.dataframe(
                    prev_table.style.map(color_balance, subset=["Υπόλοιπο"]),
                    use_container_width=True,
                    hide_index=True,
                )
