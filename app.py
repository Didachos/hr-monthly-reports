import io
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

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


def leave_balance_table(leaves: pd.DataFrame, month: int) -> pd.DataFrame:
    rows = []
    for _, r in leaves.iterrows():
        curr_taken = int(r["Κανονική Άδεια από Τρέχον Έτος"])
        curr_total = int(r["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"])
        prev_remaining = int(r["Υπόλοιπο Προηγούμενου Έτους Μετά"])

        row = {
            "Επώνυμο": r["Επώνυμο"],
            "Όνομα": r["Όνομα"],
            "Τρέχον Έτος (χρησιμοποιήθηκαν/σύνολο)": f"{curr_taken}/{curr_total}",
            "Υπόλοιπο Τρέχοντος": int(r["Υπόλοιπο Τρέχοντος Έτους Μετά"]),
        }
        if month <= 3:
            row["Υπόλοιπο Προηγ. Έτους"] = prev_remaining
        rows.append(row)
    return pd.DataFrame(rows)


# =========================
# UI
# =========================

st.set_page_config(page_title="Εργάνη - Απουσίες", page_icon="📋", layout="wide")
st.title("📋 Εργάνη — Διαχείριση Απουσιών")

tab_run, tab_history, tab_balances = st.tabs(["▶ Εκτέλεση", "📁 Ιστορικό", "📊 Υπόλοιπα Αδειών"])


# =========================
# TAB: ΕΚΤΕΛΕΣΗ
# =========================

with tab_run:
    st.subheader("Περίοδος")
    col1, col2 = st.columns(2)
    with col1:
        year = st.number_input("Έτος", min_value=2020, max_value=2100, value=2026, step=1)
    with col2:
        month = st.selectbox("Μήνας", options=list(MONTHS.keys()), format_func=lambda m: MONTHS[m])

    st.subheader("Αρχεία Εισόδου")
    raw_file = st.file_uploader("Αρχείο παρουσίας (.xlsx)", type=["xlsx"])
    employees_file = st.file_uploader("employees.xlsx", type=["xlsx"])

    st.subheader("Ταξινόμηση Απουσιών (προαιρετικό)")
    st.caption("Αν έχεις ήδη συμπληρώσει το classified_absences, ανέβασέ το εδώ για να παραχθεί το πλήρες report.")
    classified_file = st.file_uploader(
        f"classified_absences_{year}_{month:02d}.xlsx",
        type=["xlsx"],
    )

    st.divider()
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

                if classified_file:
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

            if not classified_file:
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

        except Exception as e:
            st.error(f"Σφάλμα: {e}")


# =========================
# TAB: ΙΣΤΟΡΙΚΟ
# =========================

with tab_history:
    st.subheader("Παλαιότερα Αρχεία")

    if not OUTPUT_DIR.exists():
        st.info("Δεν βρέθηκε φάκελος output. Τρέξε πρώτα το script τουλάχιστον μία φορά.")
    else:
        reports = sorted(OUTPUT_DIR.glob("monthly_report_*.xlsx"), reverse=True)
        classified_files = sorted(OUTPUT_DIR.glob("classified_absences_*.xlsx"), reverse=True)
        ergani_files = sorted(OUTPUT_DIR.glob("ergani_export_*.xlsx"), reverse=True)

        if not reports and not classified_files and not ergani_files:
            st.info("Δεν υπάρχουν αρχεία ακόμα.")
        else:
            # Ομαδοποίηση ανά περίοδο (YYYY_MM)
            periods = sorted(
                set(
                    "_".join(f.stem.split("_")[-2:])
                    for f in [*reports, *classified_files, *ergani_files]
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
                    period_reports = [f for f in reports if f.stem.endswith(period)]
                    period_classified = [f for f in classified_files if f.stem.endswith(period)]
                    period_ergani = [f for f in ergani_files if period in f.stem]

                    for f in [*period_reports, *period_classified, *period_ergani]:
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

    if leaves_df is None:
        # Φόρτωσε από το πιο πρόσφατο monthly report
        if OUTPUT_DIR.exists():
            reports = sorted(OUTPUT_DIR.glob("monthly_report_*.xlsx"), reverse=True)
            if reports:
                latest = reports[0]
                try:
                    leaves_df = pd.read_excel(latest, sheet_name="Άδειες")
                    # Εξαγωγή μήνα από το όνομα αρχείου (monthly_report_YYYY_MM.xlsx)
                    parts = latest.stem.split("_")
                    leaves_month = int(parts[-1])
                    st.caption(f"Από: {latest.name}")
                except Exception:
                    pass

    if leaves_df is None:
        st.info("Δεν υπάρχουν δεδομένα. Τρέξε πρώτα μια εκτέλεση.")
    else:
        table = leave_balance_table(leaves_df, leaves_month)

        if leaves_month <= 3:
            st.caption("📌 Εντός Q1 — εμφανίζεται και το υπόλοιπο προηγούμενου έτους.")
        else:
            st.caption("📌 Μετά τον Μάρτιο — το υπόλοιπο προηγούμενου έτους έχει λήξει.")

        # Χρωματισμός υπολοίπου τρέχοντος έτους
        def color_balance(val):
            if isinstance(val, int):
                if val <= 3:
                    return "color: red"
                elif val <= 7:
                    return "color: orange"
                return "color: green"
            return ""

        st.dataframe(
            table.style.map(color_balance, subset=["Υπόλοιπο Τρέχοντος"]),
            use_container_width=True,
            hide_index=True,
        )
