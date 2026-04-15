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
# UI
# =========================

st.set_page_config(page_title="Εργάνη - Απουσίες", page_icon="📋", layout="centered")
st.title("📋 Εργάνη — Διαχείριση Απουσιών")

# --- Επιλογή περιόδου ---
st.subheader("Περίοδος")
col1, col2 = st.columns(2)
with col1:
    year = st.number_input("Έτος", min_value=2020, max_value=2100, value=2026, step=1)
with col2:
    month = st.selectbox("Μήνας", options=list(MONTHS.keys()), format_func=lambda m: MONTHS[m])

# --- Αρχεία εισόδου ---
st.subheader("Αρχεία Εισόδου")
raw_file = st.file_uploader("raw_attendance.xlsx", type=["xlsx"])
employees_file = st.file_uploader("employees.xlsx", type=["xlsx"])

st.subheader("Ταξινόμηση Απουσιών (προαιρετικό)")
st.caption("Αν έχεις ήδη συμπληρώσει το classified_absences, ανέβασέ το εδώ για να παραχθεί το πλήρες report.")
classified_file = st.file_uploader(
    f"classified_absences_{year}_{month:02d}.xlsx",
    type=["xlsx"],
)

# --- Εκτέλεση ---
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

            leaves = build_leave_summary(classified, employees, year)
            validation = build_validation_report(raw_df, df, employees, absences, classified, year, month)
            alerts = build_alerts_report(employees, absences, classified, workdays, overtime_s, leaves, year)
            ergani_df = build_ergani_export_df(classified, employees, year)

        # --- Αποτελέσματα ---
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
            # Πρώτη εκτέλεση — μόνο template
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
            # Δεύτερη εκτέλεση — πλήρες report
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
                label=f"⬇ Κατέβασε monthly_report_{year}_{month:02d}.xlsx",
                data=report_bytes,
                file_name=f"monthly_report_{year}_{month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Ergani exports ανά παράρτημα
            if not ergani_df.empty:
                st.markdown("**Ergani exports ανά παράρτημα:**")
                for branch_value, branch_df in ergani_df.groupby("ΑΑ Παραρτηματος", dropna=False):
                    branch_out = branch_df.drop(columns=["ΑΑ Παραρτηματος"]).copy()
                    branch_bytes = excel_bytes({"Άδειες": branch_out})
                    st.download_button(
                        label=f"⬇ Ergani — Παράρτημα {int(branch_value) if pd.notna(branch_value) else 'unknown'}",
                        data=branch_bytes,
                        file_name=f"ergani_export_parartima_{int(branch_value) if pd.notna(branch_value) else 'unknown'}_{year}_{month:02d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    except Exception as e:
        st.error(f"Σφάλμα: {e}")
