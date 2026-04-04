from pathlib import Path
from calendar import monthrange
import sys
import pandas as pd
import holidays


STANDARD_WORK_MINUTES = 8 * 60

VALID_ABSENCE_TYPES = {
    "Κανονική άδεια",
    "Άδεια ασθενείας",
    "Άνευ αποδοχών άδεια",
}


# =========================
# HELPERS
# =========================

def force_text_column(worksheet, header_name: str) -> None:
    target_col_idx = None

    for cell in worksheet[1]:
        if cell.value == header_name:
            target_col_idx = cell.column
            break

    if target_col_idx is None:
        return

    for row in worksheet.iter_rows(
        min_row=2,
        min_col=target_col_idx,
        max_col=target_col_idx
    ):
        cell = row[0]
        if cell.value is not None:
            cell.value = str(cell.value)
            cell.number_format = "@"


def safe_str_series(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


def minutes_to_hhmm(minutes: int) -> str:
    minutes = int(minutes)
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"


def to_minutes(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None

    parsed = pd.to_datetime(text, format="%H:%M", errors="coerce")
    if pd.isna(parsed):
        parsed = pd.to_datetime(text, errors="coerce")

    if pd.isna(parsed):
        return None

    return parsed.hour * 60 + parsed.minute


# =========================
# LOADERS
# =========================

def load_attendance(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"Δεν βρέθηκε το raw αρχείο: {file_path}")

    df = pd.read_excel(file_path, dtype={"ΑΦΜ": str})
    df.columns = [str(c).strip() for c in df.columns]

    required_columns = [
        "ΑΑ Παραρτηματος",
        "ΑΦΜ",
        "Επώνυμο",
        "Όνομα",
        "Ημ/νία",
        "Από",
        "Έως",
    ]

    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise ValueError(f"Λείπουν στήλες από το raw_attendance.xlsx: {missing}")

    return df


def load_employees(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"Δεν βρέθηκε το employees.xlsx: {file_path}")

    df = pd.read_excel(file_path, dtype={"ΑΦΜ": str})
    df.columns = [str(c).strip() for c in df.columns]

    required_columns = [
        "ΑΑ Παραρτηματος",
        "ΑΦΜ",
        "Επώνυμο",
        "Όνομα",
        "Ημερομηνία Πρόσληψης",
        "Ημερομηνία Αποχώρησης",
        "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους",
        "Υπόλοιπο Προηγούμενου Έτους",
        "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους",
    ]

    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise ValueError(f"Λείπουν στήλες από το employees.xlsx: {missing}")

    df["ΑΑ Παραρτηματος"] = pd.to_numeric(df["ΑΑ Παραρτηματος"], errors="coerce")
    df["ΑΦΜ"] = safe_str_series(df["ΑΦΜ"])
    df["Επώνυμο"] = safe_str_series(df["Επώνυμο"])
    df["Όνομα"] = safe_str_series(df["Όνομα"])

    df["Ημερομηνία Πρόσληψης"] = pd.to_datetime(
        df["Ημερομηνία Πρόσληψης"],
        dayfirst=True,
        errors="coerce"
    ).dt.normalize()

    df["Ημερομηνία Αποχώρησης"] = pd.to_datetime(
        df["Ημερομηνία Αποχώρησης"],
        dayfirst=True,
        errors="coerce"
    ).dt.normalize()

    df["Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους"] = pd.to_numeric(
        df["Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους"], errors="coerce"
    ).fillna(0).astype(int)

    df["Υπόλοιπο Προηγούμενου Έτους"] = pd.to_numeric(
        df["Υπόλοιπο Προηγούμενου Έτους"], errors="coerce"
    ).fillna(0).astype(int)

    df["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"] = pd.to_numeric(
        df["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"], errors="coerce"
    ).fillna(0).astype(int)

    if df["Ημερομηνία Πρόσληψης"].isna().any():
        missing_afm = df.loc[df["Ημερομηνία Πρόσληψης"].isna(), "ΑΦΜ"].tolist()
        raise ValueError(
            f"Υπάρχουν εργαζόμενοι χωρίς Ημερομηνία Πρόσληψης: {missing_afm}"
        )

    return df.drop_duplicates().reset_index(drop=True)


def load_classified_absences(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        return pd.DataFrame()

    df = pd.read_excel(file_path, dtype={"ΑΦΜ": str})
    df.columns = [str(c).strip() for c in df.columns]

    required_columns = [
        "ΑΑ Παραρτηματος",
        "ΑΦΜ",
        "Επώνυμο",
        "Όνομα",
        "Ημ/νία",
        "Τύπος Απουσίας",
        "Έτος Άδειας",
    ]

    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise ValueError(
            f"Λείπουν στήλες από το classified_absences αρχείο: {missing}"
        )

    df["ΑΑ Παραρτηματος"] = pd.to_numeric(df["ΑΑ Παραρτηματος"], errors="coerce")
    df["ΑΦΜ"] = safe_str_series(df["ΑΦΜ"])
    df["Επώνυμο"] = safe_str_series(df["Επώνυμο"])
    df["Όνομα"] = safe_str_series(df["Όνομα"])
    df["Ημ/νία"] = pd.to_datetime(
        df["Ημ/νία"], dayfirst=True, errors="coerce"
    ).dt.normalize()

    df["Τύπος Απουσίας"] = df["Τύπος Απουσίας"].fillna("").astype(str).str.strip()
    df["Έτος Άδειας"] = pd.to_numeric(df["Έτος Άδειας"], errors="coerce")

    df = df.dropna(subset=["ΑΦΜ", "Ημ/νία"]).copy()

    filled = df[df["Τύπος Απουσίας"] != ""].copy()

    invalid_types = sorted(set(filled["Τύπος Απουσίας"]) - VALID_ABSENCE_TYPES)
    if invalid_types:
        raise ValueError(
            f"Μη αποδεκτοί τύποι απουσίας: {invalid_types}. "
            f"Επιτρεπτές τιμές: {sorted(VALID_ABSENCE_TYPES)}"
        )

    annual_leave_missing_year = filled[
        (filled["Τύπος Απουσίας"] == "Κανονική άδεια") &
        (filled["Έτος Άδειας"].isna())
    ]
    if not annual_leave_missing_year.empty:
        raise ValueError(
            "Υπάρχουν γραμμές με 'Κανονική άδεια' χωρίς συμπληρωμένο 'Έτος Άδειας'."
        )

    filled.loc[
        filled["Τύπος Απουσίας"] != "Κανονική άδεια",
        "Έτος Άδειας"
    ] = pd.NA

    return filled.drop_duplicates().reset_index(drop=True)


# =========================
# CLEAN
# =========================

def clean_attendance(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df["ΑΑ Παραρτηματος"] = pd.to_numeric(df["ΑΑ Παραρτηματος"], errors="coerce")
    df["ΑΦΜ"] = safe_str_series(df["ΑΦΜ"])
    df["Επώνυμο"] = safe_str_series(df["Επώνυμο"])
    df["Όνομα"] = safe_str_series(df["Όνομα"])
    df["Ημ/νία"] = pd.to_datetime(
        df["Ημ/νία"], dayfirst=True, errors="coerce"
    ).dt.normalize()
    df["Από"] = df["Από"].astype(str).str.strip()
    df["Έως"] = df["Έως"].astype(str).str.strip()

    df = df.dropna(subset=["ΑΦΜ", "Ημ/νία"])
    df = df.drop_duplicates()

    return df


# =========================
# ABSENCES
# =========================

def build_month_dates(year: int, month: int):
    last = monthrange(year, month)[1]
    return pd.date_range(
        start=f"{year}-{month:02d}-01",
        end=f"{year}-{month:02d}-{last:02d}",
        freq="D"
    ).normalize()


def get_holidays(year: int):
    return {
        pd.Timestamp(d).normalize()
        for d in holidays.country_holidays("GR", years=year).keys()
    }


def find_absences(
    df: pd.DataFrame,
    employees: pd.DataFrame,
    year: int,
    month: int
) -> pd.DataFrame:
    month_df = df[
        (df["Ημ/νία"].dt.year == year) &
        (df["Ημ/νία"].dt.month == month)
    ].copy()

    employees_unique = employees[
        [
            "ΑΑ Παραρτηματος",
            "ΑΦΜ",
            "Επώνυμο",
            "Όνομα",
            "Ημερομηνία Πρόσληψης",
            "Ημερομηνία Αποχώρησης",
        ]
    ].drop_duplicates()

    dates = build_month_dates(year, month)

    full = employees_unique.assign(key=1).merge(
        pd.DataFrame({"Ημ/νία": dates, "key": 1}),
        on="key"
    ).drop(columns="key")

    holidays_set = get_holidays(year)

    full = full[
        (full["Ημ/νία"].dt.weekday < 5) &
        (~full["Ημ/νία"].isin(holidays_set))
    ].copy()

    full = full[full["Ημ/νία"] >= full["Ημερομηνία Πρόσληψης"]]
    full = full[
        full["Ημερομηνία Αποχώρησης"].isna() |
        (full["Ημ/νία"] <= full["Ημερομηνία Αποχώρησης"])
    ]

    present = month_df[["ΑΦΜ", "Ημ/νία"]].drop_duplicates()
    present["present"] = 1

    merged = full.merge(present, on=["ΑΦΜ", "Ημ/νία"], how="left")

    absences = merged[merged["present"].isna()].copy()
    absences["Κατάσταση"] = "ΑΠΩΝ"
    absences["Ημ/νία"] = pd.to_datetime(absences["Ημ/νία"]).dt.strftime("%d/%m/%Y")

    result = absences[
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία", "Κατάσταση"]
    ].sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ", "Ημ/νία"]).reset_index(drop=True)

    result["ΑΦΜ"] = result["ΑΦΜ"].astype(str)
    return result


def build_classified_absence_template(absences: pd.DataFrame) -> pd.DataFrame:
    template = absences[
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία"]
    ].copy()

    template["Τύπος Απουσίας"] = ""
    template["Έτος Άδειας"] = ""
    template["ΑΦΜ"] = template["ΑΦΜ"].astype(str)

    return template


# =========================
# WORK DAYS
# =========================

def calculate_work_days(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    month_df = df[
        (df["Ημ/νία"].dt.year == year) &
        (df["Ημ/νία"].dt.month == month)
    ].copy()

    unique = month_df[
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία"]
    ].drop_duplicates()

    result = (
        unique.groupby(
            ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"],
            as_index=False
        )
        .agg(**{"Σύνολο Ημερών Εργασίας": ("Ημ/νία", "nunique")})
        .sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ"])
        .reset_index(drop=True)
    )

    result["ΑΦΜ"] = result["ΑΦΜ"].astype(str)
    return result


# =========================
# OVERTIME
# =========================

def calculate_overtime(df: pd.DataFrame, year: int, month: int):
    month_df = df[
        (df["Ημ/νία"].dt.year == year) &
        (df["Ημ/νία"].dt.month == month)
    ].copy()

    month_df["start"] = month_df["Από"].apply(to_minutes)
    month_df["end"] = month_df["Έως"].apply(to_minutes)

    month_df = month_df.dropna(subset=["start", "end"]).copy()

    month_df["worked"] = month_df["end"] - month_df["start"]
    month_df.loc[month_df["worked"] < 0, "worked"] += 1440

    month_df["Υπερωρία Λεπτά"] = (
        month_df["worked"] - STANDARD_WORK_MINUTES
    ).clip(lower=0)

    month_df["Υπερωρία"] = month_df["Υπερωρία Λεπτά"].apply(
        lambda x: "ΝΑΙ" if x > 0 else ""
    )
    month_df["Συνολική Διάρκεια"] = month_df["worked"].apply(minutes_to_hhmm)
    month_df["Υπερωρία (HH:MM)"] = month_df["Υπερωρία Λεπτά"].apply(minutes_to_hhmm)

    detailed = month_df[
        [
            "ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία",
            "Από", "Έως", "Συνολική Διάρκεια", "Υπερωρία",
            "Υπερωρία Λεπτά", "Υπερωρία (HH:MM)"
        ]
    ].copy()

    detailed["Ημ/νία"] = pd.to_datetime(detailed["Ημ/νία"]).dt.strftime("%d/%m/%Y")
    detailed = detailed.sort_values(
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Ημ/νία"]
    ).reset_index(drop=True)
    detailed["ΑΦΜ"] = detailed["ΑΦΜ"].astype(str)

    summary = (
        month_df.groupby(
            ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"],
            as_index=False
        )
        .agg(**{"Σύνολο Υπερωρίας Λεπτά": ("Υπερωρία Λεπτά", "sum")})
        .sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ"])
        .reset_index(drop=True)
    )

    summary["Σύνολο Υπερωρίας (HH:MM)"] = summary["Σύνολο Υπερωρίας Λεπτά"].apply(minutes_to_hhmm)
    summary["ΑΦΜ"] = summary["ΑΦΜ"].astype(str)

    return detailed, summary


# =========================
# LEAVES
# =========================

def build_leave_summary(
    classified: pd.DataFrame,
    employees: pd.DataFrame,
    year: int
) -> pd.DataFrame:
    result = employees.copy()
    prev = year - 1

    if classified.empty:
        result["Κανονική Άδεια από Προηγούμενο Έτος"] = 0
        result["Κανονική Άδεια από Τρέχον Έτος"] = 0
        result["Σύνολο Κανονικής Άδειας"] = 0
        result["Σύνολο Ασθενείας"] = 0
        result["Σύνολο Άνευ Αποδοχών"] = 0
        result["Υπόλοιπο Προηγούμενου Έτους Μετά"] = result["Υπόλοιπο Προηγούμενου Έτους"]
        result["Υπόλοιπο Τρέχοντος Έτους Μετά"] = result["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"]
    else:
        annual_prev = classified[
            (classified["Τύπος Απουσίας"] == "Κανονική άδεια") &
            (classified["Έτος Άδειας"] == prev)
        ].groupby("ΑΦΜ").size()

        annual_curr = classified[
            (classified["Τύπος Απουσίας"] == "Κανονική άδεια") &
            (classified["Έτος Άδειας"] == year)
        ].groupby("ΑΦΜ").size()

        sick = classified[
            classified["Τύπος Απουσίας"] == "Άδεια ασθενείας"
        ].groupby("ΑΦΜ").size()

        unpaid = classified[
            classified["Τύπος Απουσίας"] == "Άνευ αποδοχών άδεια"
        ].groupby("ΑΦΜ").size()

        result["Κανονική Άδεια από Προηγούμενο Έτος"] = result["ΑΦΜ"].map(annual_prev).fillna(0).astype(int)
        result["Κανονική Άδεια από Τρέχον Έτος"] = result["ΑΦΜ"].map(annual_curr).fillna(0).astype(int)
        result["Σύνολο Ασθενείας"] = result["ΑΦΜ"].map(sick).fillna(0).astype(int)
        result["Σύνολο Άνευ Αποδοχών"] = result["ΑΦΜ"].map(unpaid).fillna(0).astype(int)

        result["Σύνολο Κανονικής Άδειας"] = (
            result["Κανονική Άδεια από Προηγούμενο Έτος"] +
            result["Κανονική Άδεια από Τρέχον Έτος"]
        )

        result["Υπόλοιπο Προηγούμενου Έτους Μετά"] = (
            result["Υπόλοιπο Προηγούμενου Έτους"] -
            result["Κανονική Άδεια από Προηγούμενο Έτος"]
        ).clip(lower=0)

        result["Υπόλοιπο Τρέχοντος Έτους Μετά"] = (
            result["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"] -
            result["Κανονική Άδεια από Τρέχον Έτος"]
        ).clip(lower=0)

    result = result[
        [
            "ΑΑ Παραρτηματος",
            "ΑΦΜ",
            "Επώνυμο",
            "Όνομα",
            "Ημερομηνία Πρόσληψης",
            "Ημερομηνία Αποχώρησης",
            "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους",
            "Υπόλοιπο Προηγούμενου Έτους",
            "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους",
            "Κανονική Άδεια από Προηγούμενο Έτος",
            "Κανονική Άδεια από Τρέχον Έτος",
            "Σύνολο Κανονικής Άδειας",
            "Σύνολο Ασθενείας",
            "Σύνολο Άνευ Αποδοχών",
            "Υπόλοιπο Προηγούμενου Έτους Μετά",
            "Υπόλοιπο Τρέχοντος Έτους Μετά",
        ]
    ].sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ"]).reset_index(drop=True)

    result["Ημερομηνία Πρόσληψης"] = pd.to_datetime(
        result["Ημερομηνία Πρόσληψης"]
    ).dt.strftime("%d/%m/%Y")

    result["Ημερομηνία Αποχώρησης"] = result["Ημερομηνία Αποχώρησης"].apply(
        lambda x: "" if pd.isna(x) else pd.to_datetime(x).strftime("%d/%m/%Y")
    )

    result["ΑΦΜ"] = result["ΑΦΜ"].astype(str)
    return result


# =========================
# MAIN
# =========================

def main():
    if len(sys.argv) != 3:
        raise ValueError("Χρήση: python3 src/main.py <year> <month>")

    year = int(sys.argv[1])
    month = int(sys.argv[2])

    root = Path(__file__).resolve().parent.parent

    raw = root / "data/input/raw_attendance.xlsx"
    employees_file = root / "data/input/employees.xlsx"
    classified_file = root / f"data/output/classified_absences_{year}_{month:02d}.xlsx"
    output = root / f"data/output/monthly_report_{year}_{month:02d}.xlsx"

    df = clean_attendance(load_attendance(raw))
    employees = load_employees(employees_file)

    absences = find_absences(df, employees, year, month)

    if not classified_file.exists():
        template = build_classified_absence_template(absences)
        classified_file.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(classified_file, engine="openpyxl") as writer:
            template.to_excel(writer, sheet_name="Sheet1", index=False)
            force_text_column(writer.sheets["Sheet1"], "ΑΦΜ")

        print("Δημιουργήθηκε template:", classified_file)

    classified = load_classified_absences(classified_file)

    workdays = calculate_work_days(df, year, month)
    overtime_d, overtime_s = calculate_overtime(df.copy(), year, month)
    leaves = build_leave_summary(classified, employees, year)

    output.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        absences.to_excel(writer, sheet_name="Απουσίες", index=False)
        workdays.to_excel(writer, sheet_name="Ημέρες", index=False)
        overtime_d.to_excel(writer, sheet_name="Υπερωρίες", index=False)
        overtime_s.to_excel(writer, sheet_name="Σύνολο Υπερωρίας", index=False)
        leaves.to_excel(writer, sheet_name="Άδειες", index=False)

        force_text_column(writer.sheets["Απουσίες"], "ΑΦΜ")
        force_text_column(writer.sheets["Ημέρες"], "ΑΦΜ")
        force_text_column(writer.sheets["Υπερωρίες"], "ΑΦΜ")
        force_text_column(writer.sheets["Σύνολο Υπερωρίας"], "ΑΦΜ")
        force_text_column(writer.sheets["Άδειες"], "ΑΦΜ")

    print("Έτοιμο:", output)
    print("Template αδειών:", classified_file)


if __name__ == "__main__":
    main()