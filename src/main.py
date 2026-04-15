from pathlib import Path
from calendar import monthrange
import sys
import pandas as pd
import holidays

STANDARD_WORK_MINUTES = 8 * 60
OVERWORK_LIMIT_MINUTES = 1 * 60


LOW_CURRENT_LEAVE_BALANCE_THRESHOLD = 3
HIGH_ABSENCE_COUNT_THRESHOLD = 5
HIGH_OVERWORK_MINUTES_THRESHOLD = 5 * 60
HIGH_OVERTIME_MINUTES_THRESHOLD = 8 * 60

VALID_ABSENCE_TYPES = {
    "Κανονική άδεια",
    "Άδεια ασθενείας",
    "Άνευ αποδοχών άδεια",
}

ERGANI_LEAVE_TYPES = {
    "Κανονική άδεια",
    "Αιμοδοτική άδεια",
    "Άδεια εξετάσεων",
    "Άδεια άνευ αποδοχών",
    "Άδεια μητρότητας",
    "Ειδική παροχή προστασίας της μητρότητας",
    "Άδεια πατρότητας",
    "Άδεια φροντίδας παιδιού",
    "Γονική άδεια",
    "Άδεια φροντιστή",
    "Απουσία από την εργασία για λόγους ανωτέρας βίας",
    "Άδεια για υποβολή σε μεθόδους ιατρικώς υποβοηθούμενης αναπαραγωγής",
    "Άδεια εξετάσεων προγεννητικού ελέγχου",
    "Άδεια γάμου",
    "Άδεια λόγω σοβαρών νοσημάτων των παιδιών",
    "Άδεια λόγω νοσηλείας των παιδιών",
    "Άδεια μονογονεϊκών οικογενειών",
    "Άδεια παρακολούθησης σχολικής επίδοσης τέκνου",
    "Άδεια λόγω ασθένειας παιδιού ή άλλου εξαρτώμενου μέλους",
    "Απουσία από την εργασία λόγω επικείμενου σοβαρού κινδύνου βίας ή παρενόχλησης",
    "Άδεια ασθένειας (ανυπαίτιο κώλυμα παροχής εργασίας)",
    "Άδεια απουσίας Α.Μ.Ε.Α.",
    "Άδεια λόγω θανάτου συγγενούς",
    "Άδεια ανήλικων σπουδαστών",
    "Άδεια για μεταγγίσεις αίματος και των παραγώγων του ή αιμοκάθαρση",
    "Εκπαιδευτική άδεια για φοιτητές στο Κ.ΑΝ.Ε.Π. – Γ.Σ.Ε.Ε.",
    "Άδεια λόγω AIDS",
    "Ευέλικτες ρυθμίσεις εργασίας",
    "Άδεια φροντίδας παιδιού (ΩΡΕΣ)",
    "Γονική άδεια (ΩΡΕΣ)",
    "Απουσία από την εργασία για λόγους ανωτέρας βίας (ΩΡΕΣ)",
    "Ευέλικτες ρυθμίσεις εργασίας (ΩΡΕΣ)",
    "Άδεια εξετάσεων προγεννητικού ελέγχου (ΩΡΕΣ)",
    "Άδεια παρακολούθησης σχολικής επίδοσης τέκνου (ΩΡΕΣ)",
    "Άδεια Άλλη",
    "Άδεια Άλλη (ΩΡΕΣ)",
}

CLASSIFIED_TO_ERGANI_LEAVE_TYPE = {
    "Κανονική άδεια": "Κανονική άδεια",
    "Άδεια ασθενείας": "Άδεια ασθένειας (ανυπαίτιο κώλυμα παροχής εργασίας)",
    "Άνευ αποδοχών άδεια": "Άδεια άνευ αποδοχών",
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


def make_validation_row(
    level: str,
    category: str,
    message: str,
    aa=None,
    afm=None,
    last_name=None,
    first_name=None,
    date=None,
    value=None,
):
    return {
        "Επίπεδο": level,
        "Κατηγορία": category,
        "Μήνυμα": message,
        "ΑΑ Παραρτηματος": aa if aa is not None else "",
        "ΑΦΜ": "" if afm is None else str(afm),
        "Επώνυμο": last_name if last_name is not None else "",
        "Όνομα": first_name if first_name is not None else "",
        "Ημ/νία": date if date is not None else "",
        "Τιμή": value if value is not None else "",
    }


def format_validation_date(value):
    if pd.isna(value) or value == "":
        return ""
    return pd.to_datetime(value).strftime("%d/%m/%Y")


def branch_to_filename_part(value) -> str:
    if pd.isna(value):
        return "unknown"

    try:
        numeric = float(value)
        if numeric.is_integer():
            return str(int(numeric))
    except Exception:
        pass

    text = str(value).strip()
    return text.replace("/", "_").replace("\\", "_").replace(" ", "_")


def get_entitled_days(row, year: int):
    """Επιστρέφει τις δικαιούμενες ημέρες κανονικής άδειας για το έτος της γραμμής, ή None αν δεν εφαρμόζεται."""
    if row["Τύπος Απουσίας"] != "Κανονική άδεια":
        return None
    leave_year = row["Έτος Άδειας"]
    if pd.isna(leave_year):
        return None
    leave_year = int(leave_year)
    if leave_year == year - 1:
        return int(row["Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους"])
    if leave_year == year:
        return int(row["Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους"])
    return None


def make_alert_row(
    level: str,
    category: str,
    aa=None,
    afm=None,
    last_name=None,
    first_name=None,
    message: str = "",
    value=None,
):
    return {
        "Επίπεδο": level,
        "Κατηγορία": category,
        "ΑΑ Παραρτηματος": aa if aa is not None else "",
        "ΑΦΜ": "" if afm is None else str(afm),
        "Επώνυμο": last_name if last_name is not None else "",
        "Όνομα": first_name if first_name is not None else "",
        "Μήνυμα": message,
        "Τιμή": value if value is not None else "",
    }


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

    month_df["Extra Λεπτά"] = (
        month_df["worked"] - STANDARD_WORK_MINUTES
    ).clip(lower=0)

    month_df["Υπεργασία Λεπτά"] = month_df["Extra Λεπτά"].clip(upper=OVERWORK_LIMIT_MINUTES)
    month_df["Υπερωρία Λεπτά"] = (
        month_df["Extra Λεπτά"] - month_df["Υπεργασία Λεπτά"]
    ).clip(lower=0)

    month_df["Υπεργασία"] = month_df["Υπεργασία Λεπτά"].apply(
        lambda x: "ΝΑΙ" if x > 0 else ""
    )
    month_df["Υπερωρία"] = month_df["Υπερωρία Λεπτά"].apply(
        lambda x: "ΝΑΙ" if x > 0 else ""
    )

    month_df["Συνολική Διάρκεια"] = month_df["worked"].apply(minutes_to_hhmm)
    month_df["Υπεργασία (HH:MM)"] = month_df["Υπεργασία Λεπτά"].apply(minutes_to_hhmm)
    month_df["Υπερωρία (HH:MM)"] = month_df["Υπερωρία Λεπτά"].apply(minutes_to_hhmm)

    detailed = month_df[
        [
            "ΑΑ Παραρτηματος",
            "ΑΦΜ",
            "Επώνυμο",
            "Όνομα",
            "Ημ/νία",
            "Από",
            "Έως",
            "Συνολική Διάρκεια",
            "Υπεργασία",
            "Υπεργασία Λεπτά",
            "Υπεργασία (HH:MM)",
            "Υπερωρία",
            "Υπερωρία Λεπτά",
            "Υπερωρία (HH:MM)",
        ]
    ].copy()

    detailed = detailed.sort_values(
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Ημ/νία"]
    ).reset_index(drop=True)
    detailed["ΑΦΜ"] = detailed["ΑΦΜ"].astype(str)

    summary = (
        month_df.groupby(
            ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"],
            as_index=False
        )
        .agg(
            **{
                "Σύνολο Υπεργασίας Λεπτά": ("Υπεργασία Λεπτά", "sum"),
                "Σύνολο Υπερωρίας Λεπτά": ("Υπερωρία Λεπτά", "sum"),
            }
        )
        .sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ"])
        .reset_index(drop=True)
    )

    summary["Σύνολο Υπεργασίας (HH:MM)"] = summary["Σύνολο Υπεργασίας Λεπτά"].apply(minutes_to_hhmm)
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
        classified = pd.DataFrame(columns=["ΑΦΜ", "Τύπος Απουσίας", "Έτος Άδειας"])

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
# ERGANI EXPORT
# =========================

def build_ergani_export_df(
    classified: pd.DataFrame,
    employees: pd.DataFrame,
    year: int,
) -> pd.DataFrame:
    if classified.empty:
        return pd.DataFrame(columns=[
            "ΑΑ Παραρτηματος",
            "ΑΦΜ",
            "ΕΠΩΝΥΜΟ",
            "ΟΝΟΜΑ",
            "ΗΜΕΡΟΜΗΝΙΑ",
            "ΤΥΠΟΣ ΑΔΕΙΑΣ",
            "ΩΡΑ ΑΠΟ",
            "ΩΡΑ ΕΩΣ",
            "ΕΤΟΣ",
            "ΔΙΚ ΗΜΕΡΕΣ",
        ])

    export_df = classified.copy()

    export_df["ΤΥΠΟΣ ΑΔΕΙΑΣ"] = export_df["Τύπος Απουσίας"].map(CLASSIFIED_TO_ERGANI_LEAVE_TYPE)
    export_df["ΗΜΕΡΟΜΗΝΙΑ"] = pd.to_datetime(export_df["Ημ/νία"]).dt.strftime("%d/%m/%Y")
    export_df["ΩΡΑ ΑΠΟ"] = ""
    export_df["ΩΡΑ ΕΩΣ"] = ""

    export_df["ΕΤΟΣ"] = export_df["Έτος Άδειας"].apply(
        lambda x: "" if pd.isna(x) else int(x)
    )

    employee_leave_info = employees[
        [
            "ΑΦΜ",
            "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους",
            "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους",
        ]
    ].drop_duplicates()

    export_df = export_df.merge(
        employee_leave_info,
        on="ΑΦΜ",
        how="left"
    )

    export_df["ΔΙΚ ΗΜΕΡΕΣ"] = export_df.apply(
        lambda row: (v if (v := get_entitled_days(row, year)) is not None else ""),
        axis=1,
    )

    export_df = export_df.rename(columns={
        "Επώνυμο": "ΕΠΩΝΥΜΟ",
        "Όνομα": "ΟΝΟΜΑ",
    })

    export_df = export_df[
        [
            "ΑΑ Παραρτηματος",
            "ΑΦΜ",
            "ΕΠΩΝΥΜΟ",
            "ΟΝΟΜΑ",
            "ΗΜΕΡΟΜΗΝΙΑ",
            "ΤΥΠΟΣ ΑΔΕΙΑΣ",
            "ΩΡΑ ΑΠΟ",
            "ΩΡΑ ΕΩΣ",
            "ΕΤΟΣ",
            "ΔΙΚ ΗΜΕΡΕΣ",
        ]
    ].sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ", "ΗΜΕΡΟΜΗΝΙΑ"]).reset_index(drop=True)

    export_df["ΑΦΜ"] = export_df["ΑΦΜ"].astype(str)
    return export_df


def write_ergani_exports_by_branch(
    export_df: pd.DataFrame,
    output_dir: Path,
    year: int,
    month: int,
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)

    created_files = []

    if export_df.empty:
        return created_files

    for branch_value, branch_df in export_df.groupby("ΑΑ Παραρτηματος", dropna=False):
        branch_part = branch_to_filename_part(branch_value)
        file_path = output_dir / f"ergani_export_parartima_{branch_part}_{year}_{month:02d}.xlsx"

        branch_out = branch_df.drop(columns=["ΑΑ Παραρτηματος"]).copy()

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            branch_out.to_excel(writer, sheet_name="Άδειες", index=False)
            force_text_column(writer.sheets["Άδειες"], "ΑΦΜ")

        created_files.append(file_path)

    return created_files


# =========================
# VALIDATION
# =========================

def build_validation_report(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
    employees: pd.DataFrame,
    absences: pd.DataFrame,
    classified: pd.DataFrame,
    year: int,
    month: int,
) -> pd.DataFrame:
    rows = []

    month_raw = raw_df.copy()
    month_raw["ΑΑ Παραρτηματος"] = pd.to_numeric(month_raw["ΑΑ Παραρτηματος"], errors="coerce")
    month_raw["ΑΦΜ"] = month_raw["ΑΦΜ"].astype(str).str.strip()
    month_raw["Επώνυμο"] = month_raw["Επώνυμο"].astype(str).str.strip()
    month_raw["Όνομα"] = month_raw["Όνομα"].astype(str).str.strip()
    month_raw["Ημ/νία"] = pd.to_datetime(month_raw["Ημ/νία"], dayfirst=True, errors="coerce").dt.normalize()

    month_raw = month_raw[
        (month_raw["Ημ/νία"].dt.year == year) &
        (month_raw["Ημ/νία"].dt.month == month)
    ].copy()

    employees_afm = set(employees["ΑΦΜ"].astype(str).str.strip())

    unknown_employees = month_raw[~month_raw["ΑΦΜ"].isin(employees_afm)].drop_duplicates(
        subset=["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"]
    )
    for _, r in unknown_employees.iterrows():
        rows.append(make_validation_row(
            "ERROR",
            "Άγνωστος εργαζόμενος",
            "Το attendance περιέχει ΑΦΜ που δεν υπάρχει στο employees.xlsx",
            r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"], "", r["ΑΦΜ"]
        ))

    missing_times = month_raw[
        month_raw["Από"].isna() | month_raw["Έως"].isna() |
        (month_raw["Από"].astype(str).str.strip() == "") |
        (month_raw["Έως"].astype(str).str.strip() == "")
    ]
    for _, r in missing_times.iterrows():
        rows.append(make_validation_row(
            "WARNING",
            "Ελλιπής ώρα",
            "Λείπει ώρα Από ή/και Έως",
            r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
            format_validation_date(r["Ημ/νία"]),
            f"Από={r['Από']}, Έως={r['Έως']}"
        ))

    check_times = month_raw.copy()
    check_times["start"] = check_times["Από"].apply(to_minutes)
    check_times["end"] = check_times["Έως"].apply(to_minutes)

    invalid_times = check_times[
        (
            ~(check_times["Από"].isna() | (check_times["Από"].astype(str).str.strip() == "")) &
            check_times["start"].isna()
        ) |
        (
            ~(check_times["Έως"].isna() | (check_times["Έως"].astype(str).str.strip() == "")) &
            check_times["end"].isna()
        )
    ]

    for _, r in invalid_times.iterrows():
        rows.append(make_validation_row(
            "ERROR",
            "Μη έγκυρη ώρα",
            "Η ώρα Από ή/και Έως δεν μπορεί να διαβαστεί σωστά",
            r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
            format_validation_date(r["Ημ/νία"]),
            f"Από={r['Από']}, Έως={r['Έως']}"
        ))

    dup_day = (
        month_raw.groupby(["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία"])
        .size()
        .reset_index(name="Πλήθος")
    )
    dup_day = dup_day[dup_day["Πλήθος"] > 1]

    for _, r in dup_day.iterrows():
        rows.append(make_validation_row(
            "WARNING",
            "Πολλαπλές παρουσίες ίδια μέρα",
            "Υπάρχουν πολλαπλά attendance entries για τον ίδιο εργαζόμενο την ίδια ημέρα",
            r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
            format_validation_date(r["Ημ/νία"]),
            r["Πλήθος"]
        ))

    duration_df = check_times.dropna(subset=["start", "end"]).copy()
    duration_df["worked"] = duration_df["end"] - duration_df["start"]
    duration_df.loc[duration_df["worked"] < 0, "worked"] += 1440

    suspicious_duration = duration_df[duration_df["worked"] > 12 * 60]
    for _, r in suspicious_duration.iterrows():
        rows.append(make_validation_row(
            "WARNING",
            "Υπερβολική διάρκεια",
            "Η δηλωμένη διάρκεια εργασίας ξεπερνά τις 12 ώρες",
            r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
            format_validation_date(r["Ημ/νία"]),
            minutes_to_hhmm(r["worked"])
        ))

    cleaned_month = cleaned_df[
        (cleaned_df["Ημ/νία"].dt.year == year) &
        (cleaned_df["Ημ/νία"].dt.month == month)
    ].copy()

    worked_afm = set(cleaned_month["ΑΦΜ"].astype(str).str.strip())
    no_presence = employees[~employees["ΑΦΜ"].isin(worked_afm)].copy()

    for _, r in no_presence.iterrows():
        rows.append(make_validation_row(
            "WARNING",
            "Καμία παρουσία",
            "Ο εργαζόμενος δεν έχει καμία παρουσία στον μήνα",
            r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"], "", ""
        ))

    if not classified.empty:
        abs_ref = absences.copy()
        class_ref = classified.copy()

        merged = class_ref.merge(
            abs_ref[["ΑΦΜ", "Ημ/νία"]],
            on=["ΑΦΜ", "Ημ/νία"],
            how="left",
            indicator=True
        )
        unmatched = merged[merged["_merge"] == "left_only"]

        for _, r in unmatched.iterrows():
            rows.append(make_validation_row(
                "ERROR",
                "Μη αντιστοιχισμένη ταξινόμηση",
                "Υπάρχει classified absence που δεν αντιστοιχεί σε detected absence",
                r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
                format_validation_date(r["Ημ/νία"]),
                r["Τύπος Απουσίας"]
            ))

        dup_classified = (
            class_ref.groupby(["ΑΑ Παραρτηματος", "ΑΦΜ", "Ημ/νία", "Τύπος Απουσίας", "Έτος Άδειας"])
            .size()
            .reset_index(name="Πλήθος")
        )
        dup_classified = dup_classified[dup_classified["Πλήθος"] > 1]

        for _, r in dup_classified.iterrows():
            rows.append(make_validation_row(
                "WARNING",
                "Διπλή ταξινόμηση",
                "Υπάρχουν duplicate rows στο classified_absences",
                r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], "", "",
                format_validation_date(r["Ημ/νία"]),
                r["Πλήθος"]
            ))

        invalid_leave_year = class_ref[
            (class_ref["Τύπος Απουσίας"] == "Κανονική άδεια") &
            (~class_ref["Έτος Άδειας"].isin([year - 1, year]))
        ]
        for _, r in invalid_leave_year.iterrows():
            rows.append(make_validation_row(
                "ERROR",
                "Μη αποδεκτό έτος άδειας",
                f"Το έτος άδειας πρέπει να είναι {year - 1} ή {year}",
                r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
                format_validation_date(r["Ημ/νία"]),
                r["Έτος Άδειας"]
            ))

        unmapped_types = class_ref[class_ref["Τύπος Απουσίας"].map(CLASSIFIED_TO_ERGANI_LEAVE_TYPE).isna()]
        for _, r in unmapped_types.iterrows():
            rows.append(make_validation_row(
                "ERROR",
                "Χωρίς mapping ΕΡΓΑΝΗ",
                "Ο τύπος απουσίας δεν έχει αντιστοίχιση προς τύπο άδειας ΕΡΓΑΝΗ",
                r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
                format_validation_date(r["Ημ/νία"]),
                r["Τύπος Απουσίας"]
            ))

        mapped_types = class_ref["Τύπος Απουσίας"].map(CLASSIFIED_TO_ERGANI_LEAVE_TYPE)
        invalid_ergani_labels = class_ref[~mapped_types.fillna("").isin(ERGANI_LEAVE_TYPES)]
        for _, r in invalid_ergani_labels.iterrows():
            mapped_value = CLASSIFIED_TO_ERGANI_LEAVE_TYPE.get(r["Τύπος Απουσίας"], "")
            rows.append(make_validation_row(
                "ERROR",
                "Μη αποδεκτός τύπος ΕΡΓΑΝΗ",
                "Το mapped label δεν υπάρχει στη λίστα τύπων άδειας ΕΡΓΑΝΗ",
                r["ΑΑ Παραρτηματος"], r["ΑΦΜ"], r["Επώνυμο"], r["Όνομα"],
                format_validation_date(r["Ημ/νία"]),
                mapped_value
            ))

    validation = pd.DataFrame(rows)

    if validation.empty:
        validation = pd.DataFrame([{
            "Επίπεδο": "INFO",
            "Κατηγορία": "Validation",
            "Μήνυμα": "Δεν εντοπίστηκαν προβλήματα",
            "ΑΑ Παραρτηματος": "",
            "ΑΦΜ": "",
            "Επώνυμο": "",
            "Όνομα": "",
            "Ημ/νία": "",
            "Τιμή": "",
        }])

    validation = validation.sort_values(
        ["Επίπεδο", "Κατηγορία", "ΑΑ Παραρτηματος", "ΑΦΜ", "Ημ/νία"],
        na_position="last"
    ).reset_index(drop=True)

    validation["ΑΦΜ"] = validation["ΑΦΜ"].astype(str)
    return validation

def build_alerts_report(
    employees: pd.DataFrame,
    absences: pd.DataFrame,
    classified: pd.DataFrame,
    workdays: pd.DataFrame,
    overtime_summary: pd.DataFrame,
    leaves: pd.DataFrame,
    year: int,
) -> pd.DataFrame:
    rows = []

    # 1. Χαμηλό υπόλοιπο τρέχοντος έτους
    low_leave = leaves[
        pd.to_numeric(leaves["Υπόλοιπο Τρέχοντος Έτους Μετά"], errors="coerce").fillna(0)
        <= LOW_CURRENT_LEAVE_BALANCE_THRESHOLD
    ].copy()

    for _, r in low_leave.iterrows():
        rows.append(make_alert_row(
            level="WARNING",
            category="Χαμηλό υπόλοιπο άδειας",
            aa=r["ΑΑ Παραρτηματος"],
            afm=r["ΑΦΜ"],
            last_name=r["Επώνυμο"],
            first_name=r["Όνομα"],
            message=f"Χαμηλό υπόλοιπο τρέχοντος έτους (≤ {LOW_CURRENT_LEAVE_BALANCE_THRESHOLD})",
            value=r["Υπόλοιπο Τρέχοντος Έτους Μετά"],
        ))

    # 2. Καμία παρουσία στον μήνα
    worked_afm = set(workdays["ΑΦΜ"].astype(str).str.strip())
    no_presence = employees[~employees["ΑΦΜ"].astype(str).str.strip().isin(worked_afm)].copy()

    for _, r in no_presence.iterrows():
        rows.append(make_alert_row(
            level="WARNING",
            category="Καμία παρουσία",
            aa=r["ΑΑ Παραρτηματος"],
            afm=r["ΑΦΜ"],
            last_name=r["Επώνυμο"],
            first_name=r["Όνομα"],
            message="Ο εργαζόμενος δεν έχει καμία παρουσία στον μήνα",
            value="0",
        ))

    # 3. Πολλές απουσίες
    if not absences.empty:
        abs_tmp = absences.copy()
        abs_count = (
            abs_tmp.groupby(["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"], as_index=False)
            .size()
            .rename(columns={"size": "Πλήθος Απουσιών"})
        )

        high_abs = abs_count[abs_count["Πλήθος Απουσιών"] > HIGH_ABSENCE_COUNT_THRESHOLD]

        for _, r in high_abs.iterrows():
            rows.append(make_alert_row(
                level="WARNING",
                category="Πολλές απουσίες",
                aa=r["ΑΑ Παραρτηματος"],
                afm=r["ΑΦΜ"],
                last_name=r["Επώνυμο"],
                first_name=r["Όνομα"],
                message=f"Πλήθος απουσιών μεγαλύτερο από {HIGH_ABSENCE_COUNT_THRESHOLD}",
                value=r["Πλήθος Απουσιών"],
            ))

    # 4. Πολλή υπεργασία
    if "Σύνολο Υπεργασίας Λεπτά" in overtime_summary.columns:
        high_overwork = overtime_summary[
            pd.to_numeric(overtime_summary["Σύνολο Υπεργασίας Λεπτά"], errors="coerce").fillna(0)
            > HIGH_OVERWORK_MINUTES_THRESHOLD
        ].copy()

        for _, r in high_overwork.iterrows():
            rows.append(make_alert_row(
                level="INFO",
                category="Πολλή υπεργασία",
                aa=r["ΑΑ Παραρτηματος"],
                afm=r["ΑΦΜ"],
                last_name=r["Επώνυμο"],
                first_name=r["Όνομα"],
                message=f"Σύνολο υπεργασίας μεγαλύτερο από {minutes_to_hhmm(HIGH_OVERWORK_MINUTES_THRESHOLD)}",
                value=r["Σύνολο Υπεργασίας (HH:MM)"],
            ))

    # 5. Πολλή υπερωρία
    if "Σύνολο Υπερωρίας Λεπτά" in overtime_summary.columns:
        high_overtime = overtime_summary[
            pd.to_numeric(overtime_summary["Σύνολο Υπερωρίας Λεπτά"], errors="coerce").fillna(0)
            > HIGH_OVERTIME_MINUTES_THRESHOLD
        ].copy()

        for _, r in high_overtime.iterrows():
            rows.append(make_alert_row(
                level="WARNING",
                category="Πολλή υπερωρία",
                aa=r["ΑΑ Παραρτηματος"],
                afm=r["ΑΦΜ"],
                last_name=r["Επώνυμο"],
                first_name=r["Όνομα"],
                message=f"Σύνολο υπερωρίας μεγαλύτερο από {minutes_to_hhmm(HIGH_OVERTIME_MINUTES_THRESHOLD)}",
                value=r["Σύνολο Υπερωρίας (HH:MM)"],
            ))

    # 6. Απουσίες χωρίς classification
    if not absences.empty:
        abs_ref = absences.copy()

        if classified.empty:
            unclassified = (
                abs_ref.groupby(["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"], as_index=False)
                .size()
                .rename(columns={"size": "Αταξινόμητες Απουσίες"})
            )
        else:
            class_ref = classified.copy()

            merged = abs_ref.merge(
                class_ref[["ΑΦΜ", "Ημ/νία"]],
                on=["ΑΦΜ", "Ημ/νία"],
                how="left",
                indicator=True
            )

            unclassified_rows = merged[merged["_merge"] == "left_only"].copy()

            unclassified = (
                unclassified_rows.groupby(["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"], as_index=False)
                .size()
                .rename(columns={"size": "Αταξινόμητες Απουσίες"})
            )

        for _, r in unclassified.iterrows():
            rows.append(make_alert_row(
                level="WARNING",
                category="Απουσίες χωρίς classification",
                aa=r["ΑΑ Παραρτηματος"],
                afm=r["ΑΦΜ"],
                last_name=r["Επώνυμο"],
                first_name=r["Όνομα"],
                message="Υπάρχουν detected absences χωρίς συμπληρωμένη ταξινόμηση",
                value=r["Αταξινόμητες Απουσίες"],
            ))

    # 7. Κανονική άδεια με μηδενικό entitlement στο δηλωμένο έτος
    if not classified.empty:
        class_check = classified.copy()

        employee_leave_info = employees[
            [
                "ΑΦΜ",
                "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους",
                "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους",
            ]
        ].drop_duplicates()

        class_check = class_check.merge(employee_leave_info, on="ΑΦΜ", how="left")

        class_check["δικ_ημέρες_check"] = class_check.apply(
            lambda row: get_entitled_days(row, year), axis=1
        )

        invalid_entitlement = class_check[
            (class_check["Τύπος Απουσίας"] == "Κανονική άδεια") &
            (
                class_check["δικ_ημέρες_check"].isna() |
                (pd.to_numeric(class_check["δικ_ημέρες_check"], errors="coerce").fillna(0) <= 0)
            )
        ]

        for _, r in invalid_entitlement.iterrows():
            rows.append(make_alert_row(
                level="ERROR",
                category="Μηδενικό entitlement άδειας",
                aa=r["ΑΑ Παραρτηματος"],
                afm=r["ΑΦΜ"],
                last_name=r["Επώνυμο"],
                first_name=r["Όνομα"],
                message="Κανονική άδεια με μηδενικό ή μη διαθέσιμο entitlement στο δηλωμένο έτος",
                value=r["Έτος Άδειας"],
            ))

    alerts = pd.DataFrame(rows)

    if alerts.empty:
        alerts = pd.DataFrame([{
            "Επίπεδο": "INFO",
            "Κατηγορία": "Alerts",
            "ΑΑ Παραρτηματος": "",
            "ΑΦΜ": "",
            "Επώνυμο": "",
            "Όνομα": "",
            "Μήνυμα": "Δεν εντοπίστηκαν alerts",
            "Τιμή": "",
        }])

    alerts = alerts.sort_values(
        ["Επίπεδο", "Κατηγορία", "ΑΑ Παραρτηματος", "ΑΦΜ"],
        na_position="last"
    ).reset_index(drop=True)

    alerts["ΑΦΜ"] = alerts["ΑΦΜ"].astype(str)
    return alerts


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
    ergani_output_dir = root / "data/output"

    raw_df = load_attendance(raw)
    df = clean_attendance(raw_df)
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

    alerts = build_alerts_report(
        employees=employees,
        absences=absences,
        classified=classified,
        workdays=workdays,
        overtime_summary=overtime_s,
        leaves=leaves,
        year=year,
    )

    validation = build_validation_report(
        raw_df=raw_df,
        cleaned_df=df,
        employees=employees,
        absences=absences,
        classified=classified,
        year=year,
        month=month,
    )

    ergani_export_df = build_ergani_export_df(
        classified=classified,
        employees=employees,
        year=year,
    )

    ergani_files = write_ergani_exports_by_branch(
        export_df=ergani_export_df,
        output_dir=ergani_output_dir,
        year=year,
        month=month,
    )

    output.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        absences.to_excel(writer, sheet_name="Απουσίες", index=False)
        workdays.to_excel(writer, sheet_name="Ημέρες", index=False)
        overtime_d.to_excel(writer, sheet_name="Υπερωρίες", index=False)
        overtime_s.to_excel(writer, sheet_name="Σύνολο Extra", index=False)
        leaves.to_excel(writer, sheet_name="Άδειες", index=False)
        validation.to_excel(writer, sheet_name="Validation", index=False)
        alerts.to_excel(writer, sheet_name="Alerts", index=False)

        for sheet in writer.sheets.values():
            force_text_column(sheet, "ΑΦΜ")

    print("Έτοιμο:", output)
    print("Template αδειών:", classified_file)

    if ergani_files:
        print("Ergani export files:")
        for file_path in ergani_files:
            print("-", file_path)
    else:
        print("Δεν δημιουργήθηκαν αρχεία Ergani export (δεν υπάρχουν συμπληρωμένες ταξινομημένες άδειες).")


if __name__ == "__main__":
    main()