from pathlib import Path
from calendar import monthrange
import sys
import pandas as pd
import holidays


STANDARD_WORK_MINUTES = 8 * 60


def load_attendance(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"Δεν βρέθηκε το αρχείο: {file_path}")

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


def clean_attendance(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df["ΑΑ Παραρτηματος"] = pd.to_numeric(df["ΑΑ Παραρτηματος"], errors="coerce")
    df["ΑΦΜ"] = df["ΑΦΜ"].astype(str).str.strip()
    df["Επώνυμο"] = df["Επώνυμο"].astype(str).str.strip()
    df["Όνομα"] = df["Όνομα"].astype(str).str.strip()

    df["Ημ/νία"] = pd.to_datetime(
        df["Ημ/νία"],
        errors="coerce",
        dayfirst=True
    ).dt.normalize()

    df["Από"] = df["Από"].astype(str).str.strip()
    df["Έως"] = df["Έως"].astype(str).str.strip()

    df = df.dropna(subset=["Ημ/νία", "ΑΦΜ"])
    df = df.drop_duplicates()

    return df


def build_month_dates(year: int, month: int) -> pd.DataFrame:
    last_day = monthrange(year, month)[1]
    dates = pd.date_range(
        start=f"{year}-{month:02d}-01",
        end=f"{year}-{month:02d}-{last_day:02d}",
        freq="D",
    ).normalize()

    return pd.DataFrame({"Ημ/νία": dates})


def get_greek_holidays(year: int) -> set[pd.Timestamp]:
    gr_holidays = holidays.country_holidays("GR", years=year)
    return {pd.Timestamp(d).normalize() for d in gr_holidays.keys()}


def find_absences(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    month_df = df[
        (df["Ημ/νία"].dt.year == year) &
        (df["Ημ/νία"].dt.month == month)
    ].copy()

    if month_df.empty:
        raise ValueError(
            f"Δεν βρέθηκαν δεδομένα για {month:02d}/{year} στο raw αρχείο."
        )

    employees = (
        month_df[["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    dates_df = build_month_dates(year, month)

    employees["key"] = 1
    dates_df["key"] = 1
    full_calendar = employees.merge(dates_df, on="key").drop(columns="key")

    gr_holidays = get_greek_holidays(year)

    full_calendar = full_calendar[
        (full_calendar["Ημ/νία"].dt.weekday < 5) &
        (~full_calendar["Ημ/νία"].isin(gr_holidays))
    ].copy()

    presences = month_df[["ΑΦΜ", "Ημ/νία"]].drop_duplicates().copy()
    presences["παρουσία"] = 1

    result = full_calendar.merge(
        presences,
        on=["ΑΦΜ", "Ημ/νία"],
        how="left"
    )

    absences = result[result["παρουσία"].isna()].copy()
    absences["Κατάσταση"] = "ΑΠΩΝ"

    absences = absences[
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία", "Κατάσταση"]
    ].sort_values(
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Ημ/νία"]
    ).reset_index(drop=True)

    absences["Ημ/νία"] = pd.to_datetime(absences["Ημ/νία"]).dt.strftime("%d/%m/%Y")

    return absences


def calculate_work_days(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    month_df = df[
        (df["Ημ/νία"].dt.year == year) &
        (df["Ημ/νία"].dt.month == month)
    ].copy()

    if month_df.empty:
        raise ValueError(
            f"Δεν βρέθηκαν δεδομένα για {month:02d}/{year} στο raw αρχείο."
        )

    unique_days = month_df[
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία"]
    ].drop_duplicates()

    result = (
        unique_days
        .groupby(
            ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"],
            as_index=False
        )
        .agg(**{"Σύνολο Ημερών Εργασίας": ("Ημ/νία", "nunique")})
        .sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ"])
        .reset_index(drop=True)
    )

    return result


def parse_time_to_minutes(value: str) -> float:
    if pd.isna(value):
        return float("nan")

    text = str(value).strip()
    if not text or text.lower() == "nan":
        return float("nan")

    parsed = pd.to_datetime(text, format="%H:%M", errors="coerce")
    if pd.isna(parsed):
        parsed = pd.to_datetime(text, errors="coerce")

    if pd.isna(parsed):
        return float("nan")

    return parsed.hour * 60 + parsed.minute


def minutes_to_hhmm(total_minutes: int) -> str:
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours:02d}:{minutes:02d}"


def calculate_overtime(df: pd.DataFrame, year: int, month: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    month_df = df[
        (df["Ημ/νία"].dt.year == year) &
        (df["Ημ/νία"].dt.month == month)
    ].copy()

    if month_df.empty:
        raise ValueError(
            f"Δεν βρέθηκαν δεδομένα για {month:02d}/{year} στο raw αρχείο."
        )

    month_df["Λεπτά Από"] = month_df["Από"].apply(parse_time_to_minutes)
    month_df["Λεπτά Έως"] = month_df["Έως"].apply(parse_time_to_minutes)

    month_df = month_df.dropna(subset=["Λεπτά Από", "Λεπτά Έως"]).copy()

    month_df["Λεπτά Από"] = month_df["Λεπτά Από"].astype(int)
    month_df["Λεπτά Έως"] = month_df["Λεπτά Έως"].astype(int)

    month_df["Συνολικά Λεπτά Εργασίας"] = month_df["Λεπτά Έως"] - month_df["Λεπτά Από"]

    month_df.loc[
        month_df["Συνολικά Λεπτά Εργασίας"] < 0,
        "Συνολικά Λεπτά Εργασίας"
    ] += 24 * 60

    month_df["Υπερωρία Λεπτά"] = (
        month_df["Συνολικά Λεπτά Εργασίας"] - STANDARD_WORK_MINUTES
    ).clip(lower=0)

    month_df["Κατάσταση"] = month_df["Υπερωρία Λεπτά"].apply(
        lambda x: "ΥΠΕΡΩΡΙΑ" if x > 0 else ""
    )

    month_df["Συνολική Διάρκεια"] = month_df["Συνολικά Λεπτά Εργασίας"].apply(minutes_to_hhmm)
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
            "Κατάσταση",
            "Υπερωρία Λεπτά",
            "Υπερωρία (HH:MM)",
        ]
    ].copy()

    detailed["Ημ/νία"] = pd.to_datetime(detailed["Ημ/νία"]).dt.strftime("%d/%m/%Y")

    detailed = detailed.sort_values(
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Ημ/νία"]
    ).reset_index(drop=True)

    summary = (
        month_df.groupby(
            ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"],
            as_index=False
        )
        .agg(
            **{
                "Ημέρες με Υπερωρία": ("Υπερωρία Λεπτά", lambda s: int((s > 0).sum())),
                "Σύνολο Υπερωρίας Λεπτά": ("Υπερωρία Λεπτά", "sum"),
            }
        )
        .sort_values(["ΑΑ Παραρτηματος", "ΑΦΜ"])
        .reset_index(drop=True)
    )

    summary["Σύνολο Υπερωρίας (HH:MM)"] = summary["Σύνολο Υπερωρίας Λεπτά"].apply(minutes_to_hhmm)

    return detailed, summary


def autofit_worksheet_columns(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            try:
                cell_value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(cell_value))
            except Exception:
                pass

        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 40)


def main() -> None:
    if len(sys.argv) != 3:
        raise ValueError("Χρήση: python src/main.py <year> <month>")

    year = int(sys.argv[1])
    month = int(sys.argv[2])

    if month < 1 or month > 12:
        raise ValueError("Ο μήνας πρέπει να είναι από 1 έως 12.")

    project_root = Path(__file__).resolve().parent.parent
    input_file = project_root / "data" / "input" / "raw_attendance.xlsx"
    monthly_output_file = (
        project_root / "data" / "output" / f"monthly_report_{year}_{month:02d}.xlsx"
    )

    df = load_attendance(input_file)
    df = clean_attendance(df)

    print("Min date:", df["Ημ/νία"].min())
    print("Max date:", df["Ημ/νία"].max())
    print("Unique dates in raw:", df["Ημ/νία"].nunique())
    print("Employees in raw:", df["ΑΦΜ"].nunique())

    month_holidays = sorted(
        d for d in get_greek_holidays(year) if d.month == month
    )
    print("Αργίες μήνα:", [d.strftime("%d/%m/%Y") for d in month_holidays])

    absences = find_absences(df, year=year, month=month)
    workdays = calculate_work_days(df, year=year, month=month)
    overtime_detailed, overtime_summary = calculate_overtime(df, year=year, month=month)

    monthly_output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(monthly_output_file, engine="openpyxl") as writer:
        absences.to_excel(writer, sheet_name="Απουσίες", index=False)
        workdays.to_excel(writer, sheet_name="Ημέρες Εργασίας", index=False)
        overtime_detailed.to_excel(writer, sheet_name="Υπερωρίες Αναλυτικά", index=False)
        overtime_summary.to_excel(writer, sheet_name="Υπερωρίες Σύνολα", index=False)

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            autofit_worksheet_columns(worksheet)

    print(f"Το μηνιαίο αρχείο δημιουργήθηκε: {monthly_output_file}")
    print()
    print("Περιεχόμενα sheets:")
    print("- Απουσίες")
    print("- Ημέρες Εργασίας")
    print("- Υπερωρίες Αναλυτικά")
    print("- Υπερωρίες Σύνολα")


if __name__ == "__main__":
    main()