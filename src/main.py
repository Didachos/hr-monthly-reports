from pathlib import Path
from calendar import monthrange
import sys
import pandas as pd
import holidays


def load_attendance(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"Δεν βρέθηκε το αρχείο: {file_path}")

    # Κρατάμε το ΑΦΜ ως string για να μη χαθεί τυχόν leading zero
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

    # Ρητά dayfirst=True γιατί το αρχείο είναι dd/mm/yyyy
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
    df = df.copy()

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

    # Κρατάμε μόνο εργάσιμες: Δευτέρα-Παρασκευή και όχι αργίες
    full_calendar = full_calendar[
        (full_calendar["Ημ/νία"].dt.weekday < 5) &
        (~full_calendar["Ημ/νία"].isin(gr_holidays))
    ].copy()

    presences = (
        month_df[["ΑΦΜ", "Ημ/νία"]]
        .drop_duplicates()
        .copy()
    )
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

    # 1 μέρα εργασίας = 1 μοναδική ημερομηνία παρουσίας ανά εργαζόμενο
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


def main() -> None:
    if len(sys.argv) != 3:
        raise ValueError("Χρήση: python src/main.py <year> <month>")

    year = int(sys.argv[1])
    month = int(sys.argv[2])

    if month < 1 or month > 12:
        raise ValueError("Ο μήνας πρέπει να είναι από 1 έως 12.")

    project_root = Path(__file__).resolve().parent.parent
    input_file = project_root / "data" / "input" / "raw_attendance.xlsx"

    absences_output_file = (
        project_root / "data" / "output" / f"absences_{year}_{month:02d}.xlsx"
    )
    workdays_output_file = (
        project_root / "data" / "output" / f"workdays_{year}_{month:02d}.xlsx"
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
    absences["Ημ/νία"] = pd.to_datetime(absences["Ημ/νία"]).dt.strftime("%d/%m/%Y")

    workdays = calculate_work_days(df, year=year, month=month)

    absences_output_file.parent.mkdir(parents=True, exist_ok=True)

    absences.to_excel(absences_output_file, index=False)
    workdays.to_excel(workdays_output_file, index=False)

    print(f"Το αρχείο απουσιών δημιουργήθηκε: {absences_output_file}")
    print(f"Το αρχείο ημερών εργασίας δημιουργήθηκε: {workdays_output_file}")
    print(f"Σύνολο γραμμών απουσίας: {len(absences)}")
    print()
    print("Δείγμα ημερών εργασίας:")
    print(workdays.head(20).to_string(index=False))


if __name__ == "__main__":
    main()