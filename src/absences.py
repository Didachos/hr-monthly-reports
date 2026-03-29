from calendar import monthrange
import pandas as pd


def build_month_dates(year: int, month: int) -> pd.DataFrame:
    last_day = monthrange(year, month)[1]
    dates = pd.date_range(
        start=f"{year}-{month:02d}-01",
        end=f"{year}-{month:02d}-{last_day:02d}",
        freq="D",
    )
    return pd.DataFrame({"Ημ/νία": dates})


def normalize_afm(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )


def find_absences(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    df = df.copy()

    df["ΑΦΜ"] = normalize_afm(df["ΑΦΜ"])
    df["Ημ/νία"] = pd.to_datetime(df["Ημ/νία"], errors="coerce").dt.floor("D")

    month_df = df[
        (df["Ημ/νία"].dt.year == year) &
        (df["Ημ/νία"].dt.month == month)
    ].copy()

    employees = (
        month_df[["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    dates_df = build_month_dates(year, month)

    employees["key"] = 1
    dates_df["key"] = 1

    full_calendar = employees.merge(dates_df, on="key").drop(columns="key")

    # Δευτέρα-Παρασκευή μόνο
    full_calendar = full_calendar[full_calendar["Ημ/νία"].dt.weekday < 5].copy()

    presences = (
        month_df[["ΑΦΜ", "Ημ/νία"]]
        .drop_duplicates()
        .assign(παρουσία=1)
    )

    # Extra normalization και στα δύο πριν το merge
    full_calendar["ΑΦΜ"] = normalize_afm(full_calendar["ΑΦΜ"])
    presences["ΑΦΜ"] = normalize_afm(presences["ΑΦΜ"])

    full_calendar["Ημ/νία"] = pd.to_datetime(full_calendar["Ημ/νία"]).dt.floor("D")
    presences["Ημ/νία"] = pd.to_datetime(presences["Ημ/νία"]).dt.floor("D")

    result = full_calendar.merge(
        presences,
        on=["ΑΦΜ", "Ημ/νία"],
        how="left"
    )

    absences = result[result["παρουσία"].isna()].copy()
    absences["Κατάσταση"] = "ΑΠΩΝ"

    return absences[
        ["ΑΑ Παραρτηματος", "ΑΦΜ", "Επώνυμο", "Όνομα", "Ημ/νία", "Κατάσταση"]
    ].sort_values(["Επώνυμο", "Όνομα", "Ημ/νία"])