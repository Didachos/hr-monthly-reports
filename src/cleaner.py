import pandas as pd


def clean_attendance(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    if "ΑΦΜ" in df.columns:
        df["ΑΦΜ"] = df["ΑΦΜ"].astype(str).str.strip()

    if "Επώνυμο" in df.columns:
        df["Επώνυμο"] = df["Επώνυμο"].astype(str).str.strip()

    if "Όνομα" in df.columns:
        df["Όνομα"] = df["Όνομα"].astype(str).str.strip()

    if "Ημ/νία" in df.columns:
        df["Ημ/νία"] = pd.to_datetime(df["Ημ/νία"], errors="coerce").dt.normalize()

    return df