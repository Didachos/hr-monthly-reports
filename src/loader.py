from pathlib import Path
import pandas as pd


def load_attendance(file_path: str | Path) -> pd.DataFrame:
    file_path = Path(file_path)
    df = pd.read_excel(file_path)

    expected_columns = [
        "ΑΑ Παραρτηματος",
        "ΑΦΜ",
        "Επώνυμο",
        "Όνομα",
        "Ημ/νία",
        "Από",
        "Έως",
    ]

    missing = [col for col in expected_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Λείπουν στήλες από το Excel: {missing}")

    return df