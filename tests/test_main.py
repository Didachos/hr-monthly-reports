import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from main import (
    build_ergani_export_df,
    build_leave_summary,
    calculate_overtime,
    find_absences,
    get_entitled_days,
    minutes_to_hhmm,
    to_minutes,
)


# =========================
# FIXTURES
# =========================

def make_employees(**overrides):
    defaults = dict(
        ΑΑΠαραρτηματος=1,
        ΑΦΜ="123456789",
        Επώνυμο="Παπαδόπουλος",
        Όνομα="Γιώργης",
        ΗμερομηνίαΠρόσληψης=pd.Timestamp("2020-01-01"),
        ΗμερομηνίαΑποχώρησης=pd.NaT,
        ΔικαιούμενηΚανονικήΆδειαΠροηγούμενουΈτους=20,
        ΥπόλοιποΠροηγούμενουΈτους=5,
        ΔικαιούμενηΚανονικήΆδειαΤρέχοντοςΈτους=25,
    )
    defaults.update(overrides)
    return pd.DataFrame([{
        "ΑΑ Παραρτηματος": defaults["ΑΑΠαραρτηματος"],
        "ΑΦΜ": defaults["ΑΦΜ"],
        "Επώνυμο": defaults["Επώνυμο"],
        "Όνομα": defaults["Όνομα"],
        "Ημερομηνία Πρόσληψης": defaults["ΗμερομηνίαΠρόσληψης"],
        "Ημερομηνία Αποχώρησης": defaults["ΗμερομηνίαΑποχώρησης"],
        "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους": defaults["ΔικαιούμενηΚανονικήΆδειαΠροηγούμενουΈτους"],
        "Υπόλοιπο Προηγούμενου Έτους": defaults["ΥπόλοιποΠροηγούμενουΈτους"],
        "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους": defaults["ΔικαιούμενηΚανονικήΆδειαΤρέχοντοςΈτους"],
    }])


def make_attendance(afm="123456789", date="2026-01-05", start="09:00", end="17:00", aa=1):
    return pd.DataFrame([{
        "ΑΑ Παραρτηματος": aa,
        "ΑΦΜ": afm,
        "Επώνυμο": "Παπαδόπουλος",
        "Όνομα": "Γιώργης",
        "Ημ/νία": pd.Timestamp(date),
        "Από": start,
        "Έως": end,
    }])


# =========================
# HELPERS
# =========================

class TestMinutesToHhmm:
    def test_exact_hours(self):
        assert minutes_to_hhmm(120) == "02:00"

    def test_with_minutes(self):
        assert minutes_to_hhmm(90) == "01:30"

    def test_zero(self):
        assert minutes_to_hhmm(0) == "00:00"

    def test_over_24h(self):
        assert minutes_to_hhmm(25 * 60) == "25:00"


class TestToMinutes:
    def test_hhmm_format(self):
        assert to_minutes("09:00") == 540

    def test_with_seconds(self):
        assert to_minutes("09:00:00") == 540

    def test_nan(self):
        assert to_minutes(float("nan")) is None

    def test_empty_string(self):
        assert to_minutes("") is None

    def test_invalid(self):
        assert to_minutes("abc") is None


# =========================
# GET ENTITLED DAYS
# =========================

class TestGetEntitledDays:
    BASE_ROW = {
        "Τύπος Απουσίας": "Κανονική άδεια",
        "Έτος Άδειας": 2026,
        "Δικαιούμενη Κανονική Άδεια Προηγούμενου Έτους": 20,
        "Δικαιούμενη Κανονική Άδεια Τρέχοντος Έτους": 25,
    }

    def test_current_year(self):
        assert get_entitled_days(self.BASE_ROW, 2026) == 25

    def test_previous_year(self):
        row = {**self.BASE_ROW, "Έτος Άδειας": 2025}
        assert get_entitled_days(row, 2026) == 20

    def test_wrong_type(self):
        row = {**self.BASE_ROW, "Τύπος Απουσίας": "Άδεια ασθενείας"}
        assert get_entitled_days(row, 2026) is None

    def test_missing_year(self):
        row = {**self.BASE_ROW, "Έτος Άδειας": float("nan")}
        assert get_entitled_days(row, 2026) is None

    def test_unknown_year(self):
        row = {**self.BASE_ROW, "Έτος Άδειας": 2020}
        assert get_entitled_days(row, 2026) is None


# =========================
# FIND ABSENCES
# =========================

class TestFindAbsences:
    def _employees(self):
        return make_employees()

    def _attendance(self, dates):
        rows = []
        for d in dates:
            rows.append({
                "ΑΑ Παραρτηματος": 1,
                "ΑΦΜ": "123456789",
                "Επώνυμο": "Παπαδόπουλος",
                "Όνομα": "Γιώργης",
                "Ημ/νία": pd.Timestamp(d),
                "Από": "09:00",
                "Έως": "17:00",
            })
        return pd.DataFrame(rows)

    def test_full_month_no_absences(self):
        # Παρέχουμε παρουσία για όλες τις εργάσιμες μέρες Ιανουαρίου 2026
        import holidays as hol
        from calendar import monthrange
        gr_holidays = {pd.Timestamp(d).normalize() for d in hol.country_holidays("GR", years=2026).keys()}
        dates = [
            f"2026-01-{d:02d}"
            for d in range(1, monthrange(2026, 1)[1] + 1)
            if pd.Timestamp(f"2026-01-{d:02d}").weekday() < 5
            and pd.Timestamp(f"2026-01-{d:02d}").normalize() not in gr_holidays
        ]
        df = self._attendance(dates)
        result = find_absences(df, self._employees(), 2026, 1)
        assert result.empty

    def test_detects_missing_day(self):
        # Δίνουμε παρουσία μόνο για μία μέρα — οι υπόλοιπες εργάσιμες θα είναι απουσίες
        df = self._attendance(["2026-01-05"])
        result = find_absences(df, self._employees(), 2026, 1)
        assert not result.empty
        assert "123456789" in result["ΑΦΜ"].values

    def test_weekends_excluded(self):
        # Παρουσία μόνο Δευτέρα — Σαββατοκύριακο δεν πρέπει να εμφανίζεται ως απουσία
        df = self._attendance(["2026-01-05"])
        result = find_absences(df, self._employees(), 2026, 1)
        absent_dates = pd.to_datetime(result["Ημ/νία"]).dt.weekday
        assert (absent_dates >= 5).sum() == 0

    def test_date_stays_datetime(self):
        df = self._attendance(["2026-01-05"])
        result = find_absences(df, self._employees(), 2026, 1)
        assert pd.api.types.is_datetime64_any_dtype(result["Ημ/νία"])

    def test_employee_before_hire_excluded(self):
        employees = make_employees(ΗμερομηνίαΠρόσληψης=pd.Timestamp("2026-01-20"))
        df = self._attendance(["2026-01-20"])
        result = find_absences(df, employees, 2026, 1)
        # Μόνο μέρες μετά την πρόσληψη (20+) μπορούν να είναι απουσίες
        if not result.empty:
            assert all(pd.to_datetime(result["Ημ/νία"]) >= pd.Timestamp("2026-01-20"))

    def test_terminated_employee_excluded(self):
        employees = make_employees(ΗμερομηνίαΑποχώρησης=pd.Timestamp("2026-01-10"))
        df = self._attendance(["2026-01-05"])
        result = find_absences(df, employees, 2026, 1)
        if not result.empty:
            assert all(pd.to_datetime(result["Ημ/νία"]) <= pd.Timestamp("2026-01-10"))


# =========================
# CALCULATE OVERTIME
# =========================

class TestCalculateOvertime:
    def _df(self, start, end, date="2026-01-05"):
        return pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp(date),
            "Από": start,
            "Έως": end,
        }])

    def test_standard_8h_no_extra(self):
        _, summary = calculate_overtime(self._df("09:00", "17:00"), 2026, 1)
        assert summary["Σύνολο Υπεργασίας Λεπτά"].iloc[0] == 0
        assert summary["Σύνολο Υπερωρίας Λεπτά"].iloc[0] == 0

    def test_overwork_first_hour(self):
        # 8.5ώρες → 30 λεπτά υπεργασία, 0 υπερωρία
        _, summary = calculate_overtime(self._df("09:00", "17:30"), 2026, 1)
        assert summary["Σύνολο Υπεργασίας Λεπτά"].iloc[0] == 30
        assert summary["Σύνολο Υπερωρίας Λεπτά"].iloc[0] == 0

    def test_overtime_beyond_first_hour(self):
        # 10ώρες → 60 λεπτά υπεργασία, 60 λεπτά υπερωρία
        _, summary = calculate_overtime(self._df("09:00", "19:00"), 2026, 1)
        assert summary["Σύνολο Υπεργασίας Λεπτά"].iloc[0] == 60
        assert summary["Σύνολο Υπερωρίας Λεπτά"].iloc[0] == 60

    def test_date_stays_datetime(self):
        detailed, _ = calculate_overtime(self._df("09:00", "17:00"), 2026, 1)
        assert pd.api.types.is_datetime64_any_dtype(detailed["Ημ/νία"])


# =========================
# BUILD LEAVE SUMMARY
# =========================

class TestBuildLeaveSummary:
    def _employees(self):
        return make_employees()

    def test_empty_classified_returns_zeros(self):
        result = build_leave_summary(pd.DataFrame(), self._employees(), 2026)
        assert result["Κανονική Άδεια από Τρέχον Έτος"].iloc[0] == 0
        assert result["Σύνολο Ασθενείας"].iloc[0] == 0
        assert result["Σύνολο Άνευ Αποδοχών"].iloc[0] == 0

    def test_empty_classified_preserves_entitlement_balance(self):
        result = build_leave_summary(pd.DataFrame(), self._employees(), 2026, month=1)
        assert result["Υπόλοιπο Τρέχοντος Έτους Μετά"].iloc[0] == 25
        assert result["Υπόλοιπο Προηγούμενου Έτους Μετά"].iloc[0] == 5

    def test_annual_leave_current_year_counted(self):
        classified = pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp("2026-01-10"),
            "Τύπος Απουσίας": "Κανονική άδεια",
            "Έτος Άδειας": 2026,
        }])
        result = build_leave_summary(classified, self._employees(), 2026, month=1)
        assert result["Κανονική Άδεια από Τρέχον Έτος"].iloc[0] == 1
        assert result["Υπόλοιπο Τρέχοντος Έτους Μετά"].iloc[0] == 24

    def test_annual_leave_previous_year_counted(self):
        classified = pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp("2026-01-10"),
            "Τύπος Απουσίας": "Κανονική άδεια",
            "Έτος Άδειας": 2025,
        }])
        result = build_leave_summary(classified, self._employees(), 2026, month=1)
        assert result["Κανονική Άδεια από Προηγούμενο Έτος"].iloc[0] == 1
        assert result["Υπόλοιπο Προηγούμενου Έτους Μετά"].iloc[0] == 4

    def test_previous_year_balance_expires_after_march(self):
        result = build_leave_summary(pd.DataFrame(), self._employees(), 2026, month=4)
        assert result["Υπόλοιπο Προηγούμενου Έτους Μετά"].iloc[0] == 0

    def test_previous_year_balance_valid_in_march(self):
        result = build_leave_summary(pd.DataFrame(), self._employees(), 2026, month=3)
        assert result["Υπόλοιπο Προηγούμενου Έτους Μετά"].iloc[0] == 5

    def test_balance_cannot_go_negative(self):
        classified = pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp(f"2026-01-{d:02d}"),
            "Τύπος Απουσίας": "Κανονική άδεια",
            "Έτος Άδειας": 2025,
        } for d in range(1, 12)])  # 11 μέρες, υπόλοιπο μόνο 5
        result = build_leave_summary(classified, self._employees(), 2026, month=1)
        assert result["Υπόλοιπο Προηγούμενου Έτους Μετά"].iloc[0] == 0

    def test_sick_leave_counted(self):
        classified = pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp("2026-01-10"),
            "Τύπος Απουσίας": "Άδεια ασθενείας",
            "Έτος Άδειας": pd.NA,
        }])
        result = build_leave_summary(classified, self._employees(), 2026)
        assert result["Σύνολο Ασθενείας"].iloc[0] == 1


# =========================
# BUILD ERGANI EXPORT
# =========================

class TestBuildErganiExportDf:
    def _employees(self):
        return make_employees()

    def test_empty_classified_returns_empty_df(self):
        result = build_ergani_export_df(pd.DataFrame(), self._employees(), 2026)
        assert result.empty

    def test_entitled_days_filled_for_annual_leave(self):
        classified = pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp("2026-01-10"),
            "Τύπος Απουσίας": "Κανονική άδεια",
            "Έτος Άδειας": 2026,
        }])
        result = build_ergani_export_df(classified, self._employees(), 2026)
        assert result["ΔΙΚ ΗΜΕΡΕΣ"].iloc[0] == 25

    def test_entitled_days_empty_for_sick_leave(self):
        classified = pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp("2026-01-10"),
            "Τύπος Απουσίας": "Άδεια ασθενείας",
            "Έτος Άδειας": pd.NA,
        }])
        result = build_ergani_export_df(classified, self._employees(), 2026)
        assert result["ΔΙΚ ΗΜΕΡΕΣ"].iloc[0] == ""

    def test_ergani_leave_type_mapped(self):
        classified = pd.DataFrame([{
            "ΑΑ Παραρτηματος": 1,
            "ΑΦΜ": "123456789",
            "Επώνυμο": "Παπαδόπουλος",
            "Όνομα": "Γιώργης",
            "Ημ/νία": pd.Timestamp("2026-01-10"),
            "Τύπος Απουσίας": "Άδεια ασθενείας",
            "Έτος Άδειας": pd.NA,
        }])
        result = build_ergani_export_df(classified, self._employees(), 2026)
        assert result["ΤΥΠΟΣ ΑΔΕΙΑΣ"].iloc[0] == "Άδεια ασθένειας (ανυπαίτιο κώλυμα παροχής εργασίας)"
