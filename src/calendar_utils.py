import holidays
import pandas as pd


def get_greek_holidays(year: int) -> set[pd.Timestamp]:
    gr_holidays = holidays.country_holidays("GR", years=year)
    return {pd.Timestamp(day).normalize() for day in gr_holidays.keys()}


def is_working_day(date_value: pd.Timestamp, holiday_set: set[pd.Timestamp]) -> bool:
    date_value = pd.Timestamp(date_value).normalize()

    # 0=Monday ... 6=Sunday
    is_weekday = date_value.weekday() < 5
    is_holiday = date_value in holiday_set

    return is_weekday and not is_holiday