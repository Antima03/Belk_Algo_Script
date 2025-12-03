from datetime import datetime, timedelta
from typing import Tuple, Dict, List
from utils import get_latest_fiscal
 
MONTHS_454 = ['FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','JAN']
# NRF season split: SP = Feb–Jul, FA = Aug–Jan
SP_MONTHS = {'FEB','MAR','APR','MAY','JUN','JUL'}
FA_MONTHS = {'AUG','SEP','OCT','NOV','DEC','JAN'}
 
def _first_sunday_on_or_after(d: datetime) -> datetime:
    """Return the first Sunday on or after date d."""
    return d + timedelta(days=(6 - d.weekday()) % 7)

def retail_year_start(retail_year: int) -> datetime:
    """
    NRF retail year start for a given retail_year (e.g., 2025):
    The first Sunday on/after Feb 1 of that year.
    """
    feb1 = datetime(retail_year, 2, 1)
    return _first_sunday_on_or_after(feb1)

def weeks_per_month_for_retail_year(retail_year: int) -> Dict[str, int]:
    """
    Return a dict of MONTHS_454 -> weeks for the given NRF retail year.
    Uses standard 4-5-4; if the year has 53 weeks, the extra week goes to JAN.
    """
    ry_start = retail_year_start(retail_year)
    ry_next  = retail_year_start(retail_year + 1)

    total_weeks = (ry_next - ry_start).days // 7  # 52 or 53

    base: List[int] = [4,5,4, 4,5,4, 4,5,4, 4,5,4]  # FEB..JAN = 52 weeks
    if total_weeks == 53:
        base[-1] += 1  # add extra week to JAN

    return {m: w for m, w in zip(MONTHS_454, base)}
 
def _retail_week_end_date(retail_year: int, retail_month: str, week_number: int, wpm: Dict[str, int]) -> datetime:
    """
    Compute the last calendar date (Saturday) of the given retail week in NRF 4-5-4.
    """
    retail_month = retail_month.upper()
    if retail_month not in MONTHS_454:
        raise ValueError(f"Invalid retail_month '{retail_month}'. Must be one of {MONTHS_454}.")
 
    max_weeks = wpm[retail_month]
    if not (1 <= week_number <= max_weeks):
        raise ValueError(f"week_number {week_number} out of range for {retail_month} (1..{max_weeks}).")
 
    ry_start = retail_year_start(retail_year)  # Sunday
    month_idx = MONTHS_454.index(retail_month)
    weeks_before_month = sum(wpm[m] for m in MONTHS_454[:month_idx])  # weeks to skip before this month
    # Week starts on Sunday:
    week_start = ry_start + timedelta(weeks=weeks_before_month + (week_number - 1))
    week_end = week_start + timedelta(days=6)  # Saturday
    return week_end

def get_previous_retail_week(df) :
    """
    Retail calendar details for the week BEFORE current_date (NRF 4-5-4).
    Returns:
      (retail_month, current_month_number, rolling_method, retail_week_number,
       retail_year, last_retail_year, last_month_number, season,
       weeks_FEB, weeks_MAR, ..., weeks_JAN)
    """
 
    # Weeks-per-month for this retail year (handles 53-week years)
    
    retail_year,retail_month,retail_week_number = get_latest_fiscal(df) 
    retail_year = 2025
    retail_month = "MAY"
    retail_week_number = 4
    wpm = weeks_per_month_for_retail_year(retail_year)

    current_date = _retail_week_end_date(retail_year, retail_month, retail_week_number, wpm)
    
    last_retail_year = retail_year - 1
 
 
    # Season & rolling method
    season = "SP" if retail_month in SP_MONTHS else "FA"
    rolling_method = "Current MTH" if retail_month in {'SEP','OCT','NOV','DEC','JAN'} else "YTD"

    idx = MONTHS_454.index(retail_month) + 1

    # Last completed month is one step before
    last_completed = idx - 1
    
    if last_completed <= 0:
        last_completed = len(MONTHS_454)  # wrap around if FEB
    current_month_number = idx

    print(f"Retail year start: , Weeks per month: {wpm}")
    print(f"Current date (week end): {current_date.strftime('%Y-%m-%d')}")
    print(f"Retail month: {retail_month}, Retail week number: {retail_week_number}")
    print(f"Retail year: {retail_year}, Last retail year: {last_retail_year}")
    print(f"Current month number: {current_month_number}, Last completed month: {last_completed}")
    print(f"Season: {season}, Rolling method: {rolling_method}")

    return (
        current_date,
        retail_month,
        current_month_number,
        rolling_method,
        retail_week_number,
        retail_year,
        last_retail_year,
        last_completed,
        season,
        wpm['FEB'], wpm['MAR'], wpm['APR'], wpm['MAY'],
        wpm['JUN'], wpm['JUL'], wpm['AUG'], wpm['SEP'],
        wpm['OCT'], wpm['NOV'], wpm['DEC'], wpm['JAN'],wpm
    )