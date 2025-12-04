

import calendar
import pandas as pd
import re
import math
import json


MONTHS = ['FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC','JAN']
SPRING_MONTHS = ['FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL']

FALL_MONTHS = ['AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN']

#Alreay In Macys

def month_to_num(m):
    """Convert Month name or number to 1-12 int"""
    if pd.isna(m): return pd.NA
    s = str(m).strip()
    # numeric
    if s.isdigit():
        return int(s)
    # month name/abbr
    s = s.lower()
    months = {calendar.month_name[i].lower(): i for i in range(1,13)}
    abbrs  = {calendar.month_abbr[i].lower(): i for i in range(1,13)}
    if s in months: return months[s]
    if s in abbrs: return abbrs[s]
    if len(s)>=3 and s[:3] in abbrs: return abbrs[s[:3]]
    return pd.NA

def week_to_num(w):
    """Extract first integer from Week string"""
    if pd.isna(w): return pd.NA
    m = re.search(r'\d+', str(w))
    return int(m.group()) if m else pd.NA

def get_latest_fiscal(df):
    """
    df must have columns Year, Month, Week.
    Fiscal year starts Feb (month 2) ends Jan (month 1).
    """
    tmp = df.copy()
    tmp['Month_num'] = tmp['Month'].map(month_to_num)
    tmp['Week_num'] = tmp['Week'].map(week_to_num)
    tmp['Year_num'] = pd.to_numeric(tmp['Year'], errors='coerce')

    # compute fiscal_year: Feb–Dec use same Year, Jan belongs to previous Year
    tmp['FiscalYear'] = tmp.apply(
        lambda r: r['Year_num'] - 1 if r['Month_num'] == 1 else r['Year_num'],
        axis=1
    )

    # Sort by FiscalYear, then Month_num (with Feb=2 … Dec=12, Jan=1),
    # then Week_num
    tmp_sorted = tmp.sort_values(['FiscalYear','Month_num','Week_num'])

    # Take the last row
    last = tmp_sorted.iloc[-1]

    return  int(last['Year_num']), last['Month'].upper(),int(last['Week_num'])


def generate_std_period(month_from: str, month_to: str) -> list[str]:
    """Generate a list of 3-letter month abbreviations from month_from to month_to."""
    all_months = list(calendar.month_abbr)[1:]  # ['Jan', ..., 'Dec']
    month_map = {m.upper(): i + 1 for i, m in enumerate(all_months)}

    start_abbr = get_month_abbr(month_from)
    end_abbr = get_month_abbr(month_to)
    start_idx = month_map[start_abbr]
    end_idx = month_map[end_abbr]

    if start_idx <= end_idx:
        selected = all_months[start_idx - 1:end_idx]
    else:
        selected = all_months[start_idx - 1:] + all_months[:end_idx]

    return [m.upper() for m in selected]

def convert_month_to_abbr(month_name):
    """Convert full month name to 3-letter abbreviation."""
    return month_name[:3].upper()

def get_forecast_info(forecast_date):
    """Return forecast date, forecast month (full and abbreviated), and year."""
    forecast_month = forecast_date.strftime("%B")
    forecast_month = convert_month_to_abbr(forecast_month)
    return forecast_month


def get_week_of_month(date):
    """Get the week number of the given date within its month."""
    first_day_of_month = date.replace(day=1)
    days_diff = (date - first_day_of_month).days
    return math.ceil((days_diff + 1) / 7)


def get_month_abbr(month_name: str) -> str:
    """Convert full month name to its 3-letter uppercase abbreviation."""
    try:
        month_index = list(calendar.month_name).index(month_name.capitalize())
        return calendar.month_abbr[month_index].upper()
    except ValueError:
        raise ValueError(f"Invalid month name: {month_name}")


def calculate_std_trend(std_ty_unit_sales_list, std_ly_unit_sales_list):
    """
    Calculate STD trend from LY and TY unit sales.
    """
    # Sum LY and TY unit sales
    ly_total = sum(std_ly_unit_sales_list)
    ty_total = sum(std_ty_unit_sales_list)
 
    # Compute trend if both totals are non-zero
    if ly_total and ty_total:
        std_trend = round((ty_total - ly_total) / ly_total, 2)
    else:
        std_trend = 0
 
    return std_trend


def calculate_std_index_value(index_value_dict,STD_PERIOD):
    """
    Compute rounded index values and total for given months.
    """
    # Round index values for each STD month
    std_period_index_value_list = [index_value_dict[month] for month in STD_PERIOD]
    # Sum the index values
    total_std_period_index_value = round(sum(std_period_index_value_list),2)
 
    return total_std_period_index_value




def calculate_12th_month_forecast(std_ty_unit_sales_list, total_std_period_index_value):
    """
    Forecast 12th month using STD unit sales and index value.
    """
    # Get unit sales for STD months
    total_unit_sales = sum(std_ty_unit_sales_list)
 
    # Calculate forecast if index sum is non-zero
    if total_std_period_index_value:
        month_12_fc_index = round(total_unit_sales / total_std_period_index_value, 0)
    else:
        month_12_fc_index = 0
 
    return month_12_fc_index




def calculate_fc_by_index(index_value, month_12_fc_index):
    """
    Generate forecast for each month using index values.
    """
    fc_by_index = {}
    for month in index_value:   # loop directly on 'FEB', 'MAR', etc.
        fc_by_index[month] = round(index_value[month] * month_12_fc_index, 0)
    return fc_by_index



def calculate_fc_by_trend(s1, k1,f8, row4_values, row17_values, row39_values):
    fc_by_trend = {}
    for month,row4, row17, row39 in zip(MONTHS,row4_values, row17_values, row39_values):
        if s1 == 12 and k1 == 12:
            result = round(row17 + row17 * f8, 0)
        elif s1 == 12 and row4 < 7:
            result = round(row39 + row39 * f8, 0)
        elif s1 > 6 and row4 < 7:
            result = round(row17 + row17 * f8, 0)
        elif s1 < 7 and row4 > 6:
            result = round(row39 + row39 * f8, 0)
        elif s1 < 7 and row4 < 7:
            result = round(row39 + row39 * f8, 0)
        elif s1 > 6 and row4 > 6:
            result = round(row39 + row39 * f8, 0)
        else:
            result = None  # Fallback case
 
        fc_by_trend[month] = result
 
    return fc_by_trend




def get_last_12_months_inventory(current_month, ty_inventory, ly_inventory):
    print("Calculating last 12 months inventory...")
    print(f"Current month: {current_month}")
    print(f"TY Inventory: {ty_inventory}")
    print(f"LY Inventory: {ly_inventory}")

    # Ordered months (retail style, FEB → JAN)
    months = ["FEB", "MAR", "APR", "MAY", "JUN", "JUL",
              "AUG", "SEP", "OCT", "NOV", "DEC", "JAN"]

    # Find index of current month
    cur_idx = months.index(current_month)

    # Collect last 12 months in order
    last_12_months = []
    for i in range(12):
        # Rolling backward from current month
        month_idx = (cur_idx - 11 + i) % len(months)
        month = months[month_idx]

        # Rule: months strictly before current month → TY, else LY
        if (month_idx <= cur_idx - 1) or (cur_idx == 0 and month_idx == len(months) - 1):
            value = ty_inventory.get(month, 0)
        else:
            value = ly_inventory.get(month, 0)

        last_12_months.append(value)

    print("Last 12 months inventory:", last_12_months)
    return last_12_months
def is_maintained(eom_oh_list, threshold, door_count):
    """
    Check if average EOM OH is within threshold or above door count.
    """
    # Calculate average EOM OH
    average_eom_oh = sum(eom_oh_list) / len(eom_oh_list) if eom_oh_list else 0
 
    # Check if it's maintained
    return (average_eom_oh >= threshold * int(door_count)) or (average_eom_oh > int(door_count))
 


def compare_seasonal_forecasts_by_method(fc_by_index, fc_by_trend,season_to_compare_trend_index):
    """
    Compare seasonal forecasts and decide if average method can be used.
    """
    spring_fc_by_index_all = sum(fc_by_index[month] for month in SPRING_MONTHS)
    fall_fc_by_index_all = sum(fc_by_index[month] for month in FALL_MONTHS)
   
    spring_fc_by_trend_all = sum(fc_by_trend[month] for month in SPRING_MONTHS)
    fall_fc_by_trend_all = sum(fc_by_trend[month] for month in FALL_MONTHS)
 
    seasonal_total_fc_by_index = spring_fc_by_index_all if season_to_compare_trend_index == "SPRING" else fall_fc_by_index_all
    seasonal_total_fc_by_trend = spring_fc_by_trend_all if season_to_compare_trend_index == "SPRING" else fall_fc_by_trend_all
 
    # Compare difference
    difference = None
    if seasonal_total_fc_by_index is not None and seasonal_total_fc_by_trend is not None:
        difference = (abs(seasonal_total_fc_by_trend - seasonal_total_fc_by_index) / max(seasonal_total_fc_by_index, seasonal_total_fc_by_trend)) * 100
 
    return difference,seasonal_total_fc_by_trend,seasonal_total_fc_by_index



def decide_forecasting_method(is_inventory_maintained_ly_std_period_store,store_trend_index_difference,std_ly_unit_sales_list):
    # eom_oh=[eom_oh[month] for month in MONTHS]

    #     # Choose forecasting method
    # if is_red_box_item:
    #     forecasting_method = "FC By Index"
    #     print(f"Forecasting method chosen: {forecasting_method}")
    # elif not is_inventory_maintained_ly_std_period and( std_trend>0.75 or std_trend<-0.75):
    #     forecasting_method = "FC By Index"
    #     print(f"Forecasting method chosen: {forecasting_method}")
    # elif std_trend < 0:
    #     if seasonal_total_fc_by_trend > seasonal_total_fc_by_index:
    #         forecasting_method = "FC By Trend"
    #         print(f"Forecasting method chosen: {forecasting_method}")
    #     else:
    #         forecasting_method = "FC By Index"
    #         print(f"Forecasting method chosen: {forecasting_method}")
    # elif trend_index_difference < 25:
    #     forecasting_method = "Average"
    #     print(f"Forecasting method chosen: {forecasting_method}")
    # elif trend_index_difference > 25 and is_inventory_maintained:
    #     forecasting_method = "FC By Trend"
    #     print(f"Forecasting method chosen: {forecasting_method}")
    # else :
    #     if seasonal_total_fc_by_trend > seasonal_total_fc_by_index:
    #         forecasting_method = "FC By Trend"
    #         print(f"Forecasting method chosen: {forecasting_method}")
    #     else:
    #         forecasting_method = "FC By Index"
    #         print(f"Forecasting method chosen: {forecasting_method}")
    if store_trend_index_difference < 25:
        forecasting_method = "Average"
    elif is_inventory_maintained_ly_std_period_store and std_ly_unit_sales_list !=0:
        forecasting_method = "FC By Trend"
    else:
        forecasting_method = "FC By Index"

    return forecasting_method



def calculate_fc_by_average(fc_by_index, fc_by_trend):
    """
    Compute average forecast from index and trend forecasts.
    """
    fc_by_average= {
        key: round((fc_by_index[key] + fc_by_trend[key]) / 2)
        for key in fc_by_trend
    }
    return fc_by_average

def get_recommended_forecast(forecasting_method, fc_by_index, fc_by_trend):
    """
    Return recommended forecast based on selected forecasting method.
    """
    # Select forecast based on method
    if forecasting_method == "FC By Index":
        recommended_fc = fc_by_index
    elif forecasting_method == "FC By Trend":
        recommended_fc = fc_by_trend
    else:
        recommended_fc = calculate_fc_by_average(fc_by_index, fc_by_trend)

    print("Recommended forecast:", recommended_fc)

    return recommended_fc
 



def calculate_loss(door_count, average_value):
    """
    Calculate loss percentage from door count and average store EOM OH.
    """
    if average_value and not np.isnan(average_value):
        return (door_count / average_value) - 1
    else:
        return 0
        
# New function

def get_belk_vendor_info(vendorGuideline, pid):
    row = vendorGuideline.loc[
        vendorGuideline["Belk Style Number"] == pid,
        [
            "RL Style Number",
            "Vendor No_",
            "Vendor Name",
            "Country Of Origin Code",
            "Vendor Lead Time (Days)",
            "Minimum Order Quantity",
        ],
    ]

    if row.empty:
        return None, None, None, None, None, None

    r = row.iloc[0]
    return (
        r["RL Style Number"],
        r["Vendor No_"],
        r["Vendor Name"],
        r["Country Of Origin Code"],
        r["Vendor Lead Time (Days)"],
        r["Minimum Order Quantity"],
    )

from datetime import timedelta

def get_forecast_date(vendor_leadtime, current_date):
    """
    Calculate forecast date based on vendor lead time.

    Parameters:
        vendor_leadtime (int or float): Lead time in days.
        current_date (date or datetime): Current operating date.

    Returns:
        datetime: Forecast date (current_date + leadtime).
    """
    
    if vendor_leadtime is None:
        vendor_leadtime = 60

    return current_date + timedelta(days=int(vendor_leadtime))


def get_index_dict_of_month(index_df,category):

    index_row_data = index_df.loc[index_df['Category'].astype(str).str.lower() == category.lower()]
    months = ['FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC','JAN']

    index_value = {}
    # Loop through each month and fetch its value
    for month in months:
        index_value[month] = ((index_row_data[month].iloc[0])/100) if not index_row_data.empty else 0
    return index_value