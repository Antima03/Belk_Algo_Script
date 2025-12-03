import openpyxl
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation
import time
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import PatternFill,GradientFill
from openpyxl.styles import Border,Side, Font,Alignment
import pandas as pd
import math
from datetime import datetime, timedelta
from calendar import monthrange
from concurrent.futures import ThreadPoolExecutor
import json
import numpy as np
# pid = 'CA5012G017BG0'
start_time = time.time()

 

JSON_PATH = "belk_category_forecasts.json"   # <-- your JSON file

# ---------- LOAD JSON ----------
with open(JSON_PATH, "r", encoding="utf-8") as f:
    json_data = json.load(f)

def get_previous_retail_week():
    """
    Get the previous week's month, year of the previous month,
    last year's occurrence of that month, last month before the previous month in numeric format,
    determine SP (Spring) or FA (Fall) based on the previous month,
    and calculate the number of retail weeks for each month individually.
    """
    # Use the current date as input
    #current_date = datetime.now()
    current_date = datetime(2025, 5, 31)

    # Find the current week's Sunday
    current_sunday = current_date - timedelta(days=current_date.weekday() + 1)

    # Calculate the previous week's Sunday
    previous_week_sunday = current_sunday - timedelta(days=7)

    # Determine the previous week number in the retail month
    previous_week_number = (previous_week_sunday.day - 1) // 7 + 1

    # Get the month and retail year of the previous week
    current_month = previous_week_sunday.strftime('%b')

    # Determine the retail year logic:
    # Retail year starts from February of the previous year and ends in January of the current year
    if previous_week_sunday.month >= 2:
        retail_year_of_previous_month = previous_week_sunday.year
    else:
        retail_year_of_previous_month = previous_week_sunday.year - 1

    # Last year's retail occurrence of the same month
    last_retail_year_of_previous_month = retail_year_of_previous_month - 1
    # Determine the last month before the previous month
    last_month_of_previous_month_date = previous_week_sunday.replace(day=1) - timedelta(days=1)
    last_month = last_month_of_previous_month_date.strftime('%b').upper()

    # Custom mapping for months
    month_mapping = {
        'FEB': 1, 'MAR': 2, 'APR': 3, 'MAY': 4,
        'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8,
        'OCT': 9, 'NOV': 10, 'DEC': 11, 'JAN': 12
    }
    last_month_of_previous_month_numeric = month_mapping[last_month]

    # Determine SP (Spring) or FA (Fall/Winter) based on the previous month
    spring_months = ['FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL']
    fall_months = ['AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN']

    season = "SP" if current_month.upper() in spring_months else "FA"

    # Calculate the number of retail weeks for each month of the current year
    current_year = retail_year_of_previous_month
    # Individual variables for retail weeks of each month


    def get_retail_weeks(year, month):
        """
        Calculate the number of retail weeks in a given month.
        Retail weeks follow the Sunday-to-Saturday structure,
        and all days in a week belong to the month in which the week starts.
    
        Args:
            year (int): The year of the month.
            month (int): The month (1 for January, 12 for December).

        Returns:
            int: Number of retail weeks in the month.
        """
        # Get the first day and last day of the month
        first_day = datetime(year, month, 1)
        last_day = datetime(year, month, monthrange(year, month)[1])

        # Find the first Sunday of the month
        first_sunday = first_day + timedelta(days=(6 - first_day.weekday()) % 7)

        # Find the last Saturday of the month
        last_saturday = last_day - timedelta(days=last_day.weekday() + 1)

        # Count retail weeks
        current_week_start = first_sunday
        week_count = 0

        while current_week_start <= last_saturday:
            week_count += 1
            current_week_start += timedelta(days=7)  # Move to the next Sunday

        # Check if the final week starts in the current month (partial week rule)
        if current_week_start <= last_day:
            week_count += 1

        return week_count

    feb_weeks = get_retail_weeks(current_year,2)
    mar_weeks = get_retail_weeks(current_year,3)
    apr_weeks = get_retail_weeks(current_year,4)
    may_weeks = get_retail_weeks(current_year,5)
    jun_weeks = get_retail_weeks(current_year,6)
    jul_weeks = get_retail_weeks(current_year,7)
    aug_weeks = get_retail_weeks(current_year,8)
    sep_weeks = get_retail_weeks(current_year,9)
    oct_weeks = get_retail_weeks(current_year,10)
    nov_weeks = get_retail_weeks(current_year,11)
    dec_weeks = get_retail_weeks(current_year,12)
    jan_weeks = get_retail_weeks(current_year + 1, 1)
    month_dict = {
    "Feb": 1,
    "Mar": 2,
    "Apr": 3,
    "May": 4,
    "Jun": 5,
    "Jul": 6,
    "Aug": 7,
    "Sep": 8,
    "Oct": 9,
    "Nov": 10,
    "Dec": 11,
    "Jan": 12
}  # January belongs to the next year
    current_month_number = month_dict.get(current_month, "Month not found")
    if current_month in [ "Oct","Nov","Dec","Jan"]:
        rolling_method="Current MTH"
    else:
        rolling_method="YTD"
    return current_month,current_month_number,rolling_method, previous_week_number, retail_year_of_previous_month,last_retail_year_of_previous_month, last_month_of_previous_month_numeric,season, feb_weeks, mar_weeks, apr_weeks, may_weeks,jun_weeks, jul_weeks, aug_weeks, sep_weeks, oct_weeks,nov_weeks, dec_weeks, jan_weeks
    # return current_month, previous_week_number, year_of_previous_month, last_year_of_previous_month, last_month_of_previous_month_numeric, season

current_month,current_month_number,rolling_method, previous_week_number, year_of_previous_month,last_year_of_previous_month, last_month_of_previous_month_numeric,season, feb_weeks, mar_weeks, apr_weeks, may_weeks,jun_weeks, jul_weeks, aug_weeks, sep_weeks, oct_weeks,nov_weeks, dec_weeks, jan_weeks = get_previous_retail_week()
    # return current_month, previous_week_number, year_of_previous_month, last_year_of_previous_month, last_month_of_previous_month_numeric, season


categories = {
    # "BELK & CO PWP": 14,
    "BONDED GOLD SS": 26,
    "cOLOR PRECIOUS": 1067,
    # "COLOR SEMI-PRECIOUS": 1848,
    # "DIAMOND GOLD": 1795,
    # "DIAMOND STERLING SILVER": 901,
    # "FJ MEMO DIAMOND COLLECTIONS": 4,
    # "GOLD": 2075,
    # "LAB GROWN DIAMOND GOLD": 302,
    # "LAB GROWN DIAMOND STERLING SILVER": 242,
    # "PEARL": 843,
    # "SILVER": 843
}


month_dict = {
    "Feb": 1,
    "Mar": 2,
    "Apr": 3,
    "May": 4,
    "Jun": 5,
    "Jul": 6,
    "Aug": 7,
    "Sep": 8,
    "Oct": 9,
    "Nov": 10,
    "Dec": 11,
    "Jan": 12
}
current_month_number = month_dict.get(current_month, "Month not found")
if current_month in [ "Oct","Nov","Dec","Jan","Feb","Mar","Apr","May"]:
    rolling_method="Current MTH"
else:
    rolling_method="YTD"

for category, num_products in categories.items():
    # Get PIDs for this category from JSON
    category_data = json_data.get(category, {})
    pids_list = list(category_data.keys())
    
    # Update the data variable with category-specific values
    data = [
        ["", "TY", year_of_previous_month, "LY", last_year_of_previous_month, "Season", season, "Current Year", "Month", current_month, current_month_number, "Week", previous_week_number, "", "MAY-SEP", "", "", "Last Completed Month", last_month_of_previous_month_numeric, "", "Use EOM Actual?", rolling_method],
        ["", "Count of Items", "", 8215, "", "", "", "Last SP / FA Months", "Month", "Jul", "", "Jan", 12, "Sorted by:", "Dept Grouping >Class ID", "", "", ""],
        ["", "", "", "", "", "", "", "# of Wks in Mth", feb_weeks, mar_weeks, apr_weeks, may_weeks, jun_weeks, jul_weeks, aug_weeks, sep_weeks, oct_weeks, nov_weeks, dec_weeks, jan_weeks],
        ["", category.upper(), "", "", "Avg Sales 1st & last Mth", 8, 11, "Month #", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, "", "", ""]
    ]

    output_file = f"{category,"_whole"}.xlsx"
    # Initialize workbook and create sheets
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws_index = wb.create_sheet(title="Index")
    ws_month = wb.create_sheet(title="Month")
    ws_dropdown = wb.create_sheet(title="DropdownData")


    # Step 4: Populate Worksheet with Data
    for row_num, row_data in enumerate(data, 1):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    # Step 4: Freeze Top 4 Rows
    ws.freeze_panes = ws["A5"]
    # Define options for dropdowns

    ws['C2'] = "=B2*51+4"  # Calculate row based on count of items
    ws['D2'] = "=B2*51+4"  # Calculate row based on count of items
    ws['K1'] = "=VLOOKUP(J1,Month!A:B,2,0)"
    ws['K2'] = "=VLOOKUP(J2,Month!A:B,2,0)"
    ws['M2'] = "=VLOOKUP(L2,Month!A:B,2,0)"
    # Define additional dropdown options

    # Define your dropdown options (38 items)
    dropdown_options = [
        "BT", "Citrine", "Cross", "CZ", "Dia", "Ear", "EMER", "Garnet", "Gem",
        "GEM EAR", "Gold Chain", "GOLD EAR", "Amy", "Anklet", "Aqua", "Bridal",
        "Heart", "Heavy Gold Chain", "Jade", "KIDS", "Locket", "Mens Gold Bracelet",
        "Mens Misc", "Mens Silver chain", "Mom", "MOP", "Neck", "Onyx", "Opal",
        "Pearl", "Peridot", "Religious", "Ring", "Ruby", "Saph", "Womens Silver Chain",
        "Wrist", "Grand Total"
    ]

    # Write the dropdown options to the "DropdownData" sheet in column A
    for i, option in enumerate(dropdown_options, start=1):
        ws_dropdown.cell(row=i, column=1, value=option)

    # Define the named range for the dropdown options
    # Make sure the range covers exactly the 38 items (A1 to A38)
    named_range = DefinedName(name="DropdownOptions", attr_text="DropdownData!$A$1:$A$38")
    wb.defined_names.add(named_range)

    # Create a data validation that references the named range
    dropdown = DataValidation(type="list", formula1="DropdownOptions", allow_blank=True)
    dropdown.prompt = "Please select an option"
    dropdown.promptTitle = "Dropdown List"



    for loop in range(num_products):
        start_row = 5 + (27 * loop)  # Adjust the range as needed for your loop
        ws.add_data_validation(dropdown)
        dropdown.add(ws[f"F{start_row + 1}"])  # Apply to column F as an example

    forecast_method_options = ["FC by Index", "FC by Trend", "Average", "Current Year", "Last Year"]


    season_option = ["FA", "SP"]
    month_option = ["Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan"]
    year_option = ["Current MTH", "YTD", "SPRING", "FALL", "LY FALL"]


    # Function to add a dropdown to a specific cell
    def add_dropdown(ws, cell, options):
        # Create a data validation object with the list of options
        dropdown = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=True)
        dropdown.prompt = "Please select an option"
        dropdown.promptTitle = "Dropdown List"
        # Apply the data validation dropdown to the specified cell
        ws.add_data_validation(dropdown)
        dropdown.add(ws[cell])
    # Add dropdowns to specified cells
    add_dropdown(ws, "G1", season_option)     # G1 for Season
    add_dropdown(ws, "V1", year_option)       # V1 for Year options

    # Add dropdowns to multiple cells for Months
    for cell in ["J1", "J2", "L2"]:
        add_dropdown(ws, cell, month_option)




    ALL_VALUES = [
        "PID/BLU/MKST", "Current FC Index", "(TY/LY) STD Sales Index/12M FC", "STD Trend / 12M FC",
        "Item Status/Forecasting Method/Safe", "Current Str Cnt/Last Str Cnt/Last Updated",
        "Store Model/Com Model/TTL Model", "MA Proj/Proj Ball/Holiday Bd FC",
        "Proj Qty & $ to Release / Note", "Vendor/Min order", "RL TTL/Net Proj/ORD Unalloc/+/− to Model",
        "KPI DATA", "Last KPI Door count", "Diff to Current Door",
        "Out of Stock Locations", "Suspended Location count",
        "Click to View online", "Sub Class", "Masterstyle", "PID Desc",
        "COM 1st Live/Live Site/V2C/WebID/STD Rtn", "Web ID Description",
        "Last Reviewed Date/Code/Qty to Enter", "Current Review Comments"
    ]

    H_VALUES = [
        'ROLLING 12M FC', 'Index', 'FC by Index', 'FC by Trend', 'Recommended FC', 'Planned FC',
        'Planned Shipments', 'Planned EOH (Cal)', 
        'Planned Sell thru %', f"TOTAL {year_of_previous_month}",'Total Sales Units',
         'TOTAL EOM OH',
        'Omni Sales $',
        'Omni Sell Thru %', 'Omni Turn',
        'TY Store Sales U vs LY',  'TY Store EOH vs LY',
         f"TOTAL {last_year_of_previous_month}",
        "Total Sales Units",
        "TOTAL EOM OH",
        "Omni Sell Thru %",
        "Omni Turn",
        "Omni Sales $",
    ]

    MONTHLY_VALUES = ['FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT',
                    'NOV', 'DEC', 'JAN', 'ANNUAL', 'SPRING', 'FALL']
    def apply_round_format(ws, cell_ranges, decimal_places):
        for cell_range in cell_ranges:
            # Check if the cell range is a single cell
            if ":" in cell_range:
                # This is a range, so we can iterate through it
                for row in ws[cell_range]:
                    for cell in row:
                        # Check if cell contains a formula by verifying if the value starts with "="
                        if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                            # Wrap formula in ROUND with the specified decimal places
                            cell.value = f"=ROUND({cell.value[1:]}, {decimal_places})"
            else:
                # This is a single cell reference
                cell = ws[cell_range]
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    # Wrap formula in ROUND with the specified decimal places
                    cell.value = f"=ROUND({cell.value[1:]}, {decimal_places})"

    def apply_format(ws, cell_ranges, number_format):
        for cell_range in cell_ranges:
            # Check if the cell range is a single cell or a range
            if ":" in cell_range:
                # This is a range, so we can iterate through it
                for row in ws[cell_range]:
                    for cell in row:
                        cell.number_format = number_format
            else:
                # This is a single cell
                cell = ws[cell_range]
                cell.number_format = number_format


    # Open the source workbook
    
        for col_num, value in enumerate(row_data, start=1):
            ws_index.cell(row=row_num, column=col_num, value=value)

    # Step 5: Apply Percentage Formatting to B4:P41
    for row in ws_index.iter_rows(min_row=4, max_row=41, min_col=2, max_col=16):
        for cell in row:
            if isinstance(cell.value, (int, float)) and 0 <= cell.value <= 1:
                cell.number_format = '0.00%'  # Format as percentage


    # Step 6: Apply Filters to A3:P3
    ws_index.auto_filter.ref = "A3:P3"


    month_data = [
        ["All", ""],
        ["Feb", 1],
        ["Mar", 2],
        ["Apr", 3],
        ["May", 4],
        ["Jun", 5],
        ["Jul", 6],
        ["Aug", 7],
        ["Sep", 8],
        ["Oct", 9],
        ["Nov", 10],
        ["Dec", 11],
        ["Jan", 12],
    ]

    # Write the data to the "Month" sheet
    for row_num, (month, number) in enumerate(month_data, start=1):
        ws_month.cell(row=row_num, column=1, value=month)  # Column A
        ws_month.cell(row=row_num, column=2, value=number)  # Column B

    border_color = "D9D9D9"  # Gridline color
    gridline1 = Border(
        left=Side(style="thin", color=border_color),
        right=Side(style="thin", color=border_color),
        top=Side(style="thin", color=border_color),
        bottom=Side(style="thin", color=border_color)
    )

    # Create PatternFill objects for the colors
    light_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    light_yellow = PatternFill(start_color="FFEB9C" , end_color="FFEB9C" , fill_type="solid")
    light_pink = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    # Apply the fills to specific cells
    ws['C1'].fill = light_gray
    ws['E1'].fill = light_gray
    ws['G1'].fill = light_gray
    ws['J1'].fill = light_gray
    ws['K1'].fill = light_gray
    ws['M1'].fill = light_gray
    ws['S1'].fill = light_gray
    ws['J2'].fill = light_gray
    ws['L2'].fill = light_gray

    ws['K2'].fill = light_yellow
    ws['O1'].fill = light_yellow
    ws['M2'].fill = light_yellow

    ws['B2'].fill = yellow_fill
    ws['C2'].fill = yellow_fill
    ws['V1'].fill = light_pink

    #no boader
    # Define white border style
    white_border = Border(
        left=Side(style='thin', color="FFFFFF"),
        right=Side(style='thin', color="FFFFFF"),
        top=Side(style='thin', color="FFFFFF"),
        bottom=Side(style='thin', color="FFFFFF")
    )

    # Apply no border to the range B1:W2
    for row in range(1, 3):  # Rows 1 to 2
        for col in range(2, 24):  # Columns B (2) to W (23)
            ws.cell(row=row, column=col).border = white_border

    # Remove borders for specific cells to show gridlines
    # Define the border style
    border = Border(
        left=Side(style='thin', color='7F7F7F'),  # Thin black border on the left
        right=Side(style='thin', color='7F7F7F'),  # Thin black border on the right
        top=Side(style='thin', color='7F7F7F'),  # Thin black border on the top
        bottom=Side(style='thin', color='7F7F7F')  # Thin black border on the bottom
    )
    #no_boader list
    no_boader_list=['C1','E1','G1','J1','K1','L1','M1','O1','S1','V1','J2','K2','L2','M2']
    for i in no_boader_list:
        ws[i].border = border

    #font color
    red_font_italic= Font(color="FF0000", italic=True) 
    yellow_font = Font(color="9C5700") 
    blue_bold_font = Font(color="0563C1",bold=True) 
    ws['B4'].font = blue_bold_font
    ws['C4'].font = blue_bold_font
    ws['B4'].font = Font(bold=True)
    ws['C4'].font = Font(bold=True)
    ws['K2'].font = yellow_font
    ws['O1'].font = yellow_font
    ws['M2'].font = yellow_font

    ws['H1'].font = red_font_italic
    ws['H2'].font = red_font_italic
    ws['O2'].font = red_font_italic
    # Define the alignment (right-aligned)
    right_alignment = Alignment(horizontal='right')
    ws['H1'].alignment = right_alignment
    ws['H2'].alignment = right_alignment
    # Apply no border to the cells in the no_boader list
    bold_font_list=['C1','E1','G1','J1','K1','M1','S1','V1','L2','J2','C2','E4','F4','G4','H4',]
    for i in bold_font_list:
        ws[i].font = Font(bold=True)
    for row in ws["H3:W4"]:  # Use the dynamically calculated range
        for cell in row:
            cell.border = gridline1 

    ws.column_dimensions['c'].width = 20
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['H'].width = 25
    
    for loop in range(min(num_products, len(pids_list))):
        start_row = 5 + (27 * loop)
        g_value = loop + 1
        
        # Get PID for this iteration
        pid = pids_list[loop] if loop < len(pids_list) else "kl"
        
        # Extract data from JSON for this PID
        pid_data = category_data.get(pid, {})
        forecast_data = pid_data.get("forecast", {})
        loader_data = pid_data.get("loader", {})
        
        # Find the matching row

            # Function to get the 3-letter month abbreviation based on the month number
        # Retail calendar month sequence
        retail_months = ["Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan"]
        
        def convert_month_to_abbr(full_month):
            # Define mapping of full month names to abbreviations
            month_mapping = {
                "January": "Jan",
                "February": "Feb",
                "March": "Mar",
                "April": "Apr",
                "May": "May",
                "June": "Jun",
                "July": "Jul",
                "August": "Aug",
                "September": "Sep",
                "October": "Oct",
                "November": "Nov",
                "December": "Dec"
            }
                
            # Return the corresponding abbreviation, or None if not found
            return month_mapping.get(full_month, None)
        

        def calculate_forecast_months(lead_time_weeks, current_date,current_month):
            # Calculate the lead time in days
            lead_time_days = lead_time_weeks * 7
            
            # Calculate the forecast date
            forecast_date = current_date + timedelta(days=lead_time_days)
            
            # Extract the forecast month and year
            forecast_month = forecast_date.strftime("%B")  # Full month name
            forecast_year = forecast_date.year
            forecast_month_abbr = convert_month_to_abbr(forecast_month)

            check = False
            # Calculate the week of the forecast month
            first_day_of_month = forecast_date.replace(day=1)
            days_diff = (forecast_date - first_day_of_month).days
            week_of_forecast_month = math.ceil((days_diff + 1) / 7)  # +1 to include the current day in the week count

            if week_of_forecast_month > 2 :
                check = True
            
            # Calculate the week of the current month
            first_day_of_current_month = current_date.replace(day=1)
            days_diff_current = (current_date - first_day_of_current_month).days
            week_of_current_month = math.ceil((days_diff_current + 1) / 7)

            # Define the month names in 3-letter format
            #month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            
            # # Create the list of months
            # start = month_names.index(current_month) + 1
            # end = month_names.index(forecast_month_abbr) + 1
            forecast_month_list =  [forecast_month_abbr]



            return forecast_month_list , week_of_forecast_month , forecast_month_abbr , check


        retail_months = ["Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan"]
        #Step 2 :  Find STD period###################
        
        # Extract index_value from JSON
        index_value = forecast_data.get("index_value", {
            "FEB": 0.0, "MAR": 0.0, "APR": 0.0, "MAY": 0.0,
            "JUN": 0.0, "JUL": 0.0, "AUG": 0.0, "SEP": 0.0,
            "OCT": 0.0, "NOV": 0.0, "DEC": 0.0, "JAN": 0.0
        })
        
        # Extract TY_Unit_Sales from ty_lw_sls_u
        ty_lw_sls_u = loader_data.get("ty_lw_sls_u", {})
        print(ty_lw_sls_u)
        TY_Unit_Sales = {
            "FEB": float(ty_lw_sls_u.get("FEB", 0)),
            "MAR": float(ty_lw_sls_u.get("MAR", 0)),
            "APR": float(ty_lw_sls_u.get("APR", 0)),
            "MAY": float(ty_lw_sls_u.get("MAY", 0)),
            "JUN": float(ty_lw_sls_u.get("JUN", 0)),
            "JUL": float(ty_lw_sls_u.get("JUL", 0)),
            "AUG": float(ty_lw_sls_u.get("AUG", 0)),
            "SEP": float(ty_lw_sls_u.get("SEP", 0)),
            "OCT": float(ty_lw_sls_u.get("OCT", 0)),
            "NOV": float(ty_lw_sls_u.get("NOV", 0)),
            "DEC": float(ty_lw_sls_u.get("DEC", 0)),
            "JAN": float(ty_lw_sls_u.get("JAN", 0))
        }
        
        # Extract LY_Unit_Sales from ly_lw_sls_u
        ly_lw_sls_u = loader_data.get("ly_lw_sls_u", {})
        LY_Unit_Sales = {
            "FEB": float(ly_lw_sls_u.get("FEB", 0)),
            "MAR": float(ly_lw_sls_u.get("MAR", 0)),
            "APR": float(ly_lw_sls_u.get("APR", 0)),
            "MAY": float(ly_lw_sls_u.get("MAY", 0)),
            "JUN": float(ly_lw_sls_u.get("JUN", 0)),
            "JUL": float(ly_lw_sls_u.get("JUL", 0)),
            "AUG": float(ly_lw_sls_u.get("AUG", 0)),
            "SEP": float(ly_lw_sls_u.get("SEP", 0)),
            "OCT": float(ly_lw_sls_u.get("OCT", 0)),
            "NOV": float(ly_lw_sls_u.get("NOV", 0)),
            "DEC": float(ly_lw_sls_u.get("DEC", 0)),
            "JAN": float(ly_lw_sls_u.get("JAN", 0))
        }
        
        # Extract LY_OH_Units from ly_lw_eop_oh_u
        ly_lw_eop_oh_u = loader_data.get("ly_lw_eop_oh_u", {})
        LY_OH_Units = {
            "FEB": float(ly_lw_eop_oh_u.get("FEB", 0)),
            "MAR": float(ly_lw_eop_oh_u.get("MAR", 0)),
            "APR": float(ly_lw_eop_oh_u.get("APR", 0)),
            "MAY": float(ly_lw_eop_oh_u.get("MAY", 0)),
            "JUN": float(ly_lw_eop_oh_u.get("JUN", 0)),
            "JUL": float(ly_lw_eop_oh_u.get("JUL", 0)),
            "AUG": float(ly_lw_eop_oh_u.get("AUG", 0)),
            "SEP": float(ly_lw_eop_oh_u.get("SEP", 0)),
            "OCT": float(ly_lw_eop_oh_u.get("OCT", 0)),
            "NOV": float(ly_lw_eop_oh_u.get("NOV", 0)),
            "DEC": float(ly_lw_eop_oh_u.get("DEC", 0)),
            "JAN": float(ly_lw_eop_oh_u.get("JAN", 0))
        }
        
        # Extract TY_OH_Units from ty_lw_eop_oh_u
        ty_lw_eop_oh_u = loader_data.get("ty_lw_eop_oh_u", {})
        TY_OH_Units = {
            "FEB": float(ty_lw_eop_oh_u.get("FEB", 0)),
            "MAR": float(ty_lw_eop_oh_u.get("MAR", 0)),
            "APR": float(ty_lw_eop_oh_u.get("APR", 0)),
            "MAY": float(ty_lw_eop_oh_u.get("MAY", 0)),
            "JUN": float(ty_lw_eop_oh_u.get("JUN", 0)),
            "JUL": float(ty_lw_eop_oh_u.get("JUL", 0)),
            "AUG": float(ty_lw_eop_oh_u.get("AUG", 0)),
            "SEP": float(ty_lw_eop_oh_u.get("SEP", 0)),
            "OCT": float(ty_lw_eop_oh_u.get("OCT", 0)),
            "NOV": float(ty_lw_eop_oh_u.get("NOV", 0)),
            "DEC": float(ty_lw_eop_oh_u.get("DEC", 0)),
            "JAN": float(ty_lw_eop_oh_u.get("JAN", 0))
        }
        
        # Extract LY_sales_amt and TY_sales_amt from ty_lw_sls_dollar and ly_lw_sls_dollar
        ly_lw_sls_dollar = loader_data.get("ly_lw_sls_dollar", {})
        LY_sales_amt = {
            "FEB": float(ly_lw_sls_dollar.get("FEB", 0)),
            "MAR": float(ly_lw_sls_dollar.get("MAR", 0)),
            "APR": float(ly_lw_sls_dollar.get("APR", 0)),
            "MAY": float(ly_lw_sls_dollar.get("MAY", 0)),
            "JUN": float(ly_lw_sls_dollar.get("JUN", 0)),
            "JUL": float(ly_lw_sls_dollar.get("JUL", 0)),
            "AUG": float(ly_lw_sls_dollar.get("AUG", 0)),
            "SEP": float(ly_lw_sls_dollar.get("SEP", 0)),
            "OCT": float(ly_lw_sls_dollar.get("OCT", 0)),
            "NOV": float(ly_lw_sls_dollar.get("NOV", 0)),
            "DEC": float(ly_lw_sls_dollar.get("DEC", 0)),
            "JAN": float(ly_lw_sls_dollar.get("JAN", 0))
        }
        
        ty_lw_sls_dollar = loader_data.get("ty_lw_sls_dollar", {})
        TY_sales_amt = {
            "FEB": float(ty_lw_sls_dollar.get("FEB", 0)),
            "MAR": float(ty_lw_sls_dollar.get("MAR", 0)),
            "APR": float(ty_lw_sls_dollar.get("APR", 0)),
            "MAY": float(ty_lw_sls_dollar.get("MAY", 0)),
            "JUN": float(ty_lw_sls_dollar.get("JUN", 0)),
            "JUL": float(ty_lw_sls_dollar.get("JUL", 0)),
            "AUG": float(ty_lw_sls_dollar.get("AUG", 0)),
            "SEP": float(ty_lw_sls_dollar.get("SEP", 0)),
            "OCT": float(ty_lw_sls_dollar.get("OCT", 0)),
            "NOV": float(ty_lw_sls_dollar.get("NOV", 0)),
            "DEC": float(ty_lw_sls_dollar.get("DEC", 0)),
            "JAN": float(ty_lw_sls_dollar.get("JAN", 0))
        }
        
        # Extract planned_fc and planned_shp from fc_by_index
        fc_by_index = forecast_data.get("fc_by_index", {})
        planned_fc = {
            "FEB": float(fc_by_index.get("FEB", 0)),
            "MAR": float(fc_by_index.get("MAR", 0)),
            "APR": float(fc_by_index.get("APR", 0)),
            "MAY": float(fc_by_index.get("MAY", 0)),
            "JUN": float(fc_by_index.get("JUN", 0)),
            "JUL": float(fc_by_index.get("JUL", 0)),
            "AUG": float(fc_by_index.get("AUG", 0)),
            "SEP": float(fc_by_index.get("SEP", 0)),
            "OCT": float(fc_by_index.get("OCT", 0)),
            "NOV": float(fc_by_index.get("NOV", 0)),
            "DEC": float(fc_by_index.get("DEC", 0)),
            "JAN": float(fc_by_index.get("JAN", 0))
        }
        
        # planned_shp is same as planned_fc
        planned_shp = planned_fc.copy()
        
        # Extract other forecast parameters
        month_12_fc_index = forecast_data.get("month_12_fc_index", 0)
        std_trend = forecast_data.get("std_trend", 0)
        std_index_value = forecast_data.get("std_index_value", 0)
        std_ty_unit_sales_list = [float(x) for x in forecast_data.get("std_ty_unit_sales_list", [0, 0, 0, 0])]
        std_ly_unit_sales_list = [float(x) for x in forecast_data.get("std_ly_unit_sales_list", [0, 0, 0, 0])]
        forecasting_method = forecast_data.get("forecasting_method", "FC by Index")
        std_period = forecast_data.get("std_period", ["FEB", "MAR", "APR", "MAY"])
        pid = forecast_data.get("pid", 0)
        dynamic_formulas = {
            f"G{start_row + 18}": 1,
            f"F{start_row + 1}": f"=C{start_row + 1}",
            f"C{start_row}": pid,
            f"D{start_row}":"",
            f"F{start_row}":'MKST',
            f"O1":std_period[0]+"-"+std_period[-1],
            f"C{start_row + 1}":'Current_FC_Index',
            f"C{start_row + 2}": sum(std_ty_unit_sales_list),
            f"D{start_row + 2}": sum(std_ly_unit_sales_list),
            f"C{start_row + 3}": f"=IFERROR(ROUND((C{start_row + 2}-D{start_row + 2})/D{start_row + 2},2),0)",
            f"F{start_row + 3}": std_trend,
            f"E{start_row + 3}": "Chg Trend",
            f"D{start_row + 1}": "Change Index",
            f"E{start_row + 2}": std_index_value,
            f"F{start_row + 2}": month_12_fc_index,  # Added IFERROR to prevent divide-by-zero issues
            f"F{start_row + 4}": forecasting_method,
            f"C{start_row + 4}":  'Item_Status',
            f"C{start_row + 5}": 'Door_Count',
            f"D{start_row + 5}": 'Last_Str_Cnt',
            f"E{start_row + 5}": f"=ROUND(C{start_row + 5}-D{start_row + 5},2)",
            f"F{start_row + 5}": 'Door_count_Updated',
            f"C{start_row + 6}": 'Store_Model',
            f"D{start_row + 6}": 'Com_Model',
            f"E{start_row + 6}": f"=ROUND(C{start_row + 6}+D{start_row + 6},2)",
            f"G{start_row + 7}": 'Holiday_Bld_FC',
            f"C{start_row + 8}":'MCYOH',
            f"D{start_row + 8}": 'OO',
            f"E{start_row + 8}": f"=IF(D{start_row + 8}-E{start_row + 14}<0,0,D{start_row + 8}-E{start_row + 14})",
            f"E{start_row + 14}": 'nav_OO',
            f"F{start_row + 8}": 'MTD_SHIPMENTS',
            f"G{start_row + 8}":'LW_Shipments',
            f"C{start_row + 9}": 'Wks_of_Stock_OH',
            f"D{start_row + 9}":'Wks_of_on_Proj',
            f"F{start_row + 9}": 'Last_3Wks_Ships',
            f"C{start_row + 10}": "=0",
            f"C{start_row + 13}": 'Vendor_Name',
            f"G{start_row + 13}": 'Min_order',
            f"C{start_row + 14}": 'Proj',
            f"D{start_row + 14}": 'Net_Proj',
            f"F{start_row + 14}": 'Unalloc_Orders',
            f"G{start_row + 14}": f"=IF(C{start_row + 5}>G{start_row + 5},(C{start_row + 8}+D{start_row + 8})-C{start_row + 5},IF(OR(E{start_row + 6}>C{start_row + 5},E{start_row + 6}>G{start_row + 5}),(C{start_row + 8}+D{start_row + 8})-E{start_row + 6},(C{start_row + 8}+D{start_row + 8})-G{start_row + 5}))",
            f"C{start_row + 15}": 'RLJ_OH',
            f"D{start_row + 15}": 'FLDC',
            f"E{start_row + 15}": 'WIP',
            f"C{start_row + 11}": f"=C{start_row + 14}+F{start_row + 14}-C{start_row + 15}-E{start_row + 15}",
            f"D{start_row + 11}": f"=IF(AND(F{start_row + 14}>=C{start_row + 11},C{start_row + 11}>0),\"Demand due to Unalloc Sales Orders\",IF(C{start_row + 11}<0,\"Excess OH Units\",IF(AND(C{start_row + 11}>0,F{start_row + 14}<C{start_row + 11}),\"Demand\",\"\")))",
            f"C{start_row + 17}": 'MD_Status_MZ1',
            f"D{start_row + 17}":'Repl_Flag',
            f"E{start_row + 17}": 'MCOM_RPL',
            f"F{start_row + 17}": 'Pool_stock',
            f"C{start_row + 18}": 'st_Rec_Date',
            f"D{start_row + 18}": 'Last_Rec_Date',
            f"E{start_row + 18}": 'Item_Age',
            f"F{start_row + 18}": f"=IFERROR((NOW()-VALUE(C{start_row + 18}))/30,0)",
            f"C{start_row + 19}": 'TY_Last_Cost',
            f"D{start_row + 19}": 'Own_Retail',
            f"E{start_row + 19}": 'AWR_1st_Tkt_Ret',
            f"F{start_row + 19}": f"=(D{start_row + 19}-C{start_row + 19})/D{start_row + 19}",
            f"C{start_row + 20}": 'Metal_Lock',
            f"D{start_row + 20}":'MFG_Policy',
            f"C{start_row + 21}": 'KPI_Data_Updated',
            f"C{start_row + 22}": 'KPI_Door_count',
            f"D{start_row + 21}": f"=SUBSTITUTE(D{start_row},\"/\",\"\")",
            f"I{start_row + 1}": index_value['FEB'],
            f"J{start_row + 1}": index_value['MAR'],
            f"K{start_row + 1}": index_value['APR'],
            f"L{start_row + 1}": index_value['MAY'],
            f"M{start_row + 1}": index_value['JUN'],
            f"N{start_row + 1}": index_value['JUL'],
            f"O{start_row + 1}": index_value['AUG'],
            f"P{start_row + 1}": index_value['SEP'],
            f"Q{start_row + 1}": index_value['OCT'],
            f"R{start_row + 1}": index_value['NOV'],
            f"S{start_row + 1}": index_value['DEC'],
            f"T{start_row + 1}": index_value['JAN'],
            f"U{start_row + 1}": f"=SUM(I{start_row + 1}:T{start_row + 1})",
            f"V{start_row + 1}": f"=IFERROR(SUM(I{start_row + 1}:N{start_row + 1}),0)",
            f"W{start_row + 1}": f"=IFERROR(SUM(O{start_row + 1}:T{start_row + 1}),0)",
            f"I{start_row + 2}": f"=ROUND($F{start_row + 2}*I{start_row + 1},0)",
            f"J{start_row + 2}": f"=ROUND($F{start_row + 2}*J{start_row + 1},0)",
            f"K{start_row + 2}": f"=ROUND($F{start_row + 2}*K{start_row + 1},0)",
            f"L{start_row + 2}": f"=ROUND($F{start_row + 2}*L{start_row + 1},0)",
            f"M{start_row + 2}": f"=ROUND($F{start_row + 2}*M{start_row + 1},0)",
            f"N{start_row + 2}": f"=ROUND($F{start_row + 2}*N{start_row + 1},0)",
            f"O{start_row + 2}": f"=ROUND($F{start_row + 2}*O{start_row + 1},0)",
            f"P{start_row + 2}": f"=ROUND($F{start_row + 2}*P{start_row + 1},0)",
            f"Q{start_row + 2}": f"=ROUND($F{start_row + 2}*Q{start_row + 1},0)",
            f"R{start_row + 2}": f"=ROUND($F{start_row + 2}*R{start_row + 1},0)",
            f"S{start_row + 2}": f"=ROUND($F{start_row + 2}*S{start_row + 1},0)",
            f"T{start_row + 2}": f"=ROUND($F{start_row + 2}*T{start_row + 1},0)",
            f"U{start_row + 2}": f"=SUM(I{start_row + 2}:T{start_row + 2})",
            f"V{start_row + 2}": f"=IFERROR(SUM(I{start_row + 2}:N{start_row + 2}),0)",
            f"W{start_row + 2}": f"=IFERROR(SUM(O{start_row + 2}:T{start_row + 2}),0)",
            f"I{start_row + 13}": f"=IFERROR(IFERROR(I{start_row + 10}/(I{start_row + 10}+I{start_row + 11}),0),0)",
            f"J{start_row + 13}": f"=IFERROR(IFERROR(J{start_row + 10}/(J{start_row + 10}+J{start_row + 11}),0),0)",
            f"K{start_row + 13}": f"=IFERROR(IFERROR(K{start_row + 10}/(K{start_row + 10}+K{start_row + 11}),0),0)",
            f"L{start_row + 13}": f"=IFERROR(IFERROR(L{start_row + 10}/(L{start_row + 10}+L{start_row + 11}),0),0)",
            f"M{start_row + 13}": f"=IFERROR(IFERROR(M{start_row + 10}/(M{start_row + 10}+M{start_row + 11}),0),0)",
            f"N{start_row + 13}": f"=IFERROR(IFERROR(N{start_row + 10}/(N{start_row + 10}+N{start_row + 11}),0),0)",
            f"O{start_row + 13}": f"=IFERROR(IFERROR(O{start_row + 10}/(O{start_row + 10}+O{start_row + 11}),0),0)",
            f"P{start_row + 13}": f"=IFERROR(IFERROR(P{start_row + 10}/(P{start_row + 10}+P{start_row + 11}),0),0)",
            f"Q{start_row + 13}": f"=IFERROR(IFERROR(Q{start_row + 10}/(Q{start_row + 10}+Q{start_row + 11}),0),0)",
            f"R{start_row + 13}": f"=IFERROR(IFERROR(R{start_row + 10}/(R{start_row + 10}+R{start_row + 11}),0),0)",
            f"S{start_row + 13}": f"=IFERROR(IFERROR(S{start_row + 10}/(S{start_row + 10}+S{start_row + 11}),0),0)",
            f"T{start_row + 13}": f"=IFERROR(IFERROR(T{start_row + 10}/(T{start_row + 10}+T{start_row + 11}),0),0)",

            f"U{start_row + 13}": f"=IFERROR(IFERROR(U{start_row + 10}/(U{start_row + 10}+U{start_row + 11}-U{start_row + 8}),0),0)",
            f"V{start_row + 13}": f"=IFERROR(IFERROR(V{start_row + 10}/(V{start_row + 10}+V{start_row + 11}-V{start_row + 8}),0),0)",
            f"W{start_row + 13}": f"=IFERROR(IFERROR(W{start_row + 10}/(W{start_row + 10}+W{start_row + 11}-W{start_row + 8}),0),0)",
            f"I{start_row + 9}": "FEB",
            f"J{start_row + 9}": "MAR",
            f"K{start_row + 9}": "APR",
            f"L{start_row + 9}": "MAY",
            f"M{start_row + 9}": "JUN",
            f"N{start_row + 9}": "JUL",
            f"O{start_row + 9}": "AUG",
            f"P{start_row + 9}": "SEP",
            f"Q{start_row + 9}": "OCT",
            f"R{start_row + 9}": "NOV",
            f"S{start_row + 9}": "DEC",
            f"T{start_row + 9}": "JAN",
            f"U{start_row + 9}": "ANNUAL",
            f"V{start_row + 9}": "SPRING",
            f"W{start_row + 9}": "FALL",

            f"I{start_row + 17}": "FEB",
            f"J{start_row + 17}": "MAR",
            f"K{start_row + 17}": "APR",
            f"L{start_row + 17}": "MAY",
            f"M{start_row + 17}": "JUN",
            f"N{start_row + 17}": "JUL",
            f"O{start_row + 17}": "AUG",
            f"P{start_row + 17}": "SEP",
            f"Q{start_row + 17}": "OCT",
            f"R{start_row + 17}": "NOV",
            f"S{start_row + 17}": "DEC",
            f"T{start_row + 17}": "JAN",
            f"U{start_row + 17}": "ANNUAL",
            f"V{start_row + 17}": "SPRING",
            f"W{start_row + 17}": "FALL",


            # f"F{start_row + 2}": f"=IFERROR(C{start_row + 2}/E{start_row + 2},0)",
            f"F{start_row + 7}": f"=C{start_row + 14}+E{start_row + 8}-E{start_row + 7}",
            f"D{start_row + 10}": f"=C{start_row + 10}*C{start_row + 19}",
            f"I{start_row + 10}": TY_Unit_Sales['FEB'],
            f"J{start_row + 10}":TY_Unit_Sales['MAR'],
            f"K{start_row + 10}": TY_Unit_Sales['APR'],
            f"L{start_row + 10}":TY_Unit_Sales['MAY'],
            f"M{start_row + 10}":TY_Unit_Sales['JUN'],
            f"N{start_row + 10}": TY_Unit_Sales['JUL'],
            f"O{start_row + 10}":TY_Unit_Sales['AUG'],
            f"P{start_row + 10}": TY_Unit_Sales['SEP'],
            f"Q{start_row + 10}": TY_Unit_Sales['OCT'],
            f"R{start_row + 10}":TY_Unit_Sales['NOV'],
            f"S{start_row + 10}": TY_Unit_Sales['DEC'],
            f"T{start_row + 10}": TY_Unit_Sales['JAN'],
            f"U{start_row + 10}": f"=IFERROR(SUM(I{start_row + 10}:T{start_row + 10}),0)",
            f"V{start_row + 10}": f"=IFERROR(SUM(I{start_row + 10}:N{start_row + 10}),0)",
            f"W{start_row + 10}": f"=IFERROR(SUM(O{start_row + 10}:T{start_row + 10}),0)" ,
            f"I{start_row + 18}": LY_Unit_Sales['FEB'],
            f"J{start_row + 18}":LY_Unit_Sales['MAR'],
            f"K{start_row + 18}":LY_Unit_Sales['APR'],
            f"L{start_row + 18}":LY_Unit_Sales['MAY'],
            f"M{start_row + 18}":LY_Unit_Sales['JUN'],
            f"N{start_row + 18}":LY_Unit_Sales['JUL'],
            f"O{start_row + 18}":LY_Unit_Sales['AUG'],
            f"P{start_row + 18}":LY_Unit_Sales['SEP'],
            f"Q{start_row + 18}":LY_Unit_Sales['OCT'],
            f"R{start_row + 18}":LY_Unit_Sales['NOV'],
            f"S{start_row + 18}":LY_Unit_Sales['DEC'],
            f"T{start_row + 18}":LY_Unit_Sales['JAN'],
            f"U{start_row + 18}": f"=SUM(I{start_row + 18}:T{start_row + 18})",
            f"V{start_row + 18}": f"=SUM(I{start_row + 18}:N{start_row + 18})",
            f"W{start_row + 18}": f"=SUM(O{start_row + 18}:T{start_row + 18})",
            f"I{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(I{start_row + 10}+I{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,I$4<7),ROUND(I{start_row + 18}+I{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,I$4<7),ROUND(I{start_row + 10}+I{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,I$4>6),ROUND(I{start_row + 18}+I{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,I$4<7),ROUND(I{start_row + 18}+I{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,I$4>6),ROUND(I{start_row + 18}+I{start_row + 18}*$F{start_row + 3},0)))))))",

            f"J{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(J{start_row + 10}+J{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,J$4<7),ROUND(J{start_row + 18}+J{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,J$4<7),ROUND(J{start_row + 10}+J{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,J$4>6),ROUND(J{start_row + 18}+J{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,J$4<7),ROUND(J{start_row + 18}+J{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,J$4>6),ROUND(J{start_row + 18}+J{start_row + 18}*$F{start_row + 3},0)))))))",

            f"L{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(L{start_row + 10}+L{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,L$4<7),ROUND(L{start_row + 18}+L{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,L$4<7),ROUND(L{start_row + 10}+L{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,L$4>6),ROUND(L{start_row + 18}+L{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,L$4<7),ROUND(L{start_row + 18}+L{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,L$4>6),ROUND(L{start_row + 18}+L{start_row + 18}*$F{start_row + 3},0)))))))",

            f"M{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(M{start_row + 10}+M{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,M$4<7),ROUND(M{start_row + 18}+M{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,M$4<7),ROUND(M{start_row + 10}+M{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,M$4>6),ROUND(M{start_row + 18}+M{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,M$4<7),ROUND(M{start_row + 18}+M{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,M$4>6),ROUND(M{start_row + 18}+M{start_row + 18}*$F{start_row + 3},0)))))))",

            f"N{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(N{start_row + 10}+N{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,N$4<7),ROUND(N{start_row + 18}+N{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,N$4<7),ROUND(N{start_row + 10}+N{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,N$4>6),ROUND(N{start_row + 18}+N{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,N$4<7),ROUND(N{start_row + 18}+N{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,N$4>6),ROUND(N{start_row + 18}+N{start_row + 18}*$F{start_row + 3},0)))))))",

            f"O{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(O{start_row + 10}+O{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,O$4<7),ROUND(O{start_row + 18}+O{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,O$4<7),ROUND(O{start_row + 10}+O{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,O$4>6),ROUND(O{start_row + 18}+O{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,O$4<7),ROUND(O{start_row + 18}+O{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,O$4>6),ROUND(O{start_row + 18}+O{start_row + 18}*$F{start_row + 3},0)))))))",

            f"P{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(P{start_row + 10}+P{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,P$4<7),ROUND(P{start_row + 18}+P{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,P$4<7),ROUND(P{start_row + 10}+P{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,P$4>6),ROUND(P{start_row + 18}+P{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,P$4<7),ROUND(P{start_row + 18}+P{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,P$4>6),ROUND(P{start_row + 18}+P{start_row + 18}*$F{start_row + 3},0)))))))",

            f"Q{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(Q{start_row + 10}+Q{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,Q$4<7),ROUND(Q{start_row + 18}+Q{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,Q$4<7),ROUND(Q{start_row + 10}+Q{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,Q$4>6),ROUND(Q{start_row + 18}+Q{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,Q$4<7),ROUND(Q{start_row + 18}+Q{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,Q$4>6),ROUND(Q{start_row + 18}+Q{start_row + 18}*$F{start_row + 3},0)))))))",

            f"R{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(R{start_row + 10}+R{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,R$4<7),ROUND(R{start_row + 18}+R{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,R$4<7),ROUND(R{start_row + 10}+R{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,R$4>6),ROUND(R{start_row + 18}+R{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,R$4<7),ROUND(R{start_row + 18}+R{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,R$4>6),ROUND(R{start_row + 18}+R{start_row + 18}*$F{start_row + 3},0)))))))",

            f"S{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(S{start_row + 10}+S{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,S$4<7),ROUND(S{start_row + 18}+S{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,S$4<7),ROUND(S{start_row + 10}+S{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,S$4>6),ROUND(S{start_row + 18}+S{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,S$4<7),ROUND(S{start_row + 18}+S{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,S$4>6),ROUND(S{start_row + 18}+S{start_row + 18}*$F{start_row + 3},0)))))))",

            f"T{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(T{start_row + 10}+T{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,T$4<7),ROUND(T{start_row + 18}+T{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,T$4<7),ROUND(T{start_row + 10}+T{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,T$4>6),ROUND(T{start_row + 18}+T{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,T$4<7),ROUND(T{start_row + 18}+T{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,T$4>6),ROUND(T{start_row + 18}+T{start_row + 18}*$F{start_row + 3},0)))))))",
            f"K{start_row + 3}": f"=IF(AND($S$1=12,$K$1=12),ROUND(K{start_row + 10}+K{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1=12,K$4<7),ROUND(K{start_row + 18}+K{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,K$4<7),ROUND(K{start_row + 10}+K{start_row + 10}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,K$4>6),ROUND(K{start_row + 18}+K{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1<7,K$4<7),ROUND(K{start_row + 18}+K{start_row + 18}*$F{start_row + 3},0),"
                                f"IF(AND($S$1>6,K$4>6),ROUND(K{start_row + 18}+K{start_row + 18}*$F{start_row + 3},0)))))))",

            # Continue similarly for columns L through T
            f"U{start_row + 3}": f"=SUM(I{start_row + 3}:T{start_row + 3})",
            f"V{start_row + 3}": f"=IFERROR(SUM(I{start_row + 3}:N{start_row + 3}),0)",
            f"W{start_row + 3}": f"=IFERROR(SUM(O{start_row + 3}:T{start_row + 3}),0)",
            
            f"I{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},I{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},I{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",I{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",I{start_row + 18},"
                                f"ROUND(AVERAGE(I{start_row + 2}:I{start_row + 3}),0)))))",

            f"J{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},J{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},J{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",J{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",J{start_row + 18},"
                                f"ROUND(AVERAGE(J{start_row + 2}:J{start_row + 3}),0)))))",

            f"K{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},K{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},K{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",K{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",K{start_row + 18},"
                                f"ROUND(AVERAGE(K{start_row + 2}:K{start_row + 3}),0)))))",

            f"L{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},L{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},L{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",L{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",L{start_row + 18},"
                                f"ROUND(AVERAGE(L{start_row + 2}:L{start_row + 3}),0)))))",

            f"M{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},M{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},M{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",M{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",M{start_row + 18},"
                                f"ROUND(AVERAGE(M{start_row + 2}:M{start_row + 3}),0)))))",

            f"N{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},N{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},N{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",N{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",N{start_row + 18},"
                                f"ROUND(AVERAGE(N{start_row + 2}:N{start_row + 3}),0)))))",

            f"O{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},O{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},O{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",O{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",O{start_row + 18},"
                                f"ROUND(AVERAGE(O{start_row + 2}:O{start_row + 3}),0)))))",

            f"P{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},P{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},P{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",P{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",P{start_row + 18},"
                                f"ROUND(AVERAGE(P{start_row + 2}:P{start_row + 3}),0)))))",

            f"Q{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},Q{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},Q{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",Q{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",Q{start_row + 18},"
                                f"ROUND(AVERAGE(Q{start_row + 2}:Q{start_row + 3}),0)))))",

            f"R{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},R{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},R{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",R{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",R{start_row + 18},"
                                f"ROUND(AVERAGE(R{start_row + 2}:R{start_row + 3}),0)))))",

            f"S{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},S{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},S{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",S{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",S{start_row + 18},"
                                f"ROUND(AVERAGE(S{start_row + 2}:S{start_row + 3}),0)))))",

            f"T{start_row + 4}": f"=IF($F{start_row + 4}=$H{start_row + 2},T{start_row + 2},"
                                f"IF($F{start_row + 4}=$H{start_row + 3},T{start_row + 3},"
                                f"IF($F{start_row + 4}=\"Current year\",T{start_row + 10},"
                                f"IF($F{start_row + 4}=\"Last year\",T{start_row + 18},"
                                f"ROUND(AVERAGE(T{start_row + 2}:T{start_row + 3}),0)))))",

            f"U{start_row + 4}": f"=SUM(I{start_row + 4}:T{start_row + 4})",
            f"V{start_row + 4}": f"=IFERROR(SUM(I{start_row + 4}:N{start_row + 4}),0)",
            f"W{start_row + 4}": f"=IFERROR(SUM(O{start_row + 4}:T{start_row + 4}),0)",
            f"I{start_row + 12}": TY_sales_amt['FEB'],
            f"J{start_row + 12}": TY_sales_amt['MAR'],
            f"K{start_row + 12}": TY_sales_amt['APR'],
            f"L{start_row + 12}":TY_sales_amt['MAY'],
            f"M{start_row + 12}":TY_sales_amt['JUN'],
            f"N{start_row + 12}": TY_sales_amt['JUL'],
            f"O{start_row + 12}": TY_sales_amt['AUG'],
            f"P{start_row + 12}":TY_sales_amt['SEP'],
            f"Q{start_row + 12}":TY_sales_amt['OCT'],
            f"R{start_row + 12}": TY_sales_amt['NOV'],
            f"S{start_row + 12}":TY_sales_amt['DEC'],
            f"T{start_row + 12}": TY_sales_amt['JAN'],


            f"I{start_row + 22}": LY_sales_amt['FEB'],
            f"J{start_row + 22}": LY_sales_amt['MAR'],
            f"K{start_row + 22}": LY_sales_amt['APR'],
            f"L{start_row + 22}":LY_sales_amt['MAY'],
            f"M{start_row + 22}":LY_sales_amt['JUN'],
            f"N{start_row + 22}": LY_sales_amt['JUL'],
            f"O{start_row + 22}": LY_sales_amt['AUG'],
            f"P{start_row + 22}":LY_sales_amt['SEP'],
            f"Q{start_row + 22}":LY_sales_amt['OCT'],
            f"R{start_row + 22}": LY_sales_amt['NOV'],
            f"S{start_row + 22}":LY_sales_amt['DEC'],
            f"T{start_row + 22}": LY_sales_amt['JAN'],

            f"I{start_row + 19}": LY_OH_Units['FEB'],
            f"J{start_row + 19}": LY_OH_Units['MAR'],
            f"K{start_row + 19}": LY_OH_Units['APR'],
            f"L{start_row + 19}":LY_OH_Units['MAY'],
            f"M{start_row + 19}":LY_OH_Units['JUN'],
            f"N{start_row + 19}": LY_OH_Units['JUL'],
            f"O{start_row + 19}": LY_OH_Units['AUG'],
            f"P{start_row + 19}":LY_OH_Units['SEP'],
            f"Q{start_row + 19}":LY_OH_Units['OCT'],
            f"R{start_row + 19}": LY_OH_Units['NOV'],
            f"S{start_row + 19}":LY_OH_Units['DEC'],
            f"T{start_row + 19}": LY_OH_Units['JAN'],
            f"U{start_row + 19}": f"=AVERAGEIFS(I{start_row + 19}:T{start_row + 19}, $I$4:$T$4, \"<=\"&$T$4)",
            f"V{start_row + 19}": f"=AVERAGEIFS(I{start_row + 19}:N{start_row + 19}, $I$4:$N$4, \"<=\"&$N$4)",
            f"W{start_row + 19}": f"=AVERAGEIFS(O{start_row + 19}:T{start_row + 19}, $O$4:$T$4, \"<=\"&$T$4)",
            f"I{start_row + 5}": f"=IF(AND($V$1=\"YTD\",I$4<$K$1),I{start_row + 10},IF(AND($V$1=\"YTD\",I$4=$K$1),I{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",I$4=$K$1),I{start_row + 10},IF(AND($V$1=\"SPRING\",I$4<7),I{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",I$4>6),I{start_row + 10},IF(AND($V$1=\"LY FALL\",I$4>6),I{start_row + 18},I{start_row + 4}))))))",

            f"J{start_row + 5}": f"=IF(AND($V$1=\"YTD\",J$4<$K$1),J{start_row + 10},IF(AND($V$1=\"YTD\",J$4=$K$1),J{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",J$4=$K$1),J{start_row + 10},IF(AND($V$1=\"SPRING\",J$4<7),J{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",J$4>6),J{start_row + 10},IF(AND($V$1=\"LY FALL\",J$4>6),J{start_row + 18},J{start_row + 4}))))))",

            f"K{start_row + 5}": f"=IF(AND($V$1=\"YTD\",K$4<$K$1),K{start_row + 10},IF(AND($V$1=\"YTD\",K$4=$K$1),K{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",K$4=$K$1),K{start_row + 10},IF(AND($V$1=\"SPRING\",K$4<7),K{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",K$4>6),K{start_row + 10},IF(AND($V$1=\"LY FALL\",K$4>6),K{start_row + 18},K{start_row + 4}))))))",

            f"L{start_row + 5}": f"=IF(AND($V$1=\"YTD\",L$4<$K$1),L{start_row + 10},IF(AND($V$1=\"YTD\",L$4=$K$1),L{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",L$4=$K$1),L{start_row + 10},IF(AND($V$1=\"SPRING\",L$4<7),L{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",L$4>6),L{start_row + 10},IF(AND($V$1=\"LY FALL\",L$4>6),L{start_row + 18},L{start_row + 4}))))))",

            f"M{start_row + 5}": f"=IF(AND($V$1=\"YTD\",M$4<$K$1),M{start_row + 10},IF(AND($V$1=\"YTD\",M$4=$K$1),M{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",M$4=$K$1),M{start_row + 10},IF(AND($V$1=\"SPRING\",M$4<7),M{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",M$4>6),M{start_row + 10},IF(AND($V$1=\"LY FALL\",M$4>6),M{start_row + 18},M{start_row + 4}))))))",

            f"N{start_row + 5}": f"=IF(AND($V$1=\"YTD\",N$4<$K$1),N{start_row + 10},IF(AND($V$1=\"YTD\",N$4=$K$1),N{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",N$4=$K$1),N{start_row + 10},IF(AND($V$1=\"SPRING\",N$4<7),N{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",N$4>6),N{start_row + 10},IF(AND($V$1=\"LY FALL\",N$4>6),N{start_row + 18},N{start_row + 4}))))))",

            f"O{start_row + 5}": f"=IF(AND($V$1=\"YTD\",O$4<$K$1),O{start_row + 10},IF(AND($V$1=\"YTD\",O$4=$K$1),O{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",O$4=$K$1),O{start_row + 10},IF(AND($V$1=\"SPRING\",O$4<7),O{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",O$4>6),O{start_row + 10},IF(AND($V$1=\"LY FALL\",O$4>6),O{start_row + 18},O{start_row + 4}))))))",
            f"P{start_row + 5}": f"=IF(AND($V$1=\"YTD\",P$4<$K$1),P{start_row + 10},IF(AND($V$1=\"YTD\",P$4=$K$1),P{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",P$4=$K$1),P{start_row + 10},IF(AND($V$1=\"SPRING\",P$4<7),P{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",P$4>6),P{start_row + 10},IF(AND($V$1=\"LY FALL\",P$4>6),P{start_row + 18},P{start_row + 4}))))))",

            f"Q{start_row + 5}": f"=IF(AND($V$1=\"YTD\",Q$4<$K$1),Q{start_row + 10},IF(AND($V$1=\"YTD\",Q$4=$K$1),Q{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",Q$4=$K$1),Q{start_row + 10},IF(AND($V$1=\"SPRING\",Q$4<7),Q{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",Q$4>6),Q{start_row + 10},IF(AND($V$1=\"LY FALL\",Q$4>6),Q{start_row + 18},Q{start_row + 4}))))))",
            
            f"R{start_row + 5}": f"=IF(AND($V$1=\"YTD\",R$4<$K$1),R{start_row + 10},IF(AND($V$1=\"YTD\",R$4=$K$1),R{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",R$4=$K$1),R{start_row + 10},IF(AND($V$1=\"SPRING\",R$4<7),R{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",R$4>6),R{start_row + 10},IF(AND($V$1=\"LY FALL\",R$4>6),R{start_row + 18},R{start_row + 4}))))))",
            
            f"S{start_row + 5}": f"=IF(AND($V$1=\"YTD\",S$4<$K$1),S{start_row + 10},IF(AND($V$1=\"YTD\",S$4=$K$1),S{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",S$4=$K$1),S{start_row + 10},IF(AND($V$1=\"SPRING\",S$4<7),S{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",S$4>6),S{start_row + 10},IF(AND($V$1=\"LY FALL\",S$4>6),S{start_row + 18},S{start_row + 4}))))))",
            
            f"T{start_row + 5}": f"=IF(AND($V$1=\"YTD\",T$4<$K$1),T{start_row + 10},IF(AND($V$1=\"YTD\",T$4=$K$1),T{start_row + 10},"
                                f"IF(AND($V$1=\"CURRENT MTH\",T$4=$K$1),T{start_row + 10},IF(AND($V$1=\"SPRING\",T$4<7),T{start_row + 10},"
                                f"IF(AND($V$1=\"FALL\",T$4>6),T{start_row + 10},IF(AND($V$1=\"LY FALL\",T$4>6),T{start_row + 18},T{start_row + 4}))))))",
            
            f"U{start_row + 5}": f"=SUM(I{start_row + 5}:T{start_row + 5})",
            
            f"V{start_row + 5}": f"=IFERROR(SUM(I{start_row + 5}:N{start_row + 5}),0)",
            
            f"W{start_row + 5}": f"=IFERROR(SUM(O{start_row + 5}:T{start_row + 5}),0)",

            f"U{start_row + 6}": f"=SUM(I{start_row + 6}:T{start_row + 6})",
            f"V{start_row + 6}": f"=IFERROR(SUM(I{start_row + 6}:N{start_row + 6}),0)",
            f"W{start_row + 6}": f"=IFERROR(SUM(O{start_row + 6}:T{start_row + 6}),0)",
            f"I{start_row + 11}": TY_OH_Units['FEB'],
            f"J{start_row + 11}": TY_OH_Units['MAR'],
            f"K{start_row + 11}":TY_OH_Units['APR'],
            f"L{start_row + 11}": TY_OH_Units['MAY'],
            f"M{start_row + 11}": TY_OH_Units['JUN'],
            f"N{start_row + 11}":TY_OH_Units['JUL'],
            f"O{start_row + 11}": TY_OH_Units['AUG'],
            f"P{start_row + 11}":TY_OH_Units['SEP'],
            f"Q{start_row + 11}":TY_OH_Units['OCT'],
            f"R{start_row + 11}": TY_OH_Units['NOV'],
            f"S{start_row + 11}": TY_OH_Units['DEC'],
            f"T{start_row + 11}": TY_OH_Units['JAN'],

            f"U{start_row + 11}": f"=AVERAGEIFS(I{start_row + 11}:T{start_row + 11}, $I$4:$T$4, \"<=\"&$K$1)",
            
            f"V{start_row + 11}": f"=AVERAGEIFS(I{start_row + 11}:N{start_row + 11}, $I$4:$N$4, \"<=\"&$K$2)",
            
            f"W{start_row + 11}": f"=AVERAGEIFS(O{start_row + 11}:T{start_row + 11}, $O$4:$T$4, \"<=\"&$M$2)",

            f"I{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=I$4,I$4=1),$T{start_row + 19}+I{start_row + 6}+I{start_row + 32}-I{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=I$4),I{start_row + 11}+I{start_row + 6}-(I{start_row + 5}-I{start_row + 10}),IF(AND($V$1=\"Current Mth\",I$4=1,$K$1>1),$T{start_row + 7}+I{start_row + 6}-I{start_row + 5},IF(AND($V$1=\"YTD\",I$4<$K$1),I{start_row + 11},IF(AND($V$1=\"Spring\",I$4<7),I{start_row + 11},IF(AND($V$1=\"Fall\",I$4>6,I$4<$K$1),I{start_row + 11},IF(AND($V$1=\"Fall\",I$4=1,$K$1>1),$T{start_row + 7}+I{start_row + 6}-I{start_row + 5},IF(AND($V$1=\"Fall\",I$4>6,I$4=$K$1),H{start_row + 11}+I{start_row + 32}+I{start_row + 6}-I{start_row + 5},IF(AND($V$1=\"LY Fall\",I$4>6),I{start_row + 19},IF(AND(I$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+I{start_row + 6}+I{start_row + 6}-I{start_row + 5},IF(AND(I$4=$K$1,I$4>1),H{start_row + 11}+I{start_row + 32}+I{start_row + 6}-I{start_row + 5},H{start_row + 7}+I{start_row + 6}-I{start_row + 5})))))))))))",

            f"J{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=J$4,J$4=1),$T{start_row + 19}+J{start_row + 6}+J{start_row + 32}-J{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=J$4),J{start_row + 11}+J{start_row + 6}-(J{start_row + 5}-J{start_row + 10}),IF(AND($V$1=\"Current Mth\",J$4=1,$K$1>1),$T{start_row + 7}+J{start_row + 6}-J{start_row + 5},IF(AND($V$1=\"YTD\",J$4<$K$1),J{start_row + 11},IF(AND($V$1=\"Spring\",J$4<7),J{start_row + 11},IF(AND($V$1=\"Fall\",J$4>6,J$4<$K$1),J{start_row + 11},IF(AND($V$1=\"Fall\",J$4=1,$K$1>1),$T{start_row + 7}+J{start_row + 6}-J{start_row + 5},IF(AND($V$1=\"Fall\",J$4>6,J$4=$K$1),I{start_row + 11}+J{start_row + 32}+J{start_row + 6}-J{start_row + 5},IF(AND($V$1=\"LY Fall\",J$4>6),J{start_row + 19},IF(AND(J$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+J{start_row + 6}+J{start_row + 6}-J{start_row + 5},IF(AND(J$4=$K$1,J$4>1),I{start_row + 11}+J{start_row + 32}+J{start_row + 6}-J{start_row + 5},I{start_row + 7}+J{start_row + 6}-J{start_row + 5})))))))))))",

            f"K{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=K$4,K$4=1),$T{start_row + 19}+K{start_row + 6}+K{start_row + 32}-K{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=K$4),K{start_row + 11}+K{start_row + 6}-(K{start_row + 5}-K{start_row + 10}),IF(AND($V$1=\"Current Mth\",K$4=1,$K$1>1),$T{start_row + 7}+K{start_row + 6}-K{start_row + 5},IF(AND($V$1=\"YTD\",K$4<$K$1),K{start_row + 11},IF(AND($V$1=\"Spring\",K$4<7),K{start_row + 11},IF(AND($V$1=\"Fall\",K$4>6,K$4<$K$1),K{start_row + 11},IF(AND($V$1=\"Fall\",K$4=1,$K$1>1),$T{start_row + 7}+K{start_row + 6}-K{start_row + 5},IF(AND($V$1=\"Fall\",K$4>6,K$4=$K$1),J{start_row + 11}+K{start_row + 32}+K{start_row + 6}-K{start_row + 5},IF(AND($V$1=\"LY Fall\",K$4>6),K{start_row + 19},IF(AND(K$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+K{start_row + 6}+K{start_row + 6}-K{start_row + 5},IF(AND(K$4=$K$1,K$4>1),J{start_row + 11}+K{start_row + 32}+K{start_row + 6}-K{start_row + 5},J{start_row + 7}+K{start_row + 6}-K{start_row + 5})))))))))))",

            f"L{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=L$4,L$4=1),$T{start_row + 19}+L{start_row + 6}+L{start_row + 32}-L{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=L$4),L{start_row + 11}+L{start_row + 6}-(L{start_row + 5}-L{start_row + 10}),IF(AND($V$1=\"Current Mth\",L$4=1,$K$1>1),$T{start_row + 7}+L{start_row + 6}-L{start_row + 5},IF(AND($V$1=\"YTD\",L$4<$K$1),L{start_row + 11},IF(AND($V$1=\"Spring\",L$4<7),L{start_row + 11},IF(AND($V$1=\"Fall\",L$4>6,L$4<$K$1),L{start_row + 11},IF(AND($V$1=\"Fall\",L$4=1,$K$1>1),$T{start_row + 7}+L{start_row + 6}-L{start_row + 5},IF(AND($V$1=\"Fall\",L$4>6,L$4=$K$1),K{start_row + 11}+L{start_row + 32}+L{start_row + 6}-L{start_row + 5},IF(AND($V$1=\"LY Fall\",L$4>6),L{start_row + 19},IF(AND(L$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+L{start_row + 6}+L{start_row + 6}-L{start_row + 5},IF(AND(L$4=$K$1,L$4>1),K{start_row + 11}+L{start_row + 32}+L{start_row + 6}-L{start_row + 5},K{start_row + 7}+L{start_row + 6}-L{start_row + 5})))))))))))",

            f"M{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=M$4,M$4=1),$T{start_row + 19}+M{start_row + 6}+M{start_row + 32}-M{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=M$4),M{start_row + 11}+M{start_row + 6}-(M{start_row + 5}-M{start_row + 10}),IF(AND($V$1=\"Current Mth\",M$4=1,$K$1>1),$T{start_row + 7}+M{start_row + 6}-M{start_row + 5},IF(AND($V$1=\"YTD\",M$4<$K$1),M{start_row + 11},IF(AND($V$1=\"Spring\",M$4<7),M{start_row + 11},IF(AND($V$1=\"Fall\",M$4>6,M$4<$K$1),M{start_row + 11},IF(AND($V$1=\"Fall\",M$4=1,$K$1>1),$T{start_row + 7}+M{start_row + 6}-M{start_row + 5},IF(AND($V$1=\"Fall\",M$4>6,M$4=$K$1),L{start_row + 11}+M{start_row + 32}+M{start_row + 6}-M{start_row + 5},IF(AND($V$1=\"LY Fall\",M$4>6),M{start_row + 19},IF(AND(M$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+M{start_row + 6}+M{start_row + 6}-M{start_row + 5},IF(AND(M$4=$K$1,M$4>1),L{start_row + 11}+M{start_row + 32}+M{start_row + 6}-M{start_row + 5},L{start_row + 7}+M{start_row + 6}-M{start_row + 5})))))))))))",

            f"O{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=O$4,O$4=1),$T{start_row + 19}+O{start_row + 6}+O{start_row + 32}-O{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=O$4),O{start_row + 11}+O{start_row + 6}-(O{start_row + 5}-O{start_row + 10}),IF(AND($V$1=\"Current Mth\",O$4=1,$K$1>1),$T{start_row + 7}+O{start_row + 6}-O{start_row + 5},IF(AND($V$1=\"YTD\",O$4<$K$1),O{start_row + 11},IF(AND($V$1=\"Spring\",O$4<7),O{start_row + 11},IF(AND($V$1=\"Fall\",O$4>6,O$4<$K$1),O{start_row + 11},IF(AND($V$1=\"Fall\",O$4=1,$K$1>1),$T{start_row + 7}+O{start_row + 6}-O{start_row + 5},IF(AND($V$1=\"Fall\",O$4>6,O$4=$K$1),N{start_row + 11}+O{start_row + 32}+O{start_row + 6}-O{start_row + 5},IF(AND($V$1=\"LY Fall\",O$4>6),O{start_row + 19},IF(AND(O$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+O{start_row + 6}+O{start_row + 6}-O{start_row + 5},IF(AND(O$4=$K$1,O$4>1),N{start_row + 11}+O{start_row + 32}+O{start_row + 6}-O{start_row + 5},N{start_row + 7}+O{start_row + 6}-O{start_row + 5})))))))))))",

            f"P{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=P$4,P$4=1),$T{start_row + 19}+P{start_row + 6}+P{start_row + 32}-P{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=P$4),P{start_row + 11}+P{start_row + 6}-(P{start_row + 5}-P{start_row + 10}),IF(AND($V$1=\"Current Mth\",P$4=1,$K$1>1),$T{start_row + 7}+P{start_row + 6}-P{start_row + 5},IF(AND($V$1=\"YTD\",P$4<$K$1),P{start_row + 11},IF(AND($V$1=\"Spring\",P$4<7),P{start_row + 11},IF(AND($V$1=\"Fall\",P$4>6,P$4<$K$1),P{start_row + 11},IF(AND($V$1=\"Fall\",P$4=1,$K$1>1),$T{start_row + 7}+P{start_row + 6}-P{start_row + 5},IF(AND($V$1=\"Fall\",P$4>6,P$4=$K$1),O{start_row + 11}+P{start_row + 32}+P{start_row + 6}-P{start_row + 5},IF(AND($V$1=\"LY Fall\",P$4>6),P{start_row + 19},IF(AND(P$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+P{start_row + 6}+P{start_row + 6}-P{start_row + 5},IF(AND(P$4=$K$1,P$4>1),O{start_row + 11}+P{start_row + 32}+P{start_row + 6}-P{start_row + 5},O{start_row + 7}+P{start_row + 6}-P{start_row + 5})))))))))))",

            f"Q{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=Q$4,Q$4=1),$T{start_row + 19}+Q{start_row + 6}+Q{start_row + 32}-Q{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=Q$4),Q{start_row + 11}+Q{start_row + 6}-(Q{start_row + 5}-Q{start_row + 10}),IF(AND($V$1=\"Current Mth\",Q$4=1,$K$1>1),$T{start_row + 7}+Q{start_row + 6}-Q{start_row + 5},IF(AND($V$1=\"YTD\",Q$4<$K$1),Q{start_row + 11},IF(AND($V$1=\"Spring\",Q$4<7),Q{start_row + 11},IF(AND($V$1=\"Fall\",Q$4>6,Q$4<$K$1),Q{start_row + 11},IF(AND($V$1=\"Fall\",Q$4=1,$K$1>1),$T{start_row + 7}+Q{start_row + 6}-Q{start_row + 5},IF(AND($V$1=\"Fall\",Q$4>6,Q$4=$K$1),P{start_row + 11}+Q{start_row + 32}+Q{start_row + 6}-Q{start_row + 5},IF(AND($V$1=\"LY Fall\",Q$4>6),Q{start_row + 19},IF(AND(Q$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+Q{start_row + 6}+Q{start_row + 6}-Q{start_row + 5},IF(AND(Q$4=$K$1,Q$4>1),P{start_row + 11}+Q{start_row + 32}+Q{start_row + 6}-Q{start_row + 5},P{start_row + 7}+Q{start_row + 6}-Q{start_row + 5})))))))))))",

            f"N{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=N$4,N$4=1),$T{start_row + 19}+N{start_row + 6}+N{start_row + 32}-N{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=N$4),N{start_row + 11}+N{start_row + 6}-(N{start_row + 5}-N{start_row + 10}),IF(AND($V$1=\"Current Mth\",N$4=1,$K$1>1),$T{start_row + 7}+N{start_row + 6}-N{start_row + 5},IF(AND($V$1=\"YTD\",N$4<$K$1),N{start_row + 11},IF(AND($V$1=\"Spring\",N$4<7),N{start_row + 11},IF(AND($V$1=\"Fall\",N$4>6,N$4<$K$1),N{start_row + 11},IF(AND($V$1=\"Fall\",N$4=1,$K$1>1),$T{start_row + 7}+N{start_row + 6}-N{start_row + 5},IF(AND($V$1=\"Fall\",N$4>6,N$4=$K$1),M{start_row + 11}+N{start_row + 32}+N{start_row + 6}-N{start_row + 5},IF(AND($V$1=\"LY Fall\",N$4>6),N{start_row + 19},IF(AND(N$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+N{start_row + 6}+N{start_row + 6}-N{start_row + 5},IF(AND(N$4=$K$1,N$4>1),M{start_row + 11}+N{start_row + 32}+N{start_row + 6}-N{start_row + 5},M{start_row + 7}+N{start_row + 6}-N{start_row + 5})))))))))))",

            f"R{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=R$4,R$4=1),$T{start_row + 19}+R{start_row + 6}+R{start_row + 32}-R{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=R$4),R{start_row + 11}+R{start_row + 6}-(R{start_row + 5}-R{start_row + 10}),IF(AND($V$1=\"Current Mth\",R$4=1,$K$1>1),$T{start_row + 7}+R{start_row + 6}-R{start_row + 5},IF(AND($V$1=\"YTD\",R$4<$K$1),R{start_row + 11},IF(AND($V$1=\"Spring\",R$4<7),R{start_row + 11},IF(AND($V$1=\"Fall\",R$4>6,R$4<$K$1),R{start_row + 11},IF(AND($V$1=\"Fall\",R$4=1,$K$1>1),$T{start_row + 7}+R{start_row + 6}-R{start_row + 5},IF(AND($V$1=\"Fall\",R$4>6,R$4=$K$1),Q{start_row + 11}+R{start_row + 32}+R{start_row + 6}-R{start_row + 5},IF(AND($V$1=\"LY Fall\",R$4>6),R{start_row + 19},IF(AND(R$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+R{start_row + 6}+R{start_row + 6}-R{start_row + 5},IF(AND(R$4=$K$1,R$4>1),Q{start_row + 11}+R{start_row + 32}+R{start_row + 6}-R{start_row + 5},Q{start_row + 7}+R{start_row + 6}-R{start_row + 5})))))))))))",

            f"S{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=S$4,S$4=1),$T{start_row + 19}+S{start_row + 6}+S{start_row + 32}-S{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=S$4),S{start_row + 11}+S{start_row + 6}-(S{start_row + 5}-S{start_row + 10}),IF(AND($V$1=\"Current Mth\",S$4=1,$K$1>1),$T{start_row + 7}+S{start_row + 6}-S{start_row + 5},IF(AND($V$1=\"YTD\",S$4<$K$1),S{start_row + 11},IF(AND($V$1=\"Spring\",S$4<7),S{start_row + 11},IF(AND($V$1=\"Fall\",S$4>6,S$4<$K$1),S{start_row + 11},IF(AND($V$1=\"Fall\",S$4=1,$K$1>1),$T{start_row + 7}+S{start_row + 6}-S{start_row + 5},IF(AND($V$1=\"Fall\",S$4>6,S$4=$K$1),R{start_row + 11}+S{start_row + 32}+S{start_row + 6}-S{start_row + 5},IF(AND($V$1=\"LY Fall\",S$4>6),S{start_row + 19},IF(AND(S$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+S{start_row + 6}+S{start_row + 6}-S{start_row + 5},IF(AND(S$4=$K$1,S$4>1),R{start_row + 11}+S{start_row + 32}+S{start_row + 6}-S{start_row + 5},R{start_row + 7}+S{start_row + 6}-S{start_row + 5})))))))))))",

            f"T{start_row + 7}": f"=IF(AND($V$1=\"Current Mth\",$K$1=T$4,T$4=1),$T{start_row + 19}+T{start_row + 6}+T{start_row + 32}-T{start_row + 5},IF(AND($V$1=\"Current Mth\",$K$1=T$4),T{start_row + 11}+T{start_row + 6}-(T{start_row + 5}-T{start_row + 10}),IF(AND($V$1=\"Current Mth\",T$4=1,$K$1>1),$T{start_row + 7}+T{start_row + 6}-T{start_row + 5},IF(AND($V$1=\"YTD\",T$4<$K$1),T{start_row + 11},IF(AND($V$1=\"Spring\",T$4<7),T{start_row + 11},IF(AND($V$1=\"Fall\",T$4>6,T$4<$K$1),T{start_row + 11},IF(AND($V$1=\"Fall\",T$4=1,$K$1>1),$T{start_row + 7}+T{start_row + 6}-T{start_row + 5},IF(AND($V$1=\"Fall\",T$4>6,T$4=$K$1),S{start_row + 11}+T{start_row + 32}+T{start_row + 6}-T{start_row + 5},IF(AND($V$1=\"LY Fall\",T$4>6),T{start_row + 19},IF(AND(T$4=1,$V$1=\"LY FALL\"),$T{start_row + 19}+T{start_row + 6}+T{start_row + 6}-T{start_row + 5},IF(AND(T$4=$K$1,T$4>1),S{start_row + 11}+T{start_row + 32}+T{start_row + 6}-T{start_row + 5},S{start_row + 7}+T{start_row + 6}-T{start_row + 5})))))))))))",

            f"U{start_row + 7}": f"=AVERAGEIFS(I{start_row + 7}:T{start_row + 7},$I$4:$T$4,\"<=\"&$K$1)",
            f"V{start_row + 7}": f"=AVERAGEIFS(I{start_row + 7}:N{start_row + 7},$I$4:$N$4,\"<=\"&$K$2)",
            f"W{start_row + 7}": f"=AVERAGEIFS(O{start_row + 7}:T{start_row + 7},$O$4:$T$4,\"<=\"&$M$2)",

            f"I{start_row + 6}": planned_shp["FEB"],
            f"J{start_row + 6}": planned_shp["MAR"],
            f"K{start_row + 6}": planned_shp["APR"],
            f"L{start_row + 6}": planned_shp["MAY"],
            f"M{start_row + 6}": planned_shp["JUN"],
            f"N{start_row + 6}": planned_shp["JUL"],
            f"O{start_row + 6}": planned_shp["AUG"],
            f"P{start_row + 6}":planned_shp["SEP"],
            f"Q{start_row + 6}":planned_shp["OCT"],
            f"R{start_row + 6}":planned_shp["NOV"],
            f"S{start_row + 6}": planned_shp["DEC"],
            f"T{start_row + 6}": planned_shp["JAN"],
            f"U{start_row + 8}": f"=SUM(I{start_row + 8}:T{start_row + 8})",
            f"V{start_row + 8}": f"=IFERROR(SUM(I{start_row + 8}:N{start_row + 8}),0)",
            f"W{start_row + 8}": f"=IFERROR(SUM(O{start_row + 8}:T{start_row + 8}),0)",


            f"U{start_row + 9}": f"=SUM(I{start_row + 9}:T{start_row + 9})",
            f"V{start_row + 9}": f"=IFERROR(SUM(I{start_row + 9}:N{start_row + 9}),0)",
            f"W{start_row + 9}": f"=IFERROR(SUM(O{start_row + 9}:T{start_row + 9}),0)",
            f"I{start_row + 8}": f"=I{start_row + 5}/(I{start_row + 5}+I{start_row + 7})",
            f"J{start_row + 8}": f"=J{start_row + 5}/(J{start_row + 5}+J{start_row + 7})",
            f"K{start_row + 8}": f"=K{start_row + 5}/(K{start_row + 5}+K{start_row + 7})",
            f"L{start_row + 8}": f"=L{start_row + 5}/(L{start_row + 5}+L{start_row + 7})",
            f"M{start_row + 8}": f"=M{start_row + 5}/(M{start_row + 5}+M{start_row + 7})",
            f"N{start_row + 8}": f"=N{start_row + 5}/(N{start_row + 5}+N{start_row + 7})",
            f"O{start_row + 8}": f"=O{start_row + 5}/(O{start_row + 5}+O{start_row + 7})",
            f"P{start_row + 8}": f"=P{start_row + 5}/(P{start_row + 5}+P{start_row + 7})",
            f"Q{start_row + 8}": f"=Q{start_row + 5}/(Q{start_row + 5}+Q{start_row + 7})",
            f"R{start_row + 8}": f"=R{start_row + 5}/(R{start_row + 5}+R{start_row + 7})",
            f"S{start_row + 8}": f"=S{start_row + 5}/(S{start_row + 5}+S{start_row + 7})",
            f"T{start_row + 8}": f"=T{start_row + 5}/(T{start_row + 5}+T{start_row + 7})",
            f"U{start_row + 8}": f"=U{start_row + 5}/(U{start_row + 5}+U{start_row + 7})",
            f"V{start_row + 8}": f"=V{start_row + 5}/(V{start_row + 5}+V{start_row + 7})",
            f"W{start_row + 8}": f"=W{start_row + 5}/(W{start_row + 5}+W{start_row + 7})",


            f"U{start_row + 14}": f"=IFERROR(SUM(I{start_row + 14}:T{start_row + 14}),0)",
            f"V{start_row + 14}": f"=IFERROR(SUM(I{start_row + 14}:N{start_row + 14}),0)",
            f"W{start_row + 14}": f"=IFERROR(SUM(O{start_row + 14}:T{start_row + 14}),0)",


            f"I{start_row + 15}": f"=IFERROR(I{start_row + 14}/I{start_row + 10},0)",
            f"J{start_row + 15}": f"=IFERROR(J{start_row + 14}/J{start_row + 10},0)",
            f"K{start_row + 15}": f"=IFERROR(K{start_row + 14}/K{start_row + 10},0)",
            f"L{start_row + 15}": f"=IFERROR(L{start_row + 14}/L{start_row + 10},0)",
            f"M{start_row + 15}": f"=IFERROR(M{start_row + 14}/M{start_row + 10},0)",
            f"N{start_row + 15}": f"=IFERROR(N{start_row + 14}/N{start_row + 10},0)",
            f"O{start_row + 15}": f"=IFERROR(O{start_row + 14}/O{start_row + 10},0)",
            f"P{start_row + 15}": f"=IFERROR(P{start_row + 14}/P{start_row + 10},0)",
            f"Q{start_row + 15}": f"=IFERROR(Q{start_row + 14}/Q{start_row + 10},0)",
            f"R{start_row + 15}": f"=IFERROR(R{start_row + 14}/R{start_row + 10},0)",
            f"S{start_row + 15}": f"=IFERROR(S{start_row + 14}/S{start_row + 10},0)",
            f"T{start_row + 15}": f"=IFERROR(T{start_row + 14}/T{start_row + 10},0)",
            f"U{start_row + 15}": f"=IFERROR(U{start_row + 14}/U{start_row + 10},0)",
            f"V{start_row + 15}": f"=IFERROR(V{start_row + 14}/V{start_row + 10},0)",
            f"W{start_row + 15}": f"=IFERROR(W{start_row + 14}/W{start_row + 10},0)",


            f"U{start_row + 18}": f"=AVERAGEIFS(I{start_row + 18}:T{start_row + 18},$I$4:$T$4,\"<=\"&$K$1)",
            f"V{start_row + 18}": f"=AVERAGEIFS(I{start_row + 18}:N{start_row + 18},$I$4:$N$4,\"<=\"&$K$2)",
            f"W{start_row + 18}": f"=AVERAGEIFS(O{start_row + 18}:T{start_row + 18},$O$4:$T$4,\"<=\"&$M$2)",


            f"I{start_row + 19}": f"=IFERROR(I{start_row + 18}/I{start_row + 11},0)",
            f"J{start_row + 19}": f"=IFERROR(J{start_row + 18}/J{start_row + 11},0)",
            f"K{start_row + 19}": f"=IFERROR(K{start_row + 18}/K{start_row + 11},0)",
            f"L{start_row + 19}": f"=IFERROR(L{start_row + 18}/L{start_row + 11},0)",
            f"M{start_row + 19}": f"=IFERROR(M{start_row + 18}/M{start_row + 11},0)",
            f"N{start_row + 19}": f"=IFERROR(N{start_row + 18}/N{start_row + 11},0)",
            f"O{start_row + 19}": f"=IFERROR(O{start_row + 18}/O{start_row + 11},0)",
            f"P{start_row + 19}": f"=IFERROR(P{start_row + 18}/P{start_row + 11},0)",
            f"Q{start_row + 19}": f"=IFERROR(Q{start_row + 18}/Q{start_row + 11},0)",
            f"R{start_row + 19}": f"=IFERROR(R{start_row + 18}/R{start_row + 11},0)",
            f"S{start_row + 19}": f"=IFERROR(S{start_row + 18}/S{start_row + 11},0)",
            f"T{start_row + 19}": f"=IFERROR(T{start_row + 18}/T{start_row + 11},0)",
            f"U{start_row + 19}": f"=IFERROR(U{start_row + 18}/U{start_row + 11},0)",
            f"V{start_row + 19}": f"=IFERROR(V{start_row + 18}/V{start_row + 11},0)",
            f"W{start_row + 19}": f"=IFERROR(W{start_row + 18}/W{start_row + 11},0)",


            # Aggregates
            f"U{start_row + 20}": f"=IFERROR(SUM(I{start_row + 20}:T{start_row + 20}),0)",
            f"V{start_row + 20}": f"=IFERROR(SUM(I{start_row + 20}:N{start_row + 20}),0)",
            f"W{start_row + 20}": f"=IFERROR(SUM(O{start_row + 20}:T{start_row + 20}),0)",


            f"U{start_row + 21}": f"=IFERROR(SUM(I{start_row + 21}:T{start_row + 21}),0)",
            f"V{start_row + 21}": f"=IFERROR(SUM(I{start_row + 21}:N{start_row + 21}),0)",
            f"W{start_row + 21}": f"=IFERROR(SUM(O{start_row + 21}:T{start_row + 21}),0)",


            f"G{start_row + 3}": f"=IFERROR(U{start_row + 18}+U{start_row + 18}*F{start_row + 3},0)",  # Dynamic references from row 5 onward
            f"C{start_row + 7}": f"=IF(G1=\"SP\",V{start_row + 9},W{start_row + 9})",
            f"D{start_row + 7}": f"=IF($G$1=\"SP\",C{start_row + 7}-V{start_row + 32},C{start_row + 7}-W{start_row + 32})",
            f"E{start_row + 7}": f"=IF($G$1=\"SP\",C{start_row + 7}-V{start_row + 32},C{start_row + 7}-W{start_row + 32})",
            f"E{start_row + 9}": f"=SUMIFS(I{start_row + 5}:T{start_row + 5},$I$4:$T$4,\">=\"&$F$4,$I$4:$T$4,\"<=\"&$G$4)/SUMIFS($I$3:$T$3,$I$4:$T$4,\">=\"&$F$4,$I$4:$T$4,\"<=\"&$G$4)",  # Static references in rows 1 to 4
            f"C{start_row + 10}": f"=IFERROR(C{start_row + 15}/E{start_row + 9},0)",
            f"D{start_row + 10}": f"=IFERROR(C{start_row + 14}/E{start_row + 9},0)",
            f"C{start_row + 11}": f"=IFERROR((D{start_row + 10}-D{start_row + 9})*E{start_row + 9},0)",  # Dynamic D and E references from row 5 onward
            f"D{start_row + 11}": f"=C{start_row + 11}*C{start_row + 19}",  # C24 kept static because it was a specific reference
            f"F{start_row + 11}": f"=IFERROR(IF(U{start_row + 6}>C{start_row + 14},(U{start_row + 6}-C{start_row + 14}+C{start_row + 14}-F{start_row + 10}-C{start_row + 10})/E{start_row + 9},(C{start_row + 14}-F{start_row + 10}-C{start_row + 10})/E{start_row + 9}),0)",
            f"G{start_row + 11}": f"=F{start_row + 11}*E{start_row + 9}",
            f"G{start_row + 19}": f"=((U{start_row + 20}/U{start_row + 10})-C{start_row + 19})/(U{start_row + 20}/U{start_row + 10})",
            f"E{start_row + 33}": f"=ROUND((U{start_row + 6}-U{start_row + 8}-E{start_row + 8})/5, 0) * 5",
            f"A{start_row + 1}": f"=$C{start_row}&H{start_row + 1}",
            f"A{start_row + 2}": f"=$C{start_row}&H{start_row + 2}",
            f"A{start_row + 3}": f"=$C{start_row}&H{start_row + 3}",
            f"A{start_row + 4}": f"=$C{start_row}&H{start_row + 4}",
            f"A{start_row + 5}": f"=$C{start_row}&H{start_row + 5}",
            f"A{start_row + 6}": f"=$C{start_row}&H{start_row + 6}",
            f"A{start_row + 7}": f"=$C{start_row}&H{start_row + 7}",
            f"A{start_row + 8}": f"=$C{start_row}&H{start_row + 8}",
            f"A{start_row + 9}": f"=$C{start_row}&H{start_row + 9}",
            f"A{start_row + 10}": f"=$C{start_row}&H{start_row +10}",
            f"A{start_row + 11}": f"=$C{start_row + 11}&\"Excess Proj\"",
            f"A{start_row + 10}": f"=$C{start_row + 10}&\"Qty to Release\"",
            f"A{start_row + 13}": f"=$C{start_row}&H{start_row +12}",
            f"A{start_row + 14}": f"=$C{start_row}&H{start_row +13}",
            f"A{start_row + 15}": f"=$C{start_row}&H{start_row +14}",
            f"A{start_row + 11}": f"=$C{start_row}&H{start_row +16}",
            f"A{start_row + 17}": f"=$C{start_row}&H{start_row +17}",
            f"A{start_row + 18}": f"=$C{start_row}&H{start_row +18}",

            f"A{start_row + 20}": f"=$C{start_row}&H{start_row +20}",
            f"A{start_row + 21}": f"=$C{start_row}&H{start_row +21}",
            f"A{start_row + 22}": f"=$C{start_row}&H{start_row +22}",
            f"I{start_row + 5}":planned_fc["FEB"],
            f"J{start_row + 5}":planned_fc["MAR"],
            f"K{start_row + 5}":planned_fc["APR"],
            f"L{start_row + 5}":planned_fc["MAY"],
            f"M{start_row + 5}":planned_fc["JUN"],
            f"N{start_row + 5}":planned_fc["JUL"],
            f"O{start_row + 5}":planned_fc["AUG"],
            f"P{start_row + 5}":planned_fc["SEP"],
            f"Q{start_row + 5}":planned_fc["OCT"],
            f"R{start_row + 5}":planned_fc["NOV"],
            f"S{start_row + 5}":planned_fc["DEC"],
            f"T{start_row + 5}":planned_fc["JAN"],
            f"I{start_row + 14}": f"=IFERROR(IFERROR(I{start_row + 10}/I{start_row + 11},0),0)",
            f"J{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:J{start_row + 10})/AVERAGE(I{start_row + 11}:J{start_row + 11}),0),0)",
            f"K{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:K{start_row + 10})/AVERAGE(I{start_row + 11}:K{start_row + 11}),0),0)",
            f"L{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:L{start_row + 10})/AVERAGE(I{start_row + 11}:L{start_row + 11}),0),0)",
            f"M{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:M{start_row + 10})/AVERAGE(I{start_row + 11}:M{start_row + 11}),0),0)",
            f"N{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:N{start_row + 10})/AVERAGE(I{start_row + 11}:N{start_row + 11}),0),0)",
            f"O{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:O{start_row + 10})/AVERAGE(I{start_row + 11}:O{start_row + 11}),0),0)",
            f"P{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:P{start_row + 10})/AVERAGE(I{start_row + 11}:P{start_row + 11}),0),0)",
            f"Q{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:Q{start_row + 10})/AVERAGE(I{start_row + 11}:Q{start_row + 11}),0),0)",
            f"R{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:R{start_row + 10})/AVERAGE(I{start_row + 11}:R{start_row + 11}),0),0)",
            f"S{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:S{start_row + 10})/AVERAGE(I{start_row + 11}:S{start_row + 11}),0),0)",
            f"T{start_row + 14}": f"=IFERROR(IFERROR(SUM(I{start_row + 10}:T{start_row + 10})/AVERAGE(I{start_row + 11}:T{start_row + 11}),0),0)",
            f"U{start_row + 14}": f"=IFERROR(U{start_row + 10}/U{start_row + 11},0)",
            f"V{start_row + 14}": f"=IFERROR(V{start_row + 10}/V{start_row + 11},0)",
            f"W{start_row + 14}": f"=IFERROR(W{start_row + 10}/W{start_row + 11},0)",

            f"I{start_row + 21}": f"=IFERROR(IFERROR(I{start_row + 18}/I{start_row + 19},0),0)",
            f"J{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:J{start_row + 18})/AVERAGE(I{start_row + 19}:J{start_row + 19}),0),0)",
            f"K{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:K{start_row + 18})/AVERAGE(I{start_row + 19}:K{start_row + 19}),0),0)",
            f"L{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:L{start_row + 18})/AVERAGE(I{start_row + 19}:L{start_row + 19}),0),0)",
            f"M{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:M{start_row + 18})/AVERAGE(I{start_row + 19}:M{start_row + 19}),0),0)",
            f"N{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:N{start_row + 18})/AVERAGE(I{start_row + 19}:N{start_row + 19}),0),0)",
            f"O{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:O{start_row + 18})/AVERAGE(I{start_row + 19}:O{start_row + 19}),0),0)",
            f"P{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:P{start_row + 18})/AVERAGE(I{start_row + 19}:P{start_row + 19}),0),0)",
            f"Q{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:Q{start_row + 18})/AVERAGE(I{start_row + 19}:Q{start_row + 19}),0),0)",
            f"R{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:R{start_row + 18})/AVERAGE(I{start_row + 19}:R{start_row + 19}),0),0)",
            f"S{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:S{start_row + 18})/AVERAGE(I{start_row + 19}:S{start_row + 19}),0),0)",
            f"T{start_row + 21}": f"=IFERROR(IFERROR(SUM(I{start_row + 18}:T{start_row + 18})/AVERAGE(I{start_row + 19}:T{start_row + 19}),0),0)",
            f"U{start_row + 21}": f"=IFERROR(U{start_row + 18}/U{start_row + 19},0)",
            f"V{start_row + 21}": f"=IFERROR(V{start_row + 18}/V{start_row + 19},0)",
            f"W{start_row + 21}": f"=IFERROR(W{start_row + 18}/W{start_row + 19},0)",

            f"J{start_row + 15}": f"=IFERROR(IF(J{start_row + 10}+J{start_row + 18}=0,0,(J{start_row + 10}-J{start_row + 18})/J{start_row + 18}),1)",
            f"I{start_row + 15}": f"=IFERROR(IF(I{start_row + 10}+I{start_row + 18}=0,0,(I{start_row + 10}-I{start_row + 18})/I{start_row + 18}),1)",
            f"K{start_row + 15}": f"=IFERROR(IF(K{start_row + 10}+K{start_row + 18}=0,0,(K{start_row + 10}-K{start_row + 18})/K{start_row + 18}),1)",
            f"L{start_row + 15}": f"=IFERROR(IF(L{start_row + 10}+L{start_row + 18}=0,0,(L{start_row + 10}-L{start_row + 18})/L{start_row + 18}),1)",
            f"M{start_row + 15}": f"=IFERROR(IF(M{start_row + 10}+M{start_row + 18}=0,0,(M{start_row + 10}-M{start_row + 18})/M{start_row + 18}),1)",
            f"N{start_row + 15}": f"=IFERROR(IF(N{start_row + 10}+N{start_row + 18}=0,0,(N{start_row + 10}-N{start_row + 18})/N{start_row + 18}),1)",
            f"O{start_row + 15}": f"=IFERROR(IF(O{start_row + 10}+O{start_row + 18}=0,0,(O{start_row + 10}-O{start_row + 18})/O{start_row + 18}),1)",
            f"P{start_row + 15}": f"=IFERROR(IF(P{start_row + 10}+P{start_row + 18}=0,0,(P{start_row + 10}-P{start_row + 18})/P{start_row + 18}),1)",
            f"Q{start_row + 15}": f"=IFERROR(IF(Q{start_row + 10}+Q{start_row + 18}=0,0,(Q{start_row + 10}-Q{start_row + 18})/Q{start_row + 18}),1)",
            f"R{start_row + 15}": f"=IFERROR(IF(R{start_row + 10}+R{start_row + 18}=0,0,(R{start_row + 10}-R{start_row + 18})/R{start_row + 18}),1)",
            f"S{start_row + 15}": f"=IFERROR(IF(S{start_row + 10}+S{start_row + 18}=0,0,(S{start_row + 10}-S{start_row + 18})/S{start_row + 18}),1)",
            f"T{start_row + 15}": f"=IFERROR(IF(T{start_row + 10}+T{start_row + 18}=0,0,(T{start_row + 10}-T{start_row + 18})/T{start_row + 18}),1)",
            f"U{start_row + 15}": f"=IFERROR(IF(U{start_row + 10}+U{start_row + 18}=0,0,(U{start_row + 10}-U{start_row + 18})/U{start_row + 18}),1)",
            f"V{start_row + 15}": f"=IFERROR(IF(V{start_row + 10}+V{start_row + 18}=0,0,(V{start_row + 10}-V{start_row + 18})/V{start_row + 18}),1)",
            f"W{start_row + 15}": f"=IFERROR(IF(W{start_row + 10}+W{start_row + 18}=0,0,(W{start_row + 10}-W{start_row + 18})/W{start_row + 18}),1)",
            f"I{start_row + 16}": f"=IFERROR(IF(I{start_row + 11}+I{start_row + 19}=0,0,(I{start_row + 11}-I{start_row + 19})/I{start_row + 19}),1)",
            f"J{start_row + 16}": f"=IFERROR(IF(J{start_row + 11}+J{start_row + 19}=0,0,(J{start_row + 11}-J{start_row + 19})/J{start_row + 19}),1)",
            f"K{start_row + 16}": f"=IFERROR(IF(K{start_row + 11}+K{start_row + 19}=0,0,(K{start_row + 11}-K{start_row + 19})/K{start_row + 19}),1)",
            f"L{start_row + 16}": f"=IFERROR(IF(L{start_row + 11}+L{start_row + 19}=0,0,(L{start_row + 11}-L{start_row + 19})/L{start_row + 19}),1)",
            f"M{start_row + 16}": f"=IFERROR(IF(M{start_row + 11}+M{start_row + 19}=0,0,(M{start_row + 11}-M{start_row + 19})/M{start_row + 19}),1)",
            f"N{start_row + 16}": f"=IFERROR(IF(N{start_row + 11}+N{start_row + 19}=0,0,(N{start_row + 11}-N{start_row + 19})/N{start_row + 19}),1)",
            f"O{start_row + 16}": f"=IFERROR(IF(O{start_row + 11}+O{start_row + 19}=0,0,(O{start_row + 11}-O{start_row + 19})/O{start_row + 19}),1)",
            f"P{start_row + 16}": f"=IFERROR(IF(P{start_row + 11}+P{start_row + 19}=0,0,(P{start_row + 11}-P{start_row + 19})/P{start_row + 19}),1)",
            f"Q{start_row + 16}": f"=IFERROR(IF(Q{start_row + 11}+Q{start_row + 19}=0,0,(Q{start_row + 11}-Q{start_row + 19})/Q{start_row + 19}),1)",
            f"R{start_row + 16}": f"=IFERROR(IF(R{start_row + 11}+R{start_row + 19}=0,0,(R{start_row + 11}-R{start_row + 19})/R{start_row + 19}),1)",
            f"S{start_row + 16}": f"=IFERROR(IF(S{start_row + 11}+S{start_row + 19}=0,0,(S{start_row + 11}-S{start_row + 19})/S{start_row + 19}),1)",
            f"T{start_row + 16}": f"=IFERROR(IF(T{start_row + 11}+T{start_row + 19}=0,0,(T{start_row + 11}-T{start_row + 19})/T{start_row + 19}),1)",
            f"U{start_row + 16}": f"=IFERROR(IF(U{start_row + 11}+U{start_row + 19}=0,0,(U{start_row + 11}-U{start_row + 19})/U{start_row + 19}),1)",
            f"V{start_row + 16}": f"=IFERROR(IF(V{start_row + 11}+V{start_row + 19}=0,0,(V{start_row + 11}-V{start_row + 19})/V{start_row + 19}),1)",
            f"W{start_row + 16}": f"=IFERROR(IF(W{start_row + 11}+W{start_row + 19}=0,0,(W{start_row + 11}-W{start_row + 19})/W{start_row + 19}),1)",

                }
                # Add dropdown validation to the specific cell (e.g., F column) for each loop

        # Apply formulas
        percentage_ranges = [
        f"I{start_row + 1}:W{start_row + 1}",

        f"E{start_row + 2}",

        f"F{start_row + 3}",
    ]
        rounded_ranges = [
            f"F{start_row + 2}",
            f"G{start_row + 3}",
            f"E{start_row + 9}",
            f"C{start_row + 10}",
            f"D{start_row + 10}",
            f"C{start_row + 11}",
            f"F{start_row + 11}",
            f"F{start_row + 18}",
            f"U{start_row + 7}:W{start_row + 7}",
            f"U{start_row + 11}:W{start_row + 11}",
            f"U{start_row + 18}:W{start_row + 18}",
            f"I{start_row + 20}:W{start_row + 20}",
            f"I{start_row + 21}:W{start_row + 21}",
            f"U{start_row + 19}:W{start_row + 19}",
            f"U{start_row + 39}:W{start_row + 39}",
            f"U{start_row + 40}:W{start_row + 40}",


        ]
        # rounded_ranges_one = [

        #     f"I{start_row + 25}:T{start_row + 25}",
        #     f"I{start_row + 26}:T{start_row + 26}",

        # ]
        # rounded_ranges_two = [
        #     f"U{start_row + 25}:W{start_row + 25}",
        #     f"U{start_row + 26}:W{start_row + 26}",

        # ]
        currency_ranges = [
        f"D{start_row + 11}",
        f"C{start_row + 19}",
        f"D{start_row + 19}",
        f"C{start_row + 20}",
        f"D{start_row + 10}",
        f"E{start_row + 19}",
        f"I{start_row + 20}:W{start_row + 20}",
        f"I{start_row + 21}:W{start_row + 21}",

        
        # Add more ranges as needed
    ]
        sheet = ws
                # Add values to column B from ALL_VALUES and apply alignment
        for i, value in enumerate(ALL_VALUES, start=start_row):
            cell = ws.cell(row=i, column=2, value=value)  # Column B is the 2nd column
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Add values to column H from H_VALUES and apply alignment
        for i, value in enumerate(H_VALUES, start=start_row):
            cell = ws.cell(row=i, column=8, value=value)  # Column H is the 8th column
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Add monthly values starting from column I (9th column) in the specified row
        for col, value in enumerate(MONTHLY_VALUES, start=9):
            ws.cell(row=start_row, column=col, value=value)
        for cell, formula in dynamic_formulas.items():
            try:
                ws[cell] = formula  # Set the formula directly in the specified cell
            except Exception as e:
                print(f"Error setting formula in {cell}: {e}")
        # Apply ROUND formatting with different decimal places
        apply_round_format(ws, rounded_ranges, decimal_places=0)
        # apply_round_format(ws, rounded_ranges_one, decimal_places=1)
        # apply_round_format(ws, rounded_ranges_two, decimal_places=2)
    # Apply formats
        apply_format(ws, percentage_ranges, "0%")        # Apply percentage format
        apply_format(ws, currency_ranges, "$#,##0") 
        # add_dropdown(ws, f"F{start_row + 1}", category_options)
        add_dropdown(ws, f"F{start_row + 4}", forecast_method_options)




    #######loop formatinf
        red_font= Font(color="FF0000") 
        blue_font=Font(color="0000FF") 
        white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        dark_pink = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        gray = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")
        pink_font = Font(color="9C0006") 
        for row in ws[f"C{start_row}:G{start_row + 8}"]:
            for cell in row:
                cell.fill = gray

        for row in ws[f"C{start_row + 13}:G{start_row + 20}"]:  # Use the dynamically constructed range
            for cell in row:
                cell.fill = gray


        ws[f"D{start_row + 1}"].fill = white  # D6 becomes D(start_row + 1)
        ws[f"E{start_row + 3}"].fill = white  # E8 becomes E(start_row + 3)





        pink_fill_font = [
            f"C{start_row + 3}",  # C8 -> start_row + 3
            f"F{start_row + 3}",  # F8 -> start_row + 3
            f"F{start_row + 14}",  # F19 -> start_row + 14
            f"G{start_row + 14}",  # G19 -> start_row + 14
            f"C{start_row + 11}",  # C21 -> start_row + 16
        ]
        for cell_address in pink_fill_font:
            ws[cell_address].fill = dark_pink
            ws[cell_address].font = pink_font





        # For the range I14:W14
        for row in ws[f"I{start_row + 9}:W{start_row + 9}"]:  # I14:W14 -> start_row + 9
            for cell in row:
                cell.font = red_font  
        yelow_fill_font_list=['E13','F13','Q13','R13','S13','G39',]
        for i in yelow_fill_font_list:
            ws[i].fill = light_yellow
            ws[i].font = yellow_font
        #Hyperlink
        # ws[f"B{start_row + 26}"].font = Font(color="0563C1", underline="single", bold=False)  # B31 -> start_row + 26

        # Dynamically merge cells
        # ws.merge_cells(f"D{start_row + 21}:G{start_row + 22}")  # D26:G31
        # ws.merge_cells(f"D{start_row + 11}:E{start_row + 11}")  # D16:E16
        # ws.merge_cells(f"D{start_row + 1}:E{start_row + 1}")    # D6:E6
        # ws.merge_cells(f"D{start_row}:E{start_row}")            # D5:E5
        # ws.merge_cells(f"D{start_row + 20}:F{start_row + 20}")  # D25:F25
        # ws.merge_cells(f"C{start_row + 13}:F{start_row + 13}")  # C18:F18
        # ws.merge_cells(f"D{start_row + 11}:G{start_row + 11}")  # D21:G21
        # Dynamically set alignment
        ws[f"D{start_row + 21}"].alignment = Alignment(horizontal='center', vertical='center')  # D26
        ws[f"D{start_row + 11}"].alignment = Alignment(horizontal='center', vertical='center')  # D16
        ws[f"D{start_row}"].alignment = Alignment(horizontal='center', vertical='center')       # D5
        ws[f"D{start_row + 1}"].alignment = Alignment(horizontal='right', vertical='center')    # D6
        ws[f"C{start_row + 3}"].alignment = Alignment(horizontal='center', vertical='center')   # C8
        ws[f"D{start_row + 11}"].alignment = Alignment(horizontal='left', vertical='center')    # D21
        ws[f"D{start_row + 20}"].alignment = Alignment(horizontal='center', vertical='center')  # D25

        #gradient bg
        two_yellow_gradient_fill = GradientFill(type="linear",degree=90,stop=("FFFFFF", "FFFF99", "FFFFFF")) 
        top_green_gradient_fill = GradientFill(type="linear", degree=90, stop=("E2EFDA", "FFFFFF"))
        full_green_gradient_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        bott_green_gradient_fill = GradientFill(type="linear", degree=90, stop=("FFFFFF","E2EFDA"))
        bott_light_gray_gradient_fill = GradientFill(type="linear", degree=90, stop=( "FFFFFF","EDEDED"))
        top_gray_gradient_fill = GradientFill(type="linear", degree=90, stop=("D9D9D9", "FFFFFF"))
        dark_gray_gradient_fill = GradientFill(type="linear", degree=90, stop=( "FFFFFF","DBDBDB"))
        top_orange_gradient_fill = GradientFill(type="linear", degree=90, stop=("FFE699", "FFFFFF"))
        bott_orange_gradient_fill = GradientFill(type="linear", degree=90, stop=( "FFFFFF","FFE699"))
        top_yellow_gradient_fill = GradientFill(type="linear", degree=90, stop=("FFFFCC", "FFFFFF"))
        full_yellow_gradient_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
        bott_yellow_gradient_fill = GradientFill(type="linear", degree=90, stop=("FFFFFF","FFFFCC"))
        bott_blue_gradient_fill = GradientFill(type="linear", degree=90, stop=("FFFFFF","DFECF7"))
        full_blue_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        top_pink_gradient_fill = GradientFill(type="linear", degree=90, stop=("FCE4D6", "FFFFFF"))
        full_pink_gradient_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        bott_pink_gradient_fill = GradientFill(type="linear", degree=90, stop=("FFFFFF","FCE4D6"))
        # Pink
        for row in ws[f"C{start_row + 9}:G{start_row + 9}"]:  # C14:G14 -> start_row + 9
            for cell in row:
                cell.fill = bott_pink_gradient_fill  

        for row in ws[f"C{start_row + 10}:G{start_row + 11}"]:  # C15:G16 -> start_row + 10 to start_row + 11
            for cell in row:
                cell.fill = full_pink_gradient_fill  

        for row in ws[f"C{start_row + 10}:G{start_row + 10}"]:  # C17:G17 -> start_row + 12
            for cell in row:
                cell.fill = top_pink_gradient_fill 

        # Yellow
        for row in ws[f"I{start_row + 4}:W{start_row + 4}"]:  # I9:W9 -> start_row + 4
            for cell in row:
                cell.fill = two_yellow_gradient_fill  

        # Green
        for row in ws[f"I{start_row + 5}:W{start_row + 5}"]:  # I10:W10 -> start_row + 5
            for cell in row:
                cell.fill = bott_green_gradient_fill  

        for row in ws[f"I{start_row + 6}:W{start_row + 6}"]:  # I11:W11 -> start_row + 6
            for cell in row:
                cell.fill = full_green_gradient_fill  

        for row in ws[f"I{start_row + 7}:W{start_row + 7}"]:  # I12:W12 -> start_row + 7
            for cell in row:
                cell.fill = top_green_gradient_fill  

        # Blue
        bott_blue_fill_list = [
            f"I{start_row + 10}:W{start_row + 10}",  # I17:W17 -> start_row + 12
            f"I{start_row + 20}:W{start_row + 20}",  # I25:W25 -> start_row + 20
            f"I{start_row + 18}:W{start_row + 18}",  # I39:W39 -> start_row + 34
            f"I{start_row + 47}:W{start_row + 47}"   # I52:W52 -> start_row + 47
        ]
        for i in bott_blue_fill_list:
            for row in ws[i]:
                for cell in row:
                    cell.fill = bott_blue_gradient_fill  
                    cell.font = Font(bold=True)  

        full_blue_fill_list = [
            f"I{start_row + 13}:W{start_row + 14}",  # I18:W19 -> start_row + 13 to start_row + 14
            f"I{start_row + 21}:W{start_row + 22}",  # I26:W27 -> start_row + 21 to start_row + 22
            f"I{start_row + 35}:W{start_row + 36}",  # I40:W41 -> start_row + 35 to start_row + 36
            f"I{start_row + 48}:W{start_row + 49}"   # I53:W54 -> start_row + 48 to start_row + 49
        ]
        for i in full_blue_fill_list:
            for row in ws[i]:
                for cell in row:
                    cell.fill = full_blue_fill  
        # Gray
        for row in ws[f"I{start_row + 8}:W{start_row + 8}"]:  # I13:W13 -> start_row + 8
            for cell in row:
                cell.fill = bott_light_gray_gradient_fill 

        for row in ws[f"I{start_row + 9}:W{start_row + 9}"]:  # I14:W14 -> start_row + 9
            for cell in row:
                cell.fill = top_gray_gradient_fill 

        dark_gray_fill_list = [
            f"H{start_row}:W{start_row}",       # H5:W5 -> start_row
            f"H{start_row + 9}:W{start_row + 9}",  # H16:W16 -> start_row + 11
            f"H{start_row + 17}:W{start_row + 17}"   # H38:W38 -> start_row + 33
        ]
        for i in dark_gray_fill_list:
            for row in ws[i]:
                for cell in row:
                    cell.fill = dark_gray_gradient_fill 
                    cell.font=Font(bold=True)

        # Individual cells
        ws[f"F{start_row + 1}"].fill = dark_gray_gradient_fill  # F6 -> start_row + 1
        ws[f"F{start_row + 4}"].fill = dark_gray_gradient_fill  # F9 -> start_row + 4



 

        # Orange
        top_orange_fill_list = [
            f"I{start_row + 15}:W{start_row + 15}",  # I20:W20 -> start_row + 15
            f"I{start_row + 37}:W{start_row + 37}"  # I42:W42 -> start_row + 37
        ]
        bott_orange_fill_list = [
            f"I{start_row + 19}:W{start_row + 19}",  # I24:W24 -> start_row + 19
            f"I{start_row + 41}:W{start_row + 41}"  # I46:W46 -> start_row + 41
        ]

        # Apply top orange gradient fill
        for i in top_orange_fill_list:
            for row in ws[i]:
                for cell in row:
                    cell.fill = top_orange_gradient_fill 

        # Apply bottom orange gradient fill
        for i in bott_orange_fill_list:
            for row in ws[i]:
                for cell in row:
                    cell.fill = bott_orange_gradient_fill
        bold_list=['C5','D5','F5','C6','B26','B39','B44','C38','E38','G39','H20','H21','H24','H32','H33','H34','H46','H43','H42']
        for i in bold_list:
            ws[i].font = Font(bold=True)


        for row in ws[f"C{start_row}:W{start_row + 49}"]:  # Use the dynamically calculated range
            for cell in row:
                cell.border = gridline1  # Apply the border
                # Hide column A
        # Define border style for top and bottom only

        gridline_top_bottom = Border(
            left=Side(style="thin", color='FFFFFF'),
            right=Side(style="thin", color='FFFFFF'),
            top=Side(style="thin", color='D9D9D9'),
            bottom=Side(style="thin", color='D9D9D9')
        )

        for row in ws[f"B{start_row+50}:W{start_row+50}"]:
            for cell in row:
                cell.border = gridline_top_bottom
        # Define border style for left and right only

    # Hide column A - moved outside the product loop
    ws.column_dimensions.group("A", "A", outline_level=1, hidden=True)

    # Save the workbook - outside the product loop but inside the category loop
    wb.save(output_file)
    print(f"Successfully saved {output_file}")
    # writer.close()
    wb.close()


 
end_time = time.time()
runtime = end_time - start_time

# Display runtime in seconds
