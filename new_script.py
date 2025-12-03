import polars as pl
import pandas as pd
import time
from getereatilinfo import get_previous_retail_week
from var import VariableLoader
from datetime import datetime
from belk_algo import algorithm
from utils import *
from logger import logging
RAW_FILE = "Belk Weekly Raw Data.xlsx"
VENDOR_FILE = "vendor_sheet.xlsx"
TARGET_YEAR = None


# ---------- Safe Autocast ----------
def safe_autocast(df: pl.DataFrame) -> pl.DataFrame:
    """
    Convert string columns to numbers, but ONLY if the values look numeric.
    """
    new_cols = {}

    for col in df.columns:
        if df[col].dtype == pl.Utf8:

            # sample first 200 values (fast)
            sample = df[col].head(200).drop_nulls().to_list()

            # check if ALL non-null values look numeric
            def looks_number(x):
                x = str(x).replace(",", "").strip()
                return x.replace(".", "", 1).isdigit()

            if len(sample) > 0 and all(looks_number(v) for v in sample):
                # cast to float
                new_cols[col] = pl.col(col).str.replace_all(",", "").cast(pl.Float64)

    # return DF with safe changes applied
    return df.with_columns(new_cols.values()) if new_cols else df


# ---------- Load each sheet using polars (super fast) ----------
def load_all_sheets_fast(path):
    xls = pd.ExcelFile(path)
    sheet_names = xls.sheet_names

    print("📂 FAST loading sheets:")

    dfs = {}
    for name in sheet_names:
        print(f"  ✔ {name}")

        df = pl.read_excel(path, sheet_name=name, infer_schema_length=None)

        # FIX: Only numeric-looking strings are cast
        df = safe_autocast(df)

        dfs[name] = df

    return dfs



# -------- Convert all polars DFs → one combined DF ----------
def combine_sheets(dfs):
    return pl.concat(list(dfs.values()), how="vertical")


# ---------- Compute Last Week ----------
def compute_last_week(df):
    if "Week (Month)" in df.columns:
        week_col = "Week (Month)"
    elif "Week" in df.columns:
        week_col = "Week"
    else:
        return None

    return (
        df.group_by(["Year", "Month"])
          .agg(pl.col(week_col).max().alias("Week"))
    )



# ---------- Aggregate ----------
import polars as pl

def aggregate_data(df: pl.DataFrame, target_year=None) -> pl.DataFrame:
    """
    Aggregation rules:
      - All numeric columns -> SUM
      - Except:
          "LW EOP OH U",
          "LW EOP OH U LY",
          "LW Repl Act Loc Count"
        -> take value from LAST week of that month (per Year, Month, Belk Style #)
      - AUR / % columns -> MEAN
      - Non-numeric -> FIRST
    """

    # Optional year filter
    if target_year is not None:
        df = df.filter(pl.col("Year") == target_year)

    group_keys = ["Year", "Month", "Belk Style #"]

    # Columns that should use LAST WEEK value instead of sum
    LAST_WEEK_COLS = [
        "LW EOP OH U",
        "LW EOP OH U LY",
        "LW Repl Act Loc Count",
    ]

    # ---------- 1) Normal numeric aggregation (excluding last-week cols) ----------
    numeric_cols = [
        c for c, dt in zip(df.columns, df.dtypes)
        if dt in (pl.Float64, pl.Float32, pl.Int64, pl.Int32)
        and c not in group_keys
        and c not in ["Week", "Week (Month)"]
    ]

    agg_numeric_cols = [c for c in numeric_cols if c not in LAST_WEEK_COLS]

    agg_exprs = []

    for col in agg_numeric_cols:
        col_l = col.lower()
        if "aur" in col_l or "%" in col_l:
            agg_exprs.append(pl.col(col).mean().alias(col))
        else:
            agg_exprs.append(pl.col(col).sum().alias(col))

    # ---------- 2) Non-numeric columns -> first() ----------
    non_numeric_cols = [
        c for c in df.columns
        if c not in numeric_cols and c not in group_keys and c not in ["Week", "Week (Month)"]
    ]

    for col in non_numeric_cols:
        agg_exprs.append(pl.col(col).first().alias(col))

    # Base aggregated table (without LAST_WEEK_COLS yet)
    aggregated = (
        df.group_by(group_keys)
          .agg(agg_exprs)
          .sort(group_keys)
    )

    # ---------- 3) Compute last-week rows for LAST_WEEK_COLS ----------
    # Decide which week column to use
    if "Week (Month)" in df.columns:
        week_col = "Week (Month)"
    elif "Week" in df.columns:
        week_col = "Week"
    else:
        # No week column; just return aggregated as-is
        return aggregated

    # For each (Year, Month, Style), keep the row with max week
    # 1) Find last week per group
    week_max = (
        df.group_by(group_keys)
          .agg(pl.col(week_col).max().alias(week_col))
    )

    # 2) Keep only rows that are in that last week
    df_last_week = df.join(
        week_max,
        on=group_keys + [week_col],
        how="inner",
    )

    # 3) Compute a combined score from LAST_WEEK_COLS (blank/null -> 0)
    score_exprs = [
        pl.when(pl.col(c).is_null()).then(0).otherwise(pl.col(c))
        for c in LAST_WEEK_COLS
        if c in df_last_week.columns
    ]

    if score_exprs:
        df_last_week = df_last_week.with_columns(
            (sum(score_exprs)).alias("_last_week_score")
        )
    else:
        # No LAST_WEEK_COLS present: just take any row in last week
        df_last_week = df_last_week.with_columns(
            pl.lit(0).alias("_last_week_score")
        )

    # 4) Within those last-week rows, pick row with max score
    last_week_rows = (
        df_last_week
        .sort(
            group_keys + ["_last_week_score"],
            descending=[False] * len(group_keys) + [True],
        )
        .group_by(group_keys)
        .head(1)
        .select(
            group_keys
            + [c for c in LAST_WEEK_COLS if c in df_last_week.columns]
        )
    )

    # ---------- 4) Join last-week metrics onto aggregated ----------
    aggregated = aggregated.join(
        last_week_rows,
        on=group_keys,
        how="left",
    )

    return aggregated



# ---------- Index table ----------
def create_index_table(aggregated_pl):
    """
    Build Category × Month index table using only POLARS.
    Returns a Pandas df at the end (ready for your return statement).
    """
    # Step 1: Aggregate LW Sls U by Category × Month
    df = (
        aggregated_pl
        .group_by(["Category", "Month"])
        .agg(pl.col("LW Sls U").sum().alias("Units"))
        .sort(["Category", "Month"])
    )

    # Step 2: Pivot POLARS (wide format)
    pivot = df.pivot(
        values="Units",
        index="Category",
        on="Month"
    ).fill_null(0)

    # Step 3: Normalize rows so each row = 100%
    rename_map = {col: col.upper() for col in pivot.columns if col != "Category"}
    if rename_map:
        pivot = pivot.rename(rename_map)

    month_cols = [c for c in pivot.columns if c != "Category"]
    print(f"month_cols: {month_cols}")
   

    pivot = pivot.with_columns(
        [
            (pl.col(col) / pl.sum_horizontal(month_cols) * 100).alias(col)
            for col in month_cols
        ]
    )

    # Step 4: Return as pandas
    return pivot.to_pandas()



# ==================== PIPELINE ====================
def run_belk_pipeline(raw_file, vendor_file):
    start = time.time()

    # 1) Load vendor guideline (for style filter + category mapping)
    vendorGuideline = pl.read_excel(vendor_file)       # POLARS

    # 2) Fast sheet loading
    sheet_dfs = load_all_sheets_fast(raw_file)

    # 3) Combine using polars (instant)
    combined = pl.concat(list(sheet_dfs.values()))  # POLARS

    # ---- REMOVE CLEARANCE ROWS HERE ----
    if "Clearance?" in combined.columns:
        combined = combined.filter(pl.col("Clearance?") != "Clearance")
    # ------------------------------------

    # 4) Filter to Belk styles present in vendor file and attach Category
    #    Assumes vendor has columns: "Belk Style Number" and "Go-Forward Category"
    if "Belk Style #" in combined.columns and "Belk Style Number" in vendorGuideline.columns:
        vendor_styles = (
            vendorGuideline
            .select([
                pl.col("Belk Style Number"),
                pl.col("Go-Forward Category").alias("Category"),
            ])
            .unique(subset=["Belk Style Number"])
        )

        combined = combined.join(
            vendor_styles,
            left_on="Belk Style #",
            right_on="Belk Style Number",
            how="inner",   # keep only styles present in vendor file
        )

    # 5) Aggregate after vendor filter / category join
    aggregated = aggregate_data(combined)           # POLARS

    # If join created a vendor category column like "Category_right",
    # drop the original Category and use the vendor-mapped one.
    if "Category_right" in aggregated.columns:
        if "Category" in aggregated.columns:
            aggregated = aggregated.drop("Category")
        aggregated = aggregated.rename({"Category_right": "Category"})

    last_week_map = compute_last_week(combined)    # POLARS
   
    aggregated = aggregated.join(last_week_map, on=["Year","Month"], how="left")
    index_df = create_index_table(aggregated)          # returns Pandas
    # 6) Merge Last week
    if last_week_map is not None:
        aggregated = aggregated.join(last_week_map, on=["Year", "Month"], how="left")

    # 7) Create Index
    index_df = create_index_table(aggregated)

    aggregated_pd = aggregated.to_pandas()
    index_df_pd = index_df  # already pandas
    vendorGuideline_pd = pd.DataFrame(vendorGuideline,columns=vendorGuideline.columns)

    end = time.time()
    print(f"\n⏱️ TOTAL TIME: {end-start:.2f} sec")

    # Return ONLY required dfs
    return (
        aggregated_pd,
        index_df_pd,
        vendorGuideline_pd,
    )


def process_belk_data(raw_file, vendor_file,std_start_month,std_end_month):
    aggregated, index_df, vendorGuideline = run_belk_pipeline(raw_file,vendor_file)
    print(f"vendorGuideline: {vendorGuideline.head()}.")
    (current_date,current_month,current_month_number,rolling_method, previous_week_number, retail_year,last_retail_year, last_month_of_previous_month_numeric,season, feb_weeks, mar_weeks, apr_weeks, may_weeks,jun_weeks, jul_weeks, aug_weeks, sep_weeks, oct_weeks,nov_weeks, dec_weeks, jan_weeks,wpm_dict) = previous_retail_week_info = get_previous_retail_week(aggregated)
    std_period = generate_std_period(std_start_month,std_end_month)
    # 2. Filter only Year = 2025

    print(f"Previous Retail Week Info: {previous_retail_week_info}")
    
    category_payload = {}
    # 3. Loop category-wise, then Belk Style # wise
    for category, cat_df in aggregated.groupby("Category"):
        print(f"\n=== Category: {category} ===")
        print(cat_df.head())
        category_payload.setdefault(category, {})
        description = cat_df["Description"].iloc[0]
        product = cat_df["Product"].iloc[0]
        clearance = cat_df["Clearance?"].iloc[0]
    
        
        # Unique Belk Style # in this category for 2025
        unique_pids = cat_df["Belk Style #"].unique()
        print("Belk Style # in this category (2025):", unique_pids)
        
        # Inner loop for each Belk Style #
        for pid in unique_pids:
            pid_df = cat_df[cat_df["Belk Style #"] == pid]
            loader = VariableLoader(pid_df,pid,retail_year,last_retail_year,last_retail_year-1)
            print("LW Sls U and LW EOP OH U Dicts for Style 014OUL18:", loader)
            print(f"ty_lw_sls_u : {loader.ty_lw_sls_u}")
            print(f"ly_lw_sls_u : {loader.ly_lw_sls_u}")
            print(f"lly_lw_sls_u : {loader.lly_lw_sls_u}")
            print(f"ty_lw_eop_oh_u : {loader.ty_lw_eop_oh_u}")
            print(f"ly_lw_eop_oh_u : {loader.ly_lw_eop_oh_u}")
            print(f"lly_lw_eop_oh_u : {loader.lly_lw_eop_oh_u}")
            print(f"ty_lw_rel_act_loc_count : {loader.ty_latest_rel_act_loc_value}")
            print(f"ly_lw_rel_act_loc_count : {loader.ly_latest_rel_act_loc_value}")
            print(f"lly_lw_rel_act_loc_count : {loader.lly_latest_rel_act_loc_value}")

            forecast_dict = algorithm(category,pid,std_period,loader,vendorGuideline,previous_retail_week_info,index_df,description,product,clearance)
            print(f"Forecast Dict: {forecast_dict}")

            category_payload[category][pid] = {
                    "loader": {
                        "ty_lw_sls_u": loader.ty_lw_sls_u,
                        "ly_lw_sls_u": loader.ly_lw_sls_u,
                        "lly_lw_sls_u": loader.lly_lw_sls_u,
                        "ty_lw_eop_oh_u": loader.ty_lw_eop_oh_u,
                        "ly_lw_eop_oh_u": loader.ly_lw_eop_oh_u,
                        "lly_lw_eop_oh_u": loader.lly_lw_eop_oh_u,
                        "ty_latest_rel_act_loc_value": loader.ty_latest_rel_act_loc_value,
                        "ly_latest_rel_act_loc_value": loader.ly_latest_rel_act_loc_value,
                        "lly_latest_rel_act_loc_value": loader.lly_latest_rel_act_loc_value,
                        "ty_lw_rel_act_loc_count": loader.ty_lw_rel_act_loc_count,
                        "ly_lw_rel_act_loc_count": loader.ly_lw_rel_act_loc_count,
                        "lly_lw_rel_act_loc_count": loader.lly_lw_rel_act_loc_count,
                        "ty_lw_sls_dollar": loader.ty_lw_sls_dollar,
                        "ly_lw_sls_dollar": loader.ly_lw_sls_dollar,
                        "lly_lw_sls_dollar": loader.lly_lw_sls_dollar,
                    },
                    "forecast": forecast_dict,
                }

        
        

            # print(dsdef)

    output_path = "belk_category_forecasts.json"
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump(category_payload, fh, indent=2)

    print(f"\nSaved category forecast JSON to {output_path}")




            # Here you can do whatever you want with this subset
            # Example: show a small preview
        



# Run test
if __name__ == "__main__":
    # aggregated, index_df, vendorGuideline = run_belk_pipeline(RAW_FILE,VENDOR_FILE)
    # writer = pd.ExcelWriter("Belk_Aggregate.xlsx", engine='openpyxl')
    # aggregated.to_excel(writer, sheet_name="aggregated", index=False)
    # index_df.to_excel(writer, sheet_name="index", index=False)
    # vendorGuideline.to_excel(writer, sheet_name="vendor_guideline", index=False)
    # writer.close()
    # (current_date,current_month,current_month_number,rolling_method, previous_week_number, retail_year,last_retail_year, last_month_of_previous_month_numeric,season, feb_weeks, mar_weeks, apr_weeks, may_weeks,jun_weeks, jul_weeks, aug_weeks, sep_weeks, oct_weeks,nov_weeks, dec_weeks, jan_weeks,wpm_dict) = previous_retail_week_info = get_previous_retail_week(aggregated)

    # # previous_retail_week_info = get_previous_retail_week(aggregated)
    # Var = build_lw_sls_u_dicts_for_style(aggregated, '014OUL18',retail_year,last_retail_year,last_retail_year-1)
    # print("Previous Retail Week Info:", previous_retail_week_info)
    # print("LW Sls U and LW EOP OH U Dicts for Style 014OUL18:", Var)
    process_belk_data(RAW_FILE, VENDOR_FILE,"February","May")


