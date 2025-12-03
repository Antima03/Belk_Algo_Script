# variable_loader.py

import pandas as pd
import logging

class VariableLoader:

    def __init__(self, all_data, product_id, this_year_value, last_year_value, last_to_last_year_value):
        
        # ---------------------------------
        #  PREPARE MONTH COLUMN
        # ---------------------------------
        if "Month_upper" not in all_data.columns:
            all_data = all_data.copy()
            all_data["Month_upper"] = all_data["Month"].astype(str).str.upper()

        # ---------------------------------
        #  FILTER DATA BY PRODUCT + YEAR
        # ---------------------------------
        self.this_year_data = all_data.loc[
            (all_data["Belk Style #"] == product_id) &
            (all_data["Year"] == this_year_value)
        ]

        self.last_year_data = all_data.loc[
            (all_data["Belk Style #"] == product_id) &
            (all_data["Year"] == last_year_value)
        ]

        self.last_to_last_year_data = all_data.loc[
            (all_data["Belk Style #"] == product_id) &
            (all_data["Year"] == last_to_last_year_value)
        ]

        print(f"\n=== PID: {product_id} ===")
        print("This year rows:", len(self.this_year_data))
        print("Last year rows:", len(self.last_year_data))
        print("Last-to-last year rows:", len(self.last_to_last_year_data))

        # BELK MONTH ORDER
        months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]

        # ---------------------------------
        #  INIT DICTIONARIES 
        # ---------------------------------
        self.ty_lw_sls_u = {m: 0 for m in months}
        self.ly_lw_sls_u = {m: 0 for m in months}
        self.lly_lw_sls_u = {m: 0 for m in months}

        self.ty_lw_eop_oh_u = {m: 0 for m in months}
        self.ly_lw_eop_oh_u = {m: 0 for m in months}
        self.lly_lw_eop_oh_u = {m: 0 for m in months}

        self.ty_lw_rel_act_loc_count = {m: 0 for m in months}
        self.ly_lw_rel_act_loc_count = {m: 0 for m in months}
        self.lly_lw_rel_act_loc_count = {m: 0 for m in months}

        self.ty_latest_rel_act_loc_month = None
        self.ly_latest_rel_act_loc_month = None
        self.lly_latest_rel_act_loc_month = None
        self.ty_latest_rel_act_loc_value = 0
        self.ly_latest_rel_act_loc_value = 0
        self.lly_latest_rel_act_loc_value = 0

        self.ty_lw_sls_dollar = {m: 0 for m in months}
        self.ly_lw_sls_dollar = {m: 0 for m in months}
        self.lly_lw_sls_dollar = {m: 0 for m in months}

        month_order = {m: idx for idx, m in enumerate(months)}

        def latest_month_available(df):
            if df.empty:
                return None
            unique_months = [m for m in df["Month_upper"].unique() if m in month_order]
            if not unique_months:
                return None
            return max(unique_months, key=lambda m: month_order[m])

        def _sum_value(df, month, column):
            value = df.loc[df["Month_upper"] == month, column].sum()
            if pd.isna(value):
                return 0
            return value.item() if hasattr(value, "item") else value

        ty_latest = latest_month_available(self.this_year_data)
        ly_latest = latest_month_available(self.last_year_data)
        lly_latest = latest_month_available(self.last_to_last_year_data)
        self.ty_latest_rel_act_loc_month = ty_latest
        self.ly_latest_rel_act_loc_month = ly_latest
        self.lly_latest_rel_act_loc_month = lly_latest

        # ---------------------------------
        #  FILL VALUES MONTH-WISE
        # ---------------------------------
        months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]

        for month in months:

            # ========================== 
            # LW SALES UNITS 
            # ==========================
            self.ty_lw_sls_u[month] = _sum_value(self.this_year_data, month, "LW Sls U")
            self.ly_lw_sls_u[month] = _sum_value(self.last_year_data, month, "LW Sls U")
            self.lly_lw_sls_u[month] = _sum_value(self.last_to_last_year_data, month, "LW Sls U")

            # ========================== 
            # LW EOP OH UNITS
            # ==========================
            self.ty_lw_eop_oh_u[month] = _sum_value(self.this_year_data, month, "LW EOP OH U")
            self.ly_lw_eop_oh_u[month] = _sum_value(self.last_year_data, month, "LW EOP OH U")
            self.lly_lw_eop_oh_u[month] = _sum_value(self.last_to_last_year_data, month, "LW EOP OH U") 

           # ==========================
            #  LW REL ACT LOC COUNT - all months
            # ==========================
            self.ty_lw_rel_act_loc_count[month]  = _sum_value(self.this_year_data,        month, "LW Repl Act Loc Count")
            self.ly_lw_rel_act_loc_count[month]  = _sum_value(self.last_year_data,        month, "LW Repl Act Loc Count")
            self.lly_lw_rel_act_loc_count[month] = _sum_value(self.last_to_last_year_data, month, "LW Repl Act Loc Count")

            # ==========================
            # LW REL ACT LOC COUNT – latest month only
            # ==========================
            if month == ty_latest:
                self.ty_latest_rel_act_loc_value = self.ty_lw_rel_act_loc_count[month]
            if month == ly_latest:
                self.ly_latest_rel_act_loc_value = self.ly_lw_rel_act_loc_count[month]
            if month == lly_latest:
                self.lly_latest_rel_act_loc_value = self.lly_lw_rel_act_loc_count[month]
            # ========================== 
            # LW SALES DOLLAR
            # ==========================
            self.ty_lw_sls_dollar[month] = _sum_value(self.this_year_data, month, "LW Sls $")
            self.ly_lw_sls_dollar[month] = _sum_value(self.last_year_data, month, "LW Sls $")
            self.lly_lw_sls_dollar[month] = _sum_value(self.last_to_last_year_data, month, "LW Sls $")

            print(
                f"[{month}] "
                f"TY SLS={self.ty_lw_sls_u[month]}, "
                f"LY SLS={self.ly_lw_sls_u[month]}, "
                f"LLY SLS={self.lly_lw_sls_u[month]}, "
                f"TY EOP={self.ty_lw_eop_oh_u[month]}, "
                f"LY EOP={self.ly_lw_eop_oh_u[month]}, "
                f"LLY EOP={self.lly_lw_eop_oh_u[month]}, "
                f"TY REL LOC={self.ty_latest_rel_act_loc_value if month == ty_latest else 0}, "
                f"LY REL LOC={self.ly_latest_rel_act_loc_value if month == ly_latest else 0}, "
                f"LLY REL LOC={self.lly_latest_rel_act_loc_value if month == lly_latest else 0}"
            )

        logging.info("LW Sls U + LW EOP OH loaded successfully.")
