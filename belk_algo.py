from utils import *
from logger import logging
from datetime import datetime
MONTHS = ['FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC','JAN']

def algorithm(category,pid,std_period,loader,vendorGuideline,previous_retail_week_info,index_df,description,product,clearance):


    (current_date,current_month,current_month_number,rolling_method, previous_week_number, retail_year,last_retail_year, last_month_of_previous_month_numeric,season, feb_weeks, mar_weeks, apr_weeks, may_weeks,jun_weeks, jul_weeks, aug_weeks, sep_weeks, oct_weeks,nov_weeks, dec_weeks, jan_weeks,wpm_dict) = previous_retail_week_info
    rj_style_number,vendor_no,vendor_name,country,lead_time_days,min_order = get_belk_vendor_info(vendorGuideline,pid)
    print(f"rj_style_number: {rj_style_number}")
    print(f"vendor_no: {vendor_no}")
    print(f"vendor_name: {vendor_name}")
    print(f"country: {country}")
    print(f"lead_time_days: {lead_time_days}")
    print(f"min_order: {min_order}")
    # current_date = datetime(2025, 5, 31)
    # current_month = current_date.strftime("%b").upper()
    # current_month_number = current_date.month

    std_ty_unit_sales_list = [loader.ty_lw_sls_u[month] for month in std_period]
    std_ly_unit_sales_list = [loader.ly_lw_sls_u[month] for month in std_period]
    month_num_list = [i+1 for i in range(12)]
    ty_lw_sls_u_list = [loader.ty_lw_sls_u[month] for month in MONTHS]
    ly_lw_sls_u_list = [loader.ly_lw_sls_u[month] for month in MONTHS]
    ly_avg_eoh=sum(std_ly_unit_sales_list)/len(std_ly_unit_sales_list)
    print(f"std_ty_unit_sales_list: {std_ty_unit_sales_list}")
    print(f"std_ly_unit_sales_list: {std_ly_unit_sales_list}")
    print(f"month_num_list: {month_num_list}")
    print(f"ty_lw_sls_u_list: {ty_lw_sls_u_list}")
    print(f"ly_lw_sls_u_list: {ly_lw_sls_u_list}")
    forecast_date = get_forecast_date(lead_time_days, current_date)
    forecast_month = get_forecast_info(forecast_date)
    print(f"forecast_month: {forecast_month}")
    week_of_month = get_week_of_month(forecast_date)
    print(f"week_of_month: {week_of_month}")

    std_trend = calculate_std_trend(std_ty_unit_sales_list, std_ly_unit_sales_list)
    print(f"std_trend: {std_trend}")


    index_value_dict = get_index_dict_of_month(index_df,category)
    print(f"index_value_dict: {index_value_dict}")

    std_index_value = calculate_std_index_value(index_value_dict,std_period)
    print(f"std_index_value: {std_index_value}")


    month_12_fc_index = calculate_12th_month_forecast(std_ty_unit_sales_list, std_index_value)
    print(f"month_12_fc_index: {month_12_fc_index}")
    loss=calculate_loss(loader.ty_latest_rel_act_loc_value, ly_avg_eoh)
    month_12_fc_index =month_12_fc_index*(1+loss)
    fc_by_index = calculate_fc_by_index(index_value_dict, month_12_fc_index)
    print(f"fc_by_index: {fc_by_index}")
    # Calculate FC by Trend
    fc_by_trend = calculate_fc_by_trend(last_month_of_previous_month_numeric,current_month_number,std_trend, month_num_list, ty_lw_sls_u_list, ly_lw_sls_u_list)
    print(f"fc_by_trend: {fc_by_trend}")
    
    # Door Count
    ly_lw_eop_oh_u = [loader.ly_lw_eop_oh_u[month] for month in std_period]
    print(f"ly_lw_eop_oh_u: {ly_lw_eop_oh_u}")
    first = std_period[0]
    idx = MONTHS.index(first)
    prev_month = MONTHS[(idx - 1) % len(MONTHS)]
    std_period_with_previous_month= [prev_month] + std_period
    logging.info(f"std_period_with_previous_month : {std_period_with_previous_month}")
    print("="*50)
    print("this year relp count",loader.ty_latest_rel_act_loc_value)
    print("last year relp count",loader.ly_latest_rel_act_loc_value)
    print("last last year relp count",loader.lly_latest_rel_act_loc_value)
    print("="*50)


    ly_lw_eop_oh_u_for_std = [loader.ly_lw_eop_oh_u[month] for month in std_period_with_previous_month]
    last_12_months_eop_oh = get_last_12_months_inventory(current_month,loader.ty_lw_eop_oh_u, loader.ly_lw_eop_oh_u)
    is_inventory_maintained_store=is_maintained(last_12_months_eop_oh, 0.95, loader.ty_latest_rel_act_loc_value)
    is_inventory_maintained_ly_std_period_store=is_maintained(ly_lw_eop_oh_u, 0.95, loader.ty_latest_rel_act_loc_value)
    store_trend_index_difference,store_seasonal_total_fc_by_trend,store_seasonal_total_fc_by_index=compare_seasonal_forecasts_by_method(fc_by_index,fc_by_trend,season)
    logging.info(f'store_trend_index_difference: {store_trend_index_difference}')

    store_forecasting_method=decide_forecasting_method(is_inventory_maintained_ly_std_period_store,store_trend_index_difference,std_ly_unit_sales_list)
    forecasting_method_original=store_forecasting_method

    recommended_fc=get_recommended_forecast(store_forecasting_method, fc_by_index, fc_by_trend)





    return{"fc_by_index":fc_by_index,
            "fc_by_trend":fc_by_trend,
            "std_index_value":std_index_value,
            "std_trend":std_trend,
            "std_ty_unit_sales_list":std_ty_unit_sales_list,
            "std_ly_unit_sales_list":std_ly_unit_sales_list,
            "month_num_list":month_num_list,
            "ty_lw_sls_u_list":ty_lw_sls_u_list,
            "ly_lw_sls_u_list":ly_lw_sls_u_list,
            "forecast_date":forecast_date.isoformat(),

            "forecast_month":forecast_month,
            "week_of_month":week_of_month,
            "std_period":std_period,
            "category":category,
            "description":description,
            "product":product,
            "clearance":clearance,
            "pid":pid,
            "std_period":std_period,
            "month_12_fc_index":month_12_fc_index,
            "rj_style_number":rj_style_number,
            "vendor_no":vendor_no,
            "vendor_name":vendor_name,
            "country":country,
            "lead_time_days":lead_time_days,
            "min_order":min_order,
            "forecasting_method":forecasting_method_original,
            "index_value":index_value_dict,
            "recommended_fc":recommended_fc}