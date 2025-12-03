import json
import os

# Adjust path if this file is somewhere else
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE_DIR, "belk_category_forecasts.json")

MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]

def is_all_zero_for_three_years(loader_dict):
    """
    loader_dict is something like:
      {
        "ty_lw_sls_u": {...},
        "ly_lw_sls_u": {...},
        "lly_lw_sls_u": {...},
        ...
      }
    Returns True if for every month in MONTHS, all three values are 0.
    """
    ty = loader_dict.get("ty_lw_sls_u", {})
    ly = loader_dict.get("ly_lw_sls_u", {})
    lly = loader_dict.get("lly_lw_sls_u", {})

    for m in MONTHS:
        if (ty.get(m, 0) != 0) or (ly.get(m, 0) != 0) or (lly.get(m, 0) != 0):
            return False
    return True

def find_zero_sales_styles(json_path):
    with open(json_path, "r") as f:
        data = json.load(f)

    zero_styles = []  # list of tuples: (category, belk_style)

    # Structure: {category: {belk_style: {"loader": {...}, "forecast": {...}}}}
    for category, styles in data.items():
        for belk_style, content in styles.items():
            loader = content.get("loader", {})
            if is_all_zero_for_three_years(loader):
                zero_styles.append((category, belk_style))

    return zero_styles

def main():
    if not os.path.exists(JSON_PATH):
        print(f"JSON file not found: {JSON_PATH}")
        return

    zero_styles = find_zero_sales_styles(JSON_PATH)

    print("belk_style with lw_sls_u = 0 for all 12 months in all three years:")
    print("Count:", len(zero_styles))
    print("-" * 80)
    for category, belk_style in zero_styles:
        print(f"Category: {category} | belk_style: {belk_style}")

    # OPTIONAL: also write to a CSV for easier review
    out_csv = os.path.join(BASE_DIR, "belk_styles_zero_sales_all_3yrs.csv")
    try:
        with open(out_csv, "w") as f:
            f.write("category,belk_style\n")
            for category, belk_style in zero_styles:
                f.write(f"{category},{belk_style}\n")
        print(f"\nSaved list to: {out_csv}")
    except Exception as e:
        print(f"Could not write CSV: {e}")

if __name__ == "__main__":
    main()