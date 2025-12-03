import json
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE_DIR, "belk_category_forecasts.json")

MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]

def is_first_time_2025(loader_dict):
    """
    True if:
      - any ty_lw_sls_u[month] > 0
      - and all ly_lw_sls_u[month] == 0
      - and all lly_lw_sls_u[month] == 0
    """
    ty  = loader_dict.get("ty_lw_sls_u",  {})
    ly  = loader_dict.get("ly_lw_sls_u",  {})
    lly = loader_dict.get("lly_lw_sls_u", {})

    # 2025 has some sales
    has_ty_sales = any(ty.get(m, 0) != 0 for m in MONTHS)

    # previous years have no sales at all
    ly_all_zero  = all(ly.get(m, 0) == 0 for m in MONTHS)
    lly_all_zero = all(lly.get(m, 0) == 0 for m in MONTHS)

    return has_ty_sales and ly_all_zero and lly_all_zero

def find_first_time_2025_styles(json_path):
    with open(json_path, "r") as f:
        data = json.load(f)

    styles = []   # (category, belk_style)

    # {category: {belk_style: {"loader": {...}, "forecast": {...}}}}
    for category, styles_dict in data.items():
        for belk_style, content in styles_dict.items():
            loader = content.get("loader", {})
            if is_first_time_2025(loader):
                styles.append((category, belk_style))

    return styles

def main():
    if not os.path.exists(JSON_PATH):
        print(f"JSON file not found: {JSON_PATH}")
        return

    styles = find_first_time_2025_styles(JSON_PATH)

    print("belk_style with sales first time in 2025 (no sales in prior years):")
    print("Count:", len(styles))
    print("-" * 80)
    for category, belk_style in styles:
        print(f"Category: {category} | belk_style: {belk_style}")

    # Optional CSV
    out_csv = os.path.join(BASE_DIR, "belk_styles_first_time_2025.csv")
    try:
        with open(out_csv, "w") as f:
            f.write("category,belk_style\n")
            for category, belk_style in styles:
                f.write(f"{category},{belk_style}\n")
        print(f"\nSaved list to: {out_csv}")
    except Exception as e:
        print(f"Could not write CSV: {e}")

if __name__ == "__main__":
    main()