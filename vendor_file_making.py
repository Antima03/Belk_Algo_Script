import pandas as pd
from pathlib import Path

# ---------- Paths ----------
BASE_DIR = Path(r"c:\Users\phefa\OneDrive\Desktop\Belk_forecasti")

VENDOR_FILE   = BASE_DIR / "Harish Request 11.11.25.xlsx"
CATEGORY_FILE = BASE_DIR / "BELK CATEGORY LIST.xlsx"
RAW_FILE      = BASE_DIR / "Belk Weekly Raw Data.xlsx"

OUTPUT_FILE   = BASE_DIR / "Harish Request 11.11.25_with_category.xlsx"

# ---------- Column Names (change here if your headers differ) ----------
VENDOR_RL_COL   = "RL Style Number"  
Vendor_BELK_COL = "Belk Style Number" # in vendor file
CAT_RL_COL      = "RL STYLE #"   # in BELK CATEGORY LIST
CAT_CATEGORY_COL = "Go-Forward Category"  # in BELK CATEGORY LIST

RAW_RL_COL      = "Belk Style #"   # in Belk Weekly Raw Data
RAW_CATEGORY_COL = "Category"  # in Belk Weekly Raw Data

def main():
    # 1. Read files
    print("Reading vendor file...")
    df_vendor = pd.read_excel(VENDOR_FILE)

    print("Reading category list...")
    df_cat = pd.read_excel(CATEGORY_FILE)

    print("Reading raw weekly file (this may take some time)...")
    df_raw = pd.read_excel(RAW_FILE)

    print("Files read. Columns in each file:")
    print("Vendor columns:", list(df_vendor.columns))
    print("Category list columns:", list(df_cat.columns))
    print("Raw file columns:", list(df_raw.columns))

    # 2. Normalize RL style keys for safer join
        # 2. Normalize keys for safer join
    for df, col in [
        (df_vendor, VENDOR_RL_COL),
        (df_vendor, Vendor_BELK_COL),
        (df_cat, CAT_RL_COL),
        (df_raw, RAW_RL_COL),
    ]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)
            )
        else:
            print(f"WARNING: Column {col!r} not found in DataFrame; columns are {list(df.columns)}")
            
    # 3. PRIMARY: merge vendor with BELK CATEGORY LIST
        # 3. PRIMARY: merge vendor with BELK CATEGORY LIST on RL style
    df_vendor = df_vendor.merge(
        df_cat[[CAT_RL_COL, CAT_CATEGORY_COL]].drop_duplicates(),
        left_on=VENDOR_RL_COL,
        right_on=CAT_RL_COL,
        how="left",
    )

    # After merge, category column from category list is CAT_CATEGORY_COL
    primary_cat_col = CAT_CATEGORY_COL

    # 4. Find rows where category is still missing after category file
    mask_missing = df_vendor[primary_cat_col].isna()

    if mask_missing.any():
        # Use only rows that have a Belk Style Number present
        mask_belk_available = df_vendor[Vendor_BELK_COL].notna()
        mask_fallback = mask_missing & mask_belk_available

        if mask_fallback.any():
            # 5. Build Belk Style -> Category mapping from RAW file
            raw_lookup = (
                df_raw[[RAW_RL_COL, RAW_CATEGORY_COL]]
                .drop_duplicates()
            )

            # Get distinct Belk Style Numbers that need fallback
            missing_belk_styles = (
                df_vendor.loc[mask_fallback, Vendor_BELK_COL]
                .dropna()
                .unique()
            )

            df_missing = pd.DataFrame({Vendor_BELK_COL: missing_belk_styles})
            df_missing = df_missing.merge(
                raw_lookup,
                left_on=Vendor_BELK_COL,
                right_on=RAW_RL_COL,
                how="left",
            )

            # Map Belk Style Number -> Category from raw
            raw_map = dict(
                zip(df_missing[Vendor_BELK_COL], df_missing[RAW_CATEGORY_COL])
            )

            # 6. Fill missing categories from RAW mapping using Belk Style Number
            df_vendor.loc[mask_fallback, primary_cat_col] = (
                df_vendor.loc[mask_fallback, Vendor_BELK_COL].map(raw_map)
            )

    # 7. Save result
    df_vendor.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved new vendor file with category to:\n{OUTPUT_FILE}")

if __name__ == "__main__":
    main()