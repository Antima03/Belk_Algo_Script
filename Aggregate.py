import pandas as pd
import openpyxl
from datetime import datetime

# ---------------- CONFIG ----------------
EXCEL_FILE = "Belk Weekly Raw Data.xlsx"   # your workbook with 3 sheets
TARGET_YEAR = None              # None = ALL years; or set like 2023
SHEETS_TO_USE = None            # None = all sheets; or ["2023","2024","2025"]
OUTPUT_FILE = "belk_aggregated_all_years_data_1.xlsx"
CHUNK_SIZE = 50000              # Process in chunks for better memory management
# ---------------------------------------


def load_all_sheets_chunked(path, sheets=None, chunksize=50000):
    """Read sheets in chunks using openpyxl for memory efficiency."""
    import time
    
    print(f"📂 Loading: {path}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    
    sheet_names = sheets if sheets is not None else wb.sheetnames
    print(f"📊 Processing {len(sheet_names)} sheet(s): {sheet_names}")
    
    all_chunks = []
    
    for sheet_idx, sheet_name in enumerate(sheet_names, 1):
        print(f"  [{sheet_idx}/{len(sheet_names)}] {sheet_name}...", end=" ", flush=True)
        start_time = time.time()
        ws = wb[sheet_name]
        
        # Get headers from first row
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        
        # Read data in chunks row by row
        chunk_data = []
        row_count = 0
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            chunk_data.append(row)
            row_count += 1
            
            # When chunk is full, convert to DataFrame
            if len(chunk_data) >= chunksize:
                df_chunk = pd.DataFrame(chunk_data, columns=headers)
                df_chunk["SourceSheet"] = sheet_name
                all_chunks.append(df_chunk)
                chunk_data = []
        
        # Add remaining rows
        if chunk_data:
            df_chunk = pd.DataFrame(chunk_data, columns=headers)
            df_chunk["SourceSheet"] = sheet_name
            all_chunks.append(df_chunk)
        
        elapsed = time.time() - start_time
        print(f"✓ {row_count:,} rows ({elapsed:.1f}s)")
    
    wb.close()
    
    print(f"\n🔗 Concatenating {len(all_chunks)} chunk(s)...", end=" ", flush=True)
    total = pd.concat(all_chunks, ignore_index=True)
    print(f"✓")
    return total


def build_agg_dict(df):
    """
    Build aggregation rules:
    - numeric columns with '%' or 'AUR' -> mean
    - other numeric columns -> sum
    """
    numeric_cols = df.select_dtypes(include="number").columns.tolist()

    # never aggregate grouping columns
    ignore_cols = ["Year", "Month", "Belk Style #", "Week", "Week (Month)"]
    for col in ignore_cols:
        if col in numeric_cols:
            numeric_cols.remove(col)

    agg = {}
    for col in numeric_cols:
        col_lower = col.lower()
        if "%" in col or "aur" in col_lower:
            agg[col] = "mean"
        else:
            agg[col] = "sum"
    return agg


def aggregate_chunked(data, year=None, chunksize=100000):
    """Aggregate in chunks to reduce memory usage, preserving all columns."""
    print(f"🔧 Starting aggregation ({len(data):,} rows)...")
    
    if year is not None:
        data = data[data["Year"] == year]
        print(f"   Filtered to year {year}: {len(data):,} rows")

    if data.empty:
        print("⚠️  No data found for filter.")
        return data

    group_cols = ["Year", "Month", "Belk Style #"]
    
    # Build aggregation dict for numeric columns
    agg_dict = build_agg_dict(data)
    
    # For non-numeric columns (like Category, Product, Description), take first value
    non_numeric_cols = data.select_dtypes(exclude="number").columns.tolist()
    for col in non_numeric_cols:
        if col not in group_cols and col != "SourceSheet":
            agg_dict[col] = "first"  # Take first occurrence
    
    print(f"   Aggregating {len(agg_dict)} columns (numeric + text)...")
    
    # Process in chunks if data is large
    if len(data) > chunksize:
        print(f"   Processing in chunks of {chunksize:,}...")
        aggregated_chunks = []
        
        for i in range(0, len(data), chunksize):
            chunk = data.iloc[i:i+chunksize]
            agg_chunk = chunk.groupby(group_cols, as_index=False).agg(agg_dict)
            aggregated_chunks.append(agg_chunk)
            print(f"   Chunk {i//chunksize + 1}: {len(chunk):,} -> {len(agg_chunk):,} rows")
        
        # Combine and re-aggregate
        print("   Final aggregation...")
        combined = pd.concat(aggregated_chunks, ignore_index=True)
        
        # For final aggregation, numeric columns use same rules, text columns use 'first'
        final_agg_dict = build_agg_dict(combined)
        for col in non_numeric_cols:
            if col not in group_cols and col != "SourceSheet" and col in combined.columns:
                final_agg_dict[col] = "first"
        
        result = combined.groupby(group_cols, as_index=False).agg(final_agg_dict)
    else:
        result = data.groupby(group_cols, as_index=False).agg(agg_dict)
    
    result = result.sort_values(["Year", "Month", "Belk Style #"])
    print(f"✅ Aggregation complete: {len(result):,} rows with {len(result.columns)} columns")
    return result


if __name__ == "__main__":
    import time
    start = time.time()
    
    print("="*60)
    print("🚀 BELK DATA AGGREGATION (CHUNK-BASED)")
    print("="*60)
    
    # 1) Load all sheets with chunking
    all_data = load_all_sheets_chunked(EXCEL_FILE, SHEETS_TO_USE, CHUNK_SIZE)
    print(f"📊 Total loaded: {len(all_data):,} rows\n")

    # 2) Aggregate with chunking
    final_output = aggregate_chunked(all_data, TARGET_YEAR, CHUNK_SIZE)

    # 3) Save result
    print(f"\n💾 Saving to: {OUTPUT_FILE}...")
    final_output.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
    
    elapsed = time.time() - start
    print("="*60)
    print(f"✅ SUCCESS! Saved {len(final_output):,} rows")
    print(f"⏱️  Total time: {elapsed:.1f}s")
    print("="*60)
