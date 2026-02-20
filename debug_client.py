# debug_client.py
import sys
import os
sys.path.append('.')

import pandas as pd
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def test_column_finding():
    """Test column finding for your specific file"""
    from app.analysis import _normalize_header, _read_excel_cached
    
    # Update with your actual file path
    file_path = r"C:\DQA\data\uploaded_files\3e8582ce__Client_Level.xls"
    
    print(f"=== Testing file: {os.path.basename(file_path)} ===")
    
    # Read the file
    df = _read_excel_cached(file_path)
    
    print(f"\nTotal columns: {len(df.columns)}")
    print(f"Total rows: {len(df)}")
    
    print("\n=== ALL COLUMNS ===")
    for i, col in enumerate(df.columns):
        col_letter = chr(65 + i) if i < 26 else f"{chr(65 + (i//26 - 1))}{chr(65 + (i%26))}"
        normalized = _normalize_header(col)
        print(f"{i:3d} ({col_letter:>3s}): '{col}' -> normalized: '{normalized}'")
    
    # Look for specific columns
    target_start = "Sex"
    target_end = "Biometric enrollment form available in clients folder"
    
    normalized_start = _normalize_header(target_start)
    normalized_end = _normalize_header(target_end)
    
    print(f"\n=== LOOKING FOR COLUMNS ===")
    print(f"Looking for start: '{target_start}' -> normalized: '{normalized_start}'")
    print(f"Looking for end: '{target_end}' -> normalized: '{normalized_end}'")
    
    found_start = None
    found_end = None
    
    for i, col in enumerate(df.columns):
        normalized_col = _normalize_header(col)
        if normalized_col == normalized_start:
            found_start = i
            print(f"  ✓ Found start column at index {i} (column {chr(65+i)})")
        if normalized_col == normalized_end:
            found_end = i
            print(f"  ✓ Found end column at index {i} (column {chr(65+i)})")
    
    if found_start is not None and found_end is not None:
        print(f"\n=== COLUMN RANGE ===")
        print(f"Start: column {found_start} ({chr(65+found_start)})")
        print(f"End: column {found_end} ({chr(65+found_end)})")
        print(f"Number of columns: {found_end - found_start + 1}")
        
        # Show data from these columns
        print(f"\n=== SAMPLE DATA (first 5 rows) ===")
        for i in range(min(5, len(df))):
            row_data = []
            for j in range(found_start, found_end + 1):
                value = df.iloc[i, j]
                row_data.append(f"{'T' if str(value).upper() == 'TRUE' else 'F'}")
            print(f"Row {i}: {', '.join(row_data)}")
    else:
        print("\n✗ Could not find one or both columns!")
        if found_start is None:
            print(f"  - Start column '{target_start}' not found")
            print(f"  - Looking for normalized: '{normalized_start}'")
        if found_end is None:
            print(f"  - End column '{target_end}' not found")
            print(f"  - Looking for normalized: '{normalized_end}'")
        
        # Show similar columns
        print(f"\n=== SIMILAR COLUMNS FOUND ===")
        for i, col in enumerate(df.columns):
            normalized_col = _normalize_header(col)
            if 'sex' in normalized_col or 'gender' in normalized_col:
                print(f"  Potential start column: '{col}' -> '{normalized_col}'")
            if 'biometric' in normalized_col or 'enrollment' in normalized_col:
                print(f"  Potential end column: '{col}' -> '{normalized_col}'")

if __name__ == "__main__":
    test_column_finding()