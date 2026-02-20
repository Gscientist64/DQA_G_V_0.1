# debug_vl.py
import pandas as pd
import sys

def debug_vl_file():
    file_path = r"C:\DQA\data\uploaded_files\3e8582ce_VL_Unsupressed.xlsx"  # Update with your actual VL file
    
    print(f"=== DEBUGGING VL FILE ===")
    
    try:
        df = pd.read_excel(file_path, engine='xlrd')
        
        print(f"\nShape: {df.shape[0]} rows, {df.shape[1]} columns")
        
        print("\n=== ALL COLUMNS ===")
        for i, col in enumerate(df.columns):
            col_letter = chr(65 + i) if i < 26 else f"{chr(65 + (i//26 - 1))}{chr(65 + (i%26))}"
            print(f"{i:3d} ({col_letter:>3s}): '{col}'")
        
        # Look for Sex column
        print("\n=== LOOKING FOR SEX COLUMNS ===")
        sex_columns = []
        for i, col in enumerate(df.columns):
            if 'sex' in str(col).lower():
                col_letter = chr(65 + i) if i < 26 else f"{chr(65 + (i//26 - 1))}{chr(65 + (i%26))}"
                sex_columns.append((i, col_letter, col))
                print(f"  Found at {i} ({col_letter}): '{col}'")
        
        # Look for EAC column
        print("\n=== LOOKING FOR EAC COLUMNS ===")
        eac_columns = []
        for i, col in enumerate(df.columns):
            if 'eac' in str(col).lower() or 'vl result' in str(col).lower():
                col_letter = chr(65 + i) if i < 26 else f"{chr(65 + (i//26 - 1))}{chr(65 + (i%26))}"
                eac_columns.append((i, col_letter, col))
                print(f"  Found at {i} ({col_letter}): '{col}'")
        
        # Show TRUE/FALSE data from potential columns
        if sex_columns and eac_columns:
            start_idx = sex_columns[-1][0]  # Last sex column (should be the TRUE/FALSE one)
            end_idx = eac_columns[-1][0]    # Last EAC column
            
            print(f"\n=== PROPOSED COLUMN RANGE ===")
            print(f"Start: {start_idx} ({sex_columns[-1][1]}) - '{sex_columns[-1][2]}'")
            print(f"End: {end_idx} ({eac_columns[-1][1]}) - '{eac_columns[-1][2]}'")
            print(f"Number of columns: {end_idx - start_idx + 1}")
            
            # Check if this range looks like TRUE/FALSE data
            print(f"\n=== CHECKING DATA IN THIS RANGE ===")
            sub_df = df.iloc[:, start_idx:end_idx + 1]
            
            print(f"First 3 rows:")
            for i in range(min(3, len(sub_df))):
                row_data = []
                for val in sub_df.iloc[i]:
                    if pd.isna(val):
                        row_data.append("NaN")
                    elif isinstance(val, bool):
                        row_data.append("TRUE" if val else "FALSE")
                    elif isinstance(val, (int, float)):
                        if val == 1 or val == 1.0:
                            row_data.append("1.0")
                        elif val == 0 or val == 0.0:
                            row_data.append("0.0")
                        else:
                            row_data.append(str(val))
                    else:
                        row_data.append(str(val)[:10])
                print(f"  Row {i}: {row_data}")
            
            # Count TRUE/FALSE
            print(f"\n=== TRUE/FALSE COUNTS ===")
            for j, col in enumerate(sub_df.columns):
                col_data = sub_df[col]
                true_count = 0
                total = 0
                
                for val in col_data:
                    if pd.isna(val):
                        continue
                    total += 1
                    if isinstance(val, bool) and val:
                        true_count += 1
                    elif isinstance(val, (int, float)) and (val == 1 or val == 1.0):
                        true_count += 1
                    elif isinstance(val, str) and val.upper() == "TRUE":
                        true_count += 1
                    elif isinstance(val, str) and val == "1":
                        true_count += 1
                
                if total > 0:
                    pct = (true_count / total) * 100
                    print(f"  Column {j}: {true_count}/{total} TRUE ({pct:.1f}%)")
                    
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_vl_file()