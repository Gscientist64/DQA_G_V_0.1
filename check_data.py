# check_data.py
import json
import os

def check_data_files():
    """Check the data files we have."""
    data_dir = 'data'
    
    if not os.path.exists(data_dir):
        print("No data directory found!")
        return
    
    print("Found data directory with files:")
    for file in os.listdir(data_dir):
        filepath = os.path.join(data_dir, file)
        if os.path.isfile(filepath):
            size = os.path.getsize(filepath)
            print(f"  - {file} ({size} bytes)")
            
            if file.endswith('.json'):
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    print(f"    Contains {len(data)} records" if isinstance(data, list) else "    JSON data loaded")
                except:
                    print("    Could not read JSON")

if __name__ == "__main__":
    check_data_files()