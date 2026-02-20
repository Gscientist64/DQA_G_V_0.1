# app/storage.py
import os
import json
import logging
import sys
import shutil
from datetime import datetime

logger = logging.getLogger(__name__)

def get_app_base_dir():
    """
    Get the base directory where the application is running from.
    In standalone: folder containing the EXE
    In development: project root folder
    """
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller executable
        return os.path.dirname(sys.executable)
    else:
        # Running in development
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_bundled_data_path():
    """
    Get the path to bundled data files inside the EXE.
    """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller >= 5.0
        return os.path.join(sys._MEIPASS, "data")
    else:
        # Development or older PyInstaller
        return os.path.join(get_app_base_dir(), "data")

def setup_application_data():
    """
    Setup the application's data directory structure.
    Copies initial data from bundled resources if needed.
    Returns the path to the application's data directory.
    """
    base_dir = get_app_base_dir()
    app_data_dir = os.path.join(base_dir, "DQA_Data")
    
    # Create main data directory
    os.makedirs(app_data_dir, exist_ok=True)
    
    # Create subdirectories
    subdirs = ["uploaded_files", "backups", "exports"]
    for subdir in subdirs:
        os.makedirs(os.path.join(app_data_dir, subdir), exist_ok=True)
    
    # Copy essential data files from bundled resources if they don't exist
    bundled_data_path = get_bundled_data_path()
    
    # List of essential files to copy on first run
    essential_files = ["results.json"]
    
    for filename in essential_files:
        src_path = os.path.join(bundled_data_path, filename)
        dst_path = os.path.join(app_data_dir, filename)
        
        # Only copy if source exists and destination doesn't
        if os.path.exists(src_path) and not os.path.exists(dst_path):
            try:
                shutil.copy2(src_path, dst_path)
                logger.info(f"Copied initial data: {filename}")
            except Exception as e:
                logger.warning(f"Could not copy {filename}: {e}")
    
    return app_data_dir

# Initialize application data directory
APP_DATA_DIR = setup_application_data()

# Define file paths relative to application data directory
RESULTS_FILE = os.path.join(APP_DATA_DIR, "results.json")
UPLOAD_DIR = os.path.join(APP_DATA_DIR, "uploaded_files")
BACKUP_DIR = os.path.join(APP_DATA_DIR, "backups")
EXPORT_DIR = os.path.join(APP_DATA_DIR, "exports")

logger.info(f"Application data directory: {APP_DATA_DIR}")
logger.info(f"Results file: {RESULTS_FILE}")

# ========== REST OF YOUR EXISTING CODE ==========
# Keep all your existing functions but update them to use the new constants

def _ensure_data_dir():
    """Ensure all required directories exist."""
    # Already handled by setup_application_data()
    pass

def _create_backup():
    """Create a backup of the results file before modification."""
    if not os.path.exists(RESULTS_FILE):
        return
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(BACKUP_DIR, f"results_backup_{timestamp}.json")
    
    try:
        shutil.copy2(RESULTS_FILE, backup_file)
        logger.debug(f"Created backup: {backup_file}")
    except Exception as e:
        logger.warning(f"Failed to create backup: {e}")


def load_results():
    """
    Load all facility analyses from results.json.
    Returns a list of dicts. If file doesn't exist, returns [].
    """
    if not os.path.exists(RESULTS_FILE):
        logger.info(f"Results file does not exist: {RESULTS_FILE}")
        return []

    try:
        with open(RESULTS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # DEBUG: Log what we found
        logger.info(f"Loaded data from {RESULTS_FILE}")
        logger.info(f"Data type: {type(data)}")
        
        # Handle the new format (dictionary with "facilities" key)
        if isinstance(data, dict):
            logger.info(f"Data keys: {list(data.keys())}")
            if "facilities" in data:
                facilities_list = data["facilities"]
                if isinstance(facilities_list, list):
                    logger.info(f"✓ Found {len(facilities_list)} facilities")
                    return facilities_list
                else:
                    logger.warning(f"'facilities' key is not a list, type: {type(facilities_list)}")
                    return []
            else:
                logger.warning("Data is dict but no 'facilities' key found")
                return []
        
        # Handle old format (direct list)
        elif isinstance(data, list):
            logger.info(f"✓ Found {len(data)} facilities in direct list format")
            return data
        
        else:
            logger.warning(f"Unexpected data type in results.json: {type(data)}")
            return []
            
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse results.json: {e}")
        return []
    except Exception as e:
        logger.error(f"Unexpected error loading results: {e}")
        return []    


def save_results(facilities):
    """
    Save the list of facility analyses back to results.json.
    Creates a backup before saving.
    """
    # Create backup before saving
    _create_backup()
    
    # Add metadata
    save_data = {
        "version": "1.1",
        "last_updated": datetime.now().isoformat(),
        "total_facilities": len(facilities),
        "facilities": facilities
    }
    
    try:
        # Use temporary file to prevent corruption
        temp_file = RESULTS_FILE + ".tmp"
        
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(save_data, f, ensure_ascii=False, indent=2)
        
        # Atomic replace
        os.replace(temp_file, RESULTS_FILE)
        
        logger.info(f"Successfully saved {len(facilities)} facilities to {RESULTS_FILE}")
        
        # Clean up old backups (keep last 10)
        _cleanup_old_backups()
        
    except Exception as e:
        logger.error(f"Failed to save results: {e}")
        raise


def _cleanup_old_backups(max_backups=10):
    """Clean up old backup files, keeping only the most recent ones."""
    try:
        if not os.path.exists(BACKUP_DIR):
            return
        
        backup_files = []
        for filename in os.listdir(BACKUP_DIR):
            if filename.startswith("results_backup_") and filename.endswith(".json"):
                filepath = os.path.join(BACKUP_DIR, filename)
                backup_files.append((filepath, os.path.getmtime(filepath)))
        
        # Sort by modification time (oldest first)
        backup_files.sort(key=lambda x: x[1])
        
        # Remove oldest backups if we have more than max_backups
        if len(backup_files) > max_backups:
            files_to_remove = len(backup_files) - max_backups
            for i in range(files_to_remove):
                os.remove(backup_files[i][0])
                logger.debug(f"Removed old backup: {backup_files[i][0]}")
                
    except Exception as e:
        logger.warning(f"Failed to cleanup old backups: {e}")


def get_facility_by_name(facility_name):
    """
    Get a specific facility result by name.
    Returns None if not found.
    """
    facilities = load_results()
    
    for facility in facilities:
        if facility.get("Facility") == facility_name:
            return facility
    
    return None


def delete_facility(facility_name):
    """
    Delete a facility result by name.
    Returns True if deleted, False if not found.
    """
    facilities = load_results()
    
    # Filter out the facility to delete
    original_count = len(facilities)
    filtered_facilities = [f for f in facilities if f.get("Facility") != facility_name]
    
    if len(filtered_facilities) == original_count:
        logger.info(f"Facility '{facility_name}' not found for deletion")
        return False
    
    # Save the updated list
    save_results(filtered_facilities)
    
    logger.info(f"Deleted facility '{facility_name}'")
    return True


def get_summary_statistics():
    """
    Get summary statistics from all facilities.
    """
    facilities = load_results()
    
    if not facilities:
        return {
            "total_facilities": 0,
            "average_score": 0,
            "score_distribution": {},
            "ratings_distribution": {},
            "last_updated": 0,
            "last_updated_formatted": "Never"
        }
    
    # Calculate average scores
    total_score = sum(f.get("Score", 0) for f in facilities)
    avg_score = total_score / len(facilities)
    
    # Score distribution
    score_distribution = {
        "90-100": sum(1 for f in facilities if 90 <= f.get("Score", 0) <= 100),
        "75-89": sum(1 for f in facilities if 75 <= f.get("Score", 0) < 90),
        "50-74": sum(1 for f in facilities if 50 <= f.get("Score", 0) < 75),
        "0-49": sum(1 for f in facilities if 0 <= f.get("Score", 0) < 50),
    }
    
    # Ratings distribution
    ratings = {}
    for f in facilities:
        rating = f.get("rating", "Unknown")
        ratings[rating] = ratings.get(rating, 0) + 1
    
    # Last updated timestamp
    last_updated = max(f.get("processing_timestamp", 0) for f in facilities)
    
    # Format last updated date
    try:
        last_updated_formatted = datetime.fromtimestamp(last_updated).strftime("%Y-%m-%d")
    except:
        last_updated_formatted = "Unknown"
    
    return {
        "total_facilities": len(facilities),
        "average_score": round(avg_score, 2),
        "score_distribution": score_distribution,
        "ratings_distribution": ratings,
        "last_updated": last_updated,
        "last_updated_formatted": last_updated_formatted
    }


def clear_all_results():
    """
    Clear all results (with backup).
    Returns True if successful.
    """
    try:
        _create_backup()
        
        if os.path.exists(RESULTS_FILE):
            os.remove(RESULTS_FILE)
            logger.info("Cleared all results")
        
        return True
    except Exception as e:
        logger.error(f"Failed to clear results: {e}")
        return False


def export_results_to_json(output_path=None):
    """
    Export all results to a JSON file.
    If output_path is None, creates a timestamped file in EXPORT_DIR.
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(EXPORT_DIR, f"results_export_{timestamp}.json")
    
    facilities = load_results()
    
    export_data = {
        "export_version": "1.0",
        "export_timestamp": datetime.now().isoformat(),
        "total_records": len(facilities),
        "data": facilities
    }
    
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(export_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"Exported {len(facilities)} records to {output_path}")
        return output_path
    except Exception as e:
        logger.error(f"Failed to export results: {e}")
        raise