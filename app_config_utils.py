# app_config_utils.py
import json
import os
import sys
import logging

CONFIG_FILE_NAME = "user_settings.json"

DEFAULT_SETTINGS = {
    "db_path": None,  # User must configure this
    "user_name": "Default User",
    "work_dir": os.getcwd() # Default to current working directory
}

def get_config_file_path():
    """Determines the path to the user_settings.json file.
    It's placed in the same directory as the executable or the script being run.
    """
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (e.g., PyInstaller)
        application_path = os.path.dirname(sys.executable)
    else:
        # If run as a script
        application_path = os.path.dirname(os.path.abspath(__file__))
        # If app_config_utils.py is in a subdirectory of the main app dir, adjust path:
        # For instance, if main script is in '.../app/' and this is in '.../app/utils/'
        # application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(application_path, CONFIG_FILE_NAME)

def load_user_config():
    """Loads user configurations from the JSON file.
    Returns default settings if the file doesn't exist or is invalid.
    """
    config_path = get_config_file_path()
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                # Ensure all default keys are present, merge with defaults
                # User's settings take precedence
                merged_settings = DEFAULT_SETTINGS.copy()
                merged_settings.update(settings)
                logging.info(f"User settings loaded from {config_path}")
                return merged_settings
        else:
            logging.info(f"Settings file not found at {config_path}. Using default settings.")
            return DEFAULT_SETTINGS.copy()
    except (json.JSONDecodeError, IOError) as e:
        logging.error(f"Error loading settings from {config_path}: {e}. Using default settings.")
        return DEFAULT_SETTINGS.copy()

def save_user_config(settings_dict):
    """Saves the provided settings dictionary to the JSON file."""
    config_path = get_config_file_path()
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(settings_dict, f, indent=4)
        logging.info(f"User settings saved to {config_path}")
        return True
    except IOError as e:
        logging.error(f"Error saving settings to {config_path}: {e}")
        return False

if __name__ == '__main__':
    # Example usage (for testing this module directly)
    print(f"Config file will be at: {get_config_file_path()}")
    
    # Load current or default settings
    current_settings = load_user_config()
    print(f"Loaded settings: {current_settings}")

    # Modify a setting
    # current_settings["user_name"] = "Test User One"
    # current_settings["db_path"] = "/path/to/my/database.db"
    # current_settings["work_dir"] = "/my/test/workdir"

    # Save settings
    # if save_user_config(current_settings):
    #     print("Settings saved.")
    #     reloaded_settings = load_user_config()
    #     print(f"Reloaded settings: {reloaded_settings}")
    # else:
    #     print("Failed to save settings.")