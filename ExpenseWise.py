"""ExpenseWise - A Personal Finance Tracker Application using Tkinter"""

import tkinter as tk
from tkinter import ttk, messagebox
import datetime
import random
import csv
import os
import json
import logging

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Configuration ---

# Theme Definitions
THEME_DARK = {
    "background": "#141414",
    "foreground": "#e0e0e0",
    "sidebar": "#2a5c54",
    "card": "#222222",
    "accent": "#4CAF50",
    "accent_darker": "#388E3C",
    "red": "#E50914",
    "blue": "#2196F3",
    "yellow": "#FFEB3B",
    "disabled": "#757575",
    "account_icon_bg": "#333333",
    "dialog_bg": "#1f1f1f",
    "dialog_fg": "#e0e0e0",
    "dialog_card": "#333333",
    "treeview_heading_bg": "#141414",
    "scrollbar_trough": "#222222",
    "scrollbar_bg": "#141414",
    "button_fg": "#ffffff",
    "combobox_list_bg": "#222222",
    "combobox_list_fg": "#e0e0e0",
    "combobox_list_select_bg": "#4CAF50",
    "combobox_list_select_fg": "#ffffff",
}

THEME_LIGHT = {
    "background": "#f5f5f5",
    "foreground": "#212121",
    "sidebar": "#e0f2f1",
    "card": "#ffffff",
    "accent": "#00796b",
    "accent_darker": "#004d40",
    "red": "#d32f2f",
    "blue": "#1976D2",
    "yellow": "#FBC02D",
    "disabled": "#bdbdbd",
    "account_icon_bg": "#e0e0e0",
    "dialog_bg": "#f0f0f0",
    "dialog_fg": "#212121",
    "dialog_card": "#ffffff",
    "treeview_heading_bg": "#f5f5f5",
    "scrollbar_trough": "#ffffff",
    "scrollbar_bg": "#f5f5f5",
    "button_fg": "#ffffff",
    "combobox_list_bg": "#ffffff",
    "combobox_list_fg": "#212121",
    "combobox_list_select_bg": "#00796b",
    "combobox_list_select_fg": "#ffffff",
}

# Global Theme Variables (Defaults to Dark)
theme_colors = THEME_DARK.copy()

# --- Fonts ---
FONT_FAMILY = "Segoe UI"
FONT_NORMAL = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10, "bold")
FONT_LARGE = (FONT_FAMILY, 14, "bold")
FONT_XLARGE = (FONT_FAMILY, 18, "bold")
FONT_XXLARGE = (FONT_FAMILY, 24, "bold")
FONT_ACCOUNT_NAME = (FONT_FAMILY, 12)
FONT_TITLE = (FONT_FAMILY, 12, "bold")
FONT_SMALL = (FONT_FAMILY, 8)

# --- Data Store ---
app_data = {
    "user_profiles": {},
    "current_user_id": None,
    "wallets": {},
    "budgets": {},
    "goals": {},
    "transactions": [],
    "activity_log": [],
    "settings": {"theme": "dark"},
    "categories": {},
}

# --- Core Categories ---
BASE_CATEGORIES = {
    "rent": {"name": "Rent/Mortgage", "icon": "🏠", "type": "expense"},
    "groceries": {"name": "Groceries", "icon": "🛒", "type": "expense"},
    "utilities": {"name": "Utilities", "icon": "💡", "type": "expense"},
    "transport": {"name": "Transportation", "icon": "🚗", "type": "expense"},
    "dining": {"name": "Dining", "icon": "🍽️", "type": "expense"},
    "personal_care": {"name": "Personal Care", "icon": "💇", "type": "expense"},
    "healthcare": {"name": "Healthcare","icon":"⚕","type":"expense"},
    "shopping": {"name": "Shopping", "icon": "🛍️", "type": "expense"},
    "entertainment": {"name": "Entertainment", "icon": "🎬", "type": "expense"},
    "internet": {"name": "Internet", "icon": "🌐", "type": "expense"},
    "subscriptions": {"name": "Subscriptions", "icon": "📺", "type": "expense"},
    "home_improvement": {"name": "Home", "icon": "🛠️", "type": "expense"},
    "education": {"name": "Education", "icon": "📚", "type": "expense"},
    "travel": {"name": "Travel", "icon": "✈️", "type": "expense"},
    "others": {"name": "Other", "icon": "❓", "type": "expense"},

    # Income Categories
    "salary": {"name": "Salary", "icon": "💼", "type": "income"},
    "freelance": {"name": "Freelance", "icon": "💡", "type": "income"},
    "investment": {"name": "Investment", "icon": "📈", "type": "income"},
    "business_income": {"name": "Business", "icon": "🏢", "type": "income"},
    "rental_income": {"name": "Rental", "icon": "🏘️", "type": "income"},
    "refunds" : {"name": "Refunds", "icon": "💰","type":"income"},
    "allowance": {"name": "Allowance", "icon": "💸", "type": "income"},
    "gifts": {"name": "Gifts", "icon": "💝", "type": "income"},
    "other_income": {"name": "Other", "icon": "➕", "type": "income"},
}

# --- File Paths & Constants ---
DATA_DIR = "ExpenseWiseData"
USER_PROFILES_CSV = os.path.join(DATA_DIR, "user_profiles.csv")
ACCOUNT_ICON_COLORS = ["#E57373", "#81C784", "#64B5F6", "#FFD54F", "#BA68C8", "#4DB6AC", "#F06292", "#A1887F"]
MAX_ACTIVITY_LOG_SIZE = 150

# --- Utility Functions ---
def create_stylish_button(parent, text, command, style="TButton", **kwargs):
    """Creates a ttk button with common styling."""
    return ttk.Button(parent, text=text, command=command, style=style, **kwargs)

def create_card_frame(parent):
    """Creates a standard card frame."""
    return tk.Frame(parent, bg=theme_colors["card"], relief=tk.FLAT, bd=0)

def format_currency(amount):
    """Formats a number as Philippine Peso currency."""
    if amount is None:
        return "₱ N/A"
    try:
        return f"₱ {float(amount):,.2f}"
    except (ValueError, TypeError):
        logging.warning(f"Invalid amount for currency formatting: {amount}")
        return "₱ Invalid"

def log_activity(action):
    """Adds an entry to the activity log for the current user."""
    user_id = app_data.get("current_user_id")
    if not user_id:
        logging.warning("Attempted to log activity with no user selected.")
        return
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {"timestamp": timestamp, "action": action}

    if not isinstance(app_data.get("activity_log"), list):
        app_data["activity_log"] = [] # Ensure log is a list

    app_data["activity_log"].append(log_entry)
    # Limit activity log size
    if len(app_data["activity_log"]) > MAX_ACTIVITY_LOG_SIZE:
        app_data["activity_log"].pop(0)

def get_unique_id(prefix):
    """Generates a simple unique ID (timestamp + random)."""
    return f"{prefix}_{int(datetime.datetime.now().timestamp())}_{random.randint(1000, 9999)}"

def ensure_data_dir():
    """Creates the data directory if it doesn't exist."""
    if not os.path.exists(DATA_DIR):
        try:
            os.makedirs(DATA_DIR)
            logging.info(f"Created data directory: {DATA_DIR}")
        except OSError as e:
            logging.error(f"Could not create data directory '{DATA_DIR}': {e}")
            messagebox.showerror("Directory Error", f"Could not create data directory '{DATA_DIR}':\n{e}\nApplication cannot continue.")
            exit(1) # Critical failure

# --- CSV/JSON Handling for User Profiles ---
def load_user_profiles_from_csv():
    """Loads user profile data from user_profiles.csv."""
    ensure_data_dir()
    profiles = {}
    created_demo = False
    required_fields = ['user_id', 'name', 'icon_color']

    if not os.path.exists(USER_PROFILES_CSV):
        logging.warning(f"'{USER_PROFILES_CSV}' not found. Creating demo user profile.")
        demo_id = get_unique_id("user_demo")
        profiles[demo_id] = {"name": "Demo User", "icon_color": random.choice(ACCOUNT_ICON_COLORS)}
        created_demo = True
    else:
        try:
            with open(USER_PROFILES_CSV, mode='r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                # Check for required CSV columns
                if not reader.fieldnames or not all(col in reader.fieldnames for col in required_fields):
                    raise ValueError("User profiles CSV is missing required columns (user_id, name, icon_color).")

                for row_num, row in enumerate(reader, 1):
                    try:
                        user_id = row.get('user_id', '').strip()
                        name = row.get('name', '').strip()
                        icon_color = row.get('icon_color', random.choice(ACCOUNT_ICON_COLORS)).strip()

                        if not user_id or not name:
                            logging.warning(f"Skipping invalid row {row_num} in user profiles CSV: {row}")
                            continue
                        # Basic color validation
                        if not icon_color.startswith('#') or len(icon_color) != 7:
                             icon_color = random.choice(ACCOUNT_ICON_COLORS)
                             logging.warning(f"Invalid icon_color in row {row_num}, assigning random.")

                        profiles[user_id] = {"name": name, "icon_color": icon_color}
                    except Exception as e:
                        logging.error(f"Error processing user profile row {row_num} ({row}): {e}. Skipping.")
                        continue
        except FileNotFoundError:
            logging.error(f"'{USER_PROFILES_CSV}' disappeared during read attempt. Creating demo profile.")
            demo_id = get_unique_id("user_demo")
            profiles[demo_id] = {"name": "Demo User", "icon_color": random.choice(ACCOUNT_ICON_COLORS)}
            created_demo = True
        except (ValueError, csv.Error, Exception) as e:
            logging.exception(f"Failed to load user profiles from '{USER_PROFILES_CSV}': {e}")
            messagebox.showerror("CSV Load Error",
                                 f"Failed to load user profiles from '{USER_PROFILES_CSV}':\n{e}\n\nPlease check the file or delete it to start fresh with a demo user.")
            # Load demo as fallback on error
            profiles.clear()
            demo_id = get_unique_id("user_demo")
            profiles[demo_id] = {"name": "Demo User", "icon_color": random.choice(ACCOUNT_ICON_COLORS)}
            created_demo = True

    # Ensure at least one profile exists, even after errors
    if not profiles:
        logging.warning("No valid user profiles loaded or file was empty. Creating demo user profile.")
        demo_id = get_unique_id("user_demo")
        profiles[demo_id] = {"name": "Demo User", "icon_color": random.choice(ACCOUNT_ICON_COLORS)}
        created_demo = True

    app_data["user_profiles"] = profiles
    if created_demo:
        save_user_profiles_to_csv() # Save the newly created demo user

def save_user_profiles_to_csv():
    """Saves the current app_data['user_profiles'] to user_profiles.csv."""
    ensure_data_dir()
    profiles_to_save = app_data.get("user_profiles")
    if not profiles_to_save:
        logging.warning("Attempted to save user profiles, but none are loaded in memory.")
        try:
            with open(USER_PROFILES_CSV, mode='w', newline='', encoding='utf-8') as csvfile:
                 fieldnames = ['user_id', 'name', 'icon_color']
                 writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                 writer.writeheader()
            logging.info(f"Created empty user profiles file with header: '{USER_PROFILES_CSV}'.")
        except IOError as e:
            logging.error(f"Could not write header to empty '{USER_PROFILES_CSV}': {e}")
        return

    try:
        with open(USER_PROFILES_CSV, mode='w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['user_id', 'name', 'icon_color']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for user_id, details in profiles_to_save.items():
                 if not isinstance(details, dict):
                     logging.warning(f"Skipping saving invalid profile data for ID {user_id}: {details}")
                     continue
                 row_data = {
                    'user_id': user_id,
                    'name': details.get('name', f'Unnamed User {user_id}'),
                    'icon_color': details.get('icon_color', random.choice(ACCOUNT_ICON_COLORS))
                 }
                 writer.writerow(row_data)
        logging.info(f"User profiles saved successfully to '{USER_PROFILES_CSV}'.")
    except IOError as e:
        logging.error(f"Could not write to '{USER_PROFILES_CSV}': {e}")
        messagebox.showerror("CSV Save Error", f"Could not write user profiles to '{USER_PROFILES_CSV}':\n{e}")
    except Exception as e:
        logging.exception(f"An unexpected error occurred while saving user profiles: {e}")
        messagebox.showerror("CSV Save Error", f"An unexpected error occurred while saving user profiles:\n{e}")


# --- CSV/JSON Handling for Specific User Data ---
def get_user_data_file_path(user_id, data_type):
    """Generates the file path for a specific user's data type."""
    ensure_data_dir()
    base_filename = f"{data_type}_{user_id}"
    extension = ".json" if data_type == "settings" else ".csv"
    return os.path.join(DATA_DIR, f"{base_filename}{extension}")

# --- Data Loading Helpers ---
def _load_json_data(file_path, default_value=None):
    """Loads data from a JSON file."""
    if default_value is None: default_value = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        logging.warning(f"JSON file not found: {file_path}. Returning default.")
        return default_value
    except json.JSONDecodeError as e:
        logging.error(f"Error decoding JSON from {file_path}: {e}. Returning default.")
        return default_value
    except Exception as e:
        logging.exception(f"Unexpected error loading JSON {file_path}: {e}")
        return default_value

def _load_csv_data(file_path, expected_fields, id_field=None, numeric_fields=None):
    """Loads data from a CSV file into a list or dictionary."""
    if numeric_fields is None: numeric_fields = []
    data_list = [] # Temporarily store all rows read

    logging.debug(f"Executing _load_csv_data for: {file_path}")
    logging.debug(f"  Expected fields: {expected_fields}")
    logging.debug(f"  ID field: {id_field}")

    try:
        with open(file_path, mode='r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)

            # Check header existence and content
            if not reader.fieldnames:
                 logging.warning(f"CSV file '{file_path}' appears empty or has no header. Returning empty data.")
                 return {} if id_field else []

            processed_rows = 0
            for row_num, row in enumerate(reader, 1):
                processed_row = {}
                try:
                    # Process only expected fields
                    for field in expected_fields:
                        val = row.get(field)

                        # Handle numeric conversion
                        if field in numeric_fields:
                            try:
                                processed_row[field] = float(val) if val not in [None, ''] else 0.0
                            except (ValueError, TypeError):
                                logging.warning(f"Invalid numeric value '{val}' for field '{field}' in row {row_num} of {file_path}. Using 0.0.")
                                processed_row[field] = 0.0

                        else:
                             processed_row[field] = val if val is not None else ''

                    data_list.append(processed_row)
                    processed_rows += 1
                except Exception as e:
                     logging.error(f"Error processing row {row_num} in {file_path}: {e}. Skipping row: {row}")
                     continue # Skip to the next row

            logging.debug(f"  Read and processed {processed_rows} rows from {file_path}.")

    except FileNotFoundError:
        logging.warning(f"CSV file not found: {file_path}. Returning empty data.")
        return {} if id_field else []
    except (csv.Error, Exception) as e:
        logging.exception(f"Error reading CSV file {file_path}: {e}")
        return {} if id_field else [] # Return empty on error

    # Dictionary Construction (if id_field is provided)
    if id_field:
        data_dict = {}
        duplicate_ids = 0
        missing_ids = 0
        successful_adds = 0
        if id_field not in expected_fields:
             logging.error(f"ID field '{id_field}' specified for {file_path} is not in expected_fields list: {expected_fields}. Cannot build dictionary.")
             return {} # Return empty dict as we can't key it

        for item in data_list:
            item_id = item.get(id_field)

            if item_id is not None and item_id != '':
                 if item_id in data_dict:
                     logging.warning(f"Duplicate ID '{item_id}' found in {file_path}. Overwriting with later entry: {item}")
                     duplicate_ids += 1
                 data_dict[item_id] = item
                 successful_adds += 1
            else:
                logging.warning(f"Skipping item with missing or empty ID field '{id_field}' in {file_path}: {item}")
                missing_ids += 1

        logging.debug(f"  Constructed dictionary for {file_path}: {successful_adds} items added, {duplicate_ids} duplicates overwritten, {missing_ids} missing IDs skipped.")
        return data_dict
    else:
        logging.debug(f"  Returning list for {file_path} (no ID field specified).")
        return data_list

def load_user_data(user_id):
    """Loads all data for the specified user_id into the global app_data."""
    logging.info(f"Loading data for user: {user_id}")
    app_data["current_user_id"] = user_id

    # Configure data types for loading (JSON or CSV, with expected fields)
    data_types_config = {
        "wallets": {"type": dict, "fields": ['wallet_id', 'name', 'balance'], "id_field": "wallet_id", "numeric_fields": ["balance"]},
        "budgets": {"type": dict, "fields": ['budget_id', 'name', 'allocated', 'cycle'], "id_field": "budget_id", "numeric_fields": ["allocated"]},
        "goals": {"type": dict, "fields": ['goal_id', 'name', 'target', 'saved', 'due_date'], "id_field": "goal_id", "numeric_fields": ["target", "saved"]},
        "transactions": {"type": list, "fields": ['date', 'time', 'timestamp', 'title', 'wallet', 'amount', 'category', 'type', 'from_account', 'to_account', 'linked_budget', 'linked_goal'], "numeric_fields": ["amount"]},
        "activity_log": {"type": list, "fields": ['timestamp', 'action']},
        "settings": {"type": dict, "is_json": True},
    }

    # Load data for each type
    for data_key, config in data_types_config.items():
        file_path = get_user_data_file_path(user_id, data_key)
        is_json = config.get("is_json", False)
        default_value = {} if config["type"] == dict else []
        id_field_to_use = config.get("id_field")

        logging.info(f"Attempting to load '{data_key}' from {file_path} (Expected type: {config['type']}, ID Field: {id_field_to_use})")

        if is_json:
            loaded_data = _load_json_data(file_path, default_value=default_value)
            if data_key == "settings":
                loaded_data.setdefault("theme", "dark") # Ensure default theme if missing
            app_data[data_key] = loaded_data
            logging.info(f"  Loaded JSON data for '{data_key}'.")
        else: # CSV
            loaded_data = _load_csv_data(
                file_path,
                config["fields"],
                id_field=id_field_to_use,
                numeric_fields=config.get("numeric_fields", [])
            )
            app_data[data_key] = loaded_data
            logging.info(f"  Loaded CSV data for '{data_key}'. Result type: {type(loaded_data)}, Length: {len(loaded_data) if hasattr(loaded_data, '__len__') else 'N/A'}")

        # Post-Load Handling & Defaults
        if data_key == "wallets" and not app_data[data_key]:
             logging.info(f"No wallets loaded for user {user_id}. Creating default 'Cash' wallet.")
             wallet_id = get_unique_id("wallet")
             if not isinstance(app_data[data_key], dict): app_data[data_key] = {}
             app_data[data_key][wallet_id] = {"wallet_id": wallet_id, "name": "Cash", "balance": 0.0}

    # Ensure Categories are Loaded (Global/Shared Structure)
    core_categories = BASE_CATEGORIES
    app_data["categories"] = core_categories
    logging.info("Global categories loaded/reset.")

    logging.info(f"Data loading finished for user: {user_id}")

# --- Data Saving Helpers ---
def _save_json_data(file_path, data):
    """Saves data to a JSON file."""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
        return True
    except (IOError, TypeError) as e:
        logging.error(f"Error saving JSON data to {file_path}: {e}")
        return False
    except Exception as e:
        logging.exception(f"Unexpected error saving JSON {file_path}: {e}")
        return False

def _save_csv_data(file_path, data, fields):
    """Saves list or dictionary data to a CSV file."""
    logging.debug(f"Executing _save_csv_data for: {file_path}")
    logging.debug(f"  Data type received: {type(data)}")
    if hasattr(data, '__len__'): logging.debug(f"  Data length: {len(data)}")
    logging.debug(f"  Fields to write: {fields}")

    if not fields:
        logging.error(f"Cannot save CSV data to {file_path}: 'fields' list is missing or empty.")
        return False

    try:
        list_to_save = []

        if isinstance(data, dict):
            # Assume the first field in 'fields' list is the ID field name
            if not fields:
                 logging.error(f"Field list is empty for dictionary data in {file_path}. Cannot determine ID field.")
                 return False
            id_field_name = fields[0]
            logging.debug(f"  Identified ID field for dictionary as: '{id_field_name}' (from fields[0])")

            valid_items = 0
            skipped_items = 0
            for item_id, item_data in data.items():
                if isinstance(item_data, dict):
                    row_dict = item_data.copy()
                    row_dict[id_field_name] = item_id
                    list_to_save.append(row_dict)
                    valid_items += 1
                else:
                    logging.warning(f"Skipping non-dictionary value for ID '{item_id}' in {file_path}.")
                    skipped_items += 1
            logging.debug(f"  Converted dictionary: {valid_items} valid items added to list, {skipped_items} skipped.")

        elif isinstance(data, list):
            valid_rows = [row for row in data if isinstance(row, dict)]
            if len(valid_rows) != len(data):
                logging.warning(
                    f"Some non-dictionary items found in list data for {file_path}. Only saving valid rows.")
            list_to_save = valid_rows
            logging.debug(f"  Saving list: Using {len(list_to_save)} valid rows for writing.")
        else:
            logging.error(
                f"Invalid data type ({type(data)}) provided for CSV saving to {file_path}. Expected list or dict.")
            return False

        # Writing the data
        with open(file_path, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fields, extrasaction='ignore', restval='')
            writer.writeheader()
            if not list_to_save:
                logging.info(f"  No data rows to write for {file_path}. Only header written.")
            else:
                writer.writerows(list_to_save)
                logging.debug(f"  Successfully wrote {len(list_to_save)} rows to {file_path}.")
        return True

    except (IOError, csv.Error, TypeError, KeyError) as e:
        logging.error(f"Error saving CSV data to {file_path}: {e}")
        return False
    except Exception as e:
        logging.exception(f"Unexpected error saving CSV {file_path}: {e}")
        return False

def save_user_data(user_id):
    """Saves all data for the specified user_id from app_data to files."""
    if not user_id:
        logging.error("Cannot save data: No user ID specified.")
        return

    logging.info(f"Saving data for user: {user_id}")
    ensure_data_dir()

    data_types_config = {
        "wallets": {"type": dict, "fields": ['wallet_id', 'name', 'balance']},
        "budgets": {"type": dict, "fields": ['budget_id', 'name', 'allocated', 'cycle']},
        "goals": {"type": dict, "fields": ['goal_id', 'name', 'target', 'saved', 'due_date']},
        "transactions": {"type": list, "fields": ['date', 'time', 'timestamp', 'title', 'wallet', 'amount', 'category', 'type', 'from_account', 'to_account', 'linked_budget', 'linked_goal']},
        "activity_log": {"type": list, "fields": ['timestamp', 'action']},
        "settings": {"type": dict, "is_json": True},
    }

    save_success = True
    for data_key, config in data_types_config.items():
        file_path = get_user_data_file_path(user_id, data_key)
        is_json = config.get("is_json", False)
        data_to_save = app_data.get(data_key)

        if data_to_save is None:
            logging.warning(f"No data found in app_data for '{data_key}'. Skipping save for {file_path}.")
            continue
        else:
            data_type_info = type(data_to_save)
            data_len_info = f" (Length: {len(data_to_save)})" if hasattr(data_to_save, '__len__') else ""
            logging.info(f"Attempting to save '{data_key}' (Type: {data_type_info}{data_len_info}) to {file_path}")

        success = False
        if is_json:
            success = _save_json_data(file_path, data_to_save)
        else: # CSV
            csv_fields = config.get("fields")
            if not csv_fields:
                 logging.error(f"Missing 'fields' configuration for CSV data key '{data_key}'. Cannot save.")
                 success = False
            else:
                 success = _save_csv_data(file_path, data_to_save, csv_fields)

        if not success:
            save_success = False
        else:
            logging.info(f"Successfully saved '{data_key}' to '{file_path}'.")

    if save_success:
        logging.info(f"Data saving finished successfully for user: {user_id}")
    else:
        logging.error(f"Data saving process encountered errors for user: {user_id}. Some data might not be saved.")

# --- Accounts Page Class (User Profile Selection) ---
class AccountsPage(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ExpenseWise - Select User Profile")
        self.geometry("800x500")
        self.configure(bg=THEME_DARK["background"])

        load_user_profiles_from_csv()
        self.selected_user_id = None

        # Styling
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('Account.TButton', background=THEME_DARK["account_icon_bg"], foreground=THEME_DARK["foreground"], font=FONT_XXLARGE, borderwidth=0, relief=tk.FLAT, width=5, height=2)
        self.style.map('Account.TButton', background=[('active', '#555555')])
        self.style.configure('AddAccount.TButton', background=THEME_DARK["background"], foreground=THEME_DARK["disabled"], font=FONT_XXLARGE, borderwidth=1, relief=tk.SOLID, bordercolor=THEME_DARK["disabled"])
        self.style.map('AddAccount.TButton', foreground=[('active', THEME_DARK["foreground"])], bordercolor=[('active', THEME_DARK["foreground"])])
        self.style.configure('Exit.TButton', font=FONT_NORMAL, foreground=THEME_DARK["foreground"], background="#555555", borderwidth=1, relief=tk.SOLID)
        self.style.map('Exit.TButton', background=[('active', THEME_DARK["red"])], foreground=[('active', THEME_DARK["button_fg"])])

        # Layout
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        title_label = tk.Label(self, text="Who's Using ExpenseWise?", font=FONT_XLARGE, bg=THEME_DARK["background"], fg=THEME_DARK["foreground"])
        title_label.grid(row=0, column=0, pady=(40, 30))
        self.accounts_frame = tk.Frame(self, bg=THEME_DARK["background"])
        self.accounts_frame.grid(row=1, column=0, sticky="n", pady=20)
        self.accounts_frame.grid_columnconfigure(0, weight=1)
        self.display_user_profiles()
        exit_button = ttk.Button(self, text="Exit Application", command=self.exit_app, style="Exit.TButton")
        exit_button.place(relx=0.98, rely=0.95, anchor='se', x=-20, y=-20)
        self.center_window()
        self.protocol("WM_DELETE_WINDOW", self.exit_app)

    def center_window(self):
        """Centers the main window on the screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        self.attributes('-alpha', 1.0)

    def display_user_profiles(self):
        """Displays user profile icons and names."""
        for widget in self.accounts_frame.winfo_children():
            widget.destroy()
        inner_frame = tk.Frame(self.accounts_frame, bg=THEME_DARK["background"])
        inner_frame.pack(pady=10)

        max_cols = 4
        col_count = 0
        row_count = 0
        profiles = app_data.get("user_profiles", {})
        sorted_profiles = sorted(profiles.items(), key=lambda item: item[1].get('name', '').lower())

        for user_id, details in sorted_profiles:
            if col_count >= max_cols:
                col_count = 0
                row_count += 1
            account_container = tk.Frame(inner_frame, bg=THEME_DARK["background"])
            account_container.grid(row=row_count, column=col_count, padx=20, pady=10, sticky="n")
            name = details.get('name', 'Unknown')
            initial = name[0].upper() if name else "?"
            icon_color = details.get('icon_color', random.choice(ACCOUNT_ICON_COLORS))
            icon_frame = tk.Frame(account_container, bg=icon_color, width=100, height=100, cursor="hand2")
            icon_frame.pack(pady=(0, 5))
            icon_frame.pack_propagate(False)
            initial_label = tk.Label(icon_frame, text=initial, font=FONT_XXLARGE, bg=icon_color, fg=THEME_DARK["button_fg"])
            initial_label.place(relx=0.5, rely=0.5, anchor="center")
            name_label = tk.Label(account_container, text=name, font=FONT_ACCOUNT_NAME, bg=THEME_DARK["background"], fg=THEME_DARK["disabled"])
            name_label.pack()
            icon_frame.bind("<Button-1>", lambda e, u_id=user_id: self.select_user(u_id))
            initial_label.bind("<Button-1>", lambda e, u_id=user_id: self.select_user(u_id))
            name_label.bind("<Button-1>", lambda e, u_id=user_id: self.select_user(u_id))
            col_count += 1

        # Add "+" Button for new profile
        if col_count >= max_cols:
            col_count = 0
            row_count += 1
        add_container = tk.Frame(inner_frame, bg=THEME_DARK["background"])
        add_container.grid(row=row_count, column=col_count, padx=20, pady=10, sticky="n")
        add_icon_frame = tk.Frame(add_container, bg=THEME_DARK["background"], width=100, height=100, cursor="hand2", relief=tk.SOLID, borderwidth=2, highlightbackground=THEME_DARK["disabled"], highlightthickness=2, highlightcolor=THEME_DARK["disabled"])
        add_icon_frame.pack(pady=(0, 5))
        add_icon_frame.pack_propagate(False)
        add_label = tk.Label(add_icon_frame, text="+", font=FONT_XXLARGE, bg=THEME_DARK["background"], fg=THEME_DARK["disabled"])
        add_label.place(relx=0.5, rely=0.5, anchor="center")
        add_name_label = tk.Label(add_container, text="Add Profile", font=FONT_ACCOUNT_NAME, bg=THEME_DARK["background"], fg=THEME_DARK["disabled"])
        add_name_label.pack()
        add_icon_frame.bind("<Button-1>", self.add_user_profile_dialog)
        add_label.bind("<Button-1>", self.add_user_profile_dialog)
        add_name_label.bind("<Button-1>", self.add_user_profile_dialog)

    def add_user_profile_dialog(self, event=None):
        """Opens a dialog to add a new user profile."""
        dialog = SimpleEntryDialog(self, "Create New User Profile", {
            "name": {"label": "Profile Name:", "type": "text", "required": True},
        })
        result = dialog.result
        if result and isinstance(result, dict):
            try:
                name = result.get("name", "").strip()
                if not name: raise ValueError("Profile name cannot be empty.")
                current_profiles = app_data.get("user_profiles", {})
                if any(prof.get('name', '').lower() == name.lower() for prof in current_profiles.values() if isinstance(prof, dict)):
                    raise ValueError(f"A profile named '{name}' already exists.")
                new_id = get_unique_id("user")
                new_profile = {"name": name, "icon_color": random.choice(ACCOUNT_ICON_COLORS)}
                app_data["user_profiles"][new_id] = new_profile
                save_user_profiles_to_csv()
                self.display_user_profiles()
            except ValueError as e: messagebox.showerror("Invalid Input", str(e), parent=self)
            except Exception as e:
                logging.exception("Error creating profile")
                messagebox.showerror("Error", f"Could not create profile.\n{e}", parent=self)

    def select_user(self, user_id):
        """Selects a user profile and closes the accounts page."""
        profiles = app_data.get("user_profiles", {})
        if user_id in profiles:
            self.selected_user_id = user_id
            user_name = profiles[user_id].get('name', user_id)
            logging.info(f"Selected user profile: {user_name} (ID: {user_id})")
            self.destroy()
        else:
            logging.error(f"Attempted to select non-existent user ID: {user_id}")
            messagebox.showerror("Error", "Selected user profile not found.", parent=self)
            load_user_profiles_from_csv()
            self.display_user_profiles()

    def exit_app(self):
        """Exits the application from the Accounts Page."""
        logging.info("Exiting ExpenseWise from Accounts Page.")
        self.quit()
        self.destroy()


# --- Main Application Class (ExpenseWiseApp) ---
class ExpenseWiseApp(tk.Tk):
    def __init__(self, user_id):
        super().__init__()
        self.current_user_id = user_id
        self._page_creation_lock = False

        load_user_data(self.current_user_id)
        self.current_theme = app_data.get("settings", {}).get("theme", "dark")
        self.apply_theme_colors()

        self.title(f"ExpenseWise - {app_data['user_profiles'].get(user_id, {}).get('name', 'User')}")
        self.geometry("1200x700")
        self.configure(bg=theme_colors["background"])

        # Style Configuration
        self.style = ttk.Style(self)
        self.configure_styles()

        # Layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Sidebar
        self.sidebar = Sidebar(self, self.show_page)
        self.sidebar.grid(row=0, column=0, sticky="nsw")

        # Main Content Area
        self.main_frame = tk.Frame(self, bg=theme_colors["background"])
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.current_page_frame = None

        # Floating Action Button (FAB)
        self.fab = create_stylish_button(self, "+", self.open_add_transaction_dialog, style="FAB.TButton")
        self.fab.place(relx=0.98, rely=0.95, anchor='se')

        # Initial Page
        self.show_page("Home")
        if self.sidebar and self.sidebar.winfo_exists():
            self.sidebar.highlight_button("Home")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        logging.info(f"ExpenseWiseApp initialized for user {user_id}.")

    def on_closing(self):
        """Handles application closing, saving user data."""
        logging.info(f"Closing application for user {self.current_user_id}...")
        save_user_data(self.current_user_id)
        logging.info("User data saved. Exiting main application window.")
        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
            self.sidebar.stop_timer()
        self.destroy()

    def perform_full_exit(self):
        """Handles saving data and completely exiting the application."""
        logging.info(f"Performing full application exit for user {self.current_user_id}...")
        save_user_data(self.current_user_id)
        logging.info("User data saved.")
        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
            try:
                self.sidebar.stop_timer()
            except Exception as e:
                logging.warning(f"Ignoring error stopping sidebar timer during exit: {e}")

        self._full_exit_requested = True # Flag for the main loop
        try:
            self.destroy()
        except tk.TclError as e:
            logging.warning(f"TclError during full exit: {e}")
        logging.info("Full exit procedures initiated.")

    def apply_theme_colors(self):
        """Applies chosen theme colors to global variable."""
        global theme_colors
        if self.current_theme == "light":
            theme_colors.update(THEME_LIGHT)
            logging.info("Applied Light Theme")
        else:
            theme_colors.update(THEME_DARK)
            logging.info("Applied Dark Theme")

    def configure_styles(self):
        """Configures and updates all ttk widget styles based on current theme."""
        self.style.theme_use('clam')
        # General Widget Styling
        self.style.configure('.', background=theme_colors["background"], foreground=theme_colors["foreground"], font=FONT_NORMAL)
        self.style.configure('TFrame', background=theme_colors["background"])
        self.style.configure('TLabel', background=theme_colors["background"], foreground=theme_colors["foreground"], font=FONT_NORMAL)
        # Specific Label Styles
        self.style.configure('Sidebar.TLabel', background=theme_colors["sidebar"], foreground=theme_colors["foreground"])
        self.style.configure('Card.TLabel', background=theme_colors["card"], foreground=theme_colors["foreground"])
        self.style.configure('Title.TLabel', background=theme_colors["background"], foreground=theme_colors["foreground"], font=FONT_LARGE)
        self.style.configure('CardTitle.TLabel', background=theme_colors["card"], foreground=theme_colors["foreground"], font=FONT_BOLD)
        self.style.configure('Accent.TLabel', background=theme_colors["card"], foreground=theme_colors["accent"], font=FONT_BOLD)
        self.style.configure('Error.TLabel', background=theme_colors["background"], foreground=theme_colors["red"], font=FONT_BOLD)
        date_time_fg = theme_colors["disabled"] if self.current_theme == 'dark' else theme_colors["accent_darker"]
        self.style.configure('DateTime.TLabel', background=theme_colors["sidebar"], foreground=date_time_fg, font=FONT_SMALL)
        # Button Styling
        self.style.configure('TButton', background=theme_colors["accent"], foreground=theme_colors["button_fg"], font=FONT_BOLD, padding=6, borderwidth=0, relief=tk.FLAT)
        self.style.map('TButton', background=[('active', theme_colors["accent_darker"])], foreground=[('active', theme_colors["button_fg"])])
        # Sidebar button styling
        self.style.configure('Sidebar.TButton', background=theme_colors["sidebar"], foreground=theme_colors["foreground"], font=FONT_BOLD, anchor='w', padding=(15, 8), borderwidth=0, relief=tk.FLAT)
        self.style.map('Sidebar.TButton', background=[('active', theme_colors["accent"]), ('selected', theme_colors["accent"])], foreground=[('active', theme_colors["button_fg"]), ('selected', theme_colors["button_fg"])])
        # FAB styling
        self.style.configure('FAB.TButton', background=theme_colors["accent"], foreground=theme_colors["button_fg"], font=(FONT_FAMILY, 18, "bold"), padding=10, borderwidth=0, relief=tk.FLAT)
        self.style.map('FAB.TButton', background=[('active', theme_colors["accent_darker"])])
        # Treeview Styling
        self.style.configure("Treeview", background=theme_colors["card"], foreground=theme_colors["foreground"], fieldbackground=theme_colors["card"], rowheight=25, borderwidth=0, relief=tk.FLAT)
        self.style.configure("Treeview.Heading", background=theme_colors["treeview_heading_bg"], foreground=theme_colors["foreground"], font=FONT_BOLD, relief="flat", padding=(5, 5))
        self.style.map("Treeview.Heading", background=[('active', theme_colors["accent"])])
        self.style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        # Progressbar Styling
        self.style.configure("TProgressbar", thickness=10, background=theme_colors["accent"], troughcolor=theme_colors["card"])
        # Combobox Styling
        self.option_add('*TCombobox*Listbox*Background', theme_colors["combobox_list_bg"])
        self.option_add('*TCombobox*Listbox*Foreground', theme_colors["combobox_list_fg"])
        self.option_add('*TCombobox*Listbox*selectBackground', theme_colors["combobox_list_select_bg"])
        self.option_add('*TCombobox*Listbox*selectForeground', theme_colors["combobox_list_select_fg"])
        self.style.configure('TCombobox', background=theme_colors["card"], foreground=theme_colors["foreground"], fieldbackground=theme_colors["card"], selectbackground=theme_colors["card"], selectforeground=theme_colors["foreground"], arrowcolor=theme_colors["foreground"], borderwidth=0, padding=5)
        self.style.map('TCombobox', fieldbackground=[('readonly', theme_colors["card"])], selectbackground=[('readonly', theme_colors["card"])], selectforeground=[('readonly', theme_colors["foreground"])])
        # Entry Styling
        self.style.configure('TEntry', background=theme_colors["card"], foreground=theme_colors["foreground"], fieldbackground=theme_colors["card"], insertcolor=theme_colors["foreground"], borderwidth=0, padding=5)
        self.style.map('TEntry', fieldbackground=[('focus', theme_colors["card"])])
        # Notebook Styling
        self.style.configure('TNotebook', background=theme_colors["background"], borderwidth=0)
        self.style.configure('TNotebook.Tab', font=FONT_BOLD, padding=[10, 5], background=theme_colors["card"], foreground=theme_colors["foreground"], borderwidth=0)
        self.style.map('TNotebook.Tab', background=[('selected', theme_colors["accent"])], foreground=[('selected', theme_colors["button_fg"])])
        # Scrollbar Styling
        self.style.configure("Vertical.TScrollbar", background=theme_colors["scrollbar_bg"], troughcolor=theme_colors["scrollbar_trough"], borderwidth=0, arrowcolor=theme_colors["foreground"], relief=tk.FLAT)
        self.style.map("Vertical.TScrollbar", background=[('active', theme_colors["accent"])])
        # Checkbutton/Radiobutton Styling
        self.style.configure("TCheckbutton", background=theme_colors["background"], foreground=theme_colors["foreground"], font=FONT_NORMAL)
        self.style.map("TCheckbutton", indicatorcolor=[('selected', theme_colors["accent"])])
        self.style.configure("TRadiobutton", background=theme_colors["background"], foreground=theme_colors["foreground"], font=FONT_NORMAL)
        self.style.map("TRadiobutton", indicatorcolor=[('selected', theme_colors["accent"])])
        # Card Radiobutton Style
        self.style.configure("Card.TRadiobutton", background=theme_colors["card"], foreground=theme_colors["foreground"])
        self.style.map("Card.TRadiobutton", background=[('active', theme_colors["card"])], indicatorcolor=[('selected', theme_colors["accent"])])
        # Toolbutton style (for category Radiobuttons)
        self.style.configure("Toolbutton", anchor="center", padding=5, font=FONT_NORMAL, background=theme_colors["card"], foreground=theme_colors["foreground"], borderwidth=1, relief="raised")
        self.style.map("Toolbutton", relief=[('selected', 'sunken'), ('active', 'raised')], background=[('selected', theme_colors["accent"]), ('active', theme_colors["card"])], foreground=[('selected', theme_colors["button_fg"]), ('active', theme_colors["foreground"])])

        # Update non-ttk widgets
        self.configure(bg=theme_colors["background"])
        if hasattr(self, 'main_frame') and self.main_frame and self.main_frame.winfo_exists():
            self.main_frame.configure(bg=theme_colors["background"])
        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
            self.sidebar.configure(bg=theme_colors["sidebar"])
        if hasattr(self, 'fab') and self.fab and self.fab.winfo_exists():
             self.fab.configure(style="FAB.TButton")
        logging.debug("Styles reconfigured for theme: %s", self.current_theme)

    def switch_theme(self, theme_name):
        """Switches the application theme and refreshes UI."""
        if theme_name == self.current_theme: return
        if self._page_creation_lock: return

        logging.info(f"Switching theme to: {theme_name}")
        self.current_theme = theme_name
        app_data["settings"]["theme"] = theme_name
        self.apply_theme_colors()

        try:
            self.configure_styles()

            current_page_name = "Home"
            if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
                current_page_name = self.sidebar.get_current_page_name() or "Home"
                self.sidebar.stop_timer()
                self.sidebar.destroy()

            self.sidebar = Sidebar(self, self.show_page)
            self.sidebar.grid(row=0, column=0, sticky="nsw")

            self._page_creation_lock = True
            try:
                 self.show_page(current_page_name)
            finally:
                self._page_creation_lock = False

            if hasattr(self, 'fab') and self.fab and self.fab.winfo_exists():
                self.fab.configure(style="FAB.TButton")

            log_activity(f"Theme switched to {theme_name}")
            logging.info(f"Theme switched successfully to {theme_name}")

        except tk.TclError as e:
             logging.error(f"TclError during theme switch: {e}")
             messagebox.showwarning("Theme Switch Issue", "An minor error occurred applying the theme.", parent=self)
             try:
                 self._page_creation_lock = True
                 self.show_page("Home")
             except Exception as nested_e:
                  logging.error(f"Failed to recover after theme switch error: {nested_e}")
             finally:
                  self._page_creation_lock = False
        except Exception as e:
             logging.exception("Unexpected error during theme switch")
             messagebox.showerror("Theme Switch Error", f"An unexpected error occurred during theme switch:\n{e}", parent=self)

    def show_page(self, page_name):
        """Displays the specified page in the main content area."""
        if self._page_creation_lock:
             logging.warning(f"show_page('{page_name}') called while lock is active. Ignoring.")
             return

        logging.info(f"Switching to page: {page_name}")

        # Destroy current page frame
        if self.current_page_frame and self.current_page_frame.winfo_exists():
            if isinstance(self.current_page_frame, HomePage):
                self.current_page_frame.unbind_mousewheel()
            try:
                self.current_page_frame.destroy()
                self.current_page_frame = None
            except tk.TclError as e:
                 logging.warning(f"TclError destroying previous page frame: {e}")

        # Map page names to their classes
        page_mapping = {
            "Home": HomePage,
            "Transactions": TransactionsPage,
            "Budgets": BudgetPage,
            "Goals": GoalsPage,
            "Wallets": WalletsPage,
            "All Spending": AllSpendingPage,
            "Activity Log": ActivityLogPage,
            "Settings": SettingsPage
        }

        page_class = page_mapping.get(page_name)
        page = None

        if page_class:
            try:
                 page = page_class(self.main_frame, self)
                 self.current_page_frame = page
                 page.grid(row=0, column=0, sticky="nsew")
            except Exception as e:
                logging.exception(f"Error creating page '{page_name}'")
                messagebox.showerror("Page Load Error", f"Could not load page '{page_name}':\n{e}", parent=self)
                try:
                     page = PlaceholderPage(self.main_frame, f"Error Loading {page_name}", self)
                     self.current_page_frame = page
                     page.grid(row=0, column=0, sticky="nsew")
                except: pass
        else:
            logging.warning(f"Page class not found for '{page_name}'. Showing placeholder.")
            page = PlaceholderPage(self.main_frame, page_name, self)
            self.current_page_frame = page
            page.grid(row=0, column=0, sticky="nsew")

        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
             self.sidebar.highlight_button(page_name)
             self.sidebar.current_page_name = page_name

    def open_add_transaction_dialog(self):
        """Opens the Add Transaction dialog."""
        dialog = AddTransactionDialog(self)

    def refresh_current_page(self):
        """Refreshes the content of the currently displayed page."""
        logging.debug("Refreshing current page...")
        current_page = "Home"
        if hasattr(self, 'sidebar') and self.sidebar and self.sidebar.winfo_exists():
             current_page = self.sidebar.get_current_page_name() or "Home"

        if current_page:
             self.show_page(current_page)
        else:
             logging.warning("Could not determine current page to refresh, defaulting to Home.")
             self.show_page("Home")

# --- Sidebar Class ---
class Sidebar(tk.Frame):
    def __init__(self, parent, show_page_callback):
        super().__init__(parent, bg=theme_colors["sidebar"], width=200)
        self.app = parent
        self.show_page = show_page_callback
        self.buttons = {}
        self.current_page_name = None
        self._timer_id = None

        sidebar_items_config = [
            {"name": "Home", "type": "page"},
            {"name": "Transactions", "type": "page"},
            {"name": "Budgets", "type": "page"},
            {"name": "Goals", "type": "page"},
            {"name": "Wallets", "type": "page"},
            {"name": "All Spending", "type": "page"},
            {"name": "Activity Log", "type": "page"},
            {"name": "Settings", "type": "page"},
        ]

        title_label = ttk.Label(self, text="ExpenseWise", font=FONT_XLARGE, style="Sidebar.TLabel", anchor='w', padding=(15, 15))
        title_label.pack(pady=(10, 0), fill="x")

        self.datetime_label = ttk.Label(self, text="", style="DateTime.TLabel", anchor='w', padding=(15, 0))
        self.datetime_label.pack(pady=(0, 15), fill="x")
        self.update_datetime()

        # Navigation Buttons
        for item_config in sidebar_items_config:
            item_name = item_config["name"]
            btn = create_stylish_button(self, item_name,
                                        lambda name=item_name: self.show_page(name),
                                        style="Sidebar.TButton")
            btn.pack(fill="x")
            self.buttons[item_name] = btn

        tk.Frame(self, bg=theme_colors["sidebar"]).pack(expand=True, fill="y")

    def update_datetime(self):
        """Updates the current date and time displayed in the sidebar."""
        now = datetime.datetime.now()
        datetime_str = now.strftime("%a, %d %b %Y | %H:%M:%S")
        try:
            if self.datetime_label.winfo_exists():
                self.datetime_label.configure(text=datetime_str)
                self._timer_id = self.after(1000, self.update_datetime)
            else:
                self._timer_id = None
        except tk.TclError:
            self._timer_id = None

    def stop_timer(self):
        """Stops the datetime update timer."""
        if self._timer_id:
            try: self.after_cancel(self._timer_id)
            except tk.TclError: pass
            finally: self._timer_id = None

    def highlight_button(self, page_name):
        """Highlights the selected sidebar button."""
        if not page_name: return
        self.current_page_name = page_name

        for name, button in self.buttons.items():
             if not button or not button.winfo_exists(): continue
             try:
                 if name == page_name:
                     button.state(['selected'])
                 else:
                     button.state(['!selected'])
             except tk.TclError:
                 logging.warning(f"TclError setting state for button '{name}'.")

    def get_current_page_name(self):
        """Returns the name of the currently selected page."""
        return self.current_page_name

    def destroy(self):
        """Destroys the sidebar and stops its timer."""
        self.stop_timer()
        super().destroy()

    def __del__(self):
        """Ensures timer is stopped on object deletion."""
        self.stop_timer()


# --- Base Page Class ---
class BasePage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=theme_colors["background"])
        self.app = app


# --- Placeholder Page Class ---
class PlaceholderPage(BasePage):
    def __init__(self, parent, page_name, app):
        super().__init__(parent, app)
        label = ttk.Label(self, text=f"{page_name}\n(Content Placeholder)", style="Title.TLabel", justify=tk.CENTER, anchor=tk.CENTER)
        label.pack(expand=True, fill="both", padx=20, pady=20)

# --- HomePage Class ---
class HomePage(BasePage):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.max_wallets_per_row = 4
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Scrollable Setup
        self.canvas = tk.Canvas(self, bg=theme_colors["background"], highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        self.scrollable_frame = tk.Frame(self.canvas, bg=theme_colors["background"])
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", tags="scrollable_frame")
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.bind_mousewheel()

        # Content inside scrollable_frame
        content_frame = self.scrollable_frame
        content_frame.grid_columnconfigure(0, weight=1)
        self.create_wallets_section(content_frame, row=0)
        self.create_budgets_section(content_frame, row=1)
        self.create_goals_section(content_frame, row=2)
        self.update_idletasks()
        # Configure scrollregion after content is placed
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def _on_frame_configure(self, event=None):
        """Updates scroll region when the inner frame's size changes."""
        if self.canvas.winfo_exists():
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
         """Updates the width of the inner frame to match the canvas width."""
         if self.canvas.winfo_exists():
             canvas_width = self.canvas.winfo_width()
             self.canvas.itemconfig("scrollable_frame", width=canvas_width)

    def bind_mousewheel(self):
        """Binds mouse wheel scrolling events."""
        self.canvas.bind_all("<MouseWheel>", lambda e: self._on_mousewheel(e), add='+')
        self.canvas.bind_all("<Button-4>", lambda e: self._on_mousewheel(e), add='+')
        self.canvas.bind_all("<Button-5>", lambda e: self._on_mousewheel(e), add='+')

    def unbind_mousewheel(self):
        """Unbinds all mouse wheel events."""
        try:
             self.canvas.unbind_all("<MouseWheel>")
             self.canvas.unbind_all("<Button-4>")
             self.canvas.unbind_all("<Button-5>")
        except tk.TclError as e:
             logging.warning(f"TclError unbinding mousewheel: {e}")

    def _on_mousewheel(self, event):
        """Handles mouse wheel scrolling for the canvas."""
        widget_under_mouse = self.winfo_containing(event.x_root, event.y_root)
        if not widget_under_mouse: return

        # Check if the widget under the mouse is canvas or its descendant
        is_related = False
        curr = widget_under_mouse
        while curr is not None:
            if curr == self.canvas:
                is_related = True
                break
            try:
                if hasattr(curr, 'master'):
                    curr = curr.master
                else:
                    break
            except AttributeError:
                break

        if is_related:
            delta = 0
            if event.num == 4:
                delta = -1
            elif event.num == 5:
                delta = 1
            elif event.delta > 0:
                delta = -1
            elif event.delta < 0:
                delta = 1

            if delta != 0 and self.canvas.winfo_exists():
                self.canvas.yview_scroll(delta, "units")

    def create_wallets_section(self, parent_frame, row):
        """Creates and populates the Wallets summary section."""
        wallets_frame = tk.Frame(parent_frame, bg=theme_colors["background"])
        wallets_frame.grid(row=row, column=0, sticky="ew", pady=(0, 20))
        ttk.Label(wallets_frame, text="Wallets", style="Title.TLabel").pack(anchor="w", pady=(0, 10))
        wallets_grid_frame = tk.Frame(wallets_frame, bg=theme_colors["background"])
        wallets_grid_frame.pack(fill="x")
        for i in range(self.max_wallets_per_row):
            wallets_grid_frame.grid_columnconfigure(i, weight=1, uniform="wallet_col")
        user_wallets = app_data.get("wallets", {})
        if not user_wallets:
            ttk.Label(wallets_grid_frame, text="No wallets created yet.", style="Card.TLabel", foreground=theme_colors["disabled"]).grid(row=0, column=0, columnspan=self.max_wallets_per_row, padx=10, pady=5, sticky="w")
        else:
            sorted_wallets = sorted(user_wallets.items(), key=lambda item: str(item[1].get('name', item[0])).lower() if isinstance(item[1], dict) else "")
            grid_row, grid_col = 0, 0
            for wallet_id, details in sorted_wallets:
                if not isinstance(details, dict): continue
                card = create_card_frame(wallets_grid_frame)
                card.grid(row=grid_row, column=grid_col, sticky="nsew", padx=10, pady=5)
                card.grid_columnconfigure(0, weight=1)
                ttk.Label(card, text=details.get("name", "Unnamed"), style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
                ttk.Label(card, text=format_currency(details.get('balance', 0.0)), style="Card.TLabel", font=FONT_LARGE).grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
                grid_col += 1
                if grid_col >= self.max_wallets_per_row: grid_col = 0; grid_row += 1

    def create_budgets_section(self, parent_frame, row):
        """Creates and populates the Budgets summary section."""
        budget_frame = tk.Frame(parent_frame, bg=theme_colors["background"])
        budget_frame.grid(row=row, column=0, sticky="ew", pady=20)
        ttk.Label(budget_frame, text="Budgets", style="Title.TLabel").pack(anchor="w", pady=(0, 10))
        budget_grid_frame = tk.Frame(budget_frame, bg=theme_colors["background"])
        budget_grid_frame.pack(fill="x")
        max_cols = 3
        for i in range(max_cols):
            budget_grid_frame.grid_columnconfigure(i, weight=1, uniform="budget_col")
        user_budgets = app_data.get("budgets", {})
        if not user_budgets:
            ttk.Label(budget_grid_frame, text="No budgets created yet. Go to 'Budgets'.", style="Card.TLabel", foreground=theme_colors["disabled"]).grid(row=0, column=0, columnspan=max_cols, padx=10, pady=5, sticky="w")
        else:
            sorted_budgets = sorted(user_budgets.items(), key=lambda item: str(item[1].get('name', item[0])).lower() if isinstance(item[1], dict) else "")
            grid_row, grid_col = 0, 0
            for bud_id, details in sorted_budgets:
                if not isinstance(details, dict): continue
                card = create_card_frame(budget_grid_frame)
                card.grid(row=grid_row, column=grid_col, sticky="nsew", padx=10, pady=5)
                card.grid_columnconfigure(0, weight=1)
                cycle_text = f" ({details.get('cycle', 'N/A')})"
                budget_name = details.get("name", "Unnamed") + cycle_text
                ttk.Label(card, text=budget_name, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
                allocated = details.get('allocated', 0.0)
                spent = self.calculate_budget_spent(details.get("name"))
                spent_text = f"{format_currency(spent)} / {format_currency(allocated)}"
                ttk.Label(card, text=spent_text, style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=10)
                progress = (spent / allocated) * 100 if allocated and allocated > 0 else 0
                pb = ttk.Progressbar(card, orient="horizontal", length=150, mode="determinate", value=min(progress, 100), style="TProgressbar")
                pb.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 10))
                grid_col += 1
                if grid_col >= max_cols: grid_col = 0; grid_row += 1

    def create_goals_section(self, parent_frame, row):
        """Creates and populates the Goals summary section, including linked expenses."""
        goals_frame = tk.Frame(parent_frame, bg=theme_colors["background"])
        goals_frame.grid(row=row, column=0, sticky="ew", pady=20)
        ttk.Label(goals_frame, text="Goals", style="Title.TLabel").pack(anchor="w", pady=(0, 10))
        goals_grid_frame = tk.Frame(goals_frame, bg=theme_colors["background"])
        goals_grid_frame.pack(fill="x")
        max_cols = 3
        for i in range(max_cols):
            goals_grid_frame.grid_columnconfigure(i, weight=1, uniform="goal_col")

        user_goals = app_data.get("goals", {})
        user_transactions = app_data.get("transactions", [])
        if not isinstance(user_transactions, list): user_transactions = []
        if not isinstance(user_goals, dict): user_goals = {}

        if not user_goals:
            ttk.Label(goals_grid_frame, text="No goals set yet. Go to 'Goals'.", style="Card.TLabel",
                      foreground=theme_colors["disabled"]).grid(row=0, column=0, columnspan=max_cols, padx=10,
                                                                pady=5, sticky="w")
        else:
            def goal_sort_key(item):
                # Sort by due date, then alphabetically
                details = item[1]
                due_date_obj = datetime.date.max
                name_str = ""
                if isinstance(details, dict):
                    due_date_str = details.get("due_date")
                    name_str = details.get("name", "").lower()
                    if due_date_str:
                        try:
                            due_date_obj = datetime.datetime.strptime(due_date_str, "%Y-%m-%d").date()
                        except (ValueError, TypeError):
                            pass
                return (due_date_obj, name_str)

            sorted_goals = sorted(user_goals.items(), key=goal_sort_key)

            grid_row, grid_col = 0, 0
            for goal_id, details in sorted_goals:
                if not isinstance(details, dict): continue

                goal_name = details.get("name", "Unnamed Goal")
                target = details.get('target', 0.0)
                base_saved = details.get('saved', 0.0)

                # Calculate linked contribution from expenses
                linked_expense_contribution = 0.0
                for tx in user_transactions:
                    if (isinstance(tx, dict) and
                            tx.get("linked_goal") == goal_name and
                            tx.get("type") == "expense" and
                            isinstance(tx.get("amount"), (int, float)) and
                            tx["amount"] < 0):
                        linked_expense_contribution += abs(tx["amount"])

                # Effective saved amount includes base + linked expenses
                effective_saved = base_saved + linked_expense_contribution

                card = create_card_frame(goals_grid_frame)
                card.grid(row=grid_row, column=grid_col, sticky="nsew", padx=10, pady=5)
                card.grid_columnconfigure(0, weight=1)

                ttk.Label(card, text=goal_name, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w",
                                                                               padx=10, pady=(10, 5))
                saved_text = f"{format_currency(effective_saved)} / {format_currency(target)}"
                ttk.Label(card, text=saved_text, style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=10)

                progress = (effective_saved / target) * 100 if target and target > 0 else 0
                visual_progress = min(progress, 100.0)
                pb = ttk.Progressbar(card, orient="horizontal", length=200, mode="determinate",
                                     value=visual_progress, style="TProgressbar")
                pb.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

                # Due date logic
                due_date_str = details.get("due_date")
                remaining_text = "No Due Date"
                if due_date_str:
                    try:
                        due_date = datetime.datetime.strptime(due_date_str, "%Y-%m-%d").date()
                        today = datetime.date.today()
                        delta = due_date - today
                        if delta.days == 0:
                            remaining_text = f"Due Today ({due_date_str})"
                        elif delta.days > 0:
                            remaining_text = f"{delta.days} days left (Due: {due_date_str})"
                        else:
                            remaining_text = f"Overdue by {abs(delta.days)} days (Due: {due_date_str})"
                    except (ValueError, TypeError):
                        remaining_text = f"Invalid Due Date ({due_date_str})"

                ttk.Label(card, text=remaining_text, style="Card.TLabel",
                          foreground=theme_colors["disabled"]).grid(row=3, column=0, sticky="w", padx=10,
                                                                    pady=(0, 10))

                grid_col += 1
                if grid_col >= max_cols: grid_col = 0; grid_row += 1

    def calculate_budget_spent(self, budget_name):
        """Calculates total spending linked to a specific budget."""
        if not budget_name: return 0.0
        total_spent = 0.0
        user_transactions = app_data.get("transactions", [])
        if not isinstance(user_transactions, list): return 0.0

        for tx in user_transactions:
            # Check if it's an expense transaction linked to the budget
            if (isinstance(tx, dict) and
                    tx.get("type") == "expense" and
                    tx.get("linked_budget") == budget_name and
                    isinstance(tx.get("amount"), (int, float))):
                total_spent += abs(tx["amount"])
        return total_spent

    def destroy(self):
        """Destroys the HomePage instance and unbinds mousewheel events."""
        self.unbind_mousewheel()
        super().destroy()

# --- TransactionsPage Class ---
class TransactionsPage(BasePage):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        control_frame = tk.Frame(self, bg=theme_colors["background"])
        control_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Label(control_frame, text="Transactions", style="Title.TLabel").pack(side=tk.LEFT, padx=(0, 20))

        columns = ("date", "title", "wallet", "category", "amount")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", style="Treeview")
        self.tree.heading("date", text="Date")
        self.tree.heading("title", text="Title")
        self.tree.heading("wallet", text="Wallet")
        self.tree.heading("category", text="Category")
        self.tree.heading("amount", text="Amount")
        self.tree.column("date", width=100, anchor=tk.W, stretch=tk.NO)
        self.tree.column("title", width=250, anchor=tk.W)
        self.tree.column("wallet", width=120, anchor=tk.W, stretch=tk.YES)
        self.tree.column("category", width=120, anchor=tk.W, stretch=tk.YES)
        self.tree.column("amount", width=100, anchor=tk.E, stretch=tk.NO)

        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview, style="Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")

        self.tree.tag_configure('expense', foreground=theme_colors["red"])
        self.tree.tag_configure('income', foreground=theme_colors["accent"])
        self.tree.tag_configure('transfer', foreground=theme_colors["blue"])
        self.tree.tag_configure('other', foreground=theme_colors["foreground"])

        self.populate_transactions()

    def populate_transactions(self):
        """Populates the transaction treeview with user data."""
        try:
             for item in self.tree.get_children(): self.tree.delete(item)
        except tk.TclError as e: logging.warning(f"TclError clearing transaction tree: {e}")

        user_transactions = app_data.get("transactions", [])
        if not isinstance(user_transactions, list): user_transactions = []

        # Sort transactions by timestamp (most recent first)
        def sort_key(tx):
             if not isinstance(tx, dict): return datetime.datetime.min
             ts_str = tx.get('timestamp'); date_str = tx.get('date')
             try:
                 if ts_str:
                     try: return datetime.datetime.strptime(ts_str, "%Y-%m-%d %H:%M")
                     except ValueError: return datetime.datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
                 elif date_str: return datetime.datetime.strptime(date_str, "%Y-%m-%d")
                 else: return datetime.datetime.min
             except (ValueError, TypeError): return datetime.datetime.min
        try:
             valid_transactions = [tx for tx in user_transactions if isinstance(tx, dict)]
             sorted_transactions = sorted(valid_transactions, key=sort_key, reverse=True)
        except Exception as e:
             logging.exception(f"Error sorting transactions: {e}. Displaying unsorted.")
             sorted_transactions = valid_transactions

        for tx in sorted_transactions:
            try:
                amount = tx.get('amount', 0.0)
                amount_str = format_currency(amount)
                category_name = tx.get("category", "Uncategorized")
                wallet_name = tx.get("wallet", "N/A")
                tx_type = tx.get("type", "").lower()
                tag = 'other'
                if tx_type == "income" or tx_type == "transfer_in": tag = 'income'
                elif tx_type == "expense": tag = 'expense'
                elif tx_type == "transfer_out": tag = 'transfer'
                elif amount > 0: tag = 'income'
                elif amount < 0: tag = 'expense'
                values = (tx.get('date', 'N/A'), tx.get('title', 'N/A'), wallet_name, category_name, amount_str,)
                self.tree.insert("", tk.END, values=values, tags=(tag,))
            except Exception as e:
                 logging.error(f"Error inserting transaction row for '{tx.get('title', 'N/A')}': {e}")
                 try: self.tree.insert("", tk.END, values=("Error", "Error processing row", "", "", ""), tags=('expense',))
                 except: pass

# --- ActivityLogPage Class ---
class ActivityLogPage(BasePage):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        ttk.Label(self, text="Activity Log", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        columns = ("timestamp", "action")
        tree = ttk.Treeview(self, columns=columns, show="headings", style="Treeview")
        tree.heading("timestamp", text="Timestamp")
        tree.heading("action", text="Action")
        tree.column("timestamp", width=150, anchor=tk.W, stretch=tk.NO)
        tree.column("action", width=500, anchor=tk.W)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=tree.yview, style="Vertical.TScrollbar")
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")
        activity_log = app_data.get("activity_log", [])

        if not isinstance(activity_log, list): activity_log = []
        try:
             for item in tree.get_children(): tree.delete(item)
             for log_entry in reversed(activity_log):
                 if isinstance(log_entry, dict):
                     tree.insert("", tk.END, values=(log_entry.get('timestamp', 'N/A'), log_entry.get('action', 'N/A')))
                 else: logging.warning(f"Skipping invalid activity log entry: {log_entry}")
        except tk.TclError as e: logging.warning(f"TclError populating activity log: {e}")
        except Exception as e: logging.exception("Error populating activity log")


# --- AllSpendingPage Class ---
class AllSpendingPage(BasePage):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        ttk.Label(self, text="Spending Summary", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 15))

        # Summary Cards Frame
        summary_frame = tk.Frame(self, bg=theme_colors["background"])
        summary_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        summary_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="summary_col")

        # Calculate Summaries
        total_income, total_expense = 0.0, 0.0
        expense_by_category = {}
        transactions = app_data.get("transactions", [])
        if not isinstance(transactions, list): transactions = []
        for tx in transactions:
            if not isinstance(tx, dict): continue
            tx_type = tx.get('type', '').lower(); amount = tx.get('amount'); category = tx.get('category', 'Uncategorized')
            if not isinstance(amount, (int, float)): continue
            if tx_type.startswith("transfer"): continue
            is_income = (tx_type == "income") or (tx_type != "expense" and amount > 0)
            is_expense = (tx_type == "expense") or (tx_type != "income" and amount < 0)
            if is_income: total_income += amount
            elif is_expense:
                expense_amount = abs(amount)
                total_expense += expense_amount
                expense_by_category[category] = expense_by_category.get(category, 0.0) + expense_amount
        net_total = total_income - total_expense

        # Display Summary Cards
        card_net = create_card_frame(summary_frame); card_net.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        ttk.Label(card_net, text="Net Income", style="CardTitle.TLabel").pack(padx=10, pady=(10, 0), anchor='w')
        net_color = theme_colors["accent"] if net_total >= 0 else theme_colors["red"]
        ttk.Label(card_net, text=format_currency(net_total), style="Card.TLabel", font=FONT_LARGE, foreground=net_color).pack(padx=10, pady=(0, 10), anchor='w')
        card_income = create_card_frame(summary_frame); card_income.grid(row=0, column=1, sticky="nsew", padx=10, pady=5)
        ttk.Label(card_income, text="Total Income", style="CardTitle.TLabel").pack(padx=10, pady=(10, 0), anchor='w')
        ttk.Label(card_income, text=format_currency(total_income), style="Card.TLabel", font=FONT_LARGE, foreground=theme_colors["accent"]).pack(padx=10, pady=(0, 10), anchor='w')
        card_expense = create_card_frame(summary_frame); card_expense.grid(row=0, column=2, sticky="nsew", padx=10, pady=5)
        ttk.Label(card_expense, text="Total Expenses", style="CardTitle.TLabel").pack(padx=10, pady=(10, 0), anchor='w')
        ttk.Label(card_expense, text=format_currency(total_expense), style="Card.TLabel", font=FONT_LARGE, foreground=theme_colors["red"]).pack(padx=10, pady=(0, 10), anchor='w')

        # Expense Breakdown Section
        ttk.Label(self, text="Expense Breakdown by Category", style="Title.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 10))
        breakdown_frame = tk.Frame(self, bg=theme_colors["background"])
        breakdown_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 0))
        breakdown_frame.grid_columnconfigure(0, weight=1); breakdown_frame.grid_rowconfigure(0, weight=1)
        breakdown_cols = ("category", "amount", "percentage")
        self.breakdown_tree = ttk.Treeview(breakdown_frame, columns=breakdown_cols, show="headings", style="Treeview")
        self.breakdown_tree.heading("category", text="Category"); self.breakdown_tree.heading("amount", text="Amount Spent"); self.breakdown_tree.heading("percentage", text="% of Total")
        self.breakdown_tree.column("category", width=200, anchor=tk.W); self.breakdown_tree.column("amount", width=150, anchor=tk.E, stretch=tk.NO); self.breakdown_tree.column("percentage", width=100, anchor=tk.E, stretch=tk.NO)
        scrollbar = ttk.Scrollbar(breakdown_frame, orient="vertical", command=self.breakdown_tree.yview, style="Vertical.TScrollbar")
        self.breakdown_tree.configure(yscrollcommand=scrollbar.set)
        self.breakdown_tree.grid(row=0, column=0, sticky="nsew"); scrollbar.grid(row=0, column=1, sticky="ns")
        sorted_categories = sorted(expense_by_category.items(), key=lambda item: item[1], reverse=True)
        if not sorted_categories: self.breakdown_tree.insert("", tk.END, values=("No expenses recorded.", "", ""))
        else:
            for category, amount in sorted_categories:
                percentage = (amount / total_expense * 100) if total_expense else 0
                self.breakdown_tree.insert("", tk.END, values=(category if category else "Uncategorized", format_currency(amount), f"{percentage:.1f}%"))

# --- Base Class for Editing Lists/Dicts ---
class EditListPageBase(BasePage):
    def __init__(self, parent, app, title, data_key, columns, column_config,
                 item_name_singular, add_dialog_fields, edit_dialog_fields):
        super().__init__(parent, app)
        self.data_key = data_key
        self.item_name = item_name_singular
        self.columns = columns
        self.column_config = column_config
        self.add_dialog_fields = add_dialog_fields
        self.edit_dialog_fields = edit_dialog_fields

        self.grid_rowconfigure(1, weight=1); self.grid_columnconfigure(0, weight=1)

        # Title and Action Buttons
        title_frame = tk.Frame(self, bg=theme_colors["background"])
        title_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        ttk.Label(title_frame, text=title, style="Title.TLabel").pack(side=tk.LEFT, padx=(0, 20))

        # Action Buttons (Add and Delete Selected)
        create_stylish_button(title_frame, f"+ Add {self.item_name}", self.add_item).pack(side=tk.RIGHT, padx=(5, 0))

        self.delete_selected_button = create_stylish_button(
            title_frame,
            f"Delete Selected {self.item_name}",
            self.delete_selected_item,
            style="TButton"
        )
        self.delete_selected_button.pack(side=tk.RIGHT, padx=(5, 0))

        # Treeview Setup
        tree_columns_ids = list(self.columns.keys())
        self.tree = ttk.Treeview(self, columns=tree_columns_ids, show="headings", style="Treeview", selectmode="browse")
        for col_id, heading_text in self.columns.items(): self.tree.heading(col_id, text=heading_text)
        for col_id, config in self.column_config.items():
            stretch_val = config.get('stretch', tk.YES); config_copy = config.copy(); config_copy.pop('stretch', None)
            self.tree.column(col_id, **config_copy, stretch=stretch_val)
        self.tree.heading("#0", text="Actions"); self.tree.column("#0", width=100, anchor=tk.CENTER, stretch=tk.NO)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview, style="Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=1, column=0, sticky="nsew"); scrollbar.grid(row=1, column=1, sticky="ns")
        self.tree.bind("<ButtonRelease-1>", self.on_action_click)
        self.populate_data()

    def delete_selected_item(self):
        """Deletes the currently selected item in the treeview."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", f"Please select a {self.item_name} from the list to delete.", parent=self)
            return
        if len(selection) > 1:
            messagebox.showwarning("Multiple Selection", f"Please select only one {self.item_name} to delete.", parent=self)
            return

        item_id = selection[0]
        self.delete_item(item_id)

    def populate_data(self):
        """Populates the treeview with data from app_data."""
        try:
            for item in self.tree.get_children(): self.tree.delete(item)
        except tk.TclError as e: logging.warning(f"TclError clearing tree in {self.__class__.__name__}: {e}")

        data_source = app_data.get(self.data_key)
        if not isinstance(data_source, dict):
            logging.warning(f"Data source '{self.data_key}' is not a dict. Initializing as empty.")
            data_source = {}; app_data[self.data_key] = data_source
        if not data_source: logging.info(f"No data found for '{self.data_key}'."); return

        try:
            # Sort items by name
            def sort_key(item_tuple):
                item_id, details = item_tuple
                if isinstance(details, dict): return str(details.get('name', item_id)).lower()
                return str(item_id).lower()
            sorted_items = sorted(data_source.items(), key=sort_key)
        except Exception as e:
            logging.exception(f"Error sorting {self.data_key}: {e}. Using unsorted."); sorted_items = data_source.items()

        for item_id, details in sorted_items:
            if not isinstance(details, dict): logging.warning(f"Skipping invalid item '{item_id}' in '{self.data_key}'"); continue
            try:
                values = self.get_values_for_item(details)
                self.tree.insert("", tk.END, iid=item_id, values=values, text=" Edit | Delete ")
            except tk.TclError as e: logging.error(f"TclError inserting item {item_id}: {e}")
            except Exception as e:
                logging.exception(f"Error populating item {item_id} for {self.data_key}: {e}")
                try: self.tree.insert("", tk.END, iid=f"error_{item_id}", values=("Error",)*len(self.columns), text="ERROR")
                except: pass

    def get_values_for_item(self, details):
        """
        Abstract method: Subclasses must implement this to return a tuple
        of values corresponding to the columns defined for that page.
        """
        raise NotImplementedError(f"Subclass {self.__class__.__name__} must implement get_values_for_item")

    def _validate_and_process_dialog_result(self, result, is_edit=False, item_id=None):
        """Validates and processes input from a dialog."""
        if result is None: return None
        processed_data = {}
        try:
            dialog_config = self.edit_dialog_fields if is_edit else self.add_dialog_fields
            for field, config in dialog_config.items():
                value = result.get(field); field_type = config.get("type", "text")
                is_required = config.get("required", True); label = config.get('label', field)
                if field_type == "boolean": processed_data[field] = bool(value)
                elif value is None or (isinstance(value, str) and value.strip() == ""):
                    if is_required: raise ValueError(f"Field '{label}' is required.")
                    elif field_type == "date": processed_data[field] = None
                    else: processed_data[field] = None
                elif field_type in ["number", "currency"]:
                    try: processed_data[field] = float(str(value).replace(",", "").strip())
                    except (ValueError, TypeError): raise ValueError(f"Invalid number format for '{label}'.")
                elif field_type == "date":
                    try:
                         date_obj = datetime.datetime.strptime(str(value).strip(), "%Y-%m-%d").date()
                         processed_data[field] = date_obj.strftime("%Y-%m-%d")
                    except (ValueError, TypeError):
                        if not is_required: processed_data[field] = None
                        else: raise ValueError(f"Invalid date format (YYYY-MM-DD) for '{label}'.")
                else: processed_data[field] = str(value).strip()
            processed_data = self.validate_specific_fields(processed_data, is_edit, item_id)
            return processed_data
        except ValueError as e: messagebox.showerror("Invalid Input", f"Please check your input:\n{e}", parent=self); return None
        except Exception as e:
            logging.exception("Error processing dialog result"); messagebox.showerror("Error", f"Error processing input:\n{e}", parent=self); return None

    def validate_specific_fields(self, data, is_edit, item_id):
        """Placeholder: Subclasses override this for specific validation rules."""
        return data

    def add_item(self):
        """Opens a dialog to add a new item and saves it."""
        self._update_dynamic_dialog_fields(self.add_dialog_fields)
        dialog = SimpleEntryDialog(self, f"Add New {self.item_name}", self.add_dialog_fields)
        processed_data = self._validate_and_process_dialog_result(dialog.result, is_edit=False)
        if processed_data:
            try:
                id_prefix = self.data_key[:-1] if self.data_key.endswith('s') else self.data_key
                new_id = get_unique_id(id_prefix)
                id_field_name = f"{id_prefix}_id"
                if id_field_name in self.columns: processed_data[id_field_name] = new_id
                if not isinstance(app_data.get(self.data_key), dict): app_data[self.data_key] = {}
                app_data[self.data_key][new_id] = processed_data
                log_activity(f"Added {self.item_name}: {processed_data.get('name', new_id)}")
                self.populate_data()
                logging.info(f"Added new {self.item_name} with ID {new_id}")
            except Exception as e:
                logging.exception(f"Error adding {self.item_name}"); messagebox.showerror("Error", f"Could not add {self.item_name}.\n{e}", parent=self)

    def edit_item(self, item_id):
        """Opens a dialog to edit an existing item and saves changes."""
        data_source = app_data.get(self.data_key)
        if not isinstance(data_source, dict) or item_id not in data_source:
            messagebox.showerror("Error", f"{self.item_name} ID '{item_id}' not found.", parent=self); self.populate_data(); return
        original_data = data_source[item_id]
        if not isinstance(original_data, dict): messagebox.showerror("Error", f"Invalid data for {self.item_name} ID '{item_id}'.", parent=self); return

        edit_fields_with_initials = {}
        for field, config in self.edit_dialog_fields.items():
            new_config = config.copy(); new_config['initial'] = original_data.get(field); edit_fields_with_initials[field] = new_config
        self._update_dynamic_dialog_fields(edit_fields_with_initials)
        dialog = SimpleEntryDialog(self, f"Edit {self.item_name}", edit_fields_with_initials)
        processed_data = self._validate_and_process_dialog_result(dialog.result, is_edit=True, item_id=item_id)
        if processed_data:
            try:
                if item_id in app_data.get(self.data_key, {}):
                    id_prefix = self.data_key[:-1] if self.data_key.endswith('s') else self.data_key
                    id_field_name = f"{id_prefix}_id"
                    processed_data.pop(id_field_name, None)

                    app_data[self.data_key][item_id].update(processed_data)
                    log_activity(f"Edited {self.item_name}: {processed_data.get('name', item_id)}")
                    self.populate_data()
                    logging.info(f"Edited {self.item_name} with ID {item_id}")
                else: messagebox.showerror("Error", f"{self.item_name} removed before edit saved.", parent=self); self.populate_data()
            except Exception as e:
                logging.exception(f"Error updating {self.item_name} {item_id}"); messagebox.showerror("Error", f"Could not update {self.item_name}.\n{e}", parent=self)

    def delete_item(self, item_id):
        """Deletes a selected item after confirmation and dependency check."""
        data_source = app_data.get(self.data_key)
        if not isinstance(data_source, dict) or item_id not in data_source:
            logging.warning(f"Attempted to delete non-existent {self.item_name} ID: {item_id}")
            self.populate_data()
            return
        item_details = data_source[item_id]; item_name_display = item_details.get("name", item_id) if isinstance(item_details, dict) else item_id
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{item_name_display}'?\nThis action cannot be undone.", parent=self):
            try:
                can_delete, reason = self.check_can_delete(item_id)
                if not can_delete: messagebox.showwarning("Cannot Delete", reason, parent=self); return
                if item_id in app_data.get(self.data_key, {}):
                    del app_data[self.data_key][item_id]
                    log_activity(f"Deleted {self.item_name}: {item_name_display}")
                    self.populate_data()
                    logging.info(f"Deleted {self.item_name}: {item_name_display} (ID: {item_id})")
                else:
                    messagebox.showwarning("Delete Warning", f"{self.item_name} '{item_name_display}' was already removed.", parent=self)
                    self.populate_data()
            except Exception as e:
                logging.exception(f"Error deleting {self.item_name} {item_id}"); messagebox.showerror("Error", f"Could not delete {self.item_name}.\n{e}", parent=self)

    def check_can_delete(self, item_id):
        """Placeholder: Subclasses override to check dependencies (e.g., transactions)."""
        return True, ""

    def on_action_click(self, event):
        """Handles clicks on the action column (Edit/Delete)."""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column_id = self.tree.identify_column(event.x); item_iid = self.tree.identify_row(event.y)
            if item_iid and column_id == "#0":
                action_text = self.tree.item(item_iid, "text")
                if not action_text or "|" not in action_text: return
                data_source = app_data.get(self.data_key)
                if not isinstance(data_source, dict) or item_iid not in data_source:
                    messagebox.showwarning("Action Error", "Item not found.", parent=self); self.populate_data(); return
                try:
                    col_box = self.tree.bbox(item_iid, column="#0")
                    if col_box:
                        click_x_relative = event.x - col_box[0]; separator_pos = col_box[2] * 0.5
                        if "Edit" in action_text and click_x_relative < separator_pos: self.edit_item(item_iid)
                        elif "Delete" in action_text and click_x_relative >= separator_pos: self.delete_item(item_iid)
                    else: logging.warning(f"Could not get bbox for action cell of item {item_iid}")
                except tk.TclError as e: logging.error(f"TclError processing action click for {item_iid}: {e}")
                except Exception as e:
                    logging.exception(f"Error processing action click for {item_iid}"); messagebox.showerror("Error", "Could not process action.", parent=self)

    def _update_dynamic_dialog_fields(self, dialog_fields):
        """Updates 'combo' type fields in dialog configs with current data."""
        wallets_data = app_data.get("wallets", {}); categories_data = app_data.get("categories", {}); budgets_data = app_data.get("budgets", {})
        wallet_names = sorted([w.get("name", "") for w in wallets_data.values()
                               if isinstance(w, dict) and "name" in w and w["name"] is not None],
                              key=lambda x: str(x).lower())

        category_names = sorted([c.get("name", "") for c in categories_data.values()
                                 if isinstance(c, dict) and "name" in c and c["name"] is not None],
                                key=lambda x: str(x).lower())

        budget_names = sorted([b.get("name", "") for b in budgets_data.values()
                               if isinstance(b, dict) and "name" in b and b["name"] is not None],
                              key=lambda x: str(x).lower())
        budget_cycles = ["Once", "Daily", "Weekly", "Monthly", "Yearly"]
        try:
            for field, config in dialog_fields.items():
                if config.get("type") == "combo":
                    if field in ["wallet", "account", "from_account", "to_account"]: config["values"] = wallet_names
                    elif field == "category": config["values"] = category_names
                    elif field == "cycle": config["values"] = budget_cycles
                    elif field == "budget_name": config["values"] = budget_names
        except Exception as e: logging.exception(f"Error updating dynamic dialog fields: {e}")


# --- Subclasses for Specific Edit Pages ---

# WalletsPage
class WalletsPage(EditListPageBase):
    def __init__(self, parent, app):
        columns = {"name": "Wallet Name", "balance": "Balance"}
        column_config = {
            "name": {"width": 200, "anchor": tk.W, "stretch": tk.YES},
            "balance": {"width": 120, "anchor": tk.E, "stretch": tk.NO}
        }

        add_dialog_fields = {
            "name": {"label": "Wallet Name:", "type": "text", "required": True}
        }
        edit_dialog_fields = {
            "name": {"label": "Wallet Name:", "type": "text", "required": True},
            "balance": {"label": "Corrected Balance:", "type": "currency", "required": True}
        }
        super().__init__(parent, app, "Edit Wallets", "wallets", columns, column_config, "Wallet", add_dialog_fields, edit_dialog_fields)

    def get_values_for_item(self, details):
        """Returns display values for a wallet item."""
        return (
            details.get('name', 'N/A'),
            format_currency(details.get('balance', 0.0))
        )

    def validate_specific_fields(self, data, is_edit, item_id):
        """Validates wallet name uniqueness and balance input."""
        name_to_check = data.get("name", "").strip(); wallets_dict = app_data.get("wallets", {})
        if not name_to_check: raise ValueError("Wallet name cannot be empty.")
        for w_id, w_details in wallets_dict.items():
            if is_edit and w_id == item_id: continue
            if isinstance(w_details, dict) and w_details.get("name", "").lower() == name_to_check.lower(): raise ValueError(f"Wallet named '{data.get('name')}' already exists.")

        if 'balance' in data and data.get('balance') is None:
             raise ValueError("Corrected Balance cannot be empty.")
        return data

    def add_item(self):
        """Opens dialog to add new wallet, sets initial balance to 0.0, and saves."""
        self._update_dynamic_dialog_fields(self.add_dialog_fields)
        dialog = SimpleEntryDialog(self, f"Add New {self.item_name}", self.add_dialog_fields)
        processed_data = self._validate_and_process_dialog_result(dialog.result, is_edit=False)

        if processed_data:
            try:
                processed_data['balance'] = 0.0

                id_prefix = self.data_key[:-1] if self.data_key.endswith('s') else self.data_key
                new_id = get_unique_id(id_prefix)
                id_field_name = f"{id_prefix}_id"
                if id_field_name in self.columns: processed_data[id_field_name] = new_id

                if not isinstance(app_data.get(self.data_key), dict): app_data[self.data_key] = {}
                app_data[self.data_key][new_id] = processed_data
                log_activity(f"Added {self.item_name}: {processed_data.get('name', new_id)}")
                self.populate_data()
                logging.info(f"Added new {self.item_name} with ID {new_id} (Balance: 0.00)")
            except Exception as e:
                logging.exception(f"Error adding {self.item_name}"); messagebox.showerror("Error", f"Could not add {self.item_name}.\n{e}", parent=self)

    def check_can_delete(self, item_id):
        """Checks if a wallet is used in any transactions before deletion."""
        wallets_dict = app_data.get("wallets", {}); transactions = app_data.get("transactions", [])
        if item_id not in wallets_dict or not isinstance(wallets_dict[item_id], dict): return False, "Wallet not found."
        wallet_name = wallets_dict[item_id].get("name")
        if not wallet_name: return True, ""
        if not isinstance(transactions, list): transactions = []
        if any(isinstance(tx, dict) and wallet_name in [tx.get("wallet"), tx.get("from_account"), tx.get("to_account")] for tx in transactions):
            return False, f"Cannot delete wallet '{wallet_name}' used in transactions."
        return True, ""

# BudgetPage
class BudgetPage(EditListPageBase):
    def __init__(self, parent, app):
        columns = {"name": "Budget Name", "allocated": "Allocated", "cycle": "Cycle", "spent": "Spent"}
        column_config = {
            "name": {"width": 200, "anchor": tk.W, "stretch": tk.YES},
            "allocated": {"width": 120, "anchor": tk.E, "stretch": tk.NO},
            "cycle": {"width": 80, "anchor": tk.W, "stretch": tk.NO},
            "spent": {"width": 120, "anchor": tk.E, "stretch": tk.NO}
        }
        budget_cycles = ["Once", "Daily", "Weekly", "Monthly", "Yearly"]
        add_dialog_fields = {
            "name": {"label": "Budget Name:", "type": "text", "required": True},
            "allocated": {"label": "Allocated Amount:", "type": "currency", "required": True, "initial": "100.00"},
            "cycle": {"label": "Cycle:", "type": "combo", "values": budget_cycles, "required": True,
                      "initial": "Monthly"}
        }
        edit_dialog_fields = {
            "name": {"label": "Budget Name:", "type": "text", "required": True},
            "allocated": {"label": "Allocated Amount:", "type": "currency", "required": True},
            "cycle": {"label": "Cycle:", "type": "combo", "values": budget_cycles, "required": True}
        }
        super().__init__(parent, app, "Budgets", "budgets", columns, column_config, "Budget", add_dialog_fields,edit_dialog_fields)

    def get_values_for_item(self, details):
        """Gets display values for a budget, calculating current spending."""
        budget_name = details.get('name', 'N/A')
        calculated_spent = self._calculate_spent_for_budget(budget_name)
        return (
            budget_name,
            format_currency(details.get('allocated', 0.0)),
            details.get('cycle', 'N/A'),
            format_currency(calculated_spent)
        )

    def _calculate_spent_for_budget(self, budget_name):
        """Calculates total spending linked to a specific budget name."""
        if not budget_name: return 0.0
        total_spent = 0.0
        transactions = app_data.get("transactions", [])
        if not isinstance(transactions, list): return 0.0

        for tx in transactions:
            if (isinstance(tx, dict) and
                    tx.get("type") == "expense" and
                    tx.get("linked_budget") == budget_name and
                    isinstance(tx.get("amount"), (int, float))):
                total_spent += abs(tx["amount"])
        return total_spent

    def validate_specific_fields(self, data, is_edit, item_id):
        """Validates budget name uniqueness and positive allocation."""
        name_to_check = data.get("name", "").strip(); budgets_dict = app_data.get("budgets", {})
        if not name_to_check: raise ValueError("Budget name cannot be empty.")
        for bud_id, bud_details in budgets_dict.items():
            if is_edit and bud_id == item_id: continue
            if isinstance(bud_details, dict) and bud_details.get("name", "").lower() == name_to_check.lower():
                raise ValueError(f"Budget named '{data.get('name')}' already exists.")
        allocated = data.get('allocated'); cycle = data.get('cycle')
        if allocated is None or allocated <= 0: raise ValueError("Allocated amount must be positive.")
        if not cycle or cycle not in ["Once", "Daily", "Weekly", "Monthly", "Yearly"]:
            raise ValueError("Please select a valid budget cycle.")
        return data

    def check_can_delete(self, item_id):
        """Checks if a budget is linked to any transactions before deletion."""
        budgets_dict = app_data.get("budgets", {}); transactions = app_data.get("transactions", [])
        if item_id not in budgets_dict or not isinstance(budgets_dict[item_id], dict):
            return False, "Budget not found."

        budget_name = budgets_dict[item_id].get("name")
        if not budget_name:
            return True, ""

        if not isinstance(transactions, list): transactions = []
        if any(isinstance(tx, dict) and tx.get("linked_budget") == budget_name for tx in transactions):
            return False, f"Cannot delete budget '{budget_name}' as it is linked to existing transactions."

        return True, ""

# --- GoalsPage (Subclass) ---
class GoalsPage(EditListPageBase):
    def __init__(self, parent, app):
        columns = {"name": "Goal Name", "target": "Target Amount", "saved": "Amount Saved", "due_date": "Due Date"}
        column_config = {
            "name": {"width": 200, "anchor": tk.W, "stretch": tk.YES},
            "target": {"width": 120, "anchor": tk.E, "stretch": tk.NO},
            "saved": {"width": 120, "anchor": tk.E, "stretch": tk.NO},
            "due_date": {"width": 100, "anchor": tk.CENTER, "stretch": tk.NO}
        }

        add_dialog_fields = {
            "name": {"label": "Goal Name:", "type": "text", "required": True},
            "target": {"label": "Target Amount:", "type": "currency", "required": True, "initial": "1000.00"},
            "due_date": {"label": "Target Date (YYYY-MM-DD):", "type": "date", "required": False}
        }
        edit_dialog_fields = {
            "name": {"label": "Goal Name:", "type": "text", "required": True},
            "target": {"label": "Target Amount:", "type": "currency", "required": True},
            "saved": {"label": "Base Amount Saved:", "type": "currency", "required": True},
            "due_date": {"label": "Target Date (YYYY-MM-DD):", "type": "date", "required": False}
        }
        super().__init__(parent, app, "Goals", "goals", columns, column_config, "Goal", add_dialog_fields,
                         edit_dialog_fields)

    def get_values_for_item(self, details):
        """Gets display values for a goal, calculating effective saved amount including linked expenses."""
        goal_name = details.get('name', 'N/A')
        target = details.get('target', 0.0)
        base_saved = details.get('saved', 0.0)

        linked_expense_contribution = 0.0
        user_transactions = app_data.get("transactions", [])
        if not isinstance(user_transactions, list): user_transactions = []

        # Calculate linked contribution from expenses
        for tx in user_transactions:
            if (isinstance(tx, dict) and
                    tx.get("linked_goal") == goal_name and
                    tx.get("type") == "expense" and
                    isinstance(tx.get("amount"), (int, float)) and
                    tx["amount"] < 0):
                linked_expense_contribution += abs(tx["amount"])

        effective_saved = base_saved + linked_expense_contribution

        return (
            goal_name,
            format_currency(target),
            format_currency(effective_saved),
            details.get('due_date', '') or ""
        )

    def validate_specific_fields(self, data, is_edit, item_id):
        """Validates goal name uniqueness, amounts, and date format."""
        name_to_check = data.get("name", "").strip(); goals_dict = app_data.get("goals", {})
        if not name_to_check: raise ValueError("Goal name cannot be empty.")
        for g_id, g_details in goals_dict.items():
            if is_edit and g_id == item_id: continue
            if isinstance(g_details, dict) and g_details.get("name", "").lower() == name_to_check.lower():
                raise ValueError(f"Goal named '{data.get('name')}' already exists.")

        target = data.get('target'); saved = data.get('saved')
        if target is None or target <= 0: raise ValueError("Target amount must be positive.")

        if 'saved' in data:
            if saved is None:
                 raise ValueError("Base saved amount is required when editing.")
            if saved < 0:
                 raise ValueError("Saved amount must be zero or positive.")

        # Validate date format if provided
        due_date_str = data.get('due_date')
        if due_date_str:
            try:
                 datetime.datetime.strptime(due_date_str, "%Y-%m-%d")
            except (ValueError, TypeError):
                 raise ValueError("Invalid date format. Use YYYY-MM-DD or leave empty.")

        return data

    def add_item(self):
        """Opens dialog to add new goal, sets initial saved amount to 0.0, and saves."""
        self._update_dynamic_dialog_fields(self.add_dialog_fields)
        dialog = SimpleEntryDialog(self, f"Add New {self.item_name}", self.add_dialog_fields)
        processed_data = self._validate_and_process_dialog_result(dialog.result, is_edit=False)

        if processed_data:
            try:
                processed_data['saved'] = 0.0

                id_prefix = self.data_key[:-1] if self.data_key.endswith('s') else self.data_key
                new_id = get_unique_id(id_prefix)
                id_field_name = f"{id_prefix}_id"
                if id_field_name in self.columns: processed_data[id_field_name] = new_id

                if not isinstance(app_data.get(self.data_key), dict): app_data[self.data_key] = {}
                app_data[self.data_key][new_id] = processed_data
                log_activity(f"Added {self.item_name}: {processed_data.get('name', new_id)}")
                self.populate_data()
                logging.info(f"Added new {self.item_name} with ID {new_id} (Saved: 0.00)")
            except Exception as e:
                logging.exception(f"Error adding {self.item_name}"); messagebox.showerror("Error", f"Could not add {self.item_name}.\n{e}", parent=self)

    def check_can_delete(self, item_id):
        """Checks if a goal is linked to any transactions before deletion."""
        goals_dict = app_data.get("goals", {}); transactions = app_data.get("transactions", [])
        if item_id not in goals_dict or not isinstance(goals_dict[item_id], dict):
            return False, "Goal not found."
        goal_name = goals_dict[item_id].get("name")
        if not goal_name: return True, ""
        if not isinstance(transactions, list): transactions = []
        if any(isinstance(tx, dict) and tx.get("linked_goal") == goal_name for tx in transactions):
            return False, f"Cannot delete goal '{goal_name}' as it is linked to existing income transactions."
        return True, ""


# --- SettingsPage Class ---
class SettingsPage(BasePage):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.grid_columnconfigure(0, weight=1)
        self.create_ui_elements()

    def create_ui_elements(self):
        """Creates and arranges UI elements for the settings page."""
        for widget in self.winfo_children(): widget.destroy()

        ttk.Label(self, text="Settings", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 20))

        # Theme Settings Card
        self.theme_frame = create_card_frame(self)
        self.theme_frame.grid(row=1, column=0, sticky="ew", padx=0, pady=10)
        self.theme_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(self.theme_frame, text="Theme", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2,
                                                                                 sticky="w", padx=10, pady=(10, 5))
        self.theme_var = tk.StringVar(value=app_data.get("settings", {}).get("theme", "dark"))
        light_rb = ttk.Radiobutton(self.theme_frame, text="Light Mode", variable=self.theme_var, value="light",
                                   command=self.change_theme, style="Card.TRadiobutton")
        light_rb.grid(row=1, column=0, sticky="w", padx=(20, 10), pady=2)
        dark_rb = ttk.Radiobutton(self.theme_frame, text="Dark Mode", variable=self.theme_var, value="dark",
                                  command=self.change_theme, style="Card.TRadiobutton")
        dark_rb.grid(row=2, column=0, sticky="w", padx=(20, 10), pady=(2, 10))

        # User Profile Actions Card
        self.action_frame = create_card_frame(self)
        self.action_frame.grid(row=2, column=0, sticky="ew", padx=0, pady=10)
        ttk.Label(self.action_frame, text="User Profile & Application", style="CardTitle.TLabel").pack(padx=10, pady=(10, 5),
                                                                                         anchor='w')

        buttons_frame = tk.Frame(self.action_frame, bg=theme_colors["card"])
        buttons_frame.pack(fill="x", padx=10, pady=(5, 10))

        # Buttons packed from left to right
        switch_button = create_stylish_button(buttons_frame, "Switch User Profile", self.switch_user, style="TButton")
        switch_button.pack(side=tk.LEFT, padx=(0, 5))

        reset_button = create_stylish_button(buttons_frame, "Reset Data", self.reset_data, style="TButton")
        reset_button.pack(side=tk.LEFT, padx=5)

        delete_button = create_stylish_button(buttons_frame, "Delete User", self.delete_user, style="TButton")
        delete_button.pack(side=tk.LEFT, padx=5)

        exit_button = create_stylish_button(
            buttons_frame,
            "Exit Application",
            self.exit_application,
            style="TButton"
        )
        exit_button.pack(side=tk.LEFT, padx=5)

    def change_theme(self):
        """Changes the application's visual theme."""
        new_theme = self.theme_var.get()
        logging.info(f"Theme selection changed to: {new_theme}")
        self.app.switch_theme(new_theme)
        self.create_ui_elements()
        self.configure(bg=theme_colors["background"])

    def switch_user(self):
        """Closes the current application and returns to the user selection screen."""
        logging.info("Switch User action initiated.")
        self.app.on_closing()

    def reset_data(self):
        """Resets all financial data for the current user."""
        if messagebox.askyesno("Reset Data",
                               "This will permanently clear all transactions, wallets, budgets, and goals for the current user.\n\n"
                               "Your user profile will remain, but all its associated financial data will be reset to defaults.\n\n"
                               "This action cannot be undone. Are you absolutely sure?",
                               icon='warning', parent=self):

            try:
                user_id = self.app.current_user_id
                logging.warning(f"Resetting all data for user {user_id}")
                # Clear data in memory
                app_data["wallets"] = {}
                app_data["budgets"] = {}
                app_data["goals"] = {}
                app_data["transactions"] = []
                app_data["activity_log"] = []

                # Create default wallet again
                wallet_id = get_unique_id("wallet")
                app_data["wallets"][wallet_id] = {"wallet_id": wallet_id, "name": "Cash", "balance": 0.0}

                # Save the now empty/default data
                save_user_data(user_id)
                log_activity("Reset all user data")

                messagebox.showinfo("Data Reset", "All financial data for this user has been reset successfully.", parent=self)
                self.app.show_page("Home")
            except Exception as e:
                logging.exception("Error resetting user data")
                messagebox.showerror("Error", f"Could not reset data:\n{e}", parent=self)

    def delete_user(self):
        """Deletes the current user profile and all associated data."""
        if messagebox.askyesno("Delete User",
                               "This will permanently delete the current user profile and ALL associated data (transactions, wallets, etc.).\n\n"
                               "This action cannot be undone. Are you absolutely sure?",
                               icon='warning', parent=self):
            try:
                user_id = self.app.current_user_id
                if not user_id:
                    messagebox.showerror("Error", "Cannot determine the current user ID.", parent=self)
                    return

                user_name = app_data.get("user_profiles", {}).get(user_id, {}).get("name", f"ID: {user_id}")
                logging.warning(f"Attempting to delete user: {user_name} (ID: {user_id})")

                # Remove user from profiles dictionary
                if user_id in app_data.get("user_profiles", {}):
                    del app_data["user_profiles"][user_id]
                    save_user_profiles_to_csv()

                # Delete associated user data files
                data_types = ["wallets", "budgets", "goals", "transactions", "activity_log", "settings"]
                for data_type in data_types:
                    file_path = get_user_data_file_path(user_id, data_type)
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                            logging.info(f"Deleted user data file: {file_path}")
                        except OSError as e:
                            logging.error(f"Could not delete file {file_path}: {e}")

                messagebox.showinfo("User Deleted", f"User profile '{user_name}' and all associated data have been permanently deleted.", parent=self)

                # Close current application and return to account selection
                self.app.on_closing()

            except Exception as e:
                logging.exception("Error deleting user")
                messagebox.showerror("Error", f"An error occurred while deleting the user:\n{e}", parent=self)

    def exit_application(self):
        """Saves data and immediately closes the entire application."""
        logging.info("Exit Application action initiated from Settings (no confirmation).")
        self.app.perform_full_exit()

# --- Add Transaction Dialog ---
class AddTransactionDialog(tk.Toplevel):
    def __init__(self, parent_app):
        super().__init__(parent_app)
        self.app = parent_app
        self.configure(bg=theme_colors["dialog_bg"], padx=20, pady=20)
        self.title("Add Transaction")
        self.geometry("550x650")
        self.resizable(False, False)
        self.transient(parent_app); self.grab_set()
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(0, weight=1)

        container = tk.Frame(self, bg=theme_colors["dialog_bg"])
        container.grid(row=0, column=0, sticky='nsew')
        container.grid_rowconfigure(0, weight=1); container.grid_columnconfigure(0, weight=1)

        # Data Variables
        self.amount_var = tk.StringVar(value="0.00"); self.title_var = tk.StringVar()
        self.wallet_var = tk.StringVar(); self.to_wallet_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.time_var = tk.StringVar(value=datetime.datetime.now().strftime("%H:%M"))
        self.selected_category_var = tk.StringVar()
        self.budget_var = tk.StringVar(value="None")
        self.goal_var = tk.StringVar(value="None")

        # Styles
        self.dialog_style = ttk.Style(self); self.dialog_style.theme_use('clam')
        self._configure_dialog_styles()

        # Notebook for transaction types
        self.notebook = ttk.Notebook(container, style='TNotebook')
        self.expense_tab = tk.Frame(self.notebook, bg=theme_colors["dialog_card"], padx=10, pady=10)
        self.income_tab = tk.Frame(self.notebook, bg=theme_colors["dialog_card"], padx=10, pady=10)
        self.transfer_tab = tk.Frame(self.notebook, bg=theme_colors["dialog_card"], padx=10, pady=10)
        self.create_tab_widgets(self.expense_tab, "expense")
        self.create_tab_widgets(self.income_tab, "income")
        self.create_tab_widgets(self.transfer_tab, "transfer")
        self.notebook.add(self.expense_tab, text=" Expense "); self.notebook.add(self.income_tab, text=" Income "); self.notebook.add(self.transfer_tab, text=" Transfer ")
        self.notebook.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)
        self.on_tab_change()

        # Action Buttons
        action_frame = tk.Frame(container, bg=theme_colors["dialog_bg"])
        action_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        action_frame.grid_columnconfigure((0, 1), weight=1)
        cancel_style = "Cancel.Dialog.TButton" if "Cancel.Dialog.TButton" in self.dialog_style.element_names() else "TButton"
        add_style = "Dialog.TButton" if "Dialog.TButton" in self.dialog_style.element_names() else "TButton"
        self.cancel_button = ttk.Button(action_frame, text="Cancel", command=self.destroy, style=cancel_style)
        self.cancel_button.grid(row=0, column=0, sticky="ew", padx=(0, 5), ipady=5)
        self.add_button = ttk.Button(action_frame, text="Add Transaction", command=self.add_transaction, style=add_style)
        self.add_button.grid(row=0, column=1, sticky="ew", padx=(5, 0), ipady=5)

        self.bind("<Map>", self._set_initial_focus, add="+")
        self.center_dialog(parent_app)
        self.wait_window(self)

    def _configure_dialog_styles(self):
        """Configures ttk styles specifically for this dialog."""
        bg=theme_colors["dialog_bg"]; fg=theme_colors["dialog_fg"]; card_bg=theme_colors["dialog_card"]
        accent=theme_colors["accent"]; accent_darker=theme_colors["accent_darker"]; disabled_fg=theme_colors["disabled"]
        button_fg=theme_colors["button_fg"]; red=theme_colors["red"]
        self.dialog_style.configure('Dialog.TLabel', background=card_bg, foreground=fg, font=FONT_NORMAL)
        self.dialog_style.configure('Dialog.TEntry', fieldbackground=card_bg, foreground=fg, insertcolor=fg, borderwidth=0, padding=5)
        self.dialog_style.map('Dialog.TEntry', fieldbackground=[('focus', card_bg)])
        self.dialog_style.configure('Dialog.TCombobox', fieldbackground=card_bg, foreground=fg, selectbackground=card_bg, selectforeground=fg, arrowcolor=fg, borderwidth=0, padding=5)
        self.dialog_style.map('Dialog.TCombobox', fieldbackground=[('readonly', card_bg)])
        self.dialog_style.configure('Dialog.TButton', background=accent, foreground=button_fg, font=FONT_BOLD)
        self.dialog_style.map('Dialog.TButton', background=[('active', accent_darker)])
        self.dialog_style.configure('Cancel.Dialog.TButton', background=disabled_fg, foreground=fg)
        self.dialog_style.map('Cancel.Dialog.TButton', background=[('active', red)], foreground=[('active', button_fg)])
        self.dialog_style.configure("Dialog.Toolbutton", anchor="w", padding=(2,5), font=FONT_NORMAL, background=card_bg, foreground=fg, borderwidth=1, relief="raised")
        self.dialog_style.map("Dialog.Toolbutton", relief=[('selected', 'sunken'), ('active', 'raised')], background=[('selected', accent), ('active', card_bg)], foreground=[('selected', button_fg), ('active', fg)])

    def _set_initial_focus(self, event=None):
        """Sets focus to the amount entry field when dialog opens."""
        try:
            current_tab_widget = self.notebook.nametowidget(self.notebook.select())
            for widget in current_tab_widget.winfo_children():
                if isinstance(widget, ttk.Entry) and widget.cget("textvariable") == str(self.amount_var):
                    widget.focus_set(); widget.select_range(0, tk.END); break
            self.unbind("<Map>")
        except (tk.TclError, AttributeError) as e:
            logging.warning(f"Error setting initial focus in AddTransactionDialog: {e}")
            try: self.unbind("<Map>")
            except: pass

    def center_dialog(self, parent):
        """Centers the dialog window relative to its parent."""
        try:
            self.update_idletasks(); parent_x=parent.winfo_rootx(); parent_y=parent.winfo_rooty()
            parent_w=parent.winfo_width(); parent_h=parent.winfo_height(); dialog_w=self.winfo_width(); dialog_h=self.winfo_height()
            x = parent_x + (parent_w // 2) - (dialog_w // 2); y = parent_y + (parent_h // 3) - (dialog_h // 2)
            screen_w=self.winfo_screenwidth(); screen_h=self.winfo_screenheight()
            x = max(0, min(x, screen_w - dialog_w)); y = max(0, min(y, screen_h - dialog_h))
            self.geometry(f"+{x}+{y}")
        except Exception as e: logging.warning(f"Could not center AddTransactionDialog: {e}")

    def create_tab_widgets(self, tab_frame, tab_type):
        """Creates input widgets for expense, income, or transfer tabs."""
        tab_frame.grid_columnconfigure(1, weight=1)
        row_num = 0
        label_style = 'Dialog.TLabel'
        entry_style = 'Dialog.TEntry'
        combo_style = 'Dialog.TCombobox'
        cat_button_style = 'Dialog.Toolbutton'

        # Amount entry
        ttk.Label(tab_frame, text="Amount:", font=FONT_BOLD, style=label_style).grid(row=row_num, column=0, sticky="w", padx=(0, 10), pady=(5, 2))
        amount_entry = ttk.Entry(tab_frame, textvariable=self.amount_var, font=FONT_LARGE, justify=tk.RIGHT, style=entry_style)
        amount_entry.grid(row=row_num, column=1, sticky="ew", ipady=5)
        row_num += 1

        # Title entry
        ttk.Label(tab_frame, text="Title:", font=FONT_BOLD, style=label_style).grid(row=row_num, column=0, sticky="w", padx=(0, 10), pady=(10, 2))
        title_entry = ttk.Entry(tab_frame, textvariable=self.title_var, font=FONT_NORMAL, style=entry_style)
        title_entry.grid(row=row_num, column=1, sticky="ew")
        row_num += 1

        # Wallet selection
        wallet_data = app_data.get("wallets", {})
        wallet_names = sorted(
            [w.get('name', '') for w in wallet_data.values()
             if isinstance(w, dict) and 'name' in w and w['name'] is not None],
            key=lambda x: str(x).lower()
        )
        if not wallet_names:
            ttk.Label(tab_frame, text="No wallets found. Please create one first.", foreground=theme_colors["red"], style=label_style).grid(row=row_num, column=0, columnspan=2, sticky="ew", pady=5)
            row_num += (1 if tab_type != 'transfer' else 2)
        else:
            if tab_type == "transfer":
                ttk.Label(tab_frame, text="From Wallet:", font=FONT_BOLD, style=label_style).grid(row=row_num, column=0, sticky="w", padx=(0, 10), pady=(10, 2))
                from_wallet_combo = ttk.Combobox(tab_frame, textvariable=self.wallet_var, values=wallet_names, state='readonly', font=FONT_NORMAL, style=combo_style)
                from_wallet_combo.grid(row=row_num, column=1, sticky="ew")
                row_num += 1
                if wallet_names: from_wallet_combo.current(0)

                ttk.Label(tab_frame, text="To Wallet:", font=FONT_BOLD, style=label_style).grid(row=row_num, column=0, sticky="w", padx=(0, 10), pady=(10, 2))
                to_wallet_combo = ttk.Combobox(tab_frame, textvariable=self.to_wallet_var, values=wallet_names, state='readonly', font=FONT_NORMAL, style=combo_style)
                to_wallet_combo.grid(row=row_num, column=1, sticky="ew")
                row_num += 1
                if len(wallet_names) > 1:
                    to_wallet_combo.current(1)
                elif wallet_names:
                    to_wallet_combo.current(0)
            else:
                ttk.Label(tab_frame, text="Wallet:", font=FONT_BOLD, style=label_style).grid(row=row_num, column=0, sticky="w", padx=(0, 10), pady=(10, 2))
                wallet_combo = ttk.Combobox(tab_frame, textvariable=self.wallet_var, values=wallet_names, state='readonly', font=FONT_NORMAL, style=combo_style)
                wallet_combo.grid(row=row_num, column=1, sticky="ew")
                row_num += 1
                if wallet_names: wallet_combo.current(0)

        # Date/Time entry
        datetime_frame = tk.Frame(tab_frame, bg=theme_colors["dialog_card"])
        datetime_frame.grid(row=row_num, column=0, columnspan=2, sticky="ew", pady=(10, 5))
        datetime_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        ttk.Label(datetime_frame, text="Date:", font=FONT_BOLD, style=label_style).grid(row=0, column=0, sticky="w")
        date_entry = ttk.Entry(datetime_frame, textvariable=self.date_var, font=FONT_NORMAL, width=12, style=entry_style)
        date_entry.grid(row=0, column=1, sticky="ew", padx=(0, 15))
        ttk.Label(datetime_frame, text="Time:", font=FONT_BOLD, style=label_style).grid(row=0, column=2, sticky="w", padx=(15, 0))
        time_entry = ttk.Entry(datetime_frame, textvariable=self.time_var, font=FONT_NORMAL, width=8, style=entry_style)
        time_entry.grid(row=0, column=3, sticky="ew")
        row_num += 1

        if tab_type in ["expense", "income"]:
            category_label_text = "Category:" if tab_type == "expense" else "Income Source:"
            ttk.Label(tab_frame, text=category_label_text, font=FONT_BOLD, style=label_style).grid(row=row_num,
                                                                                                   column=0,
                                                                                                   sticky="nw",
                                                                                                   padx=(0, 10),
                                                                                                   pady=(15, 5))
            cat_outer_frame = tk.Frame(tab_frame, bg=theme_colors["dialog_card"])
            cat_outer_frame.grid(row=row_num, column=1, sticky="nsew", pady=(15, 10))
            cat_outer_frame.grid_rowconfigure(0, weight=1)
            cat_outer_frame.grid_columnconfigure(0, weight=1)
            cat_canvas = tk.Canvas(cat_outer_frame, bg=theme_colors["dialog_card"], highlightthickness=0,
                                   height=150)
            cat_scrollbar = ttk.Scrollbar(cat_outer_frame, orient="vertical", command=cat_canvas.yview,
                                          style="Vertical.TScrollbar")
            category_buttons_frame = tk.Frame(cat_canvas, bg=theme_colors["dialog_card"])
            cat_canvas.create_window((0, 0), window=category_buttons_frame, anchor="nw", tags="cat_frame")
            cat_canvas.configure(yscrollcommand=cat_scrollbar.set)
            category_buttons_frame.bind("<Configure>",
                                        lambda e, canvas=cat_canvas: canvas.configure(scrollregion=canvas.bbox("all")))
            cat_canvas.bind("<Configure>",
                            lambda e, canvas=cat_canvas, frame=category_buttons_frame: canvas.itemconfig("cat_frame",
                                                                                                         width=e.width))
            cat_canvas.grid(row=0, column=0, sticky="nsew")
            cat_scrollbar.grid(row=0, column=1, sticky="ns")
            cat_cols = 3
            for i in range(cat_cols): category_buttons_frame.grid_columnconfigure(i, weight=0)

            all_categories = app_data.get("categories", {})
            relevant_categories = {}
            default_category_name = None

            relevant_categories = {
                k: v for k, v in all_categories.items() if
                isinstance(v, dict) and v.get('type') == tab_type
            }

            # Determine a suitable default category
            if tab_type == "expense":
                 prefs = ["Groceries", "Fuel", "Rent/Mortgage", "Utilities"]
                 for pref in prefs:
                     if any(c.get("name") == pref for c in relevant_categories.values()):
                         default_category_name = pref; break
                 if not default_category_name and relevant_categories:
                      sorted_relevant = sorted(relevant_categories.values(), key=lambda x: x.get('name',''))
                      default_category_name = sorted_relevant[0].get('name') if sorted_relevant else None

            elif tab_type == "income":
                 prefs = ["Salary", "Freelance", "Business"]
                 for pref in prefs:
                     if any(c.get("name") == pref for c in relevant_categories.values()):
                         default_category_name = pref; break
                 if not default_category_name and relevant_categories:
                      sorted_relevant = sorted(relevant_categories.values(), key=lambda x: x.get('name',''))
                      default_category_name = sorted_relevant[0].get('name') if sorted_relevant else None

            if not relevant_categories:
                ttk.Label(category_buttons_frame, text=f"No suitable {tab_type} categories.",
                          foreground=theme_colors["disabled"], style=label_style).grid(row=0, column=0,
                                                                                       columnspan=cat_cols)
                self.selected_category_var.set("Other" if tab_type == "expense" else "Other Income")
            else:
                current_selection = self.selected_category_var.get()
                valid_names = [details['name'] for details in relevant_categories.values()]

                if not current_selection or current_selection not in valid_names:
                    final_default = default_category_name if default_category_name in valid_names else (valid_names[0] if valid_names else ("Other" if tab_type == "expense" else "Other Income"))
                    self.selected_category_var.set(final_default)

                cat_row, cat_col = 0, 0
                sorted_cat_items = sorted(relevant_categories.items(), key=lambda item: item[1].get('name', '').lower())
                for cat_id, details in sorted_cat_items:
                    cat_name = details['name']
                    cat_icon = details.get('icon', '')
                    cat_text = f"{cat_icon} {cat_name}".strip()
                    rb = ttk.Radiobutton(category_buttons_frame, text=cat_text, variable=self.selected_category_var,
                                         value=cat_name, style=cat_button_style, width=14)
                    rb.grid(row=cat_row, column=cat_col, sticky="w", padx=2, pady=2)
                    cat_col += 1
                    if cat_col >= cat_cols: cat_col = 0; cat_row += 1
            row_num += 1

            # Budget and Goal Sections
            if tab_type == "expense":
                ttk.Label(tab_frame, text="Deduct from Budget:", font=FONT_BOLD, style=label_style).grid(row=row_num,
                                                                                                         column=0,
                                                                                                         sticky="w",
                                                                                                         padx=(0, 10),
                                                                                                         pady=(10, 2))
                budget_data = app_data.get("budgets", {})
                budget_names = sorted(
                    [b.get('name', '') for b in budget_data.values()
                     if isinstance(b, dict) and 'name' in b and b['name'] is not None],
                    key=lambda x: str(x).lower()
                )
                budget_choices = ["None"] + budget_names
                budget_combo = ttk.Combobox(tab_frame, textvariable=self.budget_var, values=budget_choices, state='readonly', font=FONT_NORMAL, style=combo_style)
                budget_combo.grid(row=row_num, column=1, sticky="ew", pady=(10, 5))
                if self.budget_var.get() not in budget_choices: self.budget_var.set("None")
                budget_combo.set(self.budget_var.get())
                row_num += 1

                ttk.Label(tab_frame, text="Add to Goal:", font=FONT_BOLD, style=label_style).grid(row=row_num, column=0,
                                                                                                  sticky="w",
                                                                                                  padx=(0, 10),
                                                                                                  pady=(10, 2))
                goal_data = app_data.get("goals", {})
                goal_names = sorted(
                    [g.get('name', '') for g in goal_data.values()
                     if isinstance(g, dict) and 'name' in g and g['name'] is not None],
                    key=lambda x: str(x).lower()
                )
                goal_choices = ["None"] + goal_names
                goal_combo = ttk.Combobox(tab_frame, textvariable=self.goal_var, values=goal_choices, state='readonly', font=FONT_NORMAL, style=combo_style)
                goal_combo.grid(row=row_num, column=1, sticky="ew", pady=(10, 5))
                if self.goal_var.get() not in goal_choices: self.goal_var.set("None")
                goal_combo.set(self.goal_var.get())
                row_num += 1

        tk.Frame(tab_frame, height=10, bg=theme_colors["dialog_card"]).grid(row=row_num, column=0, columnspan=2)

    def on_tab_change(self, event=None):
        """Adjusts category selection based on the active tab (Expense/Income/Transfer)."""
        try:
            current_tab_index = self.notebook.index(self.notebook.select())
            all_categories = app_data.get("categories", {}); default_to_set = None
            if not all_categories: return

            current_selection = self.selected_category_var.get()

            if current_tab_index == 0:  # Expense
                 relevant_type = 'expense'
                 valid_names = [v['name'] for v in all_categories.values() if isinstance(v,dict) and v.get('type') == relevant_type]
                 if not current_selection or current_selection not in valid_names:
                      prefs = ["Groceries", "Fuel", "Rent/Mortgage", "Utilities"]
                      for pref in prefs:
                          if pref in valid_names: default_to_set = pref; break
                      if not default_to_set and valid_names:
                          default_to_set = sorted(valid_names)[0]
                      elif not default_to_set:
                           default_to_set = "Other"

            elif current_tab_index == 1:  # Income
                 relevant_type = 'income'
                 valid_names = [v['name'] for v in all_categories.values() if isinstance(v,dict) and v.get('type') == relevant_type]
                 if not current_selection or current_selection not in valid_names:
                     prefs = ["Salary", "Freelance", "Business"]
                     for pref in prefs:
                         if pref in valid_names: default_to_set = pref; break
                     if not default_to_set and valid_names:
                          default_to_set = sorted(valid_names)[0]
                     elif not default_to_set:
                          default_to_set = "Other Income"

            elif current_tab_index == 2:  # Transfer
                 transfer_cat = next((v for k,v in all_categories.items() if isinstance(v, dict) and k == 'cat_transfer'), None)
                 transfer_cat_name = transfer_cat['name'] if transfer_cat else 'Transfer'
                 self.selected_category_var.set(transfer_cat_name); default_to_set = None

            if default_to_set is not None:
                self.selected_category_var.set(default_to_set)

            # Reset budget/goal selections when switching away from expense tab
            if current_tab_index != 0:
                 self.budget_var.set("None")
                 self.goal_var.set("None")

        except (tk.TclError, AttributeError, IndexError) as e: logging.warning(f"Error during tab change handling: {e}")
        except Exception as e: logging.exception("Unexpected error during tab change")

    def add_transaction(self):
        """Validates input, creates a new transaction(s), and updates wallet balances."""
        try:
            # Input Validation
            amount_str = self.amount_var.get().replace(",", "").strip()
            if not amount_str: raise ValueError("Amount cannot be empty.")
            try:
                amount = abs(float(amount_str))
                if amount <= 0: raise ValueError("Amount must be positive.")
            except ValueError:
                raise ValueError("Invalid amount entered.")

            title = self.title_var.get().strip()
            if not title: raise ValueError("Title cannot be empty.")

            wallet_name = self.wallet_var.get()
            if not wallet_name: raise ValueError("Please select a wallet.")

            user_wallets = app_data.get("wallets", {})
            valid_wallet_names = [w['name'] for w in user_wallets.values() if isinstance(w, dict)]
            if wallet_name not in valid_wallet_names: raise ValueError(f"Wallet '{wallet_name}' is invalid.")

            date_str = self.date_var.get()
            time_str = self.time_var.get()
            try:
                datetime.datetime.strptime(date_str, "%Y-%m-%d")
                datetime.datetime.strptime(time_str, "%H:%M")
            except ValueError:
                 raise ValueError("Invalid date or time format (Use YYYY-MM-DD and HH:MM).")
            timestamp = f"{date_str} {time_str}"

            current_tab_index = self.notebook.index(self.notebook.select())
            category = ""
            final_amount = 0.0
            tx_type = ""
            log_message = ""
            linked_budget_name = None
            linked_goal_name = None

            if current_tab_index == 0:  # Expense
                final_amount = -amount
                tx_type = "expense"
                category = self.selected_category_var.get()
                if not category: raise ValueError("Please select a category.")

                cat_details = next((v for k,v in app_data.get("categories", {}).items() if isinstance(v, dict) and v.get('name') == category), None)
                if not cat_details or cat_details.get('type') != 'expense':
                    raise ValueError(f"Invalid category '{category}' selected for an expense.")

                budget_selection = self.budget_var.get()
                if budget_selection != "None":
                    if budget_selection in [b['name'] for b in app_data.get("budgets", {}).values() if isinstance(b, dict)]:
                        linked_budget_name = budget_selection
                    else:
                        logging.warning(f"Selected budget '{budget_selection}' no longer exists. Ignoring link.")
                        self.budget_var.set("None")

                goal_selection = self.goal_var.get()
                if goal_selection != "None":
                    if goal_selection in [g['name'] for g in app_data.get("goals", {}).values() if isinstance(g, dict)]:
                        linked_goal_name = goal_selection
                    else:
                        logging.warning(f"Selected goal '{goal_selection}' no longer exists. Ignoring link.")
                        self.goal_var.set("None")

                log_message = f"Added Expense: {title} ({format_currency(final_amount)}) to {category}"
                if linked_budget_name: log_message += f" (Budget: {linked_budget_name})"
                if linked_goal_name: log_message += f" (Goal: {linked_goal_name})"

            elif current_tab_index == 1:  # Income
                final_amount = amount
                tx_type = "income"
                category = self.selected_category_var.get()
                if not category: raise ValueError("Please select income source.")

                cat_details = next((v for k,v in app_data.get("categories", {}).items() if isinstance(v, dict) and v.get('name') == category), None)
                if not cat_details or cat_details.get('type') != 'income':
                     raise ValueError(f"Invalid category '{category}' selected for income.")

                linked_goal_name = None

                log_message = f"Added Income: {title} ({format_currency(final_amount)}) from {category}"

            elif current_tab_index == 2:  # Transfer
                to_wallet_name = self.to_wallet_var.get()
                if not to_wallet_name: raise ValueError("Please select 'To Wallet'.")
                if wallet_name == to_wallet_name: raise ValueError("'From' and 'To' wallets must be different.")
                if to_wallet_name not in valid_wallet_names: raise ValueError(f"'To Wallet' ({to_wallet_name}) is invalid.")

                all_categories = app_data.get("categories", {})
                transfer_cat = next((v for k,v in all_categories.items() if isinstance(v, dict) and k == 'cat_transfer'), None)
                transfer_cat_name = transfer_cat['name'] if transfer_cat else 'Transfer'


                tx_out = {"date": date_str, "time": time_str, "timestamp": timestamp,
                          "title": f"Transfer to {to_wallet_name}", "wallet": wallet_name, "amount": -amount,
                          "category": transfer_cat_name, "type": "transfer_out", "from_account": wallet_name,
                          "to_account": to_wallet_name, "linked_budget": None, "linked_goal": None}
                tx_in = {"date": date_str, "time": time_str, "timestamp": timestamp,
                         "title": f"Transfer from {wallet_name}", "wallet": to_wallet_name, "amount": amount,
                         "category": transfer_cat_name, "type": "transfer_in", "from_account": wallet_name,
                         "to_account": to_wallet_name, "linked_budget": None, "linked_goal": None}

                if not isinstance(app_data.get("transactions"), list): app_data["transactions"] = []
                app_data["transactions"].extend([tx_out, tx_in])
                log_activity(f"Added Transfer: {format_currency(amount)} from {wallet_name} to {to_wallet_name}")
                self.update_wallet_balance(wallet_name, -amount)
                self.update_wallet_balance(to_wallet_name, amount)
                self.app.refresh_current_page()
                messagebox.showinfo("Success", "Transfer added!", parent=self)
                self.destroy()
                return

            # Create and Save Single Transaction (Expense/Income)
            new_transaction = {
                "date": date_str,
                "time": time_str,
                "timestamp": timestamp,
                "title": title,
                "wallet": wallet_name,
                "amount": final_amount,
                "category": category,
                "type": tx_type,
                "from_account": None,
                "to_account": None,
                "linked_budget": linked_budget_name,
                "linked_goal": linked_goal_name
            }

            if not isinstance(app_data.get("transactions"), list): app_data["transactions"] = []
            app_data["transactions"].append(new_transaction)

            log_activity(log_message)
            self.update_wallet_balance(wallet_name, final_amount)

            # Refresh and Close
            self.app.refresh_current_page()
            messagebox.showinfo("Success", f"{tx_type.capitalize()} added!", parent=self)
            self.destroy()

        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e), parent=self)
        except Exception as e:
            logging.exception("Error adding transaction")
            messagebox.showerror("Error", f"Could not add transaction.\n{e}", parent=self)

    def update_wallet_balance(self, wallet_name, amount_change):
        """Updates the balance of a specified wallet."""
        wallets_dict = app_data.get("wallets", {})
        if not isinstance(wallets_dict, dict): logging.error("Wallets data not a dict."); return
        wallet_id_to_update = None
        for w_id, w_details in wallets_dict.items():
            if isinstance(w_details, dict) and w_details.get("name") == wallet_name: wallet_id_to_update = w_id; break
        try:
            current_balance = float(wallets_dict[wallet_id_to_update].get('balance', 0.0))
            new_balance = current_balance + amount_change
            wallets_dict[wallet_id_to_update]["balance"] = new_balance
            logging.info(f"Updated balance for '{wallet_name}' to {format_currency(new_balance)}")
        except (ValueError, TypeError) as e:
            logging.error(f"Error converting balance for {wallet_name}: {e}")
        except Exception as e:
            logging.exception(f"Error updating balance for {wallet_name}")

# --- Simple Entry Dialog (Used for Add/Edit Items) ---
class SimpleEntryDialog(tk.Toplevel):
    def __init__(self, parent, title, fields_config):
        super().__init__(parent)
        self.transient(parent); self.parent = parent; self.title(title)
        self.fields_config = fields_config; self.result = None; self.entries = {}; self.vars = {}
        dialog_theme = THEME_DARK.copy()
        try:
            app_instance = parent.app if hasattr(parent, 'app') else parent
            if hasattr(app_instance, 'current_theme'): dialog_theme = THEME_LIGHT.copy() if app_instance.current_theme == 'light' else THEME_DARK.copy()
        except AttributeError: pass
        dialog_bg=dialog_theme.get("dialog_bg", dialog_theme["background"]); dialog_fg=dialog_theme.get("dialog_fg", dialog_theme["foreground"])
        dialog_card=dialog_theme.get("dialog_card", dialog_theme["card"]); dialog_accent=dialog_theme.get("accent")
        self.configure(bg=dialog_bg, padx=15, pady=15)
        main_frame = tk.Frame(self, bg=dialog_bg); main_frame.pack(expand=True, fill="both"); main_frame.grid_columnconfigure(1, weight=1)
        dialog_style = ttk.Style(self); dialog_style.theme_use('clam')
        dialog_style.configure('Dialog.TLabel', background=dialog_bg, foreground=dialog_fg)
        dialog_style.configure('Dialog.TEntry', fieldbackground=dialog_card, foreground=dialog_fg, insertcolor=dialog_fg, borderwidth=0, padding=3); dialog_style.map('Dialog.TEntry', fieldbackground=[('focus', dialog_card)])
        dialog_style.configure('Dialog.TCombobox', fieldbackground=dialog_card, foreground=dialog_fg, selectbackground=dialog_card, selectforeground=dialog_fg, arrowcolor=dialog_fg, borderwidth=0, padding=3); dialog_style.map('Dialog.TCombobox', fieldbackground=[('readonly', dialog_card)])
        dialog_style.configure('Dialog.TCheckbutton', background=dialog_bg, foreground=dialog_fg); dialog_style.map('Dialog.TCheckbutton', indicatorcolor=[('selected', dialog_accent)])
        dialog_style.configure('Dialog.TButton', background=dialog_accent, foreground=dialog_theme.get("button_fg", "#ffffff"), font=FONT_BOLD); dialog_style.map('Dialog.TButton', background=[('active', dialog_theme.get("accent_darker", dialog_accent))])
        dialog_style.configure('Cancel.Dialog.TButton', background=dialog_theme.get("disabled", "#555"), foreground=dialog_fg); dialog_style.map('Cancel.Dialog.TButton', background=[('active', dialog_theme.get("red", "#AA0000"))], foreground=[('active', dialog_theme.get("button_fg", "#ffffff"))])

        row_num = 0
        for name, config in fields_config.items():
            label_text = config.get("label", name.replace("_", " ").title() + ":"); field_type = config.get("type", "text")
            initial_value = config.get("initial", None); required = config.get("required", True); label_suffix = " *" if required else ""
            lbl = ttk.Label(main_frame, text=label_text + label_suffix, style='Dialog.TLabel'); lbl.grid(row=row_num, column=0, sticky="w", padx=(0, 10), pady=5)
            var = None; entry_widget = None; widget_parent = main_frame
            if field_type == "boolean":
                var = tk.BooleanVar(value=bool(initial_value)); entry_widget = ttk.Checkbutton(widget_parent, variable=var, style='Dialog.TCheckbutton')
                entry_widget.grid(row=row_num, column=1, sticky="w", pady=5)
            elif field_type == "combo":
                var = tk.StringVar(); values = config.get("values", []); entry_widget = ttk.Combobox(widget_parent, textvariable=var, values=values, state='readonly', font=FONT_NORMAL, style='Dialog.TCombobox')
                if initial_value is not None: var.set(str(initial_value))
                current_val = var.get()
                if values and current_val in values: entry_widget.set(current_val)
                elif values: entry_widget.current(0)
                elif not values: entry_widget.set(" (No options) "); entry_widget.config(state='disabled')
                entry_widget.grid(row=row_num, column=1, sticky="ew", pady=5)
            elif field_type == "date":
                 date_str = str(initial_value) if initial_value else datetime.date.today().strftime("%Y-%m-%d")
                 try: datetime.datetime.strptime(date_str, "%Y-%m-%d")
                 except (ValueError, TypeError): date_str = datetime.date.today().strftime("%Y-%m-%d")
                 var = tk.StringVar(value=date_str); entry_widget = ttk.Entry(widget_parent, textvariable=var, font=FONT_NORMAL, style='Dialog.TEntry', width=12)
                 entry_widget.grid(row=row_num, column=1, sticky="w", pady=5)
            else:
                 initial_str = str(initial_value) if initial_value is not None else ""; var = tk.StringVar(value=initial_str)
                 justify = tk.RIGHT if field_type in ["number", "currency"] else tk.LEFT
                 entry_widget = ttk.Entry(widget_parent, textvariable=var, font=FONT_NORMAL, justify=justify, style='Dialog.TEntry')
                 entry_widget.grid(row=row_num, column=1, sticky="ew", pady=5)
            if entry_widget: self.entries[name] = entry_widget
            if var: self.vars[name] = var
            row_num += 1

        button_frame = tk.Frame(main_frame, bg=dialog_bg); button_frame.grid(row=row_num, column=0, columnspan=2, sticky="e", pady=(15, 0))
        ok_button = ttk.Button(button_frame, text="OK", command=self.on_ok, style="Dialog.TButton"); ok_button.pack(side=tk.RIGHT, padx=(5, 0))
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.on_cancel, style="Cancel.Dialog.TButton"); cancel_button.pack(side=tk.RIGHT)

        self.bind("<Return>", lambda e: self.on_ok()); self.bind("<Escape>", lambda e: self.on_cancel())
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.after(10, self._set_initial_focus)
        self._center_dialog()
        self.grab_set(); self.wait_window(self)

    def _set_initial_focus(self):
        """Sets focus to the first focusable widget in the dialog."""
        first_focusable = None
        for name in self.fields_config.keys():
            widget = self.entries.get(name)
            if widget and widget.winfo_exists() and 'disabled' not in widget.state(): first_focusable = widget; break
        if first_focusable:
            first_focusable.focus_set()
            if isinstance(first_focusable, (ttk.Entry, ttk.Combobox)): first_focusable.select_range(0, tk.END)

    def _center_dialog(self):
        """Centers the dialog window relative to its parent."""
        try:
            self.update_idletasks(); parent_x, parent_y = self.parent.winfo_rootx(), self.parent.winfo_rooty()
            parent_w, parent_h = self.parent.winfo_width(), self.parent.winfo_height(); dialog_w, dialog_h = self.winfo_width(), self.winfo_height()
            x = parent_x + (parent_w // 2) - (dialog_w // 2); y = parent_y + (parent_h // 3) - (dialog_h // 2)
            screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
            x = max(0, min(x, screen_w - dialog_w)); y = max(0, min(y, screen_h - dialog_h))
            self.geometry(f"+{x}+{y}")
        except Exception as e: logging.warning(f"Could not center SimpleEntryDialog: {e}")

    def on_ok(self):
        """Collects input values and sets the result, then destroys dialog."""
        self.result = {}
        for name, var in self.vars.items():
            try: self.result[name] = var.get()
            except Exception as e: logging.error(f"Error getting value for '{name}': {e}"); self.result[name] = None
        self.destroy()

    def on_cancel(self):
        """Sets result to None and destroys dialog."""
        self.result = None; self.destroy()


# --- Main Execution Logic ---
def launch_main_app(user_id):
    """Launches the main application window for the selected user."""
    logging.info(f"Launching Main Application for user_id: {user_id}...")
    app = None
    try:
        app = ExpenseWiseApp(user_id)
        app.mainloop()
        logging.info(f"Main application mainloop finished for user {user_id}.")
        # Check if full exit was requested
        if app and getattr(app, '_full_exit_requested', False):
            logging.info("Full exit requested by application.")
            return False
        else:
            logging.info("Returning to Accounts Page.")
            return True
    except Exception as e:
         logging.exception(f"Critical error running main application for user {user_id}")
         messagebox.showerror("Application Error", f"A critical error occurred:\n{e}")
         return True

if __name__ == "__main__":
    ensure_data_dir()
    logging.info("--- ExpenseWise Application Starting ---")
    continue_running = True
    while continue_running:
        logging.info("Showing Accounts Page...")
        selected_user_id = None
        app_instance_closed = False
        try:
            accounts_page = AccountsPage()
            accounts_page.mainloop()
            selected_user_id = getattr(accounts_page, 'selected_user_id', None)
            if selected_user_id is None:
                 logging.info("Accounts Page closed without user selection.")
                 continue_running = False

        except Exception as e:
             logging.exception("Error during Accounts Page execution")
             messagebox.showerror("Startup Error", f"Error on accounts page:\n{e}")
             continue_running = False

        if selected_user_id and continue_running:
            continue_running = launch_main_app(selected_user_id)

    logging.info("--- ExpenseWise Application Finished ---")
