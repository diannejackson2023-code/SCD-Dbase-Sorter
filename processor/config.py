import os

# Root of the project
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Data directories
DATA_DIR = os.path.join(BASE_DIR, "data")
CONFIG_DIR = os.path.join(DATA_DIR, "config")
MASTER_DIR = os.path.join(DATA_DIR, "master")
HOSPITALS_DIR = os.path.join(DATA_DIR, "hospitals")
LOGS_DIR = os.path.join(DATA_DIR, "logs")
STAGING_DIR = os.path.join(DATA_DIR, "staging")

# File paths
MASTER_DB_PATH = os.path.join(MASTER_DIR, "Master_Database.xlsx")
MASTER_KEY_PATH = os.path.join(CONFIG_DIR, ".master.key")
ALIASES_PATH = os.path.join(CONFIG_DIR, "aliases.json")
HOSPITAL_EMAILS_PATH = os.path.join(CONFIG_DIR, "hospital_emails.csv")
AUDIT_LOG_PATH = os.path.join(LOGS_DIR, "audit_log.jsonl")
DISCOVERY_REQUESTS_PATH = os.path.join(CONFIG_DIR, "discovery_requests.json")
QUEUE_FILE = os.path.join(STAGING_DIR, "queue.json")
