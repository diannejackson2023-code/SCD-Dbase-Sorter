import logging
import os
from datetime import datetime
import json

LOG_DIR = "/home/team/shared/SCD_Dbase_Sorter/data/logs"
AUDIT_LOG_FILE = os.path.join(LOG_DIR, "audit_log.jsonl")

class AuditLogger:
    def __init__(self):
        os.makedirs(LOG_DIR, exist_ok=True)
        self.logger = logging.getLogger("audit_logger")
        self.logger.setLevel(logging.INFO)
        
        # Avoid adding multiple handlers if already initialized
        if not self.logger.handlers:
            handler = logging.FileHandler(AUDIT_LOG_FILE)
            self.logger.addHandler(handler)

    def log_action(self, action, user="system", details=None):
        """
        Logs an action in JSONL format.
        """
        log_entry = {
            "timestamp": datetime.now().isoformat(),
            "action": action,
            "user": user,
            "details": details or {}
        }
        self.logger.info(json.dumps(log_entry))

# Global instance
audit_logger = AuditLogger()
