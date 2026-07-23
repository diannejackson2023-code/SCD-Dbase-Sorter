import logging
import os
from datetime import datetime
import json

try:
    from .config import LOGS_DIR as LOG_DIR, AUDIT_LOG_PATH as AUDIT_LOG_FILE
except ImportError:
    from config import LOGS_DIR as LOG_DIR, AUDIT_LOG_PATH as AUDIT_LOG_FILE

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
