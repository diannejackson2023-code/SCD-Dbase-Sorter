import os
from cryptography.fernet import Fernet

KEY_FILE = "/home/team/shared/SCD_Dbase_Sorter/data/config/.master.key"

def ensure_key():
    """Ensures a master key exists."""
    if not os.path.exists(KEY_FILE):
        os.makedirs(os.path.dirname(KEY_FILE), exist_ok=True)
        key = Fernet.generate_key()
        with open(KEY_FILE, "wb") as f:
            f.write(key)
    else:
        with open(KEY_FILE, "rb") as f:
            key = f.read()
    return key

def encrypt_file(file_path):
    """Encrypts a file in place."""
    if not os.path.exists(file_path):
        return
    
    key = ensure_key()
    fernet = Fernet(key)
    
    with open(file_path, "rb") as f:
        data = f.read()
    
    encrypted_data = fernet.encrypt(data)
    
    with open(file_path, "wb") as f:
        f.write(encrypted_data)

def decrypt_file_to_memory(file_path):
    """Decrypts a file and returns the data in memory."""
    if not os.path.exists(file_path):
        return None
    
    key = ensure_key()
    fernet = Fernet(key)
    
    with open(file_path, "rb") as f:
        encrypted_data = f.read()
    
    try:
        decrypted_data = fernet.decrypt(encrypted_data)
        return decrypted_data
    except Exception:
        # If decryption fails, maybe it's not encrypted? 
        # For safety, return the original data or handle error
        return encrypted_data
