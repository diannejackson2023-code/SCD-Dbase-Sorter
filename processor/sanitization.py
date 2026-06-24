import html

def sanitize_value(val):
    """
    Sanitizes a single value (string or other).
    """
    if not isinstance(val, str):
        return val
    
    # 1. Prevent Excel Formula Injection
    # Common triggers for Excel formulas are: =, +, -, @
    if val.startswith(('=', '+', '-', '@')):
        # Add a single quote to the front to treat it as text in Excel
        val = "'" + val
        
    # 2. Prevent XSS by escaping HTML tags
    val = html.escape(val)
    
    return val

def sanitize_dataframe(df):
    """
    Applies sanitization to all string columns in a DataFrame.
    """
    # Use applymap for element-wise sanitization
    # Note: applymap is deprecated in newer pandas, use apply(lambda x: x.map(sanitize_value)) or similar
    # For compatibility across versions:
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].apply(sanitize_value)
    return df
