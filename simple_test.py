#!/usr/bin/env python3
"""
Test sederhana untuk fitur baru
"""

import re
import pandas as pd

def extract_first_value(value_string):
    """Extract the first value from a comma-separated string"""
    if not value_string or pd.isna(value_string):
        return None
    
    # Convert to string and split by comma
    str_value = str(value_string).strip()
    if ',' in str_value:
        first_value = str_value.split(',')[0].strip()
    else:
        first_value = str_value
    
    return first_value if first_value else None

def clean_gsm_operators(gsm_value):
    """Remove operators (>, <, >=, <=) from GSM value and return the numeric value"""
    if not gsm_value or pd.isna(gsm_value):
        return None
    
    # Convert to string
    str_value = str(gsm_value).strip()
    
    # Remove operators: >=, <=, >, <
    # Pattern to match operators at the beginning of the string
    pattern = r'^(>=|<=|>|<)\s*'
    cleaned_value = re.sub(pattern, '', str_value)
    
    return cleaned_value if cleaned_value else None

# Test cases
print("Testing extract_first_value:")
test_cases_1 = [
    "67.80,50,40",
    "100,200,300", 
    ">40,50",
    "<=30,35",
    "50"
]

for case in test_cases_1:
    result = extract_first_value(case)
    print(f"  '{case}' → '{result}'")

print("\nTesting clean_gsm_operators:")
test_cases_2 = [
    ">40",
    "<50", 
    ">=25",
    "<=30",
    "40",
    "> 40",
    "<= 30"
]

for case in test_cases_2:
    result = clean_gsm_operators(case)
    print(f"  '{case}' → '{result}'")

print("\nTesting combined (GSM case):")
gsm_cases = [
    "<=30,35,40",
    ">40,50,60", 
    "67.80,50,40"
]

for case in gsm_cases:
    first = extract_first_value(case)
    cleaned = clean_gsm_operators(first)
    numeric = pd.to_numeric(cleaned, errors='coerce') or 0
    print(f"  '{case}' → first: '{first}' → cleaned: '{cleaned}' → numeric: {numeric}")

print("\nTesting width case:")
width_cases = [
    "150,200,300",
    "100",
    "50.5,75"
]

for case in width_cases:
    first = extract_first_value(case)
    numeric = pd.to_numeric(first, errors='coerce') or 0
    print(f"  '{case}' → first: '{first}' → numeric: {numeric}")
