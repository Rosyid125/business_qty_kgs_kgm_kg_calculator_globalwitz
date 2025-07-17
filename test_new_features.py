#!/usr/bin/env python3
"""
Test script untuk fitur baru: 
1. Ekstrak nilai pertama dari string terpisah koma
2. Hapus operator dari nilai GSM
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from business_quantity_converter import BusinessQuantityConverter
import tkinter as tk
import pandas as pd

def test_extract_first_value():
    """Test fungsi extract_first_value"""
    print("=== Testing extract_first_value ===")
    
    # Create dummy converter instance
    root = tk.Tk()
    root.withdraw()  # Hide the window
    converter = BusinessQuantityConverter(root)
    
    test_cases = [
        ("67.80,50,40", "67.80"),
        ("100,200,300", "100"),
        ("50", "50"),
        ("", None),
        (None, None),
        ("  150,  250  ", "150"),
        ("abc,def,ghi", "abc"),
        ("1.5,2.5,3.5", "1.5"),
    ]
    
    for input_val, expected in test_cases:
        result = converter.extract_first_value(input_val)
        status = "✅" if result == expected else "❌"
        print(f"{status} Input: '{input_val}' → Output: '{result}' (Expected: '{expected}')")
    
    root.destroy()

def test_clean_gsm_operators():
    """Test fungsi clean_gsm_operators"""
    print("\n=== Testing clean_gsm_operators ===")
    
    # Create dummy converter instance
    root = tk.Tk()
    root.withdraw()  # Hide the window
    converter = BusinessQuantityConverter(root)
    
    test_cases = [
        (">40", "40"),
        ("<50", "50"),
        (">=25", "25"),
        ("<=30", "30"),
        ("40", "40"),
        ("> 40", "40"),
        ("<= 30", "30"),
        (">= 25", "25"),
        ("", None),
        (None, None),
        ("abc", "abc"),  # non-numeric should still work
        ("  >=  100  ", "100"),
    ]
    
    for input_val, expected in test_cases:
        result = converter.clean_gsm_operators(input_val)
        status = "✅" if result == expected else "❌"
        print(f"{status} Input: '{input_val}' → Output: '{result}' (Expected: '{expected}')")
    
    root.destroy()

def test_combined_processing():
    """Test kombinasi extract_first_value dan clean_gsm_operators"""
    print("\n=== Testing Combined Processing ===")
    
    # Create dummy converter instance
    root = tk.Tk()
    root.withdraw()  # Hide the window
    converter = BusinessQuantityConverter(root)
    
    test_cases = [
        # (input, expected_after_extract_first, expected_after_clean_operators)
        ("<=30,35,40", "<=30", "30"),
        (">40,50,60", ">40", "40"),
        (">=25,30", ">=25", "25"),
        ("<50,60", "<50", "50"),
        ("67.80,50,40", "67.80", "67.80"),  # no operators
        ("100,200", "100", "100"),  # no operators
    ]
    
    for input_val, expected_first, expected_final in test_cases:
        # Step 1: Extract first value
        first_val = converter.extract_first_value(input_val)
        # Step 2: Clean operators
        final_val = converter.clean_gsm_operators(first_val)
        
        status = "✅" if final_val == expected_final else "❌"
        print(f"{status} Input: '{input_val}' → First: '{first_val}' → Final: '{final_val}' (Expected: '{expected_final}')")
    
    root.destroy()

def test_numeric_conversion():
    """Test konversi ke numeric setelah preprocessing"""
    print("\n=== Testing Numeric Conversion ===")
    
    # Create dummy converter instance
    root = tk.Tk()
    root.withdraw()  # Hide the window
    converter = BusinessQuantityConverter(root)
    
    test_cases = [
        ("67.80,50,40", 67.80),
        ("<=30,35", 30.0),
        (">40,50", 40.0),
        (">=25.5,30", 25.5),
        ("100", 100.0),
    ]
    
    for input_val, expected_numeric in test_cases:
        # Process like in the main function
        first_val = converter.extract_first_value(input_val)
        cleaned_val = converter.clean_gsm_operators(first_val)
        numeric_val = pd.to_numeric(cleaned_val, errors='coerce') or 0
        
        status = "✅" if abs(numeric_val - expected_numeric) < 0.001 else "❌"
        print(f"{status} Input: '{input_val}' → Processed: '{cleaned_val}' → Numeric: {numeric_val} (Expected: {expected_numeric})")
    
    root.destroy()

if __name__ == "__main__":
    print("🧪 Testing New Features for GSM and Width Processing\n")
    
    test_extract_first_value()
    test_clean_gsm_operators()
    test_combined_processing()
    test_numeric_conversion()
    
    print("\n🎉 All tests completed!")
