#!/usr/bin/env python3
"""
Test script untuk memverifikasi rumus SQF yang sudah diperbaiki
"""

def test_sqf_conversion():
    """Test SQF conversion dengan data yang diberikan user"""
    
    # Data dari user
    business_quantity = 955027.95  # SQF
    gsm = 28
    
    # Rumus yang sudah diperbaiki: Business Quantity × 0.092903 × GSM ÷ 1000
    result = business_quantity * 0.092903 * gsm / 1000
    
    print("=== Test SQF Conversion ===")
    print(f"Business Quantity: {business_quantity} SQF")
    print(f"GSM: {gsm}")
    print(f"Formula: {business_quantity} × 0.092903 × {gsm} ÷ 1000")
    print(f"Calculation: {business_quantity} × 0.092903 × {gsm} ÷ 1000")
    print(f"Step 1: {business_quantity} × 0.092903 = {business_quantity * 0.092903}")
    print(f"Step 2: {business_quantity * 0.092903} × {gsm} = {business_quantity * 0.092903 * gsm}")
    print(f"Step 3: {business_quantity * 0.092903 * gsm} ÷ 1000 = {result}")
    print(f"Result: {result:.2f} KG")
    print()
    
    # Verifikasi dengan perhitungan manual user
    expected_result = 2484.33
    print(f"Expected result (manual calculation): {expected_result} KG")
    print(f"Difference: {abs(result - expected_result):.2f} KG")
    
    if abs(result - expected_result) < 1:  # tolerance 1 KG
        print("✅ PASSED: Formula is correct!")
    else:
        print("❌ FAILED: Formula needs adjustment")
    
    return result

if __name__ == "__main__":
    test_sqf_conversion()
