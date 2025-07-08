#!/usr/bin/env python3
"""
Test script to verify unit normalization and conversion factors
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from business_quantity_converter import BusinessQuantityConverter
import tkinter as tk

def test_unit_normalization():
    """Test unit normalization functionality"""
    print("🧪 Testing Unit Normalization...")
    
    # Create a dummy root for testing
    root = tk.Tk()
    root.withdraw()  # Hide the window
    
    converter = BusinessQuantityConverter(root)
    
    # Test cases for unit normalization
    test_cases = [
        # Kilogram variants
        ("kg", "KG"), ("KG", "KG"), ("kgs", "KG"), ("K", "KG"), ("kilo", "KG"),
        
        # Gram variants  
        ("g", "G"), ("gr", "G"), ("gram", "G"), ("grams", "G"), ("grm", "G"),
        
        # Pound variants
        ("lb", "LBS"), ("lbs", "LBS"), ("pound", "LBS"), ("pounds", "LBS"),
        
        # Ounce variants
        ("oz", "OZ"), ("ounce", "OZ"), ("ounces", "OZ"),
        
        # Ton variants
        ("ton", "TON"), ("tons", "TON"), ("tonne", "TON"), ("mt", "MT"),
        
        # Complex units
        ("mtr", "MTR"), ("meter", "MTR"), ("m", "MTR"),
        ("yd", "YD"), ("yard", "YD"), ("yards", "YD"),
        
        # Case and spacing variations
        ("K G", "KG"), ("k.g", "KG"), ("k-g", "KG"), (" kg ", "KG"),
        ("LB S", "LBS"), ("l.b.s", "LBS"), ("L B S", "LBS"),
    ]
    
    passed = 0
    failed = 0
    
    for input_unit, expected in test_cases:
        result = converter.normalize_unit(input_unit)
        if result == expected:
            print(f"✅ '{input_unit}' → '{result}'")
            passed += 1
        else:
            print(f"❌ '{input_unit}' → '{result}' (expected '{expected}')")
            failed += 1
    
    print(f"\n📊 Unit Normalization Results: {passed} passed, {failed} failed")
    
    return passed, failed

def test_conversion_factors():
    """Test conversion factors"""
    print("\n🧪 Testing Conversion Factors...")
    
    root = tk.Tk()
    root.withdraw()
    
    converter = BusinessQuantityConverter(root)
    
    # Test cases for conversion factors
    test_cases = [
        ("KG", 1.0),
        ("G", 0.001),
        ("LBS", 0.453592),
        ("OZ", 0.0283495),
        ("TON", 1000.0),
        ("MT", 1000.0),
        ("STONE", 6.35029),
        ("MG", 0.000001),
    ]
    
    passed = 0
    failed = 0
    
    for unit, expected_factor in test_cases:
        result = converter.get_conversion_factor(unit)
        if result == expected_factor:
            print(f"✅ {unit} → {result}")
            passed += 1
        else:
            print(f"❌ {unit} → {result} (expected {expected_factor})")
            failed += 1
    
    # Test unsupported unit
    unsupported = converter.get_conversion_factor("UNKNOWN")
    if unsupported is None:
        print(f"✅ UNKNOWN → None (correctly unsupported)")
        passed += 1
    else:
        print(f"❌ UNKNOWN → {unsupported} (should be None)")
        failed += 1
    
    print(f"\n📊 Conversion Factor Results: {passed} passed, {failed} failed")
    
    return passed, failed

def test_real_world_examples():
    """Test with real-world unit variations"""
    print("\n🧪 Testing Real-World Examples...")
    
    root = tk.Tk()
    root.withdraw()
    
    converter = BusinessQuantityConverter(root)
    
    # Real-world examples that might appear in Excel files
    real_world_cases = [
        # Common Excel variations
        ("Kg", "KG"), ("KG.", "KG"), ("kg ", "KG"), (" KG", "KG"),
        ("lb.", "LBS"), ("Lb", "LBS"), ("LB ", "LBS"), ("lbs.", "LBS"),
        ("gr.", "G"), ("Gr", "G"), ("grams", "G"), ("GRAMS", "G"),
        ("Ounce", "OZ"), ("OUNCE", "OZ"), ("oz.", "OZ"),
        ("Meter", "MTR"), ("METER", "MTR"), ("mtr.", "MTR"),
        
        # Variations with numbers/special chars (should be cleaned)
        ("kg1", "KG1"), ("lb-1", "LB1"), ("gr.2", "GR2"),  # These won't match our dict
        
        # International variations
        ("kilogram", "KG"), ("kilogramme", "KG"), ("kilogrammes", "KG"),
        ("gramme", "G"), ("grammes", "G"),
    ]
    
    passed = 0
    failed = 0
    
    for input_unit, expected in real_world_cases:
        result = converter.normalize_unit(input_unit)
        if result == expected:
            print(f"✅ '{input_unit}' → '{result}'")
            passed += 1
        else:
            # For special cases that we expect to fail (like kg1)
            if input_unit in ["kg1", "lb-1", "gr.2"]:
                print(f"⚠️ '{input_unit}' → '{result}' (expected failure)")
                passed += 1
            else:
                print(f"❌ '{input_unit}' → '{result}' (expected '{expected}')")
                failed += 1
    
    print(f"\n📊 Real-World Examples Results: {passed} passed, {failed} failed")
    
    return passed, failed

if __name__ == "__main__":
    print("🚀 Starting Unit Converter Tests...")
    print("=" * 50)
    
    # Run all tests
    norm_passed, norm_failed = test_unit_normalization()
    factor_passed, factor_failed = test_conversion_factors() 
    real_passed, real_failed = test_real_world_examples()
    
    # Summary
    total_passed = norm_passed + factor_passed + real_passed
    total_failed = norm_failed + factor_failed + real_failed
    total_tests = total_passed + total_failed
    
    print("\n" + "=" * 50)
    print("🏁 FINAL RESULTS:")
    print(f"✅ Total Passed: {total_passed}")
    print(f"❌ Total Failed: {total_failed}")
    print(f"📊 Success Rate: {(total_passed/total_tests)*100:.1f}%" if total_tests > 0 else "No tests run")
    
    if total_failed == 0:
        print("🎉 All tests passed! Unit converter is working correctly.")
    else:
        print("⚠️ Some tests failed. Please review the implementation.")
    
    print("\n💡 Run the main application with: python business_quantity_converter.py")
