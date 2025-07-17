#!/usr/bin/env python3
"""
Create sample Excel file with test data for the new features
"""

import pandas as pd
import os

# Create sample data with comma-separated values and operators
data = {
    'Unit of Weight': ['KG', 'MTR', 'MTK', 'KG', 'MTR', 'G', 'LBS'],
    'Business Quantity': [100, 50, 25, 200, 75, 1000, 10],
    'Unit Price (USD)': [15.50, 25.00, 30.00, 12.75, 40.00, 5.00, 20.00],
    'Width': ['150,200,250', '120', '180,220', '100,150', '160,200,300', '140', '170,190'],
    'GSM': ['>200,250,300', '<=150,180', '>=180', '<100,120', '>300,350', '250,280,320', '200']
}

df = pd.DataFrame(data)

# Create input folder if it doesn't exist
input_folder = "./input"
os.makedirs(input_folder, exist_ok=True)

# Save to Excel file
output_file = os.path.join(input_folder, "test_data_new_features.xlsx")
df.to_excel(output_file, index=False, sheet_name="TestSheet")

print(f"✅ Sample data created: {output_file}")
print("\nSample data preview:")
print(df.to_string(index=False))

print("\n📋 Expected processing results:")
print("Width column (extract first value):")
for i, width in enumerate(df['Width']):
    first_val = width.split(',')[0] if ',' in str(width) else str(width)
    print(f"  Row {i+1}: '{width}' → '{first_val}'")

print("\nGSM column (extract first + remove operators):")
for i, gsm in enumerate(df['GSM']):
    first_val = str(gsm).split(',')[0] if ',' in str(gsm) else str(gsm)
    # Remove operators
    import re
    cleaned = re.sub(r'^(>=|<=|>|<)\s*', '', first_val)
    print(f"  Row {i+1}: '{gsm}' → first: '{first_val}' → cleaned: '{cleaned}'")
