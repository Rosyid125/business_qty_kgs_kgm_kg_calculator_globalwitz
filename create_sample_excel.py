import pandas as pd
import os

# Sample data with various unit formats
sample_data = [
    # Direct conversions
    {"Unit of Weight": "kg", "Business Quantity": 100, "Unit Price (USD)": 50, "Width": 150, "GSM": 200, "Description": "Kilogram standard"},
    {"Unit of Weight": "KG", "Business Quantity": 50, "Unit Price (USD)": 30, "Width": 120, "GSM": 180, "Description": "Kilogram uppercase"},
    {"Unit of Weight": "kgs", "Business Quantity": 25, "Unit Price (USD)": 40, "Width": 140, "GSM": 220, "Description": "Kilogram plural"},
    {"Unit of Weight": "K", "Business Quantity": 75, "Unit Price (USD)": 35, "Width": 160, "GSM": 190, "Description": "K abbreviation"},
    
    {"Unit of Weight": "g", "Business Quantity": 5000, "Unit Price (USD)": 25, "Width": 100, "GSM": 150, "Description": "Gram lowercase"},
    {"Unit of Weight": "GR", "Business Quantity": 3000, "Unit Price (USD)": 45, "Width": 130, "GSM": 170, "Description": "Gram GR format"},
    {"Unit of Weight": "gram", "Business Quantity": 2000, "Unit Price (USD)": 55, "Width": 110, "GSM": 160, "Description": "Gram full word"},
    {"Unit of Weight": "grams", "Business Quantity": 4000, "Unit Price (USD)": 60, "Width": 140, "GSM": 180, "Description": "Gram plural"},
    
    {"Unit of Weight": "lb", "Business Quantity": 220, "Unit Price (USD)": 70, "Width": 150, "GSM": 200, "Description": "Pound lb"},
    {"Unit of Weight": "LBS", "Business Quantity": 150, "Unit Price (USD)": 80, "Width": 120, "GSM": 210, "Description": "Pound LBS"},
    {"Unit of Weight": "pound", "Business Quantity": 100, "Unit Price (USD)": 90, "Width": 160, "GSM": 190, "Description": "Pound full word"},
    {"Unit of Weight": "pounds", "Business Quantity": 180, "Unit Price (USD)": 75, "Width": 140, "GSM": 170, "Description": "Pound plural"},
    
    {"Unit of Weight": "oz", "Business Quantity": 32, "Unit Price (USD)": 100, "Width": 150, "GSM": 200, "Description": "Ounce oz"},
    {"Unit of Weight": "OZ", "Business Quantity": 48, "Unit Price (USD)": 85, "Width": 130, "GSM": 180, "Description": "Ounce uppercase"},
    {"Unit of Weight": "ounce", "Business Quantity": 64, "Unit Price (USD)": 95, "Width": 140, "GSM": 190, "Description": "Ounce full word"},
    {"Unit of Weight": "ounces", "Business Quantity": 80, "Unit Price (USD)": 105, "Width": 160, "GSM": 210, "Description": "Ounce plural"},
    
    {"Unit of Weight": "ton", "Business Quantity": 2, "Unit Price (USD)": 500, "Width": 200, "GSM": 300, "Description": "Metric ton"},
    {"Unit of Weight": "MT", "Business Quantity": 1.5, "Unit Price (USD)": 600, "Width": 180, "GSM": 250, "Description": "Metric ton MT"},
    {"Unit of Weight": "tonne", "Business Quantity": 3, "Unit Price (USD)": 550, "Width": 220, "GSM": 280, "Description": "Tonne spelling"},
    
    # Complex conversions
    {"Unit of Weight": "mtr", "Business Quantity": 100, "Unit Price (USD)": 50, "Width": 150, "GSM": 200, "Description": "Meter mtr"},
    {"Unit of Weight": "MTR", "Business Quantity": 200, "Unit Price (USD)": 60, "Width": 120, "GSM": 180, "Description": "Meter uppercase"},
    {"Unit of Weight": "meter", "Business Quantity": 150, "Unit Price (USD)": 55, "Width": 140, "GSM": 190, "Description": "Meter full word"},
    {"Unit of Weight": "metres", "Business Quantity": 180, "Unit Price (USD)": 65, "Width": 160, "GSM": 210, "Description": "Meter plural UK"},
    
    {"Unit of Weight": "mtk", "Business Quantity": 50, "Unit Price (USD)": 80, "Width": 150, "GSM": 250, "Description": "Square meter mtk"},
    {"Unit of Weight": "MTK", "Business Quantity": 75, "Unit Price (USD)": 90, "Width": 140, "GSM": 220, "Description": "Square meter uppercase"},
    {"Unit of Weight": "m2", "Business Quantity": 60, "Unit Price (USD)": 85, "Width": 130, "GSM": 240, "Description": "Square meter m2"},
    {"Unit of Weight": "sqm", "Business Quantity": 40, "Unit Price (USD)": 95, "Width": 160, "GSM": 230, "Description": "Square meter sqm"},
    
    {"Unit of Weight": "yd", "Business Quantity": 100, "Unit Price (USD)": 45, "Width": 150, "GSM": 200, "Description": "Yard yd"},
    {"Unit of Weight": "YD", "Business Quantity": 120, "Unit Price (USD)": 50, "Width": 140, "GSM": 180, "Description": "Yard uppercase"},
    {"Unit of Weight": "yard", "Business Quantity": 110, "Unit Price (USD)": 48, "Width": 160, "GSM": 190, "Description": "Yard full word"},
    {"Unit of Weight": "yards", "Business Quantity": 130, "Unit Price (USD)": 52, "Width": 170, "GSM": 210, "Description": "Yard plural"},
    
    {"Unit of Weight": "roll", "Business Quantity": 10, "Unit Price (USD)": 200, "Width": 150, "GSM": 300, "Description": "Roll full word"},
    {"Unit of Weight": "ROLL", "Business Quantity": 15, "Unit Price (USD)": 250, "Width": 140, "GSM": 280, "Description": "Roll uppercase"},
    {"Unit of Weight": "rol", "Business Quantity": 12, "Unit Price (USD)": 220, "Width": 160, "GSM": 320, "Description": "Roll rol"},
    {"Unit of Weight": "rolls", "Business Quantity": 8, "Unit Price (USD)": 180, "Width": 130, "GSM": 270, "Description": "Roll plural"},
    
    # Precision units
    {"Unit of Weight": "mg", "Business Quantity": 500000, "Unit Price (USD)": 10, "Width": 100, "GSM": 150, "Description": "Milligram"},
    {"Unit of Weight": "MG", "Business Quantity": 750000, "Unit Price (USD)": 15, "Width": 120, "GSM": 160, "Description": "Milligram uppercase"},
    {"Unit of Weight": "carat", "Business Quantity": 5000, "Unit Price (USD)": 1000, "Width": 50, "GSM": 100, "Description": "Carat full word"},
    {"Unit of Weight": "CT", "Business Quantity": 2500, "Unit Price (USD)": 1200, "Width": 60, "GSM": 120, "Description": "Carat CT"},
    
    # Variations with spaces and punctuation
    {"Unit of Weight": "k g", "Business Quantity": 80, "Unit Price (USD)": 40, "Width": 150, "GSM": 200, "Description": "KG with space"},
    {"Unit of Weight": "k.g", "Business Quantity": 90, "Unit Price (USD)": 45, "Width": 140, "GSM": 180, "Description": "KG with dot"},
    {"Unit of Weight": "l.b.s", "Business Quantity": 200, "Unit Price (USD)": 65, "Width": 160, "GSM": 190, "Description": "LBS with dots"},
    {"Unit of Weight": " gr ", "Business Quantity": 6000, "Unit Price (USD)": 30, "Width": 130, "GSM": 170, "Description": "Gram with spaces"},
    
    # Imperial units
    {"Unit of Weight": "stone", "Business Quantity": 15, "Unit Price (USD)": 300, "Width": 180, "GSM": 250, "Description": "Stone UK"},
    {"Unit of Weight": "ST", "Business Quantity": 12, "Unit Price (USD)": 280, "Width": 170, "GSM": 240, "Description": "Stone ST"},
    {"Unit of Weight": "quintal", "Business Quantity": 5, "Unit Price (USD)": 800, "Width": 200, "GSM": 300, "Description": "Quintal"},
    {"Unit of Weight": "Q", "Business Quantity": 3, "Unit Price (USD)": 750, "Width": 190, "GSM": 280, "Description": "Quintal Q"},
]

# Create DataFrame
df = pd.DataFrame(sample_data)

# Create input directory if it doesn't exist
os.makedirs("input", exist_ok=True)

# Save to Excel file with multiple sheets
with pd.ExcelWriter("input/sample_units_test.xlsx", engine='openpyxl') as writer:
    # Main sheet with all data
    df.to_excel(writer, sheet_name="All_Units", index=False)
    
    # Separate sheets by category
    direct_units = df[df["Unit of Weight"].str.upper().isin(
        ["KG", "KGS", "K", "G", "GR", "GRAM", "GRAMS", "LB", "LBS", "POUND", "POUNDS", 
         "OZ", "OUNCE", "OUNCES", "TON", "MT", "TONNE", "MG", "CARAT", "CT", "STONE", "ST", "QUINTAL", "Q"]
    )]
    direct_units.to_excel(writer, sheet_name="Direct_Conversions", index=False)
    
    complex_units = df[df["Unit of Weight"].str.upper().isin(
        ["MTR", "METER", "METRES", "MTK", "M2", "SQM", "YD", "YARD", "YARDS", "ROLL", "ROL", "ROLLS"]
    )]
    complex_units.to_excel(writer, sheet_name="Complex_Conversions", index=False)
    
    variations = df[df["Unit of Weight"].isin(["k g", "k.g", "l.b.s", " gr "])]
    variations.to_excel(writer, sheet_name="Spacing_Variations", index=False)

print("✅ Sample Excel file created: input/sample_units_test.xlsx")
print(f"📊 Generated {len(df)} test records across 4 sheets:")
print("   - All_Units: Complete dataset")
print("   - Direct_Conversions: Units with direct conversion factors")
print("   - Complex_Conversions: Units requiring additional parameters")
print("   - Spacing_Variations: Units with unusual spacing/punctuation")
print("\n💡 You can now test the converter with this file!")
print("🚀 Run: python business_quantity_converter.py")
