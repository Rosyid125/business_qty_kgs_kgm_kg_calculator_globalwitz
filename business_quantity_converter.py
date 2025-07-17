import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import sys
from pathlib import Path
import threading
import openpyxl
from openpyxl import load_workbook

class BusinessQuantityConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Business Quantity to KG Converter")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Variables
        self.input_folder = "./input"
        self.output_folder = "./output"
        self.selected_file = None
        self.workbook = None
        self.valid_sheets = []
        self.selected_columns = {}
        
        # Create folders if they don't exist
        os.makedirs(self.input_folder, exist_ok=True)
        os.makedirs(self.output_folder, exist_ok=True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Business Quantity to KG Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="1. Select Excel File", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Excel Files:").grid(row=0, column=0, sticky=tk.W)
        
        self.file_listbox = tk.Listbox(file_frame, height=4)
        self.file_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        ttk.Button(file_frame, text="Refresh Files", 
                  command=self.refresh_files).grid(row=1, column=2, sticky=tk.W, padx=(10, 0))
        
        # Sheet selection section
        sheet_frame = ttk.LabelFrame(main_frame, text="2. Select Sheet to Process", padding="10")
        sheet_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        sheet_frame.columnconfigure(1, weight=1)
        
        # Single sheet dropdown selection
        ttk.Label(sheet_frame, text="Choose sheet to process:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, 
                                       state="readonly", width=50)
        self.sheet_combo.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_select)
        
        # Column mapping section
        column_frame = ttk.LabelFrame(main_frame, text="3. Map Columns", padding="10")
        column_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        column_frame.columnconfigure(1, weight=1)
        
        # Column mapping controls
        self.column_vars = {}
        column_labels = [
            ("Unit of Weight:", "unit_of_weight"),
            ("Business Quantity:", "business_quantity"),
            ("Unit Price (USD):", "unit_price"),
            ("Width:", "width"),
            ("GSM:", "gsm")
        ]
        
        for i, (label, key) in enumerate(column_labels):
            ttk.Label(column_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2)
            self.column_vars[key] = tk.StringVar()
            self.column_vars[key].trace('w', lambda *args: self.check_ready_to_process())
            combo = ttk.Combobox(column_frame, textvariable=self.column_vars[key], 
                               state="readonly", width=30)
            combo.grid(row=i, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
            setattr(self, f"{key}_combo", combo)
        
        # Process button
        process_frame = ttk.Frame(main_frame)
        process_frame.grid(row=4, column=0, columnspan=3, pady=(10, 0))
        
        self.process_btn = ttk.Button(process_frame, text="Start Conversion", 
                                     command=self.start_conversion, state="disabled")
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Help button for supported units
        help_btn = ttk.Button(process_frame, text="Show Supported Units", 
                             command=self.show_supported_units)
        help_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.progress = ttk.Progressbar(process_frame, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Log section
        log_frame = ttk.LabelFrame(main_frame, text="Process Log", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Initial setup log
        self.log("📋 Business Quantity to KG Converter v2.0 initialized")
        self.log("📁 Place your Excel files in the 'input' folder")
        self.log("🎯 NEW: Robust unit recognition with 50+ unit variants!")
        self.log("💡 Click 'Show Supported Units' to see all available units")
        self.refresh_files()
    
    def log(self, message):
        """Add message to log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def show_supported_units(self):
        """Show window with all supported units"""
        units_window = tk.Toplevel(self.root)
        units_window.title("Supported Units")
        units_window.geometry("800x600")
        units_window.resizable(True, True)
        
        # Create notebook for tabs
        notebook = ttk.Notebook(units_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Direct conversion tab
        direct_frame = ttk.Frame(notebook)
        notebook.add(direct_frame, text="Direct Conversion")
        
        direct_text = scrolledtext.ScrolledText(direct_frame, wrap=tk.WORD)
        direct_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        direct_units = """DIRECT CONVERSION UNITS (No additional parameters needed)

📏 KILOGRAM GROUP (Factor: 1.0):
KG, KGS, KGM, K, KILO, KILOS, KILOGRAM, KILOGRAMME

📏 GRAM GROUP (Factor: 0.001):
G, GR, GRM, GRAM, GRAMS, GRAMME, GM, GMS

📏 POUND GROUP (Factor: 0.453592):
LB, LBS, POUND, POUNDS, PND, PNDS, LBM

📏 OUNCE GROUP (Factor: 0.0283495):
OZ, OUNCE, OUNCES, ONZ

📏 TON GROUP:
• TON, TONS, TONNE, TONNES, T (Factor: 1000.0)
• MT, METRICTON, METRICTONS (Factor: 1000.0)
• SHORTTON (Factor: 907.185)
• LONGTON (Factor: 1016.05)

📏 IMPERIAL UNITS:
• STONE, STONES, ST (Factor: 6.35029)
• QUINTAL, QUINTALS, Q, QTL (Factor: 100.0)

📏 PRECISION UNITS:
• GRAIN, GRAINS, GRN (Factor: 0.00006479891)
• CARAT, CARATS, CT, CAR (Factor: 0.0002)
• MG, MILLIGRAM, MILLIGRAMS (Factor: 0.000001)
• UG, MCG, MICROGRAM, MICROGRAMS (Factor: 0.000000001)

📏 ADDITIONAL IMPERIAL:
• DRAM (Factor: 0.0017718)
• SCRUPLE (Factor: 0.001296)
• PENNYWEIGHT (Factor: 0.001555)
• SLUG (Factor: 14.5939)
• HUNDREDWEIGHT (Factor: 50.8023)
• USHUNDREDWEIGHT (Factor: 45.3592)

Note: All units are case-insensitive and ignore spaces/punctuation.
Example: "kg" = "KG" = "k g" = "k.g" = "kilo"
"""
        direct_text.insert(tk.END, direct_units)
        direct_text.config(state=tk.DISABLED)
        
        # Complex conversion tab
        complex_frame = ttk.Frame(notebook)
        notebook.add(complex_frame, text="Complex Conversion")
        
        complex_text = scrolledtext.ScrolledText(complex_frame, wrap=tk.WORD)
        complex_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        complex_units = """COMPLEX CONVERSION UNITS (Require additional parameters)

📐 LINEAR UNITS (Need: Unit Price, Width, GSM):
MTR, METER, METRE, M, MTS
Formula: (Unit Price × 1000) ÷ (Width × GSM)

📐 YARD UNITS (Need: Unit Price, Width, GSM):
YD, YARD, YARDS, YDS
Formula: ((Unit Price ÷ 0.9144) × 1000) ÷ (Width × GSM)

📐 AREA UNITS (Need: Unit Price, GSM):
MTK, MTR2, M2, SQM, SQMETER, SQUAREMETER
Formula: (Unit Price × 1000) ÷ GSM

📐 SQUARE FEET UNITS (Need: Business Quantity, GSM):
SQF, SQFT, SQUAREFEET, SQUAREFOOT
Formula: Business Quantity × 0.092903 × GSM ÷ 1000

📐 ROLL UNITS (Need: Business Quantity, GSM):
ROL, ROLL, ROLLS
Formula: Business Quantity ÷ GSM

REQUIRED COLUMNS FOR COMPLEX CONVERSION:
• Unit Price (USD): Required for MTR, YD, MTK calculations
• Width: Required for MTR and YD calculations  
• GSM: Required for all complex conversions (MTR, YD, MTK, SQF, ROLL)
• Business Quantity: Always required

CONVERSION PRIORITY:
1. Try direct conversion first (if unit is in direct list)
2. If not direct, try complex conversion (if required params available)
3. If neither works, mark as unconvertible

TIPS:
• Make sure all required columns are mapped correctly
• Check that numeric values are valid (not text or empty)
• Review conversion statistics in the log for success rates
"""
        complex_text.insert(tk.END, complex_units)
        complex_text.config(state=tk.DISABLED)
        
        # Examples tab
        examples_frame = ttk.Frame(notebook)
        notebook.add(examples_frame, text="Examples")
        
        examples_text = scrolledtext.ScrolledText(examples_frame, wrap=tk.WORD)
        examples_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        examples_content = """CONVERSION EXAMPLES

✅ DIRECT CONVERSIONS:
• 100 KG → 100 KG (no change)
• 1000 G → 1 KG (1000 × 0.001)
• 10 LBS → 4.53592 KG (10 × 0.453592)
• 16 OZ → 0.453592 KG (16 × 0.0283495)
• 1 TON → 1000 KG (1 × 1000)

✅ COMPLEX CONVERSIONS:
Example data for MTR conversion:
• Business Quantity: 100 MTR
• Unit Price: $50 USD
• Width: 150 cm
• GSM: 200
• Result: (50 × 1000) ÷ (150 × 200) = 1.67 KG

Example data for MTK conversion:
• Business Quantity: 10 MTK
• Unit Price: $30 USD
• GSM: 300
• Result: (30 × 1000) ÷ 300 = 100 KG

Example data for ROLL conversion:
• Business Quantity: 5 ROLL
• GSM: 250
• Result: 5 ÷ 250 = 0.02 KG

Example data for SQF conversion:
• Business Quantity: 1000 SQF
• GSM: 200
• Result: 1000 × 0.092903 × 200 ÷ 1000 = 18.5806 KG

🔧 UNIT RECOGNITION EXAMPLES:
Input → Recognized As:
• "kg" → KG
• "kilograms" → KG
• "lb" → LBS
• "pounds" → LBS
• "g" → G
• "grams" → G
• "oz" → OZ
• "ounces" → OZ
• "mtr" → MTR
• "meters" → MTR
• "sqf" → SQF
• "square feet" → SQF

❌ COMMON ISSUES:
• Empty business quantity → Result: "-"
• Invalid unit → Result: "-"
• Missing required parameters for complex units → Result: "-"
• Non-numeric values → Result: "-"
"""
        examples_text.insert(tk.END, examples_content)
        examples_text.config(state=tk.DISABLED)
        
        # Close button
        close_btn = ttk.Button(units_window, text="Close", command=units_window.destroy)
        close_btn.pack(side=tk.BOTTOM, pady=10)
    
    def refresh_files(self):
        """Refresh the list of Excel files"""
        self.file_listbox.delete(0, tk.END)
        try:
            files = [f for f in os.listdir(self.input_folder) 
                    if f.lower().endswith(('.xlsx', '.xls'))]
            for file in files:
                self.file_listbox.insert(tk.END, file)
            
            if not files:
                self.log("⚠️ No Excel files found in input folder")
            else:
                self.log(f"✅ Found {len(files)} Excel file(s)")
        except Exception as e:
            self.log(f"❌ Error reading input folder: {str(e)}")
    
    def on_file_select(self, event):
        """Handle file selection"""
        selection = self.file_listbox.curselection()
        if selection:
            self.selected_file = self.file_listbox.get(selection[0])
            self.log(f"📁 Selected file: {self.selected_file}")
            self.load_sheets()
    
    def load_sheets(self):
        """Load and validate sheets from selected file"""
        if not self.selected_file:
            return
        
        try:
            file_path = os.path.join(self.input_folder, self.selected_file)
            self.log(f"📖 Loading sheets from {self.selected_file}...")
            
            # Load workbook
            self.workbook = pd.ExcelFile(file_path)
            all_sheets = self.workbook.sheet_names
            
            self.log(f"🔍 Found {len(all_sheets)} sheet(s) in file")
            
            # Validate sheets
            self.valid_sheets = []
            
            for sheet_name in all_sheets:
                try:
                    # Try to read a small sample to check if sheet has data
                    df_sample = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
                    
                    if not df_sample.empty and len(df_sample.columns) > 0:
                        # Check if there's actual data (not all NaN)
                        has_data = not df_sample.dropna(how='all').empty
                        
                        if has_data:
                            self.valid_sheets.append(sheet_name)
                            self.log(f"✅ Sheet '{sheet_name}': Valid data found")
                        else:
                            self.log(f"⚠️ Sheet '{sheet_name}': Empty, skipped")
                    else:
                        self.log(f"⚠️ Sheet '{sheet_name}': No data, skipped")
                
                except Exception as e:
                    self.log(f"❌ Sheet '{sheet_name}': Error - {str(e)}")
            
            if self.valid_sheets:
                self.log(f"📊 {len(self.valid_sheets)} valid sheet(s) available")
                
                # Populate sheet dropdown
                self.sheet_combo['values'] = self.valid_sheets
                
                # Auto-select first sheet
                if self.valid_sheets:
                    self.sheet_combo.set(self.valid_sheets[0])
                    self.load_columns()
            else:
                self.log("❌ No valid sheets found")
                
        except Exception as e:
            self.log(f"❌ Error loading file: {str(e)}")
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
    
    def on_sheet_select(self, event=None):
        """Handle sheet dropdown selection"""
        if self.sheet_var.get():
            self.log(f"📋 Selected sheet: '{self.sheet_var.get()}'")
            self.load_columns()
    
    def load_columns(self):
        """Load columns from selected sheet in dropdown"""
        selected_sheet = self.sheet_var.get()
        if not selected_sheet or not self.valid_sheets or selected_sheet not in self.valid_sheets:
            return
        
        try:
            file_path = os.path.join(self.input_folder, self.selected_file)
            
            # Read header row from the selected sheet
            df_header = pd.read_excel(file_path, sheet_name=selected_sheet, nrows=0)
            columns = list(df_header.columns)
            
            # Update all comboboxes
            for combo_name in ['unit_of_weight_combo', 'business_quantity_combo', 
                             'unit_price_combo', 'width_combo', 'gsm_combo']:
                combo = getattr(self, combo_name)
                combo['values'] = columns
                combo.set('')  # Clear selection
            
            self.log(f"📋 Loaded {len(columns)} columns from sheet '{selected_sheet}'")
            self.check_ready_to_process()
            
        except Exception as e:
            self.log(f"❌ Error loading columns: {str(e)}")
    
    def check_ready_to_process(self):
        """Check if ready to process and enable/disable button"""
        ready = (self.selected_file and 
                self.valid_sheets and 
                self.column_vars['unit_of_weight'].get() and 
                self.column_vars['business_quantity'].get())
        
        self.process_btn.config(state="normal" if ready else "disabled")
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        self.process_btn.config(state="disabled")
        self.progress.start()
        
        # Run conversion in separate thread to prevent UI freezing
        thread = threading.Thread(target=self.convert_data)
        thread.daemon = True
        thread.start()
    
    def convert_data(self):
        """Main conversion logic"""
        try:
            self.log("\n🚀 Starting conversion process...")
            
            # Get selected sheet from dropdown
            selected_sheet = self.sheet_var.get()
            if not selected_sheet:
                self.log("❌ No sheet selected")
                return
            
            # Get column mappings
            columns = {
                'unit_of_weight': self.column_vars['unit_of_weight'].get(),
                'business_quantity': self.column_vars['business_quantity'].get(),
                'unit_price': self.column_vars['unit_price'].get() or None,
                'width': self.column_vars['width'].get() or None,
                'gsm': self.column_vars['gsm'].get() or None
            }
            
            file_path = os.path.join(self.input_folder, self.selected_file)
            
            # Create output workbook with single sheet
            with pd.ExcelWriter(os.path.join(self.output_folder, f"converted_{self.selected_file}"), 
                               engine='openpyxl') as writer:
                
                self.log(f"\n📊 Processing sheet: '{selected_sheet}'")
                
                # Read data
                df = pd.read_excel(file_path, sheet_name=selected_sheet)
                self.log(f"📖 Read {len(df)} rows from sheet")
                
                # Convert data
                df_converted = self.convert_business_quantity_to_kg(df, columns)
                
                # Write to output
                df_converted.to_excel(writer, sheet_name=selected_sheet, index=False)
                self.log(f"✅ Sheet '{selected_sheet}' processed successfully")
            
            output_path = os.path.join(self.output_folder, f"converted_{self.selected_file}")
            file_size = os.path.getsize(output_path) / (1024 * 1024)  # MB
            
            self.log(f"\n🎉 Conversion completed successfully!")
            self.log(f"📁 Output file: {output_path}")
            self.log(f"📦 File size: {file_size:.2f} MB")
            
            messagebox.showinfo("Success", f"Conversion completed!\nOutput saved to: converted_{self.selected_file}")
            
        except Exception as e:
            error_msg = f"❌ Conversion failed: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Error", f"Conversion failed:\n{str(e)}")
        
        finally:
            self.root.after(0, self._conversion_finished)
    
    def _conversion_finished(self):
        """Called when conversion is finished (runs in main thread)"""
        self.progress.stop()
        self.process_btn.config(state="normal")
    
    def extract_first_value(self, value_string):
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
    
    def clean_gsm_operators(self, gsm_value):
        """Remove operators (>, <, >=, <=) from GSM value and return the numeric value"""
        if not gsm_value or pd.isna(gsm_value):
            return None
        
        # Convert to string
        str_value = str(gsm_value).strip()
        
        # Remove operators: >=, <=, >, <
        import re
        # Pattern to match operators at the beginning of the string
        pattern = r'^(>=|<=|>|<)\s*'
        cleaned_value = re.sub(pattern, '', str_value)
        
        return cleaned_value if cleaned_value else None

    def normalize_unit(self, unit_string):
        """Normalize unit string to standard format"""
        if not unit_string or pd.isna(unit_string):
            return ''
        
        # Convert to uppercase and remove spaces/punctuation
        unit = str(unit_string).upper().strip().replace(' ', '').replace('.', '').replace('-', '')
        
        # Unit mapping dictionary - maps various spellings to standard units
        unit_mappings = {
            # Kilogram variants
            'KG': 'KG', 'KGS': 'KG', 'KGM': 'KG', 'K': 'KG', 'KILO': 'KG', 'KILOS': 'KG',
            'KILOGRAM': 'KG', 'KILOGRAMS': 'KG', 'KILOGRAMME': 'KG', 'KILOGRAMMES': 'KG',
            
            # Gram variants
            'G': 'G', 'GR': 'G', 'GRM': 'G', 'GRAM': 'G', 'GRAMS': 'G', 'GRAMME': 'G', 'GRAMMES': 'G',
            'GMS': 'G', 'GM': 'G',
            
            # Pound variants
            'LB': 'LBS', 'LBS': 'LBS', 'POUND': 'LBS', 'POUNDS': 'LBS', 'PND': 'LBS', 'PNDS': 'LBS',
            'LBM': 'LBS', 'LBMASS': 'LBS',
            
            # Ounce variants
            'OZ': 'OZ', 'OUNCE': 'OZ', 'OUNCES': 'OZ', 'ONZ': 'OZ',
            
            # Ton variants
            'TON': 'TON', 'TONS': 'TON', 'TONNE': 'TON', 'TONNES': 'TON', 'T': 'TON',
            'MT': 'MT', 'METRICTON': 'MT', 'METRICTONS': 'MT',
            
            # Stone variants
            'ST': 'STONE', 'STONE': 'STONE', 'STONES': 'STONE',
            
            # Quintal variants
            'Q': 'QUINTAL', 'QTL': 'QUINTAL', 'QUINTAL': 'QUINTAL', 'QUINTALS': 'QUINTAL',
            
            # Grain variants
            'GRN': 'GRAIN', 'GRAIN': 'GRAIN', 'GRAINS': 'GRAIN',
            
            # Carat variants
            'CT': 'CARAT', 'CARAT': 'CARAT', 'CARATS': 'CARAT', 'CAR': 'CARAT',
            
            # Milligram variants
            'MG': 'MG', 'MILLIGRAM': 'MG', 'MILLIGRAMS': 'MG', 'MILLIGRAMME': 'MG',
            
            # Microgram variants
            'UG': 'UG', 'MCG': 'UG', 'MICROGRAM': 'UG', 'MICROGRAMS': 'UG',
            
            # Complex units (require additional parameters)
            'MTR': 'MTR', 'METER': 'MTR', 'METRE': 'MTR', 'METERS': 'MTR', 'METRES': 'MTR',
            'M': 'MTR', 'MTS': 'MTR',
            
            'MTK': 'MTK', 'MTR2': 'MTK', 'M2': 'MTK', 'SQM': 'MTK', 'SQMETER': 'MTK',
            'SQUAREMETER': 'MTK', 'SQUAREMETERS': 'MTK',
            
            'SQF': 'SQF', 'SQFT': 'SQF', 'SQUAREFEET': 'SQF', 'SQUAREFOOT': 'SQF',
            'SQFEET': 'SQF', 'SF': 'SQF',
            
            'YD': 'YD', 'YARD': 'YD', 'YARDS': 'YD', 'YDS': 'YD',
            
            'ROL': 'ROLL', 'ROLL': 'ROLL', 'ROLLS': 'ROLL',
            
            # Feet variants
            'FT': 'FT', 'FOOT': 'FT', 'FEET': 'FT',
            
            # Inch variants
            'IN': 'IN', 'INCH': 'IN', 'INCHES': 'IN'
        }
        
        return unit_mappings.get(unit, unit)
    
    def get_conversion_factor(self, unit):
        """Get conversion factor to convert unit to KG"""
        # Conversion factors to KG
        conversion_factors = {
            # Direct conversions (no additional parameters needed)
            'KG': 1.0,                    # Kilogram (base unit)
            'G': 0.001,                   # Gram
            'LBS': 0.453592,              # Pound
            'OZ': 0.0283495,              # Ounce
            'TON': 1000.0,                # Metric Ton
            'MT': 1000.0,                 # Metric Ton
            'STONE': 6.35029,             # Stone (UK)
            'QUINTAL': 100.0,             # Quintal
            'GRAIN': 0.00006479891,       # Grain
            'CARAT': 0.0002,              # Carat
            'MG': 0.000001,               # Milligram
            'UG': 0.000000001,            # Microgram
            
            # Imperial tons
            'SHORTTON': 907.185,          # US short ton
            'LONGTON': 1016.05,           # UK long ton
            
            # Additional weight units
            'DRAM': 0.0017718,            # Dram
            'SCRUPLE': 0.001296,          # Scruple
            'PENNYWEIGHT': 0.001555,      # Pennyweight
            'SLUG': 14.5939,              # Slug
            'HUNDREDWEIGHT': 50.8023,     # Hundredweight (UK)
            'USHUNDREDWEIGHT': 45.3592,   # Hundredweight (US)
        }
        
        return conversion_factors.get(unit, None)
    
    def convert_business_quantity_to_kg(self, df, columns):
        """Convert business quantity to KG with robust unit recognition"""
        self.log("🔄 Starting unit conversion...")
        
        # Create a copy to avoid modifying original
        df_result = df.copy()
        
        # Initialize result column
        df_result['BUSINESS QUANTITY (KG)'] = '-'
        
        converted_count = 0
        total_rows = len(df_result)
        unit_stats = {}  # Track conversion statistics
        
        for index, row in df_result.iterrows():
            # Get values
            raw_unit = row.get(columns['unit_of_weight'], '-')
            normalized_unit = self.normalize_unit(raw_unit)
            business_quantity = pd.to_numeric(row.get(columns['business_quantity'], 0), errors='coerce') or 0
            unit_price = pd.to_numeric(row.get(columns['unit_price'], 0), errors='coerce') or 0 if columns['unit_price'] else 0
            
            # Process Width: extract first value from comma-separated string
            width_raw = row.get(columns['width'], 0) if columns['width'] else 0
            width_first = self.extract_first_value(width_raw) if width_raw else None
            width = pd.to_numeric(width_first, errors='coerce') or 0 if width_first else 0
            
            # Process GSM: extract first value and remove operators
            gsm_raw = row.get(columns['gsm'], 0) if columns['gsm'] else 0
            gsm_first = self.extract_first_value(gsm_raw) if gsm_raw else None
            gsm_cleaned = self.clean_gsm_operators(gsm_first) if gsm_first else None
            gsm = pd.to_numeric(gsm_cleaned, errors='coerce') or 0 if gsm_cleaned else 0
            
            result = '-'
            conversion_method = 'none'
            
            # Track unit usage
            if normalized_unit not in unit_stats:
                unit_stats[normalized_unit] = {'count': 0, 'converted': 0}
            unit_stats[normalized_unit]['count'] += 1
            
            if business_quantity <= 0:
                result = '-'
                conversion_method = 'invalid_quantity'
            else:
                # Try direct conversion first
                conversion_factor = self.get_conversion_factor(normalized_unit)
                
                if conversion_factor is not None:
                    # Direct conversion
                    result = business_quantity * conversion_factor
                    conversion_method = 'direct'
                    unit_stats[normalized_unit]['converted'] += 1
                
                elif normalized_unit in ['MTR', 'MTK', 'YD', 'SQF', 'ROLL'] and gsm > 0:
                    # Complex conversions requiring additional parameters
                    if normalized_unit == 'MTR' and width > 0 and unit_price > 0:
                        result = (unit_price * 1000) / (width * gsm)
                        conversion_method = 'mtr_complex'
                        unit_stats[normalized_unit]['converted'] += 1
                    elif normalized_unit == 'MTK' and unit_price > 0:
                        result = (unit_price * 1000) / gsm
                        conversion_method = 'mtk_complex'
                        unit_stats[normalized_unit]['converted'] += 1
                    elif normalized_unit == 'YD' and width > 0 and unit_price > 0:
                        result = ((unit_price / 0.9144) * 1000) / (width * gsm)
                        conversion_method = 'yd_complex'
                        unit_stats[normalized_unit]['converted'] += 1
                    elif normalized_unit == 'SQF':
                        # SQF conversion: SQF * 0.092903 * GSM / 1000
                        result = business_quantity * 0.092903 * gsm / 1000
                        conversion_method = 'sqf_complex'
                        unit_stats[normalized_unit]['converted'] += 1
                    elif normalized_unit == 'ROLL':
                        result = business_quantity / gsm
                        conversion_method = 'roll_complex'
                        unit_stats[normalized_unit]['converted'] += 1
                    else:
                        conversion_method = 'missing_parameters'
                else:
                    conversion_method = 'unsupported_unit'
            
            # Store result
            if isinstance(result, (int, float)) and result != '-':
                df_result.at[index, 'BUSINESS QUANTITY (KG)'] = round(result, 6)
                converted_count += 1
            else:
                df_result.at[index, 'BUSINESS QUANTITY (KG)'] = '-'
            
            # Progress update
            if index % 50 == 0 and index > 0:
                self.log(f"  Progress: {index}/{total_rows} rows processed...")
        
        # Log conversion statistics
        self.log(f"\n📊 Conversion Statistics:")
        for unit, stats in unit_stats.items():
            if stats['count'] > 0:
                success_rate = (stats['converted'] / stats['count']) * 100
                self.log(f"  {unit}: {stats['converted']}/{stats['count']} converted ({success_rate:.1f}%)")
        
        self.log(f"\n✅ Conversion completed: {converted_count}/{total_rows} rows converted")
        return df_result

def main():
    root = tk.Tk()
    app = BusinessQuantityConverter(root)
    
    # Center window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
