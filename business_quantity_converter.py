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
        self.refresh_files()
    
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
    
    def log(self, message):
        """Add message to log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
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
    
    def convert_business_quantity_to_kg(self, df, columns):
        """Convert business quantity to KG"""
        self.log("🔄 Starting unit conversion...")
        
        # Create a copy to avoid modifying original
        df_result = df.copy()
        
        # Initialize result column
        df_result['BUSINESS QUANTITY (KG)'] = '-'
        
        converted_count = 0
        total_rows = len(df_result)
        
        for index, row in df_result.iterrows():
            # Get values
            unit_of_weight = str(row.get(columns['unit_of_weight'], '-')).upper().strip()
            business_quantity = pd.to_numeric(row.get(columns['business_quantity'], 0), errors='coerce') or 0
            unit_price = pd.to_numeric(row.get(columns['unit_price'], 0), errors='coerce') or 0 if columns['unit_price'] else 0
            width = pd.to_numeric(row.get(columns['width'], 0), errors='coerce') or 0 if columns['width'] else 0
            gsm = pd.to_numeric(row.get(columns['gsm'], 0), errors='coerce') or 0 if columns['gsm'] else 0
            
            result = '-'
            
            # Conversion logic
            if unit_of_weight in ['GRM', 'GR'] and business_quantity > 0:
                result = business_quantity * 1000  # GR to KG
            elif unit_of_weight in ['KG', 'KGM', 'KGS', 'K'] and business_quantity > 0:
                result = business_quantity  # Already in KG
            elif unit_of_weight == 'LBS' and business_quantity > 0:
                result = business_quantity * 0.453592  # LBS to KG
            elif business_quantity > 0 and unit_price > 0 and width > 0 and gsm > 0:
                if unit_of_weight == 'MTR':
                    result = (unit_price * 1000) / (width * gsm)
                elif unit_of_weight in ['MTK', 'MTR2']:
                    result = (unit_price * 1000) / gsm
                elif unit_of_weight == 'YD':
                    result = ((unit_price / 0.9144) * 1000) / (width * gsm)
                elif unit_of_weight in ['ROL', 'ROLL']:
                    result = business_quantity / gsm
            
            df_result.at[index, 'BUSINESS QUANTITY (KG)'] = result
            
            if result != '-':
                converted_count += 1
            
            # Progress update
            if index % 50 == 0 and index > 0:
                self.log(f"  Progress: {index}/{total_rows} rows processed...")
        
        self.log(f"✅ Conversion completed: {converted_count}/{total_rows} rows converted")
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
