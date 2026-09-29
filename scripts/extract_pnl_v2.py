import os
import re
import pandas as pd
from openpyxl import load_workbook

FOLDER_PATH = 'Data_Folder' 
OUTPUT_FILE = 'Master_PnL_Complete.xlsx'

CATEGORIES_IN_ORDER = [
    "Sales Revenue",
    "Other Revenue",
    "Sales Discounts",
    "Sales Returns",
    "Cost of raw materials",
    "Cost of parts used",
    "Direct labor costs",
    "Overhead costs",
    "Automobile",
    "Rented Equipment",
    "Insurance",
    "Job expenses",
    "Legal and Professional Fees",
    "Maintenance and Repair",
    "Meals",
    "Office Expenses",
    "Rent or Lease",
    "Utilities",
    "Vehicle Expenses",
    "Miscellaneous Expenses"
]

TOTALS_IN_ORDER = [
    "Total Revenue",
    "Total COGS",
    "Gross Profit",
    "Total Expenses",
    "Net Operating Profit",
    "Total Other Expenses",
    "Net Profit"
]

def parse_filename(filename):
    """
    Extracts Month, Year, and Project Name from inconsistent filename formats.
    Input examples: '1.5_1.16.2025 - Project Name.xlsx', '2.11-12.2025 - Event.xlsx'
    """
    name = filename.replace('.xlsx', '')
    
    # Extract Month (first number in filename)
    month_match = re.match(r'^(\d{1,2})', name)
    month = int(month_match.group(1)) if month_match else 0
    
    # Extract Year (4-digit number anywhere in filename)
    year_match = re.search(r'(\d{4})', name)
    year = int(year_match.group(1)) if year_match else 2025
    
    # Extract Project Name (remove date prefix with various separators)
    project_name = re.sub(r'^\d{1,2}[.\-_]\d{1,2}[.\-_]?\d{0,2}\s*[-]?\s*', '', name)
    
    return month, year, project_name

def clean_label(text):
    if not text: return ""
    return text.split('(')[0].strip()

def find_value_in_row(ws, row_num, label_text):
    """Find value at column  C, then find value at column E"""
    for col in ['C', 'D', 'E', 'F', 'G']:  # Coba beberapa kolom untuk label
        cell = ws[f'{col}{row_num}']
        if cell.value and label_text.upper() in str(cell.value).upper():
            # Ketemu labelnya, sekarang cari nilai di kolom E (atau kolom setelahnya)
            value_cell = ws[f'E{row_num}']
            return value_cell.value if value_cell.value else 0
    return 0

def extract_data_from_file(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    
    data = {key: 0 for key in TOTALS_IN_ORDER}
    for cat in CATEGORIES_IN_ORDER:
        data[cat] = 0

    # Scan rows 1-50 (covers all data sections + noise)
    for row_num in range(1, 51):
        cell_c = ws[f'C{row_num}']
        cell_e = ws[f'E{row_num}']
        
        if not cell_c.value:
            continue
            
        text_c = str(cell_c.value).upper()
        val_e = cell_e.value if cell_e.value is not None else 0
        
        # Match against TOTALS keywords
        for total_name in TOTALS_IN_ORDER:
            if total_name.upper() in text_c:
                data[total_name] = val_e
                break
        
        # Match against DETAIL CATEGORIES keywords
        for cat in CATEGORIES_IN_ORDER:
            if cat.upper() in text_c:
                data[cat] = val_e
                break

    wb.close()
    return data

def main():
    all_data = []
    
    # Define final column order (metadata + totals + details)
    final_columns = ['Filename', 'Project Name', 'Month', 'Year'] + TOTALS_IN_ORDER + CATEGORIES_IN_ORDER

    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith('.xlsx') and not filename.startswith('~$'):
            filepath = os.path.join(FOLDER_PATH, filename)
            
            month, year, project = parse_filename(filename)
            data = extract_data_from_file(filepath)
            
            row_data = {
                'Filename': filename,
                'Project Name': project,
                'Month': month,
                'Year': year
            }
            row_data.update(data)
            all_data.append(row_data)

    df = pd.DataFrame(all_data)
    
    # Sort chronologically
    df = df.sort_values(by=['Year', 'Month'])
    
    # Ensure consistent column order and fill missing values with 0
    df = df.reindex(columns=final_columns, fill_value=0)
    
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Selesai! Data tersimpan di {OUTPUT_FILE}")
    print(f"Total file diproses: {len(all_data)}")

if __name__ == '__main__':
    main()