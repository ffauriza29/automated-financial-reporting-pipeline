# 📊 Automated Financial Reporting Pipeline

### Transforming 107+ Unstructured Excel Files into an Interactive Financial Dashboard

![Annual Dashboard](screenshots/annual_dashboard.png)

---

## 🎯 Project Overview

Built an end-to-end automated data pipeline using Python to extract, clean, and transform messy financial data from 107 individual Excel files into professional Annual and Monthly Profit & Loss (P&L) reports. The project automated a process that would have taken 20+ hours of manual work into minutes, delivering accurate results within a tight 36-hour deadline.

---

## 💼 The Business Problem

A US-based event company needed to consolidate financial reports from 107 individual projects. Key challenges included:

- 📂 **Fragmented Data:** 107 separate Excel files with inconsistent filename formats (e.g., `1.5_1.16.2025`, `2.11-12.2025`)
- 🔄 **Shifting Row Positions:** Critical data positions (Gross Profit, Total Expenses) shifted between files due to variable "noise" rows (payment check numbers)
- 🗑️ **Noisy Data:** Advertisements and add-on instructions at the bottom of files needed to be ignored
- ⏰ **Tight Deadline:** Reports needed to be delivered within 36 hours

---

## 🛠️ Tech Stack

| Category | Tools |
|----------|-------|
| **Programming** | Python 3 |
| **Data Processing** | Pandas, Openpyxl |
| **Text Parsing** | Regular Expressions (Regex) |
| **Reporting** | Google Sheets (Advanced Formulas, Pivot Tables, Conditional Formatting) |
| **Visualization** | Google Sheets Charts, Sparklines |

---

## ⚙️ Technical Deep Dive: The Data Pipeline

### Phase 1: Data Ingestion & Filename Parsing

The first challenge was extracting metadata (Month, Year, Project Name) from highly inconsistent filenames. I used **Regular Expressions (Regex)** to parse varying date patterns.

```python
def parse_filename(filename):
    name = filename.replace('.xlsx', '')
    month_match = re.match(r'^(\d{1,2})', name)
    month = int(month_match.group(1)) if month_match else 0
    year_match = re.search(r'(\d{4})', name)
    year = int(year_match.group(1)) if year_match else 2025
    project_name = re.sub(r'^\d{1,2}[.\-_]\d{1,2}[.\-_]?\d{0,2}\s*[-]?\s*', '', name)
    return month, year, project_name
```


### Phase 2: Data Cleaning & Noise Reduction

Since row numbers shifted between files, a position-based approach (e.g., "get data from row 20") would fail. I designed a Keyword Search & State Machine algorithm to scan label columns.
Key Solutions:
✅ **Handling Shifting Rows**: Script scans Column C for specific keywords (TOTAL REVENUE, GROSS PROFIT) and retrieves corresponding values, ignoring row shifts
✅ **Removing Noise**: Ignores rows containing "Check#" or "Coefficient" (advertisement text)
✅ **Text Normalization**: Cleans category labels with dynamic descriptions in parentheses for proper grouping

```python
def clean_label(text):
    if not text: return ""
    return text.split('(')[0].strip()

def find_value_in_row(ws, row_num, label_text):
    """Find value at column  C, then find value at column E"""
    for col in ['C', 'D', 'E', 'F', 'G']:  
        cell = ws[f'{col}{row_num}']
        if cell.value and label_text.upper() in str(cell.value).upper():
            value_cell = ws[f'E{row_num}']
            return value_cell.value if value_cell.value else 0
    return 0

def extract_data_from_file(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    
    data = {key: 0 for key in TOTALS_IN_ORDER}
    for cat in CATEGORIES_IN_ORDER:
        data[cat] = 0

    # Scan all rows (1-50)
    for row_num in range(1, 51):
        # Ambil teks dari kolom C
        cell_c = ws[f'C{row_num}']
        cell_e = ws[f'E{row_num}']
        
        if not cell_c.value:
            continue
            
        text_c = str(cell_c.value).upper()
        val_e = cell_e.value if cell_e.value is not None else 0
        
        # Find TOTALS
        for total_name in TOTALS_IN_ORDER:
            if total_name.upper() in text_c:
                data[total_name] = val_e
                break
        
        # Find CATEGORIES (detail)
        for cat in CATEGORIES_IN_ORDER:
            if cat.upper() in text_c:
                data[cat] = val_e
                break

    wb.close()
    return data
```


### Phase 3: Data Transformation & Aggregation

After extracting raw data into a Master DataFrame, I used Pandas for reporting-ready transformations:
Column Standardization: Forced consistent column ordering matching accounting standards
Monthly & Yearly Aggregation: Used groupby and sum functions for Monthly and Annual report datasets
Missing Value Handling: Filled NaN values with 0 for expense categories not used in specific projects

```python
def main():
    print("Start data extraction (simplified version)...")
    all_data = []
    
    final_columns = ['Filename', 'Project Name', 'Month', 'Year'] + TOTALS_IN_ORDER + CATEGORIES_IN_ORDER

    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith('.xlsx') and not filename.startswith('~$'):
            filepath = os.path.join(FOLDER_PATH, filename)
            print(f"Processing: {filename}")
            
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
    df = df.sort_values(by=['Year', 'Month'])
    df = df.reindex(columns=final_columns, fill_value=0)
    
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Finished! Data stored at {OUTPUT_FILE}")
    print(f"Total processed file: {len(all_data)}")
```
---

## 📈 The Output: Financial Dashboard
The pipeline produces an integrated, clean, and decision-ready Google Sheets file.
1. **Annual P&L with KPI Scorecard**
![Annual Dashboard](screenshots/annual_dashboard.png)
Provides an executive summary at the top with key metrics (Total Revenue, Total Expenses, Net Profit, Profit Margin) dynamically pulled from the database. Expense details are organized into clear categories as requested by the client.
2. **Monthly P&L Comparative Report**
![Annual Dashboard](screenshots/monthly_dashboard.png)
Breaks down 12-month data into a horizontal format, enabling clients to compare spending month-to-month on an apple-to-apples basis. Includes Conditional Formatting (Green for profit, Red for loss) for easy visual reading.

---

## 💡 Key Takeaways
1. **Never Trust Raw Data**: Real-world data is almost always messy. Building robust logic (like keyword search instead of hard-coded row numbers) is crucial for scalability.
2. **Automation over Manual Work**: Automating the extraction of 107 files saved an estimated 20+ hours of manual work and eliminated human error risks (copy-paste mistakes).
3. **Business Context Matters**: Great Python code is meaningless if the final output isn't readable by business users. Report formatting and structure are as important as data cleanliness.


---

## How to Run This Project
1. Clone this repository
2. Install dependencies: pip install pandas openpyxl
3. Place your Excel files in the sample_data/ folder
4. Update FOLDER_PATH in scripts/extract_pnl.py
5. Run: python scripts/extract_pnl.py
6. Check the output/ folder for the consolidated data

--- 

## 📬 Contact
LinkedIn: [https://www.linkedin.com/in/mohfaurizarw]
Email: [fauriza29@gmail.com]

---
