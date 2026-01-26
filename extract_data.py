import csv
import pandas as pd
import os
import glob


def get_piezo_ports(csv_filename):
    """
    Determine piezo port columns based on drill name extracted from CSV filename.
    """
    
    with open(csv_filename, newline="", encoding="utf-8") as f :
        reader = csv.reader(f)
        header = next(reader)
    
    piezo_ports = []
    piezo_ports = [col.split('_')[0] for col in header if col.endswith('_p')]
    return piezo_ports

def add_datas_to_excel(excel_filename="", csv_filename="", drill_name="", port="", output_filename=""):
    
    # Read CSV data
    df_csv = pd.read_csv(csv_filename)
    
    # Read Excel sheet
    sheet_name = f"{drill_name} {port}"
    
    try:
        df_excel = pd.read_excel(excel_filename, sheet_name=sheet_name)
    except ValueError:
        print(f"✗ Sheet '{sheet_name}' not found, skipping...")
        return None, None
    
    # Get the last row values for copying (TC, Section, Drillhole, Piezometer)
    if len(df_excel) > 0:
        last_row = df_excel.iloc[-1]
        tc_value = last_row.get('TC', '')
        section_value = last_row.get('Section', '')
        drillhole_value = last_row.get('Drillhole', '')
        piezometer_value = last_row.get('Piezometer', '')
    else:
        # If sheet is empty, use default values or leave blank
        tc_value = ''
        section_value = ''
        drillhole_value = drill_name
        piezometer_value = port
    
    # Map CSV columns to Excel columns
    new_rows = []
    for _, row in df_csv.iterrows():
        # Get port values
        port_p = row[f'{port}_p']
        port_t = row[f'{port}_t']
        
        # Check if port is disconnected (-99 or -999999)
        if port_p in [-99, -999999] or port_t in [-99, -999999]:
            continue  # Skip this row if port is disconnected
        
        # Convert negative values to absolute (except disconnected values)
        if port_p < 0:
            port_p = abs(port_p)
        if port_t < 0:
            port_t = abs(port_t)
        
        new_row = {
            'Timestamp': row['Date'],  # Assuming 'Date' column exists in CSV
            'TC': tc_value,  # Copy from last row
            'Section': section_value,  # Copy from last row
            'Drillhole': drillhole_value,  # Copy from last row
            'Piezometer': piezometer_value,  # Copy from last row
            'B': port_p,  # Map port_p to B
            'T': port_t   # Map port_t to T
        }
        new_rows.append(new_row)
    
    # Create DataFrame from new rows
    df_new = pd.DataFrame(new_rows)
    
    # Append to existing Excel data
    df_combined = pd.concat([df_excel, df_new], ignore_index=True)
    
    print(f"✓ Added {len(new_rows)} rows to sheet '{sheet_name}'")
    
    return df_combined, sheet_name


# Main execution
if __name__ == "__main__":
    csv_file = "BH20-01_20260124_085152.csv"
    excel_file = "for tableau TC2 New VWP_24_Jan_2026.xlsx"
    
    # Generate output filename
    output_file = excel_file.replace('.xlsx', '_updated.xlsx')

    drill_name = csv_file.split('_')[0]

    piezo_ports = get_piezo_ports(csv_file)
    
    print(f"Drill name: {drill_name}")
    print(f"Piezo ports found: {piezo_ports}\n")

    # Store all updated sheets
    all_sheets = {}
    
    # Load all existing sheets from original file
    xls = pd.ExcelFile(excel_file)
    for sheet in xls.sheet_names:
        all_sheets[sheet] = pd.read_excel(excel_file, sheet_name=sheet)

    for port in piezo_ports:
        print(f"Processing port: {port}")
        result = add_datas_to_excel(excel_file, csv_file, drill_name, port, output_file)
        if result[0] is not None:
            df_combined, sheet_name = result
            all_sheets[sheet_name] = df_combined
        print()
    
    # Write all sheets to new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\n✓ Saved updated file as: {output_file}")
