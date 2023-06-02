import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

def load_raw():
    try:
        df = pd.read_excel("RAW.xlsm")
        return df
    except FileNotFoundError as e:
        print("Error: File not found:", e)
        return None

def main():
    df = load_raw()
    if df is not None:
        workbook = load_workbook("RAW.xlsm")
        writer = pd.ExcelWriter("new_file.xlsm", engine="openpyxl")  # Use openpyxl engine for writing

        # Write DataFrame to the sheet without the index column and header
        df.to_excel(writer, sheet_name="Instructions", index=False, header=False)

        # Access the workbook and sheet
        sheet = writer.book["Instructions"]

        # Insert three rows at the start of the sheet
        sheet.insert_rows(1, amount=4)

        # Apply formatting to the inserted rows and first 4 rows
        fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        font = Font(color="FFFFFF", size=14)
        for row in sheet.iter_rows(min_row=1, max_row=4, min_col=1, max_col=10):
            for cell in row:
                cell.fill = fill
                cell.font = font

        # Apply formatting to the range A11:J11
        for row in sheet.iter_rows(min_row=11, max_row=11, min_col=1, max_col=10):
            for cell in row:
                cell.fill = fill
                cell.font = font
                cell.column_letter = cell.column_letter.upper()
                sheet.column_dimensions[cell.column_letter].width = 15  # Adjust the width as desired


        # Insert text into cell A3 and increase font size
        sheet["A3"] = "PSG EMEA PRODUCTS Weight, Dimensions, Pallet data"
        sheet["A3"].font = Font(size=26)
        
        # Insert text into cell A6
        sheet["A6"] = "HP Confidential. For HP and Partner internal use only. The information contained herein is HP Confidential. Disclosure is governed by your HP Partner Agreement, HP Retail Partner Agreement, or another applicable contract (e.g., Confidential Disclosure Agreement)."
        
        # Delete content from cell A5
        sheet["A5"].value = None

        # Save the workbook
        writer.save()
        writer.close()

if __name__ == "__main__":
    main()
