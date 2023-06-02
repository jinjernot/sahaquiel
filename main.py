import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side

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
        writer = pd.ExcelWriter("new_file.xlsm", engine="openpyxl")

        # Write DataFrame to the sheet without the index column and header
        df.to_excel(writer, sheet_name="Instructions", index=False, header=False)

        # Access the workbook and sheet
        sheet = writer.book["Instructions"]

        # Insert three rows at the start of the sheet
        sheet.insert_rows(1, amount=4)

        # Apply formatting to the inserted rows and first 4 rows
        fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        font = Font(color="FFFFFF", size=26)
        smallfont = Font(color="FFFFFF", size=10)
        for row in sheet.iter_rows(min_row=1, max_row=4, min_col=1, max_col=10):
            for cell in row:
                cell.fill = fill
                cell.font = font
                sheet.row_dimensions[cell.row].height = 30

        # Apply formatting to the range A11:J11 and adjust column widths
        for row in sheet.iter_rows(min_row=11, max_row=11, min_col=1, max_col=10):
            for cell in row:
                cell.fill = fill
                cell.font = smallfont
                cell.alignment = cell.alignment.copy(wrap_text=True)  #
                cell.column_letter
                sheet.column_dimensions[cell.column_letter].width = 30
                cell.row
                sheet.row_dimensions[cell.row].height = 60
                
        # Select the ranges
        Box_Info = sheet["K8:O11"]
        Standard_AIR_IP = sheet["P5:W11"]
        Standard_TRUCK_IP = sheet["X5:AE11"]
        Standard_SEA_IP = sheet["AF5:AM11"]
        Customer_Specific_EP = sheet["AN5:AV11"]

        # Define the background colors for each range
        colors = {
            Box_Info: "e6e1eb",
            Standard_AIR_IP: "4e9fd5",
            Standard_TRUCK_IP: "a9cdf2",
            Standard_SEA_IP: "dbe6f6",
            Customer_Specific_EP: "34d1da",
        }

        # Apply background colors to each range
        for range_, color in colors.items():
            for row in range_:
                for cell in row:
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

        # Define border style
        border_style = Side(border_style="thin", color="000000")

        # Apply border to rows 9 and 10
        for row_num in range(9, 11):
            for col_num in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row_num, column=col_num)
                cell.border = Border(top=border_style, bottom=border_style, left=border_style, right=border_style)



        # Insert text into cell A3 and increase font size
        sheet["A3"] = "PSG EMEA PRODUCTS Weight, Dimensions, Pallet data"
        #sheet["A3"].font = Font(size=16)
        
        # Insert text into cell A6
        sheet["A6"] = "HP Confidential. For HP and Partner internal use only. The information contained herein is HP Confidential. Disclosure is governed by your HP Partner Agreement, HP Retail Partner Agreement, or another applicable contract (e.g., Confidential Disclosure Agreement)."
        
        # Delete content from cell A5
        sheet["A5"].value = None

        # Add autofilters to row 11
        sheet.auto_filter.ref = "A11:J11"
        

        # Save the workbook
        writer.save()
        writer.close()

if __name__ == "__main__":
    main()
