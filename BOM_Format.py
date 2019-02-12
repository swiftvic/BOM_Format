# BOM formating 
# Original BOM 9010242321 now 123240109
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

headers = ["Level", "Customer Part Number", "Qty", "Ref Des", "Item Rev", "Mara #", "Mara Description", "Cust MFR", "Cust MPN", "Cust Notes", "Higher Level", "Ext Qty"]

# Styling for headers
font_headers = Font(name = 'Arial', size = 10, bold = True)

# Column map from CustBOM
cust_bom_ref = {"Level": 1, "Number": 2, "BOM.Qty": 15, "BOM.Ref": 18, "Rev": 11, "Description": 4, "Manufacturers.MFR Name": 26, "Manufacturers.MPN": 27}

def main():

    filepath_old = "a_20180412_064608761.xlsx"
    filepath_new = "a_20181029_024744353.xlsx"

    wb_old = openpyxl.load_workbook(filepath_old)
    wb_new = openpyxl.load_workbook(filepath_new)

    cust_sheet = wb_old["Sheet0"]                                 # Select Sheet0 as CustBOM
    ws_new_sheet = wb_new["Sheet0"]                               # Select Sheet0 of new file as CustBOM

    mara_format = wb_old.create_sheet("PFormat")                  # Create new sheet called PFormat in old workbook
    
    # Print old wb sheet names
    print(wb_old.sheetnames)

    # Old sheet max rows and columns
    max_row = cust_sheet.max_row
    max_col = cust_sheet.max_column

    # New sheet max rows and columns
    max_new_row = ws_new_sheet.max_row
    max_new_col = ws_new_sheet.max_column

    # Prints stats of each file
    print("There are " + str(max_row) + " line items in " + str(filepath_old))
    print("There are " + str(max_new_row) + " line items in " + str(filepath_new))

    # Copies BOM lvl and Assy p/n to top
    for col in range(1,3):
        mara_format.cell(1, col).value = cust_sheet.cell(2, col).value
    
    # Copies Rev and Description to new format
    mara_format["E1"] = cust_sheet["K2"].value
    mara_format["G1"] = cust_sheet["D2"].value

    # Adding headers into new sheet
    col = 1
    for item in headers:
        mara_format.cell(2, col).value = item                                                           # Copies each header into each cell
        mara_format.cell(2, col).font = font_headers                                                    # Setting the styling of each header
        col += 1

    # Copying various customer columns (ref cust_bom_col variable at top) over to Mara new sheet
    m_col = 1                                                                                           # Start col 1 of new sheet
    for cust_col in cust_bom_ref.values():                                                              # Loop through each item in cust_bom_ref and use it's values
        m_row = 3                                                                                       # Start row 3 of new sheet and reset to 3 after each column is copied                                                                                     
        for cust_row in range(3, max_row):                                                              # Loop through each row in cust sheet until end         
            mara_format.cell(m_row, m_col).value = cust_sheet.cell(cust_row, cust_col).value            # Copy cells from cust sheet cell to mara format cell
            m_row += 1                                                                                  # Increase to new row of new sheet
        m_col += 1                                                                                      # Increase to next col of new sheet
        if m_col == 6 or m_col == 11:                                                                   # Skip col 6 and 11 of new sheet
            m_col += 1                                                                                  

    # Adjust columns to length of cell values
    for col in mara_format.columns:
        max_length = 0
        column = col[0].column                                                                          # Gets the Column letter/name
        for cell in col:
            try:                                                                                        # Avoid error on an empty cell
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 1)
        mara_format.column_dimensions[column].width = adjusted_width

    # Saves changes
    wb_old.save(filepath_old)

if __name__ == '__main__':
    main()