# Task 4: Add computed columns (restock status and inventory value) to the Excel file.

from src.utils import load_workbook, get_sheet, add_computed_columns, save_workbook
from src.config import INPUT_FILE, OUTPUT_FILE, SHEET_NAME

def main():

    # Load workbook and sheet
    workbook = load_workbook(INPUT_FILE)
    sheet = get_sheet(workbook, SHEET_NAME)
    
    # Add computed columns
    add_computed_columns(sheet)
    
    # Save the modified workbook
    save_workbook(workbook, OUTPUT_FILE)

if __name__ == "__main__":
    main()