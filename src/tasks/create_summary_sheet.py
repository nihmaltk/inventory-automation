# Task 5: Create a summary sheet with supplier statistics.

from src.utils import load_workbook, get_sheet, create_summary_sheet, save_workbook
from src.config import OUTPUT_FILE, SHEET_NAME

def main():

    # Load workbook and sheet
    workbook = load_workbook(OUTPUT_FILE)
    sheet = get_sheet(workbook, SHEET_NAME)
    
    # Create summary sheet
    create_summary_sheet(workbook, sheet)
    
    # Save the workbook
    save_workbook(workbook, OUTPUT_FILE)

if __name__ == "__main__":
    main()