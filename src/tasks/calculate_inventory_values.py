# Task 2: Calculate the total inventory value per supplier.

from src.utils import load_workbook, get_sheet, calculate_inventory_values
from src.config import INPUT_FILE, SHEET_NAME

def main():

    # Load workbook and sheet
    workbook = load_workbook(INPUT_FILE)
    sheet = get_sheet(workbook, SHEET_NAME)
    
    # Calculate inventory values
    total_inventory_value_per_supplier = calculate_inventory_values(sheet)
    
    # Display results
    for supplier, value in total_inventory_value_per_supplier.items():
        print(f"{supplier}: ${value:,.2f}")

if __name__ == "__main__":
    main()
