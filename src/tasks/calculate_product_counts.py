# Task 1: Calculate the number of products per supplier.

from src.utils import load_workbook, get_sheet, calculate_product_counts
from src.config import INPUT_FILE, SHEET_NAME

def main():

    # Load workbook and sheet
    workbook = load_workbook(INPUT_FILE)
    sheet = get_sheet(workbook, SHEET_NAME)
    
    # Calculate product counts
    product_count_per_supplier = calculate_product_counts(sheet)
    
    # Display results
    for supplier, count in product_count_per_supplier.items():
        print(f"{supplier}: {count} products")

if __name__ == "__main__":
    main()

