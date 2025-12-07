# Task 3: Detect products with low inventory

from src.utils import load_workbook, get_sheet, get_low_inventory_products
from src.config import INPUT_FILE, SHEET_NAME

def main():

    # Load workbook and sheet
    workbook = load_workbook(INPUT_FILE)
    sheet = get_sheet(workbook, SHEET_NAME)
    
    # Get low inventory products
    low_inventory_products = get_low_inventory_products(sheet)
    
    # Display results
    if low_inventory_products:
        for product in low_inventory_products:
            print(f"\nProduct ID: {product['product_id']}")
            print(f"  Name: {product['product_name']}")
            print(f"  Supplier: {product['supplier_name']}")
            print(f"  Current Inventory: {product['inventory']}")
            print(f"  Reorder Level: {product['reorder_level']}")
            print(f"  Shortage: {product['shortage']}")
    else:
        print("No products with low inventory.")
    
    print(f"\nTotal products needing restock: {len(low_inventory_products)}")

if __name__ == "__main__":
    main()
