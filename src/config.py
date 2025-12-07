# File paths
INPUT_FILE = "data/inventory_products.xlsx"
OUTPUT_FILE = "data/updated_inventory_products.xlsx"
SHEET_NAME = "products"
SUMMARY_SHEET_NAME = "summary"

# Column indices (using 1-based indexing as in openpyxl)
COL_PRODUCT_ID = 1
COL_PRODUCT_NAME = 2
COL_CATEGORY = 3
COL_SUPPLIER_NAME = 4
COL_PRICE = 5
COL_INVENTORY = 6
COL_REORDER_LEVEL = 7
COL_RESTOCK = 8
COL_INVENTORY_VALUE = 9

# Column names for new columns
HEADER_RESTOCK = "restock"
HEADER_INVENTORY_VALUE = "inventory_value"

# Summary sheet headers
SUMMARY_HEADERS = ["supplier_name", "products_count", "total_inventory_value"]