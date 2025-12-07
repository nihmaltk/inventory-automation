from src.tasks import calculate_product_counts
from src.tasks import calculate_inventory_values
from src.tasks import detect_low_inventory
from src.tasks import add_computed_columns
from src.tasks import create_summary_sheet

def main():
    
    try:
        # Task 1: Calculate product counts per supplier
        print("\n[Task 1/5] Calculating product counts...")
        calculate_product_counts.main()
        
        # Task 2: Calculate inventory values per supplier
        print("\n[Task 2/5] Calculating inventory values...")
        calculate_inventory_values.main()
        
        # Task 3: Detect low inventory products
        print("\n[Task 3/5] Detecting low inventory products...")
        detect_low_inventory.main()
        
        # Task 4: Add computed columns
        print("\n[Task 4/5] Adding computed columns...")
        add_computed_columns.main()
        
        # Task 5: Create summary sheet
        print("\n[Task 5/5] Creating summary sheet...")
        create_summary_sheet.main()
           
    except Exception as e:
        print(f"\n Error occurred: {e}")
        print("Please check your input file and try again.")

if __name__ == "__main__":
    main()


