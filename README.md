## Inventory Automation

A Python-based inventory management automation tool that analyzes product data, tracks supplier statistics, and identifies low inventory items from Excel files.

## Features

This tool automates essential inventory management workflows:

* **Low Stock Detection:** Automatically identifies products where **Inventory < Reorder Level**
* **Inventory Valuation:** Calculates the **total inventory value** per product and generates total value per supplier.
* **Supplier Analysis:** Counts the **total number of products** supplied by each vendor.
* **Data Enhancement:** Adds computed columns (`Restock Status`, `Inventory Value`) for better analysis.
* **Summary Reports:** Generates a **comprehensive supplier statistics sheet** in the output file.

## Tech Stack

- Python 
- openpyxl library

## Installation

1. Clone the repository:
```bash
git clone https://github.com/nihmaltk/inventory-automation.git
cd inventory-automation
```
2. Create and Activate the Virtual Environment:
```bash
python -m venv venv

To activate the environment:
   - Windows: `venv\Scripts\activate`
   - macOS/Linux: `source venv/bin/activate`
```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

**Run all tasks at once:**
```bash
python main.py
```

**Run individual tasks:**
```bash
python -m src.tasks.calculate_product_counts
python -m src.tasks.calculate_inventory_values
python -m src.tasks.detect_low_inventory
python -m src.tasks.add_computed_columns
python -m src.tasks.create_summary_sheet
```

## Input Data Format
The application expects an Excel file containing a sheet named **`products`** with the following required columns:

Column         |     Description              |
-------------- | ---------------------------- |
Product ID     |     Unique identifier        |
Product Name   |     Name of the product      |
Category       |     Product category         |
Supplier Name  |     Name of supplier         |
Price          |     Unit price               |
Inventory      |     Current stock quantity   |
Reorder Level  |     Minimum stock threshold  |

## Output & Reports

The tool generates a new Excel file named **`updated_inventory_products.xlsx`** containing:

1.  **`products` Sheet:** Updated with two new calculated columns:
    * **`restock`**: (Yes/No)
    * **`inventory_value`**: (Price * Inventory)
    
2.  **`supplier_summary` Sheet:** A new sheet containing aggregated statistics:
    * Products per supplier
    * Total inventory value per supplier

Console output provides status updates as each task completes.

## Key Concepts

This project demonstrates:

* **Python for Data Automation:** Practical application of Python in manipulating business data (Excel).
* **Excel Manipulation:** Using the `openpyxl` library to read, write, and modify complex Excel files.
* **Modular Architecture:** Organizing code into separate tasks and modules for maintainability.
* **Python Best Practices:** Proper use of **virtual environments** and a professional project structure.

