# Excel Automation Project

## Overview

This project automates tasks related to inventory management using Python and Excel. The primary script processes an input inventory file, performs necessary calculations, and outputs an updated inventory file.

## Files

- `main.py`: The main Python script that performs the automation.
- `Inventory.xlsx`: The input Excel file containing the initial inventory data.
- `updated_inventory_totals.xlsx`: The output Excel file with the updated inventory totals.

## Requirements

- Python 3.x
- pandas
- openpyxl

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/Excel_Automation.git
    cd Excel_Automation
    ```

2. Install the required packages:
    ```sh
    pip install pandas openpyxl
    ```

## Usage

1. Ensure the `Inventory.xlsx` file is placed in the same directory as `main.py`.
2. Run the Python script:
    ```sh
    python main.py
    ```
3. The script will generate an `updated_inventory_totals.xlsx` file with the updated inventory data.

## Functionality

The script `main.py` performs the following tasks:

1. Reads the initial inventory data from `Inventory.xlsx`.
2. Processes the data to update inventory totals.
3. Writes the updated data to `updated_inventory_totals.xlsx`.

## Example

Before running the script:
- Input file: `Inventory.xlsx`

After running the script:
- Output file: `updated_inventory_totals.xlsx`

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
