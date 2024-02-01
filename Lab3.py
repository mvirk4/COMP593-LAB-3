"""description: this program creates an Excel file for each order and exports the data from a sales data CSV file to the Excel file.
Author: Manjot Singh
Date: 2024-01-31"""
import sys
import os
from datetime import date
import pandas as  pd

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    if len(sys.argv) < 2:
        print("ERROR: Missing CSV file path.")
        sys.exit()
    # Check whether provide parameter is valid path of file
        if not os.path.isfile(sys.argv[1]):
            print("ERROR: Invalid invalid CSV path.")
            sys.exit()
    return sys.argv[1]

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    sales_csv_path = os.path.abspath()(sales_csv)
    sales_csv_dir = os.path.dirname(sales_csv_path)

    # Determine the name and path of the directory to hold the order data files
    current_date = date.today().isoformat()
    folder_name = f"orders_{current_date}"
    order_dir = os.path.join(sales_csv_dir, folder_name)

    # Create the order directory if it does not already exist
    if not os.path.isdir(order_dir):
        os.mkdir(order_dir)

    return order_dir

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_df = pd.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, "TOTAL PRICE", sales_df["QUANTITY"] * sales_df["PRICE"])
    
    # Remove columns from the DataFrame that are not needed
    sales_df.drop(columns =["ADDRESS","CITY","STATE","POSTAL CODE","COUNTRY"], inplace = True)
    # Group the rows in the DataFrame by order ID
    grouped_df = sales_df.groupby("ORDER ID")

    # For each order ID:
    for order_id, order_df in grouped_df:
        # Remove the "ORDER ID" column
        order_df.drop(columns=["ORDER ID"], inplace=True)
        # Sort the items by item number
        order_df.sort_values(by="ITEM NUMBER", inplace=True)
        # Append a "GRAND TOTAL" row
        total_price = order_df["TOTAL PRICE"].sum()
        grand_total_row = pd.DataFrame({
            "PRODUCT CODE": ["GRAND TOTAL"],
            "QUANTITY": [0],
            "PRICE": [total_price],
            "TOTAL PRICE": [total_price],
        }, index=[len(order_df) + 1])
        order_df = order_df.append(grand_total_row)
        # Determine the file name and full path of the Excel sheet
        order_num = order_df["ITEM NUMBER"].iloc[0]
        order_file = f"order_{order_num:06d}.xlsx"
        order_path = os.path.join(orders_dir, order_file)
    
        # Export the data to an Excel sheet
        order_df.to_excel(order_path, index=False)
        # TODO: Format the Excel sheet
    pass

if __name__ == '__main__':
    main()