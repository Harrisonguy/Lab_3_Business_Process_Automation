import sys
import os
from datetime import date
import pandas as pd
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    num_params = len(sys.argv)
    if num_params >= 1:
        csv_path = sys.argv[1]
        if os.path.isfile(csv_path):
            return csv_path
        else:
            print('Error: CSV file does not exist')
            sys.exit(1)
    else:
        print('Error: missing CSV file path')
        sys.exit(1)
    # Check whether command line parameter provided
    # Check whether provide parameter is valid path of file

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):

    csv_dir = os.path.dirname(sales_csv)
    todays_date = date.today().isoformat()
    orders_dir_name = f'Orders_{todays_date}'
    orders_dir_path = os.path.join(csv_dir, orders_dir_name)

    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)
    return orders_dir_path



    # Get directory in which sales data CSV file resides
    # Determine the name and path of the directory to hold the order data files
    # Create the order directory if it does not already exist
    return 

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    sales_df = pd.read_csv(sales_csv)
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])
    sales_df.drop(columns=['ADDRESS','CITY','POSTA CODE','COUNTRY'],inplace=True)
    for order_id, order_df in sales_df.groupby('ORDER ID'):
        order_df.drop(columns=['ORDER ID'],inplace=True)
        order_df.sort_values(by='ITEM NUMBER',inplace=True)
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.concat({'ITEM PRICE':['GRAND TOTAL'],'TOTAL PRICE':[grand_total]})
        order_df = pd.concat([order_df, grand_total_df])
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W','',customer_name)
        order_file_name = f'Order{order_id}_{customer_name}.xlsx'
        order_file_path = os.path.join(orders_dir, order_file_name)
        sheet_name = f'Order {order_id}'
        order_df.to_excel(order_file_path,index=False ,sheet_name=sheet_name)

    # Import the sales data from the CSV file into a DataFrame
    # Insert a new "TOTAL PRICE" column into the DataFrame
    # Remove columns from the DataFrame that are not needed
    # Group the rows in the DataFrame by order ID
    # For each order ID:
        # Remove the "ORDER ID" column
        # Sort the items by item number
        # Append a "GRAND TOTAL" row
        # Determine the file name and full path of the Excel sheet
        # Export the data to an Excel sheet
        # TODO: Format the Excel sheet
        print(order_df)#for debugging
        pass #change to pass later

if __name__ == '__main__':
    main()