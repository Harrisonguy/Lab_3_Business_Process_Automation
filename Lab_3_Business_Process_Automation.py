import sys 
import os
from datetime import date
import pandas as pd
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

def get_sales_csv():
    num_params = len(sys.argv) -1
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

def create_orders_dir(sales_csv):

    csv_dir = os.path.dirname(sales_csv)
    todays_date = date.today().isoformat()
    orders_dir_name = f'Orders_{todays_date}'
    orders_dir_path = os.path.join(csv_dir, orders_dir_name)

    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)
    return orders_dir_path

def process_sales_data(sales_csv, orders_dir):
    sales_df = pd.read_csv(sales_csv)
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])
    sales_df.drop(columns=['ADDRESS','CITY','POSTAL CODE','COUNTRY','STATE'],inplace=True)
    for order_id, order_df in sales_df.groupby('ORDER ID'):
        order_df.drop(columns=['ORDER ID'],inplace=True)
        order_df.sort_values(by='ITEM NUMBER',inplace=True)
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE':['GRAND TOTAL'],'TOTAL PRICE':[grand_total]})
        order_df = pd.concat([order_df, grand_total_df])
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W','',customer_name)
        order_file_name = f'Order{order_id}_{customer_name}.xlsx'
        order_file_path = os.path.join(orders_dir, order_file_name)
        sheet_name = f'Order {order_id}'
    
        writer = pd.ExcelWriter(order_file_path, engine='xlsxwriter')
        order_df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]
        money_format = workbook.add_format({'num_format': '$#,##0.00'})
        worksheet.set_column(0, 0, 11, None)
        worksheet.set_column(1, 1, 13, None)
        worksheet.set_column(2, 4, 15, None)
        worksheet.set_column(5, 6, 13, money_format)
        worksheet.set_column(7, 7, 10, None)
        worksheet.set_column(8, 8, 30, None)
        writer.close()
        pass
    return

if __name__ == '__main__':
    main()