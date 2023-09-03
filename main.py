import csv
import openpyxl as xl
from pathlib import Path

order_file_path = "test.csv"
dues_book_path = "dues.xlsx"
dues_book_exists = True if Path(dues_book_path).is_file() else False
dues_book = xl.load_workbook(dues_book_path) if dues_book_exists else xl.Workbook()
orders = {}

# if not dues_book_exists:
#     dues_book.create_sheet(title="Alumni")
#     dues_book.create_sheet(title="Grad")
#     dues_book.create_sheet(title="Undergrad")
#     dues_book.create_sheet(title="Staff")

with open(order_file_path, "r", encoding="utf-8") as order_file:
    order_dict = csv.DictReader(order_file)
    first_pass_orders = {}
    for order in order_dict:
        cleaned_order = {key: value for key, value in order.items() if value is not None and value != ""}
        if (order["Status"] == "paid"):
            first_pass_orders[order["Order #"]] = cleaned_order
        if (order["Product Name"] == "Membership Dues" and order["Order #"] in first_pass_orders.keys()):
            first_pass_orders[order["Order #"]].update(cleaned_order)
    orders = {order_dict: order_details for order_dict, order_details in first_pass_orders.items() if "Product Name" in order_details.keys()}
    
    
print(orders)
print(len(orders))
