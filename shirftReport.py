import pyodbc
import tkinter as tk
from tkinter import ttk, messagebox
from configparser import ConfigParser
from datetime import datetime
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from configparser import ConfigParser
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from matplotlib.table import Table as MPLTable
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from tkcalendar import Calendar, DateEntry
import tkinter as tk
from tkinter import ttk, messagebox
from configparser import ConfigParser
import os
import subprocess
from datetime import datetime, timedelta
import uuid
from reportlab.lib import pagesizes
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
import json
from configparser import ConfigParser
import argparse
from datetime import datetime, timedelta

# Parse command line arguments
parser = argparse.ArgumentParser(description='Generate shift report.')
parser.add_argument('--start_datetime', type=str, help='Start date and time in the format YYYY-MM-DD HH:MM')
parser.add_argument('--end_datetime', type=str, help='End date and time in the format YYYY-MM-DD HH:MM')

args = parser.parse_args()

# If not provided, use current date and time
if args.start_datetime:
    start_datetime = args.start_datetime
else:
    start_datetime = datetime.now().strftime("%Y-%m-%d %H:%M")

if args.end_datetime:
    end_datetime = args.end_datetime
else:
    end_datetime = datetime.now().strftime("%Y-%m-%d %H:%M")

def load_config():
    # Loading database configuration from ini file
    config = ConfigParser()
    config.read("config.ini")
    database_config = config["DATABASE"]
    
    # Split the codes string into a list
    if config.has_section("DISCOUNT_CODES"):
        codes_string = config.get("DISCOUNT_CODES", "codes", fallback="")
        discount_codes = codes_string.split(",")
    else:
        discount_codes = []

    # Loading additional configuration from json file (if needed)
    # Note: It's not clear why you were using a JSON file in addition to the ini file.
    # If it's not necessary, you can remove this section.
    try:
        with open('config.json') as f:
            json_config = json.load(f)
    except FileNotFoundError:
        json_config = {}

    return database_config, discount_codes, json_config


# Example usage:
database_config, discount_codes, json_config = load_config()

# Connecting to the database
server = database_config.get("server")
database = database_config.get("database")
username = database_config.get("username")
password = database_config.get("password")
auth_type = database_config.get("auth_type")



def connect_db(server, database, username, password, auth_type):
    if auth_type == "SQL":
        connection = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}")
    else:
        connection = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes")
    return connection


def fetch_report(connection, start_datetime, end_datetime):
    query = f"""
        SELECT
            D.Description AS "DRYSTOCK SALES",
            SUM(TL.Quantity) AS QTY,
            SUM(TH.TotalAfterTax) AS SALE
        FROM
            TransLines TL
        INNER JOIN
            Items I ON TL.UPC = I.UPC
        INNER JOIN
            Departments D ON I.Department = D.ID
        INNER JOIN
            TransHeaders TH ON TL.TransNo = TH.TransNo AND TL.Branch = TH.Branch AND TL.Station = TH.Station
        WHERE
            TH.Logged BETWEEN ? AND ?
            AND D.Description != 'Wet Stock'
            AND TL.Quantity >= 0
            AND TH.TotalAfterTax >= 0
        GROUP BY
            D.Description
        WITH ROLLUP
    """
    cursor = connection.cursor()
    cursor.execute(query, (start_datetime, end_datetime))
    return cursor.fetchall()

def fetch_negative_values(connection, start_datetime, end_datetime):
    query = f"""
        SELECT
            D.Description AS "DRYSTOCK SALES",
            SUM(TL.Quantity) AS QTY,
            SUM(TH.TotalAfterTax) AS SALE
        FROM
            TransLines TL
        INNER JOIN
            Items I ON TL.UPC = I.UPC
        INNER JOIN
            Departments D ON I.Department = D.ID
        INNER JOIN
            TransHeaders TH ON TL.TransNo = TH.TransNo AND TL.Branch = TH.Branch AND TL.Station = TH.Station
        WHERE
            TH.Logged BETWEEN ? AND ?
            AND (TL.Quantity < 0 OR TH.TotalAfterTax < 0)
            AND D.Description != 'Wet Stock'
        GROUP BY
            D.Description
        WITH ROLLUP
    """
    cursor = connection.cursor()
    cursor.execute(query, (start_datetime, end_datetime))
    return cursor.fetchall()



def fetch_wetstock(connection, start_datetime, end_datetime):
    query = f"""
        SELECT
            SDesc.Description AS WETSTOCK,
            SUM(TL.Quantity) AS VOLUME,
            SUM(TL.SubAfterTax) AS SALE
        FROM
            TransLines TL
        INNER JOIN
            Items I ON TL.UPC = I.UPC
        INNER JOIN
            SubDepartments SDesc ON I.SubDepartment = SDesc.ID
        INNER JOIN
            TransHeaders TH ON TL.TransNo = TH.TransNo AND TL.Branch = TH.Branch AND TL.Station = TH.Station
        WHERE
            SDesc.DepID = 2
            AND TH.Logged BETWEEN ? AND ?
        GROUP BY
            SDesc.Description
    """
    cursor = connection.cursor()
    cursor.execute(query, (start_datetime, end_datetime))
    return cursor.fetchall()


def fetch_payment_totals(connection, start_datetime, end_datetime):
    query = f"""
        SELECT
            TP.MediaName AS "PAYMENT TOTALS",
            COUNT(*) AS QTY,
            SUM(TP.Value) AS AMOUNT
        FROM
            TransPayments TP
        INNER JOIN
            TransHeaders TH ON TP.TransNo = TH.TransNo AND TP.Branch = TH.Branch AND TP.Station = TH.Station
        WHERE
            TH.Logged BETWEEN ? AND ?
        GROUP BY
            TP.MediaName
    """

    cursor = connection.cursor()
    cursor.execute(query, (start_datetime, end_datetime))
    return cursor.fetchall()


def fetch_sale_totals(connection, start_datetime, end_datetime, excluded_upcs):
    # Make a tuple from the excluded_upcs list
    excluded_upcs_str = ', '.join(f"'{item}'" for item in excluded_upcs)
    query = f"""
        -- Discounts
        SELECT
            'DISCOUNTS' AS "SALE TOTALS",
            COUNT(DISTINCT TL.TransNo) AS "TXN#",
            SUM(TL.SubAfterTax) AS AMOUNT
        FROM
            TransLines TL
        INNER JOIN
            TransHeaders TH ON TL.TransNo = TH.TransNo AND TL.Branch = TH.Branch AND TL.Station = TH.Station
        WHERE
            TL.DiscountID > 0
            AND TH.Logged BETWEEN ? AND ?
        UNION ALL
        -- Voided
        SELECT
            'VOIDED' AS "SALE TOTALS",
            COUNT(DISTINCT TL.TransNo) AS "TXN#",
            SUM(TL.SubAfterTax) AS AMOUNT
        FROM
            TransLines TL
        INNER JOIN
            TransHeaders TH ON TL.TransNo = TH.TransNo AND TL.Branch = TH.Branch AND TL.Station = TH.Station
        WHERE
            TL.IsVoided = 1
            AND TH.Logged BETWEEN ? AND ?
        UNION ALL
        -- Refunded
        SELECT
            'REFUNDED' AS "SALE TOTALS",
            COUNT(DISTINCT TL.TransNo) AS "TXN#",
            SUM(TL.SubAfterTax) AS AMOUNT
        FROM
            TransLines TL
        INNER JOIN
            TransHeaders TH ON TL.TransNo = TH.TransNo AND TL.Branch = TH.Branch AND TL.Station = TH.Station
        WHERE
            TL.SubAfterTax < 0
            AND TL.UPC NOT IN ({excluded_upcs_str})
            AND TH.Logged BETWEEN ? AND ?
    """
    # ... rest of the code

    cursor = connection.cursor()
    # Reuse the date range parameters for each part of the UNION query.
    cursor.execute(query, (start_datetime, end_datetime, start_datetime, end_datetime, start_datetime, end_datetime))
    

    return cursor.fetchall()

def fetch_yesterday_report():
    # Calculate yesterday's date
    yesterday_date = datetime.now() - timedelta(days=1)
    # Format the start and end times
    start_date = str(yesterday_date.date())
    end_date = str(yesterday_date.date())
    start_time = "00:00:00"
    end_time = "23:59:59"

    # Create a temporary file to save the PDF
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

    # Trigger the report generation with the above date and time
    generate_report(file_path, start_date, start_time, end_date, end_time)


    
def generate_report(file_path, start_date, start_time, end_date, end_time):
    config, excluded_upcs, json_config = load_config()

    server = config.get("server")
    database = config.get("database")
    username = config.get("username")
    password = config.get("password")
    auth_type = config.get("auth_type")

    # If not provided, use yesterday's date from 00:00 to 23:59
    if not start_date or not start_time:
        yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
        start_date = yesterday.strftime("%Y-%m-%d")
        start_time = "00:00"

    if not end_date or not end_time:
        yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
        end_date = yesterday.strftime("%Y-%m-%d")
        end_time = "23:59"
    
    # Concatenate date and time for the queries
    start_datetime = start_date + " " + start_time
    end_datetime = end_date + " " + end_time

    try:
        connection = connect_db(server, database, username, password, auth_type)
        report = fetch_report(connection, start_date + " " + start_time, end_date + " " + end_time)
        # Rest of your code

        # Build the document with all elements

        

        # Fetch additional information
        cursor = connection.cursor()
        cursor.execute("SELECT TOP 1 ShiftId, Branch, Station FROM TransHeaders WHERE Logged BETWEEN ? AND ?", (start_datetime, end_datetime))
        additional_info = cursor.fetchone()
        
        # Ask user for location to save the PDF
        # Ask user for location to save the PDF if it is not already provided
        
        #file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        
        # Define 80mm width (approximately 3.15 inches) and 11 inches height
        custom_pagesize = (80 * mm, 11 * 72)
        
        # Generate PDF
        if file_path:
            
            doc = SimpleDocTemplate(file_path, pagesize=custom_pagesize, leftMargin=5*mm, rightMargin=0, topMargin=10)
                    # Define table styles
            thin_border = 0.5
            table_style = TableStyle([
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),  # Use smaller font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduce padding
            ('GRID', (0, 0), (-1, -1), thin_border, colors.black)
            ])

            table_neg_style = TableStyle([
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),  # Use smaller font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduce padding
            ('GRID', (0, 0), (-1, -1), thin_border, colors.black)
            ])

            wetstock_table_style = TableStyle([
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),  # Use smaller font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduce padding
            ('GRID', (0, 0), (-1, -1), thin_border, colors.black)
            ])

            elements = []

            # Load styles and left align paragraph text
            styles = getSampleStyleSheet()
            # Define paragraph styles
            normal_style = getSampleStyleSheet()["Normal"]
            normal_style.alignment = 0  # 0 is LEFT alignment

            title_style = getSampleStyleSheet()["Heading1"]
            title_style.alignment = 1  # 1 is CENTER alignment

            # Add the title to the top
            elements.append(Paragraph("Shift Report", title_style))
            elements.append(Spacer(1, 12))
            if additional_info:
                shift_id, branch, station = additional_info
                elements.append(Paragraph(f"Shift ID: {shift_id}", normal_style))
                elements.append(Paragraph(f"Branch: {branch}", normal_style))
                elements.append(Paragraph(f"Station: {station}", normal_style))
            else:
                elements.append(Paragraph(f"No shift information found.", normal_style))

            elements.append(Paragraph(f"Start Date and Time: {start_datetime}", normal_style))
            elements.append(Paragraph(f"End Date and Time: {end_datetime}", normal_style))
            elements.append(Spacer(1, 12))
            wet_stock_sale_total = 0
            # Include table
            data = [["DRYSTOCK SALES", "QTY", "SALE"]]
            total_qty = 0
            total_sale = 0
            for row in report:
                desc, qty, sale = row
                if desc is None:
                    desc = "DRYSTOCK SALES TOTAL"
                    total_qty = qty
                    total_sale = sale

                data.append([desc, f"{qty:.2f}".rstrip("0").rstrip("."), f"{sale:.2f}".rstrip("0").rstrip(".")])

# ... rest of the code

            # Table style
            # ... table style settings
            col_widths = (30 * mm, 13 * mm, 13 * mm)
            elements.append(Table(data, style=table_style, hAlign='LEFT', colWidths=col_widths))
            elements.append(Spacer(1, 6))
            
                        # Initialize these outside the if blocks
            # Initialize these variables
            # Initialize these variables
            refund_qty = 0
            refund_sale = 0

            # Always create table headers
            data_neg = [["DRYSTOCK SALES", "QTY", "SALE"]]

            # Include table for negative values
            negative_values = fetch_negative_values(connection, start_datetime, end_datetime)
            if negative_values:
                for row in negative_values:
                    desc, qty, sale = row
                    if desc is None:
                        desc = "DRYSTOCK Refund"
                        refund_qty += qty  # Accumulate qty for refund
                        refund_sale += sale  # Accumulate sale for refund

                    data_neg.append([desc, f"{qty:.2f}".rstrip("0").rstrip("."), f"{sale:.2f}".rstrip("0").rstrip(".")])

            if refund_qty > 0:
                refund_qty *= -1
            if refund_sale > 0:
                refund_sale *= -1
            # Calculate NET DRYSTOCK TOTAL
            net_drystock_qty = total_qty + refund_qty
            net_drystock_sale = total_sale + refund_sale  # total_sale should be calculated before this block

            # Add NET DRYSTOCK TOTAL

            data_neg.append(["NET DRYSTOCK TOTAL", f"{net_drystock_qty:.2f}".rstrip("0").rstrip("."), f"{net_drystock_sale:.2f}".rstrip("0").rstrip(".")])

            elements.append(Table(data_neg, style=table_neg_style, hAlign='LEFT', colWidths=col_widths))
            elements.append(Spacer(1, 6))

            # Calculate SALES TOTAL before appending to WetStock table
            sales_total = wet_stock_sale_total + net_drystock_sale

            # Include table for wet stock
            # Calculate wet_stock_sale_total before the if block
            

            # Fetch wet_stock_info
            wet_stock_info = fetch_wetstock(connection, start_datetime, end_datetime)

            # Proceed if wet_stock_info is not empty
            if wet_stock_info:
                wet_stock_data = [["WETSTOCK", "VOLUME", "SALE"]]
                wet_stock_volume_total = 0

                # Loop through the wet_stock_info to populate wet_stock_data
                for row in wet_stock_info:
                    wet_stock, volume, sale = row
                    wet_stock_volume_total += volume
                    wet_stock_sale_total += sale
                    wet_stock_data.append([wet_stock, f"{volume:.2f}".rstrip("0").rstrip("."), f"{sale:.2f}".rstrip("0").rstrip(".")])

                # Calculate SALES TOTAL (WETSTOCK SALES TOTAL + NET DRYSTOCK TOTAL)
                sales_total = wet_stock_sale_total + net_drystock_sale

                # Append totals to wet_stock_data
                wet_stock_data.append(["WETSTOCK SALES TOTAL", f"{wet_stock_volume_total:.2f}".rstrip("0").rstrip("."), f"{wet_stock_sale_total:.2f}".rstrip("0").rstrip(".")])
                # wet_stock_data.append(["SALES TOTAL", "", f"{sales_total:.2f}".rstrip("0").rstrip(".")])
                
                # Append wet_stock_data to elements
                elements.append(Spacer(1, 12))
                elements.append(Table(wet_stock_data, style=wetstock_table_style, hAlign='LEFT', colWidths=col_widths))

            # Create a separate table for SALES TOTAL
            sales_total_table_style = TableStyle([
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 6),  # Use smaller font size
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduce padding
                ('GRID', (0, 0), (-1, -1), thin_border, colors.black)
            ])

            sales_total_data = [["SALES TOTAL", "", f"{sales_total:.2f}".rstrip("0").rstrip(".")]]
            elements.append(Spacer(1, 12))
            elements.append(Table(sales_total_data, style=sales_total_table_style, hAlign='LEFT', colWidths=col_widths))




            payment_totals_info = fetch_payment_totals(connection, start_datetime, end_datetime)
            if payment_totals_info:
                payment_totals_data = [["PAYMENT TOTALS", "QTY", "AMOUNT"]]
                payment_totals_qty_total = 0
                payment_totals_amount_total = 0
                for row in payment_totals_info:
                    payment_total, qty, amount = row
                    payment_totals_qty_total += qty
                    payment_totals_amount_total += amount
                    payment_totals_data.append([payment_total, str(qty), f"{amount:.2f}".rstrip("0").rstrip(".")])
                payment_totals_data.append(["PAYMENT TOTAL", f"{payment_totals_qty_total}", f"{payment_totals_amount_total:.2f}".rstrip("0").rstrip(".")])
                elements.append(Spacer(1, 12))
                elements.append(Table(payment_totals_data, style=table_style,hAlign='LEFT', colWidths=col_widths))

            # Fetch sale totals and include as a new table
            sale_totals = fetch_sale_totals(connection, start_datetime, end_datetime, excluded_upcs)
            if sale_totals:
                sale_totals_data = [["SALE TOTALS", "QTY", "AMOUNT"]]
                for row in sale_totals:
                    sale_total, txn, amount = row
                    sale_totals_data.append([sale_total, str(txn), f"{amount:.2f}".rstrip("0").rstrip(".") if amount is not None else '0'])
                elements.append(Spacer(1, 12))
                elements.append(Table(sale_totals_data, style=table_style, hAlign='LEFT', colWidths=col_widths))

            # # Build the document with all elements
            # doc.build(elements)

            # messagebox.showinfo("Info", "Report generated successfully")

    # Build the document with all elements


            # Build the document with all elements
            doc.build(elements)
            subprocess.run([r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe", "-print-to-default", file_path], shell=True)
            
    except Exception as e:
        messagebox.showerror("Error", str(e))

fetch_yesterday_report()


# # GUI
# root = tk.Tk()
# root.title("Shift Report Generator")
# root.geometry("400x350")
# # Date and Time configuration
# start_date_label = tk.Label(root, text="Start Date:")
# start_date_label.pack()

# start_date = DateEntry(root)
# start_date.pack()

# start_time_label = ttk.Label(root, text="Start Time (HH:MM:SS):")
# start_time_label.pack()
# start_time = ttk.Entry(root)
# start_time.pack()

# end_date_label = ttk.Label(root, text="End Date:")
# end_date_label.pack()

# end_date = DateEntry(root)
# end_date.pack()

# end_time_label = ttk.Label(root, text="End Time (HH:MM:SS):")
# end_time_label.pack()
# end_time = ttk.Entry(root)
# end_time.pack()

# def update_config():
#     config_win = tk.Toplevel(root)
#     config_win.title("Update Config")

#     # Server configuration
#     server_label = ttk.Label(config_win, text="Server:")
#     server_label.pack()
#     server_entry = ttk.Entry(config_win)
#     server_entry.pack()

#     database_label = ttk.Label(config_win, text="Database:")
#     database_label.pack()
#     database_entry = ttk.Entry(config_win)
#     database_entry.pack()

#     username_label = ttk.Label(config_win, text="Username:")
#     username_label.pack()
#     username_entry = ttk.Entry(config_win)
#     username_entry.pack()

#     password_label = ttk.Label(config_win, text="Password:")
#     password_label.pack()
#     password_entry = ttk.Entry(config_win, show="*")
#     password_entry.pack()

#     auth_label = ttk.Label(config_win, text="Auth Type (SQL/Windows):")
#     auth_label.pack()
#     auth_type_entry = ttk.Entry(config_win)
#     auth_type_entry.pack()

#     # Discount codes as a comma-separated list
#     discount_codes_label = ttk.Label(config_win, text="Fuel Delivery Codes (comma-separated):")
#     discount_codes_label.pack()
#     discount_codes_entry = ttk.Entry(config_win)
#     discount_codes_entry.pack()

#     def save_config():
#         server = server_entry.get()
#         database = database_entry.get()
#         username = username_entry.get()
#         password = password_entry.get()
#         auth_type = auth_type_entry.get()

#         discount_codes = discount_codes_entry.get()

#         config = ConfigParser()
#         config["DATABASE"] = {
#             "server": server,
#             "database": database,
#             "username": username,
#             "password": password,
#             "auth_type": auth_type
#         }

#         config["DISCOUNT_CODES"] = {
#             "codes": discount_codes
#         }

#         with open("config.ini", "w") as config_file:
#             config.write(config_file)

#         config_win.destroy()
#         messagebox.showinfo("Success", "Configuration Updated Successfully!")

#     save_button = ttk.Button(config_win, text="Save", command=save_config)
#     save_button.pack()

# update_config_button = ttk.Button(root, text="Update Config", command=update_config)
# update_config_button.pack()

# generate_button = ttk.Button(root, text="Generate Report", command=generate_report)
# generate_button.pack()

# yesterday_report_button = ttk.Button(root, text="Generate Yesterday's Report", command=fetch_yesterday_report)
# yesterday_report_button.pack()

# root.mainloop()
