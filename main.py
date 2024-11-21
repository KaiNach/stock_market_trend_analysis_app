import os

import requests
import ttkbootstrap as ttk
from ttkbootstrap import Style
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from nsetools import Nse
import tkinter as tk
from ttkbootstrap import Style
from ttkbootstrap.widgets import Treeview
import pandas as pd
from tkinter import filedialog, messagebox
# Set high-contrast theme
style = Style(theme="darkly")
# Get ticker symbol using Yahoo Finance API
import requests

import requests

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import json
import sys
app_dir = os.path.dirname(sys.executable)  # Get the folder where the app is running
SAVE_FILE_PATH = os.path.join(app_dir, 'saved_file.json')
# Define the path for saving the uploaded file path

def save_file_path(file_path):
    """Save the uploaded file path to a JSON file."""
    with open(SAVE_FILE_PATH, "w") as f:
        json.dump({"file_path": file_path}, f)

def load_saved_file_path():
    """Load the previously saved file path, if it exists."""
    if os.path.exists(SAVE_FILE_PATH):
        with open(SAVE_FILE_PATH, "r") as f:
            data = json.load(f)
            return data.get("file_path", None)
    return None
# Function to open a new window with the stock's historical data

def initialize_table():
    saved_file_path = load_saved_file_path()
    if saved_file_path and os.path.exists(saved_file_path):
        try:
            # Load the saved Excel file
            df = pd.read_excel(saved_file_path)
            update_table_data()
            display_table(df)
            print(f"Loaded saved file: {saved_file_path}")
        except Exception as e:
            messagebox.showerror("File Error", f"Could not load saved file: {e}")
    else:
        print("No saved file found. Please upload an Excel file.")
def show_historical_data(event):
    # Get the item (row) that was double-clicked
    item = treeview.selection()[0]
    column_index = treeview.identify_column(event.x)  # Get the column index of the clicked cell

    # Only proceed if the double-click was on the "Instrument" column (the first column)
    if column_index == "#1":
        instrument_name = treeview.item(item, "values")[0]  # Assuming the "Instrument" is in the first column

        # Fetch the ticker symbol for the instrument name (e.g., "IDEA")
        ticker_symbol = get_ticker_name(instrument_name)

        if ticker_symbol != "N/A":
            # Fetch the historical data for the stock using yfinance
            stock_data = yf.Ticker(ticker_symbol)
            historical_data = stock_data.history(period="1y")  # Last 1 year of data

            # Create a plot for the historical data
            fig, ax = plt.subplots(figsize=(8, 6))
            ax.plot(historical_data.index, historical_data['Close'], label=f"{ticker_symbol} Closing Price",
                    color='blue')
            ax.set_title(f"Historical Trend of {ticker_symbol}")
            ax.set_xlabel('Date')
            ax.set_ylabel('Price (INR)')
            ax.legend()

            # Create a new window to display the graph
            top = tk.Toplevel(root)
            top.title(f"Stock Trend for {ticker_symbol}")

            # Embed the plot in the new window using a canvas
            canvas = FigureCanvasTkAgg(fig, master=top)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # Start the new window's main loop
            top.mainloop()


# Function to fetch the ticker name for a given company
def get_ticker_name(company_name, exchange_code="NS"):
    try:
        # Appending the exchange code for more accurate results (e.g., ".NS" for NSE-listed stocks)
        search_query = f"{company_name}.{exchange_code}"
        search_url = f"https://query1.finance.yahoo.com/v1/finance/search?q={search_query}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept": "application/json",
        }

        # Making the request
        response = requests.get(search_url, headers=headers)
        response.raise_for_status()
        data = response.json()

        if "quotes" in data and data["quotes"]:
            return data['quotes'][0].get('symbol', 'N/A')  # Return the first symbol in the results
        return "N/A"

    except requests.exceptions.RequestException as e:
        print("Error fetching data for", company_name, ":", e)
        return "N/A"

# Refresh prices
def refresh_prices():
    for row_id in treeview.get_children():
        row_data = list(treeview.item(row_id, "values"))
        company_ticker = row_data[1]
        if not company_ticker or company_ticker == "NA":
            continue

        try:
            stock = yf.Ticker(company_ticker)
            todays_data = stock.history()
            if len(todays_data) >= 2:
                current_price = todays_data['Close'].iloc[-1]
                previous_price = todays_data['Close'].iloc[-2]
                net_change = current_price - previous_price
                net_change_percentage = (net_change / previous_price) * 100 if previous_price else 0
                new_values = row_data
                new_values[5] = f"{(float(new_values[3]) * float(new_values[4])):.2f}"
                new_values[6] = f"{current_price:.2f}"  # LTP
                new_values[7] = f"{float(row_data[3]) * float(current_price):.2f}"  # Cur. Val
                new_values[11] = f"{net_change_percentage:.2f}"  # Net Chg. in %
                new_values[9] = f"{(float(new_values[7]) - float(new_values[5])):.2f}"  # P&L
                new_values[10] = f"{(float(new_values[9]) / float(new_values[5]) *100):.2f}"
                new_values[13] = f"{(float(new_values[6]) - float(new_values[4])):.2f}"
                update_num_days()
                # Update treeview item
                treeview.item(row_id, values=new_values)

                # Conditional formatting for P&L
                p_and_l = float(new_values[9])
                p_and_l_tag = "profit" if p_and_l >= 0 else "loss"
                treeview.item(row_id, tags=(p_and_l_tag,))
                update_totals()
                current_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                timestamp_label.config(text=f"Values are as of {current_time}")
            else:
                print(f"Not enough data for {company_ticker} to calculate net change.")
        except Exception as e:
            print(f"Error fetching price for {company_ticker}: {e}")
# Load Excel file data
def save_to_file():
    try:
        # Open file dialog for user to choose location
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

        # If the user cancels the file dialog
        if not save_path:
            return

        # Get data from the table and save to the chosen Excel file
        rows = []
        for item in treeview.get_children():
            row_data = treeview.item(item, "values")
            rows.append(row_data)

        # Assuming `treeview["columns"]` contains the column names
        df = pd.DataFrame(rows, columns=treeview["columns"])

        # Save DataFrame to the specified path
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Save Successful", "Portfolio saved successfully!")
    except Exception as e:
        messagebox.showerror("Save Error", f"Could not save file: {e}")


def load_saved_file():
    try:
        # Open file dialog for user to choose file to load
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

        # If the user cancels the file dialog
        if not file_path:
            return

        # Check if the file exists and the user has read-write permissions
        if os.path.exists(file_path) and os.access(file_path, os.R_OK):
            # Load the saved Excel file
            df = pd.read_excel(file_path)

            # Ensure that the columns in the DataFrame match the expected columns in the Treeview
            expected_columns = treeview["columns"]
            if all(col in df.columns for col in expected_columns):
                # Clear the existing Treeview data
                for row in treeview.get_children():
                    treeview.delete(row)

                # Insert new rows from the loaded file into the Treeview
                for _, row in df.iterrows():
                    treeview.insert("", "end", values=row.tolist())
                messagebox.showinfo("Load Successful", "Portfolio loaded successfully!")
            else:
                messagebox.showerror("Load Error", "Loaded file does not match the expected format.")
        else:
            messagebox.showerror("Load Error", "File does not exist or cannot be read.")

    except Exception as e:
        messagebox.showerror("Load Error", f"Could not load the file: {e}")

def update_table_data():
    # Get all the rows from the Treeview
    rows = []
    for item in treeview.get_children():
        values = treeview.item(item, "values")
        rows.append(values)

    # Update your DataFrame or internal list with the new values
    df = pd.DataFrame(rows,
                      columns=["Instrument", "Company Ticker","Date", "Qty.", "Avg. cost", "Invest Value", "LTP", "Cur. Val",
                               "P&L", "Net chg.", "Day chg.", "Total P&L", "Num Days", "Price chg."])
    # Update calculations, e.g., re-calculate Invest Value
    for index, row in df.iterrows():
        df.at[index, 'invest value'] = row['qty.'] * row['avg. cost']

    # After updating internal data, call the display_table to refresh Treeview
    display_table(df)

def initialize_app():
    # Attempt to load the previously saved portfolio
    load_saved_file()
def upload_file():
    """Handle the file upload and process the Excel file."""
    global uploaded_file_path  # To save the file path for reuse
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        try:
            # Load the Excel file
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip().str.lower()

            # Ensure required columns exist
            if 'company ticker' not in df.columns:
                df.insert(1, 'company ticker', '')
            if 'date' not in df.columns:
                df.insert(2, 'date', '')  # Initial blank "Date" column
            if 'invest value' not in df.columns:
                df.insert(5, 'invest value', 0)  # New "Invest Value" column
            if 'num days' not in df.columns:
                df.insert(11, 'num days', 0)  # Insert "Num Days" column after "Cur. Val"
            if 'total p&l (%)' not in df.columns:
                df.insert(10, 'total p&l (%)', 0)
            if 'price chg.' not in df.columns:
                df.insert(13, 'price chg.', 0)

            # Populate 'Company Ticker' and calculate 'Invest Value'
            for index, row in df.iterrows():
                company_name = row['instrument']
                ticker_name = get_ticker_name(company_name)  # Fetch ticker
                df.at[index, 'company ticker'] = ticker_name
                df.at[index, 'invest value'] = row['qty.'] * row['avg. cost']

            # Display the table
            display_table(df)

            # Save the uploaded file path
            save_file_path(file_path)
            uploaded_file_path = file_path
            print(f"Uploaded and saved file: {file_path}")
            messagebox.showinfo("Upload", "File uploaded and saved successfully!")

        except Exception as e:
            messagebox.showerror("File Error", f"An error occurred while reading the file: {e}")
# Display data in Treeview table
def display_table(df):
    # Clear existing rows
    for row in treeview.get_children():
        treeview.delete(row)

    # Insert rows with alternating row colors and conditional formatting
    for idx, row in df.iterrows():
        row_tag = "evenrow" if idx % 2 == 0 else "oddrow"
        # Insert the row normally
        item_id = treeview.insert("", "end", values=(row['instrument'], row['company ticker'], row['date'],
                                                     row['qty.'], row['avg. cost'], row['invest value'],
                                                     row['ltp'], row['cur. val'], row['num days'], row['p&l'],row['total p&l (%)'],
                                                     row['net chg.'], row['day chg.'],row['price chg.']),
                                  tags=(row_tag,))

# Allow editing Company Ticker, Date, and Qty. columns
def on_treeview_double_click(event):
    item = treeview.selection()
    if not item:
        return
    column = treeview.identify_column(event.x)
    col_num = int(column[1:]) - 1  # Get column number (0-based index)

    def update_values():
        # Helper function to update the table data after the edit
        values = list(treeview.item(item, "values"))
        treeview.item(item, values=values)

    if col_num == 1 or col_num == 3:  # Allow editing for "Company Ticker" and "Qty."
        entry = ttk.Entry(treeview)
        entry.place(x=event.x, y=event.y)
        current_value = treeview.item(item)["values"][col_num]
        entry.insert(0, current_value)

        def save_edit():
            new_value = entry.get().strip()
            if new_value:
                values = list(treeview.item(item, "values"))
                values[col_num] = new_value
                if col_num == 3:  # Recalculate Invest Value when Qty. is updated
                    try:
                        values[5] = float(values[3]) * float(values[4])
                    except ValueError:
                        values[5] = 0  # Set to 0 if conversion fails
                if col_num == 1:  # Recalculate Cur. Val when Company Ticker is updated
                    try:
                        values[7] = float(values[3]) * float(values[6])
                    except ValueError:
                        values[7] = 0  # Set to 0 if conversion fails
                treeview.item(item, values=values)
                update_values()  # Update table data

            entry.after(100, entry.place_forget)  # Remove the entry field after 100ms delay

        entry.bind("<Return>", lambda event: save_edit())
        entry.bind("<FocusOut>", lambda event: save_edit())

    elif col_num == 2:  # Manual text entry for "Date" column
        entry = ttk.Entry(treeview)
        entry.place(x=event.x, y=event.y)
        current_value = treeview.item(item)["values"][col_num]

        # Set placeholder text if the current value is empty
        if not current_value:
            current_value = "DD/MM/YYYY"
        entry.insert(0, current_value)

        def on_focus_in(event):
            if entry.get() == "DD/MM/YYYY":
                entry.delete(0, tk.END)

        def save_date():
            new_date = entry.get().strip()
            if new_date:
                try:
                    # Validate and store as datetime object
                    formatted_date = datetime.strptime(new_date, "%d/%m/%Y")
                    values = list(treeview.item(item, "values"))
                    values[col_num] = formatted_date.strftime("%d/%m/%Y")
                    treeview.item(item, values=values)
                    update_values()  # Update table data

                    # Calculate Num Days after date is entered
                    num_days = (datetime.now() - formatted_date).days
                    values[8] = num_days
                    treeview.item(item, values=values)
                    update_values()  # Update table data
                except ValueError:
                    messagebox.showerror("Date Format Error", "Please enter the date in DD/MM/YYYY format.")
            entry.after(100, entry.place_forget)  # Remove the entry field after 100ms delay

        entry.bind("<Return>", lambda event: save_date())
        entry.bind("<FocusOut>", lambda event: save_date())
        entry.bind("<FocusIn>", on_focus_in)


# Add sorting functionality
def sort_column(treeview, column_index, ascending=True, numeric=False, is_date=False):
    data = []
    for child in treeview.get_children():
        item = treeview.item(child)
        value = item["values"][column_index]

        # Handle date-specific sorting
        if is_date:
            try:
                value = datetime.strptime(value, "%d/%m/%Y")
            except (ValueError, TypeError):
                value = datetime.min  # Assign the earliest possible date for empty/invalid values
        elif numeric:
            try:
                value = float(value.replace(",", ""))  # Convert numeric strings
            except (ValueError, AttributeError):
                value = float("-inf")  # Assign smallest number for empty/invalid values

        data.append((value, child))

    # Sort data
    data.sort(key=lambda x: x[0], reverse=not ascending)

    # Rearrange rows in Treeview
    for index, (_, child) in enumerate(data):
        treeview.move(child, "", index)
# Update all tickers to `.NS`
def update_tickers_to_ns():
    """
    Updates all 'Company Ticker' entries in the table by appending '.NS'.
    """
    for row_id in treeview.get_children():
        row_data = list(treeview.item(row_id, "values"))
        if row_data[1]:  # Ensure there's a ticker to update
            row_data[1] = f"{row_data[1].split('.')[0]}.NS"
            treeview.item(row_id, values=row_data)


# Add row functionality
def add_row():
    # Create a new window
    add_window = tk.Toplevel()
    add_window.title("Add New Row")

    # Create Entry widgets for each column
    label_instrument = ttk.Label(add_window, text="Instrument :")
    label_instrument.grid(row=0, column=0, padx=10, pady=5)
    entry_instrument = ttk.Entry(add_window)
    entry_instrument.grid(row=0, column=1, padx=10, pady=5)

    label_qty = ttk.Label(add_window, text="Qty. :")
    label_qty.grid(row=1, column=0, padx=10, pady=5)
    entry_qty = ttk.Entry(add_window)
    entry_qty.grid(row=1, column=1, padx=10, pady=5)

    label_avg_cost = ttk.Label(add_window, text="Avg. cost :")
    label_avg_cost.grid(row=2, column=0, padx=10, pady=5)
    entry_avg_cost = ttk.Entry(add_window)
    entry_avg_cost.grid(row=2, column=1, padx=10, pady=5)

    label_date_value = ttk.Label(add_window, text="Date :")
    label_date_value.grid(row=3, column=0, padx=10, pady=5)
    entry_date_value = ttk.Entry(add_window)
    entry_date_value.grid(row=3, column=1, padx=10, pady=5)

    # Function to handle the 'Add' button click
    def add_to_table():
        # Get values from the entries
        instrument = entry_instrument.get()
        qty = entry_qty.get()
        avg_cost = entry_avg_cost.get()
        date_value = entry_date_value.get()

        # Add the values to the main table
        # Assuming `tree` is your main table (Treeview)
        treeview.insert("", "end", values=(instrument,"",date_value,qty, avg_cost, 0,0,0,0,0,0,0,0))

        # Close the add window
        add_window.destroy()

    # Button to add the row to the table
    add_button = ttk.Button(add_window, text="Add", command=add_to_table)
    add_button.grid(row=4, column=0, columnspan=2, pady=10)

# Update headers with '%'
def update_headers():
    treeview.heading("Net Chg.", text="Net Chg. (%)")
    treeview.heading("Day Chg.", text="Day Chg. (%)")


# Display Day Chg. as a percentage

def update_num_days():
    today = datetime.now()
    for child in treeview.get_children():
        values = list(treeview.item(child)["values"])
        try:
            date_value = datetime.strptime(values[columns.index("Date")], "%d/%m/%Y")
            num_days = (today - date_value).days
            values[columns.index("Num Days")] = num_days
        except (ValueError, IndexError, TypeError):
            values[columns.index("Num Days")] = ""  # Clear the "Num Days" value for invalid/empty dates
        treeview.item(child, values=values)
# Bind sorting functionality
# Handle header double-click for sorting
# Handle header double-click for sorting
def on_column_header_double_click(event):
    region = treeview.identify_region(event.x, event.y)
    if region != "heading":  # Ensure it's a header double-click
        return

    column_index = int(treeview.identify_column(event.x)[1:]) - 1
    column_name = columns[column_index]

    def is_column_numeric():
        for child in treeview.get_children():
            value = treeview.item(child)["values"][column_index]
            if value is not None and not isinstance(value, (int, float)):
                try:
                    float(value)
                except ValueError:
                    return False
        return True

    is_numeric = is_column_numeric()
    is_date = column_name == "Date"  # Check if the column is the "Date" column

    # Create sorting popup
    popup = tk.Toplevel(root)
    popup.geometry(f"+{event.x_root}+{event.y_root}")
    popup.overrideredirect(True)

    def sort_ascending():
        sort_column(treeview, column_index, ascending=True, numeric=is_numeric, is_date=is_date)
        popup.destroy()

    def sort_descending():
        sort_column(treeview, column_index, ascending=False, numeric=is_numeric, is_date=is_date)
        popup.destroy()

    ttk.Button(popup, text="Ascending", command=sort_ascending).pack(fill=tk.BOTH)
    ttk.Button(popup, text="Descending", command=sort_descending).pack(fill=tk.BOTH)
# Handle row double-click for editing
# Handle row double-click for editing
# def on_row_double_click(event):
#     region = treeview.identify_region(event.x, event.y)
#     if region == "cell":  # Ensure it's a cell double-click
#         selected_item = treeview.selection()
#         if not selected_item:
#             return
#
#         item_id = selected_item[0]
#         column_index = int(treeview.identify_column(event.x)[1:]) - 1
#         column_name = columns[column_index]
#         editable_columns = ["Company Ticker", "Qty.", "Date"]
#         if column_name not in editable_columns:
#             return  # Skip non-editable columns
#
#         # Create a small entry widget for editing
#         cell_value = treeview.item(item_id, "values")[column_index]
#         cell_value = "" if cell_value is None else cell_value
#
#         x, y, width, height = treeview.bbox(item_id, f"#{column_index + 1}")
#         entry = ttk.Entry(root)
#         entry.place(x=x + treeview.winfo_rootx(), y=y + treeview.winfo_rooty(), width=width, height=height)
#         entry.insert(0, cell_value)
#
#         def save_edit(event):
#             new_value = entry.get()
#             entry.destroy()
#             # Update treeview and recalculate related fields if needed
#             values = list(treeview.item(item_id, "values"))
#             values[column_index] = new_value
#             treeview.item(item_id, values=values)
#
#             # Trigger "Num Days" update if the "Date" column is edited
#             if column_name == "Date":
#                 update_num_days()
#
#         entry.bind("<Return>", save_edit)
#         entry.focus()
root = tk.Tk()
root.title("Stock Price Checker & Share Manager")
root.geometry("1000x600")

# Set initial theme
style = Style(theme="darkly")  # Start with dark theme
# Function to toggle theme
# Function to toggle theme and update table styles
def toggle_theme():
    current_theme = style.theme_use()
    new_theme = "litera" if current_theme == "darkly" else "darkly"
    style.theme_use(new_theme)
    toggle_button.config(text=f"Switch to {'Dark' if new_theme == 'litera' else 'Light'} Theme")

    # Update Treeview styles
    if new_theme == "litera":
        # Light theme styles
        treeview.tag_configure("evenrow", background="white")
        treeview.tag_configure("oddrow", background="lightgray")
        treeview.tag_configure("profit", foreground="green")
        treeview.tag_configure("loss", foreground="red")
    else:
        # Dark theme styles
        treeview.tag_configure("evenrow", background="gray20")
        treeview.tag_configure("oddrow", background="gray30")
        treeview.tag_configure("profit", foreground="lightgreen")
        treeview.tag_configure("loss", foreground="salmon")

    # Refresh the Treeview display to apply the new styles
    refresh_table_view()
# Add toggle theme button


def update_totals():
    total_invest_value = 0
    total_pnl = 0
    total_cur_value = 0

    for row_id in treeview.get_children():
        values = treeview.item(row_id, "values")
        try:
            total_invest_value += float(values[5])  # Invest Value
            total_pnl += float(values[9])          # P&L
            total_cur_value += float(values[7])    # Current Value
        except (ValueError, IndexError):
            continue

    label_invest_value.config(text=f"Total Invest Value: {total_invest_value:,.2f}")
    label_pnl.config(text=f"Total P&L: {total_pnl:,.2f}")
    label_cur_value.config(text=f"Total Current Value: {total_cur_value:,.2f}")

def refresh_table_view():
    for row_id in treeview.get_children():
        values = treeview.item(row_id, "values")
        # Determine row color (even/odd)
        row_index = treeview.index(row_id)
        tag = "evenrow" if row_index % 2 == 0 else "oddrow"

        # Apply P&L conditional formatting
        pnl_value = float(values[9]) if values[9] else 0  # P&L is the 9th column
        if pnl_value > 0:
            treeview.item(row_id, tags=(tag, "profit"))
        elif pnl_value < 0:
            treeview.item(row_id, tags=(tag, "loss"))
        else:
            treeview.item(row_id, tags=(tag,))

# GUI Elements
label_heading = ttk.Label(root, text="Stock Price Checker & Share Manager", font=("Helvetica", 16, "bold"))
label_heading.pack(pady=10)

frame_buttons = ttk.Frame(root)
frame_buttons.pack(pady=10, fill=tk.X)  # Using fill=tk.X to stretch across the window

button_upload = ttk.Button(frame_buttons, text="Upload Share Portfolio", bootstyle="primary-outline", command=upload_file)
button_upload.pack(side=tk.LEFT, padx=5)

button_refresh = ttk.Button(frame_buttons, text="Refresh Prices", bootstyle="info-outline", command=refresh_prices)
button_refresh.pack(side=tk.LEFT, padx=5)

add_row_button = ttk.Button(frame_buttons, text="Add Row", command=add_row)
add_row_button.pack(side=tk.LEFT, padx=5)

button_update_tickers = ttk.Button(frame_buttons, text="Update Tickers to '.NS'", bootstyle="primary-outline", command=None)
button_update_tickers.pack(side=tk.LEFT, padx=5)

# Add toggle theme button
toggle_button = ttk.Button(frame_buttons, text="Contrast Switch", bootstyle="primary-outline", command=toggle_theme)
toggle_button.pack(side=tk.LEFT, padx=5)

save_button = ttk.Button(frame_buttons, text="Save", command=save_to_file)
save_button.pack(side=tk.LEFT, padx=5, pady=5)

# Frame for totals
totals_frame = ttk.Frame(root)
totals_frame.pack(pady=5)

# Separator Line
separator = ttk.Separator(root, orient="horizontal")
separator.pack(fill=tk.X, pady=5)

# Label for Timestamp
timestamp_label = ttk.Label(root, text="Values are as of [Not refreshed yet]", font=("Helvetica", 10, "italic"))
timestamp_label.pack(pady=5)

# Labels for totals
label_invest_value = ttk.Label(totals_frame, text="Total Invest Value: 0", font=("Helvetica", 12, "bold"))
label_invest_value.grid(row=0, column=0, padx=10)

label_pnl = ttk.Label(totals_frame, text="Total P&L: 0", font=("Helvetica", 12, "bold"))
label_pnl.grid(row=0, column=1, padx=10)

label_cur_value = ttk.Label(totals_frame, text="Total Current Value: 0", font=("Helvetica", 12, "bold"))
label_cur_value.grid(row=0, column=2, padx=10)

# Treeview table
treeview_frame = ttk.Frame(root)
treeview_frame.pack(pady=10)

columns = (
    "Instrument", "Company Ticker", "Date", "Qty.", "Avg. Cost", "Invest Value", "LTP", "Cur. Val",
    "Num Days", "P&L", "Total P&L (%)", "Net Chg.", "Day Chg.", "Price Chg."
)
treeview = ttk.Treeview(treeview_frame, columns=columns, show="headings", height=150, bootstyle="dark")

# Define columns
for col in columns:
    treeview.heading(col, text=col)
    treeview.column(col, anchor=tk.CENTER, width=100)

update_headers()  # Update headers for '%' suffix
treeview.pack(fill=tk.BOTH, expand=True)
treeview.heading("Qty.", command=lambda: sort_column(treeview, "Qty.", False))
treeview.heading("Num Days", command=lambda: sort_column(treeview, "Num Days", False))

# Remove any previous bindings for sorting
treeview.unbind("<Button-1>")
treeview.unbind("<Double-1>")
treeview.bind("<Double-2>", show_historical_data)
# Assuming your Treeview widget is already set up
treeview.bind("<Double-1>", on_treeview_double_click)
treeview.bind("<Double-1>", lambda e: on_treeview_double_click(e) if treeview.identify_region(e.x, e.y) == "cell" else on_column_header_double_click(e))

# Define additional row styles for P&L conditional formatting
treeview.tag_configure("profit", foreground="green")
treeview.tag_configure("loss", foreground="red")

# Define alternating row colors
treeview.tag_configure("evenrow", background="gray20")
treeview.tag_configure("oddrow", background="gray30")

initialize_app()
# Start the Tkinter main loop
root.mainloop()