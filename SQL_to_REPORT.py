#!/usr/bin/env python
# coding: utf-8

# In[7]:


import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as messagebox
from tkcalendar import DateEntry
import pandas as pd
from sqlalchemy import create_engine, text
from openpyxl import Workbook
from openpyxl.styles import Font
import os
from datetime import datetime
import sys
from PIL import Image, ImageTk
import openpyxl.drawing.image

# Database connection information
DATABASES = {
    "MIS_DB": {
        "name": "KINLEY_MIS_DB",
        "server": r"DESKTOP-87HT9VP\WINCC",
        "driver": "ODBC Driver 17 for SQL Server",
        "username": "admin",
        "password": "admin"
    },
    "RO_DATA": {
        "name": "KINLEY_RO_DB",
        "server": r"DESKTOP-87HT9VP\WINCC",
        "driver": "ODBC Driver 17 for SQL Server",
        "username": "admin",
        "password": "admin"
    }
}

if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app 
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Database Query to Excel")

        # Add logo at the top
        try:
            # Load and resize the logo
            logo_path = os.path.join(application_path, "logo.png")
            logo_image = Image.open(logo_path)
            # Resize the image if needed (adjust size as needed)
            logo_image = logo_image.resize((200, 100), Image.Resampling.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_image)
            
            # Create frame for logo and website
            logo_frame = tk.Frame(root)
            logo_frame.grid(row=0, columnspan=4, pady=5)
            
            # Create label for logo
            logo_label = tk.Label(logo_frame, image=logo_photo)
            logo_label.image = logo_photo  # Keep a reference!
            logo_label.pack()
            
            # Add website text below logo
            website_label = tk.Label(logo_frame, text="pisplindustry.com", font=('Arial', 10))
            website_label.pack()

            # Rest of the GUI elements shifted down by one row
            self.label_start_date = tk.Label(root, text="Start Date")
            self.label_start_date.grid(row=1, column=0, padx=10, pady=10)
            self.start_date = DateEntry(root, date_pattern='yyyy-mm-dd')
            self.start_date.grid(row=1, column=1, padx=10, pady=10)
            
            self.label_start_time = tk.Label(root, text="Start Time (HH:MM:SS)")
            self.label_start_time.grid(row=1, column=2, padx=10, pady=10)
            self.start_time = ttk.Combobox(root, values=[f'{str(h).zfill(2)}:00:00' for h in range(24)])
            self.start_time.grid(row=1, column=3, padx=10, pady=10)
            self.start_time.set("00:00:00")
            
            # End Date and Time
            self.label_end_date = tk.Label(root, text="End Date")
            self.label_end_date.grid(row=2, column=0, padx=10, pady=10)
            self.end_date = DateEntry(root, date_pattern='yyyy-mm-dd')
            self.end_date.grid(row=2, column=1, padx=10, pady=10)

            self.label_end_time = tk.Label(root, text="End Time (HH:MM:SS)")
            self.label_end_time.grid(row=2, column=2, padx=10, pady=10)
            self.end_time = ttk.Combobox(root, values=[f'{str(h).zfill(2)}:00:00' for h in range(24)])
            self.end_time.grid(row=2, column=3, padx=10, pady=10)
            self.end_time.set("23:59:59")
            
            # Database Dropdown
            self.label_database = tk.Label(root, text="Select Database")
            self.label_database.grid(row=3, column=0, padx=10, pady=10)
            self.database_var = tk.StringVar(root)
            self.database_dropdown = ttk.Combobox(root, textvariable=self.database_var, values=list(DATABASES.keys()))
            self.database_dropdown.grid(row=3, column=1, padx=10, pady=10)
            self.database_dropdown.set(list(DATABASES.keys())[0])

            # Table Dropdown
            self.label_table = tk.Label(root, text="Select Table")
            self.label_table.grid(row=4, column=0, padx=10, pady=10)
            self.table_var = tk.StringVar(root)
            self.table_dropdown = ttk.Combobox(root, textvariable=self.table_var)
            self.table_dropdown.grid(row=4, column=1, padx=10, pady=10)

            # Populate tables dropdown
            self.populate_tables()

            # Generate Button
            self.generate_button = tk.Button(root, text="Generate Excel", command=self.generate_excel)
            self.generate_button.grid(row=5, columnspan=4, pady=20)
            
            # Footer Label
            self.footer_label = tk.Label(root, text="Developed by Subhajit Duttagupta")
            self.footer_label.grid(row=6, columnspan=4, pady=10)

            # Bind database selection to update tables
            self.database_dropdown.bind('<<ComboboxSelected>>', self.on_database_change)

        except Exception as e:
            messagebox.showerror("Error", f"Could not load logo: {e}")
            # Continue with the rest of the GUI setup without the logo

    def populate_tables(self):
        try:
            selected_db = DATABASES[self.database_var.get()]
            connection_string = f"mssql+pyodbc://{selected_db['username']}:{selected_db['password']}@{selected_db['server']}/{selected_db['name']}?driver={selected_db['driver']}"
            engine = create_engine(connection_string)
            query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
            with engine.connect() as connection:
                result = connection.execute(text(query))
                tables = [row[0] for row in result]
                self.table_dropdown['values'] = tables
                if tables:
                    self.table_dropdown.set(tables[0])  # Select the first table by default
        except Exception as e:
            messagebox.showerror("Error", f"Could not fetch tables: {e}")

    def generate_excel(self):
        start_datetime = f"{self.start_date.get()} {self.start_time.get()}"
        end_datetime = f"{self.end_date.get()} {self.end_time.get()}"
        selected_table = self.table_var.get()
        selected_db = DATABASES[self.database_var.get()]

        if not start_datetime or not end_datetime or not selected_table:
            messagebox.showerror("Error", "All fields are required!")
            return

        try:
            # Connect to the database and fetch the data
            connection_string = f"mssql+pyodbc://{selected_db['username']}:{selected_db['password']}@{selected_db['server']}/{selected_db['name']}?driver={selected_db['driver']}"
            engine = create_engine(connection_string)
            
            # Convert date to DD-MM-YYYY format for the LIKE query
            date_for_query = self.start_date.get().split('-')  # splits YYYY-MM-DD
            formatted_date = f"{date_for_query[2]}-{date_for_query[1]}-{date_for_query[0]}"  # creates DD-MM-YYYY
            
            # Query that will be executed
            query = f"""
                SELECT TOP 1000 * 
                FROM [{selected_db['name']}].[dbo].[{selected_table}] 
                WHERE [DateAndTime] LIKE '{formatted_date}%'
                ORDER BY [DateAndTime] DESC
            """

            # Show query for debugging
            messagebox.showinfo("Debug - Query", f"""
            Database: {selected_db['name']}
            Table: {selected_table}
            Using date format: {formatted_date}
            Full query:
            {query}
            
            Example of expected format: 22-02-2025 17:17:22
            """)
            
            df = pd.read_sql(query, engine)
            
            # Remove specified columns based on table
            if selected_table == 'ro_tab':
                columns_to_keep = [
                    'DateAndTime',
                    'Permeate_PH',
                    'Permeate_COND',
                    'Feed_PH',
                    'Feed_COND',
                    'Feed_FLOW',
                    'Reject_FLOW'
                ]
                df = df[columns_to_keep]
            elif selected_table == 'mis_tab':
                columns_to_keep = [
                    'DateAndTime',
                    'Conductivity_1',
                    'Conductivity_2',
                    'Mis_Flow_Rate',
                    'Ozone'
                ]
                df = df[columns_to_keep]
            
            # Show data retrieval info
            debug_info = f"""
            Database: {selected_db['name']}
            Table: {selected_table}
            Retrieved Rows: {len(df)}
            Columns: {', '.join(df.columns.tolist())}
            
            First few rows:
            {df.head().to_string() if not df.empty else 'No data found'}
            """
            messagebox.showinfo("Debug - Data Retrieved", debug_info)

            # Create a new workbook
            wb = Workbook()
            ws = wb.active

            # Set database name as bold heading in center
            ws['A1'] = selected_db['name']
            ws['A1'].font = Font(bold=True, size=14)
            ws.merge_cells('A1:D1')
            ws['A1'].alignment = openpyxl.styles.Alignment(horizontal='center')

            # Set table name as heading in center
            ws['A2'] = selected_table
            ws['A2'].font = Font(bold=True, size=12)
            ws.merge_cells('A2:D2')
            ws['A2'].alignment = openpyxl.styles.Alignment(horizontal='center')

            # Set datetime range in center
            ws['A3'] = 'Start DateTime:'
            ws['B3'] = start_datetime
            ws['C3'] = 'End DateTime:'
            ws['D3'] = end_datetime
            for cell in ['A3', 'B3', 'C3', 'D3']:
                ws[cell].alignment = openpyxl.styles.Alignment(horizontal='center')

            # Write column headers in row 4
            for idx, col in enumerate(df.columns, 1):
                cell = ws.cell(row=4, column=idx, value=col)
                cell.font = Font(bold=True)
                cell.alignment = openpyxl.styles.Alignment(horizontal='center')

            # Write data starting from row 5
            for r_idx, row in enumerate(df.itertuples(index=False), 5):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center')

            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column = list(column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = adjusted_width

            # Generate filename with current datetime and table name
            current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"{selected_table}_{current_datetime}.xlsx"
            
            # Save the workbook
            wb.save(output_file)

            # Show success message with file path
            full_path = os.path.abspath(output_file)
            messagebox.showinfo("Success", f"Excel file generated successfully!\nLocation: {full_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generating Excel: {str(e)}\nQuery: {query}")

    def on_database_change(self, event=None):
        self.populate_tables()

root = tk.Tk()
app = App(root)
root.mainloop()

