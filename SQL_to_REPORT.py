#!/usr/bin/env python
# coding: utf-8

# In[7]:


import tkinter as tk
from tkinter import ttk, filedialog
import tkinter.messagebox as messagebox
from tkcalendar import DateEntry
import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL
import pyodbc
from openpyxl import Workbook
from openpyxl.styles import Font
import os
from datetime import datetime
import sys
from PIL import Image, ImageTk
import openpyxl.drawing.image
import pyodbc  # For direct ODBC connection testing

# Database connection information
DATABASES = {
    "ERGO": {
        "name": "ERGO",
        "server": r"DESKTOP-E636DTV\WINCC",  # Using the instance name since it works in sqlcmd
        "driver": "ODBC Driver 17 for SQL Server",
        "username": "admin",
        "password": "Moinadanga@1"
    }
    # "ALL_DATA": {
    #     "name": "KINLEY_RO_DB",
    #     "server": r"DESKTOP-87HT9VP\WINCC",
    #     "driver": "ODBC Driver 17 for SQL Server",
    #     "username": "admin",
    #     "password": "admin"
    # },

    # "KINLEY_MIS_DB": {
    #     "name": "KINLEY_MIS_DB",
    #     "server": r"DESKTOP-87HT9VP\WINCC",
    #     "driver": "ODBC Driver 17 for SQL Server",
    #     "username": "admin",
    #     "password": "admin"
    # }
}

# Table mapping for each database
DATABASE_TABLES = {
    "ERGO": ["REJECT_RO", "RO_TABLE"]
    # "ALL_DATA": ["all_data"],
    # "KINLEY_MIS_DB": ["mis_tab"]
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
        self.uploaded_image_path = None  # Store the path of uploaded image

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

            # Image Upload Section
            self.label_image = tk.Label(root, text="Upload Image (Optional)")
            self.label_image.grid(row=5, column=0, padx=10, pady=10)
            
            # Frame for image upload controls
            image_frame = tk.Frame(root)
            image_frame.grid(row=5, column=1, columnspan=2, padx=10, pady=10, sticky='w')
            
            self.upload_button = tk.Button(image_frame, text="Browse Image", command=self.upload_image)
            self.upload_button.pack(side=tk.LEFT, padx=(0, 10))
            
            self.image_status_label = tk.Label(image_frame, text="No image selected", fg="gray")
            self.image_status_label.pack(side=tk.LEFT)

            # Populate tables dropdown
            self.populate_tables()

            # Generate Button
            self.generate_button = tk.Button(root, text="Generate Excel", command=self.generate_excel)
            self.generate_button.grid(row=6, columnspan=4, pady=20)
            
            # Footer Label
            self.footer_label = tk.Label(root, text="Developed by Subhajit Duttagupta")
            self.footer_label.grid(row=7, columnspan=4, pady=10)

            # Bind database selection to update tables
            self.database_dropdown.bind('<<ComboboxSelected>>', self.on_database_change)

            # Set the report directory to a specific path
            self.report_dir = r"C:\Users\HP\Desktop\REPORT"
            if not os.path.exists(self.report_dir):
                os.makedirs(self.report_dir)

        except Exception as e:
            messagebox.showerror("Error", f"Could not load logo: {e}")
            # Continue with the rest of the GUI setup without the logo

    def upload_image(self):
        """Handle image upload functionality"""
        file_types = [
            ('Image files', '*.png *.jpg *.jpeg *.gif *.bmp *.tiff'),
            ('PNG files', '*.png'),
            ('JPEG files', '*.jpg *.jpeg'),
            ('All files', '*.*')
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=file_types
        )
        
        if filename:
            try:
                # Validate that it's a valid image file
                with Image.open(filename) as img:
                    # Store the image path
                    self.uploaded_image_path = filename
                    
                    # Update the status label
                    image_name = os.path.basename(filename)
                    if len(image_name) > 30:
                        image_name = image_name[:27] + "..."
                    self.image_status_label.config(text=f"Selected: {image_name}", fg="green")
                    
            except Exception as e:
                messagebox.showerror("Error", f"Invalid image file: {str(e)}")
                self.uploaded_image_path = None
                self.image_status_label.config(text="No image selected", fg="gray")

    def populate_tables(self):
        try:
            selected_db_name = self.database_var.get()
            
            # Use predefined table mapping instead of querying database
            if selected_db_name in DATABASE_TABLES:
                tables = DATABASE_TABLES[selected_db_name]
                self.table_dropdown['values'] = tables
                if tables:
                    self.table_dropdown.set(tables[0])  # Select the first table by default
            else:
                # Fallback to database query if not in predefined mapping
                selected_db = DATABASES[selected_db_name]
                connection_string = (
                    f"mssql+pyodbc://"
                    f"{selected_db['username']}:{selected_db['password']}@"
                    f"{selected_db['server']}"
                    f"?driver={selected_db['driver']}"
                    f"&database={selected_db['name']}"
                    f"&TrustServerCertificate=yes"
                    f"&encrypt=no"
                )
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

    def add_image_to_excel(self, ws, image_path):
        """Add image to Excel worksheet at cell O2"""
        try:
            # Method 1: Try direct image insertion
            try:
                # Create Excel image object directly from the original file
                excel_image = openpyxl.drawing.image.Image(image_path)
                
                # Resize the image in Excel
                excel_image.width = 80
                excel_image.height = 80
                
                # Add image to cell O2
                ws.add_image(excel_image, 'O2')
                return  # Success, exit the function
                
            except Exception as direct_error:
                print(f"Direct image insertion failed: {direct_error}")
                
            # Method 2: Fallback to temporary file method
            with Image.open(image_path) as img:
                # Calculate new size to fit in Excel cell
                max_size = (100, 100)
                img.thumbnail(max_size, Image.Resampling.LANCZOS)
                
                # Create a unique temporary file name
                import tempfile
                import uuid
                temp_filename = f"temp_image_{uuid.uuid4().hex[:8]}.png"
                temp_image_path = os.path.join(self.report_dir, temp_filename)
                
                # Ensure the report directory exists
                os.makedirs(self.report_dir, exist_ok=True)
                
                # Save resized image to temporary file
                img.save(temp_image_path, "PNG")
                
                # Verify the temporary file was created
                if not os.path.exists(temp_image_path):
                    raise Exception("Temporary image file was not created successfully")
                
                # Create Excel image object
                excel_image = openpyxl.drawing.image.Image(temp_image_path)
                excel_image.width = 80
                excel_image.height = 80
                
                # Add image to cell O2
                ws.add_image(excel_image, 'O2')
                
                # Clean up temporary file
                try:
                    if os.path.exists(temp_image_path):
                        os.remove(temp_image_path)
                except Exception as cleanup_error:
                    print(f"Warning: Could not delete temporary file: {cleanup_error}")
                    
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not add image to Excel: {str(e)}")

    def generate_excel(self):
        start_datetime = f"{self.start_date.get()} {self.start_time.get()}"
        end_datetime = f"{self.end_date.get()} {self.end_time.get()}"
        selected_table = self.table_var.get()
        selected_db = DATABASES[self.database_var.get()]

        if not start_datetime or not end_datetime or not selected_table:
            messagebox.showerror("Error", "All fields are required!")
            return

        try:
            # Show connection parameters for debugging
            # Try direct ODBC connection first to test credentials
            try:
                conn_str = (
                    f"DRIVER={{{selected_db['driver']}}};"
                    f"SERVER={selected_db['server']};"
                    f"DATABASE={selected_db['name']};"
                    f"UID={selected_db['username']};"
                    f"PWD={selected_db['password']};"
                    "TrustServerCertificate=yes;"
                    "Encrypt=no"
                )
                messagebox.showinfo("Debug - ODBC Connection String", f"Testing ODBC connection with:\n{conn_str}")
                
                # Test direct ODBC connection
                conn = pyodbc.connect(conn_str, timeout=30)
                conn.close()
                messagebox.showinfo("Success", "Direct ODBC connection test successful!")
            except Exception as odbc_error:
                messagebox.showerror("ODBC Connection Error", f"Direct ODBC connection failed: {str(odbc_error)}")
                raise  # Re-raise to prevent further execution
            
            # Connect to the database and fetch the data
            connection_url = URL.create(
                "mssql+pyodbc",
                username=selected_db['username'],
                password=selected_db['password'],
                host=selected_db['server'],
                database=selected_db['name'],
                query={
                    "driver": selected_db['driver'],
                    "TrustServerCertificate": "yes",
                    "encrypt": "no"
                }
            )
            
            # Show the exact connection string for debugging
            messagebox.showinfo("Debug - Connection URL", f"Connection URL: {connection_url}")
            
            engine = create_engine(connection_url)
            
            # Convert datetime strings to DD-MM-YYYY HH:MM:SS format (Style 103)
            start_dt = datetime.strptime(start_datetime, '%Y-%m-%d %H:%M:%S')
            end_dt = datetime.strptime(end_datetime, '%Y-%m-%d %H:%M:%S')
            
            start_datetime_sql = start_dt.strftime('%d-%m-%Y %H:%M:%S')
            end_datetime_sql = end_dt.strftime('%d-%m-%Y %H:%M:%S')
            
            # Create the query using SQLAlchemy's text() with parameters
            query = text(f"""
                SELECT *
                FROM [{selected_db['name']}].dbo.[{selected_table}]
                WHERE TRY_CONVERT(DATETIME, [DateAndTime], 103) >= TRY_CONVERT(DATETIME, :start_date, 103)
                AND TRY_CONVERT(DATETIME, [DateAndTime], 103) <= TRY_CONVERT(DATETIME, :end_date, 103)
                ORDER BY TRY_CONVERT(DATETIME, [DateAndTime], 103) DESC
            """).bindparams(
                start_date=start_datetime_sql,
                end_date=end_datetime_sql
            )
            
            # Show the complete query with parameters for debugging
            debug_query = f"""
            Query being executed:
            SELECT *
            FROM [{selected_db['name']}].dbo.[{selected_table}]
            WHERE TRY_CONVERT(DATETIME, [DateAndTime], 103) >= TRY_CONVERT(DATETIME, '{start_datetime_sql}', 103)
            AND TRY_CONVERT(DATETIME, [DateAndTime], 103) <= TRY_CONVERT(DATETIME, '{end_datetime_sql}', 103)
            ORDER BY TRY_CONVERT(DATETIME, [DateAndTime], 103) DESC
            """
            messagebox.showinfo("Debug - Full Query", debug_query)

            # Show query for debugging
            messagebox.showinfo("Debug - Query", f"""
            Database: {selected_db['name']}
            Table: {selected_table}
            Start DateTime: {start_datetime_sql}
            End DateTime: {end_datetime_sql}
            Using parameterized query for security
            """)
            
            with engine.connect() as connection:
                result = connection.execute(query)
                df = pd.DataFrame(result.fetchall(), columns=result.keys())
            
            # Show all columns in debug before filtering
            messagebox.showinfo("Debug - Available Columns", f"Available columns in {selected_table}:\n{', '.join(df.columns)}")
            
            # Remove specified columns based on table
            if selected_table == 'RO_TABLE':
                columns_to_keep = [
                    'DateAndTime',
                    'Permeate_PH',
                    'Permeate_COND',
                    'Feed_PH',
                    'Feed_COND',
                    'Feed_FLOW',
                    'Permeate_FLOW',
                    'Reject_FLOW',
                    'STAGE_2_FEED_FLOW',
                    'ACF_OUTLET_ORP',
                    'FEED_ORP',
                    'HEADER_PRESSURE',
                    'HPP_1_PID',
                    'HPP_2_PID',
                    'MICRON_1_INLET_PR',
                    'MICRON_1_OUTLET_PR',
                    'MICRON_2_INLET_PR',
                    'MICRON_2_OUTLET_PR',
                    'PERMEATE_TNS_1',
                    'PERMEATE_TNS_2',
                    'PERMEATE_TNS_3',
                    'PH_CORRECTION_DOSING',
                    'POST_MICRON_FILTER_1_INLET',
                    'POST_MICRON_FILTER_2_INLET',
                    'POST_MICRON_OUTLET_1',
                    'POST_MICRON_OUTLET_2',
                    'PST_1_LVL',
                    'PST_2_LVL',
                    'UV_HEALTHY'
                ]
            elif selected_table == 'REJECT_RO':
                columns_to_keep = [
                    'DateAndTime',
                    'Feed_Flow',
                    'Per_Flow',
                    'Reject_Flow',
                    'Feed_PH',
                    'Per_PH',
                    'Per_COND',
                    'Turbidity',
                    'CO2_PID',
                    'DOSING_PUMP_PID',
                    'HPP_PID'
                ]
            
            # Keep all columns if they exist in the DataFrame
            available_columns = [col for col in columns_to_keep if col in df.columns]
            if available_columns:
                df = df[available_columns]
            
            # Show which columns were kept
            messagebox.showinfo("Debug - Kept Columns", f"Keeping these columns:\n{', '.join(available_columns)}")
            
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

            # Load template instead of creating new workbook
            template_path = os.path.join(application_path, "Report_Template.xlsx")
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active

            # Set database name as bold heading
            ws['A6'] = selected_db['name']
            ws['A6'].font = Font(bold=True, size=14)
            ws.merge_cells('A6:D6')
            ws['A6'].alignment = openpyxl.styles.Alignment(horizontal='center')

            # Set table name as heading
            ws['A7'] = selected_table
            ws['A7'].font = Font(bold=True, size=12)
            ws.merge_cells('A7:D7')
            ws['A7'].alignment = openpyxl.styles.Alignment(horizontal='center')

            # Set datetime range
            ws['A8'] = 'Start DateTime:'
            ws['B8'] = start_datetime
            ws['C8'] = 'End DateTime:'
            ws['D8'] = end_datetime
            for cell in ['A8', 'B8', 'C8', 'D8']:
                ws[cell].alignment = openpyxl.styles.Alignment(horizontal='center')

            # Write column headers in row 9
            for idx, col in enumerate(df.columns, 1):
                cell = ws.cell(row=9, column=idx, value=col)
                cell.font = Font(bold=True)
                cell.alignment = openpyxl.styles.Alignment(horizontal='center')

            # Write data starting from row 10
            for r_idx, row in enumerate(df.itertuples(index=False), 10):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center')

            # Add uploaded image to cell O2 if an image was selected
            if self.uploaded_image_path:
                self.add_image_to_excel(ws, self.uploaded_image_path)

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

            # Save in Report folder
            current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(self.report_dir, f"{selected_table}_{current_datetime}.xlsx")
            
            # Save the workbook
            wb.save(output_file)

            # Show success message with file path
            success_message = f"Excel file generated successfully!\nLocation: {output_file}"
            if self.uploaded_image_path:
                success_message += f"\nImage added to cell O2"
            messagebox.showinfo("Success", success_message)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generating Excel: {str(e)}")

    def on_database_change(self, event=None):
        self.populate_tables()

root = tk.Tk()
app = App(root)
root.mainloop()