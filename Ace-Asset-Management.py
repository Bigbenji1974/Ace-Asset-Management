import sys
import os
import sqlite3
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QStackedWidget, QSpacerItem, QSizePolicy, QLineEdit, QTableWidget, QTableWidgetItem, QDialog, QComboBox, QGridLayout
from PyQt6.QtGui import QPixmap, QFont
from PyQt6.QtCore import Qt, QPropertyAnimation, QEvent
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
from datetime import datetime, timedelta 
import pandas as pd
import threading
import webbrowser


class WebStyleApp(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Full-Screen Web-Style GUI")

        # Enable window decorations but keep full-screen mode
        self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.CustomizeWindowHint)
        self.showFullScreen()  # Open in full-screen mode

        # Ensure the database is set up
        self.database_path = self.setup_database()

        # Initialize Stacked Widget for pages
        self.pages = QStackedWidget()
        self.pages.addWidget(self.create_home_page())  # Index 0
        self.pages.addWidget(self.create_ace_equipment_program_page())  # Index 1
        self.pages.addWidget(self.create_database_page())  # Index 2
        self.pages.addWidget(self.create_rma_page())  # Index 3
        self.pages.addWidget(self.create_print_page())  # Index 4

        # **Main Layout**
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # **Top Bar for Minimize & Close Buttons**
        top_bar_widget = QWidget()
        top_bar_widget.setFixedHeight(30)  # Keeps it small
        top_bar_widget.setStyleSheet("background-color: rgba(34, 34, 34, 0.8);")  # Dark semi-transparent background

        top_bar_layout = QHBoxLayout(top_bar_widget)
        top_bar_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins
        top_bar_layout.setSpacing(0)  # No spacing between elements

        # **Minimize & Close Buttons**
        minimize_btn = QPushButton("—")  # Minimize button
        minimize_btn.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        minimize_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        minimize_btn.setFixedSize(30, 25)
        minimize_btn.setStyleSheet("""
            QPushButton {
                color: white;
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.2);
            }
        """)
        minimize_btn.clicked.connect(self.showMinimized)

        close_btn = QPushButton("✕")  # Close button
        close_btn.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setFixedSize(30, 25)
        close_btn.setStyleSheet("""
            QPushButton {
                color: white;
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.2);
            }
        """)
        close_btn.clicked.connect(self.close)

        # Add stretch before buttons to align them to the right
        top_bar_layout.addStretch()
        top_bar_layout.addWidget(minimize_btn)
        top_bar_layout.addWidget(close_btn)

        # **Navigation Bar (Full Button Coverage)**
        nav_bar_widget = QWidget()
        nav_bar_widget.setStyleSheet("""
            background-color: #B22222;  /* Deep red background */
            border: 2px solid #800000;  /* Dark red border */
            padding: 0;
        """)
        nav_bar_widget.setFixedHeight(70)  # Keep navigation bar height

        nav_bar = QHBoxLayout(nav_bar_widget)
        nav_bar.setContentsMargins(0, 0, 0, 0)  # Remove margins
        nav_bar.setSpacing(0)  # Remove spacing between buttons

        # **Navigation Buttons (No Gaps, Full Width, Darker Hover Effect)**
        buttons = {
            "Home Page": 0,
            "Ace Equipment Program": 1,
            "Database": 2,
            "RMA": 3,
            "Print": 4
        }

        self.nav_buttons = []

        for text, index in buttons.items():
            btn = QPushButton(text)
            btn.setFont(QFont("Arial", 18, QFont.Weight.Bold))
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)  # Ensures equal width & height

            # **Updated Hover Effect with a Darker Shade**
            btn.setStyleSheet("""
                QPushButton {
                    color: white;
                    background-color: #B22222;  /* Default deep red */
                    border: 2px solid #800000;  /* Dark red border */
                    padding: 15px;
                }
                QPushButton:hover {
                    background-color: #CC2E1A;  /* Darker red on hover */
                }
                QPushButton:pressed {
                    background-color: #800000;  /* Darker red on press */
                }
            """)

            btn.clicked.connect(lambda checked, i=index, t=text: self.handle_tab_switch(i, t))
            nav_bar.addWidget(btn)
            self.nav_buttons.append(btn)

            # **Add Widgets to Layout**
            layout.addWidget(top_bar_widget)  # Minimize & Close bar
            layout.addWidget(nav_bar_widget)  # Navigation bar
            layout.addWidget(self.pages)  # Main content area

            # Set the final layout
            self.setLayout(layout)


    def handle_tab_switch(self, index, tab_name):
        """Handles tab switching and resets the input fields when entering 'Ace Equipment Program'."""
        self.pages.setCurrentIndex(index)  # Switch tab
        if tab_name == "Ace Equipment Program":
            self.reset_inputs()  # Reset all fields
            self.employee_id_input.setFocus()  # 🚀 Move cursor to Employee ID input

    def create_database_page(self):
        """Creates the Database Page with query options inside a structured card."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Card Widget ---
        card_widget = QWidget()
        card_layout = QVBoxLayout(card_widget)
        card_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.setContentsMargins(40, 40, 40, 40)

        card_widget.setStyleSheet("""
            background-color: white;
            border-radius: 10px;
            padding: 40px;
        """)

        # --- Title ---
        title_label = QLabel("Database Queries")
        title_label.setFont(QFont("Arial", 32, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("margin-bottom: 20px;")  # Adds spacing below the title

        # --- Query Dropdown ---
        self.query_type_combo = QComboBox()
        self.query_type_combo.addItems([
            "Entire Database",
            "Query Today",
            "Query Non-Turned In Equipment",
            "Query Last Month",
            "Query Mobile Printers",
            "Query Radios",
            "Query Filler Sets",
            "Query RF Guns"
        ])
        self.query_type_combo.setFont(QFont("Arial", 18))
        self.query_type_combo.setFixedWidth(800)  # Ensures dropdown has a proper width

        # 🚨 FIX: Force dropdown content to **take full width**
        self.query_type_combo.setStyleSheet("""
            QComboBox {
                background-color: #F0F0F0;
                border: 2px solid #B22222;
                border-radius: 8px;
                padding: 10px;
                font-size: 18px;
            }
            QComboBox::drop-down {
                border: none;
                width: 0px;  /* Hide dropdown button border */
            }
            QComboBox QAbstractItemView {
                background-color: white;
                border: 1px solid #B22222;
                selection-background-color: #FF5733;
                padding: 5px;
            }
            QComboBox QAbstractItemView::item {
                min-width: 800px;  /* 🚀 Force items to take full width */
                padding: 10px;
            }
            QComboBox::item:selected {
                background-color: #FF5733;
                color: white;
            }
        """)

        # --- Execute Button ---
        execute_button = QPushButton("Execute Query")
        execute_button.setFont(QFont("Arial", 20))
        execute_button.setFixedWidth(400)  # Increased width for alignment
        execute_button.setCursor(Qt.CursorShape.PointingHandCursor)
        execute_button.setStyleSheet("""
            QPushButton {
                background-color: #B22222;
                color: white;
                border-radius: 8px;
                padding: 12px;
            }
            QPushButton:hover {
                background-color: #FF5733;
            }
        """)
        execute_button.clicked.connect(self.execute_query)

        # --- Layout for Dropdown and Button ---
        column_layout = QVBoxLayout()
        column_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        column_layout.setSpacing(50)
        column_layout.addWidget(self.query_type_combo, alignment=Qt.AlignmentFlag.AlignCenter)
        column_layout.addWidget(execute_button, alignment=Qt.AlignmentFlag.AlignCenter)

        # --- Assemble Components into Card ---
        card_layout.addWidget(title_label)
        card_layout.addLayout(column_layout)

        layout.addWidget(card_widget)

        return page


    def setup_database(self):
        """Ensures the SQLite database exists and sets up the necessary tables if they do not exist."""
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "Ace Database Log")
        if not os.path.exists(desktop_path):
            os.makedirs(desktop_path)  # Create the folder if it doesn't exist

        database_path = os.path.join(desktop_path, "equipment_log.db")

        # Connect to SQLite and create tables if they do not exist
        conn = sqlite3.connect(database_path)
        cursor = conn.cursor()

        # Create equipment log table if not exists
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS equipment_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                employee_id TEXT,
                equipment_number TEXT
            )
        """)

        # 🚨 Ensure problem_equipment table exists 🚨
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS problem_equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipment_number TEXT NOT NULL,
                description TEXT NOT NULL,
                timestamp TEXT NOT NULL
            )
        """)

        conn.commit()
        conn.close()

        return database_path

    def create_home_page(self):
        """Creates the Home Page with a Header and Logo inside a styled card."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Background Styling
        page.setStyleSheet("background-color: #f5f5f5;")  # Light gray background

        # --- Header ---
        header_label = QLabel("Welcome to the Ace Hardware Equipment Management Program")
        header_label.setFont(QFont("Arial", 32, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Load Logo ---
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "Program Photos", "Ace_Logo.png")
        one_drive_path = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "Program Photos", "Ace_Logo.png")
        image_path = desktop_path if os.path.exists(desktop_path) else one_drive_path if os.path.exists(one_drive_path) else None

        img_label = QLabel()
        if image_path:
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                img_label.setPixmap(pixmap.scaled(500, 300, Qt.AspectRatioMode.KeepAspectRatio))
                img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            else:
                img_label.setText("Error: Image file cannot be loaded.")
                img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            img_label.setText("Error: Image not found in expected directories.")
            img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Card Widget ---
        card_widget = QWidget()
        card_layout = QVBoxLayout(card_widget)
        card_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.setContentsMargins(20, 20, 20, 20)

        card_widget.setStyleSheet("""
            background-color: white;
            border-radius: 10px;
            padding: 30px;
        """)

        card_layout.addWidget(header_label)
        card_layout.addWidget(img_label)

        layout.addWidget(card_widget)

        return page

    def create_ace_equipment_program_page(self):
        """Creates the Ace Equipment Program Page with structured styling inside a modern card."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Card Widget ---
        card_widget = QWidget()
        card_layout = QVBoxLayout(card_widget)
        card_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.setContentsMargins(30, 30, 30, 30)
        card_layout.setSpacing(20)  # Increased spacing between elements

        card_widget.setStyleSheet("""
            background-color: white;
            border-radius: 15px;
            padding: 40px;
        """)

        # --- Header ---
        header_label = QLabel("Sign Equipment In or Out")
        header_label.setFont(QFont("Arial", 32, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(header_label)

        # --- Input Field Container ---
        input_container = QWidget()
        input_layout = QHBoxLayout(input_container)
        input_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(100)

        # --- Employee ID Input ---
        self.employee_id_input = QLineEdit()
        self.employee_id_input.setPlaceholderText("Scan Employee ID (6 Digits) or Equipment Number")
        self.employee_id_input.setFont(QFont("Arial", 18))
        self.employee_id_input.setFixedWidth(400)
        self.employee_id_input.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.employee_id_input.setStyleSheet("""
            QLineEdit {
                background-color: #F0F0F0;
                border: 2px solid #B22222;
                border-radius: 8px;
                padding: 10px;
                font-size: 18px;
            }
            QLineEdit:focus {
                border: 2px solid #FF5733;
            }
        """)

        self.employee_id_input.returnPressed.connect(self.process_equipment_entry)

        # --- Equipment Number Input ---
        self.equipment_input = QLineEdit()
        self.equipment_input.setPlaceholderText("Enter Equipment Number (e.g., TM123, RF456)")
        self.equipment_input.setFont(QFont("Arial", 18))
        self.equipment_input.setFixedWidth(400)
        self.equipment_input.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.equipment_input.setVisible(False)
        self.equipment_input.setStyleSheet("""
            QLineEdit {
                background-color: #F0F0F0;
                border: 2px solid #B22222;
                border-radius: 8px;
                padding: 10px;
                font-size: 18px;
            }
            QLineEdit:focus {
                border: 2px solid #FF5733;
            }
        """)

        self.equipment_input.returnPressed.connect(self.process_equipment_entry)

        # Add input fields to the layout
        input_layout.addWidget(self.employee_id_input)
        input_layout.addWidget(self.equipment_input)
        card_layout.addWidget(input_container, alignment=Qt.AlignmentFlag.AlignCenter)

        # --- Error Message Label ---
        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Arial", 16))
        self.error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.error_label.setStyleSheet("color: red;")
        card_layout.addWidget(self.error_label)

        # **Increased spacing before buttons**
        card_layout.addSpacing(1)

        # --- Submit Button ---
        submit_button = QPushButton("Submit")
        submit_button.setFont(QFont("Arial", 22, QFont.Weight.Bold))  # Larger, bolder text
        submit_button.setFixedWidth(260)
        submit_button.setFixedHeight(60)  # Slightly larger button for emphasis
        submit_button.setCursor(Qt.CursorShape.PointingHandCursor)
        submit_button.setStyleSheet("""
            QPushButton {
                background-color: #B22222;
                color: white;
                border-radius: 10px;
                padding: 10px;
                border: 3px solid #800000;
            }
            QPushButton:hover {
                background-color: #FF5733;
                border: 3px solid #B22222;
            }
            QPushButton:pressed {
                background-color: #800000;
            }
        """)
        submit_button.clicked.connect(self.process_equipment_entry)
        card_layout.addWidget(submit_button, alignment=Qt.AlignmentFlag.AlignCenter)

        # **Spacing between submit button and other buttons**
        card_layout.addSpacing(30)

        # --- Button Container for "Filler Assignments" and "Problem Equipment" ---
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        button_layout.setSpacing(40)  # Increased spacing between buttons

        # --- Common Button Style for Filler Assignments and Problem Equipment ---
        common_button_style = """
            QPushButton {
                background-color: #B22222;
                color: white;
                border-radius: 10px;
                padding: 10px;
                font-size: 18px;
                border: 2px solid #800000;
            }
            QPushButton:hover {
                background-color: #FF5733;
                border: 2px solid #B22222;
            }
            QPushButton:pressed {
                background-color: #800000;
            }
        """

        # --- Filler Assignment Button ---
        filler_button = QPushButton("Filler Assignment")
        filler_button.setFont(QFont("Arial", 20))
        filler_button.setFixedWidth(260)
        filler_button.setFixedHeight(50)
        filler_button.setCursor(Qt.CursorShape.PointingHandCursor)
        filler_button.setStyleSheet(common_button_style)
        filler_button.clicked.connect(self.open_filler_assignment_window)

        # --- Problem Equipment Button ---
        problem_button = QPushButton("Problem Equipment")
        problem_button.setFont(QFont("Arial", 20))
        problem_button.setFixedWidth(260)
        problem_button.setFixedHeight(50)
        problem_button.setCursor(Qt.CursorShape.PointingHandCursor)
        problem_button.setStyleSheet(common_button_style)
        problem_button.clicked.connect(self.open_problem_equipment_window)

        # Add buttons to the layout
        button_layout.addWidget(filler_button)
        button_layout.addWidget(problem_button)
        card_layout.addWidget(button_container, alignment=Qt.AlignmentFlag.AlignCenter)

        # --- Last Action Label ---
        self.last_action_label = QLabel("")
        self.last_action_label.setFont(QFont("Arial", 18))
        self.last_action_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.last_action_label.setStyleSheet("color: blue;")
        card_layout.addWidget(self.last_action_label)

        layout.addWidget(card_widget)
        return page



    def open_filler_assignment_window(self, checked=False):
        """Opens a full-screen Tkinter window for assigning fillers to equipment numbers."""

        def create_filler_window():
            """Manages the filler assignments in a full-screen Tkinter window."""

            def update_table():
                """Fetch and display all records from the database in alphanumeric order."""
                for row in table.get_children():
                    table.delete(row)

                cursor.execute("SELECT employee_id, equipment_number FROM fillers ORDER BY equipment_number COLLATE NOCASE")
                records = cursor.fetchall()

                for record in records:
                    table.insert("", tk.END, values=record)

            def add_filler():
                """Add a new filler assignment to the database."""
                employee_id = employee_id_entry.get().strip()
                equipment_number = equipment_number_entry.get().strip().upper()

                if not employee_id or not equipment_number:
                    messagebox.showwarning("Invalid Input", "Please enter both Employee ID and Equipment Number.")
                    return

                cursor.execute("INSERT INTO fillers (employee_id, equipment_number) VALUES (?, ?)", (employee_id, equipment_number))
                conn.commit()
                update_table()
                employee_id_entry.delete(0, tk.END)
                equipment_number_entry.delete(0, tk.END)

            def modify_filler():
                """Modify a selected filler assignment."""
                selected_item = table.selection()
                if not selected_item:
                    messagebox.showwarning("No Selection", "Please select a row to modify.")
                    return

                row_values = table.item(selected_item, "values")
                if not row_values:
                    return

                old_employee_id, old_equipment_number = row_values

                def save_changes():
                    """Save the modified assignment."""
                    new_employee_id = employee_id_entry.get().strip()
                    new_equipment_number = equipment_number_entry.get().strip().upper()

                    if not new_employee_id or not new_equipment_number:
                        messagebox.showwarning("Invalid Input", "Both fields must be filled.")
                        return

                    cursor.execute("""
                        UPDATE fillers 
                        SET employee_id = ?, equipment_number = ? 
                        WHERE employee_id = ? AND equipment_number = ?
                    """, (new_employee_id, new_equipment_number, old_employee_id, old_equipment_number))
                    conn.commit()
                    update_table()
                    modify_window.destroy()

                # Create a pop-up window for editing
                modify_window = tk.Toplevel(root)
                modify_window.title("Modify Filler Assignment")
                modify_window.geometry("400x250")
                modify_window.transient(root)
                modify_window.grab_set()

                tk.Label(modify_window, text="Employee ID:", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=5, sticky="e")
                tk.Label(modify_window, text="Equipment Number:", font=("Arial", 14)).grid(row=1, column=0, padx=10, pady=5, sticky="e")

                employee_id_entry = tk.Entry(modify_window, font=("Arial", 14))
                employee_id_entry.grid(row=0, column=1, padx=10, pady=5)
                employee_id_entry.insert(0, old_employee_id)

                equipment_number_entry = tk.Entry(modify_window, font=("Arial", 14))
                equipment_number_entry.grid(row=1, column=1, padx=10, pady=5)
                equipment_number_entry.insert(0, old_equipment_number)

                save_button = tk.Button(modify_window, text="Save", font=("Arial", 14), command=save_changes)
                save_button.grid(row=2, column=0, columnspan=2, pady=10)

            def delete_filler():
                """Delete a selected filler assignment."""
                selected_item = table.selection()
                if not selected_item:
                    messagebox.showwarning("No Selection", "Please select a row to delete.")
                    return

                row_values = table.item(selected_item, "values")
                if not row_values:
                    return

                employee_id, equipment_number = row_values

                confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete assignment:\n\n{employee_id} -> {equipment_number}?")
                if confirm:
                    cursor.execute("DELETE FROM fillers WHERE employee_id = ? AND equipment_number = ?", (employee_id, equipment_number))
                    conn.commit()
                    update_table()

            # Initialize the full-screen Tkinter window
            root = tk.Tk()
            root.title("Filler Assignments")
            root.state('zoomed')

            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            window_width = int(screen_width * 0.6)
            root.geometry(f"{window_width}x{screen_height}+{int(screen_width * 0.2)}+0")  # Centered


            # SQLite Database Setup
            filler_db = os.path.join(os.path.expanduser("~"), "Desktop", "filler_assignments.db")
            conn = sqlite3.connect(filler_db)
            cursor = conn.cursor()

            # 🚨 **Fix Missing Columns** 🚨
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS fillers (
                    employee_id TEXT NOT NULL,
                    equipment_number TEXT NOT NULL
                )
            """)

            # **Check if employee_id and equipment_number exist, recreate if necessary**
            cursor.execute("PRAGMA table_info(fillers)")
            columns = [col[1] for col in cursor.fetchall()]

            if "employee_id" not in columns or "equipment_number" not in columns:
                cursor.execute("DROP TABLE fillers")  # Delete incorrect table
                cursor.execute("""
                    CREATE TABLE fillers (
                        employee_id TEXT NOT NULL,
                        equipment_number TEXT NOT NULL
                    )
                """)
                conn.commit()

            # Layout Configuration
            frame_top = tk.Frame(root)
            frame_top.pack(pady=20)

            tk.Label(frame_top, text="Employee ID:", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=5)
            tk.Label(frame_top, text="Equipment Number:", font=("Arial", 14)).grid(row=0, column=1, padx=10, pady=5)

            employee_id_entry = tk.Entry(frame_top, font=("Arial", 14))
            employee_id_entry.grid(row=1, column=0, padx=10, pady=5)

            equipment_number_entry = tk.Entry(frame_top, font=("Arial", 14))
            equipment_number_entry.grid(row=1, column=1, padx=10, pady=5)

            add_button = tk.Button(frame_top, text="Add Assignment", font=("Arial", 14), command=add_filler)
            add_button.grid(row=1, column=2, padx=10, pady=5)

            frame_table = tk.Frame(root, width=window_width)
            frame_table.pack(expand=True, fill=tk.BOTH, padx=int(screen_width * 0.2))  # Centered horizontally


            # Table for displaying assignments
            columns = ("Employee ID", "Equipment Number")
            table = ttk.Treeview(frame_table, columns=columns, show="headings", height=20)

            for col in columns:
                table.heading(col, text=col)
                table.column(col, anchor="center", width=int(window_width * 0.5))  # Each column takes 50% of table width

            table.pack(expand=True, fill=tk.BOTH)

            # Buttons for Modify and Delete
            frame_bottom = tk.Frame(root)
            frame_bottom.pack(pady=20)

            modify_button = tk.Button(frame_bottom, text="Modify Selected", font=("Arial", 14), command=modify_filler)
            modify_button.grid(row=0, column=0, padx=10)

            delete_button = tk.Button(frame_bottom, text="Delete Selected", font=("Arial", 14), command=delete_filler)
            delete_button.grid(row=0, column=1, padx=10)

            # Populate table initially
            update_table()

            # Run the Tkinter event loop
            root.mainloop()

        # Fix **Threading Issue**
        if threading.active_count() > 1:
            messagebox.showwarning("Warning", "Filler window is already open.")
            return

        # Start the filler assignment window in a separate thread to prevent GUI freezing
        filler_thread = threading.Thread(target=create_filler_window, daemon=True)
        filler_thread.start()

    def open_problem_equipment_window(self, checked=False):
        """Opens a full-screen Tkinter window while ensuring the data takes up only 60% of the screen width."""

        def create_problem_window():
            """Manages the problem equipment database in a full-screen Tkinter window with a 60% width data panel."""

            root = tk.Tk()
            root.title("Problem Equipment Management")
            root.state('zoomed')  # **Set full-screen mode**

            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            window_width = int(screen_width * 0.6)  # **Data takes up 60% of the screen width**

            conn = sqlite3.connect(self.database_path)
            cursor = conn.cursor()

            # **Ensure Problem Equipment Table Exists**
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS problem_equipment (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    equipment_number TEXT NOT NULL,
                    description TEXT NOT NULL,
                    timestamp TEXT NOT NULL
                )
            """)
            conn.commit()

            # **Main Frame Centering**
            frame_main = tk.Frame(root, width=window_width, height=screen_height)
            frame_main.pack(expand=True, fill=tk.BOTH)
            frame_main.place(relx=0.5, rely=0.5, anchor=tk.CENTER)  # Centered at 60% of the screen width

            # **Paned Window for Left (Table) and Right (Description)**
            pane = tk.PanedWindow(frame_main, orient=tk.HORIZONTAL, sashwidth=5)
            pane.pack(fill=tk.BOTH, expand=True)

            # **Left Frame: Equipment Table (Takes 60% of the data section)**
            frame_table = tk.Frame(pane, width=int(window_width * 0.5))
            pane.add(frame_table)

            # **Right Frame: Full Description Box (Takes 40% of the data section)**
            frame_description = tk.Frame(pane, width=int(window_width * 0.5))
            pane.add(frame_description)

            # **Equipment Table Setup**
            columns = ("Equipment Number", "Timestamp")
            table = ttk.Treeview(frame_table, columns=columns, show="headings", height=20)

            table.heading("Equipment Number", text="Equipment Number", anchor="center")
            table.column("Equipment Number", anchor="center", width=int(window_width * 0.3))  # Center text

            table.heading("Timestamp", text="Timestamp")
            table.column("Timestamp", anchor="center", width=int(window_width * 0.3))

            table.pack(expand=True, fill=tk.BOTH)

            # **Description Box on the Right Side (Smaller Height)**
            tk.Label(frame_description, text="Full Description:", font=("Arial", 14)).pack(pady=10)
            description_text = tk.Text(frame_description, font=("Arial", 14), height=24, wrap="word")
            description_text.pack(expand=False, fill=tk.X, padx=10, pady=10)

            # **Buttons Below the Table**
            frame_buttons = tk.Frame(frame_main)
            frame_buttons.pack(pady=10)

            tk.Label(frame_buttons, text="Equipment Number:", font=("Arial", 12)).grid(row=0, column=0, padx=5, pady=5)
            equipment_entry = tk.Entry(frame_buttons, font=("Arial", 12))
            equipment_entry.grid(row=0, column=1, padx=5, pady=5)

            tk.Label(frame_buttons, text="Description:", font=("Arial", 12)).grid(row=1, column=0, padx=5, pady=5)
            description_entry = tk.Text(frame_buttons, font=("Arial", 12), height=3, width=40)
            description_entry.grid(row=1, column=1, padx=5, pady=5)

            def add_problem():
                """Add a new problem record to the database."""
                equipment_number = equipment_entry.get().strip().upper()
                description = description_entry.get("1.0", tk.END).strip()

                if not equipment_number or not description:
                    messagebox.showwarning("Invalid Input", "Please enter both Equipment Number and Description.")
                    return

                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("INSERT INTO problem_equipment (equipment_number, description, timestamp) VALUES (?, ?, ?)", 
                            (equipment_number, description, timestamp))
                conn.commit()
                update_table()
                equipment_entry.delete(0, tk.END)
                description_entry.delete("1.0", tk.END)

            def update_table():
                """Fetch and display all records from the problem equipment database."""
                for row in table.get_children():
                    table.delete(row)

                cursor.execute("SELECT equipment_number, description, timestamp FROM problem_equipment ORDER BY timestamp DESC")
                records = cursor.fetchall()

                for record in records:
                    equipment_number, description, timestamp = record
                    table.insert("", tk.END, values=(equipment_number, timestamp), tags=(description,))

            def on_row_click(event):
                """Show full description in the text widget when a row is clicked."""
                selected_item = table.selection()
                if selected_item:
                    description_text.delete("1.0", tk.END)  # Clear previous text
                    description = table.item(selected_item, "tags")[0]  # Get stored description
                    description_text.insert(tk.END, description)  # Insert full description

            def delete_problem():
                """Delete a selected problem record."""
                selected_item = table.selection()
                if not selected_item:
                    messagebox.showwarning("No Selection", "Please select a row to delete.")
                    return

                equipment_number = table.item(selected_item, "values")[0]
                confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete this record for {equipment_number}?")
                if confirm:
                    cursor.execute("DELETE FROM problem_equipment WHERE equipment_number = ?", (equipment_number,))
                    conn.commit()
                    update_table()
                    description_text.delete("1.0", tk.END)  # Clear description box after deletion

            # **Bind row selection event**
            table.bind("<ButtonRelease-1>", on_row_click)

            # **Add and Delete Buttons**
            button_frame = tk.Frame(frame_main)
            button_frame.pack(pady=10)

            add_button = tk.Button(button_frame, text="Add Problem", font=("Arial", 12), command=add_problem, width=15)
            add_button.grid(row=0, column=0, padx=5, pady=5)

            delete_button = tk.Button(button_frame, text="Delete Selected", font=("Arial", 12), command=delete_problem, width=15)
            delete_button.grid(row=0, column=1, padx=5, pady=5)

            # **Populate table initially**
            update_table()
            root.mainloop()

        # **Start problem equipment window in a separate thread**
        problem_thread = threading.Thread(target=create_problem_window, daemon=True)
        problem_thread.start()


    def check_and_submit_employee_id(self):
        """Checks if the input is 6 digits and triggers the 'Enter' key event."""
        input_text = self.employee_id_input.text().strip()
        if input_text.isdigit() and (len(input_text) == 6 or len(input_text) == 5):  # Check if input is exactly 6 digits
            self.employee_id_input.returnPressed.emit()  # Emit the returnPressed signal

    def check_and_submit_equipment_number(self):
        """Checks if the input matches the valid equipment number format and triggers 'Enter'."""
        input_text = self.equipment_input.text().strip()
        if self.validate_equipment_number(input_text):  # Use your existing validation method
            self.equipment_input.returnPressed.emit()  # Emit the returnPressed signal


    def process_equipment_entry(self):
        """Handles user input for Employee ID and Equipment Number, ensuring correct field visibility and validation."""
        print("Processing Equipment Entry...")  # Debugging Print

        input_text = self.employee_id_input.text().strip()
        print(f"User Input: {input_text}")  # Debugging Print

        if not hasattr(self, "validate_equipment_number"):
            self.error_label.setText("Error: Equipment validation function is missing.")
            return

        # If the Equipment Number field is visible, process as a sign-out
        if self.equipment_input.isVisible():
            equipment_number = self.equipment_input.text().strip().upper()

            if not equipment_number:
                self.error_label.setText("Please enter an Equipment Number.")
                return

            if not self.validate_equipment_number(equipment_number):
                self.error_label.setText("Invalid Equipment Number. Must start with TM, RF, GN, and be followed by 3 digits.")
                self.equipment_input.clear()
                return

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # **🚨 Check if the same equipment type was already assigned today 🚨**
            prefix = equipment_number[:2]  # Extracts "TM", "RF", "GN" etc.
            today = datetime.now().strftime('%Y-%m-%d')

            conn = sqlite3.connect(self.database_path)
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM equipment_log 
                WHERE employee_id = ? 
                AND equipment_number LIKE ?
                AND timestamp LIKE ?
            """, (self.current_employee_id, f"{prefix}%", f"{today}%"))
            count = cursor.fetchone()[0]
            conn.close()

            if count > 0:
                messagebox.showwarning("Duplicate Equipment Type", f"Employee {self.current_employee_id} has already been assigned a {prefix} equipment today.")

            # Log the sign-out attempt
            self.log_to_database(self.current_employee_id, equipment_number, timestamp)

            # Update the last action display
            self.display_last_action(timestamp, self.current_employee_id, equipment_number, sign_out=True)

            self.reset_inputs()
            return

        # 🚨 FIX: Prevent infinite loop by handling Employee ID input correctly
        if input_text.isdigit() and (len(input_text) == 6 or len(input_text) == 5):
            if not self.equipment_input.isVisible():
                self.current_employee_id = input_text
                print(f"Employee ID {input_text} entered. Awaiting Equipment Number...")

                # Show Equipment Number field and focus on it
                self.equipment_input.setVisible(True)
                self.equipment_input.setFocus()
                self.error_label.setText("")  # Clear error message
                return

        # If input matches Equipment Number format, assume sign-in attempt
        if self.validate_equipment_number(input_text):
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"Equipment {input_text} signed back in at {timestamp}.")

            # Log the sign-in attempt
            self.log_to_database("N/A", input_text, timestamp)

            # Update the last action display
            self.display_last_action(timestamp, "N/A", input_text, sign_out=False)

            self.reset_inputs()
            return

        self.error_label.setText("Invalid input. Please enter a valid Employee ID or Equipment Number.")

        
    def display_last_action(self, timestamp, employee_id, equipment_number, sign_out):
        """Displays the last sign-in or sign-out attempt below the submit button."""
        
        # 🚨 FIX: Remove unnecessary `input_text` reference
        if sign_out:
            message = f"Last Sign-Out: {timestamp}\nEmployee ID: {employee_id}\nEquipment: {equipment_number}"
        else:
            message = f"Last Sign-In: {timestamp}\nEquipment: {equipment_number}"

        # Update the label with the last action
        self.last_action_label.setText(message)
        self.last_action_label.setStyleSheet("color: blue; font-size: 18px;")

    def validate_equipment_number(self, equipment_number):
        """Validates if the equipment number follows the correct format."""
        valid_prefixes = ["TM", "RF", "GN"]  # Ensure valid prefixes are correct
        return (
            len(equipment_number) == 5
            and equipment_number[:2] in valid_prefixes
            and equipment_number[2:].isdigit()
        )

    def reset_inputs(self):
        """Clears all input fields, resets visibility, and refocuses the cursor on Employee ID."""
        self.employee_id_input.clear()
        self.equipment_input.clear()
        self.equipment_input.setVisible(False)  # Hide Equipment Number field
        self.error_label.setText("")  # Clear error messages
        self.current_employee_id = None

        # 🚀 Automatically focus on Employee ID input field
        self.employee_id_input.setFocus()

    def log_to_database(self, employee_id, equipment_number, timestamp):
        """Logs the equipment entry into the SQLite database with a timestamp."""
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO equipment_log (timestamp, employee_id, equipment_number)
            VALUES (?, ?, ?)
        """, (timestamp, employee_id, equipment_number))

        conn.commit()
        conn.close()

    def create_rma_page(self):
        """Creates the RMA Page with links to Zebra and Honeywell RMA portals and a database management button."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Title Label
        title = QLabel("RMA Portal Links")
        title.setFont(QFont("Arial", 36, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Common Button Style
        button_style = """
            QPushButton {
                background-color: #B22222;
                color: white;
                border-radius: 10px;
                padding: 15px;
                font-size: 20px;
                border: 2px solid #800000;
            }
            QPushButton:hover {
                background-color: #FF5733;
                border: 2px solid #B22222;
            }
            QPushButton:pressed {
                background-color: #800000;
            }
        """

        # Button Widths
        button_width = 300
        combined_button_width = button_width * 2 + 40  # Matches Zebra & Honeywell combined width

        # Zebra RMA Button
        zebra_button = QPushButton("Zebra RMA Portal")
        zebra_button.setFont(QFont("Arial", 22))
        zebra_button.setFixedWidth(button_width)
        zebra_button.setFixedHeight(60)
        zebra_button.setCursor(Qt.CursorShape.PointingHandCursor)
        zebra_button.setStyleSheet(button_style)
        zebra_button.clicked.connect(lambda: webbrowser.open("https://www.zebra.com/us/en/support-downloads/request-repair.html"))

        # Honeywell RMA Button
        honeywell_button = QPushButton("Honeywell RMA Portal")
        honeywell_button.setFont(QFont("Arial", 22))
        honeywell_button.setFixedWidth(button_width)
        honeywell_button.setFixedHeight(60)
        honeywell_button.setCursor(Qt.CursorShape.PointingHandCursor)
        honeywell_button.setStyleSheet(button_style)
        honeywell_button.clicked.connect(lambda: webbrowser.open("https://sps-support.honeywell.com/s/pss/pss-rma"))

        # Horizontal Layout for Portal Buttons
        button_row = QHBoxLayout()
        button_row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        button_row.addWidget(zebra_button)
        button_row.addSpacing(40)  # Space between buttons
        button_row.addWidget(honeywell_button)

        # RMA Database Button
        rma_button = QPushButton("Manage RMAs")
        rma_button.setFont(QFont("Arial", 22))
        rma_button.setFixedWidth(combined_button_width)  # Matches the width of both Zebra & Honeywell buttons combined
        rma_button.setFixedHeight(60)
        rma_button.setCursor(Qt.CursorShape.PointingHandCursor)
        rma_button.setStyleSheet(button_style)
        rma_button.clicked.connect(self.open_rma_management_window)

        # Add Components to Layout
        layout.addWidget(title)
        layout.addSpacing(30)
        layout.addLayout(button_row)  # Horizontal row of buttons
        layout.addSpacing(30)
        layout.addWidget(rma_button, alignment=Qt.AlignmentFlag.AlignCenter)

        return page
    
    def open_rma_management_window(self):
        """Opens a Tkinter window for managing RMA records without displaying the ID column, adjusting column width to 60% of screen width."""

        def create_rma_window():
            """Creates and manages the Tkinter RMA database window."""
            def setup_database():
                """Ensures the RMA database table exists."""
                conn = sqlite3.connect(self.database_path)
                cursor = conn.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS rma_records (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        equipment_number TEXT NOT NULL,
                        serial_number TEXT NOT NULL,
                        date TEXT NOT NULL,
                        rma_number TEXT
                    )
                """)
                conn.commit()
                conn.close()

            def update_table():
                """Fetch and display all RMA records."""
                for row in table.get_children():
                    table.delete(row)

                conn = sqlite3.connect(self.database_path)
                cursor = conn.cursor()
                cursor.execute("SELECT equipment_number, serial_number, date, rma_number FROM rma_records ORDER BY date DESC")
                records = cursor.fetchall()
                conn.close()

                for record in records:
                    table.insert("", tk.END, values=record)

            def add_rma_record():
                """Adds a new RMA record to the database."""
                equipment_number = equipment_entry.get().strip().upper()
                serial_number = serial_entry.get().strip().upper()
                date = date_entry.get().strip()
                rma_number = rma_entry.get().strip()

                if not equipment_number or not serial_number or not date:
                    messagebox.showwarning("Invalid Input", "Please enter Equipment Number, Serial Number, and Date.")
                    return

                conn = sqlite3.connect(self.database_path)
                cursor = conn.cursor()
                cursor.execute("INSERT INTO rma_records (equipment_number, serial_number, date, rma_number) VALUES (?, ?, ?, ?)",
                            (equipment_number, serial_number, date, rma_number))
                conn.commit()
                conn.close()

                update_table()
                equipment_entry.delete(0, tk.END)
                serial_entry.delete(0, tk.END)
                date_entry.delete(0, tk.END)
                rma_entry.delete(0, tk.END)

            def delete_rma_record():
                """Deletes a selected RMA record."""
                selected_item = table.selection()
                if not selected_item:
                    messagebox.showwarning("No Selection", "Please select a row to delete.")
                    return

                row_values = table.item(selected_item, "values")
                if not row_values:
                    return

                equipment_number, serial_number, date, rma_number = row_values

                confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this record?")
                if confirm:
                    conn = sqlite3.connect(self.database_path)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM rma_records WHERE equipment_number = ? AND serial_number = ? AND date = ? AND rma_number IS ?",
                                (equipment_number, serial_number, date, rma_number if rma_number else None))
                    conn.commit()
                    conn.close()
                    update_table()

            # Initialize Tkinter Window
            root = tk.Tk()
            root.title("RMA Management")
            root.state('zoomed')

            # Ensure Database is Setup
            setup_database()

            # Input Frame
            input_frame = tk.Frame(root)
            input_frame.pack(pady=20)

            tk.Label(input_frame, text="Equipment Number:", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=5)
            tk.Label(input_frame, text="Serial Number:", font=("Arial", 14)).grid(row=0, column=1, padx=10, pady=5)
            tk.Label(input_frame, text="Date (YYYY-MM-DD):", font=("Arial", 14)).grid(row=0, column=2, padx=10, pady=5)
            tk.Label(input_frame, text="RMA Number (Optional):", font=("Arial", 14)).grid(row=0, column=3, padx=10, pady=5)

            equipment_entry = tk.Entry(input_frame, font=("Arial", 14))
            equipment_entry.grid(row=1, column=0, padx=10, pady=5)

            serial_entry = tk.Entry(input_frame, font=("Arial", 14))
            serial_entry.grid(row=1, column=1, padx=10, pady=5)

            date_entry = tk.Entry(input_frame, font=("Arial", 14))
            date_entry.grid(row=1, column=2, padx=10, pady=5)

            rma_entry = tk.Entry(input_frame, font=("Arial", 14))
            rma_entry.grid(row=1, column=3, padx=10, pady=5)

            add_button = tk.Button(input_frame, text="Add RMA Record", font=("Arial", 14), command=add_rma_record)
            add_button.grid(row=1, column=4, padx=10, pady=5)

            # **Table for Displaying RMA Records (ID Column Removed & Width Adjusted to 60% of Screen)**
            columns = ("Equipment Number", "Serial Number", "Date", "RMA Number")
            table = ttk.Treeview(root, columns=columns, show="headings", height=20)

            for col in columns:
                table.heading(col, text=col)

            # **Set column widths dynamically to take up 60% of the window**
            screen_width = root.winfo_screenwidth()
            table_width = int(screen_width * 0.6)
            col_width = table_width // len(columns)  # Divide evenly across columns

            for col in columns:
                table.column(col, anchor="center", width=col_width)

            table.pack(expand=True, fill=tk.BOTH, padx=int(screen_width * 0.2))  # Center the table

            # Delete Button
            delete_button = tk.Button(root, text="Delete Selected Record", font=("Arial", 14), command=delete_rma_record)
            delete_button.pack(pady=10)

            update_table()
            root.mainloop()

        threading.Thread(target=create_rma_window, daemon=True).start()



    def create_print_page(self):
        """Creates the Print Page"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Print Page")
        title.setFont(QFont("Arial", 36))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(title)
        return page


    def open_tkinter_database_window(self):
        """Open a Tkinter window to display the entire database."""
        
        def create_tk_window():
            """Create and display the Tkinter window."""
            root = tk.Tk()
            root.title("Database Contents")
            root.geometry("800x600")  # Set a default size for the window

            # Set up the Treeview widget (similar to a table)
            tree = ttk.Treeview(root, columns=("ID", "Timestamp", "Employee ID", "Equipment Number"), show="headings")
            tree.heading("ID", text="ID")
            tree.heading("Timestamp", text="Timestamp")
            tree.heading("Employee ID", text="Employee ID")
            tree.heading("Equipment Number", text="Equipment Number")

            # Fetch all records from the database
            conn = sqlite3.connect(self.database_path)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM equipment_log")
            records = cursor.fetchall()

            # Insert records into the Treeview
            for record in records:
                tree.insert("", tk.END, values=record)

            # Add the Treeview to the window and pack it
            tree.pack(expand=True, fill=tk.BOTH)

            # Run the Tkinter event loop
            root.mainloop()

        # Create and start the Tkinter window in a separate thread
        thread = Thread(target=create_tk_window)
        thread.start()
            
    def query_entire_database(self):
        """Query and display all records from the entire database."""
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        cursor.execute("SELECT timestamp, employee_id, equipment_number FROM equipment_log ORDER BY timestamp")
        records = cursor.fetchall()
        conn.close()
        self.open_tkinter_window("Entire Database", records)    

    def display_database_contents(self):
        """Display all records from the entire database."""
        self.open_tkinter_window("Entire Database")

    def execute_query(self):
        """Executes the database query based on the selected dropdown option."""
        selected_query = self.query_type_combo.currentText()

        if selected_query == "Entire Database":
            self.query_entire_database()
        elif selected_query == "Query Today":
            self.query_today()
        elif selected_query == "Query Non-Turned In Equipment":
            self.query_non_turned_in_equipment()
        elif selected_query == "Query Last Month":
            self.query_last_month()
        elif selected_query == "Query Mobile Printers":
            self.query_mobile_printers()
        elif selected_query == "Query Radios":
            self.query_radios()
        elif selected_query == "Query Filler Sets":
            self.query_filler_sets()
        elif selected_query == "Query RF Guns":
            self.query_rf_guns()


    def query_today(self):
        """Query and display records for today."""
        today = datetime.today().strftime('%Y-%m-%d')
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        cursor.execute("SELECT timestamp, employee_id, equipment_number FROM equipment_log WHERE timestamp LIKE ?", (today + "%",))
        records = cursor.fetchall()
        conn.close()
        self.open_tkinter_window(f"Records for {today}", records)

    def query_non_turned_in_equipment(self):
        """Query and display equipment that has been signed out but NOT signed back in."""
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        query = """
            SELECT timestamp, employee_id, equipment_number
            FROM equipment_log e1
            WHERE e1.id IN (
                SELECT MAX(id) FROM equipment_log WHERE employee_id != 'N/A' GROUP BY equipment_number
            )
            AND (
                (substr(e1.equipment_number, 1, 2) = 'TM' AND substr(e1.equipment_number, 3, 3) BETWEEN '001' AND '102') OR
                (substr(e1.equipment_number, 1, 2) = 'RF' AND substr(e1.equipment_number, 3, 3) BETWEEN '001' AND '216') OR
                (substr(e1.equipment_number, 1, 2) = 'RD' AND substr(e1.equipment_number, 3, 3) BETWEEN '001' AND '024') OR
                (substr(e1.equipment_number, 1, 2) = 'ZM' AND substr(e1.equipment_number, 3, 3) BETWEEN '001' AND '024')
            )
            AND NOT EXISTS (
                SELECT 1 FROM equipment_log e2 WHERE e2.equipment_number = e1.equipment_number
                AND e2.employee_id = 'N/A' AND e2.timestamp > e1.timestamp
            )
            ORDER BY e1.equipment_number COLLATE NOCASE
        """
        cursor.execute(query)
        records = cursor.fetchall()
        conn.close()
        self.open_tkinter_window("Non-Turned In Equipment", records)


    def query_last_month(self):
        """Query and display records for the last month."""
        last_month = (datetime.today().replace(day=1) - timedelta(days=1)).replace(day=1).strftime('%Y-%m-%d')
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        cursor.execute("SELECT timestamp, employee_id, equipment_number FROM equipment_log WHERE timestamp >= ?", (last_month,))
        records = cursor.fetchall()
        conn.close()
        self.open_tkinter_window("Records for Last Month", records)

    def query_radios(self):
        """Query and display the last sign-out instance for each unique Radio (equipment numbers start with 'RD'),
        sorted alphanumerically by equipment number."""
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT e1.timestamp, e1.employee_id, e1.equipment_number
            FROM equipment_log e1
            WHERE e1.equipment_number LIKE 'RD%'
            AND e1.timestamp = (
                SELECT MAX(e2.timestamp) 
                FROM equipment_log e2 
                WHERE e2.equipment_number = e1.equipment_number
            )
            ORDER BY e1.equipment_number COLLATE NOCASE
        """)
        records = cursor.fetchall()
        conn.close()
        if not records:
            print(f"\n--- No results found for Last Sign-Out for Radios ---")
        self.open_tkinter_window("Last Sign-Out for Radios", records)


    def query_mobile_printers(self):
        """Query and display the last sign-out instance for each unique Mobile Printer (equipment numbers start with 'ZM')."""
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT e1.timestamp, e1.employee_id, e1.equipment_number
            FROM equipment_log e1
            WHERE e1.equipment_number LIKE 'ZM%'
            AND e1.timestamp = (
                SELECT MAX(e2.timestamp) 
                FROM equipment_log e2 
                WHERE e2.equipment_number = e1.equipment_number
            )
            ORDER BY e1.equipment_number COLLATE NOCASE
        """)
        records = cursor.fetchall()
        conn.close()
        if not records:
            print(f"\n--- No results found for Last Sign-Out for Mobile Printers ---")
        self.open_tkinter_window("Last Sign-Out for Mobile Printers", records)


    def query_filler_sets(self):
        """Query and display the last sign-out instance for each unique Filler Set (equipment numbers start with 'TM')."""
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT e1.timestamp, e1.employee_id, e1.equipment_number
            FROM equipment_log e1
            WHERE e1.equipment_number LIKE 'TM%'
            AND LENGTH(e1.equipment_number) = 5
            AND substr(e1.equipment_number, 3, 3) BETWEEN '001' AND '102'
            AND e1.timestamp = (
                SELECT MAX(e2.timestamp) 
                FROM equipment_log e2 
                WHERE e2.equipment_number = e1.equipment_number
            )
            ORDER BY e1.equipment_number COLLATE NOCASE
        """)
        records = cursor.fetchall()
        conn.close()
        if not records:
            print(f"\n--- No results found for Last Sign-Out for Filler Sets ---")
        self.open_tkinter_window("Last Sign-Out for Filler Sets", records)


    def query_rf_guns(self):
        """Query and display the last sign-out instance for each unique RF Gun (equipment numbers start with 'RF' and follow the format 'RF001' to 'RF216')."""
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT e1.timestamp, e1.employee_id, e1.equipment_number
            FROM equipment_log e1
            WHERE e1.equipment_number LIKE 'RF%'
            AND LENGTH(e1.equipment_number) = 5
            AND substr(e1.equipment_number, 3, 3) BETWEEN '001' AND '216'
            AND e1.timestamp = (
                SELECT MAX(e2.timestamp) 
                FROM equipment_log e2 
                WHERE e2.equipment_number = e1.equipment_number
            )
            ORDER BY e1.equipment_number COLLATE NOCASE
        """)
        
        records = cursor.fetchall()
        conn.close()
        
        if not records:
            print("\n--- No results found for Last Sign-Out for RF Guns ---")
        
        self.open_tkinter_window("Last Sign-Out for RF Guns", records)



    def open_tkinter_window(self, query_type, records=None):
        """Opens a Tkinter window to display query results, ensuring the table width is 60% of the screen width."""

        def create_tk_window():
            """Creates and displays the Tkinter window with the adjusted table width."""
            root = tk.Tk()
            root.title(query_type)
            root.state('zoomed')  # Full-screen mode

            # Table Columns
            columns = ("Timestamp", "Employee ID", "Equipment Number")
            table = ttk.Treeview(root, columns=columns, show="headings", style="Treeview")

            # Set Column Headings
            for col in columns:
                table.heading(col, text=col)

            # **Set column widths dynamically to take up 60% of the window**
            screen_width = root.winfo_screenwidth()
            table_width = int(screen_width * 0.6)
            col_width = table_width // len(columns)  # Divide evenly across columns

            for col in columns:
                table.column(col, anchor="center", width=col_width)

            # Pack the Table and Center It
            table.pack(expand=True, fill=tk.BOTH, padx=int(screen_width * 0.2))

            # Ensure the table is cleared before inserting new data
            for row in table.get_children():
                table.delete(row)

            # Populate Table with Records
            if records:
                for record in records:
                    table.insert("", tk.END, values=record)

            # If no records are found, show a message in the console
            if not records:
                print(f"\n--- No results found for {query_type} ---")

            # **Frame for action buttons**
            button_frame = tk.Frame(root)
            button_frame.pack(side=tk.BOTTOM, pady=20)

            # Modify Row Button
            modify_button = tk.Button(button_frame, text="Modify Row", command=lambda: self.modify_row(table),
                                    font=("Arial", 14), width=15)
            # Delete Row Button
            delete_button = tk.Button(button_frame, text="Delete Row", command=lambda: self.delete_row(table),
                                    font=("Arial", 14), width=15)
            # Save to Spreadsheet Button
            save_button = tk.Button(button_frame, text="Save to Spreadsheet", command=lambda: self.save_to_spreadsheet(query_type, records),
                                    font=("Arial", 14), width=20)

            # Pack buttons side by side
            modify_button.pack(side=tk.LEFT, padx=10)
            delete_button.pack(side=tk.LEFT, padx=10)
            save_button.pack(side=tk.LEFT, padx=10)

            # Close window function
            def on_closing():
                root.destroy()

            root.protocol("WM_DELETE_WINDOW", on_closing)

            # Run the Tkinter event loop
            root.mainloop()

        # Start Tkinter window in a separate thread
        threading.Thread(target=create_tk_window, daemon=True).start()



    def modify_row(self, tree):
        """Open a pop-up window to modify a selected row and update the database."""
        selected_item = tree.selection()

        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a row to modify.")
            return

        # Get the selected row's values
        row_values = tree.item(selected_item, "values")
        if not row_values or len(row_values) != 3:
            messagebox.showwarning("Error", "Invalid row selection.")
            return

        # Extract data (No ID column anymore)
        timestamp, employee_id, equipment_number = row_values

        # Create a new pop-up window
        edit_window = tk.Toplevel()
        edit_window.title("Modify Row")
        edit_window.geometry("400x300")
        edit_window.transient(tree)  # Attach to the main window
        edit_window.grab_set()  # Make it modal (disables interaction with the main window)

        # Labels and Entry Fields
        tk.Label(edit_window, text="Timestamp:", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=5, sticky="e")
        tk.Label(edit_window, text="Employee ID:", font=("Arial", 14)).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        tk.Label(edit_window, text="Equipment Number:", font=("Arial", 14)).grid(row=2, column=0, padx=10, pady=5, sticky="e")

        timestamp_entry = tk.Entry(edit_window, font=("Arial", 14))
        timestamp_entry.grid(row=0, column=1, padx=10, pady=5)
        timestamp_entry.insert(0, timestamp)

        employee_id_entry = tk.Entry(edit_window, font=("Arial", 14))
        employee_id_entry.grid(row=1, column=1, padx=10, pady=5)
        employee_id_entry.insert(0, employee_id)

        equipment_number_entry = tk.Entry(edit_window, font=("Arial", 14))
        equipment_number_entry.grid(row=2, column=1, padx=10, pady=5)
        equipment_number_entry.insert(0, equipment_number)

        def save_changes():
            """Save the modified data to the database and update the table."""
            new_timestamp = timestamp_entry.get()
            new_employee_id = employee_id_entry.get()
            new_equipment_number = equipment_number_entry.get()

            if not new_timestamp or not new_employee_id or not new_equipment_number:
                messagebox.showwarning("Invalid Input", "Fields cannot be empty.")
                return

            # Update the database
            conn = sqlite3.connect(self.database_path)
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE equipment_log
                SET timestamp = ?, employee_id = ?, equipment_number = ?
                WHERE timestamp = ? AND employee_id = ? AND equipment_number = ?
            """, (new_timestamp, new_employee_id, new_equipment_number, timestamp, employee_id, equipment_number))
            conn.commit()
            conn.close()

            # Update the table in Tkinter
            tree.item(selected_item, values=(new_timestamp, new_employee_id, new_equipment_number))
            messagebox.showinfo("Success", "Row modified successfully.")

            edit_window.destroy()  # Close the pop-up window

        # Save Button
        save_button = tk.Button(edit_window, text="Save", font=("Arial", 14), command=save_changes)
        save_button.grid(row=3, column=0, columnspan=2, pady=20)

        edit_window.mainloop()

    def delete_row(self, tree):
        """Delete the selected row from the database and the Tkinter table."""
        selected_item = tree.selection()

        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a row to delete.")
            return

        # Get the selected row's values
        row_values = tree.item(selected_item, "values")
        if not row_values:
            return

        row_id = row_values[0]

        # Confirm deletion
        confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this row?")
        if not confirm:
            return

        # Delete from the database
        conn = sqlite3.connect(self.database_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM equipment_log WHERE id = ?", (row_id,))
        conn.commit()
        conn.close()

        # Remove from Tkinter table
        tree.delete(selected_item)
        messagebox.showinfo("Success", "Row deleted successfully.")

    def save_to_spreadsheet(self, query_type, records):
        """Save the current query results to an Excel spreadsheet in the 'Ace Logs' folder on the Desktop."""
        if not records:
            messagebox.showwarning("No Data", "No data available to save.")
            return

        # Get desktop path and create "Ace Logs" folder if it doesn't exist
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        logs_folder = os.path.join(desktop_path, "Ace Logs")

        # Ensure the "Ace Logs" folder exists
        try:
            os.makedirs(logs_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Could not create 'Ace Logs' folder: {e}")
            return

        # Generate filename with query type and timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{query_type.replace(' ', '_')}_{timestamp}.xlsx"
        filepath = os.path.join(logs_folder, filename)

        # **Dynamically adjust column headers**
        num_columns = len(records[0]) if records else 3  # Default to 3 columns if no data
        column_headers = ["Timestamp", "Employee ID", "Equipment Number"] if num_columns == 3 else ["ID", "Timestamp", "Employee ID", "Equipment Number"]

        # Convert records to a DataFrame and save as Excel file
        try:
            df = pd.DataFrame(records, columns=column_headers)
            df.to_excel(filepath, index=False)

            # Show success message in a temporary Tkinter window
            root = tk.Tk()
            root.withdraw()  # Hide main window
            messagebox.showinfo("Success", f"Spreadsheet saved successfully!\n\nLocation:\n{filepath}")
            root.destroy()  # Close Tkinter window after message
        except Exception as e:
            messagebox.showerror("Error", f"Could not save spreadsheet: {e}")

# Run the Application
app = QApplication(sys.argv)
window = WebStyleApp()
window.show()
sys.exit(app.exec())

