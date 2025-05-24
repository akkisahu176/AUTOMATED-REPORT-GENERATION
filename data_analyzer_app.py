import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import tempfile
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import numpy as np
import io
import os
# Set the appearance mode and default color theme
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class DataAnalyzerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Data Analyzer")
        self.geometry("1000x650")
        self.minsize(800, 600)
        
        # Initialize variables
        self.data = None
        self.file_path = None
        self.preview_widget = None
        self.report_title = "Data Analysis Report"
        self.temp_image_files = []
        
        # Create the main layout
        self.create_layout()
        
    def create_layout(self):
        # Create main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Left side: Controls
        self.controls_frame = ctk.CTkFrame(self.main_frame, width=300)
        self.controls_frame.pack(side="left", fill="y", padx=10, pady=10)
        
        # Right side: Preview area
        self.preview_frame = ctk.CTkFrame(self.main_frame)
        self.preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        
        # Add widgets to the control frame
        self.create_control_widgets()
        
        # Add widgets to the preview frame
        self.create_preview_widgets()
        
    def create_control_widgets(self):
                # Create main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Left side: Controls
        self.controls_frame = ctk.CTkFrame(self.main_frame, width=300)
        self.controls_frame.pack(side="left", fill="y", padx=10, pady=10)
        
        # Right side: Preview area
        self.preview_frame = ctk.CTkFrame(self.main_frame)
        self.preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        
        # Add widgets to the control frame
        self.create_control_widgets()
        
        # Add widgets to the preview frame
        self.create_preview_widgets()
        
    def create_control_widgets(self):
        # Create a scrollable frame for controls so they don't go off-screen
        self.controls_scrollable = ctk.CTkScrollableFrame(self.controls_frame)
        self.controls_scrollable.pack(fill="both", expand=True, padx=5, pady=5)
        
        # File upload section
        self.file_section = ctk.CTkFrame(self.controls_scrollable)
        self.file_section.pack(fill="x", padx=10, pady=(10, 5))
        
        self.file_label = ctk.CTkLabel(self.file_section, text="Data File", font=("Arial", 14, "bold"))
        self.file_label.pack(anchor="w", padx=5, pady=5)
        
        self.file_button = ctk.CTkButton(self.file_section, text="Upload File", command=self.upload_file)
        self.file_button.pack(fill="x", padx=5, pady=5)
        
        self.file_info = ctk.CTkLabel(self.file_section, text="No file selected", wraplength=280)
        self.file_info.pack(fill="x", padx=5, pady=5)
        
        # Analysis options section
        self.analysis_section = ctk.CTkFrame(self.controls_scrollable)
        self.analysis_section.pack(fill="x", padx=10, pady=5)
        
        self.analysis_label = ctk.CTkLabel(self.analysis_section, text="Analysis Options", font=("Arial", 14, "bold"))
        self.analysis_label.pack(anchor="w", padx=5, pady=5)
        
        # Checkboxes for analysis types
        self.summary_stats_var = ctk.BooleanVar(value=True)
        self.summary_stats_cb = ctk.CTkCheckBox(self.analysis_section, text="Summary Statistics", 
                                               variable=self.summary_stats_var)
        self.summary_stats_cb.pack(anchor="w", padx=5, pady=2)
        
        self.value_counts_var = ctk.BooleanVar(value=True)
        self.value_counts_cb = ctk.CTkCheckBox(self.analysis_section, text="Value Counts", 
                                              variable=self.value_counts_var)
        self.value_counts_cb.pack(anchor="w", padx=5, pady=2)
        
        self.correlation_var = ctk.BooleanVar(value=True)
        self.correlation_cb = ctk.CTkCheckBox(self.analysis_section, text="Correlation Matrix", 
                                             variable=self.correlation_var)
        self.correlation_cb.pack(anchor="w", padx=5, pady=2)
        
        self.distribution_var = ctk.BooleanVar(value=True)
        self.distribution_cb = ctk.CTkCheckBox(self.analysis_section, text="Column Distribution", 
                                              variable=self.distribution_var)
        self.distribution_cb.pack(anchor="w", padx=5, pady=2)
        
        # Column selection
        self.column_label = ctk.CTkLabel(self.analysis_section, text="Column for Analysis")
        self.column_label.pack(anchor="w", padx=5, pady=(10, 2))
        
        self.column_var = ctk.StringVar()
        self.column_dropdown = ctk.CTkOptionMenu(self.analysis_section, variable=self.column_var, values=["No Data Loaded"])
        self.column_dropdown.pack(fill="x", padx=5, pady=2)
        
        # Chart types
        self.chart_label = ctk.CTkLabel(self.analysis_section, text="Chart Type")
        self.chart_label.pack(anchor="w", padx=5, pady=(10, 2))
        
        self.chart_type = ctk.CTkOptionMenu(self.analysis_section, values=["Bar Chart", "Line Chart", "Pie Chart", "Histogram"])
        self.chart_type.pack(fill="x", padx=5, pady=2)
        
        # Report section
        self.report_section = ctk.CTkFrame(self.controls_scrollable)
        self.report_section.pack(fill="x", padx=10, pady=5)
        
        self.report_label = ctk.CTkLabel(self.report_section, text="Report Options", font=("Arial", 14, "bold"))
        self.report_label.pack(anchor="w", padx=5, pady=5)
        
        # Report title
        self.title_label = ctk.CTkLabel(self.report_section, text="Report Title")
        self.title_label.pack(anchor="w", padx=5, pady=2)
        
        self.title_entry = ctk.CTkEntry(self.report_section)
        self.title_entry.pack(fill="x", padx=5, pady=2)
        self.title_entry.insert(0, self.report_title)
        
        # Report comment
        self.comment_label = ctk.CTkLabel(self.report_section, text="Comments")
        self.comment_label.pack(anchor="w", padx=5, pady=2)
        
        self.comment_text = ctk.CTkTextbox(self.report_section, height=60)
        self.comment_text.pack(fill="x", padx=5, pady=2)
        
        # Create a fixed frame at the bottom for action buttons
        self.action_buttons_frame = ctk.CTkFrame(self.controls_frame)
        self.action_buttons_frame.pack(fill="x", padx=10, pady=10, side="bottom")
        
        # Action buttons in the fixed bottom frame
        self.analyze_button = ctk.CTkButton(
            self.action_buttons_frame, 
            text="Analyze Data", 
            command=self.analyze_data,
            height=35,
            fg_color="#2a9d8f"
        )
        self.analyze_button.pack(fill="x", padx=5, pady=5)
        
        self.generate_pdf_button = ctk.CTkButton(
            self.action_buttons_frame, 
            text="Generate PDF Report", 
            command=self.generate_pdf,
            height=35,
            fg_color="#e76f51"
        )
        self.generate_pdf_button.pack(fill="x", padx=5, pady=5)
        
        # Export data option
        self.export_button = ctk.CTkButton(
            self.action_buttons_frame, 
            text="Export Analyzed Data", 
            command=self.export_data,
            height=35,
            fg_color="#457b9d"
        )
        self.export_button.pack(fill="x", padx=5, pady=5)
        
        self.reset_button = ctk.CTkButton(
            self.action_buttons_frame, 
            text="Reset", 
            command=self.reset_app,
            height=35,
            fg_color="#6c757d"
        )
        self.reset_button.pack(fill="x", padx=5, pady=5)
        
        # Appearance toggle
        self.appearance_label = ctk.CTkLabel(self.controls_scrollable, text="Appearance")
        self.appearance_label.pack(anchor="w", padx=15, pady=(10, 0))
        
        self.appearance_mode = ctk.CTkOptionMenu(self.controls_scrollable, 
                                               values=["System", "Light", "Dark"],
                                               command=self.change_appearance_mode)
        self.appearance_mode.pack(fill="x", padx=15, pady=(0, 10))
    def create_preview_widgets(self):
        # Create a tabview for different previews
        self.preview_tabview = ctk.CTkTabview(self.preview_frame)
        self.preview_tabview.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add tabs
        self.data_tab = self.preview_tabview.add("Data")
        self.stats_tab = self.preview_tabview.add("Statistics")
        self.chart_tab = self.preview_tabview.add("Charts")
        self.filter_tab = self.preview_tabview.add("Filter Data")
        
        # Add scrollable frame in the data tab
        self.data_scroll = ctk.CTkScrollableFrame(self.data_tab)
        self.data_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add a label in the stats tab
        self.stats_scroll = ctk.CTkScrollableFrame(self.stats_tab)
        self.stats_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add a frame for charts
        self.chart_canvas_frame = ctk.CTkFrame(self.chart_tab)
        self.chart_canvas_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add filter controls
        self.setup_filter_controls()
    
    def setup_filter_controls(self):
        """Setup the filter controls in the filter tab"""
        self.filter_frame = ctk.CTkFrame(self.filter_tab)
        self.filter_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Filter header
        self.filter_header_label = ctk.CTkLabel(
            self.filter_frame, 
            text="Filter Data", 
            font=("Arial", 16, "bold")
        )
        self.filter_header_label.pack(anchor="w", padx=10, pady=10)
        
        # Filter description
        self.filter_description = ctk.CTkLabel(
            self.filter_frame,
            text="Create filters to analyze specific subsets of your data.",
            wraplength=600
        )
        self.filter_description.pack(anchor="w", padx=10, pady=(0, 10))
        
        # Filter controls frame
        self.filter_controls_frame = ctk.CTkFrame(self.filter_frame)
        self.filter_controls_frame.pack(fill="x", padx=10, pady=5)
        
        # Column selection
        self.filter_column_label = ctk.CTkLabel(self.filter_controls_frame, text="Column:")
        self.filter_column_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.filter_column_var = ctk.StringVar()
        self.filter_column_dropdown = ctk.CTkOptionMenu(
            self.filter_controls_frame, 
            variable=self.filter_column_var,
            values=["No Data Loaded"],
            width=150
        )
        self.filter_column_dropdown.grid(row=0, column=1, padx=5, pady=5)
        
        # Operator selection
        self.filter_operator_label = ctk.CTkLabel(self.filter_controls_frame, text="Operator:")
        self.filter_operator_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        self.filter_operator_var = ctk.StringVar()
        self.filter_operator_dropdown = ctk.CTkOptionMenu(
            self.filter_controls_frame, 
            variable=self.filter_operator_var,
            values=["equals", "not equals", "greater than", "less than", "contains", "starts with", "ends with"],
            width=150
        )
        self.filter_operator_dropdown.grid(row=0, column=3, padx=5, pady=5)
        
        # Value entry
        self.filter_value_label = ctk.CTkLabel(self.filter_controls_frame, text="Value:")
        self.filter_value_label.grid(row=0, column=4, padx=5, pady=5, sticky="w")
        
        self.filter_value_entry = ctk.CTkEntry(self.filter_controls_frame, width=150)
        self.filter_value_entry.grid(row=0, column=5, padx=5, pady=5)
        
        # Add filter button
        self.add_filter_button = ctk.CTkButton(
            self.filter_controls_frame, 
            text="Add Filter", 
            command=self.add_filter
        )
        self.add_filter_button.grid(row=0, column=6, padx=10, pady=5)
        
        # Frame for active filters
        self.active_filters_label = ctk.CTkLabel(
            self.filter_frame, 
            text="Active Filters:", 
            font=("Arial", 14, "bold")
        )
        self.active_filters_label.pack(anchor="w", padx=10, pady=(15, 5))
        
        self.active_filters_frame = ctk.CTkScrollableFrame(self.filter_frame, height=150)
        self.active_filters_frame.pack(fill="x", padx=10, pady=5)
        
        # Add placeholder for no filters
        self.no_filters_label = ctk.CTkLabel(
            self.active_filters_frame, 
            text="No filters applied. Data will be analyzed without filtering."
        )
        self.no_filters_label.pack(anchor="w", padx=5, pady=5)
        
        # Filter actions frame
        self.filter_actions_frame = ctk.CTkFrame(self.filter_frame)
        self.filter_actions_frame.pack(fill="x", padx=10, pady=10)
        
        # Apply filters button
        self.apply_filters_button = ctk.CTkButton(
            self.filter_actions_frame, 
            text="Apply Filters", 
            command=self.apply_filters
        )
        self.apply_filters_button.pack(side="left", padx=5, pady=5)
        
        # Reset filters button
        self.reset_filters_button = ctk.CTkButton(
            self.filter_actions_frame, 
            text="Reset Filters", 
            command=self.reset_filters
        )
        self.reset_filters_button.pack(side="left", padx=5, pady=5)
        
        # Data preview after filtering
        self.filtered_data_label = ctk.CTkLabel(
            self.filter_frame, 
            text="Filtered Data Preview:", 
            font=("Arial", 14, "bold")
        )
        self.filtered_data_label.pack(anchor="w", padx=10, pady=(15, 5))
        
        self.filtered_data_info = ctk.CTkLabel(
            self.filter_frame,
            text="No filters applied yet."
        )
        self.filtered_data_info.pack(anchor="w", padx=10, pady=(0, 5))
        
        # Initialize filter variables
        self.active_filters = []
        self.filtered_data = None
        
    def upload_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Determine file type and read accordingly
                if file_path.endswith('.csv'):
                    self.data = pd.read_csv(file_path)
                elif file_path.endswith(('.xlsx', '.xls')):
                    self.data = pd.read_excel(file_path)
                else:
                    messagebox.showerror("Error", "Unsupported file format")
                    return
                
                # Set file path and display file info
                self.file_path = file_path
                file_name = os.path.basename(file_path)
                file_size = os.path.getsize(file_path) / 1024  # size in KB
                
                self.file_info.configure(text=f"File: {file_name}\nSize: {file_size:.1f}KB\nRows: {len(self.data)}\nColumns: {len(self.data.columns)}")
                
                # Update column dropdown
                self.update_column_dropdown()
                
                # Display data preview
                self.display_data_preview()
                
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file: {str(e)}")
    
    def display_data_preview(self):
        # Clear existing widgets
        for widget in self.data_scroll.winfo_children():
            widget.destroy()
        
        if self.data is None:
            return
        
        # Display first 10 rows of data
        preview_data = self.data.head(10)
        
        # Create a table with column headers
        header_frame = ctk.CTkFrame(self.data_scroll)
        header_frame.pack(fill="x", padx=2, pady=2)
        
        # Add column headers
        for i, col in enumerate(preview_data.columns):
            ctk.CTkLabel(
                header_frame, 
                text=str(col), 
                font=("Arial", 12, "bold"),
                width=120,
                corner_radius=4,
                fg_color=("lightgray", "gray30"),
                anchor="center"
            ).grid(row=0, column=i, padx=2, pady=2, sticky="ew")
        
        # Add data rows
        for i, row in enumerate(preview_data.itertuples(index=False)):
            row_frame = ctk.CTkFrame(self.data_scroll)
            row_frame.pack(fill="x", padx=2, pady=1)
            
            for j, value in enumerate(row):
                ctk.CTkLabel(
                    row_frame, 
                    text=str(value)[:20] + ('...' if len(str(value)) > 20 else ''),
                    width=120,
                    corner_radius=0,
                    anchor="w"
                ).grid(row=0, column=j, padx=2, pady=1, sticky="ew")
        
    def analyze_data(self):
        
        if self.data is None:
            messagebox.showinfo("Info", "Please upload a file first.")
            return
        
        # Get report title from the entry
        self.report_title = self.title_entry.get() or "Data Analysis Report"
        
        # Clear temp files from previous analysis
        for file in self.temp_image_files:
            try:
                if os.path.exists(file):
                    os.remove(file)
            except:
                pass
        self.temp_image_files = []
        
        # Clear existing stats
        for widget in self.stats_scroll.winfo_children():
            widget.destroy()
        
        # Clear existing charts
        for widget in self.chart_canvas_frame.winfo_children():
            widget.destroy()
        
        # Generate all required statistics based on selected options
        if self.summary_stats_var.get():
            self.display_summary_statistics()
        
        if self.value_counts_var.get():
            self.display_value_counts()
        
        if self.correlation_var.get():
            self.display_correlation()
        
        # Generate chart based on selected type
        self.generate_chart()
        
    def display_summary_statistics(self):
        # Create a frame for summary stats
        stats_frame = ctk.CTkFrame(self.stats_scroll)
        stats_frame.pack(fill="x", padx=5, pady=5)
        
        # Add a header
        ctk.CTkLabel(
            stats_frame, 
            text="Summary Statistics", 
            font=("Arial", 14, "bold")
        ).pack(anchor="w", padx=5, pady=5)
        
        # Get summary stats
        numeric_data = self.data.select_dtypes(include=[np.number])
        if not len(numeric_data.columns):
            ctk.CTkLabel(
                stats_frame, 
                text="No numeric columns found for summary statistics."
            ).pack(anchor="w", padx=5, pady=5)
            return
            
        stats = numeric_data.describe().round(2)
        
        # Create a table to display stats
        table_frame = ctk.CTkFrame(stats_frame)
        table_frame.pack(fill="x", padx=5, pady=5)
        
        # Add column headers
        ctk.CTkLabel(
            table_frame, 
            text="Statistic", 
            font=("Arial", 12, "bold"),
            width=100,
            corner_radius=4,
            fg_color=("lightgray", "gray30"),
            anchor="center"
        ).grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        for i, col in enumerate(stats.columns):
            ctk.CTkLabel(
                table_frame, 
                text=str(col), 
                font=("Arial", 12, "bold"),
                width=100,
                corner_radius=4,
                fg_color=("lightgray", "gray30"),
                anchor="center"
            ).grid(row=0, column=i+1, padx=2, pady=2, sticky="ew")
        
        # Add stat rows
        for i, stat in enumerate(stats.index):
            ctk.CTkLabel(
                table_frame, 
                text=str(stat), 
                font=("Arial", 12, "bold"),
                width=100,
                corner_radius=0,
                anchor="w"
            ).grid(row=i+1, column=0, padx=2, pady=1, sticky="ew")
            
            for j, col in enumerate(stats.columns):
                ctk.CTkLabel(
                    table_frame, 
                    text=str(stats.loc[stat, col]),
                    width=100,
                    corner_radius=0,
                    anchor="e"
                ).grid(row=i+1, column=j+1, padx=2, pady=1, sticky="ew")
    
    def display_value_counts(self):
        if self.data is None or len(self.data.columns) == 0:
            return
        
        # Create a frame for value counts
        value_counts_frame = ctk.CTkFrame(self.stats_scroll)
        value_counts_frame.pack(fill="x", padx=5, pady=5)
        
        # Add a header
        ctk.CTkLabel(
            value_counts_frame, 
            text="Value Counts (Top 5)", 
            font=("Arial", 14, "bold")
        ).pack(anchor="w", padx=5, pady=5)
        
        # Only process first 5 categorical columns
        cat_columns = self.data.select_dtypes(include=['object', 'category']).columns[:5]
        
        if len(cat_columns) == 0:
            ctk.CTkLabel(
                value_counts_frame, 
                text="No categorical columns found for value counts."
            ).pack(anchor="w", padx=5, pady=5)
            return
        
        # Create tabs for different columns
        columns_tabview = ctk.CTkTabview(value_counts_frame)
        columns_tabview.pack(fill="x", padx=5, pady=5)
        
        for col in cat_columns:
            tab = columns_tabview.add(col)
            
            try:
                value_counts = self.data[col].value_counts().head(10)
                
                # Create a table for the value counts
                table_frame = ctk.CTkFrame(tab)
                table_frame.pack(fill="x", padx=5, pady=5)
                
                # Headers
                ctk.CTkLabel(
                    table_frame, 
                    text="Value", 
                    font=("Arial", 12, "bold"),
                    width=150,
                    corner_radius=4,
                    fg_color=("lightgray", "gray30"),
                    anchor="center"
                ).grid(row=0, column=0, padx=2, pady=2, sticky="ew")
                
                ctk.CTkLabel(
                    table_frame, 
                    text="Count", 
                    font=("Arial", 12, "bold"),
                    width=80,
                    corner_radius=4,
                    fg_color=("lightgray", "gray30"),
                    anchor="center"
                ).grid(row=0, column=1, padx=2, pady=2, sticky="ew")
                
                # Value rows
                for i, (value, count) in enumerate(value_counts.items()):
                    ctk.CTkLabel(
                        table_frame, 
                        text=str(value)[:30] + ('...' if len(str(value)) > 30 else ''),
                        width=150,
                        corner_radius=0,
                        anchor="w"
                    ).grid(row=i+1, column=0, padx=2, pady=1, sticky="ew")
                    
                    ctk.CTkLabel(
                        table_frame, 
                        text=str(count),
                        width=80,
                        corner_radius=0,
                        anchor="e"
                    ).grid(row=i+1, column=1, padx=2, pady=1, sticky="ew")
                    
            except Exception as e:
                ctk.CTkLabel(
                    tab, 
                    text=f"Error processing column: {str(e)}"
                ).pack(anchor="w", padx=5, pady=5)
    
    def display_correlation(self):
        if self.data is None:
            return
        
        # Only process numeric columns
        numeric_data = self.data.select_dtypes(include=[np.number])
        
        if len(numeric_data.columns) < 2:
            corr_frame = ctk.CTkFrame(self.stats_scroll)
            corr_frame.pack(fill="x", padx=5, pady=5)
            
            ctk.CTkLabel(
                corr_frame, 
                text="Correlation Matrix", 
                font=("Arial", 14, "bold")
            ).pack(anchor="w", padx=5, pady=5)
            
            ctk.CTkLabel(
                corr_frame, 
                text="Insufficient numeric columns for correlation analysis."
            ).pack(anchor="w", padx=5, pady=5)
            return
        
        # Create a correlation matrix visualization
        corr_matrix = numeric_data.corr()
        
        # Create the heatmap
        fig, ax = plt.subplots(figsize=(10, 8))
        sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt=".2f", ax=ax)
        ax.set_title('Correlation Matrix')
        
        # Save to file for PDF report
        corr_file = self.add_chart_to_temp_files(fig, "Correlation Matrix")
        
        # Display in the Charts tab
        for widget in self.chart_canvas_frame.winfo_children():
            widget.destroy()
            
        canvas = FigureCanvasTkAgg(fig, master=self.chart_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
    def generate_chart(self):
        if self.data is None:
            return
        
        chart_type = self.chart_type.get()
        selected_column = self.column_var.get()
        
        # Check if selected column exists and is valid
        if selected_column not in self.data.columns or selected_column in ["No Data Loaded", "No Columns Found"]:
            # Fallback to first numeric column
            numeric_cols = self.data.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) == 0:
                messagebox.showinfo("Info", "No numeric columns found for charting.")
                return
            selected_column = numeric_cols[0]
        
        # Check if selected column is numeric for certain chart types
        if chart_type in ["Line Chart", "Histogram"] and not np.issubdtype(self.data[selected_column].dtype, np.number):
            messagebox.showinfo("Info", f"Column '{selected_column}' is not numeric. Please select a numeric column for {chart_type}.")
            return
            
        fig, ax = plt.subplots(figsize=(10, 6))
        
        if chart_type == "Bar Chart":
            if len(self.data) > 25:  # If too many rows, use the first 25
                data_subset = self.data.head(25)
            else:
                data_subset = self.data
                
            data_subset[selected_column].plot(kind='bar', ax=ax)
            ax.set_title(f'Bar Chart of {selected_column}')
            ax.set_ylabel(selected_column)
            ax.set_xlabel('Index')
            
        elif chart_type == "Line Chart":
            if len(self.data) > 100:  # If too many rows, use the first 100
                data_subset = self.data.head(100)
            else:
                data_subset = self.data
                
            data_subset[selected_column].plot(kind='line', ax=ax)
            ax.set_title(f'Line Chart of {selected_column}')
            ax.set_ylabel(selected_column)
            ax.set_xlabel('Index')
            
        elif chart_type == "Pie Chart":
            if np.issubdtype(self.data[selected_column].dtype, np.number):
                # For numeric columns, bin the data
                bins = pd.cut(self.data[selected_column], bins=5)
                value_counts = bins.value_counts()
                value_counts.plot(kind='pie', ax=ax, autopct='%1.1f%%')
            else:
                # For categorical columns, use value counts
                if len(self.data[selected_column].unique()) > 10:
                    # Too many unique values, use top 10
                    top_values = self.data[selected_column].value_counts().head(10)
                    top_values.plot(kind='pie', ax=ax, autopct='%1.1f%%')
                else:
                    self.data[selected_column].value_counts().plot(kind='pie', ax=ax, autopct='%1.1f%%')
                
            ax.set_title(f'Pie Chart of {selected_column}')
            ax.set_ylabel('')
            
        elif chart_type == "Histogram":
            self.data[selected_column].plot(kind='hist', bins=20, ax=ax)
            ax.set_title(f'Histogram of {selected_column}')
            ax.set_ylabel('Frequency')
            ax.set_xlabel(selected_column)
        
        # Save to file for PDF report
        chart_file = self.add_chart_to_temp_files(fig, chart_type, selected_column)
        
        # Display in Charts tab
        for widget in self.chart_canvas_frame.winfo_children():
            widget.destroy()
            
        canvas = FigureCanvasTkAgg(fig, master=self.chart_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def generate_pdf(self):
        if self.data is None:
            messagebox.showinfo("Info", "Please upload and analyze data first.")
            return
        
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Create a PDF options window to customize report
            pdf_options = self.create_pdf_options_window()
            
            # Wait for the window to be closed
            self.wait_window(pdf_options)
            
            # If user canceled, return
            if not hasattr(self, 'pdf_options_result') or not self.pdf_options_result:
                return
                
            # Get selected options
            options = self.pdf_options_result
            
            # Determine which data to use
            report_data = self.data
            if options.get('use_filtered_data', False) and hasattr(self, 'filtered_data') and self.filtered_data is not None:
                report_data = self.filtered_data
            
            # Create PDF
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            elements = []
            
            # Styles
            styles = getSampleStyleSheet()
            title_style = styles['Title']
            heading_style = styles['Heading2']
            normal_style = styles['Normal']
            
            # Title
            elements.append(Paragraph(self.report_title, title_style))
            elements.append(Spacer(1, 12))
            
            # Date & time
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            elements.append(Paragraph(f"Generated on: {now}", normal_style))
            elements.append(Spacer(1, 12))
            
            # Basic file info
            if options.get('include_file_info', True):
                file_name = os.path.basename(self.file_path)
                elements.append(Paragraph("File Information", heading_style))
                elements.append(Paragraph(f"File name: {file_name}", normal_style))
                elements.append(Paragraph(f"Original rows: {len(self.data)}", normal_style))
                elements.append(Paragraph(f"Original columns: {len(self.data.columns)}", normal_style))
                
                # Add filtered data info if applicable
                if options.get('use_filtered_data', False) and hasattr(self, 'filtered_data') and self.filtered_data is not None:
                    elements.append(Paragraph(f"Filtered rows: {len(report_data)}", normal_style))
                    elements.append(Paragraph(f"Rows excluded by filters: {len(self.data) - len(report_data)}", normal_style))
                
                elements.append(Spacer(1, 12))
            
            # Filter details section
            if options.get('include_filter_details', False) and hasattr(self, 'active_filters') and len(self.active_filters) > 0:
                elements.append(Paragraph("Applied Filters", heading_style))
                elements.append(Paragraph(f"Number of filters applied: {len(self.active_filters)}", normal_style))
                
                # Create filter details table
                filter_data = [["Column", "Operator", "Value"]]
                for filter_obj in self.active_filters:
                    filter_data.append([
                        filter_obj['column'],
                        filter_obj['operator'],
                        filter_obj['value']
                    ])
                
                filter_table = Table(filter_data, colWidths=[150, 100, 150])
                filter_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                elements.append(filter_table)
                elements.append(Spacer(1, 12))
            
            # Add comments if provided
            comment = self.comment_text.get("1.0", "end-1c").strip()
            if comment and options.get('include_comments', True):
                elements.append(Paragraph("Comments", heading_style))
                elements.append(Paragraph(comment, normal_style))
                elements.append(Spacer(1, 12))
            
            # Summary statistics using report_data (filtered or original)
            if options.get('include_summary_stats', True) and self.summary_stats_var.get():
                numeric_data = report_data.select_dtypes(include=[np.number])
                if len(numeric_data.columns) > 0:
                    elements.append(Paragraph("Summary Statistics", heading_style))
                    if options.get('use_filtered_data', False) and len(report_data) != len(self.data):
                        elements.append(Paragraph("(Based on filtered data)", normal_style))
                    
                    stats = numeric_data.describe().round(2).reset_index()
                    
                    # Convert to table data
                    table_data = [["Statistic"] + list(stats.columns[1:])]
                    for _, row in stats.iterrows():
                        table_data.append([row[0]] + [str(x) for x in row[1:]])
                    
                    # Create table
                    t = Table(table_data)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ]))
                    elements.append(t)
                    elements.append(Spacer(1, 12))
            
            # Value counts using report_data (filtered or original)
            if options.get('include_value_counts', True) and self.value_counts_var.get():
                cat_columns = report_data.select_dtypes(include=['object', 'category']).columns[:3]
                
                if len(cat_columns) > 0:
                    elements.append(Paragraph("Value Counts (Top 5)", heading_style))
                    if options.get('use_filtered_data', False) and len(report_data) != len(self.data):
                        elements.append(Paragraph("(Based on filtered data)", normal_style))
                    
                    for col in cat_columns:
                        elements.append(Paragraph(f"Column: {col}", styles['Heading3']))
                        
                        try:
                            value_counts = report_data[col].value_counts().head(5)
                            
                            # Convert to table data
                            table_data = [["Value", "Count"]]
                            for value, count in value_counts.items():
                                table_data.append([str(value)[:30] + ('...' if len(str(value)) > 30 else ''), str(count)])
                            
                            # Create table
                            t = Table(table_data, colWidths=[300, 100])
                            t.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                            ]))
                            elements.append(t)
                            elements.append(Spacer(1, 12))
                            
                        except Exception as e:
                            elements.append(Paragraph(f"Error processing column: {str(e)}", normal_style))
                            elements.append(Spacer(1, 12))
            
            # Add charts based on user selection
            selected_charts = options.get('selected_charts', [])
            for img_file in selected_charts:
                if os.path.exists(img_file):
                    chart_name = os.path.basename(img_file).replace('_', ' ').replace('.png', '')
                    elements.append(Paragraph(f"Visualization: {chart_name.title()}", heading_style))
                    elements.append(Image(img_file, width=6*inch, height=4*inch))
                    elements.append(Spacer(1, 12))
            
            # Generate PDF
            doc.build(elements)
            
            messagebox.showinfo("Success", f"PDF report saved to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF report: {str(e)}")
        
    def create_pdf_options_window(self):
        """Create a window to customize PDF report options"""
        options_window = ctk.CTkToplevel(self)
        options_window.title("PDF Report Options")
        options_window.geometry("500x700")  # Increased height for filter options
        options_window.resizable(True, True)
        options_window.grab_set()
        
        # Create main container
        main_container = ctk.CTkFrame(options_window)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create scrollable frame for content
        scroll_frame = ctk.CTkScrollableFrame(main_container)
        scroll_frame.pack(fill="both", expand=True, padx=5, pady=(5, 60))
        
        # Add title
        title_label = ctk.CTkLabel(scroll_frame, text="Customize PDF Report", font=("Arial", 16, "bold"))
        title_label.pack(anchor="w", padx=10, pady=10)
        
        # Basic options
        basic_frame = ctk.CTkFrame(scroll_frame)
        basic_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(basic_frame, text="Basic Options", font=("Arial", 14, "bold")).pack(anchor="w", padx=5, pady=5)
        
        # Checkboxes for basic options
        include_file_info_var = ctk.BooleanVar(value=True)
        include_file_info_cb = ctk.CTkCheckBox(basic_frame, text="Include File Information", variable=include_file_info_var)
        include_file_info_cb.pack(anchor="w", padx=20, pady=2)
        
        include_comments_var = ctk.BooleanVar(value=True)
        include_comments_cb = ctk.CTkCheckBox(basic_frame, text="Include Comments", variable=include_comments_var)
        include_comments_cb.pack(anchor="w", padx=20, pady=2)
        
        include_summary_stats_var = ctk.BooleanVar(value=True)
        include_summary_stats_cb = ctk.CTkCheckBox(basic_frame, text="Include Summary Statistics", variable=include_summary_stats_var)
        include_summary_stats_cb.pack(anchor="w", padx=20, pady=2)
        
        include_value_counts_var = ctk.BooleanVar(value=True)
        include_value_counts_cb = ctk.CTkCheckBox(basic_frame, text="Include Value Counts", variable=include_value_counts_var)
        include_value_counts_cb.pack(anchor="w", padx=20, pady=2)
        
        # Filter options - NEW SECTION
        filter_frame = ctk.CTkFrame(scroll_frame)
        filter_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(filter_frame, text="Filter Data Options", font=("Arial", 14, "bold")).pack(anchor="w", padx=5, pady=5)
        
        # Check if filters are available
        has_filters = hasattr(self, 'active_filters') and len(self.active_filters) > 0
        has_filtered_data = hasattr(self, 'filtered_data') and self.filtered_data is not None
        
        if has_filters:
            # Show active filters info
            filter_info_text = f"Active Filters ({len(self.active_filters)}):\n"
            for f in self.active_filters:
                filter_info_text += f"• {f['column']} {f['operator']} '{f['value']}'\n"
            
            filter_info_label = ctk.CTkLabel(
                filter_frame, 
                text=filter_info_text,
                justify="left",
                wraplength=400
            )
            filter_info_label.pack(anchor="w", padx=20, pady=5)
            
            # Option to use filtered data
            use_filtered_data_var = ctk.BooleanVar(value=True)
            use_filtered_data_cb = ctk.CTkCheckBox(
                filter_frame, 
                text="Use Filtered Data for Analysis", 
                variable=use_filtered_data_var
            )
            use_filtered_data_cb.pack(anchor="w", padx=20, pady=2)
            
            # Option to include filter details in report
            include_filter_details_var = ctk.BooleanVar(value=True)
            include_filter_details_cb = ctk.CTkCheckBox(
                filter_frame, 
                text="Include Filter Details in Report", 
                variable=include_filter_details_var
            )
            include_filter_details_cb.pack(anchor="w", padx=20, pady=2)
            
            # Show filtered data stats
            if has_filtered_data:
                filter_stats_text = f"Filtered Data: {len(self.filtered_data)} rows (from {len(self.data)} original rows)"
                ctk.CTkLabel(filter_frame, text=filter_stats_text).pack(anchor="w", padx=20, pady=2)
        else:
            ctk.CTkLabel(
                filter_frame, 
                text="No filters applied. All data will be used in the report."
            ).pack(anchor="w", padx=20, pady=5)
            use_filtered_data_var = ctk.BooleanVar(value=False)
            include_filter_details_var = ctk.BooleanVar(value=False)
        
        # Chart selection
        chart_frame = ctk.CTkFrame(scroll_frame)
        chart_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(chart_frame, text="Select Charts to Include", font=("Arial", 14, "bold")).pack(anchor="w", padx=5, pady=5)
        
        # Create a fixed height frame for chart selection
        chart_scroll = ctk.CTkScrollableFrame(chart_frame, height=150)
        chart_scroll.pack(fill="x", padx=5, pady=5)
        
        # Add checkboxes for each chart
        chart_vars = {}
        if hasattr(self, 'temp_image_files') and self.temp_image_files:
            for img_file in self.temp_image_files:
                if os.path.exists(img_file):
                    chart_name = os.path.basename(img_file).replace('_', ' ').replace('.png', '')
                    var = ctk.BooleanVar(value=True)
                    chart_vars[img_file] = var
                    ctk.CTkCheckBox(chart_scroll, text=chart_name.title(), variable=var).pack(anchor="w", padx=20, pady=2)
        else:
            ctk.CTkLabel(chart_scroll, text="No charts available. Please analyze data first.").pack(anchor="w", padx=20, pady=2)
        
        # Fixed buttons frame at the bottom
        buttons_frame = ctk.CTkFrame(main_container)
        buttons_frame.pack(fill="x", side="bottom", padx=5, pady=5)
        
        # Define what happens when user confirms or cancels
        def on_confirm():
            options = {
                'include_file_info': include_file_info_var.get(),
                'include_comments': include_comments_var.get(),
                'include_summary_stats': include_summary_stats_var.get(),
                'include_value_counts': include_value_counts_var.get(),
                'use_filtered_data': use_filtered_data_var.get() if has_filters else False,
                'include_filter_details': include_filter_details_var.get() if has_filters else False,
                'selected_charts': [file for file, var in chart_vars.items() if var.get()]
            }
            self.pdf_options_result = options
            options_window.destroy()
        
        def on_cancel():
            self.pdf_options_result = None
            options_window.destroy()
        
        # Confirm and cancel buttons
        cancel_button = ctk.CTkButton(buttons_frame, text="Cancel", command=on_cancel, width=100)
        cancel_button.pack(side="left", padx=5, pady=5)
        
        confirm_button = ctk.CTkButton(buttons_frame, text="Generate PDF", command=on_confirm, fg_color="#2a9d8f", width=120)
        confirm_button.pack(side="right", padx=5, pady=5)
        
        # Add a status label
        status_label = ctk.CTkLabel(buttons_frame, text="Configure your report options above, then click 'Generate PDF'")
        status_label.pack(side="left", padx=20, pady=5)
        
        return options_window
    
    def add_chart_to_temp_files(self, fig, chart_type, column_name=None):
        """Add a chart to temp files with a more descriptive name"""
        # Clean up chart type and column name for filename
        safe_chart_type = ''.join(c if c.isalnum() else '_' for c in chart_type.lower())
        
        if column_name:
            safe_column_name = ''.join(c if c.isalnum() else '_' for c in column_name)
            filename = f"{safe_chart_type}_{safe_column_name}.png"
        else:
            filename = f"{safe_chart_type}.png"
        
        # Use a counter to ensure unique filenames instead of timestamps
        counter = 1
        base_filename = filename
        chart_file = os.path.join(tempfile.gettempdir(), filename)
        
        while chart_file in self.temp_image_files or os.path.exists(chart_file):
            name_part, ext = os.path.splitext(base_filename)
            filename = f"{name_part}_{counter}{ext}"
            chart_file = os.path.join(tempfile.gettempdir(), filename)
            counter += 1
        
        fig.savefig(chart_file, bbox_inches='tight', dpi=100)
        self.temp_image_files.append(chart_file)
        return chart_file
        
    def update_column_dropdown(self):
        """Update the column dropdown with available columns from the data"""
        if self.data is None:
            self.column_dropdown.configure(values=["No Data Loaded"])
            self.filter_column_dropdown.configure(values=["No Data Loaded"])
            return
            
        # Get column names
        columns = list(self.data.columns)
        if not columns:
            self.column_dropdown.configure(values=["No Columns Found"])
            self.filter_column_dropdown.configure(values=["No Columns Found"])
            return
            
        # Update dropdowns
        self.column_dropdown.configure(values=columns)
        self.column_dropdown.set(columns[0])  # Set first column as default
        
        self.filter_column_dropdown.configure(values=columns)
        self.filter_column_dropdown.set(columns[0])  # Set first column as default
    
    def export_data(self):
        """Export the analyzed data to a CSV file"""
        if self.data is None:
            messagebox.showinfo("Info", "Please upload data first.")
            return
            
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            title="Export Data",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            # Export based on file extension
            if file_path.endswith('.csv'):
                self.data.to_csv(file_path, index=False)
            elif file_path.endswith(('.xlsx', '.xls')):
                self.data.to_excel(file_path, index=False)
            else:
                # Default to CSV
                self.data.to_csv(file_path, index=False)
                
            messagebox.showinfo("Success", f"Data exported to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def add_filter(self):
        """Add a filter to the active filters list"""
        if self.data is None:
            messagebox.showinfo("Info", "Please upload data first.")
            return
        
        column = self.filter_column_var.get()
        operator = self.filter_operator_var.get()
        value = self.filter_value_entry.get()
        
        if not value:
            messagebox.showinfo("Info", "Please enter a filter value.")
            return
        
        # Create a unique ID for this filter
        filter_id = len(self.active_filters)
        
        # Create a filter object
        filter_obj = {
            'id': filter_id,
            'column': column,
            'operator': operator,
            'value': value
        }
        
        # Add to active filters
        self.active_filters.append(filter_obj)
        
        # Remove placeholder if this is the first filter
        if len(self.active_filters) == 1:
            if hasattr(self, 'no_filters_label') and self.no_filters_label.winfo_exists():
                self.no_filters_label.destroy()
        
        # Create a filter display item
        filter_item_frame = ctk.CTkFrame(self.active_filters_frame)
        filter_item_frame.pack(fill="x", padx=5, pady=5)
        
        # Display the filter details
        filter_text = f"{column} {operator} '{value}'"
        filter_label = ctk.CTkLabel(filter_item_frame, text=filter_text)
        filter_label.pack(side="left", padx=5, pady=5)
        
        # Add a remove button
        remove_button = ctk.CTkButton(
            filter_item_frame,
            text="Remove",
            width=80,
            command=lambda fid=filter_id: self.remove_filter(fid, filter_item_frame)
        )
        remove_button.pack(side="right", padx=5, pady=5)
        
        # Clear the value entry
        self.filter_value_entry.delete(0, "end")

    def remove_filter(self, filter_id, filter_frame):
        """Remove a filter from the active filters list"""
        # Remove the filter item from the UI
        filter_frame.destroy()
        
        # Remove the filter from the list
        self.active_filters = [f for f in self.active_filters if f['id'] != filter_id]
        
        # If no filters remain, add the placeholder back
        if len(self.active_filters) == 0:
            self.no_filters_label = ctk.CTkLabel(
                self.active_filters_frame, 
                text="No filters applied. Data will be analyzed without filtering."
            )
            self.no_filters_label.pack(anchor="w", padx=5, pady=5)

    def apply_filters(self):
        """Apply all active filters to the data"""
        if self.data is None:
            messagebox.showinfo("Info", "Please upload data first.")
            return
        
        if not self.active_filters:
            messagebox.showinfo("Info", "No filters to apply.")
            return
        
        try:
            # Start with a copy of the original data
            filtered_data = self.data.copy()
            
            # Apply each filter
            for filter_obj in self.active_filters:
                column = filter_obj['column']
                operator = filter_obj['operator']
                value = filter_obj['value']
                
                # Apply the filter based on operator
                if operator == "equals":
                    filtered_data = filtered_data[filtered_data[column].astype(str) == value]
                elif operator == "not equals":
                    filtered_data = filtered_data[filtered_data[column].astype(str) != value]
                elif operator == "greater than":
                    # Try to convert to numeric if possible
                    try:
                        filtered_data = filtered_data[filtered_data[column].astype(float) > float(value)]
                    except:
                        messagebox.showerror("Error", f"Cannot apply '{operator}' to non-numeric data.")
                        return
                elif operator == "less than":
                    # Try to convert to numeric if possible
                    try:
                        filtered_data = filtered_data[filtered_data[column].astype(float) < float(value)]
                    except:
                        messagebox.showerror("Error", f"Cannot apply '{operator}' to non-numeric data.")
                        return
                elif operator == "contains":
                    filtered_data = filtered_data[filtered_data[column].astype(str).str.contains(value, na=False)]
                elif operator == "starts with":
                    filtered_data = filtered_data[filtered_data[column].astype(str).str.startswith(value, na=False)]
                elif operator == "ends with":
                    filtered_data = filtered_data[filtered_data[column].astype(str).str.endswith(value, na=False)]
            
            # Store the filtered data
            self.filtered_data = filtered_data
            
            # Update the info label
            self.filtered_data_info.configure(
                text=f"Filtered data: {len(filtered_data)} rows (from original {len(self.data)} rows)"
            )
            
            # Display preview of filtered data
            self.display_filtered_data_preview()
            
            # Alert success
            messagebox.showinfo("Success", f"Filters applied. {len(filtered_data)} rows match your criteria.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error applying filters: {str(e)}")

    def reset_filters(self):
        """Reset all filters"""
        # Clear active filters list
        self.active_filters = []
        
        # Reset filtered data
        self.filtered_data = None
        
        # Clear the filter frames
        for widget in self.active_filters_frame.winfo_children():
            widget.destroy()
        
        # Add the placeholder back
        self.no_filters_label = ctk.CTkLabel(
            self.active_filters_frame, 
            text="No filters applied. Data will be analyzed without filtering."
        )
        self.no_filters_label.pack(anchor="w", padx=5, pady=5)
        
        # Update the info label
        self.filtered_data_info.configure(text="No filters applied yet.")

    def display_filtered_data_preview(self):
        """Display a preview of the filtered data"""
        if self.filtered_data is None or len(self.filtered_data) == 0:
            return
        
        # Create a new window for the preview
        preview_window = ctk.CTkToplevel(self)
        preview_window.title("Filtered Data Preview")
        preview_window.geometry("800x500")
        
        # Create a scrollable frame for the data
        preview_scroll = ctk.CTkScrollableFrame(preview_window)
        preview_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create header frame
        header_frame = ctk.CTkFrame(preview_scroll)
        header_frame.pack(fill="x", padx=2, pady=2)
        
        # Display first 20 rows of filtered data
        preview_data = self.filtered_data.head(20)
        
        # Add column headers
        for i, col in enumerate(preview_data.columns):
            ctk.CTkLabel(
                header_frame, 
                text=str(col), 
                font=("Arial", 12, "bold"),
                width=120,
                corner_radius=4,
                fg_color=("lightgray", "gray30"),
                anchor="center"
            ).grid(row=0, column=i, padx=2, pady=2, sticky="ew")
        
        # Add data rows
        for i, row in enumerate(preview_data.itertuples(index=False)):
            row_frame = ctk.CTkFrame(preview_scroll)
            row_frame.pack(fill="x", padx=2, pady=1)
            
            for j, value in enumerate(row):
                ctk.CTkLabel(
                    row_frame, 
                    text=str(value)[:20] + ('...' if len(str(value)) > 20 else ''),
                    width=120,
                    corner_radius=0,
                    anchor="w"
                ).grid(row=0, column=j, padx=2, pady=1, sticky="ew")        

    def reset_app(self):
        # Reset variables
        self.data = None
        self.file_path = None
        self.report_title = "Data Analysis Report"
        
        # Clear temp files
        for file in self.temp_image_files:
            try:
                if os.path.exists(file):
                    os.remove(file)
            except:
                pass
        self.temp_image_files = []
        
        # Reset UI
        self.file_info.configure(text="No file selected")
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, self.report_title)
        self.comment_text.delete("1.0", "end")
        
        # Reset column dropdown
        self.column_dropdown.configure(values=["No Data Loaded"])
        self.column_dropdown.set("No Data Loaded")
        
        # Clear preview areas
        for widget in self.data_scroll.winfo_children():
            widget.destroy()
            
        for widget in self.stats_scroll.winfo_children():
            widget.destroy()
            
        for widget in self.chart_canvas_frame.winfo_children():
            widget.destroy()
    
    def change_appearance_mode(self, new_appearance_mode):
        ctk.set_appearance_mode(new_appearance_mode)

if __name__ == "__main__":
    app = DataAnalyzerApp()
    app.mainloop()
