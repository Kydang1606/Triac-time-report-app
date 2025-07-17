import pandas as pd
import datetime
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import app_logic # Import các hàm logic từ file app_logic.py

# Để bỏ qua cảnh báo UserWarning: Data Validation extension is not supported
# import openpyxl.worksheet._read_only as read_only
# if hasattr(read_only, 'ColumnDimension'):
#     del read_only.ColumnDimension
# if hasattr(read_only, 'RowDimension'):
#     del read_only.RowDimension

class TimeReportApp:
    def __init__(self, master):
        self.master = master
        master.title("TRIAC Time Report Generator")
        master.geometry("800x700")

        self.paths = core_logic.setup_paths()
        self.template_file = self.paths['template_file']

        self.df_raw = pd.DataFrame()
        self.current_config = {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame()}
        self.comparison_config = {'years': [], 'months': [], 'selected_projects': []}

        self.create_widgets()
        self.load_initial_data()

    def create_widgets(self):
        # Notebook for tabs
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(pady=10, expand=True, fill='both')

        # Tab 1: Standard Report
        self.standard_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.standard_frame, text='Standard Report')
        self.create_standard_report_tab(self.standard_frame)

        # Tab 2: Comparison Report
        self.comparison_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.comparison_frame, text='Comparison Report')
        self.create_comparison_report_tab(self.comparison_frame)

    def create_standard_report_tab(self, parent_frame):
        # Frame for file selection
        file_frame = ttk.LabelFrame(parent_frame, text="File Selection", padding="10")
        file_frame.pack(padx=10, pady=5, fill='x')

        ttk.Label(file_frame, text="Template File:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.template_path_entry = ttk.Entry(file_frame, width=50)
        self.template_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.template_path_entry.insert(0, self.template_file)
        self.browse_button = ttk.Button(file_frame, text="Browse", command=self.browse_template)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)
        self.load_data_button = ttk.Button(file_frame, text="Load Data", command=self.load_initial_data)
        self.load_data_button.grid(row=0, column=3, padx=5, pady=5)
        file_frame.grid_columnconfigure(1, weight=1)

        # Frame for Configuration
        config_frame = ttk.LabelFrame(parent_frame, text="Configuration (from Config_Year_Mode sheet)", padding="10")
        config_frame.pack(padx=10, pady=5, fill='x')

        ttk.Label(config_frame, text="Mode:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.mode_label = ttk.Label(config_frame, text="")
        self.mode_label.grid(row=0, column=1, padx=5, pady=5, sticky='w')

        ttk.Label(config_frame, text="Year:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.year_label = ttk.Label(config_frame, text="")
        self.year_label.grid(row=1, column=1, padx=5, pady=5, sticky='w')

        ttk.Label(config_frame, text="Months:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.months_label = ttk.Label(config_frame, text="")
        self.months_label.grid(row=2, column=1, padx=5, pady=5, sticky='w')

        # Frame for Project Filter
        project_filter_frame = ttk.LabelFrame(parent_frame, text="Project Filter (from Config_Project_Filter sheet)", padding="10")
        project_filter_frame.pack(padx=10, pady=5, fill='both', expand=True)

        self.project_filter_listbox = tk.Listbox(project_filter_frame, selectmode=tk.MULTIPLE, height=10)
        self.project_filter_listbox.pack(padx=5, pady=5, fill='both', expand=True)
        self.project_filter_listbox.bind('<<ListboxSelect>>', self.update_selected_projects_display)

        self.selected_projects_display_label = ttk.Label(project_filter_frame, text="Selected Projects: None")
        self.selected_projects_display_label.pack(padx=5, pady=5, fill='x')

        # Actions Frame
        action_frame = ttk.LabelFrame(parent_frame, text="Actions", padding="10")
        action_frame.pack(padx=10, pady=10, fill='x')

        self.generate_excel_button = ttk.Button(action_frame, text="Generate Excel Report", command=self.generate_standard_excel_report)
        self.generate_excel_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.generate_pdf_button = ttk.Button(action_frame, text="Generate PDF Report", command=self.generate_standard_pdf_report)
        self.generate_pdf_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.status_label = ttk.Label(parent_frame, text="Status: Ready")
        self.status_label.pack(padx=10, pady=5, fill='x')

    def create_comparison_report_tab(self, parent_frame):
        # Frame for comparison mode selection
        comparison_mode_frame = ttk.LabelFrame(parent_frame, text="Comparison Mode", padding="10")
        comparison_mode_frame.pack(padx=10, pady=5, fill='x')

        self.comparison_mode_var = tk.StringVar()
        self.comparison_mode_combobox = ttk.Combobox(comparison_mode_frame, textvariable=self.comparison_mode_var,
                                                     values=["So Sánh Dự Án Trong Một Tháng",
                                                             "So Sánh Dự Án Trong Một Năm",
                                                             "So Sánh Một Dự Án Qua Các Tháng/Năm"])
        self.comparison_mode_combobox.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
        self.comparison_mode_combobox.set("So Sánh Dự Án Trong Một Tháng") # Default
        self.comparison_mode_combobox.bind("<<ComboboxSelected>>", self.on_comparison_mode_change)
        comparison_mode_frame.grid_columnconfigure(0, weight=1)


        # Frame for comparison filters
        comparison_filter_frame = ttk.LabelFrame(parent_frame, text="Comparison Filters", padding="10")
        comparison_filter_frame.pack(padx=10, pady=5, fill='both', expand=True)

        # Years
        ttk.Label(comparison_filter_frame, text="Select Year(s):").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.comp_year_listbox = tk.Listbox(comparison_filter_frame, selectmode=tk.MULTIPLE, height=5)
        self.comp_year_listbox.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        self.comp_year_listbox.bind('<<ListboxSelect>>', self.update_comparison_config)
        
        # Months
        ttk.Label(comparison_filter_frame, text="Select Month(s):").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.comp_month_listbox = tk.Listbox(comparison_filter_frame, selectmode=tk.MULTIPLE, height=5)
        self.comp_month_listbox.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        self.comp_month_listbox.bind('<<ListboxSelect>>', self.update_comparison_config)

        # Projects
        ttk.Label(comparison_filter_frame, text="Select Project(s):").grid(row=4, column=0, padx=5, pady=5, sticky='w')
        self.comp_project_listbox = tk.Listbox(comparison_filter_frame, selectmode=tk.MULTIPLE, height=10)
        self.comp_project_listbox.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        self.comp_project_listbox.bind('<<ListboxSelect>>', self.update_comparison_config)

        comparison_filter_frame.grid_columnconfigure(1, weight=1) # Make entry widgets expand

        # Display selected comparison filters
        self.selected_comp_years_label = ttk.Label(comparison_filter_frame, text="Selected Years: None")
        self.selected_comp_years_label.grid(row=6, column=0, columnspan=2, padx=5, pady=2, sticky='w')
        self.selected_comp_months_label = ttk.Label(comparison_filter_frame, text="Selected Months: None")
        self.selected_comp_months_label.grid(row=7, column=0, columnspan=2, padx=5, pady=2, sticky='w')
        self.selected_comp_projects_label = ttk.Label(comparison_filter_frame, text="Selected Projects: None")
        self.selected_comp_projects_label.grid(row=8, column=0, columnspan=2, padx=5, pady=2, sticky='w')


        # Actions Frame for comparison
        comp_action_frame = ttk.LabelFrame(parent_frame, text="Comparison Actions", padding="10")
        comp_action_frame.pack(padx=10, pady=10, fill='x')

        self.generate_comp_excel_button = ttk.Button(comp_action_frame, text="Generate Comparison Excel", command=self.generate_comparison_excel_report)
        self.generate_comp_excel_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.generate_comp_pdf_button = ttk.Button(comp_action_frame, text="Generate Comparison PDF", command=self.generate_comparison_pdf_report)
        self.generate_comp_pdf_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.comp_status_label = ttk.Label(parent_frame, text="Status: Ready")
        self.comp_status_label.pack(padx=10, pady=5, fill='x')

    def browse_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm *.xlsx")])
        if file_path:
            self.template_path_entry.delete(0, tk.END)
            self.template_path_entry.insert(0, file_path)
            self.template_file = file_path
            self.load_initial_data()

    def load_initial_data(self):
        self.status_label.config(text="Status: Loading data...")
        self.comp_status_label.config(text="Status: Loading data...")
        self.master.update_idletasks()

        template_file_path = self.template_path_entry.get()
        if not os.path.exists(template_file_path):
            messagebox.showerror("Error", f"Template file not found at: {template_file_path}")
            self.status_label.config(text="Status: Error loading data")
            self.comp_status_label.config(text="Status: Error loading data")
            return

        try:
            self.current_config = app_logic.read_configs(template_file_path)
            self.df_raw = app_logic.load_raw_data(template_file_path)

            if self.df_raw.empty:
                messagebox.showwarning("Warning", "Raw Data sheet is empty or could not be loaded.")
                self.status_label.config(text="Status: Data loaded with warnings")
                self.comp_status_label.config(text="Status: Data loaded with warnings")
                return

            self.update_standard_config_display()
            self.populate_comparison_filters()
            self.status_label.config(text="Status: Data loaded successfully")
            self.comp_status_label.config(text="Status: Data loaded successfully")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading data: {e}")
            self.status_label.config(text="Status: Error loading data")
            self.comp_status_label.config(text="Status: Error loading data")

    def update_standard_config_display(self):
        # Update Standard Report Tab
        self.mode_label.config(text=self.current_config.get('mode', 'N/A').capitalize())
        self.year_label.config(text=str(self.current_config.get('year', 'N/A')))
        self.months_label.config(text=', '.join(self.current_config.get('months', [])) if self.current_config.get('months') else "All")

        self.project_filter_listbox.delete(0, tk.END)
        if not self.current_config['project_filter_df'].empty:
            for index, row in self.current_config['project_filter_df'].iterrows():
                project_name = row['Project Name']
                include = str(row.get('Include', 'no')).lower()
                self.project_filter_listbox.insert(tk.END, project_name)
                if include == 'yes':
                    self.project_filter_listbox.selection_set(index)
            self.update_selected_projects_display()
        else:
            self.selected_projects_display_label.config(text="Selected Projects: No projects in filter sheet")
        
    def update_selected_projects_display(self, event=None):
        selected_indices = self.project_filter_listbox.curselection()
        selected_projects = [self.project_filter_listbox.get(i) for i in selected_indices]
        if selected_projects:
            self.selected_projects_display_label.config(text=f"Selected Projects: {', '.join(selected_projects)}")
        else:
            self.selected_projects_display_label.config(text="Selected Projects: None")

    def populate_comparison_filters(self):
        # Populate Year Listbox
        self.comp_year_listbox.delete(0, tk.END)
        if not self.df_raw.empty:
            years = sorted(self.df_raw['Year'].unique().tolist())
            for year in years:
                self.comp_year_listbox.insert(tk.END, year)

        # Populate Month Listbox
        self.comp_month_listbox.delete(0, tk.END)
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        for month in month_order:
            self.comp_month_listbox.insert(tk.END, month)

        # Populate Project Listbox
        self.comp_project_listbox.delete(0, tk.END)
        if not self.df_raw.empty:
            projects = sorted(self.df_raw['Project name'].unique().tolist())
            for project in projects:
                self.comp_project_listbox.insert(tk.END, project)

        self.update_comparison_config() # Initial update

    def update_comparison_config(self, event=None):
        # Update years
        selected_year_indices = self.comp_year_listbox.curselection()
        self.comparison_config['years'] = [self.comp_year_listbox.get(i) for i in selected_year_indices]
        self.selected_comp_years_label.config(text=f"Selected Years: {', '.join(map(str, self.comparison_config['years'])) if self.comparison_config['years'] else 'None'}")

        # Update months
        selected_month_indices = self.comp_month_listbox.curselection()
        self.comparison_config['months'] = [self.comp_month_listbox.get(i) for i in selected_month_indices]
        self.selected_comp_months_label.config(text=f"Selected Months: {', '.join(self.comparison_config['months']) if self.comparison_config['months'] else 'None'}")

        # Update projects
        selected_project_indices = self.comp_project_listbox.curselection()
        self.comparison_config['selected_projects'] = [self.comp_project_listbox.get(i) for i in selected_project_indices]
        self.selected_comp_projects_label.config(text=f"Selected Projects: {', '.join(self.comparison_config['selected_projects']) if self.comparison_config['selected_projects'] else 'None'}")
        
        # Enforce single project selection for specific comparison mode
        if self.comparison_mode_var.get() == "So Sánh Một Dự Án Qua Các Tháng/Năm":
            if len(self.comparison_config['selected_projects']) > 1:
                messagebox.showwarning("Selection Warning", "For 'So Sánh Một Dự Án Qua Các Tháng/Năm' mode, please select only ONE project.")
                # Optionally deselect all but the first one, or clear all
                if selected_project_indices:
                    self.comp_project_listbox.selection_clear(0, tk.END)
                    self.comp_project_listbox.selection_set(selected_project_indices[0])
                    self.comparison_config['selected_projects'] = [self.comp_project_listbox.get(selected_project_indices[0])]
                    self.selected_comp_projects_label.config(text=f"Selected Projects: {self.comparison_config['selected_projects'][0]}")


    def on_comparison_mode_change(self, event=None):
        # Logic to enable/disable listbox selection modes based on the chosen comparison mode
        selected_mode = self.comparison_mode_var.get()
        if selected_mode == "So Sánh Dự Án Trong Một Tháng":
            self.comp_year_listbox.config(selectmode=tk.SINGLE)
            self.comp_month_listbox.config(selectmode=tk.SINGLE)
            self.comp_project_listbox.config(selectmode=tk.MULTIPLE)
        elif selected_mode == "So Sánh Dự Án Trong Một Năm":
            self.comp_year_listbox.config(selectmode=tk.SINGLE)
            self.comp_month_listbox.config(selectmode=tk.MULTIPLE) # Allow multiple months
            self.comp_project_listbox.config(selectmode=tk.MULTIPLE)
        elif selected_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
            self.comp_year_listbox.config(selectmode=tk.MULTIPLE)
            self.comp_month_listbox.config(selectmode=tk.MULTIPLE)
            self.comp_project_listbox.config(selectmode=tk.SINGLE) # Only one project

        self.update_comparison_config() # Re-evaluate selections based on new mode

    def generate_standard_excel_report(self):
        if self.df_raw.empty:
            messagebox.showwarning("Warning", "No raw data loaded. Please load data first.")
            return

        self.status_label.config(text="Status: Generating Excel report...")
        self.master.update_idletasks()

        # Re-read selected projects from the listbox for the standard report
        selected_indices = self.project_filter_listbox.curselection()
        selected_projects_from_gui = [self.project_filter_listbox.get(i) for i in selected_indices]
        
        # Create a temporary project_filter_df for the current run
        temp_project_filter_df = pd.DataFrame({'Project Name': selected_projects_from_gui, 'Include': ['yes'] * len(selected_projects_from_gui)})
        
        # Update current_config with the selected projects from the GUI
        # This ensures that only projects explicitly selected by the user are used
        current_run_config = self.current_config.copy()
        current_run_config['project_filter_df'] = temp_project_filter_df
        
        # Apply filters
        df_filtered = app_logic.apply_filters(self.df_raw, current_run_config)

        if df_filtered.empty:
            messagebox.showwarning("Warning", "No data after applying filters. Report not generated.")
            self.status_label.config(text="Status: Report generation failed (no data)")
            return

        output_file = self.paths['output_file']
        success = app_logic.export_report(df_filtered, current_run_config, output_file)

        if success:
            messagebox.showinfo("Success", f"Standard Excel report generated successfully at:\n{os.path.abspath(output_file)}")
            self.status_label.config(text="Status: Excel report generated.")
            os.startfile(os.path.abspath(output_file))
        else:
            messagebox.showerror("Error", "Failed to generate standard Excel report.")
            self.status_label.config(text="Status: Excel report generation failed.")

    def generate_standard_pdf_report(self):
        if self.df_raw.empty:
            messagebox.showwarning("Warning", "No raw data loaded. Please load data first.")
            return
        
        self.status_label.config(text="Status: Generating PDF report...")
        self.master.update_idletasks()

        # Re-read selected projects from the listbox for the standard report
        selected_indices = self.project_filter_listbox.curselection()
        selected_projects_from_gui = [self.project_filter_listbox.get(i) for i in selected_indices]
        
        # Create a temporary project_filter_df for the current run
        temp_project_filter_df = pd.DataFrame({'Project Name': selected_projects_from_gui, 'Include': ['yes'] * len(selected_projects_from_gui)})
        
        # Update current_config with the selected projects from the GUI
        current_run_config = self.current_config.copy()
        current_run_config['project_filter_df'] = temp_project_filter_df

        df_filtered = app_logic.apply_filters(self.df_raw, current_run_config)

        if df_filtered.empty:
            messagebox.showwarning("Warning", "No data after applying filters. PDF report not generated.")
            self.status_label.config(text="Status: PDF report generation failed (no data)")
            return
            
        pdf_report_path = self.paths['pdf_report']
        logo_path = self.paths['logo_path']
        success = app_logic.export_pdf_report(df_filtered, current_run_config, pdf_report_path, logo_path)

        if success:
            messagebox.showinfo("Success", f"Standard PDF report generated successfully at:\n{os.path.abspath(pdf_report_path)}")
            self.status_label.config(text="Status: PDF report generated.")
            os.startfile(os.path.abspath(pdf_report_path))
        else:
            messagebox.showerror("Error", "Failed to generate standard PDF report.")
            self.status_label.config(text="Status: PDF report generation failed.")

    def generate_comparison_excel_report(self):
        if self.df_raw.empty:
            messagebox.showwarning("Warning", "No raw data loaded. Please load data first.")
            return

        self.comp_status_label.config(text="Status: Generating comparison Excel report...")
        self.master.update_idletasks()

        comparison_mode = self.comparison_mode_var.get()
        df_comparison, message = app_logic.apply_comparison_filters(self.df_raw, self.comparison_config, comparison_mode)

        if df_comparison.empty:
            messagebox.showwarning("Warning", f"No data for comparison. {message}")
            self.comp_status_label.config(text="Status: Comparison report failed (no data)")
            return

        output_file = self.paths['comparison_output_file']
        success = app_logic.export_comparison_report(df_comparison, self.comparison_config, output_file, comparison_mode)

        if success:
            messagebox.showinfo("Success", f"Comparison Excel report generated successfully at:\n{os.path.abspath(output_file)}")
            self.comp_status_label.config(text="Status: Comparison Excel report generated.")
            os.startfile(os.path.abspath(output_file))
        else:
            messagebox.showerror("Error", "Failed to generate comparison Excel report.")
            self.comp_status_label.config(text="Status: Comparison Excel report generation failed.")

    def generate_comparison_pdf_report(self):
        if self.df_raw.empty:
            messagebox.showwarning("Warning", "No raw data loaded. Please load data first.")
            return

        self.comp_status_label.config(text="Status: Generating comparison PDF report...")
        self.master.update_idletasks()

        comparison_mode = self.comparison_mode_var.get()
        df_comparison, message = app_logic.apply_comparison_filters(self.df_raw, self.comparison_config, comparison_mode)

        if df_comparison.empty:
            messagebox.showwarning("Warning", f"No data for comparison. {message}")
            self.comp_status_label.config(text="Status: Comparison PDF report failed (no data)")
            return
            
        pdf_report_path = self.paths['comparison_pdf_report']
        logo_path = self.paths['logo_path']
        success = app_logic.export_comparison_pdf_report(df_comparison, self.comparison_config, pdf_report_path, comparison_mode, logo_path)

        if success:
            messagebox.showinfo("Success", f"Comparison PDF report generated successfully at:\n{os.path.abspath(pdf_report_path)}")
            self.comp_status_label.config(text="Status: Comparison PDF report generated.")
            os.startfile(os.path.abspath(pdf_report_path))
        else:
            messagebox.showerror("Error", "Failed to generate comparison PDF report.")
            self.comp_status_label.config(text="Status: Comparison PDF report generation failed.")

def main():
    root = tk.Tk()
    app = TimeReportApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
