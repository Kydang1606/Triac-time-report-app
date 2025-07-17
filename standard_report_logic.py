import pandas as pd
import datetime
import os
import io
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from matplotlib import pyplot as plt
import tempfile
import shutil
import streamlit as st # Đảm bảo dòng này có
from data_processing_utils import sanitize_filename # Đảm bảo import sanitize_filename

# Hàm xuất báo cáo tiêu chuẩn ra Excel
def export_standard_report_excel(filtered_df, config_data, translations, is_yoy_mode=False):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet_name_summary = translations["standard_report_tab"] + " Summary"
        sheet_name_raw = translations["standard_report_tab"] + " Raw Data"

        # Prepare data for summary (Total Hours by Project)
        summary_df_project = filtered_df.groupby('Project name')['Hours'].sum().reset_index()
        summary_df_project.rename(columns={'Hours': translations["total_hours"]}, inplace=True)
        
        # Prepare data for summary (Total Hours by Month)
        month_order = {
            'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
            'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
        }
        # Sắp xếp tháng dựa trên số thứ tự
        temp_df_month = filtered_df.copy()
        temp_df_month['Month_Num'] = temp_df_month['Month'].map(month_order)
        summary_df_month = temp_df_month.groupby('Month')['Hours'].sum().loc[temp_df_month.sort_values('Month_Num')['Month'].unique()].reset_index()
        summary_df_month.rename(columns={'Hours': translations["total_hours"]}, inplace=True)


        # Write Summary Dataframes
        summary_df_project.to_excel(writer, sheet_name=sheet_name_summary, startrow=0, startcol=0, index=False)
        summary_df_month.to_excel(writer, sheet_name=sheet_name_summary, startrow=summary_df_project.shape[0] + 2, startcol=0, index=False)

        # Write Filtered Raw Data
        filtered_df.to_excel(writer, sheet_name=sheet_name_raw, index=False)

        # Add charts (optional, can be complex in openpyxl directly)
        # For simplicity, we might just export data and let user create charts or use matplotlib in PDF export

    output.seek(0)
    return output

# Hàm xuất báo cáo tiêu chuẩn ra PDF
def export_standard_report_pdf(filtered_df, translations, logo_path, is_yoy_mode=False):
    output = io.BytesIO()
    pdf = FPDF(unit="mm", format="A4")
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Add logo
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=10, y=10, w=30) # Điều chỉnh vị trí và kích thước logo
    
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, translations["standard_report_tab"], ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, translations["summary_data_header"], ln=True)
    pdf.ln(5)

    # Summary by Project
    summary_df_project = filtered_df.groupby('Project name')['Hours'].sum().reset_index()
    summary_df_project.rename(columns={'Hours': translations["total_hours"]}, inplace=True)

    pdf.set_font("Arial", "B", 10)
    pdf.cell(50, 7, "Project Name", 1)
    pdf.cell(40, 7, translations["total_hours"], 1, ln=True)
    pdf.set_font("Arial", "", 10)
    for index, row in summary_df_project.iterrows():
        pdf.cell(50, 7, str(row['Project name']), 1)
        pdf.cell(40, 7, f"{row[translations['total_hours']]:.2f}", 1, ln=True)
    pdf.ln(10)

    # Summary by Month
    month_order = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
        'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
    }
    temp_df_month = filtered_df.copy()
    temp_df_month['Month_Num'] = temp_df_month['Month'].map(month_order)
    summary_df_month = temp_df_month.groupby('Month')['Hours'].sum().loc[temp_df_month.sort_values('Month_Num')['Month'].unique()].reset_index()
    summary_df_month.rename(columns={'Hours': translations["total_hours"]}, inplace=True)

    pdf.set_font("Arial", "B", 10)
    pdf.cell(50, 7, "Month", 1)
    pdf.cell(40, 7, translations["total_hours"], 1, ln=True)
    pdf.set_font("Arial", "", 10)
    for index, row in summary_df_month.iterrows():
        pdf.cell(50, 7, str(row['Month']), 1)
        pdf.cell(40, 7, f"{row[translations['total_hours']]:.2f}", 1, ln=True)
    pdf.ln(10)

    # Add filtered raw data
    pdf.add_page() # New page for raw data if needed
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, translations["filtered_data_header"], ln=True)
    pdf.ln(5)

    # Adjust font size for large tables
    if len(filtered_df.columns) > 5 or len(filtered_df) > 20:
        pdf.set_font("Arial", "", 8)
    else:
        pdf.set_font("Arial", "", 10)

    # Headers
    col_widths = {
        'Project name': 40, 'Task name': 40, 'Employee name': 40,
        'Hours': 20, 'Year': 20, 'Month': 20
    }
    header_cols = ['Project name', 'Task name', 'Employee name', 'Hours', 'Year', 'Month']
    
    # Ensure all header columns exist in filtered_df, if not, skip them or handle
    available_cols = [col for col in header_cols if col in filtered_df.columns]

    for col in available_cols:
        pdf.cell(col_widths.get(col, 25), 7, col, 1, 0, 'C')
    pdf.ln()

    # Data rows
    for index, row in filtered_df.iterrows():
        for col in available_cols:
            value = str(row[col])
            # Handle float formatting for 'Hours'
            if col == 'Hours':
                value = f"{float(value):.2f}"
            pdf.cell(col_widths.get(col, 25), 7, value, 1, 0)
        pdf.ln()

    output.write(pdf.output(dest='S').encode('latin-1'))
    output.seek(0)
    return output
