import pandas as pd
import datetime
import os
import io
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from matplotlib import pyplot as plt
import tempfile
import shutil
import streamlit as st # Đảm bảo dòng này có
from data_processing_utils import sanitize_filename # Đảm bảo import sanitize_filename

# Hàm áp dụng bộ lọc và tính toán so sánh
def apply_comparison_filters(raw_df, year, reporting_month, comparison_month, projects):
    # Lọc dữ liệu cho tháng báo cáo
    reporting_df = raw_df[
        (raw_df['Year'] == year) &
        (raw_df['Month'] == reporting_month) &
        (raw_df['Project name'].isin(projects))
    ].groupby('Project name')['Hours'].sum().reset_index()
    reporting_df.rename(columns={'Hours': f'Hours_{reporting_month}'}, inplace=True)

    # Lọc dữ liệu cho tháng so sánh
    comparison_df = raw_df[
        (raw_df['Year'] == year) &
        (raw_df['Month'] == comparison_month) &
        (raw_df['Project name'].isin(projects))
    ].groupby('Project name')['Hours'].sum().reset_index()
    comparison_df.rename(columns={'Hours': f'Hours_{comparison_month}'}, inplace=True)

    # Gộp hai DataFrame
    merged_df = pd.merge(
        reporting_df,
        comparison_df,
        on='Project name',
        how='outer'
    ).fillna(0) # Điền 0 cho các dự án không có giờ trong một tháng

    # Tính toán chênh lệch và phần trăm thay đổi
    merged_df['Delta Hours'] = merged_df[f'Hours_{reporting_month}'] - merged_df[f'Hours_{comparison_month}']
    merged_df['Percentage Change'] = (merged_df['Delta Hours'] / merged_df[f'Hours_{comparison_month}']) * 100
    merged_df.loc[merged_df[f'Hours_{comparison_month}'] == 0, 'Percentage Change'] = 0 # Tránh chia cho 0

    return merged_df

# Hàm xuất báo cáo so sánh ra Excel
def export_comparison_report_excel(comparison_df, translations, reporting_month, comparison_month, config_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet_name = translations["comparison_report_tab"]

        # Rename columns for clarity in Excel if needed, using translations
        excel_df = comparison_df.rename(columns={
            f'Hours_{reporting_month}': translations['total_hours'] + f' ({reporting_month})',
            f'Hours_{comparison_month}': translations['total_hours'] + f' ({comparison_month})',
            'Delta Hours': translations['delta_hours'],
            'Percentage Change': translations['percentage_change']
        })
        
        excel_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # You can add charts here using openpyxl if desired

    output.seek(0)
    return output

# Hàm xuất báo cáo so sánh ra PDF
def export_comparison_report_pdf(comparison_df, translations, logo_path, reporting_month, comparison_month):
    output = io.BytesIO()
    pdf = FPDF(unit="mm", format="A4")
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Add logo
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=10, y=10, w=30) # Điều chỉnh vị trí và kích thước logo

    pdf.set_font("Arial", "B", 16)
    title = translations["comparison_title"].format(current_month=reporting_month, comparison_month=comparison_month)
    pdf.cell(0, 10, title, ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("Arial", "B", 10)
    # Headers for comparison report
    headers = [
        "Project Name",
        translations['total_hours'] + f' ({reporting_month})',
        translations['total_hours'] + f' ({comparison_month})',
        translations['delta_hours'],
        translations['percentage_change']
    ]
    col_widths = [50, 35, 35, 30, 30] # Adjust widths as needed

    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 7, header, 1, 0, 'C')
    pdf.ln()

    pdf.set_font("Arial", "", 10)
    for index, row in comparison_df.iterrows():
        pdf.cell(col_widths[0], 7, str(row['Project name']), 1)
        pdf.cell(col_widths[1], 7, f"{row[f'Hours_{reporting_month}']:.2f}", 1, 0, 'R')
        pdf.cell(col_widths[2], 7, f"{row[f'Hours_{comparison_month}']:.2f}", 1, 0, 'R')
        pdf.cell(col_widths[3], 7, f"{row['Delta Hours']:.2f}", 1, 0, 'R')
        pdf.cell(col_widths[4], 7, f"{row['Percentage Change']:.2f}%", 1, 0, 'R')
        pdf.ln()
    pdf.ln(10)

    output.write(pdf.output(dest='S').encode('latin-1'))
    output.seek(0)
    return output
