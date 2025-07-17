import pandas as pd
import datetime
import os
from openpyxl import load_workbook
import tempfile
import re
import streamlit as st 

# Hàm hỗ trợ làm sạch tên file/sheet
def sanitize_filename(name):
    invalid_chars = re.compile(r'[\\/*?[\]:;|=,<>]')
    s = invalid_chars.sub("_", str(name))
    s = ''.join(c for c in s if c.isprintable())
    return s[:31] # Limit to 31 chars for Excel sheet names

def setup_paths():
    """Sets up default file paths. Used for naming conventions, not direct file access."""
    today = datetime.datetime.today().strftime('%Y%m%d')
    return {
        'template_file': "Time_report.xlsm", # This is a placeholder, actual file comes from upload
        'output_file': f"Time_report_Standard_{today}.xlsx",
        'pdf_report': f"Time_report_Standard_{today}.pdf",
        'comparison_output_file': f"Time_report_Comparison_{today}.xlsx",
        'comparison_pdf_report': f"Time_report_Comparison_{today}.pdf",
        'logo_path': "triac_logo.png" # This is a placeholder, actual logo comes from upload
    }

@st.cache_data(show_spinner=False)
def read_configs(template_file_path):
    """Reads configuration from the Excel template file."""
    try:
        year_mode_df = pd.read_excel(template_file_path, sheet_name='Config_Year_Mode', engine='openpyxl')
        project_filter_df = pd.read_excel(template_file_path, sheet_name='Config_Project_Filter', engine='openpyxl')

        mode_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value']
        mode = str(mode_row.values[0]).strip().lower() if not mode_row.empty and pd.notna(mode_row.values[0]) else 'year'

        year_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'year', 'Value']
        year = int(year_row.values[0]) if not year_row.empty and pd.notna(year_row.values[0]) and pd.api.types.is_number(year_row.values[0]) else datetime.datetime.now().year
        
        months_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'months', 'Value']
        months = [m.strip().capitalize() for m in str(months_row.values[0]).split(',')] if not months_row.empty and pd.notna(months_row.values[0]) else []
        
        if 'Include' in project_filter_df.columns:
            project_filter_df['Include'] = project_filter_df['Include'].astype(str).str.lower()

        return {
            'mode': mode,
            'year': year,
            'months': months,
            'project_filter_df': project_filter_df
        }
    except FileNotFoundError:
        st.error(f"Error: Template file not found at {template_file_path}. Please upload the correct Excel file.")
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}
    except Exception as e:
        st.error(f"Error reading configuration: {e}. Please ensure 'Config_Year_Mode' and 'Config_Project_Filter' sheets are correct.")
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}

@st.cache_data(show_spinner=False)
def load_raw_data(template_file_path):
    """Loads raw data from the Excel template file."""
    try:
        df = pd.read_excel(template_file_path, sheet_name='Raw Data', engine='openpyxl')
        df.columns = df.columns.str.strip()
        df.rename(columns={'Hou': 'Hours', 'Team member': 'Employee', 'Project Name': 'Project name'}, inplace=True)
        
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date']) 
        
        df['Year'] = df['Date'].dt.year
        df['MonthName'] = df['Date'].dt.month_name()
        df['Week'] = df['Date'].dt.isocalendar().week.astype(int)
        
        df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce').fillna(0)
        
        return df
    except Exception as e:
        st.error(f"Error loading raw data from 'Raw Data' sheet: {e}. Please ensure the sheet exists and columns are correct.")
        return pd.DataFrame()

def apply_filters(df, config):
    """Applies data filters based on configuration."""
    df_filtered = df.copy()

    if 'years' in config and config['years']: # For multi-year comparison
        df_filtered = df_filtered[df_filtered['Year'].isin(config['years'])]
    elif 'year' in config and config['year']: # For single-year standard report
        df_filtered = df_filtered[df_filtered['Year'] == config['year']]

    if config['months']:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    # Apply project filter based on 'Include' column in Config_Project_Filter
    if 'project_filter_df' in config and not config['project_filter_df'].empty:
        included_projects = config['project_filter_df'][
            config['project_filter_df']['Include'].astype(str).str.lower() == 'yes'
        ]['Project Name'].tolist()
        
        if included_projects:
            df_filtered = df_filtered[df_filtered['Project name'].isin(included_projects)]
        else:
            # If no projects are marked as 'yes' for inclusion, return empty DataFrame
            return pd.DataFrame(columns=df.columns)
    # If no project_filter_df or it's empty, no project filtering happens here, 
    # as per original logic, which assumed all projects if no explicit filter.
    # But for robustness, if it's explicitly about "included projects", then if none, return empty.
    # I'll keep the original logic's intent, if no filter, don't filter projects.
    # However, for the standard report, it explicitly uses the included_projects from config.
    # So, let's make this apply_filters more general:
    # If 'selected_projects' is provided in config, use that.
    elif 'selected_projects' in config and config['selected_projects']: # For comparison, where user selects projects
        df_filtered = df_filtered[df_filtered['Project name'].isin(config['selected_projects'])]
    
    return df_filtered
