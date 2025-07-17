import pandas as pd
import datetime
import os
from openpyxl import load_workbook
import tempfile
import re
import streamlit as st # Đảm bảo dòng này có

def sanitize_filename(filename):
    """
    Sanitizes a string to be safe for use as a filename.
    Removes invalid characters and replaces spaces with underscores.
    """
    # Remove invalid characters
    s = re.sub(r'[\\/:*?"<>|]', '', filename)
    # Replace spaces with underscores
    s = s.replace(' ', '_')
    return s.strip()

def setup_paths():
    # Không còn cần thiết nếu bạn đang đọc file cố định và không có cấu hình path phức tạp
    # Có thể để trống hoặc xóa hàm này nếu không còn sử dụng
    pass

@st.cache_data(show_spinner=False)
def read_configs(file_path):
    config_data = {}
    try:
        xls = pd.ExcelFile(file_path)
        
        # Read Config_Year_Mode
        if 'Config_Year_Mode' in xls.sheet_names:
            df_year_mode = pd.read_excel(xls, sheet_name='Config_Year_Mode')
            config_data['year_mode'] = df_year_mode.set_index('Year Mode')['Value'].to_dict()
        else:
            st.warning(f"Warning: Configuration sheet 'Config_Year_Mode' not found in {file_path}. Using default values.")
            config_data['year_mode'] = {'Single Year': 'Single Year', 'Year Over Year (YoY)': 'Year Over Year (YoY)'} # Default values

        # Read Config_Project_Filter
        if 'Config_Project_Filter' in xls.sheet_names:
            df_project_filter = pd.read_excel(xls, sheet_name='Config_Project_Filter')
            config_data['project_filter'] = df_project_filter.set_index('Key')['Value'].to_dict()
        else:
            st.warning(f"Warning: Configuration sheet 'Config_Project_Filter' not found in {file_path}. No project filters applied from config.")
            config_data['project_filter'] = {} # Default to empty

    except Exception as e:
        st.error(f"Error reading configuration sheets from {file_path}: {e}")
        config_data = {'year_mode': {'Single Year': 'Single Year', 'Year Over Year (YoY)': 'Year Over Year (YoY)'}, 'project_filter': {}}
    return config_data

@st.cache_data(show_spinner=False)
def load_raw_data(file_path):
    try:
        # Load the workbook to check for sheet existence
        workbook = load_workbook(file_path, read_only=True)
        if 'Raw Data' not in workbook.sheetnames:
            st.error(f"Lỗi: Sheet 'Raw Data' không tìm thấy trong file Excel '{file_path}'.")
            return pd.DataFrame() # Trả về DataFrame rỗng nếu không tìm thấy sheet
        
        # Read 'Raw Data' sheet
        raw_df = pd.read_excel(file_path, sheet_name='Raw Data', engine='openpyxl')
        
        # Clean column names (replace spaces with underscores, make lowercase)
        raw_df.columns = [col.strip().replace(' ', '_').lower() for col in raw_df.columns]
        
        # Ensure 'date' column is datetime
        if 'date' in raw_df.columns:
            raw_df['date'] = pd.to_datetime(raw_df['date'], errors='coerce')
            raw_df = raw_df.dropna(subset=['date']) # Remove rows with invalid dates
            
            # Extract Year and Month (full month name)
            raw_df['year'] = raw_df['date'].dt.year
            raw_df['month'] = raw_df['date'].dt.strftime('%B') # Full month name
        else:
            st.error("Column 'date' not found in 'Raw Data' sheet. Please check your Excel file.")
            return pd.DataFrame()

        # Rename columns to a more user-friendly format for display
        raw_df.rename(columns={
            'project_name': 'Project name',
            'task_name': 'Task name',
            'employee_name': 'Employee name',
            'hours': 'Hours',
            'year': 'Year',
            'month': 'Month'
        }, inplace=True)
        
        # Ensure 'Hours' column is numeric
        raw_df['Hours'] = pd.to_numeric(raw_df['Hours'], errors='coerce').fillna(0)
        
        return raw_df
    except Exception as e:
        st.error(f"Error loading raw data from '{file_path}': {e}")
        return pd.DataFrame()

def apply_filters(df, years=None, months=None, projects=None):
    filtered_df = df.copy()
    
    if years:
        filtered_df = filtered_df[filtered_df['Year'].isin(years)]
    if months:
        filtered_df = filtered_df[filtered_df['Month'].isin(months)]
    if projects:
        filtered_df = filtered_df[filtered_df['Project name'].isin(projects)]
        
    return filtered_df
