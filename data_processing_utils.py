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
    s = re.sub(r'[\\/:*?"<>|]', '', filename)
    s = s.replace(' ', '_')
    return s.strip()

def setup_paths():
    pass

@st.cache_data(show_spinner=False)
def read_configs(file_path):
    config_data = {}
    try:
        workbook = load_workbook(file_path, read_only=True)

        # --- Read Config_Year_Mode ---
        if 'Config_Year_Mode' in workbook.sheetnames:
            df_year_mode = pd.read_excel(file_path, sheet_name='Config_Year_Mode', engine='openpyxl')
            
            if 'Key' in df_year_mode.columns and 'Value' in df_year_mode.columns:
                config_year_mode_dict = df_year_mode.set_index('Key')['Value'].to_dict()
                
                # Get 'Mode' (e.g., 'year')
                # Ensure it's a string and lowercase for consistent comparison
                config_data['year_mode_config'] = str(config_year_mode_dict.get('Mode', 'single year')).strip().lower() 
                
                # Get 'Year' (e.g., 2025)
                # Ensure it's an integer
                default_year_val = config_year_mode_dict.get('Year', datetime.datetime.now().year) 
                try:
                    config_data['default_year'] = int(default_year_val) 
                except (ValueError, TypeError):
                    st.warning(f"Giá trị 'Year' trong cấu hình ('{default_year_val}') không hợp lệ. Sử dụng năm hiện tại.") 
                    config_data['default_year'] = datetime.datetime.now().year 

                # Get 'Months' (e.g., 'March')
                # Ensure it's a string
                config_data['default_month'] = str(config_year_mode_dict.get('Months', datetime.datetime.now().strftime('%B'))).strip() 
                
                # Get 'Weeks' if available
                config_data['default_weeks'] = config_year_mode_dict.get('Weeks', None) 

            else:
                st.error("Lỗi cấu hình: Sheet 'Config_Year_Mode' thiếu cột 'Key' hoặc 'Value'.") 
                # Provide safe default values to prevent further errors
                config_data['year_mode_config'] = 'single year' 
                config_data['default_year'] = datetime.datetime.now().year 
                config_data['default_month'] = datetime.datetime.now().strftime('%B') 
                config_data['default_weeks'] = None 

        else:
            st.warning(f"Cảnh báo: Sheet cấu hình 'Config_Year_Mode' không tìm thấy trong '{file_path}'. Sử dụng các giá trị mặc định.") 
            config_data['year_mode_config'] = 'single year' 
            config_data['default_year'] = datetime.datetime.now().year 
            config_data['default_month'] = datetime.datetime.now().strftime('%B') 
            config_data['default_weeks'] = None 

        # --- Read Config_Project_Filter ---
        if 'Config_Project_Filter' in workbook.sheetnames: 
            df_project_filter = pd.read_excel(file_path, sheet_name='Config_Project_Filter', engine='openpyxl') 
            if 'Project Name' in df_project_filter.columns: 
                config_data['default_projects_filter'] = df_project_filter['Project Name'].dropna().tolist() 
            else:
                st.warning("Cảnh báo: Sheet 'Config_Project_Filter' không có cột 'Project Name'. Không áp dụng bộ lọc dự án từ cấu hình.") 
                config_data['default_projects_filter'] = [] 
        else:
            st.warning(f"Cảnh báo: Sheet cấu hình 'Config_Project_Filter' không tìm thấy trong '{file_path}'. Không áp dụng bộ lọc dự án từ cấu hình.") 
            config_data['default_projects_filter'] = [] 

    except Exception as e:
        st.error(f"Lỗi khi đọc các sheet cấu hình từ '{file_path}': {e}. Vui lòng đảm bảo định dạng file Excel chính xác.") 
        # Provide safe default values if a critical error occurs
        config_data = { 
            'year_mode_config': 'single year', 
            'default_year': datetime.datetime.now().year, 
            'default_month': datetime.datetime.now().strftime('%B'), 
            'default_weeks': None, 
            'default_projects_filter': [] 
        }
    return config_data

@st.cache_data(show_spinner=False)
def load_raw_data(file_path):
    try:
        workbook = load_workbook(file_path, read_only=True) 
        if 'Raw Data' not in workbook.sheetnames: 
            st.error(f"Lỗi: Sheet 'Raw Data' không tìm thấy trong file Excel '{file_path}'.") 
            return pd.DataFrame() 
        
        raw_df = pd.read_excel(file_path, sheet_name='Raw Data', engine='openpyxl') 
        
        raw_df.columns = [col.strip().replace(' ', '_').lower() for col in raw_df.columns] 
        
        if 'date' in raw_df.columns: 
            raw_df['date'] = pd.to_datetime(raw_df['date'], errors='coerce') 
            raw_df = raw_df.dropna(subset=['date']) 
            
            raw_df['year'] = raw_df['date'].dt.year 
            raw_df['month'] = raw_df['date'].dt.strftime('%B') # Full month name
        else:
            st.error("Column 'date' not found in 'Raw Data' sheet. Please check your Excel file.") 
            return pd.DataFrame() 

        raw_df.rename(columns={ 
            'project_name': 'Project name', 
            'task_name': 'Task name', 
            'employee_name': 'Employee name', 
            'hours': 'Hours', 
            'year': 'Year', 
            'month': 'Month' 
        }, inplace=True) 
        
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
