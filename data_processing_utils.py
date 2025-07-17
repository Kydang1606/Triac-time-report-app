import pandas as pd
import datetime
import os
from openpyxl import load_workbook #
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
        workbook = load_workbook(file_path, read_only=True) #

        # --- Read Config_Year_Mode ---
        if 'Config_Year_Mode' in workbook.sheetnames: #
            df_year_mode = pd.read_excel(file_path, sheet_name='Config_Year_Mode', engine='openpyxl') #
            
            # Đảm bảo các cột 'Key' và 'Value' tồn tại trước khi set_index
            if 'Key' in df_year_mode.columns and 'Value' in df_year_mode.columns: #
                config_year_mode_dict = df_year_mode.set_index('Key')['Value'].to_dict() #
                
                # Lấy giá trị 'Mode' (ví dụ: 'year')
                config_data['year_mode_config'] = str(config_year_mode_dict.get('Mode', 'single year')).strip().lower() #
                
                # Lấy giá trị 'Year' (ví dụ: 2025)
                # Cần xử lý để đảm bảo nó là số nguyên
                default_year_val = config_year_mode_dict.get('Year', datetime.datetime.now().year) #
                try:
                    config_data['default_year'] = int(default_year_val) #
                except (ValueError, TypeError):
                    st.warning(f"Giá trị 'Year' trong cấu hình ('{default_year_val}') không hợp lệ. Sử dụng năm hiện tại.") #
                    config_data['default_year'] = datetime.datetime.now().year #

                # Lấy giá trị 'Months' (ví dụ: 'March')
                config_data['default_month'] = str(config_year_mode_dict.get('Months', datetime.datetime.now().strftime('%B'))).strip() #
                
                # Lấy giá trị 'Weeks' nếu có
                config_data['default_weeks'] = config_year_mode_dict.get('Weeks', None) #

            else:
                st.error("Lỗi cấu hình: Sheet 'Config_Year_Mode' thiếu cột 'Key' hoặc 'Value'.") #
                # Đặt giá trị mặc định để tránh lỗi tiếp theo
                config_data['year_mode_config'] = 'single year' #
                config_data['default_year'] = datetime.datetime.now().year #
                config_data['default_month'] = datetime.datetime.now().strftime('%B') #
                config_data['default_weeks'] = None #

        else:
            st.warning(f"Cảnh báo: Sheet cấu hình 'Config_Year_Mode' không tìm thấy trong '{file_path}'. Sử dụng các giá trị mặc định.") #
            config_data['year_mode_config'] = 'single year' #
            config_data['default_year'] = datetime.datetime.now().year #
            config_data['default_month'] = datetime.datetime.now().strftime('%B') #
            config_data['default_weeks'] = None #

        # --- Read Config_Project_Filter ---
        # Kiểm tra sheet 'Config_Project_Filter'
        if 'Config_Project_Filter' in workbook.sheetnames: #
            df_project_filter = pd.read_excel(file_path, sheet_name='Config_Project_Filter', engine='openpyxl') #
            # Giả sử sheet này có 1 cột 'Project Name' chứa danh sách các dự án mặc định
            if 'Project Name' in df_project_filter.columns: #
                config_data['default_projects_filter'] = df_project_filter['Project Name'].dropna().tolist() #
            else:
                st.warning("Cảnh báo: Sheet 'Config_Project_Filter' không có cột 'Project Name'. Không áp dụng bộ lọc dự án từ cấu hình.") #
                config_data['default_projects_filter'] = [] #
        else:
            st.warning(f"Cảnh báo: Sheet cấu hình 'Config_Project_Filter' không tìm thấy trong '{file_path}'. Không áp dụng bộ lọc dự án từ cấu hình.") #
            config_data['default_projects_filter'] = [] #

    except Exception as e:
        st.error(f"Lỗi khi đọc các sheet cấu hình từ '{file_path}': {e}. Vui lòng đảm bảo định dạng file Excel chính xác.") #
        # Đặt các giá trị mặc định an toàn để ứng dụng vẫn có thể chạy
        config_data = { #
            'year_mode_config': 'single year', #
            'default_year': datetime.datetime.now().year, #
            'default_month': datetime.datetime.now().strftime('%B'), #
            'default_weeks': None, #
            'default_projects_filter': [] #
        }
    return config_data

@st.cache_data(show_spinner=False)
def load_raw_data(file_path):
    try:
        # Load the workbook to check for sheet existence
        workbook = load_workbook(file_path, read_only=True) #
        if 'Raw Data' not in workbook.sheetnames: #
            st.error(f"Lỗi: Sheet 'Raw Data' không tìm thấy trong file Excel '{file_path}'.") #
            return pd.DataFrame() # Trả về DataFrame rỗng nếu không tìm thấy sheet
        
        # Read 'Raw Data' sheet
        raw_df = pd.read_excel(file_path, sheet_name='Raw Data', engine='openpyxl') #
        
        # Clean column names (replace spaces with underscores, make lowercase)
        raw_df.columns = [col.strip().replace(' ', '_').lower() for col in raw_df.columns] #
        
        # Ensure 'date' column is datetime
        if 'date' in raw_df.columns: #
            raw_df['date'] = pd.to_datetime(raw_df['date'], errors='coerce') #
            raw_df = raw_df.dropna(subset=['date']) # Remove rows with invalid dates
            
            # Extract Year and Month (full month name)
            raw_df['year'] = raw_df['date'].dt.year #
            raw_df['month'] = raw_df['date'].dt.strftime('%B') # Full month name
        else:
            st.error("Column 'date' not found in 'Raw Data' sheet. Please check your Excel file.") #
            return pd.DataFrame() #

        # Rename columns to a more user-friendly format for display
        raw_df.rename(columns={ #
            'project_name': 'Project name', #
            'task_name': 'Task name', #
            'employee_name': 'Employee name', #
            'hours': 'Hours', #
            'year': 'Year', #
            'month': 'Month' #
        }, inplace=True) #
        
        # Ensure 'Hours' column is numeric
        raw_df['Hours'] = pd.to_numeric(raw_df['Hours'], errors='coerce').fillna(0) #
        
        return raw_df
    except Exception as e:
        st.error(f"Error loading raw data from '{file_path}': {e}") #
        return pd.DataFrame() #

def apply_filters(df, years=None, months=None, projects=None):
    filtered_df = df.copy() #
    
    if years: #
        # 'Year' đã được xử lý thành số trong load_raw_data
        filtered_df = filtered_df[filtered_df['Year'].isin(years)] #
    if months: #
        # 'Month' là tên tháng, so sánh với list tên tháng
        filtered_df = filtered_df[filtered_df['Month'].isin(months)] #
    if projects: #
        # 'Project name' là tên dự án
        filtered_df = filtered_df[filtered_df['Project name'].isin(projects)] #
        
    return filtered_df
