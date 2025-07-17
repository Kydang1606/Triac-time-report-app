import pandas as pd
import datetime
import os
from openpyxl import load_workbook
import tempfile
import re
import streamlit as st

def sanitize_filename(filename):
    """
    Sanitizes a string to be safe for use as a filename.
    Removes invalid characters and replaces spaces with underscores.
    """
    # Loại bỏ các ký tự không hợp lệ
    s = re.sub(r'[\\/:*?"<>|]', '', filename)
    # Thay thế khoảng trắng bằng dấu gạch dưới
    s = s.replace(' ', '_')
    return s.strip()

def setup_paths():
    # Hàm này không còn thực sự cần thiết nếu đường dẫn được xử lý động
    # hoặc nếu file được tải lên trực tiếp. Giữ lại làm chỗ giữ chỗ.
    pass

@st.cache_data(show_spinner=False)
def read_configs(file_path):
    config_data = {}
    try:
        workbook = load_workbook(file_path, read_only=True)

        # --- Đọc Config_Year_Mode ---
        if 'Config_Year_Mode' in workbook.sheetnames:
            df_year_mode = pd.read_excel(file_path, sheet_name='Config_Year_Mode', engine='openpyxl')
            
            # Đảm bảo các cột 'Key' và 'Value' tồn tại trước khi set_index
            if 'Key' in df_year_mode.columns and 'Value' in df_year_mode.columns:
                config_year_mode_dict = df_year_mode.set_index('Key')['Value'].to_dict()
                
                # Lấy giá trị 'Mode' (ví dụ: 'year')
                config_data['year_mode_config'] = str(config_year_mode_dict.get('Mode', 'single year')).strip().lower()
                
                # Lấy giá trị 'Year' (ví dụ: 2025)
                default_year_val = config_year_mode_dict.get('Year', datetime.datetime.now().year)
                try:
                    config_data['default_year'] = int(default_year_val)
                except (ValueError, TypeError):
                    st.warning(f"Giá trị 'Year' trong cấu hình ('{default_year_val}') không hợp lệ. Sử dụng năm hiện tại.")
                    config_data['default_year'] = datetime.datetime.now().year

                # Lấy giá trị 'Months' (ví dụ: 'March')
                config_data['default_month'] = str(config_year_mode_dict.get('Months', datetime.datetime.now().strftime('%B'))).strip()
                
                # Lấy giá trị 'Weeks' nếu có
                config_data['default_weeks'] = config_year_mode_dict.get('Weeks', None)

            else:
                st.error("Lỗi cấu hình: Sheet 'Config_Year_Mode' thiếu cột 'Key' hoặc 'Value'.")
                # Đặt giá trị mặc định để tránh lỗi tiếp theo
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

        # --- Đọc Config_Project_Filter ---
        if 'Config_Project_Filter' in workbook.sheetnames:
            df_project_filter = pd.read_excel(file_path, sheet_name='Config_Project_Filter', engine='openpyxl')
            if 'Project Name' in df_project_filter.columns:
                # Đảm bảo các giá trị trong cột 'Project Name' là chuỗi và loại bỏ NaN
                config_data['default_projects_filter'] = df_project_filter['Project Name'].astype(str).dropna().tolist()
            else:
                st.warning("Cảnh báo: Sheet 'Config_Project_Filter' không có cột 'Project Name'. Không áp dụng bộ lọc dự án từ cấu hình.")
                config_data['default_projects_filter'] = []
        else:
            st.warning(f"Cảnh báo: Sheet cấu hình 'Config_Project_Filter' không tìm thấy trong '{file_path}'. Không áp dụng bộ lọc dự án từ cấu hình.")
            config_data['default_projects_filter'] = []

    except Exception as e:
        st.error(f"Lỗi khi đọc các sheet cấu hình từ '{file_path}': {e}. Vui lòng đảm bảo định dạng file Excel chính xác.")
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
        
        # Làm sạch tên cột: loại bỏ khoảng trắng đầu/cuối, thay khoảng trắng bằng _, chuyển thành chữ thường
        raw_df.columns = [col.strip().replace(' ', '_').lower() for col in raw_df.columns]
        
        # --- Xử lý cột 'date' một cách mạnh mẽ hơn ---
        if 'date' in raw_df.columns:
            # Chuyển đổi tất cả các giá trị trong cột 'date' sang kiểu string trước khi chuyển đổi sang datetime.
            # Điều này giúp tránh lỗi float-str nếu có giá trị số không phải ngày.
            raw_df['date'] = raw_df['date'].astype(str)
            raw_df['date'] = pd.to_datetime(raw_df['date'], errors='coerce')
            
            invalid_dates_count = raw_df['date'].isna().sum()
            if invalid_dates_count > 0:
                st.warning(f"Có {invalid_dates_count} giá trị không hợp lệ trong cột 'date' của sheet 'Raw Data'. Các hàng này sẽ bị loại bỏ.")
                # Ghi nhật ký hoặc hiển thị các giá trị gốc không hợp lệ để dễ debug
                # st.write("Một số giá trị 'date' không hợp lệ gốc:")
                # st.write(raw_df[raw_df['date'].isna()]['date_original_column'].head()) # Nếu bạn giữ bản sao cột gốc
            
            raw_df = raw_df.dropna(subset=['date']) # Loại bỏ các hàng có ngày không hợp lệ

            if raw_df.empty:
                st.warning("Sau khi loại bỏ các hàng có ngày không hợp lệ, DataFrame 'Raw Data' trở nên rỗng.")
                return pd.DataFrame()
            
            raw_df['year'] = raw_df['date'].dt.year
            raw_df['month'] = raw_df['date'].dt.strftime('%B') # Tên tháng đầy đủ
        else:
            st.error("Lỗi: Cột 'date' không tìm thấy trong sheet 'Raw Data'. Vui lòng kiểm tra file Excel của bạn.")
            return pd.DataFrame()

        # Đổi tên cột sang định dạng thân thiện hơn để hiển thị
        raw_df.rename(columns={
            'project_name': 'Project name',
            'task_name': 'Task name',
            'employee_name': 'Employee name',
            'hours': 'Hours',
            'year': 'Year',
            'month': 'Month'
        }, inplace=True)
        
        # --- Đảm bảo cột 'Hours' là số và xử lý các kiểu hỗn hợp một cách phòng thủ ---
        if 'Hours' in raw_df.columns:
            # Chuyển đổi cột 'Hours' sang kiểu số, ép buộc lỗi thành NaN, sau đó điền 0 cho NaN
            # Điều này cũng giúp xử lý nếu có string trong cột hours
            raw_df['Hours'] = pd.to_numeric(raw_df['Hours'], errors='coerce').fillna(0)
        else:
            st.warning("Cảnh báo: Cột 'hours' (sau khi làm sạch tên) không tìm thấy trong sheet 'Raw Data'. Đặt cột 'Hours' về 0.")
            raw_df['Hours'] = 0 # Thêm cột nếu nó bị thiếu

        # Đảm bảo các cột 'Year' và 'Month' tồn tại sau khi đổi tên
        if 'Year' not in raw_df.columns or 'Month' not in raw_df.columns:
            st.error("Lỗi: Không thể trích xuất 'Year' hoặc 'Month' từ cột 'date'. Vui lòng kiểm tra định dạng dữ liệu ngày.")
            return pd.DataFrame()
            
        return raw_df
    except Exception as e:
        st.error(f"Lỗi khi tải dữ liệu thô từ '{file_path}': {e}. Vui lòng kiểm tra định dạng và nội dung file Excel.")
        return pd.DataFrame()

def apply_filters(df, years=None, months=None, projects=None):
    filtered_df = df.copy()
    
    # Đảm bảo các cột được sử dụng để lọc tồn tại trước khi áp dụng bộ lọc
    if 'Year' not in filtered_df.columns:
        st.warning("Cột 'Year' không tồn tại trong DataFrame để áp dụng bộ lọc. Bỏ qua bộ lọc năm.")
        years = None # Vô hiệu hóa bộ lọc nếu cột bị thiếu
    if 'Month' not in filtered_df.columns:
        st.warning("Cột 'Month' không tồn tại trong DataFrame để áp dụng bộ lọc. Bỏ qua bộ lọc tháng.")
        months = None # Vô hiệu hóa bộ lọc nếu cột bị thiếu
    if 'Project name' not in filtered_df.columns:
        st.warning("Cột 'Project name' không tồn tại trong DataFrame để áp dụng bộ lọc. Bỏ qua bộ lọc dự án.")
        projects = None # Vô hiệu hóa bộ lọc nếu cột bị thiếu

    # Áp dụng bộ lọc năm
    if years and not filtered_df.empty:
        # Chuyển đổi cột 'Year' sang kiểu số nguyên an toàn hơn
        filtered_df['Year'] = pd.to_numeric(filtered_df['Year'], errors='coerce').fillna(-1).astype(int)
        # Đảm bảo các năm trong bộ lọc cũng là số nguyên để so sánh an toàn
        years_numeric = [int(y) for y in years if isinstance(y, (int, float, str)) and str(y).replace('.', '').isdigit()]
        filtered_df = filtered_df[filtered_df['Year'].isin(years_numeric)]

    # Áp dụng bộ lọc tháng
    if months and not filtered_df.empty:
        # Đảm bảo cột 'Month' là kiểu string để so sánh
        filtered_df['Month'] = filtered_df['Month'].astype(str)
        filtered_df = filtered_df[filtered_df['Month'].isin(months)]
    
    # Áp dụng bộ lọc dự án
    if projects and not filtered_df.empty:
        # Đảm bảo cột 'Project name' là kiểu string để so sánh
        filtered_df['Project name'] = filtered_df['Project name'].astype(str)
        filtered_df = filtered_df[filtered_df['Project name'].isin(projects)]
        
    return filtered_df
