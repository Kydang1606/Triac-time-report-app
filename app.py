import streamlit as st
import pandas as pd
import datetime
import os
import io
import tempfile

# Import logic modules
# Đảm bảo các hàm trong các module này không gọi st trực tiếp nếu không cần thiết
# Hoặc đã import streamlit as st trong từng module đó nếu có dùng st
from data_processing_utils import setup_paths, read_configs, load_raw_data, apply_filters, sanitize_filename
from standard_report_logic import export_standard_report_excel, export_standard_report_pdf
from comparison_report_logic import apply_comparison_filters, export_comparison_report_excel, export_comparison_report_pdf

# --- Translation Dictionaries ---
translations = {
    "en": {
        "app_title": "Triac Time Report Application",
        "email_auth_header": "Email Authentication",
        "enter_email_label": "Enter your Triac Email",
        "auth_button_label": "Authenticate",
        "auth_success": "Authentication successful! Welcome, {email}.",
        "auth_invalid": "Invalid email or unauthorized access.",
        "contact_support_prompt": "Please contact IT support if you believe this is an error.",
        "sidebar_header": "Navigation",
        "dashboard_tab": "Dashboard",
        "standard_report_tab": "Standard Report",
        "comparison_report_tab": "Comparison Report",
        "year_mode_selection": "Select Year Mode:",
        "single_year": "Single Year",
        "year_over_year": "Year Over Year (YoY)",
        "select_year": "Select Year:",
        "select_current_year": "Select Current Year:",
        "select_previous_year": "Select Previous Year:",
        "select_project": "Select Project (Multi-select available):",
        "select_project_single": "Select Project (Single-select available):",
        "select_month": "Select Month:",
        "select_reporting_month": "Select Reporting Month:",
        "select_comparison_month": "Select Comparison Month:",
        "apply_filters": "Apply Filters",
        "export_excel": "Export to Excel",
        "export_pdf": "Export to PDF",
        "summary_data_header": "Summary Data",
        "comparison_data_header": "Comparison Data",
        "filtered_data_header": "Filtered Data",
        "error_processing_excel": "Error processing Excel file: {error}. Please ensure the file format is correct and try again.",
        "info_data_read": "Data loaded successfully! Total rows: {num_rows}, Projects found: {num_projects}.",
        "error_empty_raw_data": "Raw data sheet is empty or could not be read. Please check your Excel file.",
        "loading_charts_data": "Loading charts and data...",
        "no_data_for_selected_filters": "No data available for the selected filters. Please adjust your selections.",
        "timesheet_hours_chart": "Timesheet Hours by Project",
        "timesheet_hours_month_chart": "Timesheet Hours by Month",
        "comparison_hours_chart": "Comparison of Hours by Project",
        "hours_by_project_title": "Total Hours by Project",
        "hours_by_month_title": "Total Hours by Month",
        "comparison_title": "Comparison Report - {current_month} vs {comparison_month}",
        "export_success": "Report exported successfully!",
        "export_error": "Error exporting report: {error}",
        "no_data_for_export": "No data to export with current filters.",
        "loading_message": "Loading Application...",
        "sidebar_settings_header": "Settings",
        "select_language_label": "Select Language:",
        "welcome_message": "Welcome to Triac Time Report Application. Please select a report type.",
        "year_total": "Year Total",
        "month_total": "Month Total",
        "total_hours": "Total Hours",
        "delta_hours": "Delta Hours",
        "percentage_change": "Percentage Change",
        "config_not_found": "Configuration sheet '{sheet_name}' not found in Excel file.",
        "raw_data_not_found": "Raw Data sheet not found in Excel file."
    },
    "vi": {
        "app_title": "Ứng dụng Báo cáo Thời gian Triac",
        "email_auth_header": "Xác thực Email",
        "enter_email_label": "Nhập Email Triac của bạn",
        "auth_button_label": "Xác thực",
        "auth_success": "Xác thực thành công! Chào mừng, {email}.",
        "auth_invalid": "Email không hợp lệ hoặc truy cập không được phép.",
        "contact_support_prompt": "Vui lòng liên hệ bộ phận IT hỗ trợ nếu bạn nghĩ đây là lỗi.",
        "sidebar_header": "Điều hướng",
        "dashboard_tab": "Trang tổng quan",
        "standard_report_tab": "Báo cáo Tiêu chuẩn",
        "comparison_report_tab": "Báo cáo So sánh",
        "year_mode_selection": "Chọn Chế độ Năm:",
        "single_year": "Một Năm",
        "year_over_year": "So sánh Năm (YoY)",
        "select_year": "Chọn Năm:",
        "select_current_year": "Chọn Năm Hiện tại:",
        "select_previous_year": "Chọn Năm Trước:",
        "select_project": "Chọn Dự án (Có thể chọn nhiều):",
        "select_project_single": "Chọn Dự án (Chọn một):",
        "select_month": "Chọn Tháng:",
        "select_reporting_month": "Chọn Tháng Báo cáo:",
        "select_comparison_month": "Chọn Tháng So sánh:",
        "apply_filters": "Áp dụng Bộ lọc",
        "export_excel": "Xuất ra Excel",
        "export_pdf": "Xuất ra PDF",
        "summary_data_header": "Dữ liệu Tóm tắt",
        "comparison_data_header": "Dữ liệu So sánh",
        "filtered_data_header": "Dữ liệu đã lọc",
        "error_processing_excel": "Lỗi xử lý file Excel: {error}. Vui lòng đảm bảo định dạng file chính xác và thử lại.",
        "info_data_read": "Dữ liệu đã tải thành công! Tổng số dòng: {num_rows}, Dự án tìm thấy: {num_projects}.",
        "error_empty_raw_data": "Sheet dữ liệu thô trống hoặc không thể đọc được. Vui lòng kiểm tra file Excel của bạn.",
        "loading_charts_data": "Đang tải biểu đồ và dữ liệu...",
        "no_data_for_selected_filters": "Không có dữ liệu cho các bộ lọc đã chọn. Vui lòng điều chỉnh lựa chọn của bạn.",
        "timesheet_hours_chart": "Giờ Bảng chấm công theo Dự án",
        "timesheet_hours_month_chart": "Giờ Bảng chấm công theo Tháng",
        "comparison_hours_chart": "So sánh Giờ theo Dự án",
        "hours_by_project_title": "Tổng số giờ theo Dự án",
        "hours_by_month_title": "Tổng số giờ theo Tháng",
        "comparison_title": "Báo cáo So sánh - {current_month} vs {comparison_month}",
        "export_success": "Báo cáo đã xuất thành công!",
        "export_error": "Lỗi khi xuất báo cáo: {error}",
        "no_data_for_export": "Không có dữ liệu để xuất với các bộ lọc hiện tại.",
        "loading_message": "Đang tải Ứng dụng...",
        "sidebar_settings_header": "Cài đặt",
        "select_language_label": "Chọn Ngôn ngữ:",
        "welcome_message": "Chào mừng đến với Ứng dụng Báo cáo Thời gian Triac. Vui lòng chọn loại báo cáo.",
        "year_total": "Tổng năm",
        "month_total": "Tổng tháng",
        "total_hours": "Tổng giờ",
        "delta_hours": "Chênh lệch giờ",
        "percentage_change": "Phần trăm thay đổi",
        "config_not_found": "Sheet cấu hình '{sheet_name}' không tìm thấy trong file Excel.",
        "raw_data_not_found": "Sheet Raw Data không tìm thấy trong file Excel."
    }
}

# --- Initial Setup ---
st.set_page_config(layout="wide", page_title=translations["en"]["app_title"])

# Khởi tạo ngôn ngữ nếu chưa có trong session_state
if 'language' not in st.session_state:
    st.session_state.language = 'en'

# Gán từ điển dịch thuật hiện tại dựa trên ngôn ngữ đã chọn
current_translations = translations[st.session_state.language]

# --- HIỂN THỊ LOGO VÀ LỰA CHỌN NGÔN NGỮ Ở ĐẦU SIDEBAR ---
logo_path = "triac_logo.png"
if os.path.exists(logo_path):
    st.sidebar.image(logo_path, width=200)
else:
    st.sidebar.warning("Logo file 'triac_logo.png' not found.")

# Lựa chọn ngôn ngữ
st.sidebar.header(current_translations["sidebar_settings_header"])
selected_language = st.sidebar.selectbox(
    current_translations["select_language_label"],
    options=['English', 'Tiếng Việt'],
    index=0 if st.session_state.language == 'en' else 1
)

# Cập nhật ngôn ngữ và rerun nếu có thay đổi
if selected_language == 'Tiếng Việt' and st.session_state.language != 'vi':
    st.session_state.language = 'vi'
    st.rerun()
elif selected_language == 'English' and st.session_state.language != 'en':
    st.session_state.language = 'en'
    st.rerun()

st.title(current_translations["app_title"])

# --- Hàm đọc email từ file CSV (Đảm bảo file 'invited_emails.csv' tồn tại) ---
def load_authorized_emails(csv_file_path):
    @st.cache_data(ttl=300) # Cache trong 5 phút
    def _load_emails(path):
        if not os.path.exists(path):
            st.error(f"Lỗi: File invited_emails.csv không tìm thấy tại {path}. Vui lòng đảm bảo file tồn tại.")
            return []
        try:
            df_emails = pd.read_csv(path, header=None, names=['email'])
            # Chuyển tất cả email thành chữ thường và loại bỏ khoảng trắng thừa
            return [email.strip().lower() for email in df_emails['email'].tolist() if pd.notna(email)]
        except Exception as e:
            st.error(f"Lỗi khi đọc file invited_emails.csv: {e}. Vui lòng kiểm tra định dạng file.")
            return []
    return _load_emails(csv_file_path)

# --- Xác định đường dẫn và tải danh sách email ---
csv_path_for_auth = "invited_emails.csv"
AUTHORIZED_EMAILS = load_authorized_emails(csv_path_for_auth)

# --- Authentication ---
def authenticate_user(email):
    processed_email = email.lower().strip()
    return processed_email in AUTHORIZED_EMAILS

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.user_email = ""

if not st.session_state.authenticated:
    st.sidebar.header(current_translations["email_auth_header"])
    user_email = st.sidebar.text_input(current_translations["enter_email_label"]).strip()
    if st.sidebar.button(current_translations["auth_button_label"]):
        if authenticate_user(user_email):
            st.session_state.authenticated = True
            st.session_state.user_email = user_email
            st.sidebar.success(current_translations["auth_success"].format(email=user_email))
            st.rerun()
        else:
            st.sidebar.error(current_translations["auth_invalid"])
            st.sidebar.write(current_translations["contact_support_prompt"])
    st.stop() # Stop execution if not authenticated

# --- Logic để đọc Time_report.xlsm CỐ ĐỊNH (chỉ chạy sau khi xác thực thành công) ---
fixed_excel_file_path = "Time_report.xlsm"

@st.cache_data(show_spinner=False) # Cache kết quả đọc file để tăng tốc độ
def get_processed_data_from_fixed_file(file_path):
    # Sử dụng hàm từ data_processing_utils để đọc cấu hình và dữ liệu thô
    config_data = read_configs(file_path)
    raw_df = load_raw_data(file_path)
    return config_data, raw_df

if st.session_state.authenticated: # Đảm bảo chỉ đọc file sau khi xác thực
    with st.spinner(current_translations["reading_processing_data"]):
        try:
            config_data, raw_df = get_processed_data_from_fixed_file(fixed_excel_file_path)

            if raw_df.empty:
                st.error(current_translations["error_empty_raw_data"])
                st.stop()

            st.session_state['raw_df'] = raw_df
            st.session_state['config_data'] = config_data

            # Lấy danh sách các dự án, năm, tháng từ dữ liệu
            all_project_names = sorted(raw_df['Project name'].unique().tolist())
            all_years_in_data = sorted(raw_df['Year'].unique().tolist())
            all_months_in_data = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

            st.session_state['all_project_names'] = all_project_names
            st.session_state['all_years_in_data'] = all_years_in_data
            st.session_state['all_months_in_data'] = all_months_in_data

            st.sidebar.info(current_translations["info_data_read"].format(num_rows=len(raw_df), num_projects=len(all_project_names)))

        except FileNotFoundError:
            st.error(f"Lỗi: File '{fixed_excel_file_path}' không tìm thấy trong thư mục ứng dụng. Vui lòng đảm bảo file đã được upload lên GitHub.")
            st.stop()
        except Exception as e:
            st.error(current_translations["error_processing_excel"].format(error=e))
            st.stop()

    # --- Sidebar and Main Content (Chỉ hiển thị sau khi xác thực và tải dữ liệu thành công) ---
    st.sidebar.header(current_translations["sidebar_header"])

    # Tabs
    tab1, tab2, tab3 = st.tabs([
        current_translations["dashboard_tab"],
        current_translations["standard_report_tab"],
        current_translations["comparison_report_tab"]
    ])

    with tab1:
        st.header(current_translations["dashboard_tab"])
        st.write(current_translations["welcome_message"])
        
        if 'raw_df' in st.session_state and not st.session_state['raw_df'].empty:
            st.subheader(current_translations["summary_data_header"])
            total_hours = st.session_state['raw_df']['Hours'].sum()
            st.metric(label=current_translations["total_hours"], value=f"{total_hours:.2f}")

            # Total Hours by Project
            hours_by_project = st.session_state['raw_df'].groupby('Project name')['Hours'].sum().sort_values(ascending=False)
            st.subheader(current_translations["hours_by_project_title"])
            st.dataframe(hours_by_project)

            # Total Hours by Month (đảm bảo sắp xếp đúng theo thứ tự tháng)
            month_order = {
                'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
                'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
            }
            # Tạo cột số tháng để sắp xếp
            temp_df_for_month_sort = st.session_state['raw_df'].copy()
            temp_df_for_month_sort['Month_Num'] = temp_df_for_month_sort['Month'].map(month_order)
            hours_by_month = temp_df_for_month_sort.groupby('Month')['Hours'].sum().loc[temp_df_for_month_sort.sort_values('Month_Num')['Month'].unique()]
            st.subheader(current_translations["hours_by_month_title"])
            st.dataframe(hours_by_month)

        else:
            st.info(current_translations["no_data_for_selected_filters"])

    with tab2:
        st.header(current_translations["standard_report_tab"])
        
        if 'all_years_in_data' not in st.session_state or not st.session_state['all_years_in_data']:
            st.warning(current_translations["no_data_for_selected_filters"])
        else:
            standard_report_year_mode = st.radio(
                current_translations["year_mode_selection"],
                (current_translations["single_year"], current_translations["year_over_year"]),
                key="standard_year_mode"
            )

            selected_years = []
            if standard_report_year_mode == current_translations["single_year"]:
                selected_year_single = st.selectbox(
                    current_translations["select_year"],
                    options=st.session_state['all_years_in_data'],
                    key="standard_select_year"
                )
                selected_years.append(selected_year_single)
            else: # Year Over Year
                col_curr_year, col_prev_year = st.columns(2)
                with col_curr_year:
                    selected_current_year = st.selectbox(
                        current_translations["select_current_year"],
                        options=st.session_state['all_years_in_data'],
                        key="standard_select_current_year"
                    )
                with col_prev_year:
                    available_previous_years = sorted([y for y in st.session_state['all_years_in_data'] if y < selected_current_year], reverse=True)
                    if available_previous_years:
                        selected_previous_year = st.selectbox(
                            current_translations["select_previous_year"],
                            options=available_previous_years,
                            index=0, # Chọn năm lớn nhất có sẵn theo mặc định
                            key="standard_select_previous_year"
                        )
                        selected_years.extend([selected_current_year, selected_previous_year])
                    else:
                        st.warning("No previous years available for comparison.")
                        selected_years.append(selected_current_year)


            selected_projects_standard = st.multiselect(
                current_translations["select_project"],
                options=st.session_state['all_project_names'],
                default=st.session_state['all_project_names'],
                key="standard_select_project"
            )
            selected_month_standard = st.selectbox(
                current_translations["select_month"],
                options=st.session_state['all_months_in_data'],
                key="standard_select_month"
            )

            if st.button(current_translations["apply_filters"], key="apply_standard_filters"):
                if not selected_projects_standard or not selected_years or not selected_month_standard:
                    st.warning("Please select at least one project, a year, and a month.")
                else:
                    filtered_df_standard = apply_filters(
                        st.session_state['raw_df'],
                        years=selected_years,
                        months=[selected_month_standard],
                        projects=selected_projects_standard
                    )

                    if filtered_df_standard.empty:
                        st.info(current_translations["no_data_for_selected_filters"])
                    else:
                        st.subheader(current_translations["filtered_data_header"])
                        st.dataframe(filtered_df_standard)

                        # Export options
                        col_excel_std, col_pdf_std = st.columns(2)
                        with col_excel_std:
                            if st.button(current_translations["export_excel"], key="export_standard_excel_btn"): # Đổi key để tránh xung đột
                                with st.spinner(current_translations["loading_charts_data"]):
                                    try:
                                        excel_buffer = export_standard_report_excel(
                                            filtered_df_standard,
                                            st.session_state['config_data'],
                                            current_translations,
                                            standard_report_year_mode == current_translations["year_over_year"]
                                        )
                                        st.download_button(
                                            label=current_translations["export_excel"],
                                            data=excel_buffer.getvalue(),
                                            file_name=f"{sanitize_filename(current_translations['standard_report_tab'])}_{selected_month_standard}_{'-'.join(map(str, selected_years))}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key="download_standard_excel"
                                        )
                                        st.success(current_translations["export_success"])
                                    except Exception as e:
                                        st.error(current_translations["export_error"].format(error=e))
                        with col_pdf_std:
if st.button(current_translations["apply_filters"], key="apply_standard_filters"):
                if not selected_projects_standard or not selected_years or not selected_month_standard:
                    st.warning("Please select at least one project, a year, and a month.")
                else: # <-- This 'else' belongs to the 'if not selected_projects...'
                    filtered_df_standard = apply_filters(
                        st.session_state['raw_df'],
                        years=selected_years,
                        months=[selected_month_standard],
                        projects=selected_projects_standard
                    )

                    # Dòng 'if filtered_df_standard.empty:' này có ELSE của nó.
                    if filtered_df_standard.empty:
                        st.info(current_translations["no_data_for_selected_filters"])
                    else: # <--- ĐÂY LÀ KHỐI ELSE MÀ BẠN GẶP LỖI
                        st.subheader(current_translations["filtered_data_header"])
                        st.dataframe(filtered_df_standard)

                        # Export options
                        col_excel_std, col_pdf_std = st.columns(2)
                        with col_excel_std:
                            if st.button(current_translations["export_excel"], key="export_standard_excel_btn"):
                                with st.spinner(current_translations["loading_charts_data"]):
                                    try:
                                        excel_buffer = export_standard_report_excel(
                                            filtered_df_standard,
                                            st.session_state['config_data'],
                                            current_translations,
                                            standard_report_year_mode == current_translations["year_over_year"]
                                        )
                                        st.download_button(
                                            label=current_translations["export_excel"],
                                            data=excel_buffer.getvalue(),
                                            file_name=f"{sanitize_filename(current_translations['standard_report_tab'])}_{selected_month_standard}_{'-'.join(map(str, selected_years))}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key="download_standard_excel"
                                        )
                                        st.success(current_translations["export_success"])
                                    except Exception as e:
                                        st.error(current_translations["export_error"].format(error=e))
                        with col_pdf_std:
                            if st.button(current_translations["export_pdf"], key="export_standard_pdf_btn"): # Đổi key để tránh xung đột
                                with st.spinner(current_translations["loading_charts_data"]):
                                    try:
                                        pdf_buffer = export_standard_report_pdf(
                                            filtered_df_standard,
                                            current_translations,
                                            logo_path, # Truyền đường dẫn logo vào hàm export PDF
                                            standard_report_year_mode == current_translations["year_over_year"]
                                        )
                                        st.download_button(
                                            label=current_translations["export_pdf"],
                                            data=pdf_buffer.getvalue(),
                                            file_name=f"{sanitize_filename(current_translations['standard_report_tab'])}_{selected_month_standard}_{'-'.join(map(str, selected_years))}.pdf",
                                            mime="application/pdf",
                                            key="download_standard_pdf"
                                        )
                                        st.success(current_translations["export_success"])
                                    except Exception as e:
                                        st.error(current_translations["export_error"].format(error=e))
                        # Dòng 'else' mà bạn đang gặp lỗi phải được XÓA khỏi vị trí này.
                        # Nó không thuộc về khối if st.button(...) này.
                        # Loại bỏ hoặc thụt lề đúng để nó thuộc về if filtered_df_standard.empty:
                        # else:  <-- DÒNG NÀY LÀ NGUYÊN NHÂN LỖI CỦA BẠN. HÃY XÓA NÓ ĐI.
                        #    st.info(current_translations["no_data_for_export"])
            # else của if st.button(current_translations["apply_filters"], key="apply_standard_filters"):
            else: # <--- Đây là else đúng của if st.button("apply filters")
                 st.info("Please apply filters to see report options.") # Một thông báo bổ sung.

    with tab3:
        st.header(current_translations["comparison_report_tab"])
        if 'all_years_in_data' not in st.session_state or not st.session_state['all_years_in_data']:
            st.warning(current_translations["no_data_for_selected_filters"])
        else:
            selected_year_comp = st.selectbox(
                current_translations["select_year"],
                options=st.session_state['all_years_in_data'],
                key="comparison_select_year"
            )

            col_report_month, col_comp_month = st.columns(2)
            with col_report_month:
                selected_reporting_month = st.selectbox(
                    current_translations["select_reporting_month"],
                    options=st.session_state['all_months_in_data'],
                    key="comparison_reporting_month"
                )
            with col_comp_month:
                # Lọc bỏ tháng báo cáo khỏi lựa chọn tháng so sánh
                available_comp_months = [m for m in st.session_state['all_months_in_data'] if m != selected_reporting_month]
                selected_comparison_month = st.selectbox(
                    current_translations["select_comparison_month"],
                    options=available_comp_months,
                    # Đặt mặc định là tháng đầu tiên có sẵn nếu có
                    index=0 if available_comp_months else 0,
                    key="comparison_comparison_month"
                )

            selected_project_comparison = st.multiselect(
                current_translations["select_project"],
                options=st.session_state['all_project_names'],
                default=st.session_state['all_project_names'],
                key="comparison_select_project"
            )

            if st.button(current_translations["apply_filters"], key="apply_comparison_filters"):
                if not selected_project_comparison or not selected_year_comp or not selected_reporting_month or not selected_comparison_month:
                    st.warning("Please select a year, reporting month, comparison month, and at least one project.")
                else:
                    try:
                        comparison_df = apply_comparison_filters(
                            st.session_state['raw_df'],
                            selected_year_comp,
                            selected_reporting_month,
                            selected_comparison_month,
                            selected_project_comparison
                        )

                        if comparison_df.empty:
                            st.info(current_translations["no_data_for_selected_filters"])
                        else:
                            st.subheader(current_translations["comparison_data_header"])
                            st.dataframe(comparison_df)

                            # Export options for comparison report
                            col_excel_comp, col_pdf_comp = st.columns(2)
                            with col_excel_comp:
                                if st.button(current_translations["export_excel"], key="export_comparison_excel_btn"): # Đổi key
                                    with st.spinner(current_translations["loading_charts_data"]):
                                        try:
                                            excel_buffer = export_comparison_report_excel(
                                                comparison_df,
                                                current_translations,
                                                selected_reporting_month,
                                                selected_comparison_month,
                                                st.session_state['config_data']
                                            )
                                            st.download_button(
                                                label=current_translations["export_excel"],
                                                data=excel_buffer.getvalue(),
                                                file_name=f"{sanitize_filename(current_translations['comparison_report_tab'])}_{selected_reporting_month}_vs_{selected_comparison_month}_{selected_year_comp}.xlsx",
                                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                key="download_comparison_excel"
                                            )
                                            st.success(current_translations["export_success"])
                                        except Exception as e:
                                            st.error(current_translations["export_error"].format(error=e))
                            with col_pdf_comp:
                                if st.button(current_translations["export_pdf"], key="export_comparison_pdf_btn"): # Đổi key
                                    with st.spinner(current_translations["loading_charts_data"]):
                                        try:
                                            pdf_buffer = export_comparison_report_pdf(
                                                comparison_df,
                                                current_translations,
                                                logo_path, # Truyền đường dẫn logo vào hàm export PDF
                                                selected_reporting_month,
                                                selected_comparison_month
                                            )
                                            st.download_button(
                                                label=current_translations["export_pdf"],
                                                data=pdf_buffer.getvalue(),
                                                file_name=f"{sanitize_filename(current_translations['comparison_report_tab'])}_{selected_reporting_month}_vs_{selected_comparison_month}_{selected_year_comp}.pdf",
                                                mime="application/pdf",
                                                key="download_comparison_pdf"
                                            )
                                            st.success(current_translations["export_success"])
                                        except Exception as e:
                                            st.error(current_translations["export_error"].format(error=e))
                    except Exception as e:
                        st.error(f"Error generating comparison report: {e}")
