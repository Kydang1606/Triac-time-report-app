import streamlit as st
import pandas as pd
import datetime
import os
import io
import tempfile

# Import logic modules
from data_processing_utils import setup_paths, read_configs, load_raw_data, apply_filters
from standard_report_logic import export_standard_report_excel, export_standard_report_pdf
from comparison_report_logic import apply_comparison_filters, export_comparison_report_excel, export_comparison_report_pdf

# --- Translation Dictionaries ---
translations = {
    "en": {
        "app_title": "TRIAC Time Report Generator",
        "sidebar_upload_excel_header": "Upload Excel Data File",
        "sidebar_upload_excel_label": "Select 'Time_report.xlsm' file",
        "sidebar_upload_logo_header": "Upload Logo (Optional)",
        "sidebar_upload_logo_label": "Select logo file (.png)",
        "excel_upload_success": "Excel file uploaded successfully!",
        "reading_processing_data": "Reading and processing data from Excel...",
        "error_empty_raw_data": "No raw data in Excel file or error reading. Please check 'Raw Data' sheet.",
        "info_data_read": "Read {num_rows} data entries from {num_projects} projects.",
        "error_processing_excel": "Error processing Excel file: {error}. Please ensure 'Time_report.xlsm' is correctly formatted and has 'Raw Data', 'Config_Year_Mode', 'Config_Project_Filter' sheets.",
        "info_upload_excel_prompt": "Please upload 'Time_report.xlsm' file to start.",
        "language_select_label": "Select Language",
        "tab_standard_report": "Standard Report",
        "tab_comparison_report": "Comparison Report",
        "tab_data_review": "Data Review",
        "tab_user_guide": "User Guide",
        "section_standard_report_title": "Generate Standard Report",
        "section_config_from_excel": "Configuration from Excel file:",
        "config_mode_label": "Mode",
        "config_year_label": "Year",
        "config_months_label": "Months",
        "all_label": "All",
        "section_projects_included": "Projects Included (from Config_Project_Filter):",
        "no_projects_selected_warning": "No projects selected in 'Config_Project_Filter' sheet or 'Include' column is not 'Yes'. Standard report may be empty.",
        "button_generate_standard_report": "Generate Standard Report",
        "generating_standard_report": "Generating standard report...",
        "download_excel_standard": "Download Standard Excel Report",
        "download_pdf_standard": "Download Standard PDF Report",
        "excel_report_success": "Standard Excel report generated successfully.",
        "pdf_report_success": "Standard PDF report generated successfully.",
        "error_generating_excel": "Error generating standard Excel report.",
        "error_generating_pdf": "Error generating standard PDF report.",
        "warning_no_data_standard": "No data filtered with current configurations for standard report.",
        "section_comparison_report_title": "Generate Comparison Report",
        "select_comparison_mode": "Select comparison mode:",
        "compare_mode_project_in_month": "Compare Projects in a Month",
        "compare_mode_project_in_year": "Compare Projects in a Year",
        "compare_mode_one_project_over_time": "Compare One Project Over Time (Months/Years)",
        "comparison_config_subheader": "Comparison Configuration:",
        "select_years_compare": "Select year(s) for comparison:",
        "select_months_compare": "Select month(s) for comparison (applies to certain modes only):",
        "select_projects_compare": "Select project(s) for comparison:",
        "validation_project_in_month": "Please select ONLY ONE year, ONLY ONE month, and at least TWO projects for this mode.",
        "validation_project_in_year": "Please select ONLY ONE year and at least TWO projects for this mode. Do not select months.",
        "validation_one_project_over_time": "Please select ONLY ONE project for this mode. Then, choose one year with multiple months, OR multiple years (do not select specific months).",
        "button_generate_comparison_report": "Generate Comparison Report",
        "error_fix_config": "Please fix configuration errors before generating the report.",
        "generating_comparison_report": "Generating comparison report...",
        "download_excel_comparison": "Download Comparison Excel Report",
        "download_pdf_comparison": "Download Comparison PDF Report",
        "excel_comparison_success": "Comparison Excel report generated successfully.",
        "pdf_comparison_success": "Comparison PDF report generated successfully.",
        "error_generating_comparison_excel": "Error generating comparison Excel report.",
        "error_generating_comparison_pdf": "Error generating comparison PDF report.",
        "warning_no_data_comparison": "No data available to generate comparison report: {msg}",
        "user_guide_title": "User Guide",
        "user_guide_content": """
        ## How to use the TRIAC Time Report Generator

        1.  **Upload Excel Data File:**
            * Browse and select your `Time_report.xlsm` file. This file should contain:
                * A sheet named `Raw Data` with columns like 'Date', 'Team member', 'Project Name', 'Task', 'Workcentre', 'Hou'.
                * A sheet named `Config_Year_Mode` with 'Key' (Mode, Year, Months) and 'Value' columns.
                * A sheet named `Config_Project_Filter` with 'Project Name' and 'Include' (Yes/No) columns.
        2.  **Upload Logo (Optional):**
            * Select your company logo in `.png` format. This will be included in the PDF reports.
        3.  **Standard Report:**
            * The app will automatically read configurations for the standard report from your uploaded Excel file.
            * Click "Generate Standard Report" to create an Excel and PDF report based on these configurations.
        4.  **Comparison Report:**
            * Select a comparison mode:
                * **Compare Projects in a Month:** Select one year, one month, and at least two projects to compare total hours.
                * **Compare Projects in a Year:** Select one year and at least two projects to compare monthly hours.
                * **Compare One Project Over Time (Months/Years):** Select one project. Then, either select one year with multiple months to see monthly trends, OR select multiple years to see yearly trends.
            * Click "Generate Comparison Report" to create the corresponding Excel and PDF reports.
        5.  **Data Review:**
            * View and filter the raw data loaded from your Excel file.
        6.  **Contact Support:**
            * If you encounter any issues, please contact support at `it.support@triac.vn`.
        """,
        "data_review_title": "Data Review",
        "data_review_description": "Review the raw data and filtered data loaded from your Excel file.",
        "raw_data_subheader": "Raw Data Overview:",
        "filtered_data_subheader": "Filtered Data Preview (based on Standard Report Config):",
        "email_auth_header": "Authentication",
        "enter_email_label": "Enter your TRIAC Email",
        "auth_button_label": "Login",
        "auth_success": "Authentication successful! Welcome, {email}.",
        "auth_invalid": "Invalid Email or unauthorized. Please contact IT support.",
        "contact_support_prompt": "If you encounter any issues, please contact IT support at: it.support@triac.vn",
        "excel_summary_sheet_name": "Summary",
        "task_col_header": "Task",
        "hours_col_header": "Hours",
        "chart_title_hours_by_project": "Total Hours by Project ({mode})",
        "chart_xaxis_project": "Project",
        "chart_yaxis_hours": "Hours",
        "chart_title_hours_by_task": "{project} - Hours by Task",
        "chart_xaxis_task": "Task",
        "chart_yaxis_task_hours": "Hours",
        "excel_config_info_sheet_name": "Config_Info",
        "config_years_label": "Year(s)",
        "config_projects_included_label": "Projects Included",
        "no_projects_selected_found": "No projects selected or found",
        "warning_empty_df_excel": "Warning: Filtered DataFrame is empty, no Excel report generated.",
        "warning_no_charts_pdf": "Warning: No charts generated to include in PDF. PDF might be empty.",
        "pdf_report_title_standard": "TRIAC TIME REPORT - STANDARD",
        "generated_on_label": "Generated on",
        "project_label": "Project",
        "no_charts_generated_message": "No charts generated for this report.",
        "error_missing_columns_excel": "Error: Missing columns in DataFrame. Cannot create Excel report.",
        "error_export_excel": "Error exporting standard Excel report: {error}",
        "error_create_pdf": "Error creating PDF report: {error}",
        "error_select_at_least_one_project": "Please select at least one project for comparison.",
        "warning_no_data_comparison_mode": "No data found for comparison mode: {mode} with current selections.",
        "invalid_comparison_mode": "Invalid comparison mode.",
        "project_name_col": "Project Name",
        "total_hours_col": "Total Hours",
        "total_label": "Total",
        "total_hours_col_prefix": "Total Hours for",
        "chart_title_project_comparison_short": "Project Hours Comparison",
        "chart_xaxis_month": "Month",
        "chart_title_project_month_comparison": "Project Hours Comparison by Month",
        "month_label": "Month",
        "hours_label": "Hours",
        "warning_no_month_cols_chart": "No month columns found to create chart.",
        "chart_title_project_over_months": "Project {project} Total Hours Over Months in {year}",
        "chart_xaxis_year": "Year",
        "chart_title_project_over_years": "Project {project} Total Hours Over Years",
        "warning_invalid_time_dimension_chart": "Invalid time dimension for chart categories in over time comparison mode.",
        "message_col": "Message",
        "no_data_to_display_message": "No data to display with selected filters.",
        "excel_comparison_sheet_name": "Comparison Report",
        "comparison_report_title": "COMPARISON REPORT",
        "config_projects_label": "Projects",
        "none_label": "None",
        "warning_no_chart_data_excel": "No sufficient data to plot comparison chart after removing total row.",
        "warning_no_chart_data_pdf": "DEBUG: df_plot is empty for mode '{mode}' after dropping 'Total'. Skipping chart creation.",
        "warning_no_month_cols_chart_pdf": "DEBUG: No month columns found for line chart in mode '{mode}'. Skipping chart creation.",
        "warning_invalid_time_dimension_chart_pdf": "DEBUG: Invalid columns for chart in mode '{mode}'. Skipping chart creation.",
        "warning_unknown_comparison_mode_chart": "DEBUG: Unknown comparison mode '{mode}'. Skipping chart creation.",
        "warning_invalid_time_dimension_chart_pdf_config": "Warning: Invalid time dimension configuration for PDF chart in over time comparison mode.",
        "warning_no_charts_comparison_pdf": "Warning: No charts generated to include in comparison PDF report. PDF might be empty.",
        "pdf_report_title_comparison": "TRIAC TIME REPORT - COMPARISON",
        "error_create_comparison_pdf": "Error creating comparison PDF report: {error}",
        "error_select_only_one_project": "Error: Internal - Please select ONLY ONE project for this mode.",
    },
    "vi": {
        "app_title": "TRIAC - Công Cụ Tạo Báo Cáo Thời Gian",
        "sidebar_upload_excel_header": "Tải lên File dữ liệu Excel",
        "sidebar_upload_excel_label": "Chọn file 'Time_report.xlsm'",
        "sidebar_upload_logo_header": "Tải lên Logo (Tùy chọn)",
        "sidebar_upload_logo_label": "Chọn file logo (.png)",
        "excel_upload_success": "File Excel đã được tải lên thành công!",
        "reading_processing_data": "Đang đọc và xử lý dữ liệu từ Excel...",
        "error_empty_raw_data": "Không có dữ liệu thô trong file Excel hoặc có lỗi khi đọc. Vui lòng kiểm tra sheet 'Raw Data'.",
        "info_data_read": "Đã đọc {num_rows} dòng dữ liệu từ {num_projects} dự án.",
        "error_processing_excel": "Lỗi khi xử lý file Excel: {error}. Vui lòng đảm bảo file 'Time_report.xlsm' đúng định dạng và có các sheet 'Raw Data', 'Config_Year_Mode', 'Config_Project_Filter'.",
        "info_upload_excel_prompt": "Vui lòng tải lên file 'Time_report.xlsm' để bắt đầu.",
        "language_select_label": "Chọn Ngôn ngữ",
        "tab_standard_report": "Báo cáo Tiêu chuẩn",
        "tab_comparison_report": "Báo cáo So sánh",
        "tab_data_review": "Xem lại Dữ liệu",
        "tab_user_guide": "Hướng dẫn sử dụng",
        "section_standard_report_title": "Tạo Báo cáo Tiêu chuẩn",
        "section_config_from_excel": "Cấu hình từ file Excel:",
        "config_mode_label": "Chế độ",
        "config_year_label": "Năm",
        "config_months_label": "Tháng",
        "all_label": "Tất cả",
        "section_projects_included": "Dự án được bao gồm (từ Config_Project_Filter):",
        "no_projects_selected_warning": "Không có dự án nào được chọn trong sheet 'Config_Project_Filter' hoặc cột 'Include' không có 'Yes'. Báo cáo tiêu chuẩn có thể trống.",
        "button_generate_standard_report": "Tạo Báo cáo Tiêu chuẩn",
        "generating_standard_report": "Đang tạo báo cáo tiêu chuẩn...",
        "download_excel_standard": "Tải xuống Báo cáo Excel Tiêu chuẩn",
        "download_pdf_standard": "Tải xuống Báo cáo PDF Tiêu chuẩn",
        "excel_report_success": "Đã tạo báo cáo Excel tiêu chuẩn.",
        "pdf_report_success": "Đã tạo báo cáo PDF tiêu chuẩn.",
        "error_generating_excel": "Có lỗi khi tạo báo cáo Excel tiêu chuẩn.",
        "error_generating_pdf": "Có lỗi khi tạo báo cáo PDF tiêu chuẩn.",
        "warning_no_data_standard": "Không có dữ liệu nào được lọc với các cấu hình hiện tại để tạo báo cáo tiêu chuẩn.",
        "section_comparison_report_title": "Tạo Báo cáo So sánh",
        "select_comparison_mode": "Chọn chế độ so sánh:",
        "compare_mode_project_in_month": "So Sánh Dự Án Trong Một Tháng",
        "compare_mode_project_in_year": "So Sánh Dự Án Trong Một Năm",
        "compare_mode_one_project_over_time": "So Sánh Một Dự Án Qua Các Tháng/Năm",
        "comparison_config_subheader": "Cấu hình So sánh:",
        "select_years_compare": "Chọn năm(các năm) để so sánh:",
        "select_months_compare": "Chọn tháng(các tháng) để so sánh (chỉ áp dụng cho một số chế độ):",
        "select_projects_compare": "Chọn dự án(các dự án) để so sánh:",
        "validation_project_in_month": "Vui lòng chọn CHỈ MỘT năm, CHỈ MỘT tháng và ít nhất HAI dự án cho chế độ này.",
        "validation_project_in_year": "Vui lòng chọn CHỈ MỘT năm và ít nhất HAI dự án cho chế độ này. Không chọn tháng.",
        "validation_one_project_over_time": "Vui lòng chọn CHỈ MỘT dự án cho chế độ này. Sau đó, chọn một năm với nhiều tháng, HOẶC nhiều năm (không chọn tháng cụ thể).",
        "button_generate_comparison_report": "Tạo Báo cáo So sánh",
        "error_fix_config": "Vui lòng sửa các lỗi cấu hình trước khi tạo báo cáo.",
        "generating_comparison_report": "Đang tạo báo cáo so sánh...",
        "download_excel_comparison": "Tải xuống Báo cáo Excel So sánh",
        "download_pdf_comparison": "Tải xuống Báo cáo PDF So sánh",
        "excel_comparison_success": "Đã tạo báo cáo Excel so sánh.",
        "pdf_comparison_success": "Đã tạo báo cáo PDF so sánh.",
        "error_generating_comparison_excel": "Có lỗi khi tạo báo cáo Excel so sánh.",
        "error_generating_comparison_pdf": "Có lỗi khi tạo báo cáo PDF so sánh.",
        "warning_no_data_comparison": "Không có dữ liệu nào để tạo báo cáo so sánh: {msg}",
        "user_guide_title": "Hướng dẫn sử dụng",
        "user_guide_content": """
        ## Hướng dẫn sử dụng Công cụ tạo báo cáo thời gian TRIAC

        1.  **Tải lên file dữ liệu Excel:**
            * Duyệt và chọn file `Time_report.xlsm` của bạn. File này phải chứa:
                * Một sheet tên `Raw Data` với các cột như 'Date', 'Team member', 'Project Name', 'Task', 'Workcentre', 'Hou'.
                * Một sheet tên `Config_Year_Mode` với các cột 'Key' (Mode, Year, Months) và 'Value'.
                * Một sheet tên `Config_Project_Filter` với các cột 'Project Name' và 'Include' (Yes/No).
        2.  **Tải lên Logo (Tùy chọn):**
            * Chọn logo công ty của bạn ở định dạng `.png`. Logo này sẽ được đưa vào các báo cáo PDF.
        3.  **Báo cáo Tiêu chuẩn:**
            * Ứng dụng sẽ tự động đọc các cấu hình cho báo cáo tiêu chuẩn từ file Excel bạn đã tải lên.
            * Nhấn "Tạo Báo cáo Tiêu chuẩn" để tạo báo cáo Excel và PDF dựa trên các cấu hình này.
        4.  **Báo cáo So sánh:**
            * Chọn một chế độ so sánh:
                * **So Sánh Dự Án Trong Một Tháng:** Chọn một năm, một tháng và ít nhất hai dự án để so sánh tổng số giờ.
                * **So Sánh Dự Án Trong Một Năm:** Chọn một năm và ít nhất hai dự án để so sánh số giờ hàng tháng.
                * **So Sánh Một Dự Án Qua Các Tháng/Năm:** Chọn một dự án. Sau đó, chọn một năm với nhiều tháng để xem xu hướng hàng tháng, HOẶC chọn nhiều năm để xem xu hướng hàng năm.
            * Nhấn "Tạo Báo cáo So sánh" để tạo các báo cáo Excel và PDF tương ứng.
        5.  **Xem lại Dữ liệu:**
            * Xem và lọc dữ liệu thô đã được tải từ file Excel của bạn.
        6.  **Liên hệ Hỗ trợ:**
            * Nếu bạn gặp bất kỳ vấn đề nào, vui lòng liên hệ bộ phận hỗ trợ tại `it.support@triac.vn`.
        """,
        "data_review_title": "Xem lại Dữ liệu",
        "data_review_description": "Xem lại dữ liệu thô và dữ liệu đã lọc được tải từ file Excel của bạn.",
        "raw_data_subheader": "Tổng quan Dữ liệu thô:",
        "filtered_data_subheader": "Xem trước Dữ liệu đã lọc (dựa trên Cấu hình Báo cáo Tiêu chuẩn):",
        "email_auth_header": "Xác thực",
        "enter_email_label": "Nhập Email TRIAC của bạn",
        "auth_button_label": "Đăng nhập",
        "auth_success": "Xác thực thành công! Chào mừng, {email}.",
        "auth_invalid": "Email không hợp lệ hoặc không được cấp quyền. Vui lòng liên hệ bộ phận hỗ trợ IT.",
        "contact_support_prompt": "Nếu bạn gặp bất kỳ vấn đề nào, vui lòng liên hệ bộ phận hỗ trợ IT tại: it.support@triac.vn",
        "excel_summary_sheet_name": "Tóm tắt",
        "task_col_header": "Nhiệm vụ",
        "hours_col_header": "Giờ",
        "chart_title_hours_by_project": "Tổng số giờ theo Dự án ({mode})",
        "chart_xaxis_project": "Dự án",
        "chart_yaxis_hours": "Giờ",
        "chart_title_hours_by_task": "{project} - Giờ theo Nhiệm vụ",
        "chart_xaxis_task": "Nhiệm vụ",
        "chart_yaxis_task_hours": "Giờ",
        "excel_config_info_sheet_name": "Thông tin Cấu hình",
        "config_years_label": "Năm(các năm)",
        "config_projects_included_label": "Dự án đã bao gồm",
        "no_projects_selected_found": "Không có dự án nào được chọn hoặc tìm thấy",
        "warning_empty_df_excel": "Cảnh báo: DataFrame đã lọc trống, không có báo cáo Excel nào được tạo.",
        "warning_no_charts_pdf": "Cảnh báo: Không có biểu đồ nào được tạo để đưa vào PDF. PDF có thể trống.",
        "pdf_report_title_standard": "BÁO CÁO THỜI GIAN TRIAC - TIÊU CHUẨN",
        "generated_on_label": "Ngày tạo",
        "project_label": "Dự án",
        "no_charts_generated_message": "Không có biểu đồ nào được tạo cho báo cáo này.",
        "error_missing_columns_excel": "Lỗi: Thiếu cột trong DataFrame. Không thể tạo báo cáo Excel.",
        "error_export_excel": "Lỗi khi xuất báo cáo Excel tiêu chuẩn: {error}",
        "error_create_pdf": "Lỗi khi tạo báo cáo PDF: {error}",
        "error_select_at_least_one_project": "Vui lòng chọn ít nhất một dự án để so sánh.",
        "warning_no_data_comparison_mode": "Không tìm thấy dữ liệu cho chế độ so sánh: {mode} với các lựa chọn hiện tại.",
        "invalid_comparison_mode": "Chế độ so sánh không hợp lệ.",
        "project_name_col": "Tên Dự án",
        "total_hours_col": "Tổng số giờ",
        "total_label": "Tổng",
        "total_hours_col_prefix": "Tổng số giờ cho",
        "chart_title_project_comparison_short": "So sánh giờ theo Dự án",
        "chart_xaxis_month": "Tháng",
        "chart_title_project_month_comparison": "So sánh giờ theo Dự án và Tháng",
        "month_label": "Tháng",
        "hours_label": "Giờ",
        "warning_no_month_cols_chart": "Không tìm thấy cột tháng để tạo biểu đồ.",
        "chart_title_project_over_months": "Tổng giờ dự án {project} qua các tháng trong năm {year}",
        "chart_xaxis_year": "Năm",
        "chart_title_project_over_years": "Tổng giờ dự án {project} qua các năm",
        "warning_invalid_time_dimension_chart": "Không tìm thấy kích thước thời gian hợp lệ cho các danh mục biểu đồ trong chế độ so sánh qua tháng/năm.",
        "message_col": "Thông báo",
        "no_data_to_display_message": "Không có dữ liệu để hiển thị với các bộ lọc đã chọn.",
        "excel_comparison_sheet_name": "Báo cáo So sánh",
        "comparison_report_title": "BÁO CÁO SO SÁNH",
        "config_projects_label": "Dự án",
        "none_label": "Không có",
        "warning_no_chart_data_excel": "Không đủ dữ liệu để vẽ biểu đồ so sánh sau khi loại bỏ hàng tổng.",
        "warning_no_chart_data_pdf": "DEBUG: df_plot trống cho chế độ '{mode}' sau khi bỏ hàng 'Tổng'. Bỏ qua tạo biểu đồ.",
        "warning_no_month_cols_chart_pdf": "DEBUG: Không tìm thấy cột tháng cho biểu đồ đường trong chế độ '{mode}'. Bỏ qua tạo biểu đồ.",
        "warning_invalid_time_dimension_chart_pdf": "DEBUG: Cột không hợp lệ cho biểu đồ trong chế độ '{mode}'. Bỏ qua tạo biểu đồ.",
        "warning_unknown_comparison_mode_chart": "DEBUG: Chế độ so sánh '{mode}' không xác định. Bỏ qua tạo biểu đồ.",
        "warning_invalid_time_dimension_chart_pdf_config": "Cảnh báo: Cấu hình kích thước thời gian không hợp lệ cho biểu đồ PDF trong chế độ so sánh qua thời gian.",
        "warning_no_charts_comparison_pdf": "Cảnh báo: Không có biểu đồ nào được tạo để đưa vào PDF báo cáo so sánh. PDF có thể trống.",
        "pdf_report_title_comparison": "BÁO CÁO THỜI GIAN TRIAC - SO SÁNH",
        "error_create_comparison_pdf": "Lỗi khi tạo báo cáo PDF so sánh: {error}",
        "error_select_only_one_project": "Lỗi: Nội bộ - Vui lòng chọn CHỈ MỘT dự án cho chế độ này.",
    }
}

# --- Initial Setup ---
st.set_page_config(layout="wide", page_title=translations["en"]["app_title"]) # Default title is English

# Initialize language in session state
if 'language' not in st.session_state:
    st.session_state.language = 'en' # Set default language to English

current_translations = translations[st.session_state.language]

st.title(current_translations["app_title"])

# --- Hàm đọc email từ file CSV ---
# Đặt hàm này ở đây, trước khi bạn cần dùng AUTHORIZED_EMAILS
def load_authorized_emails(csv_file_path):
    # Sử dụng st.cache_data để cache kết quả đọc file, tránh đọc lại mỗi lần rerun
    @st.cache_data(ttl=300) # Cache trong 300 giây (5 phút), hoặc điều chỉnh
    def _load_emails(path):
        if not os.path.exists(path):
            st.error(f"Lỗi: File invited_emails.csv không tìm thấy tại {path}. Vui lòng đảm bảo file tồn tại.")
            return []
        try:
            # Đọc CSV không có header, và đặt tên cột là 'email'
            df_emails = pd.read_csv(path, header=None, names=['email'])
            # Chuyển tất cả email thành chữ thường và loại bỏ khoảng trắng thừa
            return [email.strip().lower() for email in df_emails['email'].tolist() if pd.notna(email)]
        except Exception as e:
            st.error(f"Lỗi khi đọc file invited_emails.csv: {e}. Vui lòng kiểm tra định dạng file.")
            return []
    return _load_emails(csv_file_path)

# --- Xác định đường dẫn và tải danh sách email ---
# Đảm bảo file 'invited_emails.csv' nằm cùng cấp với 'app.py'
csv_path_for_auth = "invited_emails.csv"
AUTHORIZED_EMAILS = load_authorized_emails(csv_path_for_auth)

# --- Authentication (Phần này sẽ được giữ lại, nhưng AUTHORIZED_EMAILS đã được tải động) ---
def authenticate_user(email):
    # Thêm debug logs để kiểm tra
    # st.write(f"DEBUG: Email nhập vào (raw): '{email}'")
    processed_email = email.lower().strip()
    # st.write(f"DEBUG: Email đã xử lý: '{processed_email}'")
    # st.write(f"DEBUG: Danh sách email được phép: {AUTHORIZED_EMAILS}")
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

# --- Language Selection (Move to sidebar and set default) ---
st.sidebar.header(current_translations["language_select_label"])
lang_options = {"English": "en", "Tiếng Việt": "vi"}
selected_lang_name = st.sidebar.selectbox(
    current_translations["language_select_label"],
    options=list(lang_options.keys()),
    index=list(lang_options.keys()).index("English") if st.session_state.language == 'en' else list(lang_options.keys()).index("Tiếng Việt") # Default to English
)
selected_lang_code = lang_options[selected_lang_name]

if selected_lang_code != st.session_state.language:
    st.session_state.language = selected_lang_code
    st.rerun() # Rerun to apply language change

# Update translations based on selected language
current_translations = translations[st.session_state.language]

# --- File Uploads ---
st.sidebar.header(current_translations["sidebar_upload_excel_header"])
uploaded_file = st.sidebar.file_uploader(current_translations["sidebar_upload_excel_label"], type=["xlsm"])

st.sidebar.header(current_translations["sidebar_upload_logo_header"])
uploaded_logo = st.sidebar.file_uploader(current_translations["sidebar_upload_logo_label"], type=["png"])

# Cache raw data and configs to avoid re-reading on every rerun
@st.cache_data(show_spinner=False)
def get_processed_data_from_upload(file_buffer):
    temp_excel_path = os.path.join(tempfile.gettempdir(), "uploaded_Time_report.xlsm")
    
    with open(temp_excel_path, "wb") as f:
        f.write(file_buffer.getbuffer())
    
    config_data = read_configs(temp_excel_path)
    raw_df = load_raw_data(temp_excel_path)
    
    # Do NOT delete temp_excel_path here, as read_configs and load_raw_data might be cached
    # and expect the file to be present on subsequent calls within the same session.
    # The clean-up of temp files will be handled by the OS or Streamlit's internal mechanisms
    # for temporary directories. For explicit cleanup after the session, we'd need a different strategy.
    
    return config_data, raw_df, temp_excel_path

if uploaded_file is not None:
    st.success(current_translations["excel_upload_success"])
    
    with st.spinner(current_translations["reading_processing_data"]):
        try:
            config_data, raw_df, temp_excel_path = get_processed_data_from_upload(uploaded_file)
            
            if raw_df.empty:
                st.error(current_translations["error_empty_raw_data"])
                st.stop()

            st.session_state['raw_df'] = raw_df
            st.session_state['config_data'] = config_data
            st.session_state['temp_excel_path'] = temp_excel_path # Store temp path for later use if needed
            
            all_project_names = sorted(raw_df['Project name'].unique().tolist())
            all_years_in_data = sorted(raw_df['Year'].unique().tolist())
            all_months_in_data = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

            st.session_state['all_project_names'] = all_project_names
            st.session_state['all_years_in_data'] = all_years_in_data
            st.session_state['all_months_in_data'] = all_months_in_data

            st.sidebar.info(current_translations["info_data_read"].format(num_rows=len(raw_df), num_projects=len(all_project_names)))
        
        except Exception as e:
            st.error(current_translations["error_processing_excel"].format(error=e))
            st.stop()
else:
    st.info(current_translations["info_upload_excel_prompt"])
    st.stop()

logo_path = None
if uploaded_logo is not None:
    logo_path = os.path.join(tempfile.gettempdir(), uploaded_logo.name)
    with open(logo_path, "wb") as f:
        f.write(uploaded_logo.getbuffer())
    st.sidebar.success(current_translations["sidebar_upload_logo_header"].replace("(Optional)", "").strip() + " " + current_translations["excel_upload_success"].split(" ")[-1]) # Reusing success message part

# Get data from session_state
raw_df = st.session_state.get('raw_df')
config_data = st.session_state.get('config_data')
all_project_names = st.session_state.get('all_project_names', [])
all_years_in_data = st.session_state.get('all_years_in_data', [])
all_months_in_data = st.session_state.get('all_months_in_data', [])


# --- Main Tabs ---
tab_standard, tab_comparison, tab_data_review, tab_user_guide = st.tabs([
    current_translations["tab_standard_report"],
    current_translations["tab_comparison_report"],
    current_translations["tab_data_review"],
    current_translations["tab_user_guide"]
])

with tab_standard:
    st.header(current_translations["section_standard_report_title"])

    st.subheader(current_translations["section_config_from_excel"])
    st.write(f"**{current_translations['config_mode_label']}:** `{config_data['mode'].capitalize()}`")
    st.write(f"**{current_translations['config_year_label']}:** `{config_data['year']}`")
    st.write(f"**{current_translations['config_months_label']}:** `{', '.join(config_data['months']) if config_data['months'] else current_translations['all_label']}`")
    
    st.subheader(current_translations["section_projects_included"])
    included_projects_from_config = config_data['project_filter_df'][
        config_data['project_filter_df']['Include'].astype(str).str.lower() == 'yes'
    ]['Project Name'].tolist()
    
    if included_projects_from_config:
        st.write(f"{current_translations['project_label']}: `{', '.join(included_projects_from_config)}`")
    else:
        st.warning(current_translations["no_projects_selected_warning"])

    if st.button(current_translations["button_generate_standard_report"]):
        with st.spinner(current_translations["generating_standard_report"]):
            try:
                standard_report_config = config_data.copy()
                standard_report_config['years'] = [standard_report_config['year']]
                
                df_standard_filtered = apply_filters(raw_df, standard_report_config)

                if not df_standard_filtered.empty:
                    today = datetime.datetime.today().strftime('%Y%m%d')
                    output_excel_file = os.path.join(tempfile.gettempdir(), f"Time_report_Standard_{today}.xlsx")
                    output_pdf_file = os.path.join(tempfile.gettempdir(), f"Time_report_Standard_{today}.pdf")

                    excel_success = export_standard_report_excel(df_standard_filtered, standard_report_config, output_excel_file, current_translations)
                    if excel_success:
                        with open(output_excel_file, "rb") as f:
                            st.download_button(
                                label=current_translations["download_excel_standard"],
                                data=f.read(),
                                file_name=f"Time_report_Standard_{today}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        st.success(current_translations["excel_report_success"])
                    else:
                        st.error(current_translations["error_generating_excel"])

                    pdf_success = export_standard_report_pdf(df_standard_filtered, standard_report_config, output_pdf_file, logo_path, current_translations)
                    if pdf_success:
                        with open(output_pdf_file, "rb") as f:
                            st.download_button(
                                label=current_translations["download_pdf_standard"],
                                data=f.read(),
                                file_name=f"Time_report_Standard_{today}.pdf",
                                mime="application/pdf"
                            )
                        st.success(current_translations["pdf_report_success"])
                    else:
                        st.error(current_translations["error_generating_pdf"])
                    
                    if os.path.exists(output_excel_file): os.remove(output_excel_file)
                    if os.path.exists(output_pdf_file): os.remove(output_pdf_file)

                else:
                    st.warning(current_translations["warning_no_data_standard"])
            except Exception as e:
                st.error(current_translations["error_generating_excel"].format(error=e))
                st.exception(e)

with tab_comparison:
    st.header(current_translations["section_comparison_report_title"])

    comparison_mode_options = [
        current_translations["compare_mode_project_in_month"],
        current_translations["compare_mode_project_in_year"],
        current_translations["compare_mode_one_project_over_time"]
    ]
    selected_comparison_mode = st.selectbox(current_translations["select_comparison_mode"], comparison_mode_options)

    st.subheader(current_translations["comparison_config_subheader"])
    
    comparison_years = st.multiselect(
        current_translations["select_years_compare"], 
        options=all_years_in_data, 
        default=all_years_in_data[-1:] if all_years_in_data else []
    )

    comparison_months = st.multiselect(
        current_translations["select_months_compare"], 
        options=all_months_in_data, 
        default=[]
    )

    comparison_projects = st.multiselect(
        current_translations["select_projects_compare"], 
        options=all_project_names, 
        default=all_project_names[:2] if len(all_project_names) >= 2 else all_project_names
    )
    
    validation_message = ""
    if selected_comparison_mode == current_translations["compare_mode_project_in_month"]:
        if len(comparison_years) != 1 or len(comparison_months) != 1 or len(comparison_projects) < 2:
            validation_message = current_translations["validation_project_in_month"]
    elif selected_comparison_mode == current_translations["compare_mode_project_in_year"]:
        if len(comparison_years) != 1 or len(comparison_projects) < 2 or comparison_months:
            validation_message = current_translations["validation_project_in_year"]
    elif selected_comparison_mode == current_translations["compare_mode_one_project_over_time"]:
        if len(comparison_projects) != 1 or not ((len(comparison_years) == 1 and len(comparison_months) >= 2) or (len(comparison_years) >= 2 and not comparison_months)):
            validation_message = current_translations["validation_one_project_over_time"]

    if validation_message:
        st.warning(validation_message)

    if st.button(current_translations["button_generate_comparison_report"]):
        if validation_message:
            st.error(current_translations["error_fix_config"])
        else:
            with st.spinner(current_translations["generating_comparison_report"]):
                try:
                    comp_config = {
                        'years': comparison_years,
                        'months': comparison_months,
                        'selected_projects': comparison_projects
                    }
                    
                    df_comparison, msg = apply_comparison_filters(raw_df, comp_config, selected_comparison_mode, current_translations)
                    
                    if not df_comparison.empty:
                        today = datetime.datetime.today().strftime('%Y%m%d')
                        comp_output_excel_file = os.path.join(tempfile.gettempdir(), f"Time_report_Comparison_{today}.xlsx")
                        comp_output_pdf_file = os.path.join(tempfile.gettempdir(), f"Time_report_Comparison_{today}.pdf")

                        excel_success_comp = export_comparison_report_excel(df_comparison, comp_config, comp_output_excel_file, selected_comparison_mode, current_translations)
                        if excel_success_comp:
                            with open(comp_output_excel_file, "rb") as f:
                                st.download_button(
                                    label=current_translations["download_excel_comparison"],
                                    data=f.read(),
                                    file_name=f"Time_report_Comparison_{today}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            st.success(current_translations["excel_comparison_success"])
                        else:
                            st.error(current_translations["error_generating_comparison_excel"])

                        pdf_success_comp = export_comparison_report_pdf(df_comparison, comp_config, comp_output_pdf_file, logo_path, selected_comparison_mode, current_translations)
                        if pdf_success_comp:
                            with open(comp_output_pdf_file, "rb") as f:
                                st.download_button(
                                    label=current_translations["download_pdf_comparison"],
                                    data=f.read(),
                                    file_name=f"Time_report_Comparison_{today}.pdf",
                                    mime="application/pdf"
                                )
                            st.success(current_translations["pdf_comparison_success"])
                        else:
                            st.error(current_translations["error_generating_comparison_pdf"])

                        if os.path.exists(comp_output_excel_file): os.remove(comp_output_excel_file)
                        if os.path.exists(comp_output_pdf_file): os.remove(comp_output_pdf_file)
                    else:
                        st.warning(current_translations["warning_no_data_comparison"].format(msg=msg))
                except Exception as e:
                    st.error(current_translations["error_generating_comparison_excel"].format(error=e))
                    st.exception(e)

with tab_data_review:
    st.header(current_translations["data_review_title"])
    st.write(current_translations["data_review_description"])

    st.subheader(current_translations["raw_data_subheader"])
    st.dataframe(raw_df)

    st.subheader(current_translations["filtered_data_subheader"])
    # Show filtered data based on the standard report's configuration for consistency
    standard_report_config_for_review = config_data.copy()
    standard_report_config_for_review['years'] = [standard_report_config_for_review['year']]
    df_filtered_for_review = apply_filters(raw_df, standard_report_config_for_review)
    if not df_filtered_for_review.empty:
        st.dataframe(df_filtered_for_review)
    else:
        st.info(current_translations["warning_no_data_standard"])


with tab_user_guide:
    st.header(current_translations["user_guide_title"])
    st.markdown(current_translations["user_guide_content"])
    st.markdown(current_translations["contact_support_prompt"])
