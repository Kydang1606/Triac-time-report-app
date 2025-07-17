import core_logic as cl
import pandas as pd
import os

def run_standard_report():
    """
    Hướng dẫn người dùng tạo báo cáo thời gian tiêu chuẩn.
    """
    print("\n--- Đang tạo Báo cáo thời gian tiêu chuẩn ---")
    paths = cl.setup_paths()
    config = cl.read_configs(paths['template_file'])

    if not config:
        print("Không thể tải cấu hình. Đang thoát tạo báo cáo tiêu chuẩn.")
        return

    print(f"Cấu hình đã tải: Chế độ={config['mode'].capitalize()}, Năm={config['year']}, Tháng={', '.join(config['months']) if config['months'] else 'Tất cả'}")
    
    raw_data_df = cl.load_raw_data(paths['template_file'])
    if raw_data_df.empty:
        print("Không có dữ liệu thô nào được tải. Đang thoát tạo báo cáo tiêu chuẩn.")
        return

    # Lọc dự án dựa trên cột 'Include' trong Config_Project_Filter
    if not config['project_filter_df'].empty:
        included_projects_df = config['project_filter_df'][config['project_filter_df']['Include'].astype(str).str.lower() == 'yes']
        if included_projects_df.empty:
            print("Không có dự án nào được chọn để đưa vào báo cáo. Vui lòng cập nhật sheet 'Config_Project_Filter'.")
            return
        # Cập nhật cấu hình để chỉ bao gồm các dự án 'yes' để lọc
        config['project_filter_df'] = included_projects_df
    else:
        print("Không tìm thấy cấu hình bộ lọc dự án. Vui lòng đảm bảo sheet 'Config_Project_Filter' tồn tại và được cấu hình.")
        return

    filtered_df = cl.apply_filters(raw_data_df, config)

    if filtered_df.empty:
        print("Không có dữ liệu sau khi áp dụng bộ lọc. Báo cáo sẽ không được tạo.")
        return

    print(f"Dữ liệu đã lọc cho {len(filtered_df['Project name'].unique())} dự án.")

    excel_success = cl.export_report(filtered_df, config, paths['output_file'])
    if excel_success:
        print(f"Báo cáo Excel tiêu chuẩn được tạo thành công: {paths['output_file']}")
    else:
        print("Không thể tạo báo cáo Excel tiêu chuẩn.")

    pdf_success = cl.export_pdf_report(filtered_df, config, paths['pdf_report'], paths['logo_path'])
    if pdf_success:
        print(f"Báo cáo PDF tiêu chuẩn được tạo thành công: {paths['pdf_report']}")
    else:
        print("Không thể tạo báo cáo PDF tiêu chuẩn.")

def run_comparison_report():
    """
    Hướng dẫn người dùng tạo báo cáo thời gian so sánh.
    """
    print("\n--- Đang tạo Báo cáo thời gian so sánh ---")
    paths = cl.setup_paths()
    template_file = paths['template_file']
    logo_path = paths['logo_path']

    try:
        raw_data_df = cl.load_raw_data(template_file)
        if raw_data_df.empty:
            print("Không có dữ liệu thô nào được tải. Đang thoát tạo báo cáo so sánh.")
            return

        all_projects = raw_data_df['Project name'].unique().tolist()
        all_years = sorted(raw_data_df['Year'].unique().tolist())
        all_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

        print("\nChọn Chế độ so sánh:")
        print("1. So sánh các dự án trong một tháng")
        print("2. So sánh các dự án trong một năm")
        print("3. So sánh một dự án theo thời gian (Tháng/Năm)")
        
        mode_choice = input("Nhập lựa chọn (1/2/3): ").strip()
        
        comparison_config = {}
        selected_years = []
        selected_months = []
        selected_projects = []
        comparison_mode_name = ""

        if mode_choice == '1':
            comparison_mode_name = "So sánh các dự án trong một tháng"
            print("\n--- So sánh các dự án trong một tháng ---")
            print(f"Các năm có sẵn: {', '.join(map(str, all_years))}")
            year_input = input("Nhập MỘT năm (ví dụ, 2023): ").strip()
            if year_input.isdigit():
                selected_years.append(int(year_input))
            else:
                print("Đầu vào năm không hợp lệ.")
                return

            print(f"Các tháng có sẵn: {', '.join(all_months)}")
            month_input = input("Nhập MỘT tháng (ví dụ, January): ").strip().capitalize()
            if month_input in all_months:
                selected_months.append(month_input)
            else:
                print("Đầu vào tháng không hợp lệ.")
                return

            print(f"Các dự án có sẵn: {', '.join(all_projects)}")
            projects_input = input("Nhập HAI hoặc nhiều tên dự án, cách nhau bằng dấu phẩy (ví dụ, Dự án A, Dự án B): ").strip()
            selected_projects = [p.strip() for p in projects_input.split(',') if p.strip() in all_projects]
            if len(selected_projects) < 2:
                print("Vui lòng chọn ít nhất hai dự án hợp lệ.")
                return

        elif mode_choice == '2':
            comparison_mode_name = "So sánh các dự án trong một năm"
            print("\n--- So sánh các dự án trong một năm ---")
            print(f"Các năm có sẵn: {', '.join(map(str, all_years))}")
            year_input = input("Nhập MỘT năm (ví dụ, 2023): ").strip()
            if year_input.isdigit():
                selected_years.append(int(year_input))
            else:
                print("Đầu vào năm không hợp lệ.")
                return

            print(f"Các dự án có sẵn: {', '.join(all_projects)}")
            projects_input = input("Nhập HAI hoặc nhiều tên dự án, cách nhau bằng dấu phẩy (ví dụ, Dự án A, Dự án B): ").strip()
            selected_projects = [p.strip() for p in projects_input.split(',') if p.strip() in all_projects]
            if len(selected_projects) < 2:
                print("Vui lòng chọn ít nhất hai dự án hợp lệ.")
                return

        elif mode_choice == '3':
            comparison_mode_name = "So sánh một dự án theo thời gian (Tháng/Năm)"
            print("\n--- So sánh một dự án theo thời gian (Tháng/Năm) ---")
            print(f"Các dự án có sẵn: {', '.join(all_projects)}")
            project_input = input("Nhập MỘT tên dự án (ví dụ, Dự án X): ").strip()
            if project_input in all_projects:
                selected_projects.append(project_input)
            else:
                print("Tên dự án không hợp lệ.")
                return

            time_mode_choice = input("So sánh theo (M)tháng trong một năm HOẶC (Y)năm theo thời gian? (M/Y): ").strip().upper()
            if time_mode_choice == 'M':
                print(f"Các năm có sẵn: {', '.join(map(str, all_years))}")
                year_input = input("Nhập MỘT năm để so sánh các tháng trong đó (ví dụ, 2023): ").strip()
                if year_input.isdigit():
                    selected_years.append(int(year_input))
                else:
                    print("Đầu vào năm không hợp lệ.")
                    return
                # Tất cả các tháng sẽ được xem xét nếu không được chỉ định cho dự án này trong năm này
                selected_months = all_months 
            elif time_mode_choice == 'Y':
                print(f"Các năm có sẵn: {', '.join(map(str, all_years))}")
                years_input = input("Nhập HAI hoặc nhiều năm, cách nhau bằng dấu phẩy (ví dụ, 2022, 2023): ").strip()
                selected_years = [int(y.strip()) for y in years_input.split(',') if y.strip().isdigit() and int(y.strip()) in all_years]
                if len(selected_years) < 2:
                    print("Vui lòng chọn ít nhất hai năm hợp lệ.")
                    return
            else:
                print("Lựa chọn so sánh thời gian không hợp lệ.")
                return
        else:
            print("Lựa chọn không hợp lệ. Đang thoát tạo báo cáo so sánh.")
            return

        comparison_config = {
            'years': selected_years,
            'months': selected_months,
            'selected_projects': selected_projects
        }

        df_comparison, msg = cl.apply_comparison_filters(raw_data_df, comparison_config, comparison_mode_name)
        if df_comparison.empty:
            print(f"Lỗi khi áp dụng bộ lọc so sánh: {msg}")
            return

        excel_success = cl.export_comparison_report(df_comparison, comparison_config, paths['comparison_output_file'], comparison_mode_name)
        if excel_success:
            print(f"Báo cáo Excel so sánh được tạo thành công: {paths['comparison_output_file']}")
        else:
            print("Không thể tạo báo cáo Excel so sánh.")

        pdf_success = cl.export_comparison_pdf_report(df_comparison, comparison_config, paths['comparison_pdf_report'], comparison_mode_name, logo_path)
        if pdf_success:
            print(f"Báo cáo PDF so sánh được tạo thành công: {paths['comparison_pdf_report']}")
        else:
            print("Không thể tạo báo cáo PDF so sánh.")

    except Exception as e:
        print(f"Đã xảy ra lỗi trong quá trình tạo báo cáo so sánh: {e}")

def main():
    """
    Hàm chính để cung cấp giao diện dòng lệnh cho ứng dụng.
    """
    print("Chào mừng đến với Công cụ tạo báo cáo thời gian TRIAC!")
    while True:
        print("\n--- Menu chính ---")
        print("1. Tạo báo cáo thời gian tiêu chuẩn (Excel & PDF)")
        print("2. Tạo báo cáo thời gian so sánh (Excel & PDF)")
        print("3. Thoát")
        
        choice = input("Nhập lựa chọn của bạn (1/2/3): ").strip()

        if choice == '1':
            run_standard_report()
        elif choice == '2':
            run_comparison_report()
        elif choice == '3':
            print("Đang thoát ứng dụng. Tạm biệt!")
            break
        else:
            print("Lựa chọn không hợp lệ. Vui lòng nhập 1, 2 hoặc 3.")

if __name__ == "__main__":
    main()
