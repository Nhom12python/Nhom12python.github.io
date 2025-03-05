import PySimpleGUI as sg
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
import sys

# Sử dụng backend phù hợp
if sys.platform == 'win32':
    matplotlib.use('TkAgg')  # Sử dụng TkAgg trên Windows
else:
    matplotlib.use('Agg')  # Sử dụng Agg trên hệ thống không có giao diện đồ họa

# Đọc dữ liệu từ file Excel
try:
    df = pd.read_excel("data_demo.xlsx")
    table_data = df.values.tolist()
    table_header = df.columns.values.tolist()
except Exception as e:
    sg.popup(f"Lỗi khi đọc file Excel: {str(e)}")
    sys.exit()

# Thiết lập theme và layout giao diện
sg.theme("Green")
layout = [
    [sg.Text("THÔNG TIN QUẢN LÝ CHI TIÊU QUỸ LỚP 9/3", text_color="white", justification="center", font=("Arial", 14, "bold"), expand_x=True)],
    [sg.Text("Số ID:", size=(10, 1)), sg.Input(key="Số ID", size=47), sg.Text('Họ Tên:', size=(6, 1)), sg.Input(key="Họ Tên", size=(59, 2))],
    [sg.Text('Giới tính', size=(10, 1)), sg.Combo(['Nam', 'Nữ'], key='Giới tính', size=(45, 2))],
    [sg.Text('Danh mục', size=(10,1)), sg.Combo(['Văn nghệ trường', 'Trang trí trường', 'Phí sinh hoạt CLB', 'In tài liệu', 'Tổ chức sự kiện', 'Khác'], key='Danh mục', size=(45, 5)), 
     sg.Text("Ngày:", size=(6, 1)), sg.Input(key="Ngày", size=(59, 1)), sg.CalendarButton("Chọn ngày", target="Ngày", format="%d/%m/%Y")],
    [sg.Text("Số tiền chi", size=(10, 1)), sg.Input(key="Số tiền chi", size=47), sg.Text("Ghi Chú:", size=(6, 1)), sg.Input(key="Ghi chú", size=(59, 1))],
    [sg.Button("Save", button_color="green"), sg.Button("Modify", button_color="green"), 
     sg.Button("Delete", button_color="green"), sg.Button("Show", button_color="green"), 
     sg.Button("Statistics", button_color="green"), sg.Button("Chart", button_color="green"), 
     sg.Button("Nhắc nhở & Phân tích", button_color="green"), sg.Button("Exit", button_color="green")],
]

# Hàm cập nhật bảng dữ liệu
def update_table():
    global df
    table_data = df.values.tolist()
    return table_data

# Hàm xóa dữ liệu nhập
def clear_input():
    for key in ['Số ID', 'Họ Tên', 'Giới tính', 'Danh mục', 'Ngày', 'Số tiền chi', 'Ghi chú']:
        window[key].update('')

# Hàm vẽ biểu đồ lên canvas
def draw_figure(canvas, figure):
    figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
    figure_canvas_agg.draw()
    figure_canvas_agg.get_tk_widget().pack(side='top', fill='both', expand=1)
    return figure_canvas_agg

# Hàm phân tích và nhắc nhở
def analyze_and_remind():
    try:
        # Tính toán tổng chi tiêu
        total_spending = df['Số tiền chi'].sum()
        
        # Tính toán chi tiêu theo danh mục
        category_spending = df.groupby('Danh mục')['Số tiền chi'].sum().sort_values(ascending=False)
        
        # Phân tích chi tiêu
        analysis = []
        analysis.append(f"📊 **Tổng chi tiêu:** {total_spending:,.0f} VND\n")
        
        for category, amount in category_spending.items():
            percentage = (amount / total_spending) * 100
            analysis.append(f"🔹 **{category}:** {amount:,.0f} VND ({percentage:.1f}%)\n")
        
        # Đề xuất
        suggestions = []
        suggestions.append("\n💡 **Đề xuất:**\n")
        
        if category_spending.get('Tổ chức sự kiện', 0) > 0.3 * total_spending:
            suggestions.append("- Giảm chi tiêu **Tổ chức sự kiện** bằng cách chỉ mua vật liệu cần thiết.\n")
        
        if category_spending.get('Phí sinh hoạt CLB', 0) > 0.2 * total_spending:
            suggestions.append("- Lập kế hoạch trước 3-6 tuần.\n")
        
        if category_spending.get('Trang trí lớp', 0) > 0.15 * total_spending:
            suggestions.append("- Đặt giới hạn hàng tháng cho **Trang trí lớp** (ví dụ: 300,000 VND).\n")
        
        if category_spending.get('In tài liệu', 0) > 0.1 * total_spending:
            suggestions.append("- Nên in những tài liệu cần thiết cho lớp.\n")
        
        if category_spending.get('Khác', 0) < 0.1 * total_spending:
            suggestions.append("- Tỷ lệ **Khác** chỉ ít nhất 10% quỹ lớp đã chi tiêu.\n")
        
        # Hiển thị kết quả
        sg.popup_scrolled(''.join(analysis) + ''.join(suggestions), title="Phân tích & Nhắc nhở")
    except Exception as e:
        sg.popup(f"Lỗi khi phân tích dữ liệu: {str(e)}")

# Tạo cửa sổ chính
window = sg.Window("Giao diện người dùng", layout, resizable=True)

# Xử lý sự kiện chính
while True:
    event, values = window.read()
    
    # Thoát chương trình
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    
    # Chức năng lưu dữ liệu
    if event == "Save":
        try:
            new_data = {
                'Số ID': int(values['Số ID']),
                'Họ Tên': values['Họ Tên'],
                'Giới tính': values['Giới tính'],
                'Danh mục': values['Danh mục'],
                'Ngày': values['Ngày'],
                'Số tiền chi': float(values['Số tiền chi'].replace(',', '')),
                'Ghi chú': values['Ghi chú']
            }
            df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
            df.to_excel("data_demo.xlsx", index=False)
            sg.popup("Lưu thành công!")
            clear_input()
        except Exception as e:
            sg.popup(f"Lỗi khi lưu dữ liệu: {str(e)}")

    # Chức năng hiển thị bảng dữ liệu
    if event == "Show":
        try:
            layout_show = [
                [sg.Table(values=update_table(),
                          headings=table_header,
                          auto_size_columns=True,
                          display_row_numbers=True,
                          justification='center',
                          key='-TABLE-',
                          row_height=35)]
            ]
            window_show = sg.Window("Danh sách chi tiêu", layout_show, resizable=True)
            while True:
                event_show, _ = window_show.read()
                if event_show in (sg.WIN_CLOSED, 'Exit'):
                    break
            window_show.close()
        except Exception as e:
            sg.popup(f"Lỗi khi hiển thị bảng: {str(e)}")

    # Chức năng sửa dữ liệu
    if event == "Modify":
        try:
            if values["Số ID"] == "":
                layout = [
                    [sg.Text("Vui lòng nhập ID để sửa đổi !!")],
                    [sg.Text('ID: ', size=(15, 1)), sg.InputText(key="id_in")],
                    [sg.Submit(), sg.Cancel()]
                ]
                window1 = sg.Window("Sửa theo ID", layout)
                event1, value_id = window1.read()
                id_in = int(value_id["id_in"])
                indexa = df.loc[df['Số ID'] == id_in].index.tolist()[0]

                df_id = df.iloc[indexa]
                dicta = df_id.to_dict()
                window1.close()
                for key, value in dicta.items():
                    window[key].update(value)
            else:
                indexa = df.loc[df['Số ID'] == id_in].index.tolist()[0]
                header_list = list(df.columns.values)
                
                for key in header_list:
                    df.loc[indexa, key] = values[key]

                df.to_excel("data_demo.xlsx", index=False)
                sg.popup("Chỉnh sửa thành công!")
                clear_input()
        except Exception as e:
            sg.popup(f"Lỗi khi sửa dữ liệu: {str(e)}")

    # Chức năng xóa dữ liệu
    if event == "Delete":
        if values["Số ID"] == "":
                layout = [
                    [sg.Text("Vui lòng nhập ID cần xóa !!")],
                    [sg.Text('ID: ', size=(15, 1)), sg.InputText(key="id_in")],
                    [sg.Submit(), sg.Cancel()]
                ]
                window1 = sg.Window("Xóa theo ID", layout)
                event1, value_id = window1.read()
                id_in = int(value_id["id_in"])
                indexa = df.loc[df['Số ID'] == id_in].index.tolist()[0]

                df_id = df.iloc[indexa]
                dicta = df_id.to_dict()
                window1.close()
                for key, value in dicta.items():
                    window[key].update(value)
        else:
                indexa = df.loc[df['Số ID'] == id_in].index.tolist()[0]
                delete_df = df.drop(indexa)
                delete_df.to_excel("data_demo.xlsx", index=False)
                sg.popup("Xóa thành công!")
                clear_input()
        
            
        

    # Chức năng thống kê
    if event == "Statistics":
        try:
            stats = []
            total = df['Số tiền chi'].sum()
            stats.append(f"Tổng chi tiêu: {total:,.0f} VND")
            
            by_category = df.groupby('Danh mục')['Số tiền chi'].agg(['sum', 'count', 'mean'])
            for idx, row in by_category.iterrows():
                stats.append(f"\n{idx}:")
                stats.append(f" - Tổng chi: {row['sum']:,.0f} VND")
                stats.append(f" - Số lần chi: {row['count']}")
                stats.append(f" - Trung bình: {row['mean']:,.0f} VND")
            
            sg.popup_scrolled('\n'.join(stats), title="Thống kê chi tiêu")
        except Exception as e:
            sg.popup(f"Lỗi khi thống kê: {str(e)}")

    # Chức năng vẽ biểu đồ
    if event == "Chart":
        try:
            # Tạo figure và axes
            fig, ax = plt.subplots(figsize=(8, 6))
            
            # Tính toán dữ liệu
            category_totals = df.groupby('Danh mục')['Số tiền chi'].sum()
            
            # Vẽ biểu đồ cột
            bars = ax.bar(category_totals.index, category_totals.values, color='skyblue')
            
            # Thêm giá trị trên mỗi cột
            for bar in bars:
                height = bar.get_height()
                ax.annotate(f'{height:,.0f}',
                            xy=(bar.get_x() + bar.get_width() / 2, height),
                            xytext=(0, 3),  # 3 points vertical offset
                            textcoords="offset points",
                            ha='center', va='bottom')
            
            # Thiết lập tiêu đề và nhãn
            ax.set_title('TỔNG CHI TIÊU THEO DANH MỤC', fontsize=14)
            ax.set_xlabel('Danh mục', fontsize=12)
            ax.set_ylabel('Tổng chi tiêu (VND)', fontsize=12)
            
            # Xoay nhãn trục x để dễ đọc
            plt.xticks(rotation=45, ha='right')
            
            # Lưu biểu đồ thành ảnh
            plt.savefig('bieu_do.png')
            sg.popup("Biểu đồ đã được lưu thành 'bieu_do.png'")
            
            # Hiển thị biểu đồ (nếu sử dụng TkAgg)
            if matplotlib.get_backend() == 'TkAgg':
                layout_chart = [
                    [sg.Canvas(key='-CANVAS-')],
                    [sg.Button('Đóng')]
                ]
                window_chart = sg.Window('Biểu đồ cột', layout_chart, finalize=True)
                draw_figure(window_chart['-CANVAS-'].TKCanvas, fig)
                
                while True:
                    event_chart, _ = window_chart.read()
                    if event_chart in (sg.WIN_CLOSED, 'Đóng'):
                        break
                
                window_chart.close()
            
            # Đóng figure
            plt.close('all')
        except Exception as e:
            sg.popup(f"Lỗi khi vẽ biểu đồ: {str(e)}")

    # Chức năng nhắc nhở & phân tích
    if event == "Nhắc nhở & Phân tích":
        analyze_and_remind()

# Đóng cửa sổ chính
window.close()