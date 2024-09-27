import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageTk
import json
import os
import win32print
import win32ui
from PIL import ImageWin

# 保存设置的文件
SETTINGS_FILE = "settings.json"

# 加载或初始化设置
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, 'r') as f:
            settings = json.load(f)
            if 'dpi' not in settings:
                settings['dpi'] = 300  # 默认DPI
            return settings
    else:
        return {"spacing": 10, "dpi": 300, "label_width": 34, "label_height": 23, "columns": 3}  # 默认设置

def save_settings(settings):
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f)

# 生成条码图像
def generate_barcode(data, dpi):
    barcode_writer = ImageWriter()
    barcode = Code128(data, writer=barcode_writer)
    barcode_image = barcode.render(writer_options={"module_height": 10, "module_width": 0.2, "dpi": dpi})
    return barcode_image

# 创建单行标签布局用于预览
def create_preview_label(data, spacing, dpi, label_width, label_height, columns):
    label_width_px = int(label_width * dpi / 25.4)  # mm to pixels
    label_height_px = int(label_height * dpi / 25.4)  # mm to pixels

    preview_sheet_width = label_width_px * columns + spacing * (columns - 1)

    preview_label = Image.new("RGB", (preview_sheet_width, label_height_px), "white")

    for col in range(columns):
        barcode_image = generate_barcode(data, dpi).resize((label_width_px, label_height_px))
        x = col * (label_width_px + spacing)
        preview_label.paste(barcode_image, (x, 0))

    return preview_label

# 创建标签布局
def create_label_sheet(data, spacing, quantity, dpi, label_width, label_height, columns):
    label_width_px = int(label_width * dpi / 25.4)  # mm to pixels
    label_height_px = int(label_height * dpi / 25.4)  # mm to pixels

    total_rows = (quantity + columns - 1) // columns
    sheet_height = label_height_px * total_rows + spacing * (total_rows - 1)

    label_sheet = Image.new("RGB", (label_width_px * columns + spacing * (columns - 1), sheet_height), "white")

    for q in range(quantity):
        col = q % columns
        row = q // columns
        barcode_image = generate_barcode(data, dpi).resize((label_width_px, label_height_px))
        x = col * (label_width_px + spacing)
        y = row * (label_height_px + spacing)
        label_sheet.paste(barcode_image, (x, y))

    return label_sheet

# 打印条码到打印机
def print_barcode_image(image, printer_name, quantity):

    # 获取当前用户的主目录
    user_directory = os.path.expanduser("~")

     # 设置保存路径为用户目录
    save_path = os.path.join(user_directory, "temp_barcode_sheet.bmp")  # 用户目录
     # 保存图像到用户目录
    image.save(save_path)
   
    
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)
    printable_area = hDC.GetDeviceCaps(8), hDC.GetDeviceCaps(10)
    printer_size = hDC.GetDeviceCaps(110), hDC.GetDeviceCaps(111)
    
    image_width, image_height = image.size
    hDC.StartDoc("Barcode Sheet")
    
    for _ in range(quantity):  # 根据数量打印多次
        hDC.StartPage()
        
        dib = ImageWin.Dib(image)
        scale_x = printer_size[0] / image_width
        scale_y = printer_size[1] / image_height
        scale = min(scale_x, scale_y)

        scaled_width = int(image_width * scale)
        scaled_height = int(image_height * scale)

        dib.draw(hDC.GetHandleOutput(), (0, 0, scaled_width, scaled_height))
        
        hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

# 显示打印预览
def show_preview(data, spacing, dpi, label_width, label_height, columns, printer_name):
    preview_window = tk.Toplevel()
    preview_window.title("打印预览")
    preview_window.geometry("800x600")  # 固定窗口大小

    # 创建Canvas以便于调整内容
    canvas = tk.Canvas(preview_window, bg="white")
    canvas.pack(fill="both", expand=True)

    # 创建预览标签
    image = create_preview_label(data, spacing, dpi, label_width, label_height, columns)

    # 调整图像大小并保持比例
    img_width, img_height = image.size
    aspect_ratio = img_height / img_width
    new_width = 600  # 固定宽度
    new_height = int(new_width * aspect_ratio)  # 按比例计算高度
    img = ImageTk.PhotoImage(image.resize((new_width, new_height)))  # 缩放图像

    # 在Canvas上创建图像
    canvas_img = canvas.create_image(0, 0, anchor='nw', image=img)
    canvas.image = img  # 保持引用，防止被垃圾回收

    # 标签数量输入框
    tk.Label(preview_window, text="标签数量（需要打印次数）:", font=("Arial", 14)).pack(pady=10)
    quantity_entry = tk.Entry(preview_window, font=("Arial", 14), width=10)
    quantity_entry.pack(pady=10)

    # 打印按钮
    btn_frame = tk.Frame(preview_window)
    btn_frame.pack(side="bottom", pady=10)
    btn = tk.Button(btn_frame, text="打印", font=("Arial", 14), bg="#4CAF50", fg="white", 
                    command=lambda: (print_barcode_image(image, printer_name, int(quantity_entry.get())), preview_window.destroy()))
    btn.pack(pady=10)

# 获取可用打印机列表
def get_printer_list():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    printer_list = [printer[2] for printer in printers]
    return printer_list

# 主界面
def main_window():
    root = tk.Tk()
    root.title("条码打印程序")
    root.geometry("600x700")  # 设置窗口大小

    settings = load_settings()

    # 设置背景颜色和样式
    root.configure(bg="#2C3E50")

    # 标题
    title_label = tk.Label(root, text="条码打印", font=("Arial", 24), bg="#2C3E50", fg="white")
    title_label.pack(pady=20)

    # 条码输入框
    tk.Label(root, text="请输入条码内容:", font=("Arial", 14), bg="#2C3E50", fg="white").pack()
    barcode_entry = tk.Entry(root, font=("Arial", 14), width=30)
    barcode_entry.pack(pady=10)

    # 间距设置框
    tk.Label(root, text=f"标签间距 (当前: {settings['spacing']} 像素):", font=("Arial", 14), bg="#2C3E50", fg="white").pack()
    spacing_entry = tk.Entry(root, font=("Arial", 14), width=30)
    spacing_entry.insert(0, str(settings['spacing']))
    spacing_entry.pack(pady=10)

    # DPI输入框
    tk.Label(root, text=f"条码DPI（当前: {settings['dpi']}）:", font=("Arial", 14), bg="#2C3E50", fg="white").pack()
    dpi_entry = tk.Entry(root, font=("Arial", 14), width=30)
    dpi_entry.insert(0, str(settings['dpi']))
    dpi_entry.pack(pady=10)

    # 标签宽度输入框
    tk.Label(root, text=f"标签宽度 (mm, 当前: {settings['label_width']}):", font=("Arial", 14), bg="#2C3E50", fg="white").pack()
    label_width_entry = tk.Entry(root, font=("Arial", 14), width=30)
    label_width_entry.insert(0, str(settings['label_width']))
    label_width_entry.pack(pady=10)

    # 标签高度输入框
    tk.Label(root, text=f"标签高度 (mm, 当前: {settings['label_height']}):", font=("Arial", 14), bg="#2C3E50", fg="white").pack()
    label_height_entry = tk.Entry(root, font=("Arial", 14), width=30)
    label_height_entry.insert(0, str(settings['label_height']))
    label_height_entry.pack(pady=10)

    # 列数输入框
    tk.Label(root, text=f"每列数量 (当前: {settings['columns']}):", font=("Arial", 14), bg="#2C3E50", fg="white").pack()
    columns_entry = tk.Entry(root, font=("Arial", 14), width=30)
    columns_entry.insert(0, str(settings['columns']))
    columns_entry.pack(pady=10)

    # 打印机选择下拉框
    tk.Label(root, text="选择打印机:", font=("Arial", 14), bg="#2C3E50", fg="white").pack()
    printer_combobox = ttk.Combobox(root, font=("Arial", 14), values=get_printer_list())
    printer_combobox.pack(pady=10)
    printer_combobox.current(0)  # 默认选择第一个打印机

    # 打印预览按钮
    def preview():
        data = barcode_entry.get()
        spacing = int(spacing_entry.get())
        dpi = int(dpi_entry.get())
        label_width = float(label_width_entry.get())
        label_height = float(label_height_entry.get())
        columns = int(columns_entry.get())
        printer_name = printer_combobox.get()  # 获取选择的打印机
        


        # 保存用户输入的设置
        settings = {
            "spacing": spacing,
            "dpi": dpi,
            "label_width": label_width,
            "label_height": label_height,
            "columns": columns
        }
        save_settings(settings)  # 保存设置到 JSON 文件
        if not data:
            messagebox.showerror("输入错误", "条码内容不能为空！")
            return

        
        show_preview(data, spacing, dpi, label_width, label_height, columns, printer_name)

    preview_button = tk.Button(root, text="打印预览", font=("Arial", 14), command=preview)
    preview_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main_window()
