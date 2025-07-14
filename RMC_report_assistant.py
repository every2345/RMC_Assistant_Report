from re import L
import tkinter as tk
from tkinter import ttk, messagebox
import schedule
import threading
import time
import datetime
import gdown
import os
from PIL import Image, ImageTk
import pyperclip
import json

# Tạo cửa sổ chính
root = tk.Tk()
root.title("RMC Report Assistant")
root.geometry("1080x800")

# Frame chính
main_frame = tk.Frame(root)
main_frame.pack(expand=True, pady=40, padx=20)

# Frame con chứa văn bản và các nút
content_frame = tk.Frame(main_frame)
content_frame.pack()

# === FRAME CHỨA CONTACT, NOTE VÀ STATUS BÊN TRÁI ===
left_button_frame = tk.Frame(content_frame)
left_button_frame.pack(side='left', padx=10)

# === Text để hiển thị văn bản ===
output_text = tk.Text(content_frame, font=("Arial", 13), width=60, height=20, wrap="word")
output_text.pack(side='left', pady=(10, 0), padx=10)
output_text.config(state='disabled')

# === Frame chứa các danh sách ATQB, ABDNC... ===
button_frame = tk.Frame(content_frame)
button_frame.pack(side='left', padx=10)

# === Frame chứa các item xuất hiện khi chọn danh sách ===
item_frame = tk.Frame(content_frame)
item_frame.pack(side='left', padx=10)

# === Biến điều khiển đồng hồ ===
is_running = True

def update_clock():
    if is_running:
        now = datetime.datetime.now().strftime("%H:%M:%S")
        clock_label.config(text=now)
    root.after(1000, update_clock)

def catch_clock():
    global is_running
    is_running = False

def continue_clock():
    global is_running
    is_running = True

def copy_text_to_clipboard():
    text = output_text.get("1.0", "end-1c")
    pyperclip.copy(text)

def clear_text_output():
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.config(state='disabled')

def display_text_from_file(file_path):
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()
        output_text.config(state='normal')
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, content)
        output_text.config(state='disabled')
    except Exception as e:
        output_text.config(state='normal')
        output_text.insert(tk.END, f"Lỗi khi đọc file: {e}")
        output_text.config(state='disabled')

# ==== HÀM GOOGLE DRIVE ====
def download_from_drive(file_id):

    # Đường dẫn cache gốc trong ổ D
    cache_dir = r"D:\RMC_Assistant\Cache"
    os.makedirs(cache_dir, exist_ok=True)

    # Đường dẫn đến file cache
    cache_path = os.path.join(cache_dir, f"{file_id}.txt")

    # Nếu file cache tồn tại, trả về luôn
    if os.path.exists(cache_path):
        return cache_path

    # Nếu chưa có, tải file từ Google Drive và lưu vào cache
    url = f'https://drive.google.com/uc?id={file_id}'
    try:
        gdown.download(url, cache_path, quiet=True)
        return cache_path
    except Exception as e:
        return f"ERROR: {e}"

# ==== THÊM ĐỒNG HỒ ĐẾM NGƯỢC ==== 
timer_frame = tk.Frame(main_frame)
timer_frame.pack(pady=(10, 0))

timer_label = tk.Label(timer_frame, text="⏳Waiting Countdown⏳", font=("Arial", 16, "bold"), fg="blue")
timer_label.pack()

countdown_job = None
time_left = 300  # 5 phút = 300 giây

def update_timer():
    global time_left, countdown_job
    minutes, seconds = divmod(time_left, 60)
    timer_label.config(text=f"⏳ Remain: {minutes:02d}:{seconds:02d}")
    if time_left > 0:
        time_left -= 1
        countdown_job = root.after(1000, update_timer)
    else:
        timer_label.config(text="⏰Contact Site⏰")

def start_timer():
    global time_left, countdown_job
    if countdown_job:
        root.after_cancel(countdown_job)
    time_left = 300
    update_timer()

def reset_timer():
    global time_left, countdown_job
    if countdown_job:
        root.after_cancel(countdown_job)
        countdown_job = None
    time_left = 300
    timer_label.config(text="⏳Waiting Countdown⏳")

# ==== CHỨC NĂNG HIỂN THỊ VĂN BẢN VÀ THỜI GIAN====
def show_text_from_drive(file_id, is_no_error=False, start_timer_flag=True):
    file_path = download_from_drive(file_id)

    if not file_path or file_path.startswith("ERROR"):
        output_text.config(state='normal')
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, f"Lỗi tải file từ Google Drive: {file_path}")
        output_text.config(state='disabled')
        return

    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        if is_no_error:
            yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
            timestamp = yesterday.strftime("Trong ngày: %d-%m-%Y ") + '\n'
            lines = [timestamp if '[no_error_time]' in line else line for line in lines]
        else:
            delayed_time = datetime.datetime.now() - datetime.timedelta(minutes=1)
            current_time = delayed_time.strftime("+ Thời gian: %H:%M:%S %d-%m-%Y ") + '\n'
            lines = [current_time if '[time]' in line else line for line in lines]

        # Ghi lại nếu cần
        with open(file_path, 'w', encoding='utf-8', errors='ignore') as f:
            f.writelines(lines)

        display_text_from_file(file_path)

    except Exception as e:
        output_text.config(state='normal')
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, f"Lỗi khi xử lý nội dung file: {e}")
        output_text.config(state='disabled')

    if start_timer_flag:
        start_timer()

# ==== TẠO GIAO DIỆN DANH SÁCH ====
def create_list_block(parent, list_name, items, toggle_function, state):
    block_frame = tk.Frame(parent)
    block_frame.pack(pady=10, anchor='w')
    row_frame = tk.Frame(block_frame)
    row_frame.pack(anchor='w')

    color_indicator = tk.Canvas(row_frame, width=20, height=20, highlightthickness=0)
    color_indicator.pack(side='left', padx=(0, 10))
    indicator = color_indicator.create_oval(2, 2, 18, 18, fill='red')

    list_button = tk.Button(row_frame, text=list_name, command=lambda: toggle_function(state),
                            font=("Arial", 14), width=8)
    list_button.pack(side='left')

    state["indicator_canvas"] = color_indicator
    state["indicator_id"] = indicator

# ==== CÁC DANH SÁCH FILE ID (BẠN ĐIỀN VÀO SAU) ====
list1_files = {
    "FR&FC": "17MyGAgm4zfwH0dCIoT2hS2USzqqTDpoS",
    "POWER": "1wn10nS9ca33JGPia1enBKoflimi6R4bF",
    "FAN": "1x6WwI6IBF34bxk9neMsFWUBiHrOwzaAP",
    "DELICA": "1PETK7KmIPTyySoMlUeinUpzw39ViQanG",
    "SUSHI": "1642wL6grtH1K9u1Obqjnr6cjqnsqfmJd",
    "BAKERY": "1zsjHS11bpS6wa0ThVr3b4gtTU9Rwk0Oo",
    "NO_ERROR": "1kGNNzWSt5OXVq6KMZtmw6glhqOu1aAxi"
}

list2_files = {
    "FAN": "13pDhw6LfsONDh2prelUUu09JYi8MfWoG",
    "POWER": "1Mr6b7g9A8ehOJnUkmPOT3rbPy7fsyCgH",
    "FR&FC": "1cZj-Bw94AqpDBOJCO3v-fRBVCOCnN1UI",
    "LPG": "1vtIVCz8NAqNBQ3wHztW314ffgcv_hiSZ",
    "NO_ERROR": "1zaWLKejWabN_3Zp2M2qiC3p_7w-6xL5Y"
}

list3_files = {
    "FAN": "1og4OfF53YsYEfDK1MxT5zFCOsxhAXSSG",
    "POWER_1": "17gRIE14NcSGsrPKNv203tjlLVjfycfJt",
    "POWER_2": "1TSyDxCadsK4UJ2oFOjbYFa_0x4Lhh2jJ",
    "POWER_3": "1OhEx-Kgi4eJZWk7FxL4jw3TLEEp51Eet",
    "FR&FC": "11531CGpGf5b0NIaAUd6WHrXOZStZFhXz",
    "DELICA": "1TN_yaxyAyuLkf6PwDX9_SCQsFlpaWnnM",
    "NO_ERROR": "1BaMisKLfSq-0xyKgRCEn2WOWVtycn_aE"
}

contact_sample = { "1MuZBmMdeFZPkiOg9wHwNNpguSf_k-jrI" }

confirm_sample = { "1lIXTD2ryyYg9Qob0xedYemI9UGMSIqoV" }

#==== HÌNH ẢNH ====
image_data = {
    "1da-cjc1egzxy9fYY4yEqtbiFiRktzP1a": "NVL_Delica",
    "1HtmBQkPkDXjfGKxFO0RCp9JlSSx327OA": "TQB_Delica",
    "1a8AySv4aumTqPmaHYrqlZQ94Ja8eJzqf": "TQB_Sushi",
    "1m2hklLZhRYMaLioL92gf8r8lKXYaSfiz": "TQB_Bakery"
}

# ==== TẠO CỬA SỔ MỚI ====
def create_new_window_contact(title, content=None):
    new_window = tk.Toplevel(root)
    new_window.title(title)
    new_window.geometry("600x400")

    confirm_var = tk.StringVar(value="confirmed")

    # Tình trạng confirm
    confirm_frame = tk.LabelFrame(new_window, text="Tình trạng confirm", font=("Arial", 12, "bold"))
    confirm_frame.pack(padx=20, pady=10, fill="x")

    def toggle_entry_fields():
        state = "normal" if confirm_var.get() == "not_confirmed" else "disabled"
        dept_entry.config(state=state)
        device_entry.config(state=state)
        status_entry.config(state=state)
        desc_entry.config(state="normal" if state == "normal" else "disabled")

    tk.Radiobutton(confirm_frame, text="Đã confirm", variable=confirm_var, value="confirmed",
                   command=toggle_entry_fields).pack(anchor="w", padx=10, pady=2)
    tk.Radiobutton(confirm_frame, text="Chưa confirm", variable=confirm_var, value="not_confirmed",
                   command=toggle_entry_fields).pack(anchor="w", padx=10, pady=2)

    # Form nhập liệu
    form_frame = tk.Frame(new_window)
    form_frame.pack(padx=20, pady=10, fill="x")

    tk.Label(form_frame, text="Tên bộ phận:", font=("Arial", 11)).grid(row=0, column=0, sticky="w", pady=5)
    dept_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    dept_entry.grid(row=0, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Tên thiết bị:", font=("Arial", 11)).grid(row=1, column=0, sticky="w", pady=5)
    device_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    device_entry.grid(row=1, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Tình trạng:", font=("Arial", 11)).grid(row=2, column=0, sticky="w", pady=5)
    status_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    status_entry.grid(row=2, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Mô tả:", font=("Arial", 11)).grid(row=3, column=0, sticky="nw", pady=5)
    desc_entry = tk.Text(form_frame, font=("Arial", 11), height=5, width=40, state="disabled")
    desc_entry.grid(row=3, column=1, pady=5, sticky="ew")

    form_frame.columnconfigure(1, weight=1)
    toggle_entry_fields()

    def handle_ok():
        # Chỉ cho phép xử lý khi ở trạng thái "not_confirmed"
        if confirm_var.get() != "not_confirmed":
            new_window.destroy()
            return

        dept = dept_entry.get().strip()
        device = device_entry.get().strip()
        status = status_entry.get().strip()
        desc = desc_entry.get("1.0", tk.END).strip()

        # Tải file contact từ Google Drive
        file_id = next(iter(contact_sample))
        file_path = download_from_drive(file_id)

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            replaced_lines = []
            for line in lines:
                original_line = line  # lưu dòng gốc để kiểm tra sau

                line = line.replace("[title]", dept)
                line = line.replace("[device]", device)
                line = line.replace("[status]", status)
                line = line.replace("[description]", desc)

                # Nếu sau khi thay mà dòng đó trống hoặc chỉ có từ khóa không có giá trị thì bỏ qua
                stripped_line = line.strip()

                if ("[title]" in original_line and not dept) or \
                   ("[device]" in original_line and not device) or \
                   ("[status]" in original_line and not status) or \
                   ("[description]" in original_line and not desc) or \
                   not stripped_line:
                    continue  # bỏ dòng

                replaced_lines.append(stripped_line)  # thêm dòng đã xử lý, loại bỏ khoảng trắng

            # Gộp lại thành một khối văn bản liên tục
            content = '\n'.join(replaced_lines)

        except Exception as e:
            content = f"Lỗi khi xử lý file contact: {e}"

        output_text.config(state='normal')
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, content)
        output_text.config(state='disabled')

        start_timer()
        new_window.destroy()

    # Nút OK
    ok_button = tk.Button(new_window, text="OK", font=("Arial", 12, "bold"),
                          bg="green", fg="white", command=handle_ok)
    ok_button.pack(pady=10)

def create_new_window_status(title, content=None):
    new_window = tk.Toplevel(root)
    new_window.title(title)
    new_window.geometry("600x500")

    confirm_var = tk.StringVar(value="confirmed")

    # Tình trạng confirm
    confirm_frame = tk.LabelFrame(new_window, text="Đã confirm chưa?", font=("Arial", 12, "bold"))
    confirm_frame.pack(padx=20, pady=10, fill="x")

    def toggle_entry_fields():
        state = "normal" if confirm_var.get() == "not_confirmed" else "disabled"
        dept_entry.config(state=state)
        device_entry.config(state=state)
        status_entry.config(state=state)
        start_time_entry.config(state=state)
        end_time_entry.config(state=state)
        desc_entry.config(state="normal" if state == "normal" else "disabled")

    tk.Radiobutton(confirm_frame, text="Đã confirm", variable=confirm_var, value="confirmed",
                   command=toggle_entry_fields).pack(anchor="w", padx=10, pady=2)
    tk.Radiobutton(confirm_frame, text="Chưa confirm", variable=confirm_var, value="not_confirmed",
                   command=toggle_entry_fields).pack(anchor="w", padx=10, pady=2)

    # Form nhập liệu
    form_frame = tk.Frame(new_window)
    form_frame.pack(padx=20, pady=10, fill="x")

    tk.Label(form_frame, text="Tên bộ phận:", font=("Arial", 11)).grid(row=0, column=0, sticky="w", pady=5)
    dept_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    dept_entry.grid(row=0, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Tên thiết bị:", font=("Arial", 11)).grid(row=1, column=0, sticky="w", pady=5)
    device_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    device_entry.grid(row=1, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Tình trạng:", font=("Arial", 11)).grid(row=2, column=0, sticky="w", pady=5)
    status_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    status_entry.grid(row=2, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Thời gian bắt đầu (HH:MM):", font=("Arial", 11)).grid(row=3, column=0, sticky="w", pady=5)
    start_time_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    start_time_entry.grid(row=3, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Thời gian kết thúc (HH:MM):", font=("Arial", 11)).grid(row=4, column=0, sticky="w", pady=5)
    end_time_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    end_time_entry.grid(row=4, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Mô tả:", font=("Arial", 11)).grid(row=5, column=0, sticky="nw", pady=5)
    desc_entry = tk.Text(form_frame, font=("Arial", 11), height=5, width=40, state="disabled")
    desc_entry.grid(row=5, column=1, pady=5, sticky="ew")

    form_frame.columnconfigure(1, weight=1)
    toggle_entry_fields()

    def handle_ok():
        if confirm_var.get() != "not_confirmed":
            new_window.destroy()
            return

        dept = dept_entry.get().strip()
        device = device_entry.get().strip()
        status = status_entry.get().strip()
        start_time_str = start_time_entry.get().strip()
        end_time_str = end_time_entry.get().strip()
        desc = desc_entry.get("1.0", tk.END).strip()

        # Tính thời gian xử lý
        try:
            if start_time_str and end_time_str:
                fmt = "%H:%M"
                start_dt = datetime.datetime.strptime(start_time_str, fmt)
                end_dt = datetime.datetime.strptime(end_time_str, fmt)
                diff_minutes = int((end_dt - start_dt).total_seconds() / 60)
                if diff_minutes < 0:
                    diff_minutes += 24 * 60  # xử lý khi qua ngày
                time = f"{diff_minutes} phút ({start_time_str} - {end_time_str})"
            else:
                time = ""
        except ValueError:
            time = ""

        # Tải file contact từ Google Drive
        file_id = next(iter(confirm_sample))
        file_path = download_from_drive(file_id)

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            replaced_lines = []
            for line in lines:
                original_line = line

                line = line.replace("[tilte]", dept)
                line = line.replace("[device]", device)
                line = line.replace("[status]", status)
                line = line.replace("[time_process]", time)
                line = line.replace("[description]", desc)

                stripped_line = line.strip()

                if ("[tilte]" in original_line and not dept) or \
                   ("[device]" in original_line and not device) or \
                   ("[status]" in original_line and not status) or \
                   ("[time_process]" in original_line and not time) or \
                   ("[description]" in original_line and not desc) or \
                   not stripped_line:
                    continue

                replaced_lines.append(stripped_line)

            content = '\n'.join(replaced_lines)

        except Exception as e:
            content = f"Lỗi khi xử lý file contact: {e}"

        output_text.config(state='normal')
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, content)
        output_text.config(state='disabled')

        new_window.destroy()

    ok_button = tk.Button(new_window, text="OK", font=("Arial", 12, "bold"),
                          bg="green", fg="white", command=handle_ok)
    ok_button.pack(pady=10)

def create_new_window_image(title):
    def show_image_by_ids(id_set):
        for widget in image_frame.winfo_children():
            widget.destroy()

        max_columns = 4
        for idx, file_id in enumerate(id_set):
            img_path = download_from_drive(file_id)
            if img_path and not str(img_path).startswith("ERROR"):
                try:
                    img = Image.open(img_path)
                    img_thumbnail = img.copy()
                    img_thumbnail.thumbnail((100, 75), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img_thumbnail)

                    row = (idx * 2) // max_columns
                    col = idx % max_columns

                    # Display the thumbnail
                    label_img = tk.Label(image_frame, image=photo, bg="white", cursor="hand2")
                    label_img.image = photo
                    label_img.grid(row=row, column=col, padx=5, pady=(5, 0))

                    # Add the name of the image below the thumbnail
                    image_name = image_data.get(file_id, "Unknown Image")  # Get image name from dictionary
                    label_text = tk.Label(image_frame, text=image_name, bg="white", font=("Arial", 9))
                    label_text.grid(row=row + 1, column=col, padx=5, pady=(0, 10))

                    # Event to open large image
                    label_img.bind("<Button-1>", lambda e, path=img_path: open_large_image(path))

                except Exception as e:
                    image_label.config(text=f"Lỗi xử lý ảnh: {e}", image='', bg="white")
                    break
        else:
            if not id_set:
                image_label.config(text="Không có ảnh để hiển thị", image='', bg="white")

    def open_large_image(img_path):
        try:
            img = Image.open(img_path)
            img = img.resize((400, 300), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)

            popup = tk.Toplevel()
            popup.title("Xem ảnh lớn")
            popup.configure(bg="white")
            lbl = tk.Label(popup, image=photo, bg="white")
            lbl.image = photo
            lbl.pack(padx=10, pady=10)

            # Add Copy button under the image
            copy_button = tk.Button(popup, text="Edit Image", command=lambda: copy_image_to_clipboard(img))
            copy_button.pack(pady=10)

        except Exception as e:
            print(f"Lỗi khi mở ảnh lớn: {e}")

    def copy_image_to_clipboard(img):
        try:
            # Attempt to copy the image to clipboard
            # A typical way to do this involves OS-specific methods or external libraries
            img.show()  # This will open the image in the default viewer (a workaround to copy on some systems)
            pyperclip.copy("Image copied to clipboard!")  # This line can copy the text (not image directly)
            print("Image copied to clipboard!")  # Log the action
        except Exception as e:
            print(f"Lỗi khi copy ảnh: {e}")

    new_window = tk.Toplevel()
    new_window.title(title)
    new_window.geometry("600x400")
    new_window.configure(bg="white")

    left_frame = tk.Frame(new_window, width=150, bg="#f0f0f0")
    left_frame.pack(side="left", fill="y")

    right_frame = tk.Frame(new_window, bg="white")
    right_frame.pack(side="right", fill="both", expand=True)

    global image_frame
    image_frame = tk.Frame(right_frame, bg="white")
    image_frame.pack(expand=True)

    global image_label
    image_label = tk.Label(right_frame, bg="white", text="", font=("Arial", 14))
    image_label.pack()

    # Buttons to show images based on category
    delica_image = {"1da-cjc1egzxy9fYY4yEqtbiFiRktzP1a", "1HtmBQkPkDXjfGKxFO0RCp9JlSSx327OA"}
    sushi_image = {"1a8AySv4aumTqPmaHYrqlZQ94Ja8eJzqf"}
    bakery_image = {"1m2hklLZhRYMaLioL92gf8r8lKXYaSfiz"}

    tk.Button(left_frame, text="Delica", command=lambda: show_image_by_ids(delica_image),
              width=15, pady=5, bg="#FF9800", fg="white").pack(pady=10)
    tk.Button(left_frame, text="Sushi", command=lambda: show_image_by_ids(sushi_image),
              width=15, pady=5, bg="#FF9800", fg="white").pack(pady=10)
    tk.Button(left_frame, text="Bakery", command=lambda: show_image_by_ids(bakery_image),
              width=15, pady=5, bg="#FF9800", fg="white").pack(pady=10)
    
    reset_timer()

def create_new_window_note():
    # Thư mục lưu dữ liệu
    DATA_DIR = r"D:\RMC_Assistant\Note"
    os.makedirs(DATA_DIR, exist_ok=True)

    # === Schedule Thread ===
    def run_schedule():
        while True:
            schedule.run_pending()
            time.sleep(1)

    threading.Thread(target=run_schedule, daemon=True).start()

    # === Tạo Note ===
    def get_next_stt():
        count = 1
        while os.path.exists(os.path.join(DATA_DIR, f"reminders{count}.json")):
            count += 1
        return count

    def update_stt_label():
        current_stt.set(str(get_next_stt()))

    def save_reminder_to_new_file(reminder_data):
        stt = get_next_stt()
        file_path = os.path.join(DATA_DIR, f"reminders{stt}.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(reminder_data, f, ensure_ascii=False, indent=4)
        update_stt_label()

    def schedule_reminder(keyword, content, times, days, months, mode, file_path=None):
        for t in times:
            def job(t=t):
                now = datetime.datetime.now()
                if str(now.day) in days and str(now.month) in months:
                    # Hiển thị thông báo đúng luồng giao diện
                    def show_popup():
                        messagebox.showinfo(f"Thông báo: {keyword}", f"[{t}] {content}")
                    
                        if mode == "1 lần" and file_path and os.path.exists(file_path):
                            try:
                                os.remove(file_path)
                                print(f"Đã xóa file: {file_path}")
                            except Exception as e:
                                print(f"Lỗi xóa file {file_path}: {e}")

                    try:
                        note_window.after(0, show_popup)
                    except Exception as e:
                        print(f"Lỗi gọi after: {e}")
                
                    if mode == "1 lần":
                        return schedule.CancelJob
            schedule.every().day.at(t).do(job)

    def add_reminder():
        keyword = keyword_entry.get().strip()
        content = content_entry.get().strip()
        time_input = time_entry.get().strip()
        day_input = day_entry.get().strip()
        month_input = month_entry.get().strip()

        if time_input == "08:00,12:00,14:00,...": time_input = ""
        if day_input == "1,15,All": day_input = ""
        if month_input == "1,6,12,All": month_input = ""

        time_strs = time_input.split(",")

        if day_input.strip().lower() == "all":
            day_strs = [str(d) for d in range(1, 32)]
        else:
            day_strs = day_input.split(",")

        if month_input.strip().lower() == "all":
            month_strs = [str(m) for m in range(1, 13)]
        else:
            month_strs = month_input.split(",")

        mode = intensity_var.get()

        try:
            for t in time_strs:
                datetime.datetime.strptime(t.strip(), "%H:%M")
            for d in day_strs:
                d = int(d.strip())
                assert 1 <= d <= 31
            for m in month_strs:
                m = int(m.strip())
                assert 1 <= m <= 12
        except:
            messagebox.showerror("Lỗi", "Thời gian, ngày hoặc tháng không hợp lệ", parent=note_window)
            return

        times = [t.strip() for t in time_strs]
        days = [d.strip() for d in day_strs]
        months = [m.strip() for m in month_strs]

        schedule_reminder(keyword, content, times, days, months, mode)
        reminder_data = {
            "keyword": keyword,
            "content": content,
            "times": times,
            "days": days,
            "months": months,
            "mode": mode
        }
        save_reminder_to_new_file(reminder_data)
        file_path = os.path.join(DATA_DIR, f"reminders{get_next_stt()-1}.json")
        schedule_reminder(keyword, content, times, days, months, mode, file_path)
        messagebox.showinfo("Thành công", f"Đã tạo note {get_next_stt()-1}.json", parent=note_window)

    def set_placeholder(entry, text):
        entry.insert(0, text)
        entry.config(fg="gray")

        def on_focus_in(event):
            if entry.get() == text:
                entry.delete(0, tk.END)
                entry.config(fg="black")
        def on_focus_out(event):
            if not entry.get():
                entry.insert(0, text)
                entry.config(fg="gray")

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def load_all_json_files():
        if not os.path.exists(DATA_DIR):
            messagebox.showerror("Lỗi", f"Không tìm thấy thư mục: {DATA_DIR}", parent=note_window)
            return []

        all_data = []
        for filename in os.listdir(DATA_DIR):
            if filename.endswith(".json"):
                file_path = os.path.join(DATA_DIR, filename)
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if isinstance(data, dict) and "keyword" in data:
                            data["_file"] = file_path
                            all_data.append(data)
                        elif isinstance(data, list):
                            for item in data:
                                if isinstance(item, dict) and "keyword" in item:
                                    item["_file"] = file_path
                                    all_data.append(item)
                except Exception as e:
                    print(f"Lỗi đọc {filename}: {e}")
        return all_data

    def display_data(data_list):
        for row in tree.get_children():
            tree.delete(row)
        for i, item in enumerate(data_list, start=1):
            tree.insert("", tk.END, values=(
                i,
                item["keyword"],
                item["content"],
                ", ".join(item["times"]),
                ", ".join(item["days"]),
                ", ".join(item["months"]),
                item["mode"]
            ))

    def search_data():
        keyword = search_var.get().lower()
        filtered = [item for item in full_data if keyword in item["keyword"].lower() or keyword in item["content"].lower()]
        display_data(filtered)

    def refresh_data():
        global full_data
        full_data = load_all_json_files()
        display_data(full_data)

    def delete_selected_notes():
        global full_data
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn ít nhất một ghi chú để xóa.", parent=note_window)
            return

        confirm = messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa các ghi chú đã chọn?", parent=note_window)
        if not confirm:
            return

        to_delete = []
        deleted_files = set()

        for item_id in selected_items:
            values = tree.item(item_id, "values")
            keyword = values[1]
            content = values[2]
            for data_item in full_data:
                if data_item["keyword"] == keyword and data_item["content"] == content:
                    file_path = data_item.get("_file")
                    if file_path and os.path.exists(file_path):
                        deleted_files.add(file_path)
                    to_delete.append(data_item)
                    break

        for file_path in deleted_files:
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Lỗi khi xóa file {file_path}: {e}")

        full_data = [item for item in full_data if item not in to_delete]
        display_data(full_data)
        messagebox.showinfo("Thành công", f"Đã xóa {len(to_delete)} ghi chú.", parent=note_window)

    # === Giao diện chính ===
    note_window = tk.Toplevel()
    note_window.title("Trình quản lý ghi chú định kỳ")
    note_window.geometry("1000x400")

    btn_frame = tk.Frame(note_window)
    btn_frame.pack(pady=10)

    main_frame = tk.Frame(note_window)
    main_frame.pack(fill="both", expand=True)

    def show_create_note():
        for w in main_frame.winfo_children():
            w.destroy()

        update_stt_label()
        tk.Label(main_frame, text="STT ghi chú tiếp theo:").pack()
        tk.Label(main_frame, textvariable=current_stt, font=("Arial", 14, "bold"), fg="blue").pack(pady=(0, 10))

        global keyword_entry, content_entry, time_entry, day_entry, month_entry, intensity_var
        tk.Label(main_frame, text="Từ khóa:").pack()
        keyword_entry = tk.Entry(main_frame)
        keyword_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Nội dung:").pack()
        content_entry = tk.Entry(main_frame)
        content_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Thời gian báo (HH:MM, cách nhau dấu phẩy):").pack()
        time_entry = tk.Entry(main_frame)
        set_placeholder(time_entry, "08:00,12:00,14:00,...")
        time_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Ngày báo (VD: 1,15,28):").pack()
        day_entry = tk.Entry(main_frame)
        set_placeholder(day_entry, "1,15 hoặc All")
        day_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Tháng báo (VD: 1,6,12):").pack()
        month_entry = tk.Entry(main_frame)
        set_placeholder(month_entry, "1,6,12 hoặc All")
        month_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Cường độ báo:").pack()
        intensity_var = tk.StringVar(value="1 lần")
        ttk.Combobox(main_frame, textvariable=intensity_var, values=["1 lần", "Cố định"]).pack(fill="x", padx=10)

        tk.Button(main_frame, text="Thêm Nhắc", command=add_reminder).pack(pady=15)

    def show_view_notes():
        for w in main_frame.winfo_children():
            w.destroy()

        search_frame = tk.Frame(main_frame)
        search_frame.pack(padx=10, pady=(10, 0), fill=tk.X)

        tk.Label(search_frame, text="Tìm kiếm:").pack(side=tk.LEFT, padx=(0, 5))

        global search_var, tree, full_data
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(search_frame, text="Tìm", command=search_data,
                  bg="#00ccff", fg="white", activebackground="#006699", activeforeground="white").pack(side=tk.LEFT, padx=5)

        tk.Button(search_frame, text="Làm mới", command=refresh_data,
                  bg="#00cc66", fg="white", activebackground="#006600", activeforeground="white").pack(side=tk.LEFT, padx=5)

        tk.Button(search_frame, text="Xóa ghi chú đã chọn", command=delete_selected_notes,
                  bg="#cc3300", fg="white", activebackground="#990000", activeforeground="white").pack(side=tk.LEFT, padx=5)

        search_entry.bind("<Return>", lambda event: search_data())

        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        columns = ("STT", "Từ khóa", "Nội dung", "Thời gian", "Ngày", "Tháng", "Cường độ")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", yscrollcommand=scrollbar.set, selectmode="extended")
        scrollbar.config(command=tree.yview)
        # Cài đặt tiêu đề + cột
        tree.heading("STT", text="STT")
        tree.column("STT", width=50, anchor="center")

        tree.heading("Từ khóa", text="Từ khóa")
        tree.column("Từ khóa", width=120, anchor="center")

        tree.heading("Nội dung", text="Nội dung")
        tree.column("Nội dung", width=200, anchor="w")

        tree.heading("Thời gian", text="Thời gian")
        tree.column("Thời gian", width=130, anchor="center")

        tree.heading("Ngày", text="Ngày")
        tree.column("Ngày", width=100, anchor="center")

        tree.heading("Tháng", text="Tháng")
        tree.column("Tháng", width=100, anchor="center")

        tree.heading("Cường độ", text="Cường độ")
        tree.column("Cường độ", width=100, anchor="center")

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")

        tree.column("Nội dung", width=200, anchor="w")
        tree.pack(fill=tk.BOTH, expand=True)

        full_data = load_all_json_files()
        display_data(full_data)

    tk.Button(btn_frame, text="Tạo Note", width=20, command=show_create_note).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Xem Note", width=20, command=show_view_notes).pack(side="left", padx=10)

    current_stt = tk.StringVar()
    show_create_note()

    # ==== Lên lịch lại tất cả ghi chú đã lưu ====
    for reminder in load_all_json_files():
        schedule_reminder(
            reminder["keyword"],
            reminder["content"],
            reminder["times"],
            reminder["days"],
            reminder["months"],
            reminder["mode"],
            reminder.get("_file")  # thêm đường dẫn file
        )

# ==== HÀM BẬT TẮT DANH SÁCH ====
def toggle_list1(state):
    if list2_state["visible"]:
        toggle_sub_buttons(list2_state, list2_files)
    if list3_state["visible"]:
        toggle_sub_buttons(list3_state, list3_files)
    toggle_sub_buttons(state, list1_files)

def toggle_list2(state):
    if list1_state["visible"]:
        toggle_sub_buttons(list1_state, list1_files)
    if list3_state["visible"]:
        toggle_sub_buttons(list3_state, list3_files)
    toggle_sub_buttons(state, list2_files)

def toggle_list3(state):
    if list1_state["visible"]:
        toggle_sub_buttons(list1_state, list1_files)
    if list2_state["visible"]:
        toggle_sub_buttons(list2_state, list2_files)
    toggle_sub_buttons(state, list3_files)

# ==== SAO CHÉP VĂN BẢN ====
def copy_text_to_clipboard():
    root.clipboard_clear()
    text = output_text.get("1.0", tk.END)
    root.clipboard_append(text)
    root.update()

# ==== XOÁ VĂN BẢN ====
def clear_text_output():
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.config(state='disabled')
    reset_timer()

# ==== TẠO CÁC NÚT CHO DANH SÁCH ====
def toggle_sub_buttons(state, item_dict):
    if not state["visible"]:
        for label, file_id in item_dict.items():
            # Nếu là NO_ERROR thì không khởi động đếm ngược
            if "NO_ERROR" in label:
                cmd = lambda fid=file_id: show_text_from_drive(fid, is_no_error=True, start_timer_flag=False)
            else:
                cmd = lambda fid=file_id: show_text_from_drive(fid, start_timer_flag=True)

            btn = tk.Button(item_frame, text=label, font=("Arial", 12), command=cmd)
            btn.pack(anchor='w', pady=1)
            state["buttons"].append(btn)
        state["visible"] = True
        state["indicator_canvas"].itemconfig(state["indicator_id"], fill='green')
    else:
        for btn in state["buttons"]:
            btn.pack_forget()
        state["buttons"].clear()
        state["visible"] = False
        state["indicator_canvas"].itemconfig(state["indicator_id"], fill='red')

# ==== TRẠNG THÁI ====
list1_state = {"visible": False, "buttons": [], "indicator_canvas": None, "indicator_id": None}
list2_state = {"visible": False, "buttons": [], "indicator_canvas": None, "indicator_id": None}
list3_state = {"visible": False, "buttons": [], "indicator_canvas": None, "indicator_id": None}

# ==== TẠO DANH SÁCH GIAO DIỆN ====
create_list_block(button_frame, "ATQB", list1_files, toggle_list1, list1_state)
create_list_block(button_frame, "ABDNC", list2_files, toggle_list2, list2_state)
create_list_block(button_frame, "ANVL", list3_files, toggle_list3, list3_state)

# ==== NÚT COPY ====
copy_frame = tk.Frame(main_frame)
copy_frame.pack(fill='x', pady=(10, 0), padx=20)

# Nhóm bên trái: Copy và Clear
left_controls = tk.Frame(copy_frame)
left_controls.pack(side="left")

copy_button = tk.Button(left_controls, text="Copy", font=("Arial", 10, "bold"), bg="#4CAF50", fg="white",
                        command=copy_text_to_clipboard, width=15)
copy_button.pack(side="left", padx=(0, 5))

clear_button = tk.Button(left_controls, text="Clear", font=("Arial", 10, "bold"), bg="#f44336", fg="white",
                         command=clear_text_output, width=15)
clear_button.pack(side="left")

# Nhóm bên phải: Catch, Clock, Continue
right_controls = tk.Frame(copy_frame)
right_controls.pack(side="right")

# Catch (ngoài cùng bên phải)
catch_button = tk.Button(right_controls, text="Catch", font=("Arial", 10, "bold"), bg="#029B82", fg="white",
                         command=catch_clock, width=10)
catch_button.pack(side='right', padx=5)

# Clock (giữa)
clock_label = tk.Label(right_controls, font=("Roboto", 20, "bold"), fg="#D20103",)
clock_label.pack(side='right', padx=10)

# Continue (ngoài cùng bên trái của cụm)
continue_button = tk.Button(right_controls, text="Continue", font=("Arial", 10, "bold"), bg="#2196F3", fg="white",
                            command=continue_clock, width=10)
continue_button.pack(side='right', padx=5)

# ==== NÚT CONTACT ====
def contact_action():
    content = ""
    create_new_window_contact("Contact", content)
    reset_timer()
contact_button = tk.Button(left_button_frame, text="Contact", font=("Arial", 12, "bold"),
                           bg="#2196F3", fg="white", width=10, command=lambda: contact_action())
contact_button.pack(pady=5)

# ==== NÚT STATUS ====
def status_action():
    content = ""
    create_new_window_status("Status", content)
    reset_timer()
status_button = tk.Button(left_button_frame, text="Status", font=("Arial", 12, "bold"),
                          bg="#FF9800", fg="white", width=10, command=lambda: status_action())
status_button.pack(pady=5)

# ==== NÚT VÀO KHO ẢNH ====
def image_action():
    create_new_window_image("Image")
image_button = tk.Button(left_button_frame, text="Image", font=("Arial", 12, "bold"),
                          bg="#7c32d1", fg="white", width=10, command=lambda: create_new_window_image("Image"))
image_button.pack(pady=5)

# ==== NÚT NOTE ====
def note_action():
    create_new_window_note()
note_button = tk.Button(left_button_frame, text="Note", font=("Arial", 12, "bold"),
                          bg="#873e23", fg="white", width=10, command=lambda: note_action())
note_button.pack(pady=5)

# Bắt đầu cập nhật đồng hồ
update_clock()

# ==== CHẠY ỨNG DỤNG ====
root.mainloop()
