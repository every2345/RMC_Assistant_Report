from re import L
import tkinter as tk
from tkinter import ttk, messagebox
from unicodedata import category
import schedule
import threading
import time
import datetime
import gdown
import os
from PIL import Image, ImageTk
import pyperclip
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import requests
import subprocess
import re  

# Tạo cửa sổ chính
root = tk.Tk()
root.title("RMC Report Assistant")
root.geometry("1080x800")

# ==== Biến toàn cục cho box màu ====
boxes = [] # Danh sách các ô vuông
box_colors = ["white"] * 6
hint_label = None # Label để hiển thị gợi ý
box_filled = [False] * 6
first_box_filled = False

# ==== Cấu hình GG DRIVE API====
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
ROOT_CACHE_DIR = r"D:\RMC_Assistant\Cache"
ARCHIVE_DIR = os.path.join(ROOT_CACHE_DIR, "Documentary_archive")
os.makedirs(ARCHIVE_DIR, exist_ok=True)

# ==== Frame chính ====
main_frame = tk.Frame(root)
main_frame.pack(expand=True, pady=40, padx=20)

# Frame con chứa văn bản và các nút
content_frame = tk.Frame(main_frame)
content_frame.pack()

# === FRAME CHỨA CONTACT, NOTE VÀ STATUS BÊN TRÁI ===
left_button_frame = tk.Frame(content_frame)
left_button_frame.pack(side="left", fill="y", padx=10, pady=10)

# === Text để hiển thị văn bản ===
output_text = tk.Text(content_frame, font=("Arial", 13), width=60, height=12, wrap="word")
output_text.pack(side='left', pady=(10, 0), padx=10)
output_text.config(state='disabled')

# === Frame chứa các danh sách ATQB, ABDNC... ===
button_frame = tk.Frame(content_frame)
button_frame.pack(side='left', padx=10)

# === Frame chứa các item xuất hiện khi chọn danh sách ===
item_frame = tk.Frame(content_frame)
item_frame.pack(side='left', padx=10)

# ==== Frame cho NÚT COPY, CLEAR, COUNTINUE, CATCH ====
cccc_frame = tk.Frame(main_frame)
cccc_frame.pack(fill='x', pady=(10, 0), padx=20)

# ==== THÊM ĐỒNG HỒ ĐẾM NGƯỢC ==== 
timer_frame = tk.Frame(main_frame)
timer_frame.pack(pady=(10, 0))

timer_label = tk.Label(timer_frame, text="⏳Waiting Countdown⏳", font=("Arial", 16, "bold"), fg="blue")
timer_label.pack()

countdown_job = None
time_left = 300  # 5 phút = 300 giây

# === Biến điều khiển đồng hồ ===
is_running = True

# ==== Frame chứa các ô tô màu ====
box_frame = tk.Frame(main_frame)
box_frame.pack(pady=(10, 0))

# Tạo 6 ô trắng trong box_frame
for i in range(6):
    lbl = tk.Label(box_frame, width=10, height=1, bg="white", relief="solid", borderwidth=2)
    lbl.grid(row=0, column=i, padx=5)
    boxes.append(lbl)

# ==== TẠO hint_label nằm dưới box_frame ====
hint_label = tk.Label(main_frame, text="Quy trình xử lý sự cố đang đợi", wraplength=800, justify="left", font=("Arial", 11), fg="black")
hint_label.pack(pady=(10, 20))  # Giữa box_frame và nút xác nhận

# ==== Hàm xử lý tô màu ====
def on_category_click():
    global box_colors

    # Đếm số ô đã được tô xanh
    green_count = box_colors.count("green")

    # Gợi ý tương ứng từng bước
    if green_count == 1:
        update_hint("Đã ghi nhận sự cố, tiến hành báo cáo lên group chung và tiếp tục theo dõi sự cố đang diễn ra. Nếu trong vòng 5 phút, không có thông báo gì từ phía bên Site đang xảy ra lỗi lên group chung. Lập tức liên hệ vói Site theo danh sách đã cho dựa vào mức độ ưu tiên (Bấm xác nhận nếu như thông tin đã được cập nhật lên group từ bên Site). Sau khi đã liên hệ, cập nhật thông tin liên hệ lên Group chung thông qua biểu mẫu trong mục Contact.")
    elif green_count == 2:
        update_hint("Tiếp tục theo dõi và cập nhật sự cố liên tục. Nếu như sau một khoảng thời gian không nhận được thông tin gì từ phía bên Site kể từ thời điểm đã liên hệ với Site (1 - 2 tiếng) và đã thông tin lên group chung (). Tiến hành liên hệ lại với Site để xác minh tình trạng kiếm tra thiết bị và nguyên nhân (nếu có). Tiến hành cập nhập lại tình hình thiết bị lên nhóm group chung về tình hình khắc phục trình trạng hiện tại của thiết bị gây lỗi.")
    elif green_count == 3:
        update_hint("Nếu sự cố sau 1 tiếng cho đến 2 tiếng vẫn chưa được sự xử lí và cũng chưa được cập nhật lên group chung. Tiến hành liên hệ lại với số điện thoại ưu tiên để xác nhận lại sự cố, sau đó báo cáo lại tình hiên fleen group chung (Bấm 'Xác nhận' nếu sự cố đã được giải quyết trước thời điểm này).")
    elif green_count == 4:
        update_hint("Khi sự cố đã được giải quyết, báo cáo lên group chung để khách hàng và các bộ phận liên quan nắm thông tin (Bấm 'Xác nhận' nếu có trường hợp ngoại lệ xảy ra).")
    elif green_count == 5:
        update_hint("Cập nhật lên bảng Alarm List.")
    elif green_count == 6:
        update_hint("Toàn bộ các bước trong quy trình đã được hoàn tất, làm tốt lắm!")
        threading.Thread(target=reset_after_delay, daemon=True).start()

def reset_after_delay():
    time.sleep(5)
    for i in range(6):
        box_colors[i] = "white"
        box_filled[i] = False
        boxes[i].config(bg="white")
    global first_box_filled
    first_box_filled = False
    update_hint("Quy trình xử lý sự cố đang đợi")

def update_hint(text):
    if hint_label:
        hint_label.config(text=text)

# ==== Hàm xử lý bắt và tiếp tục đồng hồ ====
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

# ==== Hàm chức năng cho nút copy và nút clear ====
def copy_text_to_clipboard():
    text = output_text.get("1.0", "end-1c")
    pyperclip.copy(text)

def clear_text_output():
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.config(state='disabled')

# ==== HÀM tải flie từ google drive và Tạo thư mục để lưu các file được tải về từ GOOGLE DRIVE ====
def download_from_drive(file_id):

    # Đường dẫn tới file cache trong ổ D
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

# ==== Tìm credentials và token trong thư mục con từ file CACHE====
def find_auth_paths():
    folder_path = os.path.join(ROOT_CACHE_DIR, CREDENTIAL_FILE_ID)
    cred_dir = os.path.join(folder_path, "Credentials")
    token_dir = os.path.join(folder_path, "Token")
    os.makedirs(cred_dir, exist_ok=True)
    os.makedirs(token_dir, exist_ok=True)

    cred_path = os.path.join(cred_dir, "credentials.json")
    token_path = os.path.join(token_dir, "token.json")

    # Nếu chưa có credentials, tải từ Drive
    if not os.path.exists(cred_path):
        downloaded = download_from_drive(CREDENTIAL_FILE_ID, cred_dir, "credentials.json")
        if not downloaded:
            raise FileNotFoundError("Không tải được credentials.json")

    return cred_path, token_path

# ==== Xác thực Google Drive ====
def authenticate():
    cred_path, token_path = find_auth_paths()
    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())

    return creds, build('drive', 'v3', credentials=creds)

# ==== Lấy danh sách file từ Google Drive ====
def list_files(service):
    query = f"'{FOLDER_ID}' in parents and trashed = false"
    results = service.files().list(
        q=query,
        pageSize=100,
        fields="files(id, name)"
    ).execute()
    return results.get("files", [])

# ==== Tải file từ thư mục chứa tài liệu RMC trên Google Drive ====
def download_file(token, file_id, filename):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        filepath = os.path.join(ARCHIVE_DIR, filename)
        with open(filepath, "wb") as f:
            f.write(response.content)
    else:
        print("❌ Tải thất bại:", response.text)

# ====  Mở Folder theo đường dẫn bằng cách bấm nút ====
def open_archive_folder():
    if os.path.exists(ARCHIVE_DIR):
        subprocess.run(['explorer', ARCHIVE_DIR], shell=True)

# ==== Hàm chức năng cho đồng hồ đếm ngược ====
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
        
        content = ''.join(lines)
    except Exception as e:
        content = f"Không thể mở file: {e}"

    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, content)
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

# ==== CÁC mà đường dẫn DANH SÁCH FILE ID, hình ảnh và biểu mẫu báo cáo từ google drive ====
# ==== DỮ LIỆU ĐƯỢC TỔ CHỨC THEO KHU VỰC ====
category_images = {
    "ANVL": {
        "Sensor": {
            "1060jAOW78jUlcNTExPDveyG7sDN7gn3i": "KEF&KSF",
            "1kVfjuOyBmyz7y6vUG_kgRl6mf1Fck8K1": "DELICA",
            "10zRy0yQaT0UAqDA2lFCBflYBNTI3V_hB": "POWERSUPPLY1",
            "1ZwolrlhT4ulfuC5SdBdh5crxhR-1jWQ9": "FR&FC",
            "17Pv5XTXWSzU0W3JMWCehw4qhjWet7g0X": "POWERSUPPLY2",
            "1SFb0QQ84qB8q02shYxgGZWp0nhO4g4FQ": "POWERSUPPLY3"
        },
        "Layout": {
            "1moC360oUYfcZ8LNIojpdwDzHh79qmimX": "Layout_ANVL"
        },
        "Gateway": {
            "1aZE5oP7ngAyX54fOv0Exm4vDqG7lE-B9": "Gateway_1",
            "1CuqkedbFiK0I8EFPqs_i7b_-4oFWx27_": "Gateway_2"
        },
        "Alarm points": {
            "1UEmDpu5E42ZWLFiJE1a5l55nSNkDOw_u": "NVL_Delica"
        }
    },
    "BDNC": {
        "Sensor": {
            "14CrZrrWMdVDyV_rrGge2xWTWGOhXjX7Q": "AHU",
            "1nr9IRbU231jS8g32JTpNWzmOlKgMHdG3": "FAN&VRV",
            "1f2iiBlHBj4lQbNbVS04FLEzbkpnzXBIV": "FR&FC",
            "12wX89FsnvxTz1rTwmtfIOyYe2QnHjaLm": "LPG",
            "1Wdr23jmblvu9aveBGIUo_lfJOPUxR8Ou": "POWERSUPPLY"
        },
        "Layout": {
            "1tOoW_QYZ8Ns44KWRbOnolT9aKor-Qf7q": "Layout_BDNC"
        },
        "Gateway": {
            "1VHjuk-_5ZIb53o0e_gZBUHo5bDngU590": "Gateway_BDNC"
        },
        "Alarm points": {}
    },
    "TQB": {
        "Sensor": {
            "1FrMbcaBlT5P2XEsC4QE7E9fwUtocH9W4": "BAKERY",
            "1AiLD7HoCbgjt2VcmH1WUltPofGk4Fiey": "DELICA",
            "1DQCRERO-to6rtB9dyBFc6GMTXhxNb579": "FAN",
            "1RIl1cYtaQndVN-PRp2GvKs_Apbnvfxxv": "FR&FC",
            "1MSqa8FTV65vfdtdd9r0nw4oYk_PtC3H3": "SUSHI",
            "1cgGQZ6pWpkqFFgdlRysas6Sg3ufUWlNZ": "POWERSUPPLY"
        },
        "Layout": {
            "1OtfrUIaf4CmL3Slyf2ZspItNUFuDsbDC": "Layout_TQB"
        },
        "Gateway": {
            "1G5ABGCJw3OXJZOq42gSG3D3FXzjPWbOk": "Gateway_TQB"
        },
        "Alarm points": {
            "1CI_BRGdB9lQn6jYhI61gIGQmj0oKIIqq": "TQB_Sushi",
            "1ZxmaaIX3eV6Zv4HKuwcv6jOADmOFy0Wa": "TQB_Bakery"
        }
    }
}

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

category_images = {
    "ANVL": {
        "Sensor": {
            "1060jAOW78jUlcNTExPDveyG7sDN7gn3i": "KEF&KSF",
            "1kVfjuOyBmyz7y6vUG_kgRl6mf1Fck8K1": "DELICA",
            "10zRy0yQaT0UAqDA2lFCBflYBNTI3V_hB": "POWERSUPPLY1",
            "1ZwolrlhT4ulfuC5SdBdh5crxhR-1jWQ9": "FR&FC",
            "17Pv5XTXWSzU0W3JMWCehw4qhjWet7g0X": "POWERSUPPLY2",
            "1SFb0QQ84qB8q02shYxgGZWp0nhO4g4FQ": "POWERSUPPLY3"
        },
        "Layout": {
            "1moC360oUYfcZ8LNIojpdwDzHh79qmimX": "Layout_ANVL"
        },
        "Gateway": {
            "1aZE5oP7ngAyX54fOv0Exm4vDqG7lE-B9": "Gateway_1",
            "1CuqkedbFiK0I8EFPqs_i7b_-4oFWx27_": "Gateway_2"
        },
        "Alarm points": {
            "1UEmDpu5E42ZWLFiJE1a5l55nSNkDOw_u": "NVL_Delica"
        }
    },
    "BDNC": {
        "Sensor": {
            "14CrZrrWMdVDyV_rrGge2xWTWGOhXjX7Q": "AHU",
            "1nr9IRbU231jS8g32JTpNWzmOlKgMHdG3": "FAN&VRV",
            "1f2iiBlHBj4lQbNbVS04FLEzbkpnzXBIV": "FR&FC",
            "12wX89FsnvxTz1rTwmtfIOyYe2QnHjaLm": "LPG",
            "1Wdr23jmblvu9aveBGIUo_lfJOPUxR8Ou": "POWERSUPPLY"
        },
        "Layout": {
            "1tOoW_QYZ8Ns44KWRbOnolT9aKor-Qf7q": "Layout_BDNC"
        },
        "Gateway": {
            "1VHjuk-_5ZIb53o0e_gZBUHo5bDngU590": "Gateway_BDNC"
        },
        "Alarm points": {}
    },
    "TQB": {
        "Sensor": {
            "1FrMbcaBlT5P2XEsC4QE7E9fwUtocH9W4": "BAKERY",
            "1AiLD7HoCbgjt2VcmH1WUltPofGk4Fiey": "DELICA",
            "1DQCRERO-to6rtB9dyBFc6GMTXhxNb579": "FAN",
            "1RIl1cYtaQndVN-PRp2GvKs_Apbnvfxxv": "FR&FC",
            "1MSqa8FTV65vfdtdd9r0nw4oYk_PtC3H3": "SUSHI",
            "1cgGQZ6pWpkqFFgdlRysas6Sg3ufUWlNZ": "POWERSUPPLY"
        },
        "Layout": {
            "1OtfrUIaf4CmL3Slyf2ZspItNUFuDsbDC": "Layout_TQB"
        },
        "Gateway": {
            "1G5ABGCJw3OXJZOq42gSG3D3FXzjPWbOk": "Gateway_TQB"
        },
        "Alarm points": {
            "1CI_BRGdB9lQn6jYhI61gIGQmj0oKIIqq": "TQB_Sushi",
            "1ZxmaaIX3eV6Zv4HKuwcv6jOADmOFy0Wa": "TQB_Bakery"
        }
    }
}
# ==== ID GG DRIVE để lấy file credentials.json trên Google Drive ====
CREDENTIAL_FILE_ID = "11QGDt1o-1ABpnNpzmjz4IBcauTEBcxHj"

# ==== ID của thư mục chứa các file tài liệu RMC trên Google Drive ====
FOLDER_ID = '1XQNfRslvd-duF_VTkxkThVrk9n6vcv4T'

# ==== TẠO các CỬA SỔ MỚI ====
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

        if fill_box(1):  # Chỉ tô ô 1 nếu ô 0 đã được tô
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

        fill_box(2)  # Chỉ tô nếu ô 1 đã được tô

        new_window.destroy()

    ok_button = tk.Button(new_window, text="OK", font=("Arial", 12, "bold"),
                          bg="green", fg="white", command=handle_ok)
    ok_button.pack(pady=10)

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
        used_numbers = []
        for filename in os.listdir(DATA_DIR):
            if filename.startswith("reminders") and filename.endswith(".json"):
                try:
                    number = int(filename.replace("reminders", "").replace(".json", ""))
                    used_numbers.append(number)
                except:
                    continue
        count = 1
        while count in used_numbers:
            count += 1
        return count

    def update_stt_label():
        current_stt.set(str(len([f for f in os.listdir(DATA_DIR) if f.endswith(".json")])))

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

        # ==== Chuẩn hóa thời gian ====
        time_strs = time_input.split(",")
        normalized_times = []
        for t in time_strs:
            t = t.strip()
            if not t:
                continue
            try:
                h, m = map(int, t.split(":"))
                normalized_time = f"{h:02d}:{m:02d}"
                # Kiểm tra hợp lệ bằng datetime
                datetime.datetime.strptime(normalized_time, "%H:%M")
                normalized_times.append(normalized_time)
            except:
                messagebox.showerror("Lỗi", f"Thời gian không hợp lệ: {t}", parent=note_window)
                return

        # ==== Ngày ====
        if day_input.strip().lower() == "all":
            day_strs = [str(d) for d in range(1, 32)]
        else:
            day_strs = []
            for d in day_input.split(","):
                d = d.strip()
                if not d:
                    continue
                try:
                    val = int(d)
                    assert 1 <= val <= 31
                    day_strs.append(str(val))
                except:
                    messagebox.showerror("Lỗi", f"Ngày không hợp lệ: {d}", parent=note_window)
                    return

        # ==== Tháng ====
        if month_input.strip().lower() == "all":
            month_strs = [str(m) for m in range(1, 13)]
        else:
            month_strs = []
            for m in month_input.split(","):
                m = m.strip()
                if not m:
                    continue
                try:
                    val = int(m)
                    assert 1 <= val <= 12
                    month_strs.append(str(val))
                except:
                    messagebox.showerror("Lỗi", f"Tháng không hợp lệ: {m}", parent=note_window)
                    return
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

        times = normalized_times
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

        # Sort lại theo tên file reminders{n}.json để đảm bảo đúng STT thực tế
        data_list_sorted = sorted(data_list, key=lambda d: int(os.path.splitext(os.path.basename(d.get("_file", "reminders0.json")))[0].replace("reminders", "")))

        now = datetime.datetime.now()
        for i, item in enumerate(data_list_sorted, start=1):
            mode = item.get("mode", "")
            times = item.get("times", [])
            days = item.get("days", [])
            months = item.get("months", [])

            tag = ""
            if mode == "1 lần":
                expired = True
                for m in months:
                    for d in days:
                        for t in times:
                            try:
                                h, mn = map(int, t.split(":"))
                                scheduled_time = datetime.datetime(year=now.year, month=int(m), day=int(d), hour=h, minute=mn)
                                if scheduled_time >= now:
                                    expired = False
                                    break
                            except:
                                pass
                tag = "one_time_valid" if not expired else "one_time_expired"
            else:
                tag = "recurring"

            tree.insert("", tk.END, values=(
                i,
                item["keyword"],
                item["content"],
                ", ".join(item["times"]),
                ", ".join(item["days"]),
                ", ".join(item["months"]),
                mode
            ), tags=(tag,))

    def search_data():
        keyword = search_var.get().lower()
        filtered = [item for item in full_data if keyword in item["keyword"].lower() or keyword in item["content"].lower()]
        display_data(filtered)

    def refresh_data():
        global full_data
        full_data = load_all_json_files()
        display_data(full_data)
        update_stt_label()  # cập nhật STT ghi chú tiếp theo


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
        tk.Label(main_frame, text="Số ghi chú hiện tại:").pack()
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
        tree.tag_configure("one_time_valid", background="#d4fcd4")     # xanh lá
        tree.tag_configure("one_time_expired", background="#f8d4d4")   # đỏ
        tree.tag_configure("recurring", background="#d4eaff")          # xanh da trời
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

def create_new_window_image_daviteq(title):
    def show_image_by_ids(file_ids):
        for widget in image_frame.winfo_children():
            widget.destroy()

        for idx, file_id in enumerate(file_ids):
            img_path = download_from_drive(file_id)
            if img_path and not str(img_path).startswith("ERROR"):
                try:
                    img = Image.open(img_path)
                    img.thumbnail((100, 75), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)

                    row = (idx * 2) // max_columns
                    col = idx % max_columns

                    label_img = tk.Label(image_frame, image=photo, bg="white", cursor="hand2")
                    label_img.image = photo
                    label_img.grid(row=row, column=col, padx=5, pady=(5, 0))

                    image_name = next((group[file_id] for region in category_images.values() for group in region.values() if file_id in group), "Unknown Image")

                    label_text = tk.Label(image_frame, text=image_name, bg="white", font=("Arial", 9))
                    label_text.grid(row=row + 1, column=col, padx=5, pady=(0, 10))

                    label_img.bind("<Button-1>", lambda e, path=img_path: open_large_image(path))

                except Exception as e:
                    image_label.config(text=f"Lỗi xử lý ảnh: {e}", image='', bg="white")
                    break

    def open_large_image(img_path):
        try:
            img = Image.open(img_path)
            photo = ImageTk.PhotoImage(img)

            popup = tk.Toplevel()
            popup.title("Xem ảnh lớn")
            popup.configure(bg="white")

            lbl = tk.Label(popup, image=photo, bg="white")
            lbl.image = photo
            lbl.pack(padx=10, pady=10)

            btn = tk.Button(popup, text="Edit Image", command=lambda: copy_image_to_clipboard(img))
            btn.pack(pady=10)

        except Exception as e:
            print(f"Lỗi mở ảnh lớn: {e}")

    def copy_image_to_clipboard(img):
        try:
            img.show()
            pyperclip.copy("Image copied to clipboard!")
            print("Image copied to clipboard!")
        except Exception as e:
            print(f"Lỗi copy ảnh: {e}")

    new_window = tk.Toplevel()
    new_window.title(title)
    new_window.geometry("1000x550")
    new_window.configure(bg="white")

    left_frame = tk.Frame(new_window, width=150, bg="#f0f0f0")
    left_frame.pack(side="left", fill="y")

    sub_button_frame = tk.Frame(new_window, width=200, bg="#e8e8e8")
    sub_button_frame.pack(side="left", fill="y")

    right_frame = tk.Frame(new_window, bg="white")
    right_frame.pack(side="right", fill="both", expand=True)

    global image_frame
    image_frame = tk.Frame(right_frame, bg="white")
    image_frame.pack(expand=True)

    global image_label
    image_label = tk.Label(right_frame, bg="white", text="", font=("Arial", 14))
    image_label.pack()

    category_frames = {}
    parent_buttons = {}
    selected_sub_button = None
    selected_parent_button = None
    max_columns = 4

    def on_sub_button_click(btn_clicked, file_dict):
        nonlocal selected_sub_button
        for frame in category_frames.values():
            for widget in frame.winfo_children():
                if isinstance(widget, tk.Button):
                    widget.config(bg="white", fg="black")
        selected_sub_button = btn_clicked
        selected_sub_button.config(bg="#4CAF50", fg="white")
        show_image_by_ids(file_dict.keys())

    def toggle_sub_buttons(category_name):
        nonlocal selected_parent_button
        for btn in parent_buttons.values():
            btn.configure(bg="white", fg="black")
        selected_parent_button = parent_buttons[category_name]
        selected_parent_button.configure(bg="#4CAF50", fg="white")

        for cat, frame in category_frames.items():
            frame.pack_forget()
        category_frames[category_name].pack()

    for area, subcategories in category_images.items():
        parent_btn = tk.Button(left_frame, text=area, width=15, pady=5,
                               bg="white", fg="black", font=("Arial", 10, "bold"),
                               activebackground="#e0e0e0",
                               command=lambda a=area: toggle_sub_buttons(a))
        parent_btn.pack(pady=(10, 0))
        parent_buttons[area] = parent_btn

        sub_frame = tk.Frame(sub_button_frame, bg="#e8e8e8")
        category_frames[area] = sub_frame

        for sub_name, file_dict in subcategories.items():
            def make_sub_command(btn, files):
                return lambda: on_sub_button_click(btn, files)

            sub_btn = tk.Button(
                sub_frame, text=sub_name,
                width=15, pady=5,
                relief="raised",
                bg="white", fg="black",
                font=("Arial", 10, "bold"), bd=1,
                activebackground="#e0e0e0",
                command=make_sub_command(None, file_dict)
            )
            sub_btn.pack(padx=10, pady=3)
            sub_btn.config(command=make_sub_command(sub_btn, file_dict))


    first_category = list(category_images.keys())[0]
    toggle_sub_buttons(first_category)

def create_documentary_viewer(creds, service):
    files = list_files(service)
    filtered_files = files.copy()

    # ==== Hàm tách tag từ tên file ====
    def extract_tags(filename):
        # Tìm các tag nằm trong ngoặc đơn đầu tên file, ví dụ: (ADV)(GUIDE)
        tags = re.findall(r'\(([^)]+)\)', filename)
        return ", ".join(tags) if tags else "Khác"

    def update_table(*args):
        keyword = search_var.get().lower().strip()
        current_mode = mode.get()

        # Nếu đang ở chế độ tìm theo STT mà ô tìm kiếm rỗng, không làm gì cả
        if current_mode == "number" and keyword == "":
            return

        tree.delete(*tree.get_children())
        new_filtered = []

        if current_mode == "name":
            new_filtered = [f for f in files if keyword in f["name"].lower()]
        elif current_mode == "type":
            new_filtered = [f for f in files if keyword in extract_tags(f["name"]).lower()]
        elif current_mode == "number":
            if keyword.isdigit():
                idx = int(keyword)
                if 1 <= idx <= len(files):
                    new_filtered = [files[idx - 1]]
            else:
                # Không hợp lệ => hiển thị danh sách cũ (giữ nguyên)
                return
        else:
            new_filtered = files.copy()

        for idx, f in enumerate(new_filtered, start=1):
            tag_label = extract_tags(f["name"])
            filepath = os.path.join(ARCHIVE_DIR, f["name"])

            is_downloaded = os.path.exists(filepath)
            status_text = "✅ Đã tải" if is_downloaded else "❌ Chưa tải"
            tag = "downloaded" if is_downloaded else "not_downloaded"

            tree.insert("", "end", values=(idx, tag_label, f["name"], "⇩", status_text), tags=(tag,))

        filtered_files.clear()
        filtered_files.extend(new_filtered)

    def handle_download(event):
        selected = tree.focus()
        if not selected:
            return
        item = tree.item(selected)
        index = int(item['values'][0]) - 1
        file = filtered_files[index]
        download_file(creds.token, file['id'], file['name'])
        update_table()

    root = tk.Tk()
    root.title("📁 RMC DRIVE VIEWER")
    root.geometry("900x600")

    frame_search = tk.Frame(root)
    frame_search.pack(pady=5, padx=5, fill="x")

    search_var = tk.StringVar()
    entry_search = tk.Entry(frame_search, textvariable=search_var, font=("Arial", 12), width=50)
    entry_search.pack(side="left", padx=5)

    btn_refresh = tk.Button(frame_search, text="🔄 Làm mới", font=("Arial", 12), command=update_table)
    btn_refresh.pack(side="right", padx=5)

    btn_open_folder = tk.Button(frame_search, text="📂", font=("Arial", 12), command=open_archive_folder)
    btn_open_folder.pack(side="right", padx=5)

    entry_search.bind("<KeyRelease>", update_table)

    frame_filter = tk.Frame(root)
    frame_filter.pack(pady=5)

    mode = tk.StringVar(value="name")
    tk.Radiobutton(frame_filter, text="🔍 Tìm theo tên", variable=mode, value="name", command=update_table).pack(side="left", padx=10)
    tk.Radiobutton(frame_filter, text="🔍 Tìm theo loại", variable=mode, value="type", command=update_table).pack(side="left", padx=10)
    tk.Radiobutton(frame_filter, text="🔍 Tìm theo stt", variable=mode, value="number", command=update_table).pack(side="left", padx=10)

    frame_table = tk.Frame(root)
    frame_table.pack(pady=10, fill="both", expand=True)

    columns = ("STT", "Loại", "Tên", "Tải", "Trạng thái")
    tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=20)

    tree.heading("STT", text="STT")
    tree.column("STT", width=50, anchor="center")

    tree.heading("Loại", text="Loại")
    tree.column("Loại", width=120, anchor="center")

    tree.heading("Tên", text="Tên")
    tree.column("Tên", width=400, anchor="w")

    tree.heading("Tải", text="Tải về")
    tree.column("Tải", width=80, anchor="center")

    tree.heading("Trạng thái", text="Trạng thái")
    tree.column("Trạng thái", width=100, anchor="center")

    # Cấu hình màu thẻ
    tree.tag_configure("downloaded", background="#d0f0c0")  # Xanh nhạt
    tree.tag_configure("not_downloaded", background="#f7c6c7")  # Đỏ nhạt


    tree.pack(fill="both", expand=True)
    tree.bind("<Double-1>", handle_download)

    update_table()
    root.mainloop()

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

def handle_first_box_fill():
    global first_box_filled
    if not first_box_filled:
        boxes[0].config(bg="green")
        box_colors[0] = "green"
        box_filled[0] = True
        first_box_filled = True
        update_hint("Đã ghi nhận sự cố, tiến hành báo cáo lên group chung và tiếp tục theo dõi sự cố đang diễn ra...")
        return True
    return False

def fill_box(index):
    if index == 0 or box_filled[index - 1]:
        boxes[index].config(bg="green")
        box_colors[index] = "green"
        box_filled[index] = True
        return True
    else:
        messagebox.showwarning("Chưa hoàn tất bước trước", f"Vui lòng hoàn thành bước {index} trong quy trình xử lý sự cố trước khi tiếp tục.")
        return False

# ==== TẠO CÁC NÚT CHO DANH SÁCH ====
def toggle_sub_buttons(state, item_dict):
    if not state["visible"]:
        for label, file_id in item_dict.items():
            if "NO_ERROR" in label:
                def cmd(fid=file_id):
                    handle_first_box_fill()
                    show_text_from_drive(fid, is_no_error=True, start_timer_flag=False)
            else:
                def cmd(fid=file_id):
                    handle_first_box_fill()
                    show_text_from_drive(fid, start_timer_flag=True)
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

# === MỞ SẴN DANH SÁCH ATQB ===
toggle_sub_buttons(list1_state, list1_files)

# ==== Phân bổ các nút bấm trong khung ====
# Nhóm bên trái: Copy và Clear
left_controls = tk.Frame(cccc_frame)
left_controls.pack(side="left")

copy_button = tk.Button(left_controls, text="Copy", font=("Arial", 10, "bold"), bg="#4CAF50", fg="white",
                        command=copy_text_to_clipboard, width=15)
copy_button.pack(side="left", padx=(0, 5))

clear_button = tk.Button(left_controls, text="Clear", font=("Arial", 10, "bold"), bg="#f44336", fg="white",
                         command=clear_text_output, width=15)
clear_button.pack(side="left")

# Nhóm bên phải: Catch, Clock, Continue
right_controls = tk.Frame(cccc_frame)
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
    if fill_box(1):  # Chỉ tô nếu ô 0 đã tô
        create_new_window_contact("Contact")
        on_category_click()
        reset_timer()
contact_button = tk.Button(left_button_frame, text="Contact", font=("Arial", 12, "bold"),
                           bg="#2196F3", fg="white", width=10, command=lambda: contact_action())
contact_button.pack(pady=5)

# ==== NÚT STATUS ====
def status_action():
    if fill_box(2):  # Chỉ tô nếu ô 1 đã tô
        create_new_window_status("Status")
        on_category_click()
        reset_timer()
status_button = tk.Button(left_button_frame, text="Status", font=("Arial", 12, "bold"),
                          bg="#FF9800", fg="white", width=10, command=lambda: status_action())
status_button.pack(pady=5)

# ==== NÚT XÁC NHẬN HÀNH ĐỘNG ====
def confirm_action():
    for i in range(3, 6):
        if not box_filled[i]:
            if fill_box(i):
                on_category_click()
                break  # Đảm bảo chỉ tô một ô
            else:
                break  # Dừng lại nếu chưa đủ điều kiện
confirm_button = tk.Button(main_frame, text="Xác nhận", font=("Arial", 12, "bold"),
                           bg="#4CAF50", fg="white", command=confirm_action)
confirm_button.pack(pady=10)

# ==== NÚT NOTE ====
def note_action():
    create_new_window_note()
note_button = tk.Button(left_button_frame, text="Note", font=("Arial", 12, "bold"),
                        bg="#873e23", fg="white", width=10, command=lambda: note_action())
note_button.pack(pady=5)

# ==== NÚT TRUY CẬP KHO ẢNH DAVITEQ ====
def image_daviteq_action():
    create_new_window_image_daviteq("DAVITEQ")
image_daviteq_button = tk.Button(left_button_frame, text="DAVITEQ", font=("Arial", 12, "bold"),
                                 bg="#3fc4f3", fg="white", width=10, command=lambda: image_daviteq_action())
image_daviteq_button.pack(pady=5)

# ==== NÚT  VÀO THƯ VIỆN TÀI LIỆU RMC====
def rmc_drive_viewer_action():
    creds, service = authenticate()
    create_documentary_viewer(creds, service)

rmc_drive_viewer_button = tk.Button(left_button_frame, text="Document", font=("Arial", 12, "bold"),
                                    bg="#5A780B", fg="white", width=10, command=lambda: rmc_drive_viewer_action())
rmc_drive_viewer_button.pack(pady=5)


# Bắt đầu cập nhật đồng hồ
update_clock()

# ==== CHẠY ỨNG DỤNG ====
root.mainloop()
