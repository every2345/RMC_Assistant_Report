import tkinter as tk
from PIL import Image, ImageTk
import datetime
import os
from PIL import Image, ImageTk
import pyperclip
from tkinter import ttk
import shutil
import re
import requests
import msal
import base64
import time
import threading
from tkinter import ttk, messagebox
import json
import schedule

# ==== Khởi tạo Tkinter root trước ====
root = tk.Tk()
root.withdraw()   # Ẩn cửa sổ chính ban đầu 

# ==== Thiết lập và Cấu hình Azure AD, OneDrive, đường dẫn lưu trữ và hơn thế nữa =============================================================================================================
# == lINK ONNDRIVE OF REPORT FORM ==
nvl_report_form_share_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/REPORT%20FORM/NVL%20REPORT%20FORM"
tqb_report_form_share_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/REPORT%20FORM/TQB%20REPORT%20FORM"
bdnc_report_form_share_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/REPORT%20FORM/BDNC%20REPORT%20FORM"

# == lINK ONNEDRIVE OF HOTLINES AND CONTACT FORM ==
hotlines_and_confirm_form_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/HOTLINE_AND_CONFIRM_FORM"

# ===== KHU VỰC ẢNH DAVITEQ =====
# == GATEWAY == 
gateway_bdnc_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/GATEWAY/BDNC"
gateway_tqb_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/GATEWAY/TQB"
gateway_nvl_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/GATEWAY/NVL"

# == LAYOUT ==
layout_bdnc_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/LAYOUT/BDNC"
layout_tqb_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/LAYOUT/TQB"
layout_nvl_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/LAYOUT/NVL"

# == SENSOR ==
sensor_bdnc_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/SENSOR/BDNC"
sensor_tqb_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/SENSOR/TQB"
sensor_nvl_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/SENSOR/NVL"

# == ALARMPOINT ==
al_nvl_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/ALARM%20POINTS/NVL"
al_tqb_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DAVITEQ/IMAGE_%20ARCHIVE/ALARM%20POINTS/TQB"

# ===== LINK LƯU TRỮ CÁC TÀI LIỆU PDF =====
documentary_archive_url = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE/DOCUMENTARY"

# == Thông tin ID của ứng dụng Azure AD ==
CLIENT_ID = "ac4edccf-a8ee-41aa-bcc4-6603c4bebae1"
TENANT_ID = "5983a1d2-f46b-492d-a9b3-7e2f3609d20b"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_SCOPES = ["Files.Read"]
CACHE_DIR = r"D:\RMC_Assistant_ver1.1\Cache"
CACHE_FILE = os.path.join(CACHE_DIR, "token_cache.bin")

# ============ Đường dân local trên máy tính để lưu trữ cache ================
# == đường dẫn lưu trữ các biểu mẫu ==
REPORT_FORM_DIR = r"D:\RMC_Assistant_ver1.1\Report_Form_Cache"

# == đường dẫn lưu trữ các ghi chú ==
NOTE_ARCHIVE_DIR = r"D:\RMC_Assistant_ver1.1\NOTE"

# == đường dẫn lưu trữ các hình ảnh ==
IMAGE_LAYOUT_ARCHIVE_DIR = r"D:\RMC_Assistant_ver1.1\IMAGE\LAYOUT"
IMAGE_GATEWAY_ARCHIVE_DIR = r"D:\RMC_Assistant_ver1.1\IMAGE\GATEWAY"
IMAGE_SENSOR_ARCHIVE_DIR = r"D:\RMC_Assistant_ver1.1\IMAGE\SENSOR"
IMAGE_AL_ARCHIVE_DIR = r"D:\RMC_Assistant_ver1.1\IMAGE\ALARMPOINT"

# == đường dẫn lưu trữ các tài liệu ==
DOCUMENTARY_ARCHIVE_DIR = r"D:\RMC_Assistant_ver1.1\DOCUMENTARY"

# === Khu vực tạo các thư mục lưu trữ nếu chưa có ===
# Tạo thư mục lưu trữ cache
os.makedirs(CACHE_DIR, exist_ok=True)

# Tạo thư lục lưu trữ biểu mẫu
os.makedirs(REPORT_FORM_DIR, exist_ok=True)

#Tạo thư mục lưu trữ ghi chú
os.makedirs(NOTE_ARCHIVE_DIR, exist_ok=True)

# Tạo thư mục lưu trữ hình ảnh
os.makedirs(IMAGE_LAYOUT_ARCHIVE_DIR, exist_ok=True)
os.makedirs(IMAGE_GATEWAY_ARCHIVE_DIR, exist_ok=True)
os.makedirs(IMAGE_SENSOR_ARCHIVE_DIR, exist_ok=True)
os.makedirs(IMAGE_AL_ARCHIVE_DIR, exist_ok=True)

# Tạo thư mục lưu trữ tài liệu
os.makedirs(DOCUMENTARY_ARCHIVE_DIR, exist_ok=True)

# ==== Đăng nhập, tải file và xử lý OneDrive bằng Azure ===========================================================================
# == Cửa sổ đăng nhập Azure AD trên thiết bị mới ==
def show_device_login(flow):
    win = tk.Toplevel(root)   # ✅ gắn với root chính
    win.title("🔑 Đăng nhập Azure AD")
    win.geometry("500x300")
    win.grab_set()  # ✅ chặn tương tác ngoài login

    # Frame 1: Thông báo
    frame1 = tk.Frame(win, pady=10)
    frame1.pack(fill="x")
    tk.Label(frame1, text="Bạn đang đăng nhập trên một thiết bị mới",
             font=("Arial", 12, "bold"), fg="red").pack()

    # Frame 2: Văn bản truy cập
    frame2 = tk.Frame(win, pady=5)
    frame2.pack(fill="x")
    tk.Label(frame2, text="Truy cập vào liên kết dưới đây:", font=("Arial", 11)).pack()

    # Frame 3: Link
    frame3 = tk.Frame(win, pady=5)
    frame3.pack(fill="x")
    entry_link = tk.Entry(frame3, font=("Arial", 11), width=50)
    entry_link.insert(0, flow["verification_uri"])
    entry_link.pack(padx=10)

    # Frame 4: Văn bản nhập mã
    frame4 = tk.Frame(win, pady=5)
    frame4.pack(fill="x")
    tk.Label(frame4, text="Nhập mã sau vào trang web:", font=("Arial", 11)).pack()

    # Frame 5: Mã đăng nhập
    frame5 = tk.Frame(win, pady=5)
    frame5.pack(fill="x")
    entry_code = tk.Entry(frame5, font=("Arial", 14, "bold"), width=20, justify="center")
    entry_code.insert(0, flow["user_code"])
    entry_code.pack(padx=10)

    # Copy link và code
    def copy_link():
        win.clipboard_clear()
        win.clipboard_append(flow["verification_uri"])

    def copy_code():
        win.clipboard_clear()
        win.clipboard_append(flow["user_code"])

    btn_frame = tk.Frame(win, pady=10)
    btn_frame.pack()
    ttk.Button(btn_frame, text="📋 Copy link", command=copy_link).grid(row=0, column=0, padx=10)
    ttk.Button(btn_frame, text="📋 Copy mã", command=copy_code).grid(row=0, column=1, padx=10)

    return win

# ==== Hàm đăng nhập Azure AD ====
def authenticate():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()

    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            raise Exception("Không khởi tạo được Device Flow")

        win = show_device_login(flow)
        result_container = {"result": None}

        def do_login():
            result = app.acquire_token_by_device_flow(flow)
            result_container["result"] = result

        threading.Thread(target=do_login, daemon=True).start()

        def check_result():
            if result_container["result"] is not None:
                win.destroy()
            else:
                root.after(500, check_result)

        root.after(500, check_result)
        win.wait_window()
        result = result_container["result"]

    # ✅ Lưu lại cache sau khi login hoặc refresh thành công
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

    return result

# ==== Đăng nhập Azure ====
try:
    result = authenticate()
    if "access_token" not in result:
        raise Exception("Đăng nhập thất bại")
    access_token = result["access_token"]
except Exception as e:
    messagebox.showerror("Lỗi", str(e))
    root.destroy()
    exit()

## Sau một thời gian chương trình treo (idle) thì access_token hết hạn (thường là 1 giờ), nên không lấy được dữ liệu thường xuyên##
# ==== Quản lý phiên làm việc với Azure Graph API ====
class GraphSession:
    def __init__(self, client_id, authority, scopes, cache_file):
        self.client_id = client_id
        self.authority = authority
        self.scopes = scopes
        self.cache_file = cache_file
        self.cache = msal.SerializableTokenCache()
        if os.path.exists(cache_file):
            self.cache.deserialize(open(cache_file, "r").read())
        self.app = msal.PublicClientApplication(
            client_id, authority=authority, token_cache=self.cache
        )
        self.account = None
        self.token = None

    def save_cache(self):
        if self.cache.has_state_changed:
            with open(self.cache_file, "w") as f:
                f.write(self.cache.serialize())

    def ensure_token(self):
        """Đảm bảo luôn có access_token hợp lệ (refresh khi cần)."""
        # Nếu đã có token thì check hạn
        if self.token and "access_token" in self.token:
            expires_at = self.token.get("expires_on")
            if expires_at:
                import time
                if int(expires_at) > int(time.time()) + 60:
                    # Token còn sống > 60s thì dùng tiếp
                    return self.token["access_token"]

        # Nếu chưa có hoặc hết hạn → thử silent refresh
        accounts = self.app.get_accounts()
        if accounts:
            self.account = accounts[0]
            self.token = self.app.acquire_token_silent(self.scopes, account=self.account)

        # Nếu vẫn chưa có token hợp lệ → login lại
        if not self.token or "access_token" not in self.token:
            flow = self.app.initiate_device_flow(scopes=self.scopes)
            if "user_code" not in flow:
                raise Exception("Không khởi tạo được Device Flow")
            win = show_device_login(flow)
            result_container = {"result": None}

            def do_login():
                result = self.app.acquire_token_by_device_flow(flow)
                result_container["result"] = result

            threading.Thread(target=do_login, daemon=True).start()

            def check_result():
                if result_container["result"] is not None:
                    win.destroy()
                else:
                    root.after(500, check_result)

            root.after(500, check_result)
            win.wait_window()
            self.token = result_container["result"]

        if not self.token or "access_token" not in self.token:
            raise Exception("Đăng nhập Azure thất bại")

        self.save_cache()
        return self.token["access_token"]

# ==== Khởi tạo session Azure ====
graph_session = GraphSession(CLIENT_ID, AUTHORITY, GRAPH_SCOPES, CACHE_FILE)

# ==== Sử dụng trong các hàm API ====
# ==== Lấy danh sách file từ link chia sẻ ====
def list_files_from_url(share_url):
    token = graph_session.ensure_token()
    encoded_url = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    encoded_url = encoded_url.rstrip("=").replace("/", "_").replace("+", "-")

    url = f"https://graph.microsoft.com/v1.0/shares/u!{encoded_url}/driveItem/children"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        items = r.json().get("value", [])
        return [{"id": item["id"], "name": item["name"]} for item in items if "file" in item]
    else:
        return []

# ==== Tải file từ OneDrive ====
def download_file(file_id, filename):
    token = graph_session.ensure_token()
    cache_path = os.path.join(REPORT_FORM_DIR, filename)
    if os.path.exists(cache_path):
        return cache_path

    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers, stream=True)

    if r.status_code == 200:
        with open(cache_path, "wb") as f:
            for chunk in r.iter_content(1024):
                f.write(chunk)
        return cache_path
    else:
        return None

# === Khu vực tạo Frame để lưu trữ các thành phần ===============================================================================================
# Tạo cửa sổ chính
root.deiconify()
root.title("RMC Report Assistant")
root.geometry("1080x800")

# Frame chính
main_frame = tk.Frame(root)
main_frame.pack(expand=True, pady=40, padx=20)

# Frame con chứa văn bản và các nút bên phải
content_frame = tk.Frame(main_frame)
content_frame.pack()

# === frame chứa nút contact, status và note bên trái ===
left_button_frame = tk.Frame(content_frame)
left_button_frame.pack(side="left", fill="y", padx=10, pady=10)

# === Text để hiển thị văn bản ===
output_text = tk.Text(content_frame, font=("Arial", 13), width=60, height=15, wrap="word")
output_text.pack(side='left', pady=(10, 0), padx=10)
output_text.config(state='disabled')

# === Frame chứa các danh sách ATQB, ABDNC... với Scrollbar ===
button_container = tk.Frame(content_frame)
button_container.pack(side='left', padx=10, fill="y")

# Subframe cho thanh tìm kiếm (cố định, không cuộn)
button_search_frame = tk.Frame(button_container)
button_search_frame.pack(fill="x")

# Subframe cho phần cuộn danh sách nút
button_list_container = tk.Frame(button_container)
button_list_container.pack(fill="both", expand=True)

button_canvas = tk.Canvas(button_list_container, width=150, height=100)
button_canvas.pack(side="left", fill="both", expand=True)

button_scrollbar = tk.Scrollbar(
    button_list_container,
    orient="vertical",
    command=button_canvas.yview
)
button_scrollbar.pack(side="right", fill="y")

button_canvas.configure(yscrollcommand=button_scrollbar.set)
button_canvas.bind(
    '<Configure>',
    lambda e: button_canvas.configure(scrollregion=button_canvas.bbox("all"))
)

button_frame = tk.Frame(button_canvas)   # nơi chứa các nút cha
button_canvas.create_window((0, 0), window=button_frame, anchor="nw")

# === Frame chứa các item xuất hiện khi chọn danh sách ===
item_container = tk.Frame(content_frame)
item_container.pack(side='left', padx=10, fill="y")

# Subframe cho thanh tìm kiếm (cố định, không cuộn)
item_search_frame = tk.Frame(item_container)
item_search_frame.pack(fill="x")

# Subframe cho phần cuộn danh sách nút con
item_list_container = tk.Frame(item_container)
item_list_container.pack(fill="both", expand=True)

item_canvas = tk.Canvas(item_list_container, width=100, height=100)
item_canvas.pack(side="left", fill="both", expand=True)

item_scrollbar = tk.Scrollbar(item_list_container, orient="vertical", command=item_canvas.yview)
item_scrollbar.pack(side="right", fill="y")

item_canvas.configure(yscrollcommand=item_scrollbar.set)
item_canvas.bind('<Configure>', lambda e: item_canvas.configure(scrollregion=item_canvas.bbox("all")))

item_frame = tk.Frame(item_canvas)   # nơi chứa các nút con
item_canvas.create_window((0, 0), window=item_frame, anchor="nw")

# ==== NÚT COPY ====
copy_frame = tk.Frame(main_frame)
copy_frame.pack(fill='x', pady=(10, 0), padx=20)

# Nhóm bên phải: Catch, Clock, Continue
right_controls = tk.Frame(copy_frame)
right_controls.pack(side="right")

# Nhóm bên trái: Copy và Clear
left_controls = tk.Frame(copy_frame)
left_controls.pack(side="left")

# ==== Frame chứa các ô tô màu ====
box_frame = tk.Frame(main_frame)
box_frame.pack(pady=(10, 0))

# === Khu vực tạo và cấu hình chức năng ===========================================================================================
# ============== Chức năng tô màu ô tiến trình ==================
boxes = [] # Danh sách các ô vuông
box_colors = ["white"] * 6
hint_label = None # Label để hiển thị gợi ý
box_filled = [False] * 6
first_box_filled = False

# ==== Tạo 6 ô trắng trong box_frame ====
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

# ==== THANH TÌM KIẾM CHO DANH SÁCH CHA (button_frame) ====
search_parent_var = tk.StringVar()

def filter_parent_buttons(event=None):
    keyword = search_parent_var.get().strip().lower()

    matches = []
    non_matches = []

    for block, btn in parent_items:
        text = btn.cget("text").lower()
        if keyword == "" or keyword in text:
            matches.append((block, btn))
        else:
            non_matches.append((block, btn))

    # clear hết
    for block, btn in parent_items:
        block.pack_forget()

    # pack lại: matches lên đầu
    for block, btn in matches:
        block.pack(pady=10, anchor='w')
    for block, btn in non_matches:
        block.pack(pady=10, anchor='w')

    # cập nhật canvas
    button_canvas.update_idletasks()
    button_canvas.configure(scrollregion=button_canvas.bbox("all"))

search_parent_entry = tk.Entry(button_search_frame, textvariable=search_parent_var)
search_parent_entry.pack(fill="x", pady=5)
search_parent_entry.bind("<KeyRelease>", filter_parent_buttons)

parent_items = []  # list chứa các nút cha

# ==== THANH TÌM KIẾM CHO DANH SÁCH CON (item_frame) ====
search_child_var = tk.StringVar()

def filter_child_buttons(event=None):
    keyword = search_child_var.get().lower()

    # Chia danh sách thành 2 nhóm: khớp và không khớp
    matched = []
    unmatched = []
    for btn in child_buttons:
        if keyword in btn.cget("text").lower():
            matched.append(btn)
        else:
            unmatched.append(btn)

    # Clear layout trước
    for btn in child_buttons:
        btn.pack_forget()

    # Pack lại: matched trước, unmatched sau
    for btn in matched + unmatched:
        btn.pack(anchor='w', pady=1)

search_child_entry = tk.Entry(item_search_frame, textvariable=search_child_var)
search_child_entry.pack(fill="x", pady=5)
search_child_entry.bind("<KeyRelease>", filter_child_buttons)

child_buttons = []  # list chứa các nút con

# ==== Chức năng cho nút copy và nút clear văn bản đang hiển thị trên text box ====
def copy_text_to_clipboard():
    text = output_text.get("1.0", "end-1c")
    pyperclip.copy(text)

def clear_text_output():
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.config(state='disabled')

# == Chức năng bắt và tiếp tục đồng hồ thời gian thực của hệ thống ==
is_running = True

def catch_clock():
    global is_running
    is_running = False

def continue_clock():
    global is_running
    is_running = True

def update_clock():
    if is_running:
        now = datetime.datetime.now().strftime("%H:%M:%S")
        clock_label.config(text=now)
    root.after(1000, update_clock)

# == Chức năng lấy và hiển thị thời gian hiện tại trên hệ thống lên văn bản ==
def update_timer():
    global time_left, countdown_job
    minutes, seconds = divmod(time_left, 60)
    timer_label.config(text=f"⏳ Remain: {minutes:02d}:{seconds:02d}")
    if time_left > 0:
        time_left -= 1
        countdown_job = root.after(1000, update_timer)
    else:
        timer_label.config(text="⏰Contact Site⏰")

# == Chức năng bắt đầu và reset đồng hồ đếm ngược ==
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

# ==== thêm đồng hồ đếm ngược ==== 
timer_frame = tk.Frame(main_frame)
timer_frame.pack(pady=(10, 0))

timer_label = tk.Label(timer_frame, text="⏳Waiting Countdown⏳", font=("Arial", 16, "bold"), fg="blue")
timer_label.pack()

countdown_job = None
time_left = 300  # 5 phút = 300 giây

# === Hiển thị file văn bản từ OneDrive ===========================================================================
# ==== LẤY DANH SÁCH FILE ONE DRIVE THEO TÊN ====
def build_device_mapping(share_url, device_names):
    files = list_files_from_url(share_url)  # ✅ chỉ truyền share_url
    mapping = {}
    for dev in device_names:
        # Tìm file nào có tên chứa tên thiết bị (không phân biệt hoa thường)
        match = next((f for f in files if dev.lower() in f["name"].lower()), None)
        if match:
            mapping[dev] = match["id"]
    return mapping

# ==== CHỨC NĂNG HIỂN THỊ VĂN BẢN VÀ THỜI GIAN====
def show_text_from_drive(file_id, filename, is_no_error=False, start_timer_flag=True):
    try:
        file_path = download_file(file_id, filename)  # tải từ OneDrive
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

device_names_anvl = [
    "NVL_FR&FC", "NVL_FAN", "NVL_DELICA",
    "NVL_POWER_1", "NVL_POWER_2", "NVL_POWER_3", 
    "NVL_NO_ERROR"
]

device_names_atqb = [
    "TQB_FR&FC", "TQB_POWER", "TQB_FAN", 
    "TQB_SUSHI", "TQB_BAKERY", "TQB_DELICA",
    "TQB_NO_ERROR"
]

device_names_abnc = [
    "ABNC_FR&FC", "ABNC_POWER", "ABNC_FAN", "ABNC_LPG", "ABNC_NO_ERROR"
]

contact_sample = ["CONTACT_FORM"]
confirm_sample = ["CONFIRM_FORM"]

# Mapping riêng cho từng khu vực
nvl_report_form_files = build_device_mapping(nvl_report_form_share_url, device_names_anvl)
tqb_report_form_files = build_device_mapping(tqb_report_form_share_url, device_names_atqb)
bdnc_report_form_files = build_device_mapping(bdnc_report_form_share_url, device_names_abnc)

# ==== TẠO GIAO DIỆN DANH SÁCH ====
active_parent_button = None
active_child_button = None

def set_active_parent_button(btn):
    global active_parent_button
    # Reset nút cha cũ
    if active_parent_button and active_parent_button != btn:
        active_parent_button.config(bg="SystemButtonFace", fg="black")
    # Đổi màu nút cha mới
    btn.config(bg="green", fg="white")
    active_parent_button = btn

def set_active_child_button(btn):
    global active_child_button
    # Reset nút con cũ
    if active_child_button and active_child_button != btn:
        active_child_button.config(bg="SystemButtonFace", fg="black")
    # Đổi màu nút con mới
    btn.config(bg="blue", fg="white")
    active_child_button = btn

def create_list_block(parent, list_name, items, toggle_function, state):
    block_frame = tk.Frame(parent)
    block_frame.pack(pady=10, anchor='w')

    list_button = tk.Button(
        block_frame,
        text=list_name,
        font=("Arial", 14),
        width=12,
        command=lambda: [set_active_parent_button(list_button), toggle_function(state)]
    )
    list_button.pack(anchor='w')
    state["button"] = list_button

    # lưu cả frame và button
    parent_items.append((block_frame, list_button))

# ==== HÀM BẬT TẮT DANH SÁCH ====
def toggle_list1(state):
    if list2_state["visible"]:
        toggle_sub_buttons(list2_state, tqb_report_form_files)
    if list3_state["visible"]:
        toggle_sub_buttons(list3_state, bdnc_report_form_files)
    toggle_sub_buttons(state, nvl_report_form_files)

def toggle_list2(state):
    if list1_state["visible"]:
        toggle_sub_buttons(list1_state, nvl_report_form_files)
    if list3_state["visible"]:
        toggle_sub_buttons(list3_state, bdnc_report_form_files)
    toggle_sub_buttons(state, tqb_report_form_files)

def toggle_list3(state):
    if list1_state["visible"]:
        toggle_sub_buttons(list1_state, nvl_report_form_files)
    if list2_state["visible"]:
        toggle_sub_buttons(list2_state, tqb_report_form_files)
    toggle_sub_buttons(state, bdnc_report_form_files)

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

# ==== TÔ MÀU TIẾN TRÌNH ====
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

def make_cmd(fid, b, fname, is_no_error=False):
    def cmd():
        set_active_child_button(b)
        handle_first_box_fill()  # ✅ Bây giờ sẽ chỉ chạy khi bấm
        if is_no_error:
            show_text_from_drive(fid, fname, is_no_error=True, start_timer_flag=False)
        else:
            show_text_from_drive(fid, fname, start_timer_flag=True)
    return cmd

# ==== TẠO CÁC NÚT CHO DANH SÁCH ====
def toggle_sub_buttons(state, item_dict, auto_select_first=False):
    if not state["visible"]:
        first_child_btn = None
        for idx, (label, file_id) in enumerate(item_dict.items()):
            # ✅ Chỉ lấy phần thông tin quan trọng (bỏ prefix trước dấu _)
            short_label = label.split("_", 1)[-1] if "_" in label else label

            btn = tk.Button(item_frame, text=short_label, font=("Arial", 12))

            if "NO_ERROR" in label:
                cmd = lambda fid=file_id, b=btn, fname=label: [
                    set_active_child_button(b),
                    handle_first_box_fill(),
                    show_text_from_drive(fid, fname, is_no_error=True, start_timer_flag=False)
                ]
            else:
                cmd = lambda fid=file_id, b=btn, fname=label: [
                    set_active_child_button(b),
                    handle_first_box_fill(),
                    show_text_from_drive(fid, fname, start_timer_flag=True)
                ]
            btn.config(command=cmd)
            btn.pack(anchor='w', pady=1)
            child_buttons.append(btn)  # Thêm vào danh sách nút con
            state["buttons"].append(btn)

            if idx == 0:
                first_child_btn = btn
        state["visible"] = True

        if auto_select_first and first_child_btn:
            set_active_parent_button(state["button"])
            set_active_child_button(first_child_btn)
    else:
        for btn in state["buttons"]:
            btn.pack_forget()
        state["buttons"].clear()
        state["visible"] = False

# ==== TRẠNG THÁI ====
list1_state = {"visible": False, "buttons": [], "indicator_canvas": None, "indicator_id": None}
list2_state = {"visible": False, "buttons": [], "indicator_canvas": None, "indicator_id": None}
list3_state = {"visible": False, "buttons": [], "indicator_canvas": None, "indicator_id": None}

# ==== TẠO DANH SÁCH GIAO DIỆN ====
create_list_block(button_frame, "ANVL", nvl_report_form_files, toggle_list1, list1_state)
create_list_block(button_frame, "ATQB", tqb_report_form_files, toggle_list2, list2_state)
create_list_block(button_frame, "ABNC", bdnc_report_form_files, toggle_list3, list3_state)
toggle_sub_buttons(list1_state, nvl_report_form_files, auto_select_first=True)

# ==== khu vực tạo các cửa sổ chức năng ================================================================================================================================

# == Cửa sổ cho mục contact ==
def create_new_window_contact(title, content=None):
    new_window = tk.Toplevel(root)
    new_window.title(title)
    new_window.geometry("600x400")

    confirm_var = tk.StringVar(value="confirmed")

    # Tình trạng confirm
    confirm_frame = tk.LabelFrame(new_window, text="Tình trạng confirm", font=("Arial", 12, "bold"))
    confirm_frame.pack(padx=20, pady=10, fill="x")

    def toggle_entry_fields():
        if confirm_var.get() == "not_confirmed":
            dept_entry.config(state="normal")
            device_entry.config(state="normal")
            status_entry.config(state="readonly")   # combobox enable
            desc_entry.config(state="normal")
        else:
            dept_entry.config(state="disabled")
            device_entry.config(state="disabled")
            status_entry.config(state="disabled")   # combobox disable
            desc_entry.config(state="disabled")

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

    status_entry = ttk.Combobox(
        form_frame,
        font=("Arial", 11),
        state="disabled",   # ban đầu disable
        values=["Đang xử lý", "Đã xử lý", "Chờ xử lý", "Không chọn"]
    )
    status_entry.grid(row=2, column=1, pady=5, sticky="ew")

    # Đặt giá trị mặc định
    status_entry.set("Không chọn")

    tk.Label(form_frame, text="Mô tả:", font=("Arial", 11)).grid(row=3, column=0, sticky="nw", pady=5)
    desc_entry = tk.Text(form_frame, font=("Arial", 11), height=5, width=40, state="disabled")
    desc_entry.grid(row=3, column=1, pady=5, sticky="ew")

    form_frame.columnconfigure(1, weight=1)
    toggle_entry_fields()

    def handle_ok():
        if confirm_var.get() != "not_confirmed":
            new_window.destroy()
            return

        dept = dept_entry.get().strip()
        device = device_entry.get().strip()
        status_val = status_entry.get().strip()
        desc = desc_entry.get("1.0", tk.END).strip()

        try:
            # === Lấy danh sách file trong thư mục OneDrive ===
            files = list_files_from_url(hotlines_and_confirm_form_url)

            # === Tìm file confirm ===
            target_name = next(iter(contact_sample))  # ví dụ "CONFIRM_FORM"
            target_file = next((f for f in files if target_name in f["name"]), None)

            if not target_file:
                raise FileNotFoundError(f"Không tìm thấy file chứa '{target_name}' trong thư mục OneDrive.")

            file_id = target_file["id"]
            filename = target_file["name"]

            # === Tải file về ===
            file_path = download_file(file_id, filename)

            if not file_path or not os.path.exists(file_path):
                raise FileNotFoundError("File confirm không tồn tại sau khi tải.")

            # === Đọc file confirm template ===
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            replaced_lines = []
            for line in lines:
                original_line = line

                line = line.replace("[title]", dept)
                line = line.replace("[device]", device)
                line = line.replace("[status]", status_val)
                line = line.replace("[description]", desc)

                stripped_line = line.strip()

                if ("[title]" in original_line and not dept) or \
                   ("[device]" in original_line and not device) or \
                   ("[status]" in original_line and not status_val) or \
                   ("[description]" in original_line and not desc) or \
                   not stripped_line:
                    continue

                replaced_lines.append(stripped_line)

            content = '\n'.join(replaced_lines)

        except Exception as e:
            content = f"Lỗi khi xử lý file confirm: {e}"

        # Hiển thị ra output_text
        output_text.config(state='normal')
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, content)
        output_text.config(state='disabled')

        if fill_box(1):
            start_timer()

        new_window.destroy()

    # Nút OK
    ok_button = tk.Button(new_window, text="OK", font=("Arial", 12, "bold"),
                          bg="green", fg="white", command=handle_ok)
    ok_button.pack(pady=10)

# == Cửa sổ mục status ==
def create_new_window_status(title, content=None):
    new_window = tk.Toplevel(root)
    new_window.title(title)
    new_window.geometry("600x500")

    confirm_var = tk.StringVar(value="confirmed")

    # Tình trạng confirm
    confirm_frame = tk.LabelFrame(new_window, text="Đã confirm chưa?", font=("Arial", 12, "bold"))
    confirm_frame.pack(padx=20, pady=10, fill="x")

    def toggle_entry_fields():
        if confirm_var.get() == "not_confirmed":
            dept_entry.config(state="normal")
            device_entry.config(state="normal")
            status_entry.config(state="readonly")
            start_time_entry.config(state="normal")
            end_time_entry.config(state="normal")
            desc_entry.config(state="normal")
        else:
            dept_entry.config(state="disabled")
            device_entry.config(state="disabled")
            status_entry.config(state="readonly")
            start_time_entry.config(state="disabled")
            end_time_entry.config(state="disabled")
            desc_entry.config(state="disabled")

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

    status_entry = ttk.Combobox(
        form_frame,
        font=("Arial", 11),
        state="disabled",   # ban đầu disable
        values=["Alarm - Chưa xử lý", "Alarm - Đã xử lý", "Alarm - Chờ xử lý", "Normal - Đã xử lý", "Normal - Chờ xử lý", "Không chọn"]
    )
    status_entry.grid(row=2, column=1, pady=5, sticky="ew")

    # Đặt giá trị mặc định
    status_entry.set("Không chọn")

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

        # === Tính thời gian xử lý ===
        try:
            if start_time_str and end_time_str:
                fmt = "%H:%M"
                start_dt = datetime.datetime.strptime(start_time_str, fmt)
                end_dt = datetime.datetime.strptime(end_time_str, fmt)
                diff_minutes = int((end_dt - start_dt).total_seconds() / 60)
                if diff_minutes < 0:
                    diff_minutes += 24 * 60  # xử lý qua ngày
                time = f"{diff_minutes} phút ({start_time_str} - {end_time_str})"
            else:
                time = ""
        except ValueError:
            time = ""

        # === Tải file confirm từ OneDrive ===
        # === Lấy danh sách file trong thư mục OneDrive ===
        files = list_files_from_url(hotlines_and_confirm_form_url)

        # === Tìm file confirm ===
        target_name = next(iter(confirm_sample))  # "CONFIRM_FORM"
        target_file = next((f for f in files if target_name in f["name"]), None)

        if not target_file:
            raise FileNotFoundError(f"Không tìm thấy file chứa '{target_name}' trong thư mục OneDrive.")

        file_id = target_file["id"]
        filename = target_file["name"]

        # === Tải file về ===
        file_path = download_file(file_id, filename)

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
            content = f"Lỗi khi xử lý file confirm: {e}"

        # === Hiển thị nội dung ra output_text ===
        output_text.config(state='normal')
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, content)
        output_text.config(state='disabled')

        # Tô ô 2 nếu ô 1 đã được tô
        fill_box(2)

        new_window.destroy()

    ok_button = tk.Button(new_window, text="OK", font=("Arial", 12, "bold"),
                          bg="green", fg="white", command=handle_ok)
    ok_button.pack(pady=10)

# == Cửa sổ mục note ==
def create_new_window_note():
    # Thư mục lưu dữ liệu
    DATA_DIR = NOTE_ARCHIVE_DIR
    os.makedirs(DATA_DIR, exist_ok=True)

    # === Luồng kế hoạch ===
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

    def schedule_reminder(keyword, content, times, days, months, mode, file_path=None, delete_mode="delete"):
        for t in times:
            def job(t=t):
                now = datetime.datetime.now()
                if str(now.day) in days and str(now.month) in months:
                    # Hiển thị thông báo đúng luồng giao diện
                    def show_popup():
                        messagebox.showinfo(f"Thông báo: {keyword}", f"[{t}] {content}")

                        if mode == "1 lần" and file_path and os.path.exists(file_path):
                            try:
                                with open(file_path, "r", encoding="utf-8") as f:
                                    data = json.load(f)

                                # ✅ Nếu chọn delete → xóa file
                                if delete_mode == "delete":
                                    os.remove(file_path)
                                else:
                                    # ✅ Nếu chọn keep → chỉ update "done": True
                                    if isinstance(data, dict):
                                        data["done"] = True
                                        with open(file_path, "w", encoding="utf-8") as f:
                                            json.dump(data, f, ensure_ascii=False, indent=4)

                            except Exception as e:
                                print(f"Lỗi xử lý file {file_path}: {e}")

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
        mode = intensity_var.get()

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

        reminder_data = {
            "keyword": keyword,                     #  Từ khóa nhắc
            "content": content,                     #  Nội dung của nhắc
            "times": times,                         #  Dữ liệu giờ phút (HH:MM)
            "days": days,                           #  Dữ liệu ngày
            "months": months,                       #  Dữ liệu tháng 
            "mode": mode,                           #  Loai nhắc ("1 lần" hoặc "Cố định")
            "delete_mode": delete_mode_var.get(),   #  thêm lựa chọn delete/keep (Chỉ dành cho note nhắc 1 lần)
            "done": False                           #  đánh dấu đã nhắc hay chưa
        }
        save_reminder_to_new_file(reminder_data)
        file_path = os.path.join(DATA_DIR, f"reminders{get_next_stt()-1}.json")
        schedule_reminder(keyword, content, times, days, months, mode, file_path, delete_mode_var.get())
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
                                if isinstance(data, dict) and "keyword" in data:
                                    data["_file"] = file_path
                                    if "delete_mode" not in data:
                                        data["delete_mode"] = "delete"   # mặc định là xóa sau khi đã nhắc 1 lần 
                                    if "done" not in data:
                                        data["done"] = False             # Không xóa sau khi đã nhắc 1 lần 
                                    all_data.append(data)
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
    note_window.geometry("1000x500")

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

        global keyword_entry, content_entry, time_entry, day_entry, month_entry, intensity_var, delete_mode_var, delete_frame

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
        mode_combo = ttk.Combobox(main_frame, textvariable=intensity_var, values=["1 lần", "Cố định"])
        mode_combo.pack(fill="x", padx=10)

        # --- Frame chứa tick chọn ---
        delete_frame = tk.LabelFrame(main_frame, text="Tùy chọn khi đã nhắc (chỉ cho loại nhắc 1 lần", padx=5, pady=5)
        delete_frame.pack(fill="x", padx=10, pady=5)

        delete_mode_var = tk.StringVar(value="delete")  # mặc định là xóa sau nhắc

        rb_delete = tk.Radiobutton(
            delete_frame, text=" Xóa khi đã nhắc", variable=delete_mode_var, value="delete"
        )
        rb_keep = tk.Radiobutton(
            delete_frame, text=" Không xóa khi đã nhắc", variable=delete_mode_var, value="keep"
        )

        rb_delete.pack(side="left", padx=5)
        rb_keep.pack(side="left", padx=5)

        # Hàm enable/disable frame dựa trên mode
        def update_delete_frame_state(*args):
            if intensity_var.get() == "1 lần":
                for child in delete_frame.winfo_children():
                    child.configure(state="normal")
            else:
                for child in delete_frame.winfo_children():
                    child.configure(state="disabled")

        # Gán sự kiện thay đổi mode
        intensity_var.trace_add("write", update_delete_frame_state)

        # gọi 1 lần ban đầu để set trạng thái đúng
        update_delete_frame_state()

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
        tree.column("STT", width=10, anchor="center")

        tree.heading("Từ khóa", text="Từ khóa")
        tree.column("Từ khóa", width=60, anchor="center")

        tree.heading("Nội dung", text="Nội dung")
        tree.column("Nội dung", width=400, anchor="w")

        tree.heading("Thời gian", text="Thời gian")
        tree.column("Thời gian", width=50, anchor="center")

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
            reminder.get("_file"),                  # thêm đường dẫn file
            reminder.get("delete_mode", "delete")   # truyền delete_mode (mặc định delete nếu không có)
        )

# == Cửa sổ hình ảnh các thiết bị của Daviteq tương ứng với từng khu vực cũng như layout phân bổ thiết bị ở khu vực đó  ==
def create_new_window_image_daviteq(title):
    def show_images(file_list, category):
        for widget in image_frame.winfo_children():
            widget.destroy()

        for idx, file in enumerate(file_list):
            # Xác định thư mục lưu dựa vào category
            if category == "LAYOUT":
                save_dir = IMAGE_LAYOUT_ARCHIVE_DIR
            elif category == "GATEWAY":
                save_dir = IMAGE_GATEWAY_ARCHIVE_DIR
            elif category == "SENSOR":
                save_dir = IMAGE_SENSOR_ARCHIVE_DIR
            elif category == "ALARMPOINT":
                save_dir = IMAGE_AL_ARCHIVE_DIR
            else:
                save_dir = "."

            os.makedirs(save_dir, exist_ok=True)

            # Đường dẫn file ảnh lưu về
            local_path = os.path.join(save_dir, file["name"])

            # Nếu chưa có thì tải về
            if not os.path.exists(local_path):
                img_path = download_file(file["id"], local_path)
            else:
                img_path = local_path

            if img_path and not str(img_path).startswith("ERROR"):
                try:
                    img = Image.open(img_path)
                    img.thumbnail((100, 75), Image.Resampling.LANCZOS) #Thu nhỏ khi hiển thị thumbnail trong giao diện
                    photo = ImageTk.PhotoImage(img)

                    row = (idx * 2) // max_columns
                    col = idx % max_columns

                    label_img = tk.Label(image_frame, image=photo, bg="white", cursor="hand2")
                    label_img.image = photo
                    label_img.grid(row=row, column=col, padx=5, pady=(5, 0))

                    label_text = tk.Label(image_frame, text=file["name"], bg="white", font=("Arial", 9))
                    label_text.grid(row=row + 1, column=col, padx=5, pady=(0, 10))

                    label_img.bind("<Button-1>", lambda e, path=img_path: open_large_image(path))

                except Exception as e:
                    image_label.config(text=f"Lỗi xử lý ảnh: {e}", image='', bg="white")
                    break

    def open_large_image(img_path): #Scale ảnh, giảm 50% tỷ lệ thực tế của ảnh hiển thị 
        try:
            img = Image.open(img_path)

            #Scale ảnh theo tỷ lệ 1:2 (giảm 50%)
            scale_factor = 0.5
            new_size = (int(img.width * scale_factor), int(img.height * scale_factor))
            img_resized = img.resize(new_size, Image.Resampling.LANCZOS)

            photo = ImageTk.PhotoImage(img_resized)

            popup = tk.Toplevel()
            popup.title("DAVITEQ IMAGE DATA (Scaled 1:2)")
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

    def on_sub_button_click(btn_clicked, file_list, category):
        nonlocal selected_sub_button
        for frame in category_frames.values():
            for widget in frame.winfo_children():
                if isinstance(widget, tk.Button):
                    widget.config(bg="white", fg="black")
        selected_sub_button = btn_clicked
        selected_sub_button.config(bg="#4CAF50", fg="white")
        show_images(file_list, category)

    def toggle_sub_buttons(category_name):
        nonlocal selected_parent_button
        for btn in parent_buttons.values():
            btn.configure(bg="white", fg="black")
        selected_parent_button = parent_buttons[category_name]
        selected_parent_button.configure(bg="#247985", fg="white")

        for cat, frame in category_frames.items():
            frame.pack_forget()
        category_frames[category_name].pack()

    # ✅ Lấy dữ liệu từ OneDrive (thay Google Drive)
    category_images = {
        "GATEWAY": {
            "BDNC": list_files_from_url(gateway_bdnc_url),
            "TQB": list_files_from_url(gateway_tqb_url),
            "NVL": list_files_from_url(gateway_nvl_url),
        },
        "LAYOUT": {
            "BDNC": list_files_from_url(layout_bdnc_url),
            "TQB": list_files_from_url(layout_tqb_url),
            "NVL": list_files_from_url(layout_nvl_url),
        },
        "SENSOR": {
            "BDNC": list_files_from_url(sensor_bdnc_url),
            "TQB": list_files_from_url(sensor_tqb_url),
            "NVL": list_files_from_url(sensor_nvl_url),
        },
        "ALARMPOINT": {
            "TQB": list_files_from_url(al_tqb_url),
            "NVL": list_files_from_url(al_nvl_url),
        }
    }

    # ==== Sinh nút cha / con từ category_images ====
    for area, subcategories in category_images.items():
        parent_btn = tk.Button(left_frame, text=area, width=15, pady=5,
                               bg="white", fg="black", font=("Arial", 10, "bold"),
                               activebackground="#e0e0e0",
                               command=lambda a=area: toggle_sub_buttons(a))
        parent_btn.pack(pady=(10, 0))
        parent_buttons[area] = parent_btn

        sub_frame = tk.Frame(sub_button_frame, bg="#e8e8e8")
        category_frames[area] = sub_frame

        for sub_name, file_list in subcategories.items():
            def make_sub_command(btn, fl, cat=area):
                return lambda: on_sub_button_click(btn, fl, cat)

            sub_btn = tk.Button(
                sub_frame, text=sub_name,
                width=15, pady=5,
                relief="raised",
                bg="white", fg="black",
                font=("Arial", 10, "bold"), bd=1,
                activebackground="#e0e0e0"
            )
            sub_btn.pack(padx=10, pady=3)
            sub_btn.config(command=make_sub_command(sub_btn, file_list))

    # Chọn danh mục đầu tiên để hiển thị
    first_category = list(category_images.keys())[0]
    toggle_sub_buttons(first_category)

# == Cửa sổ hiển thị tài liệu của RMC ==
def create_documentary_viewer(share_url):
    files = list_files_from_url(share_url)  # Lấy file từ OneDrive Azure
    filtered_files = files.copy()

    # ==== Hàm tách tag từ tên file ====
    def extract_tags(filename):
        tags = re.findall(r'\(([^)]+)\)', filename)
        return ", ".join(tags) if tags else "Khác"

    def update_table(*args):
        keyword = search_var.get().lower().strip()
        current_mode = mode.get()

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
                return
        else:
            new_filtered = files.copy()

        for idx, f in enumerate(new_filtered, start=1):
            tag_label = extract_tags(f["name"])
            filepath = os.path.join(DOCUMENTARY_ARCHIVE_DIR, f["name"])

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

        # Gọi hàm download_file gốc (không sửa)
        temp_path = download_file(file['id'], file['name'])
        if not temp_path:
            messagebox.showerror("Lỗi", f"Tải file {file['name']} thất bại!")
            return

        # Đích: thư mục DOCUMENTARY_ARCHIVE_DIR
        final_path = os.path.join(DOCUMENTARY_ARCHIVE_DIR, file['name'])
    
        # Nếu file đã tồn tại thì ghi đè
        try:
            shutil.move(temp_path, final_path)  # chuyển sang thư mục đích
        except Exception as e:
            shutil.copy2(temp_path, final_path)  # fallback copy
            os.remove(temp_path)
        update_table()

    root = tk.Tk()
    root.title("📁 RMC DRIVE VIEWER (OneDrive - Azure)")
    root.geometry("900x600")

    # ==== Thanh tìm kiếm ====
    frame_search = tk.Frame(root)
    frame_search.pack(pady=5, padx=5, fill="x")

    search_var = tk.StringVar()
    entry_search = tk.Entry(frame_search, textvariable=search_var, font=("Arial", 12), width=50)
    entry_search.pack(side="left", padx=5)

    btn_refresh = tk.Button(frame_search, text="🔄 Làm mới", font=("Arial", 12), command=update_table)
    btn_refresh.pack(side="right", padx=5)

    btn_open_folder = tk.Button(frame_search, text="📂", font=("Arial", 12),
                                command=lambda: os.startfile(DOCUMENTARY_ARCHIVE_DIR))
    btn_open_folder.pack(side="right", padx=5)

    entry_search.bind("<KeyRelease>", update_table)

    # ==== Bộ lọc tìm kiếm ====
    frame_filter = tk.Frame(root)
    frame_filter.pack(pady=5)

    mode = tk.StringVar(value="name")
    tk.Radiobutton(frame_filter, text="🔍 Tìm theo tên", variable=mode, value="name", command=update_table).pack(side="left", padx=10)
    tk.Radiobutton(frame_filter, text="🔍 Tìm theo loại", variable=mode, value="type", command=update_table).pack(side="left", padx=10)
    tk.Radiobutton(frame_filter, text="🔍 Tìm theo stt", variable=mode, value="number", command=update_table).pack(side="left", padx=10)

    # ==== Bảng file ====
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

# === Khu vực tạo các thành phần =======================================================================================================================================
copy_button = tk.Button(left_controls, text="Copy", font=("Arial", 10, "bold"), bg="#4CAF50", fg="white",
                        command=copy_text_to_clipboard, width=15)
copy_button.pack(side="left", padx=(0, 5))

clear_button = tk.Button(left_controls, text="Clear", font=("Arial", 10, "bold"), bg="#f44336", fg="white",
                         command=clear_text_output, width=15)
clear_button.pack(side="left")

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

# ==== NÚT NOTE ====
def note_action():
    create_new_window_note()
note_button = tk.Button(left_button_frame, text="Note", font=("Arial", 12, "bold"),
                        bg="#873e23", fg="white", width=10, command=lambda: note_action())
note_button.pack(pady=5)

# ==== NÚT KHO ẢNH DAVITEQ ====
def image_daviteq_action():
    create_new_window_image_daviteq("DAVITEQ")
image_daviteq_button = tk.Button(left_button_frame, text="DAVITEQ", font=("Arial", 12, "bold"),
                                 bg="#3fc4f3", fg="white", width=10, command=lambda: image_daviteq_action())
image_daviteq_button.pack(pady=5)

# ==== NÚT VAOF KHO DOCUMENTARY ====
def rmc_drive_viewer_action():
    create_documentary_viewer(documentary_archive_url)

rmc_drive_viewer_button = tk.Button(
    left_button_frame,
    text="Document",
    font=("Arial", 12, "bold"),
    bg="#5A780B", fg="white",
    width=10,
    command=rmc_drive_viewer_action
)
rmc_drive_viewer_button.pack(pady=5)

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

# Bắt đầu cập nhật đồng hồ
update_clock()

# ==== CHẠY ỨNG DỤNG ====
root.mainloop()
