import tkinter as tk
from PIL import Image, ImageTk
import datetime
import os
import sys
import subprocess
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
import datetime 
from tkcalendar import DateEntry


# ==== Khởi tạo Tkinter root trước ====
root = tk.Tk()
root.withdraw()   # Ẩn cửa sổ chính ban đầu 

# ==== Thiết lập và Cấu hình Azure AD, OneDrive, đường dẫn lưu trữ và hơn thế nữa =============================================================================================================
BASE_URL = "https://aeondelight-my.sharepoint.com/personal/phuc_nguyen_aeondelight_biz/Documents/PHUC/PHUC/AZURE/RMC%20DATA%20STORAGE"

# ============= LINK ONEDRIVE OF REPORT FORM ===============
# = FOR AEONMALL ==
##AEON NGUYEN VAN LINH
nvl_report_form_share_url   = f"{BASE_URL}/REPORT%20FORM/NVL%20REPORT%20FORM"
##AEON TA QUANG BUU
tqb_report_form_share_url   = f"{BASE_URL}/REPORT%20FORM/TQB%20REPORT%20FORM"
##AEON BINH DUONOG NEW CITY
bdnc_report_form_share_url  = f"{BASE_URL}/REPORT%20FORM/BDNC%20REPORT%20FORM"
##AEON VAN GIANG
vg_report_form_share_url    = f"{BASE_URL}/REPORT%20FORM/VG%20REPORT%20FORM"
##AEON MIDORI PARK
mdr_report_form_share_url = f"{BASE_URL}/REPORT%20FORM/MDR%20REPORT%20FORM" 

# = FOR MAXVALUE =
##MAXVALUE LACASTA
lacasta_report_form_share_url = f"{BASE_URL}/REPORT%20FORM/MAXVALUE/LACASTA"

# == LINK ONEDRIVE OF HOTLINES AND CONTACT FORM ==
hotlines_and_confirm_form_url = f"{BASE_URL}/HOTLINE_AND_CONFIRM_FORM"

# ===== KHU VỰC ẢNH DAVITEQ =====
# == GATEWAY == 
gateway_bdnc_url = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/GATEWAY/BDNC"
gateway_tqb_url  = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/GATEWAY/TQB"
gateway_nvl_url  = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/GATEWAY/NVL"
gateway_vg_url = f"" #<< PENDING

# == LAYOUT ==
layout_bdnc_url = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/LAYOUT/BDNC"
layout_tqb_url  = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/LAYOUT/TQB"
layout_nvl_url  = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/LAYOUT/NVL"
layout_vg_url = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/LAYOUT/VG"

# == SENSOR ==
sensor_bdnc_url = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/SENSOR/BDNC"
sensor_tqb_url  = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/SENSOR/TQB"
sensor_nvl_url  = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/SENSOR/NVL"
sensor_vg_url = f"" #<< PENDING

# == ALARMPOINT ==
al_nvl_url = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/ALARM POINTS/NVL"
al_tqb_url = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/ALARM POINTS/TQB"
al_bdnc_url = f"" #<< NOT AVAILABLE
al_vg_url = f"{BASE_URL}/DAVITEQ/IMAGE_%20ARCHIVE/ALARM POINTS/VG"

# ===== LINK LƯU TRỮ CÁC TÀI LIỆU PDF =====
documentary_archive_url = f"{BASE_URL}/DOCUMENTARY"

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

# == Đường dẫn METADATA ==
METADATA_DIR = r"D:\RMC_Assistant_ver1.1\METADATA"

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

# Tạo thư mục METADATA
os.makedirs(METADATA_DIR, exist_ok=True)

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

# >> Sau một thời gian chương trình treo (idle) thì access_token hết hạn (thường là 1 giờ), nên không lấy được dữ liệu thường xuyên <<
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

# ==== Tải file từ OneDrive (sử dụng token truyền vào) ====
def download_file(token, file_id, filename):
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
        print(f"❌ Lỗi tải file {filename}: {r.status_code}")
        return None

# >> Quản lý và lưu trữ METADATA để thực hiện cập nhật và đồng bộ dữ liệu
# ==== Đường dẫn file metadata ====
METADATA_FILE = os.path.join(METADATA_DIR, "onedrive_metadata.json")

# ==== Đọc metadata đã lưu ====
def load_metadata():
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

# ==== Ghi metadata mới ====
def save_metadata(metadata):
    with open(METADATA_FILE, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=4, ensure_ascii=False)

# ==== Hàm đồng bộ file từ OneDrive ====
def sync_files_from_onedrive(token, share_url):
    """
    Đồng bộ một folder share_url:
     - Lấy danh sách file từ OneDrive (API Graph).
     - Với mỗi file: tìm các file cùng tên trên ổ đĩa (theo REPORT_FORM_DIR và các thư mục ảnh),
       hoặc dùng local_path có sẵn trong metadata.
     - Nếu local file cũ hơn remote -> xóa file local rồi tải lại.
     - Nếu không tồn tại file local -> tải về.
     - Cập nhật metadata (key = file_id).
    """
    # helper: tìm mọi file cùng tên (case-insensitive) trong các thư mục lưu trữ
    def find_local_paths_by_name(filename):
        filename_lower = filename.lower()
        roots = [
            REPORT_FORM_DIR, #<< Thư mục biểu mẫu
            IMAGE_LAYOUT_ARCHIVE_DIR, #<< Thư mục ảnh Layout
            IMAGE_GATEWAY_ARCHIVE_DIR, #<< Thư mục ảnh Gateway
            IMAGE_SENSOR_ARCHIVE_DIR, #<< Thư mục ảnh Sensor
            IMAGE_AL_ARCHIVE_DIR #<< Thư mục ảnh Alarm Point
        ]
        found = []
        for r in roots:
            if not r:
                continue
            for dirpath, dirs, files in os.walk(r):
                for f in files:
                    if f.lower() == filename_lower:
                        found.append(os.path.join(dirpath, f))
        return found

    # Lấy danh sách file từ OneDrive
    encoded_url = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    encoded_url = encoded_url.rstrip("=").replace("/", "_").replace("+", "-")
    url = f"https://graph.microsoft.com/v1.0/shares/u!{encoded_url}/driveItem/children"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        return

    items = r.json().get("value", [])
    files = [item for item in items if "file" in item]

    # Load metadata cũ
    local_metadata = load_metadata()

    for file in files:
        file_id = file.get("id")
        file_name = file.get("name")
        last_modified = file.get("lastModifiedDateTime")  # ISO 8601

        # Parse remote time -> timestamp (UTC)
        try:
            # OneDrive thường trả '2025-09-22T10:00:00Z' hoặc with offset
            remote_dt = datetime.datetime.fromisoformat(last_modified.replace("Z", "+00:00"))
            remote_ts = remote_dt.timestamp()
        except Exception as e:
            remote_ts = None

        # Tập hợp các candidate local paths:
        candidate_paths = []

        # 1) path lưu trong metadata (nếu có)
        if file_id in local_metadata and "local_path" in local_metadata[file_id]:
            candidate_paths.append(local_metadata[file_id]["local_path"])

        # 2) tìm bằng tên file trong các thư mục lưu trữ
        candidate_paths.extend(find_local_paths_by_name(file_name))

        # dedupe while preserving order
        seen = set()
        candidate_paths = [p for p in candidate_paths if not (p in seen or seen.add(p))]

        need_download = False
        # Nếu không tìm thấy candidate nào -> cần tải
        if not candidate_paths:
            need_download = True
        else:
            # Kiểm tra từng candidate; nếu có một file local hiện tại >= remote thì coi là ok
            local_is_fresh = False
            for p in candidate_paths:
                try:
                    if os.path.exists(p):
                        local_ts = os.path.getmtime(p)
                        # nếu không có remote_ts (fail parse) -> tải để an toàn
                        if remote_ts is None:
                            # không thể so sánh -> tải lại để đảm bảo nhất quán
                            try:
                                os.remove(p)
                            except Exception as e:
                                print(f"❌ Không thể xóa {p}: {e}")
                            need_download = True
                            break

                        # so sánh: nếu local cũ hơn remote -> xóa và đánh dấu cần tải
                        # dung sai 1 giây để tránh khác biệt nhỏ
                        if local_ts < (remote_ts - 1):
                            try:
                                os.remove(p)
                            except Exception as e:
                                print(f"❌ Lỗi xóa file {p}: {e}")
                                # nếu không xóa được, vẫn đánh dấu cần tải để ghi đè
                            need_download = True
                            # continue check other candidates (nếu có nhiều bản sao cũ)
                        else:
                            # local mới hơn hoặc bằng -> không cần tải lại
                            local_is_fresh = True
                            # update metadata entry (nếu thiếu)
                            local_metadata[file_id] = {
                                "name": file_name,
                                "lastModifiedDateTime": last_modified,
                                "local_path": p
                            }
                            break
                    else:
                        # candidate path có trong metadata nhưng file đã bị xóa -> cần tải
                        need_download = True
                except Exception as e:
                    print(f"❌ Lỗi khi xử lý file local {p}: {e}")
                    need_download = True

            if not local_is_fresh and not need_download:
                # nếu sau duyệt các candidate không tìm thấy file nào tươi (fresh), cần tải
                need_download = True

        # Nếu cần tải -> gọi download_file và cập nhật metadata
        if need_download:
            filepath = download_file(token, file_id, file_name)
            if filepath:
                # đảm bảo local path có tồn tại
                local_metadata[file_id] = {
                    "name": file_name,
                    "lastModifiedDateTime": last_modified,
                    "local_path": filepath
                }

    # Lưu metadata cuối cùng
    save_metadata(local_metadata)

# ==== Bắt đầu đồng bộ dữ liệu từ OneDrive ==== 
try:
    token = graph_session.ensure_token()

    # >> Cập nhật data cho các biểu mẫu báo cáo <<
    sync_files_from_onedrive(token, nvl_report_form_share_url)   # Gọi cho NVL
    sync_files_from_onedrive(token, tqb_report_form_share_url)   # Gọi cho TQB
    sync_files_from_onedrive(token, bdnc_report_form_share_url)  # Gọi cho BDNC
    sync_files_from_onedrive(token, vg_report_form_share_url)    # Gọi cho VG
    sync_files_from_onedrive(token, mdr_report_form_share_url)   # Gọi cho MDR
    sync_files_from_onedrive(token, lacasta_report_form_share_url) # Gọi cho LACASTA
    sync_files_from_onedrive(token, hotlines_and_confirm_form_url)  # Gọi cho HOTLINE AND CONFIRM FORM

    # >> Cập nhật data cho các hình ảnh DAVITEQ <<
    # GATEWAY
    sync_files_from_onedrive(token, gateway_bdnc_url)  
    sync_files_from_onedrive(token, gateway_tqb_url)   
    sync_files_from_onedrive(token, gateway_nvl_url)   

    #LAYOUT
    sync_files_from_onedrive(token, layout_bdnc_url)   
    sync_files_from_onedrive(token, layout_tqb_url)    
    sync_files_from_onedrive(token, layout_nvl_url)    

    #SENSOR
    sync_files_from_onedrive(token, sensor_bdnc_url)   
    sync_files_from_onedrive(token, sensor_tqb_url)    
    sync_files_from_onedrive(token, sensor_nvl_url)    

    #ALARMPOINT
    sync_files_from_onedrive(token, al_nvl_url)        
    sync_files_from_onedrive(token, al_tqb_url)        

except Exception as e:
    messagebox.showerror("❌ Lỗi đồng bộ OneDrive", f"Dữ liệu không thể đồng bộ:\n{e}")

# === Khu vực tạo Frame để lưu trữ các thành phần ===============================================================================================
# =============== Tạo cửa sổ chính =============== 
root.deiconify()
root.title("RMC Report Assistant")
root.geometry("1080x680")

# ==== Frame chính ==== <<<<<<<<<<<<<< MAIN FRAME 
main_frame = tk.Frame(root)
main_frame.pack(expand=True, pady=40, padx=20)

# ==== Frame đầu tiêu đề chứa phân loại khu vực ====
model_classification = tk.Frame(main_frame)
model_classification.pack()

# TẠO HÀNG ĐỢI PHÂN LOẠI ĐỂ CHỨA CÁC NÚT
SITE_GROUP_ORDER = {
    "AEONMALL": [],
    "MAXVALUE": []
}

# ==== Frame con chứa văn bản và các nút bên phải ====
show_text_and_multitasking_frame = tk.Frame(main_frame)
show_text_and_multitasking_frame.pack()

# ==== frame chứa nút contact, status và note bên trái ====
left_button_frame = tk.Frame(show_text_and_multitasking_frame)
left_button_frame.pack(side="left", fill="y", padx=10)

# ==== Text để hiển thị văn bản ====
output_text = tk.Text(show_text_and_multitasking_frame, font=("Arial", 13), width=60, height=15, wrap="word")
output_text.pack(side='left', pady=(10, 0), padx=10)
output_text.config(state='disabled')

# =============== Frame chứa các danh sách ATQB, ABDNC... với Scrollbar =============== 
site_container = tk.Frame(show_text_and_multitasking_frame)
site_container.pack(side='left', padx=10, fill="y")

# ==== Subframe cho thanh tìm kiếm (cố định, không cuộn) ====
site_search_frame = tk.Frame(site_container)
site_search_frame.pack(fill="x")

# ==== Subframe cho phần cuộn danh sách nút ====
site_list_container = tk.Frame(site_container)
site_list_container.pack(fill="both", expand=True)

site_canvas = tk.Canvas(site_list_container, width=150, height=100)
site_canvas.pack(side="left", fill="both", expand=True)

site_scrollbar = tk.Scrollbar(
    site_list_container,
    orient="vertical",
    command=site_canvas.yview
) 
site_scrollbar.pack(side="right", fill="y")

site_canvas.configure(yscrollcommand=site_scrollbar.set)
site_canvas.bind(
    '<Configure>',
    lambda e: site_canvas.configure(scrollregion=site_canvas.bbox("all"))
)

site_frame = tk.Frame(site_canvas)   # nơi chứa các nút cha
site_canvas.create_window((0, 0), window=site_frame, anchor="nw")

# ===============  Frame chứa các item xuất hiện khi chọn danh sách =============== 
item_container = tk.Frame(show_text_and_multitasking_frame)
item_container.pack(side='left', padx=10, fill="y")

# ==== Subframe cho thanh tìm kiếm (cố định, không cuộn) ====
item_search_frame = tk.Frame(item_container)
item_search_frame.pack(fill="x")

# ==== Subframe cho phần cuộn danh sách nút con ====
item_list_container = tk.Frame(item_container)
item_list_container.pack(fill="both", expand=True)

item_canvas = tk.Canvas(item_list_container, width=100, height=100)
item_canvas.pack(side="left", fill="both", expand=True)

item_scrollbar = tk.Scrollbar(item_list_container, orient="vertical", command=item_canvas.yview)
item_scrollbar.pack(side="right", fill="y")

item_canvas.configure(yscrollcommand=item_scrollbar.set)
item_canvas.bind('<Configure>', lambda e: item_canvas.configure(scrollregion=item_canvas.bbox("all")))

item_frame = tk.Frame(item_canvas)
item_window = item_canvas.create_window((0, 0), window=item_frame, anchor="nw")

def update_scrollregion(event=None):
    item_canvas.configure(scrollregion=item_canvas.bbox("all"))
    # Cho frame con luôn khớp chiều rộng với canvas
    item_canvas.itemconfig(item_window, width=item_canvas.winfo_width())

item_frame.bind("<Configure>", update_scrollregion)
item_canvas.bind("<Configure>", update_scrollregion)

# ==== FRAME CHO CÁC NÚT COPY, CLEAR, ĐỒNG HỒ, CATCH, CONTINUE ====
ccdcc_frame = tk.Frame(main_frame) # copy, clear, đồng hồ, catch, continue: ccdcc
ccdcc_frame.pack(fill='x', pady=(10, 0)) # giãn ngang (theo trục X)

# Nhóm chứa nút bên trái: Copy và Clear
left_controls = tk.Frame(ccdcc_frame)
left_controls.pack(side="left")

# Nhóm chứa nút bên phải: Catch, Clock, Continue
right_controls = tk.Frame(ccdcc_frame)
right_controls.pack(side="right")

# ==== Frame chứa các ô tô màu ====
box_frame = tk.Frame(main_frame)
box_frame.pack(pady=(10, 0))

# ==== thêm đồng hồ đếm ngược ==== 
timer_frame = tk.Frame(main_frame)
timer_frame.pack(pady=(10, 0))

# === KHU VỰC TẠO VÀ CẤU HÌNH CHỨC NĂNG ===========================================================================================
# ============== Chức năng tô màu ô tiến trình ==================
boxes = [] # Danh sách các ô vuông
box_colors = ["white"] * 6
hint_label = None # Label để hiển thị gợi ý
box_filled = [False] * 6
first_box_filled = False

# ==== Tạo 6 ô trắng trong box_frame ====
for i in range(6):
    lbl = tk.Label(
        box_frame, # Nằm trong box_frame
        width=22, # Rộng
        height=1, # Cao 
        bg="white", 
        relief="solid", # Kiểu viền 
        borderwidth=1) # Độ dày viền
    lbl.grid(row=0, column=i, padx=5)
    boxes.append(lbl)

# ==== TẠO hint_label nằm dưới box_frame ====
hint_label = tk.Label(
    main_frame,
    text="Quy trình xử lý sự cố đang đợi...",  # Nội dung hiển thị
    font=("Arial", 11),   # Font Arial cỡ 11
    fg="black",
    anchor="center",      # Căn giữa trong khung Label
    justify="center",     # Căn giữa nhiều dòng
    wraplength=900,       # Xuống dòng khi dài quá 900px (bạn chỉnh theo khung main_frame)
    height=3              # Chiều cao số dòng (có thể chỉnh)
)
hint_label.pack(fill="x", pady=(10, 20))

# ==== Hàm xử lý tô màu ====
def on_category_click():

    global box_colors

    # Đếm số ô đã được tô xanh
    green_count = box_colors.count("green")

    # Gợi ý tương ứng từng bước
    if green_count == 1:
        update_hint("Sau khi nhận sự cố, tiến hành báo cáo lên group chung, tiếp tục theo dõi sự cố đang diễn ra. Trong vòng 5 phút không có thông báo gì từ phía bên Site. Liên hệ với Site theo danh sách ưu tiên (Bấm xác nhận nếu như thông tin đã được cập nhật lên group từ bên Site). Sau khi liên hệ, cập nhật thông tin liên hệ lên Group chung thông qua biểu mẫu trong mục Contact.")
    elif green_count == 2:
        update_hint("Tiếp tục theo dõi và cập nhật. Sau một khoảng thời gian không nhận được thông tin gì từ Site kể từ thời điểm đã liên hệ (1 - 2 tiếng). Tiến hành liên hệ lại Site để xác minh tình trạng kiếm tra thiết bị và nguyên nhân (nếu có). Tiến hành cập nhập lại tình hình lên nhóm group chung (Bấm 'Xác nhận' để bỏ qua bước này nếu như sự cố thiết bị đã được khắc phục và nguyên nhân sự cố đã được cập nhật).")
    elif green_count == 3:
        update_hint("Nếu sự cố sau 1 tiếng cho đến 2 tiếng vẫn chưa được sự xử lí và cũng chưa được cập nhật lên group chung. Tiến hành liên hệ lại với số điện thoại ưu tiên để xác nhận lại sự cố, sau đó báo cáo lại tình lên group chung (Bấm 'Xác nhận' nếu sự cố đã được giải quyết trước thời điểm này).")
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
    site_canvas.update_idletasks()
    site_canvas.configure(scrollregion=site_canvas.bbox("all"))

search_parent_entry = tk.Entry(site_search_frame, textvariable=search_parent_var)
search_parent_entry.pack(fill="x", pady=5)
search_parent_entry.bind("<KeyRelease>", filter_parent_buttons)

parent_items = []  # list chứa các nút cha

# ==== THANH TÌM KIẾM CHO DANH SÁCH CON (item_frame) ====
search_child_var = tk.StringVar()
def filter_child_buttons(event=None):
    keyword = search_child_var.get().lower()

    # tìm state của parent đang active
    current_state = None
    for key, cfg in LIST_CONFIG.items():
        st = cfg["state"]()
        if st["visible"]:  # đang mở
            current_state = st
            break

    if not current_state:
        return  # không có list nào mở

    matched = []
    unmatched = []
    for btn in current_state["buttons"]:
        if keyword in btn.cget("text").lower():
            matched.append(btn)
        else:
            unmatched.append(btn)

    for btn in current_state["buttons"]:
        btn.pack_forget()

    for btn in matched + unmatched:
        btn.pack(anchor='w', pady=1)

    # 🔥 cập nhật scrollregion
    item_canvas.update_idletasks()
    item_canvas.configure(scrollregion=item_canvas.bbox("all"))

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

# ==== Chức năng bắt và tiếp tục đồng hồ thời gian thực của hệ thống ====
is_running = True

# >> Chức năng bắt thời gian (ngừng đồng hồ thười gian thực) <<
def catch_clock():
    global is_running
    is_running = False

# >> Chức năng tiếp tục đồng hồi thười gian thực <<
def continue_clock():
    global is_running
    is_running = True

# >> Chức năng cập nhật, đọc trạng thái sự kiện hai nút catch và continue để thực hiện cập nhật đồng hồ <<
def update_clock():
    if is_running:
        now = datetime.datetime.now().strftime("%H:%M:%S")
        clock_label.config(text=now)
    root.after(1000, update_clock)

# ==== Chức năng lấy và hiển thị thời gian hiện tại trên hệ thống lên văn bản ====
def update_timer():
    global time_left, countdown_job
    minutes, seconds = divmod(time_left, 60)
    timer_label.config(text=f"⏳ Remain: {minutes:02d}:{seconds:02d}")
    if time_left > 0:
        time_left -= 1
        countdown_job = root.after(1000, update_timer)
    else:
        timer_label.config(text="⏰Contact Site⏰")

# ==== Chức năng bắt đầu và reset đồng hồ đếm ngược, đặt thời gian đếm ngược cho đồng hồ ====
# >> Bắt đầu đếm ngược <<
def start_timer():
    global time_left, countdown_job
    if countdown_job:
        root.after_cancel(countdown_job)
    time_left = 300
    update_timer()

# >> Đặt lại đồng hồ <<
def reset_timer():
    global time_left, countdown_job
    if countdown_job:
        root.after_cancel(countdown_job)
        countdown_job = None
    time_left = 300
    timer_label.config(text="⏳Waiting Countdown⏳")

# ==== Tạo labels cho đồng hồ thời gian thực và đồng hồ đếm ngược ====
timer_label = tk.Label(timer_frame, text="⏳Waiting Countdown⏳", font=("Arial", 16, "bold"), fg="blue")
timer_label.pack()
countdown_job = None
time_left = 300  # 5 phút = 300 giây

current_visible_group = "AEONMALL"

# ==== Chức năng reset toàn bộ danh sách nút con và trạng thái chọn ====
def reset_all_lists():
    global active_parent_button, active_child_button

    # đóng toàn bộ list con
    for cfg in LIST_CONFIG.values():
        st = cfg["state"]()
        if st["visible"]:
            toggle_sub_buttons(st, cfg["files"])

        # reset màu nút CHA
        if "button" in st and st["button"]:
            st["button"].config(bg="SystemButtonFace", fg="black")

    # xóa nút con còn sót trong item_frame
    for widget in item_frame.winfo_children():
        widget.destroy()

    # reset trạng thái chọn
    active_parent_button = None
    active_child_button = None

# ==== Chức năng hiển thị nhóm site (AEONMALL hoặc MAXVALUE) ====
def show_site_group(group_name):
    global current_visible_group
    current_visible_group = group_name

    # reset toàn bộ trạng thái + màu
    reset_all_lists()

    # Ẩn toàn bộ block_frame
    for group_frames in SITE_GROUP_ORDER.values():
        for frame in group_frames:
            frame.pack_forget()

    # Pack lại đúng thứ tự
    for frame in SITE_GROUP_ORDER[group_name]:
        frame.pack(fill="x", pady=2)

    # đổi màu nút phân loại
    if group_name == "AEONMALL":
        aeon_mall_button.config(bg="#ef3eb3")
        maxvalue_button.config(bg="#a0a0a0")
    else:
        aeon_mall_button.config(bg="#a0a0a0")
        maxvalue_button.config(bg="#c4005b")

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

# ===================== CHỨC NĂNG HIỂN THỊ VĂN BẢN VÀ THỜI GIAN =====================
# ==== Hiển thị văn bản từ OneDrive vào Text widget ====
def show_text_from_drive(file_id, filename, is_no_error=False, start_timer_flag=True):
    try:
        # ✅ Lấy token hợp lệ trước khi tải
        token = graph_session.ensure_token()

        # ✅ Gọi đúng hàm download_file với 3 tham số
        file_path = download_file(token, file_id, filename)

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

    # ✅ Cập nhật vào Text widget
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, content)
    output_text.config(state='disabled')

    if start_timer_flag:
        start_timer()

# == AEON MALL ==
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
    "ABNC_FR&FC", "ABNC_POWER", "ABNC_FAN", 
    "ABNC_LPG", 
    "ABNC_NO_ERROR"
]

device_name_avg = [
    "AVG_BAKERY", "AVG_CAFE", "AVG_DELICA", 
    "AVG_FAN", "AVG_FISH", "AVG_FR&FC", 
    "AVG_MEAT", "AVG_NOODLE", "AVG_POWER1F",
    "AVG_POWER2F", "AVG_PRODUCT", "AVG_SUSHI", "AVG_BMS",
    "AVG_SECURITY",
    "AVG_NO_ERROR"
] 

device_name_amdr = [
    "MDR_BAKERY", "MDR_BMSGENERAL", "MDR_CAFE", "MDR_FIREPUMP", "MDR_FISH",
    "MDR_FR&FC", "MDR_GENERATOR", "MDR_HIGHLEVELWATERTANK", "MDR_KOFKEF",
    "MDR_LOWLEVELWATERTANK", "MDR_NO_ERROR", "MDR_NOODLE", "MDR_PRODUCT",
    "MDR_PWRSUPPLY1", "MDR_PWRSUPPLY2", "MDR_SUPPLY3", "MDR_SERVER",
    "MDR_SOCKET", "MDR_WATERSUPPLY", "MDR_WATERTANKBUMP",
    "MDR_SECURITY"
    ]

#>>> ADD SITE LISTS HERE <<<   

# == MAXVALU ==
device_names_lacasta = [
    "LACASTA_CHEST", "LACASTA_ICECREAM", "LACASTA_UPRIGHT", "LACASTA_SHOWCASE",
    "LACASTA_POWERSUPPLY",
    "LACASTA_NO_ERROR"
]

#>>> ADD SITE LISTS HERE <<<   

contact_sample = ["CONTACT_FORM"]
confirm_sample = ["CONFIRM_FORM"]
notification_sample = ["NOTIFICATION_FORM"]

# ============= MAPPING RIÊNG CHO TỪNG KHU VỰC ================
# == AEON MALL ==
nvl_report_form_files = build_device_mapping(nvl_report_form_share_url, device_names_anvl)
tqb_report_form_files = build_device_mapping(tqb_report_form_share_url, device_names_atqb)
bdnc_report_form_files = build_device_mapping(bdnc_report_form_share_url, device_names_abnc)
vg_report_form_share_url = build_device_mapping(vg_report_form_share_url, device_name_avg)
mdr_report_form_share_url = build_device_mapping(mdr_report_form_share_url, device_name_amdr)   

#>>> ADD SITE LISTS HERE <<<   

# == MAXVALUE ==
lacasta_report_form_share_url = build_device_mapping(lacasta_report_form_share_url, device_names_lacasta)

#>>> ADD SITE LISTS HERE <<<   

# ==== TẠO GIAO DIỆN DANH SÁCH ====
# >> Tạo biến trạng thái cho nút cha <<
active_parent_button = None 

# >> Tạo biến trạng thái cho nút con <<
active_child_button = None

# >> Đặt trạng thái và đổi màu nút khi được chọn CHO NÚT CHA <<
def set_active_parent_button(btn):
    global active_parent_button
    # Reset nút cha cũ
    if active_parent_button and active_parent_button != btn:
        active_parent_button.config(bg="SystemButtonFace", fg="black")
    # Đổi màu nút cha mới
    btn.config(bg="green", fg="white")
    active_parent_button = btn

# >> Đặt trạng thái và đổi màu nút khi được chọn CHO NÚT CON <<
def set_active_child_button(btn):
    global active_child_button
    # Reset nút con cũ
    if active_child_button and active_child_button != btn:
        active_child_button.config(bg="SystemButtonFace", fg="black")
    # Đổi màu nút con mới
    btn.config(bg="blue", fg="white")
    active_child_button = btn

# ==== TẠO KHỐI DANH SÁCH CHA - CON ====
def create_list_block(parent, list_name, items, toggle_function, state):
    block_frame = tk.Frame(parent)
    block_frame.pack(pady=10, anchor='w')

    list_button = tk.Button(
        block_frame,
        text=list_name,
        font=("Arial", 14),
        width=12,
        command=lambda: [
            set_active_parent_button(list_button),
            toggle_function()   
        ]
    )
    list_button.pack(anchor='w')

    state["button"] = list_button
    return block_frame

# ==== HÀM BẬT TẮT DANH SÁCH ====
# == Định nghĩa mapping cho từng list ==
LIST_CONFIG = {
    "list1-NVL": {"state": lambda: list1_state, "files": nvl_report_form_files},
    "list2-TQB": {"state": lambda: list2_state, "files": tqb_report_form_files},
    "list3-BDNC": {"state": lambda: list3_state, "files": bdnc_report_form_files},
    "list4-VG": {"state": lambda: list4_state, "files": vg_report_form_share_url},
    "list5-MDR": {"state": lambda: list5_state, "files": mdr_report_form_share_url},
    "lacasta": {"state": lambda: lacasta_state, "files": lacasta_report_form_share_url},
}

LIST_GROUP_MAP = {
    "list1-NVL": "AEONMALL",
    "list2-TQB": "AEONMALL",
    "list3-BDNC": "AEONMALL",
    "list4-VG": "AEONMALL",
    "list5-MDR": "AEONMALL",
    "lacasta": "MAXVALUE"
}

def toggle_list(target_key):
    target_state = LIST_CONFIG[target_key]["state"]()
    target_files = LIST_CONFIG[target_key]["files"]

    # 🚨 chỉ cho mở đúng group đang hiển thị
    if target_state["group"] != current_visible_group:
        return

    # đóng list khác
    for key, cfg in LIST_CONFIG.items():
        if key != target_key:
            st = cfg["state"]()
            if st["visible"]:
                toggle_sub_buttons(st, cfg["files"])

    toggle_sub_buttons(target_state, target_files, auto_select_first=True)

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

            state["buttons"].append(btn)

            if idx == 0:
                first_child_btn = btn

        state["visible"] = True

        # update lại canvas scrollregion
        item_canvas.update_idletasks()
        item_canvas.configure(scrollregion=item_canvas.bbox("all"))

        if auto_select_first and first_child_btn:
            set_active_parent_button(state["button"])
            set_active_child_button(first_child_btn)
    else:
        for btn in state["buttons"]:
            btn.pack_forget()
        state["buttons"].clear()
        state["visible"] = False

        # reset scrollregion khi đóng
        item_canvas.update_idletasks()
        item_canvas.configure(scrollregion=item_canvas.bbox("all"))

# ==== TRẠNG THÁI ====
# == AEON MALL ==
list1_state = {"visible": False, "buttons": [], "group": "AEONMALL"}
list2_state = {"visible": False, "buttons": [], "group": "AEONMALL"}
list3_state = {"visible": False, "buttons": [], "group": "AEONMALL"}
list4_state = {"visible": False, "buttons": [], "group": "AEONMALL"}
list5_state = {"visible": False, "buttons": [], "group": "AEONMALL"}

#>>> ADD SITE LISTS HERE <<<   
lacasta_state = {"visible": False, "buttons": [], "group": "MAXVALUE"}
#MAXVALU

# ==== TẠO DANH SÁCH GIAO DIỆN ====
# == AEON MALL ==
btn = create_list_block(
    site_frame, "ANVL", nvl_report_form_files,
    lambda: toggle_list("list1-NVL"),
    list1_state
)
SITE_GROUP_ORDER["AEONMALL"].append(btn)

btn = create_list_block(
    site_frame, "ATQB", tqb_report_form_files,
    lambda: toggle_list("list2-TQB"),
    list2_state
)
SITE_GROUP_ORDER["AEONMALL"].append(btn)

btn = create_list_block(
    site_frame, "ABNC", bdnc_report_form_files,
    lambda: toggle_list("list3-BDNC"),
    list3_state
)
SITE_GROUP_ORDER["AEONMALL"].append(btn)

btn = create_list_block(
    site_frame, "AVG", vg_report_form_share_url,
    lambda: toggle_list("list4-VG"),
    list4_state
)
SITE_GROUP_ORDER["AEONMALL"].append(btn)

btn = create_list_block(
    site_frame, "AMDR", mdr_report_form_share_url,
    lambda: toggle_list("list5-MDR"),
    list5_state
)
SITE_GROUP_ORDER["AEONMALL"].append(btn)

#>>> ADD SITE LISTS HERE <<<   

# == MAXVALU ==
btn = create_list_block(
    site_frame, "LACASTA", lacasta_report_form_share_url,
    lambda: toggle_list("lacasta"),
    lacasta_state
)
SITE_GROUP_ORDER["MAXVALUE"].append(btn)

#>>> ADD SITE LISTS HERE <<<   

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
            file_path = download_file(token, file_id, filename)

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
        # === Tải file về ===
        file_path = download_file(token, file_id, filename)

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

    #Cập nhật stt của note trên giao diện   
    def update_stt_label():
        current_stt.set(str(len([f for f in os.listdir(DATA_DIR) if f.endswith(".json")])))

    # Lưu dữ liệu nhắc vào file mới
    def save_reminder_to_new_file(reminder_data):
        stt = get_next_stt()
        file_path = os.path.join(DATA_DIR, f"reminders{stt}.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(reminder_data, f, ensure_ascii=False, indent=4)
        update_stt_label()

    #Chức năng lập lịch nhắc nhở (reminder) theo thời gian, ngày và tháng định sẵn, kèm theo xử lý file khi thông báo được hiển thị
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

    #Chức năng tạo nhắc (reminder)
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

    #Tạo hiệu ứng “placeholder” (gợi ý mờ bên trong ô nhập liệu)
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

    #Đọc toàn bộ các file .json trong thư mục DATA_DIR, sau đó tải nội dung từng file lên chương trình để sử dụng
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

    #Hiển thị danh sách các dữ liệu (data_list) lên bảng Treeview, phân loại (gắn màu hoặc tag) các dòng theo trạng thái của lịch nhắc (còn hạn, hết hạn, hay lặp lại)
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

    #Tìm kiếm và hiển thị kết quả lọc
    def search_data():
        keyword = search_var.get().lower()
        filtered = [item for item in full_data if keyword in item["keyword"].lower() or keyword in item["content"].lower()]
        display_data(filtered)

    #Làm mới toàn bộ dữ liệu hiển thị
    def refresh_data():
        global full_data
        full_data = load_all_json_files()
        display_data(full_data)
        update_stt_label()  # cập nhật STT ghi chú tiếp theo

    #Xóa các ghi chú (note hoặc reminder) mà người dùng chọn trong bảng (Treeview), xóa cả file JSON tương ứng trên ổ đĩa.
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

    #Giao diện tạo note
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

    #Giao diện xem note 
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

        # ====== Xử lý nhấp đúp để xem chi tiết nội dung ======
        def on_double_click(event):
            selected_item = tree.identify_row(event.y)
            selected_col = tree.identify_column(event.x)

            if not selected_item:
                return

            # Xác định chỉ số cột (ví dụ: '#3' là cột Nội dung)
            if selected_col == '#3':  # cột "Nội dung"
                values = tree.item(selected_item, "values")
                full_content = values[2]  # nội dung thực tế

                # Tạo cửa sổ popup nhỏ hiển thị nội dung
                popup = tk.Toplevel(note_window)
                popup.title("Chi tiết nội dung ghi chú")
                popup.geometry("400x250")
                popup.transient(note_window)  # nằm trên cửa sổ chính
                popup.grab_set()  # khóa focus vào popup

                # Frame chính
                frame = tk.Frame(popup, padx=10, pady=10)
                frame.pack(fill="both", expand=True)

                # Ô text cuộn để hiển thị nội dung
                text_frame = tk.Frame(frame)
                text_frame.pack(fill="both", expand=True, pady=5)

                text_scroll = tk.Scrollbar(text_frame)
                text_scroll.pack(side="right", fill="y")

                text_box = tk.Text(text_frame, wrap="word", yscrollcommand=text_scroll.set, font=("Arial", 10))
                text_box.insert("1.0", full_content)
                text_box.config(state="disabled")  # chỉ xem, không sửa
                text_box.pack(fill="both", expand=True)
                text_scroll.config(command=text_box.yview)

        # Gán sự kiện nhấp đúp chuột
        tree.bind("<Double-1>", on_double_click)
        full_data = load_all_json_files()
        display_data(full_data)

    #Giao diện tạo biểu mẫu thông báo 
    def show_notification_form(content=None):
        new_window = tk.Toplevel(root)
        new_window.geometry("650x400")
        new_window.title("Nhập thông tin Notification")
        new_window.configure(bg="#f7f9fc")

        # ====== Frame nhập liệu ======
        form_frame = tk.Frame(new_window, bg="#f7f9fc")
        form_frame.pack(padx=20, pady=10, fill="x")

        # ==== Các trường nhập liệu ====
        labels = ["Site:", "Description:", "Start Time:", "Start Date:",
                  "End Time:", "End Date:", "Devices:", "Note:"]
        entries = {}

        for i, label in enumerate(labels):
            tk.Label(form_frame, text=label, font=("Arial", 11), bg="#f7f9fc").grid(row=i, column=0, sticky="w", pady=5)
            if "Description" in label or "Note" in label:
                entry = tk.Text(form_frame, font=("Arial", 11), height=3, width=40)
                entry.grid(row=i, column=1, pady=5, sticky="ew")
            elif "Date" in label:
                entry = DateEntry(form_frame, font=("Arial", 11), width=18,
                                  background='darkblue', foreground='white', borderwidth=2)
                entry.set_date(datetime.date.today())
                entry.grid(row=i, column=1, pady=5, sticky="w")
            else:
                entry = tk.Entry(form_frame, font=("Arial", 11))
                entry.grid(row=i, column=1, pady=5, sticky="ew")
            entries[label.strip(":").lower()] = entry

        form_frame.columnconfigure(1, weight=1)

        # ====== Xử lý khi bấm OK ======
        def handle_ok():
            try:
                # === Lấy dữ liệu từ form ===
                site = entries["site"].get().strip()
                description = entries["description"].get("1.0", tk.END).strip()
                start_time = entries["start time"].get().strip()
                start_date = entries["start date"].get_date()
                end_time = entries["end time"].get().strip()
                end_date = entries["end date"].get_date()
                devices = entries["devices"].get().strip()
                note = entries["note"].get("1.0", tk.END).strip()

                # === Lấy danh sách file trong thư mục OneDrive ===
                files = list_files_from_url(hotlines_and_confirm_form_url)

                # === Tìm file notification ===
                target_name = next(iter(notification_sample))  # ví dụ "NOTIFICATION_FORM"
                target_file = next((f for f in files if target_name in f.get("name", "")), None)

                if not target_file:
                    raise FileNotFoundError(f"Không tìm thấy file chứa '{target_name}' trong thư mục OneDrive.")

                file_id = target_file["id"]
                # đảm bảo filename hợp lệ (loại bỏ path nếu có)
                filename = target_file.get("name", "")
                if not filename:
                    raise ValueError("Tên file nhận từ OneDrive rỗng.")

                filename = os.path.basename(filename)

                # đảm bảo thư mục NOTE tồn tại và có quyền ghi
                os.makedirs(NOTE_ARCHIVE_DIR, exist_ok=True)

                # === Tải file về ===
                file_path = download_file(token, file_id, filename)

                # Kiểm tra kết quả: phải tồn tại và phải là file (không phải thư mục)
                if not file_path or not os.path.exists(file_path) or not os.path.isfile(file_path):
                    raise FileNotFoundError(f"File notification không tồn tại sau khi tải: {file_path}")

                # === Đọc file notification template ===
                with open(file_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()

                replaced_lines = []
                for line in lines:
                    original_line = line
                    line = line.replace("[site]", site)
                    line = line.replace("[description]", description)
                    line = line.replace("[start_time]", start_time)
                    line = line.replace("[start_date]", str(start_date))
                    line = line.replace("[end_time]", end_time)
                    line = line.replace("[end_date]", str(end_date))
                    line = line.replace("[devices]", devices)
                    line = line.replace("[note]", note)

                    stripped_line = line.strip()
                    if not stripped_line:
                        continue
                    replaced_lines.append(stripped_line)

                content = '\n'.join(replaced_lines)

                # === Hiển thị cửa sổ xem nội dung notification ===
                result_window = tk.Toplevel(root)
                result_window.title("Notification Content")
                result_window.geometry("600x500")
                result_window.configure(bg="#f7f9fc")

                text_box = tk.Text(result_window, wrap="word", font=("Arial", 11), bg="white", height=20)
                text_box.pack(padx=15, pady=15, fill="both", expand=True)
                text_box.insert(tk.END, content)
                text_box.config(state="disabled")

                # === Nút Copy ===
                def copy_to_clipboard():
                    root.clipboard_clear()
                    root.clipboard_append(content)
                    messagebox.showinfo("Copied", "Đã sao chép nội dung notification vào clipboard!")

                copy_btn = tk.Button(result_window, text="Copy", font=("Arial", 12, "bold"),
                                     bg="#007BFF", fg="white", command=copy_to_clipboard)
                copy_btn.pack(pady=10)

                new_window.destroy()

            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi xử lý notification: {e}")

        # === Nút OK ===
        ok_button = tk.Button(new_window, text="OK", font=("Arial", 12, "bold"),
                              bg="green", fg="white", command=handle_ok)
        ok_button.pack(pady=10)

    tk.Button(btn_frame, text="Tạo Note", width=20, command=show_create_note).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Xem Note", width=20, command=show_view_notes).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Biểu mẫu thông báo", width=20, command=show_notification_form).pack(side="left", padx=10)

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

            filename = file["name"]
            local_path = os.path.join(save_dir, filename)

            # Nếu chưa có thì tải về
            if not os.path.exists(local_path):
                img_path = download_file(token, file["id"], filename)  # ✅ đúng định nghĩa
                if img_path and img_path != local_path:
                    # Nếu download_file mặc định lưu về REPORT_FORM_DIR → copy sang save_dir
                    shutil.copy(img_path, local_path)
                    img_path = local_path
            else:
                img_path = local_path

            if img_path and not str(img_path).startswith("ERROR"):
                try:
                    img = Image.open(img_path)
                    img.thumbnail((100, 75), Image.Resampling.LANCZOS)  # Thumbnail
                    photo = ImageTk.PhotoImage(img)

                    row = (idx * 2) // max_columns
                    col = idx % max_columns

                    label_img = tk.Label(image_frame, image=photo, bg="white", cursor="hand2")
                    label_img.image = photo
                    label_img.grid(row=row, column=col, padx=5, pady=(5, 0))

                    label_text = tk.Label(image_frame, text=filename, bg="white", font=("Arial", 9))
                    label_text.grid(row=row + 1, column=col, padx=5, pady=(0, 10))

                    # 👉 Khi bấm vào thumbnail thì mở bằng ứng dụng mặc định
                    label_img.bind("<Button-1>", lambda e, path=img_path: open_with_default_app(path))

                except Exception as e:
                    image_label.config(text=f"Lỗi xử lý ảnh: {e}", image='', bg="white")
                    break

    def open_with_default_app(img_path):
        try:
            if sys.platform.startswith("darwin"):  # macOS
                subprocess.call(("open", img_path))
            elif os.name == "nt":  # Windows
                os.startfile(img_path)
            elif os.name == "posix":  # Linux
                subprocess.call(("xdg-open", img_path))
        except Exception as e:
            print(f"Lỗi mở ảnh: {e}")

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
            "VG": list_files_from_url(layout_vg_url),
        },
        "SENSOR": {
            "BDNC": list_files_from_url(sensor_bdnc_url),
            "TQB": list_files_from_url(sensor_tqb_url),
            "NVL": list_files_from_url(sensor_nvl_url),
        },
        "ALARMPOINT": {
            "TQB": list_files_from_url(al_tqb_url),
            "NVL": list_files_from_url(al_nvl_url),
            "VG": list_files_from_url(al_vg_url),
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

# === Khu vực tạo các nút thành phần =======================================================================================================================================
#Nút phân loại trên cùng (maxvalue và aeon mall))
aeon_mall_button = tk.Button(
    model_classification,
    text="AEONMALL",
    font=("Arial", 10, "bold"),
    bg="#ef3eb3",
    fg="white",
    width=15,
    command=lambda: show_site_group("AEONMALL")
)
aeon_mall_button.pack(side="left", padx=(0, 5))

maxvalue_button = tk.Button(
    model_classification,
    text="MAXVALUE",
    font=("Arial", 10, "bold"),
    bg="#a0a0a0",
    fg="white",
    width=15,
    command=lambda: show_site_group("MAXVALUE")
)
maxvalue_button.pack(side="left")

# Nút copy (bên trái)
copy_button = tk.Button(
    left_controls, 
    text="Copy", 
    font=("Arial", 10, "bold"), 
    bg="#4CAF50", 
    fg="white",
    command=copy_text_to_clipboard, 
    width=15)
copy_button.pack(side="left", padx=(0, 5))

# Nút clear (bên trái)
clear_button = tk.Button(
    left_controls, text="Clear", 
    font=("Arial", 10, "bold"), 
    bg="#f44336", fg="white",
    command=clear_text_output, 
    width=15)
clear_button.pack(side="left")

# Catch (ngoài cùng bên phải)
catch_button = tk.Button(
    right_controls, 
    text="Catch", 
    font=("Arial", 10, "bold"), 
    bg="#029B82", fg="white",
    command=catch_clock, 
    width=10)
catch_button.pack(side='right', padx=5)

# Clock (giữa)
clock_label = tk.Label(
    right_controls, 
    font=("Roboto", 20, "bold"), 
    fg="#D20103",)
clock_label.pack(side='right', padx=10)

# Continue (ngoài cùng bên trái của cụm)
continue_button = tk.Button(
    right_controls, 
    text="Continue", 
    font=("Arial", 10, "bold"), 
    bg="#2196F3", 
    fg="white",
    command=continue_clock, 
    width=10)
continue_button.pack(side='right', padx=5)

# ==== NÚT CONTACT ====
def contact_action():
    if fill_box(1):  # Chỉ tô nếu ô 0 đã tô
        create_new_window_contact("Contact")
        on_category_click()
        reset_timer()

contact_button = tk.Button(
    left_button_frame, 
    text="Contact", 
    font=("Arial", 12, "bold"),
    bg="#2196F3", 
    fg="white", 
    width=10, 
    command=lambda: contact_action())
contact_button.pack(pady=5)

# ==== NÚT STATUS ====
def status_action():
    if fill_box(2):  # Chỉ tô nếu ô 1 đã tô
        create_new_window_status("Status")
        on_category_click()
        reset_timer()
status_button = tk.Button(
    left_button_frame, 
    text="Status", 
    font=("Arial", 12, "bold"),
    bg="#FF9800", 
    fg="white", 
    width=10, 
    command=lambda: status_action())
status_button.pack(pady=5)

# ==== NÚT NOTE ====
def note_action():
    create_new_window_note()
note_button = tk.Button(
    left_button_frame, 
    text="Note", 
    font=("Arial", 12, "bold"),
    bg="#873e23", 
    fg="white", 
    width=10, 
    command=lambda: note_action())
note_button.pack(pady=5)

# ==== NÚT KHO ẢNH DAVITEQ ====
def image_daviteq_action():
    create_new_window_image_daviteq("DAVITEQ")
image_daviteq_button = tk.Button(
    left_button_frame, 
    text="DAVITEQ", 
    font=("Arial", 12, "bold"),
    bg="#3fc4f3", 
    fg="white", 
    width=10, 
    command=lambda: image_daviteq_action())
image_daviteq_button.pack(pady=5)

# ==== NÚT VÀO KHO DOCUMENTARY ====
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
confirm_button = tk.Button(
    main_frame, 
    text="Xác nhận", 
    font=("Arial", 12, "bold"),
    bg="#4CAF50", fg="white", 
    command=confirm_action)
confirm_button.pack(pady=10)

# Mặc định hiển thị nhóm AEONMALL
show_site_group("AEONMALL")

# Khởi tạo auto select cho list1 (tùy chọn)
toggle_list("list1-NVL")

# Bắt đầu cập nhật đồng hồ
update_clock()

# ==== CHẠY ỨNG DỤNG ====
root.mainloop()
