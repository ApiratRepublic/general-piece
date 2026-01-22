#############
##### Transfer folders by PROV_ID
##### Workflow:
##### 1) ตรวจเช็คว่ามีโฟลเดอร์ไหนบ้างที่มี PROV_ID ตรงกับไฟล์แมป
##### 2) สร้างซิปของโฟลเดอร์เหล่านั้น
##### 3) Ping server
##### 4) ส่งไฟล์ซิปไปยังปลายทาง (จะเขียนทับถ้ามีอยู่แล้ว)
##### 5) ลบไฟล์ซิปชั่วคราวที่สร้างขึ้น
##### 6) ทำรายงานผลลัพธ์เป็น Excel
##### APIRAT RATTANAPAIBOON
##### Version 1.0 : 2025-11-12
#############

import os
import pandas as pd
import shutil
import subprocess
import logging
import sys
import errno

# ================= CONFIGURATION =================
SOURCE_DIR = r'D:\A02-Projects\WarRoom\Phase02\Transfer' #โฟลเดอร์ต้นทาง
TEMP_ZIP_DIR = r'D:\A02-Projects\WarRoom\Phase02\_temp_zip' #โฟลเดอร์เก็บไฟล์ซิปชั่วคราว

IP_FILE = r'D:\A03-Resources\API\IPADDRESS.xlsx' #ไฟล์แมป PROV_ID -> IP
IP_SHEET_NAME = "PROV_IP" #ชื่อชีตในไฟล์แมป

DEST_SUBFOLDER_NAME = r'WarRoom_2569' #ชื่อโฟลเดอร์ปลายทางบนเซิร์ฟเวอร์
API_FOLDER = "2026-01-06" #ชื่อโฟลเดอร์ย่อยปลายทางบนเซิร์ฟเวอร์

REPORT_PATH = os.path.join(
    SOURCE_DIR,
    f"{API_FOLDER}-Transfer2.xlsx" #ชื่อไฟล์รายงานผลลัพธ์
)

# ================= กำหนดคอลัมน์รายงาน =================
RESULT_COLUMNS = [
    "Original Folder Name",
    "Destination Path",
    "Status",
    "Error Category",
    "Suggested Fix",
    "PROV_ID"
]

# ================= LOGGING =================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ================= ฟังก์ชัน =================
## ดึงชื่อเซิร์ฟเวอร์จาก UNC path
def extract_server_from_unc(network_path):
    if isinstance(network_path, str) and network_path.startswith('\\\\'):
        return network_path.split('\\')[2]
    return None
## ================= ดึงข้อมูล ip จับกับ PROV_ID =================
def load_server_mapping(ip_file_path, sheet_name):
    print("[LOAD] Loading PROV_ID mapping")
    try:
        df = pd.read_excel(ip_file_path, sheet_name=sheet_name, dtype=str)
        df['PROV_ID'] = df['PROV_ID'].str.strip()
        df['IP'] = df['IP'].str.strip().str.rstrip(r'\/')
        mapping = df.set_index('PROV_ID')['IP'].to_dict()
        print(f"[LOAD] Loaded {len(mapping)} mappings")
        return mapping
    except Exception as e:
        print("[FATAL] Cannot load server mapping")
        logging.error(e, exc_info=True)
        sys.exit(1)
# ================= จัดกลุ่ม ERROR =================
def classify_transfer_error(e):
    if hasattr(e, 'winerror'):
        return {
            5: ("Access Denied", "Check permissions"),
            53: ("Network Path Not Found", "Check VPN / server"),
            64: ("Network Name Deleted", "Connection unstable"),
            112: ("Disk Full", "Free destination space"),
            121: ("Timeout", "Network unstable")
        }.get(e.winerror, ("Windows Error", str(e)))

    if hasattr(e, 'errno'):
        return {
            errno.EACCES: ("Access Denied", "Check permissions"),
            errno.ENOSPC: ("Disk Full", "Free destination space"),
            errno.ENOENT: ("Path Not Found", "Check destination path"),
        }.get(e.errno, ("OS Error", str(e)))

    return ("Unknown Error", str(e))

# ================= 1) ตรวจโฟลเดอร์ต้นทาง =================
def discover_folders(source_dir, server_map):
    print("[DISCOVER] Scanning source directory")
    folders = {}

    for name in os.listdir(source_dir):
        full_path = os.path.join(source_dir, name)
        if not os.path.isdir(full_path):
            continue

        prov_id = name.split('-', 1)[0]
        if prov_id in server_map:
            folders[prov_id] = {
                "source_path": full_path,
                "folder_name": name
            }
            print(f"[DISCOVER] OK: {name} → {prov_id}")
        else:
            print(f"[DISCOVER] SKIP: {name}")

    print(f"[DISCOVER] Total valid folders: {len(folders)}")
    return folders

# ================= 2) ZIP โฟลเดอร์ =================
def zip_all_folders(folders, temp_zip_dir):
    print("[ZIP] Zipping all folders")
    os.makedirs(temp_zip_dir, exist_ok=True)

    zipped = {}
    for prov_id, info in folders.items():
        zip_base = os.path.join(temp_zip_dir, prov_id)
        shutil.make_archive(zip_base, 'zip', info["source_path"])
        zipped[prov_id] = f"{zip_base}.zip"
        print(f"[ZIP] Created {zip_base}.zip")

    return zipped

# ================= 3/4 ping แล้วส่งไฟล์ ZIP =================
def transfer_zips(zipped, folders, server_map):
    print("[TRANSFER] Starting transfer")
    results = []

    for prov_id, zip_path in zipped.items():
        original_name = folders[prov_id]["folder_name"]
        network_path = server_map[prov_id]
        server_ip = extract_server_from_unc(network_path)

        # Ping
        try:
            subprocess.run(
                ['ping', '-n', '1', '-w', '1000', server_ip],
                check=True,
                capture_output=True
            )
            print(f"[PING] OK: {server_ip}")
        except Exception:
            print(f"[PING] FAIL: {server_ip}")
            results.append((
                original_name,
                network_path,
                "Failed",
                "Server Not Accessible",
                "Check VPN / server power",
                prov_id
            ))
            continue

        dest_dir = os.path.join(network_path, DEST_SUBFOLDER_NAME, API_FOLDER)
        dest_zip = os.path.join(dest_dir, os.path.basename(zip_path))

        try:
            os.makedirs(dest_dir, exist_ok=True)
            shutil.copy(zip_path, dest_zip)  # กำหนดให้เขียนทับไฟล์เดิม
            print(f"[TRANSFER] SUCCESS → {dest_zip}")
            results.append((
                original_name,
                dest_zip,
                "Success",
                "OK",
                "-",
                prov_id
            ))
        except Exception as e:
            category, fix = classify_transfer_error(e)
            print(f"[TRANSFER] FAIL ({category})")
            results.append((
                original_name,
                dest_zip,
                "Failed",
                category,
                fix,
                prov_id
            ))

    return results

# ================= 5) ลบไฟล์ ZIP ชั่วคราว =================
def cleanup_temp_zips(zipped, results):
    print("[CLEANUP] Removing temp ZIPs")

    for row in results:
        status, prov_id = row[2], row[5]
        if status == "Success":
            zip_path = zipped.get(prov_id)
            if zip_path and os.path.exists(zip_path):
                os.remove(zip_path)
                print(f"[CLEANUP] Deleted {zip_path}")

    print("[CLEANUP] Done")

# ================= 6) สร้างรายงาน =================
def generate_report(results, report_path):
    print("[REPORT] Writing Excel report")
    df = pd.DataFrame(results, columns=RESULT_COLUMNS)
    os.makedirs(os.path.dirname(report_path), exist_ok=True)
    df.to_excel(report_path, index=False)
    print(f"[REPORT] Saved → {report_path}")

# ================= MAIN =================
def main():
    print("========== PROCESS START ==========")

    server_map = load_server_mapping(IP_FILE, IP_SHEET_NAME)
    folders = discover_folders(SOURCE_DIR, server_map)

    if not folders:
        print("[EXIT] No valid folders")
        return

    zipped = zip_all_folders(folders, TEMP_ZIP_DIR)
    results = transfer_zips(zipped, folders, server_map)
    cleanup_temp_zips(zipped, results)
    generate_report(results, REPORT_PATH)

    print("========== APIRAT THE GREAT ==========")

if __name__ == "__main__":
    main()
