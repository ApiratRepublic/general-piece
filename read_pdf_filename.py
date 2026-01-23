###############################################################
############# อ่านชื่อไฟล์ pdf ขนาดไฟล์ และวันที่แก้ไขล่าสุดจากเซิร์ฟเวอร์
############## เขียนผลลัพธ์ลงไฟล์ข้อความ
############# Apirat Rattanapaiboon
############# 2567-06-12
############# เวอร์ชัน 1.0
###############################################################

import os
import datetime

# กำหนดตำแหน่งที่ตั้งไฟล์และไฟล์ผลลัพธ์
root_directory = r'V:\2566-2569'
report_filename = r"D:\A02-Projects\Raster\all-map-name-with-date.txt"

pdf_files_details = []

# --- ตรวจสอบว่าไดเรกทอรีมีอยู่จริงหรือไม่ ---
if not os.path.exists(root_directory):
    print(f"ผิดพลาดนะจ๊ะ ไม่มี '{root_directory}' อยู่ในระบบ")
    exit(1)

# --- สำรวจและค้นหาไฟล์ PDF ในไดเรกทอรีทั้งหมด ---
# os.walk จะสำรวจทุกโฟลเดอร์ย่อยภายใน root_directory
for root, _, files in os.walk(root_directory):
    for filename in files:
        if filename.lower().endswith('.pdf'): # ใช้ .lower() เพื่อให้เจอทั้ง .pdf และ .PDF
            file_path = os.path.join(root, filename)
            
            try:
                # --- ดึงข้อมูลเพิ่มเติม: ขนาดไฟล์ และ วันที่แก้ไขล่าสุด ---
                file_size = os.path.getsize(file_path)
                mod_time_stamp = os.path.getmtime(file_path)
                
                # --- แปลง timestamp เป็นรูปแบบวันที่ที่อ่านง่าย ---
                mod_date_time = datetime.datetime.fromtimestamp(mod_time_stamp)
                formatted_date = mod_date_time.strftime('%Y-%m-%d %H:%M:%S')

                # --- เก็บข้อมูลทั้งหมดไว้ใน list ---
                pdf_files_details.append((file_path, file_size, formatted_date))

            except FileNotFoundError:
                print(f"คำเตือน: ไม่พบไฟล์นี้ระหว่างดำเนินการ: {file_path}")
                continue

# --- ส่งผลลัพธ์ไปยังไฟล์ ---
if not pdf_files_details:
    print("ไม่พบไฟล์ PDF ใด ๆ ในไดเรกทอรีที่ระบุ")
    exit(0)

# เขียนข้อมูลลงไฟล์ text
with open(report_filename, 'w', encoding='utf-8') as file:
    # เขียนหัวคอลัมน์เพื่อความชัดเจน
    file.write("Full Path, Size (bytes), Last Modified\n")
    for path, size, mod_date in pdf_files_details:
        file.write(f'"{path}", {size}, {mod_date}\n')

print("รายชื่อไฟล์ PDF พร้อมขนาดและวันที่แก้ไขล่าสุดถูกบันทึกเรียบร้อยแล้ว")
print(f"จำนวนไฟล์ PDF ที่พบ: {len(pdf_files_details)}")