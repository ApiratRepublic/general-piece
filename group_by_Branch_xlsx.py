##############################
#### ใช้สำหรับแยกไฟล์ Excel ตามคอลัมน์ BRANCH_CODE ####
#### แล้วส่งออกเป็นไฟล์ Excel แยกต่างหาก ####
#### APIRAT RATTANAPAIBOON ####
##############################


import pandas as pd
import os

# ตำแหน่งที่ตั้ง
input_excel = r"C:\Users\i-api\Downloads\Telegram Desktop\Results_rp1_not in parcel.xlsx"
output_dir = r"D:\A02-Projects\WarRoom\SE"

# สร้างโฟลเดอร์ถ้ายังไม่มี
os.makedirs(output_dir, exist_ok=True)

# อ่านข้อมูลจาก Excel
df = pd.read_excel(input_excel)

#  ตรวจสอบว่าคอลัมน์ 'BRANCH_CODE' มีอยู่ใน DataFrame หรือไม่
if "BRANCH_CODE" not in df.columns:
    raise ValueError("Column 'BRANCH_CODE' not found in the Excel file")

# จัดกลุ่มและส่งออกเป็นไฟล์ Excel แยกตาม 'BRANCH_CODE'
for branch_code, group_df in df.groupby("BRANCH_CODE"):
    # เปลี่ยนชื่อไฟล์ตาม 'BRANCH_CODE'
    filename = f"{str(branch_code).strip()}.xlsx"
    output_path = os.path.join(output_dir, filename)

    # เขียนข้อมูลลงในไฟล์ Excel
    group_df.to_excel(output_path, index=False)

#### APIRAT RATTANAPAIBOON ####
print("ส่งออกไฟล์ Excel เรียบร้อยแล้ว")