import os
import fitz  # PyMuPDF
import pandas as pd

# 1. ระบุตำแหน่งโฟลเดอร์ E-Policy ของคุณ (ใช้ r นำหน้าเพื่อป้องกันปัญหาเครื่องหมาย \)
target_folder = r"C:\Users\Sangchai.s\Desktop\E-Policy"

results = []

# 2. วนลูปค้นหาไฟล์ .pdf ทั้งหมด (รวมในโฟลเดอร์ย่อยด้วย)
for root, dirs, files in os.walk(target_folder):
    for file in files:
        if file.lower().endswith(".pdf"):
            file_path = os.path.join(root, file)
            folder_name = os.path.basename(root)  # เก็บชื่อโฟลเดอร์ย่อยไว้ดูด้วย
            
            try:
                doc = fitz.open(file_path)
                page_count = len(doc)
                min_font_size = float('inf')
                smallest_text_sample = ""

                # 3. ตรวจสอบตัวอักษรทุกหน้าในไฟล์
                for page in doc:
                    blocks = page.get_text("dict")["blocks"]
                    for b in blocks:
                        if "lines" in b:  # ตรวจสอบเฉพาะส่วนที่เป็นข้อความ
                            for l in b["lines"]:
                                for s in l["spans"]:
                                    text = s["text"].strip()
                                    size = round(s["size"], 2)  # ขนาดฟอนต์ (ปัดเศษ 2 ตำแหน่ง)
                                    
                                    # ข้ามช่องว่างเปล่า และหาขนาดที่เล็กที่สุด
                                    if text and size > 0:
                                        if size < min_font_size:
                                            min_font_size = size
                                            smallest_text_sample = text

                doc.close()

                # กรณีไฟล์เป็นรูปภาพสแกนทั้งไฟล์ (ไม่มี Text ให้ดึงขนาดฟอนต์)
                if min_font_size == float('inf'):
                    min_font_size = "ไม่พบข้อความ (อาจเป็นไฟล์สแกน/รูปภาพ)"
                    smallest_text_sample = "-"

                results.append({
                    "โฟลเดอร์": folder_name,
                    "ชื่อไฟล์": file,
                    "จำนวนหน้า": page_count,
                    "ขนาดฟอนต์เล็กสุด (pt)": min_font_size,
                    "ตัวอย่างข้อความที่เล็กสุด": smallest_text_sample[:40] # ตัดมาแสดงแค่ 40 ตัวอักษรแรก
                })
                print(f"ตรวจสอบแล้ว: {file}")

            except Exception as e:
                results.append({
                    "โฟลเดอร์": folder_name,
                    "ชื่อไฟล์": file,
                    "จำนวนหน้า": "Error",
                    "ขนาดฟอนต์เล็กสุด (pt)": str(e),
                    "ตัวอย่างข้อความที่เล็กสุด": "-"
                })

# 4. บันทึกผลลัพธ์เป็นไฟล์ Excel ไว้ในโฟลเดอร์เดียวกัน
output_excel = os.path.join(target_folder, "PDF_Check_Report.xlsx")
df = pd.DataFrame(results)
df.to_excel(output_excel, index=False)

print(f"\nเสร็จสิ้น! บันทึกรายงานไว้ที่: {output_excel}")