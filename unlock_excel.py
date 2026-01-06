#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Password Unlocker
สคริปต์สำหรับปลด password ไฟล์ Excel
"""

import sys
import os
from pathlib import Path

def unlock_with_password(input_file, output_file, password):
    """
    ปลดล็อคไฟล์ Excel ที่มี File Protection (ต้องใส่รหัสเปิดไฟล์)
    ต้องรู้ password จึงจะใช้งานได้
    """
    try:
        import msoffcrypto
        
        print(f"🔓 กำลังพยายามปลดล็อคไฟล์ด้วย password...")
        
        with open(input_file, "rb") as f:
            file = msoffcrypto.OfficeFile(f)
            file.load_key(password=password)
            
            with open(output_file, "wb") as out:
                file.decrypt(out)
        
        print(f"✅ สำเร็จ! ไฟล์ถูกบันทึกที่: {output_file}")
        return True
        
    except Exception as e:
        print(f"❌ ไม่สามารถปลดล็อคด้วย password ได้: {str(e)}")
        print(f"   - ตรวจสอบว่า password ถูกต้องหรือไม่")
        print(f"   - หรือไฟล์อาจใช้การเข้ารหัสแบบอื่น")
        return False

def unlock_sheet_protection(input_file, output_file):
    """
    ปลดล็อค Sheet Protection (การล็อคเซลล์หรือแผ่นงาน)
    ไม่ต้องรู้ password
    """
    try:
        from openpyxl import load_workbook
        
        print(f"🔓 กำลังปลดล็อค Sheet Protection...")
        
        # โหลดไฟล์
        wb = load_workbook(input_file)
        
        # นับจำนวน sheet ที่ถูกล็อค
        locked_sheets = 0
        
        # ปลดล็อคทุก sheet
        for sheet in wb.worksheets:
            if sheet.protection.sheet:
                locked_sheets += 1
                sheet.protection.sheet = False
                sheet.protection.password = None
                print(f"   📄 ปลดล็อค sheet: {sheet.title}")
        
        if locked_sheets == 0:
            print(f"   ℹ️  ไม่พบ sheet ที่ถูกล็อค (Sheet Protection)")
            return False
        
        # บันทึกไฟล์
        wb.save(output_file)
        print(f"✅ สำเร็จ! ปลดล็อค {locked_sheets} sheet(s)")
        print(f"   ไฟล์ถูกบันทึกที่: {output_file}")
        return True
        
    except Exception as e:
        print(f"❌ ไม่สามารถปลดล็อค Sheet Protection ได้: {str(e)}")
        return False

def main():
    """ฟังก์ชันหลักของโปรแกรม"""
    
    print("=" * 60)
    print("🔐 Excel Password Unlocker")
    print("   โปรแกรมปลด password ไฟล์ Excel")
    print("=" * 60)
    print()
    
    # ตรวจสอบ dependencies
    try:
        import msoffcrypto
        import openpyxl
    except ImportError as e:
        print("❌ ขาด library ที่จำเป็น กรุณาติดตั้งก่อน:")
        print("   pip install -r requirements.txt")
        print()
        print(f"   Error: {e}")
        sys.exit(1)
    
    # รับชื่อไฟล์จากผู้ใช้
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = input("📁 ระบุชื่อไฟล์ Excel ที่ต้องการปลดล็อค: ").strip()
    
    # ตรวจสอบว่าไฟล์มีอยู่จริง
    if not os.path.exists(input_file):
        print(f"❌ ไม่พบไฟล์: {input_file}")
        sys.exit(1)
    
    # สร้างชื่อไฟล์ output
    file_path = Path(input_file)
    output_file = file_path.parent / f"unlocked_{file_path.name}"
    
    print()
    print(f"📂 ไฟล์ต้นฉบับ: {input_file}")
    print(f"📂 ไฟล์ที่จะบันทึก: {output_file}")
    print()
    
    # เลือกวิธีการปลดล็อค
    print("เลือกวิธีการปลดล็อค:")
    print("1. ปลดล็อค Sheet Protection (ไม่ต้องรู้ password)")
    print("2. ปลดล็อค File Protection (ต้องรู้ password)")
    print("3. ลองทั้งสองวิธี")
    print()
    
    choice = input("เลือก (1/2/3) [ค่าเริ่มต้น: 3]: ").strip() or "3"
    print()
    
    success = False
    
    if choice in ["1", "3"]:
        # ลองปลดล็อค Sheet Protection ก่อน
        print("--- วิธีที่ 1: Sheet Protection ---")
        if unlock_sheet_protection(input_file, output_file):
            success = True
        print()
    
    if choice in ["2", "3"]:
        # ปลดล็อค File Protection (ต้องรู้ password)
        print("--- วิธีที่ 2: File Protection ---")
        password = input("🔑 ใส่ password (ถ้ารู้): ").strip()
        
        if password:
            temp_output = file_path.parent / f"temp_{file_path.name}"
            if unlock_with_password(input_file, temp_output, password):
                # ถ้าปลดล็อคสำเร็จ ลองปลดล็อค sheet protection ด้วย
                if unlock_sheet_protection(temp_output, output_file):
                    os.remove(temp_output)
                else:
                    os.rename(temp_output, output_file)
                success = True
        else:
            print("ℹ️  ข้าม File Protection (ไม่ได้ระบุ password)")
        print()
    
    # สรุปผลการทำงาน
    print("=" * 60)
    if success:
        print("🎉 เสร็จสิ้น! ตรวจสอบไฟล์ได้ที่:")
        print(f"   {output_file}")
    else:
        print("⚠️  ไม่สามารถปลดล็อคได้")
        print()
        print("💡 คำแนะนำ:")
        print("   - ถ้าต้องใส่ password เปิดไฟล์ = File Protection")
        print("     → ต้องรู้ password จึงจะปลดล็อคได้")
        print("   - ถ้าเปิดไฟล์ได้แต่แก้ไขไม่ได้ = Sheet Protection")
        print("     → ใช้วิธีที่ 1 ปลดล็อคได้เลย (ไม่ต้องรู้ password)")
    print("=" * 60)

if __name__ == "__main__":
    main()
