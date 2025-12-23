# Student Data Splitter (VBA)

ระบบแยกไฟล์ข้อมูลนักเรียนซ้ำซ้อนตามเขตพื้นที่การศึกษา สำหรับสำนักงานศึกษาธิการจังหวัดอุดรธานี

## ความสามารถ
- แยกข้อมูลจาก Sheet หลักตามรายชื่อในคอลัมน์ O
- รักษาค่า Data Validation (Dropdown) ในคอลัมน์ P
- รองรับชื่อเขตพื้นที่และข้อมูลภาษาไทย (Unicode)
- Export ไฟล์แยกรายเขตเป็นนามสกุล .xlsx

## วิธีใช้งาน
1. ลบ บรรทัดที่ไม่ใช่ข้อมูลออกโดยเฉพาะบรรทัดก่อนหัวคอลัมน์
2. ตั้งชื่อ Range `A1:A11` ใน Sheet ข้อมูลสถานะนักเรียนซ้ำซ้อน ว่า `StatusList`
3. กด Alt + F11 > Insert > Module แล้ววาง Code
4. กดรัน Macro `SplitDataByArea`

หาก code มี ????????? ให้
1. ไปที่ Control Panel > Region.
2. เลือกแถบ Administrative.
3. กดปุ่ม Change system locale....
4. เลือกเป็น Thai (Thailand).
5. Restart เครื่อง 1 ครั้ง
6. ใน Excel VBA ไปที่ Tools > Options > Editor Format > เลือก Font เป็น "Tahoma (Thai)" หรือ "Courier New (Thai)".
