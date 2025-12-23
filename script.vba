' *************************************************************
' Student Data Splitter (Unicode & Sheet Inclusion Support)
' Purpose: Split student data by Area (Col O) and include Status Reference
' *************************************************************

Sub SplitDataByAreaWithStatus()
    Dim wsMain As Worksheet, wsStatus As Worksheet
    Dim lastRow As Long, i As Long
    Dim areaList As New Collection
    Dim areaItem As Variant
    Dim newWB As Workbook
    Dim savePath As String
    
    ' ใช้ Index เพื่อเลี่ยงปัญหา Font ภาษาไทยใน VBE
    ' สมมติ: Sheet main คือลำดับที่ 1, ข้อมูลสถานะฯ คือลำดับที่ 2
    Set wsMain = ThisWorkbook.Sheets(1)   
    Set wsStatus = ThisWorkbook.Sheets(2) 
    
    savePath = ThisWorkbook.Path & "\"
    lastRow = wsMain.Cells(wsMain.Rows.Count, "O").End(xlUp).Row
    
    ' ปิดระบบแจ้งเตือนเพื่อความรวดเร็ว
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. หา Unique Area Names จาก Column O
    On Error Resume Next
    For i = 2 To lastRow
        If wsMain.Cells(i, "O").Value <> "" Then
            areaList.Add wsMain.Cells(i, "O").Value, CStr(wsMain.Cells(i, "O").Value)
        End If
    Next i
    On Error GoTo 0
    
    ' 2. เริ่มขั้นตอนการแยกไฟล์
    For Each areaItem In areaList
        ' สร้าง Workbook ใหม่
        Set newWB = Workbooks.Add
        
        ' --- จุดสำคัญ: Copy Sheet ข้อมูลสถานะนักเรียนซ้ำซ้อน เข้าไปเป็น Sheet แรก ---
        wsStatus.Copy Before:=newWB.Sheets(1)
        newWB.Sheets(1).Name = wsStatus.Name ' ตั้งชื่อให้เหมือนต้นฉบับ
        ' หากต้องการซ่อน Sheet อ้างอิงให้ปลดคอมเมนต์บรรทัดล่างนี้
        ' newWB.Sheets(1).Visible = xlSheetHidden 
        
        ' 3. กรองข้อมูลเฉพาะเขตนั้นๆ
        wsMain.AutoFilterMode = False
        wsMain.Range("A1:R" & lastRow).AutoFilter Field:=15, Criteria1:=areaItem
        
        ' 4. Copy ข้อมูลที่กรองแล้วไปยัง Sheet ที่ 2 ของไฟล์ใหม่
        wsMain.Range("A1:R" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        
        With newWB.Sheets(2)
            .Name = "Main_Data" ' ตั้งชื่อ Sheet ข้อมูล
            .Range("A1").PasteSpecial xlPasteAll
            .Columns("A:R").AutoFit
        End With
        
        ' ลบ Sheet ที่เกินมา (Excel มักจะสร้างไฟล์ใหม่โดยมี Sheet1 ติดมาด้วย)
        If newWB.Sheets.Count > 2 Then
            newWB.Sheets(newWB.Sheets.Count).Delete
        End If
        
        ' 5. บันทึกไฟล์ (ใช้ชื่อเขตเป็นชื่อไฟล์)
        newWB.SaveAs Filename:=savePath & areaItem & ".xlsx", FileFormat:=51
        newWB.Close SaveChanges:=False
    Next areaItem
    
    ' คืนค่าระบบ
    wsMain.AutoFilterMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "สร้างไฟล์แยกรายเขตเรียบร้อยแล้ว ทั้งหมด " & areaList.Count & " ไฟล์", vbInformation
End Sub
