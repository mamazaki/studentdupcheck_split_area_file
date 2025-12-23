Sub SplitSheetByAreaAndSaveAsFile()
    Dim mainSheet As Worksheet
    Dim statusSheet As Worksheet
    Dim uniqueAreas As Collection
    Dim areaName As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim newWB As Workbook
    Dim folderPath As String
    
    Set mainSheet = ThisWorkbook.Sheets("main")
    Set statusSheet = ThisWorkbook.Sheets("ข้อมูลสถานะนักเรียนซ้ำซ้อน")
    Set uniqueAreas = New Collection
    
    ' 1. หาพาธที่จะเซฟไฟล์ (เซฟที่เดียวกับไฟล์หลัก)
    folderPath = ThisWorkbook.Path & "\"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 2. ดึงรายชื่อเขตพื้นที่ฯ ที่ไม่ซ้ำกันจาก Column O (เริ่มแถวที่ 2)
    lastRow = mainSheet.Cells(mainSheet.Rows.Count, "O").End(xlUp).Row
    On Error Resume Next
    For i = 2 To lastRow
        If mainSheet.Cells(i, "O").Value <> "" Then
            uniqueAreas.Add mainSheet.Cells(i, "O").Value, CStr(mainSheet.Cells(i, "O").Value)
        End If
    Next i
    On Error GoTo 0
    
    ' 3. วนลูปสร้างไฟล์ตามรายชื่อเขต
    For Each areaName In uniqueAreas
        ' สร้าง Workbook ใหม่
        Set newWB = Workbooks.Add
        
        ' Copy Sheet แหล่งอ้างอิงไปด้วยเพื่อให้ Validation ไม่พัง
        statusSheet.Copy Before:=newWB.Sheets(1)
        newWB.Sheets(1).Visible = xlSheetHidden ' ซ่อนไว้เพื่อความสวยงาม
        
        ' กลับมา Copy ข้อมูลที่กรองแล้ว
        mainSheet.AutoFilterMode = False
        mainSheet.Range("A1:R" & lastRow).AutoFilter Field:=15, Criteria1:=areaName
        
        ' Copy ข้อมูลไปยัง Workbook ใหม่ Sheet ที่ 2
        mainSheet.Range("A1:R" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        With newWB.Sheets(2)
            .Name = Left(areaName, 30) ' ชื่อ Sheet ตามเขต
            .Paste Destination:=.Range("A1")
            .Columns("A:R").AutoFit
        End With
        
        ' ลบ Sheet ว่างที่เหลือ
        Do While newWB.Sheets.Count > 2
            newWB.Sheets(newWB.Sheets.Count).Delete
        Loop
        
        ' บันทึกไฟล์แยกตามชื่อเขต
        newWB.SaveAs Filename:=folderPath & areaName & ".xlsx"
        newWB.Close SaveChanges:=False
    Next areaName
    
    mainSheet.AutoFilterMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "แยกไฟล์สำเร็จทั้งหมด " & uniqueAreas.Count & " เขต", vbInformation
End Sub
