Sub ImportExcelFilesAsSheets_SafeCopy()
    Dim FolderPath As String
    Dim Filename As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim wbTarget As Workbook
    Dim SheetName As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Set your folder path here
    FolderPath = "C:\Users\Eva\Downloads\Automotive_Industry\Data\"  ' <- CHANGE THIS
    
    Set wbTarget = ThisWorkbook
    Filename = Dir(FolderPath & "*.xls*")
    
    Do While Filename <> ""
        Set wbSource = Workbooks.Open(FolderPath & Filename, ReadOnly:=True)
        Set wsSource = wbSource.Sheets(1)  ' Use first sheet
        
        ' Sheet name based on file name without extension
        SheetName = Left(Filename, InStrRev(Filename, ".") - 1)
        
        ' Ensure sheet name is unique
        Dim i As Integer: i = 1
        Dim TempName As String: TempName = SheetName
        Do While SheetExists(TempName, wbTarget)
            TempName = SheetName & "_" & i
            i = i + 1
        Loop
        SheetName = TempName
        
        ' Create new sheet in target workbook
        Set wsTarget = wbTarget.Sheets.Add(After:=wbTarget.Sheets(wbTarget.Sheets.Count))
        wsTarget.Name = SheetName
        
        ' Copy contents manually
        wsSource.UsedRange.Copy Destination:=wsTarget.Range("A1")
        
        wbSource.Close False
        Filename = Dir()
    Loop
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "All files imported successfully!", vbInformation
End Sub

Function SheetExists(SheetName As String, wb As Workbook) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Sheets(SheetName) Is Nothing
    On Error GoTo 0
End Function

