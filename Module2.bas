Attribute VB_Name = "Module2"
Option Explicit

Public Sub Coordinates()
    
    '' This program takes the coordinates of the selected columns
    '' (select only the column numbers) and creates a file for QPM.
    
    '' Format csv is bad?
    
    Dim ssetObj As AcadSelectionSet
    Dim i As Integer
    Dim circle_center As Variant
    Dim Row As Integer
    Dim ExcelApplication As Excel.Application
    Dim ExcelWorksheet As Worksheet
    Dim column_prefix As String
    Dim fileName As String
    Dim CSVfilePath As String
    Dim Drawing As AcadDocument
    Dim DXFfilePath As String
    
    ''Input messages
    MsgBox ("Rysunek musi byæ w metrach!" & vbCrLf & "Je¿eli twój rysunek jest _" & _
            " w innych jednostkach, przeskaluj rysunek.")
    MsgBox ("Zaznacz tylko numery kolumn (C1, C2, C3, ... , Cn).")
    
    '' Setting selection Acad object
    Set ssetObj = ThisDrawing.SelectionSets.Add("Selection")
        ssetObj.SelectOnScreen
    
    '' Setting Excel object
    Set ExcelApplication = CreateObject("Excel.Application")
        ExcelApplication.Visible = True
        ExcelApplication.Workbooks.Add
        
    On Error GoTo EH
        Set ExcelWorksheet = ExcelApplication.ActiveWorkbook.Sheets("Arkusz1")
EH:
        Set ExcelWorksheet = ExcelApplication.ActiveWorkbook.Sheets("Sheet1")
        '' Setting sheet name
        fileName = Left(ThisDrawing.name, Len(ThisDrawing.name) - 4)
        ExcelWorksheet.name = Right(fileName, 31)
        
    '' Entering column coordinates and prefixes to Excel
    Row = 1
    For i = (ssetObj.Count - 1) To 0 Step -1
        
        '' Data
        circle_center = ssetObj.Item(i).InsertionPoint
        column_prefix = ssetObj.Item(i).TextString
        
        '' Entering
        ExcelWorksheet.Cells(Row, 1).Value = column_prefix
        ExcelWorksheet.Cells(Row, 2).Value = Round(circle_center(0), 2)
        ExcelWorksheet.Cells(Row, 3).Value = Round(circle_center(1), 2)
        ExcelWorksheet.Cells(Row, 4).Value = Round(circle_center(2), 2)
        
        Row = Row + 1
        
    Next i
    
    '' Saving files
    '' Saving to dxf file
    Set Drawing = Application.ActiveDocument
    
    DXFfilePath = "QPM_" & fileName
    
    Drawing.SaveAs DXFfilePath, ac2000_dxf

    '' Saving to csv file
    CSVfilePath = Left(Drawing.FullName, Len(Drawing.FullName) - _
    Len(Drawing.name)) + ("QPM_" & fileName)

    ActiveWorkbook.SaveAs fileName:=CSVfilePath, _
        FileFormat:=xlCSVWindows, CreateBackup:=False

    '' Deleting Acad object
    ssetObj.Delete

End Sub
