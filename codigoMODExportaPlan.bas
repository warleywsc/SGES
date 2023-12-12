Attribute VB_Name = "MODExportaPlan"
'@Folder "Tests"
Option Explicit


Public Sub ExportWorksheets()
    
    Dim wbSource As Workbook, wbTarget As Workbook
    Dim worksheetList As String        'Use colon as seperator since you cannot have colon in your worksheet name
    Dim worksheetArr As Variant
    Dim arrIndx As Long

    On Error GoTo errHandle
    'worksheetList = "Mapa Atual:Movimentação:Serviços:"
    
    worksheetArr = MapaAtual.ListObjects(1)
    
    If UBound(worksheetArr) = -1 Then Exit Sub
    
    Set wbSource = ThisWorkbook
    
    Set wbTarget = Workbooks.Add
    
    For arrIndx = LBound(worksheetArr) To UBound(worksheetArr)
        ThisWorkbook.Worksheets(worksheetArr(arrIndx)).Copy wbTarget.Worksheets(wbTarget.Worksheets.Count)
    Next arrIndx
    
    MsgBox "Export complete.", vbInformation
     
CleanObjects:
    Set wbTarget = Nothing
    Set wbSource = Nothing

    Exit Sub

errHandle:
    MsgBox "Error: " & Err.Description, vbExclamation
    GoTo CleanObjects
End Sub
