Attribute VB_Name = "MODTelacheia"
Option Explicit
'@Folder("SGES2020")

Public Sub Telacheia()

    Dim ws    As Worksheet

    For Each ws In ActiveWorkbook.Sheets
   
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        ws.Activate
        Application.StatusBar = False
    
        ActiveWindow.DisplayWorkbookTabs = False
    
        Application.DisplayFormulaBar = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        Application.DisplayFullScreen = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    
    Next
End Sub

