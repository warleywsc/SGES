Attribute VB_Name = "Módulo1"
'@Folder("Tests")
Option Explicit

Sub ListPivotsInfor()
    'Update 20141112
    Dim St    As Worksheet
    Dim NewSt As Worksheet
    Dim pt    As PivotTable
    Dim i, k  As Long
    Application.ScreenUpdating = False
    Set NewSt = Worksheets.Add
    i = 1: k = 2
    With NewSt
        .Cells(i, 1) = "Name"
        .Cells(i, 2) = "Source"
        .Cells(i, 3) = "Refreshed by"
        .Cells(i, 4) = "Refreshed"
        .Cells(i, 5) = "Sheet"
        .Cells(i, 6) = "Location"
        For Each St In ActiveWorkbook.Worksheets
            For Each pt In St.PivotTables
                i = i + 1
                .Cells(i, 1).Value = pt.Name
                .Cells(i, 2).Value = pt.SourceData
                .Cells(i, 3).Value = pt.RefreshName
                .Cells(i, 4).Value = pt.RefreshDate
                .Cells(i, 5).Value = St.Name
                .Cells(i, 6).Value = pt.TableRange1.Address
            Next
        Next
        .Activate
    End With
    Application.ScreenUpdating = True
End Sub
