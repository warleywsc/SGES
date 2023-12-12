Attribute VB_Name = "Módulo8"
'@Folder "Tests"
Option Explicit

Public Sub tempodecorrido()

    Dim tempo As Double

    tempo = Timer
    'Call statusservico
    atualizamapaatual
    MsgBox Int(Timer - tempo) * 1000



End Sub


Public Sub addinfo()

    Dim tblfonte As ListObject
    Dim tbldestino As ListObject

    'Get a reference to the table you want to copy
    Set tblfonte = Pesquisa.ListObjects(1)
    'Get a reference to the destination table
    Set tbldestino = Sheets(tblfonte.ListRows(1).Range.Cells(1).Value).ListObjects(1)

    'Copy the source to a new row
    tblfonte.DataBodyRange.Copy tbldestino.ListRows.Add.Range

End Sub


Public Sub maiuscula()
    Dim cell  As Range
    Application.ScreenUpdating = False
    For Each cell In Application.Selection
        cell.Value = UCase$(cell.Value)
    Next cell
    Application.ScreenUpdating = True
End Sub

