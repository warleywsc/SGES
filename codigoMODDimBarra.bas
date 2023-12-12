Attribute VB_Name = "MODDimBarra"
'@Folder("SGES2020")
Option Explicit

Public Sub dimbarra()

    '    'dimensiona barra de rolagem em info
    '    Dim i As Integer, n As Double, t As Double
    '    For i = 19 To 26
    '
    '    n = WorksheetFunction.Sum(Range("u" & i).Rows.Height)
    '    t = t + n
    '
    '    Next
    '
    '    ActiveSheet.Shapes.Range(Array("Scroll Bar 26")).Height = t
    Info.Shapes("Scroll Bar 26").Width = 16.4
    Info.Shapes("Scroll Bar 48").Width = 16.4
    Info.Shapes("Scroll Bar 26").Height = Info.Range("tbHistMov").Height
    Info.Shapes("Scroll Bar 26").Top = Info.Range("tbHistMov").Top
    Info.Shapes("btnSalvaAtualExt").Top = Info.Range("P26").Top + Info.Range("OBS").Height + 5
    Info.Shapes("Scroll Bar 48").Height = Info.Range("tbHistServ").Height
    Info.Shapes("Scroll Bar 48").Top = Info.Range("tbHistServ").Top
    'Info.Shapes("btnSalvaAtualExt").Top = Info.Range("P26").Top + Info.Range("OBS").Height + 5
    Info.Shapes("Scroll Bar 48").Left = Info.ListObjects("tbHistMov").DataBodyRange.Left - 13
    Info.Shapes("Scroll Bar 26").Left = Info.ListObjects("tbHistServ").DataBodyRange.Left - 13
End Sub

