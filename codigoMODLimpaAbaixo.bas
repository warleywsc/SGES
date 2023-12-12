Attribute VB_Name = "MODLimpaAbaixo"
Option Explicit

Public Sub LIMPALIXO()

    Range("G11").End(xlDown).Select
    Range(Cells(Selection.Offset(1, 0).Row, Selection.Offset(1, 0).Column), Cells(1048576, 30)).Clear
End Sub

