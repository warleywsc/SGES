Attribute VB_Name = "Módulo3"
Option Explicit

Sub Macro1()
    '
    ' Macro1 Macro
    '

    '
    Selection.NumberFormat = "m/d/yyyy"
    Range("U9,Z9,AA9,AB9").Select
    Range("Tabela25[Data de Envio]").Activate
    Selection.NumberFormat = "m/d/yyyy"
End Sub
Sub Macro2()
    '
    ' Macro2 Macro
    '

    '
    Selection.NumberFormat = "0.00"
End Sub
Sub Macro3()
    '
    ' Macro3 Macro
    '

    '
    Selection.NumberFormat = "@"
End Sub
Sub Macro4()
    '
    ' Macro4 Macro
    '

    '
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
