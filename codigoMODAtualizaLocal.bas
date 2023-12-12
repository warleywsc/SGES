Attribute VB_Name = "MODAtualizaLocal"
Option Explicit
'@Folder("SGES2020")

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub AtualizaLocal()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub AtualizaLocal()
    Dim ws    As Worksheet
    Set ws = ActiveWorkbook.Sheets.[_Default]("Locais")
    Dim tbl   As ListObject
    Set tbl = ws.ListObjects.[_Default]("tbLocalNovo")
    Dim sortcolumn As Range
    Set sortcolumn = Range("tbLocalNovo[LOCAL]")
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

End Sub


Public Sub Atualizatblocal()

    With ActiveWorkbook.Worksheets("Locais").ListObjects("tbLocalNovo").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("tbLocalNovo[[#All],[Local]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub




