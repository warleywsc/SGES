Attribute VB_Name = "MODAdditemfiltropvt"
Option Explicit

Public Sub addtofiltropvt()
    '
    ' Macro1 Macro
    '

    '
    With Planilha8.PivotTables("Tabela din�mica2").PivotFields("S�rie")
        .PivotItems(vbNullString & .Range("N2").Value & vbNullString).Visible = True
    End With
End Sub
Public Sub removefiltropvt()
    '
    ' removefiltropvt Macro
    '

    '
    With Planilha8.PivotTables("Tabela din�mica2").PivotFields("S�rie")
        .PivotItems(vbNullString & .Range("N4").Value & vbNullString).Visible = False
    End With
End Sub
