Attribute VB_Name = "MODatualizatbdinlocal"
Option Explicit
'@Folder("SGES2020")

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub atualizatbdinlocal()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub atualizatbdinlocal()
  
    locais.PivotTables("tbdimLocal").PivotCache.Refresh
    locais.PivotTables("tbdimLocal").PivotFields("Local").AutoSort _
        xlAscending, "Local"
End Sub









