Attribute VB_Name = "MODFiltrapesquisa"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Concei��o
' Contato...: warleywsc@gmail.com - Rotina: Sub filtrapesquisa()
    ' Data......: 26/11/2020
    ' Descricao.: Filtro da Planilha Pesquisa
    '---------------------------------------------------------------------------------------
Public Sub filtrapesquisa()

    Dim ultlinhaP As Long
    Dim ultlinhaI As Long
    Dim ultlinhaj As Long
    Dim rng   As Range
    Dim tbl   As ListObject

    Pesquisa.Unprotect "brigada"
    Pesquisa.Range("G12:AB1048576").ClearContents
    
    
    ultlinhaI = Cells(Rows.Count, 14).End(xlUp).Row 'ultima linha usada na coluna s�rie
    
    MapaAtual.Range("tbMapaAtual[#All]").AdvancedFilter Action:= _
                                                        xlFilterCopy, CriteriaRange:=Range("criteriostudo"), CopyToRange:=Range( _
                                                                                                                           "tbPesquisaMapaAtual[#Headers]"), Unique:=True
    ultlinhaj = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[S�rie]]").End(xlDown).Row
    
    With ActiveWorkbook.Worksheets("Pesquisa").ListObjects("tbPesquisaMapaAtual")
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range("tbPesquisaMapaAtual[Edif�cio]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("tbPesquisaMapaAtual[�rea]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.Apply
    End With
  
   
   
    
    If ultlinhaj < ultlinhaI Then
    
    
        Set tbl = Pesquisa.ListObjects("tbPesquisaMapaAtual")



        Set rng = Pesquisa.Range("tbPesquisaMapaAtual[#All]").Resize(ultlinhaj - 10, 23)
        tbl.Resize rng
    
    Else
        Set tbl = Pesquisa.ListObjects("tbPesquisaMapaAtual")

        ultlinhaI = Cells(Rows.Count, 14).End(xlUp).Row 'ultima linha usada na coluna s�rie

        Set rng = Pesquisa.Range("tbPesquisaMapaAtual[#All]").Resize(ultlinhaI - 10, 23)
        tbl.Resize rng
    
    End If
    Pesquisa.Calculate
    Pesquisa.Protect "brigada", DrawingObjects:=True, Contents:=True, Scenarios:=True _

End Sub


'---------------------------------------------------------------------------------------
' Programador.....: Warley S Concei��o
' Contato...: warleywsc@gmail.com - Rotina: Sub filtraimpressao()
    ' Data......: 26/11/2020
    ' Descricao.: Copia Resultado da pesquis para planilha Impressao
    '---------------------------------------------------------------------------------------
Public Sub filtraimpressao()

    Dim ultlinhaP As Long
    Dim ultlinhaI As Long
    Dim rng   As Range
    Dim tbl   As ListObject


    Impressao.Range("E3:K1048576").ClearContents
    Application.CutCopyMode = False
    Pesquisa.Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                               xlFilterCopy, CriteriaRange:=Range("criteriostudo"), CopyToRange:=Range( _
                                                                                                                                  "tbImpressao[#Headers]"), Unique:=False

    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[S�rie]]").End(xlDown).Row

    Set tbl = Impressao.ListObjects("tbImpressao")



    Set rng = Impressao.Range("tbImpressao[#aLL]").Resize(ultlinhaI - 10, 22)
    tbl.Resize rng


End Sub

Public Sub filtraPOT()

    Dim ultlinhaP As Long
    Dim ultlinhaI As Long
    Dim rng   As Range
    Dim tbl   As ListObject


    Impressao.Range("E3:K1048576").ClearContents
    Application.CutCopyMode = False
    Pesquisa.Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                               xlFilterCopy, CriteriaRange:=Range("criteriostudo"), CopyToRange:=Range( _
                                                                                                                                  "tbImpressaopot[#Headers]"), Unique:=False

    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[S�rie]]").End(xlDown).Row

    Set tbl = Impressao1.ListObjects("tbImpressaopot")



    Set rng = Impressao1.Range("tbImpressaopot[#aLL]").Resize(ultlinhaI - 10, 23)
    tbl.Resize rng


End Sub
'---------------------------------------------------------------------------------------
' Programador.....: Warley S Concei��o
' Contato...: warleywsc@gmail.com - Rotina: Sub filtraformpes()
    ' Data......: 26/11/2020
    ' Descricao.: Gera Form de Pesagem
    '---------------------------------------------------------------------------------------
Public Sub filtraformpes()

    Dim ultlinhaP As Long
    Dim ultlinhaI As Long

    Dim rng   As Range
    Dim tbl   As ListObject
    Impressaopes.Range("E3:R1048576").ClearContents

    Application.CutCopyMode = False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaopes[[#Headers],[Sup]:[S�rie]]"), Unique:=False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaopes[[#Headers],[Pr�ximo Teste]:[Pr�ximo Teste]]"), Unique:=False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaopes[[#Headers],[Pr�xima Recarga]:[Pr�xima Recarga]]"), Unique:=False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaopes[[#Headers],[Observa��o]:[Observa��o]]"), Unique:=False

    ultlinhaP = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[S�rie]]").End(xlDown).Row
    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[S�rie]]").End(xlDown).Row


    Set tbl = Impressaopes.ListObjects("tbImpressaopes")

    Set rng = Range("tbImpressaopes[#All]").Resize(ultlinhaI - 10, 14)
    tbl.DataBodyRange.Rows.RowHeight = 40
    tbl.HeaderRowRange.Font.Size = 16
    tbl.DataBodyRange.Font.Size = 14
    tbl.Resize rng
    tbl.DataBodyRange.Rows.RowHeight = 40
    Impressaopes.Range("K12:K" & ultlinhaP).Columns.AutoFit

    '######################################

End Sub


'---------------------------------------------------------------------------------------
' Programador.....: Warley S Concei��o
' Contato...: warleywsc@gmail.com - Rotina: Sub filtraformtroca()
    ' Data......: 26/11/2020
    ' Descricao.: Gera Form de Reposi��o
    '---------------------------------------------------------------------------------------
Public Sub filtraformtroca()
    Dim ultlinhaP As Long
    Dim ultlinhaI As Long

    Dim rng   As Range
    Dim tbl   As ListObject

    Impressaotroca.Range("E3:W1048576").ClearContents


    Application.CutCopyMode = False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaotroca[[#Headers],[Sup]:[S�rie]]"), Unique:=False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaotroca[[#Headers],[Observa��o]:[Observa��o]]"), Unique:=False
    ultlinhaP = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[S�rie]]").End(xlDown).Row
    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[S�rie]]").End(xlDown).Row

    Set tbl = Impressaotroca.ListObjects("tbImpressaotroca")


    Set rng = Range("tbImpressaotroca[#All]").Resize(ultlinhaI - 10, 19)
    tbl.Resize rng
    tbl.DataBodyRange.Rows.RowHeight = 40
    tbl.HeaderRowRange.Font.Size = 16
    tbl.DataBodyRange.Font.Size = 14
    Impressaotroca.Range("E12:E" & ultlinhaP).Columns.AutoFit
    Impressaotroca.Range("G12:G" & ultlinhaP).Columns.AutoFit
    Impressaotroca.Range("H12:H" & ultlinhaP).Columns.AutoFit

End Sub




