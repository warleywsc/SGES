Attribute VB_Name = "MODfiltroavancado"
Option Explicit
'@Folder("SGES2020")

Public Sub filtreme()

    Dim ultlinhaP As Long
    Dim ultlinhaI As Long
    
    Dim rng   As Range
    Dim tbl   As ListObject
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    '    Pesquisa.Range("G12:AB1048576").ClearContents
    '
    '    MapaAtual.Range("tbMapaAtual[#All]").AdvancedFilter Action:= _
    '    xlFilterCopy, CriteriaRange:=Range("criteriostudo"), CopyToRange:=Range( _
    '     "tbPesquisaMapaAtual[#Headers]"), Unique:=False
    Impressao.Range("E3:K1048576").ClearContents
    MapaAtual.Range("tbMapaAtual[#All]").AdvancedFilter Action:= _
                                                        xlFilterCopy, CriteriaRange:=Range("criteriostudo"), CopyToRange:=Range( _
                                                                                                                           "tbImpressao[#Headers]"), Unique:=False
   
    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    
    Set tbl = Impressao.ListObjects("tbImpressao")



    Set rng = Impressao.Range("tbImpressao[#aLL]").Resize(ultlinhaI - 10, 22)
    tbl.Resize rng

    '    '#########################
    '
    '    'form troca
    Impressaotroca.Range("E3:K1048576").ClearContents
    
    
    Application.CutCopyMode = False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaotroca[[#Headers],[Sup]:[Série]]"), Unique:=False
    ultlinhaP = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row

    Set tbl = Impressaotroca.ListObjects("tbImpressaotroca")


    Set rng = Range("tbImpressaotroca[#All]").Resize(ultlinhaI - 10, 19)
    tbl.Resize rng
    tbl.DataBodyRange.Rows.RowHeight = 40
    Impressaotroca.Range("E12:E" & ultlinhaP).Columns.AutoFit
    Impressaotroca.Range("G12:G" & ultlinhaP).Columns.AutoFit
    Impressaotroca.Range("H12:H" & ultlinhaP).Columns.AutoFit

    '    '########################
    'form pesagem
    Impressaopes.Range("E3:K1048576").ClearContents
        
    Application.CutCopyMode = False
    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
                                                                         xlFilterCopy, CriteriaRange:=Range("Pesquisa!Criteria"), CopyToRange:=Range _
                                                                                                                                                ("tbImpressaopes[[#Headers],[Sup]:[Série]]"), Unique:=False
    
    ultlinhaP = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    

    Set tbl = Impressaopes.ListObjects("tbImpressaopes")

    Set rng = Range("tbImpressaopes[#All]").Resize(ultlinhaI - 10, 14)
    tbl.Resize rng
    tbl.DataBodyRange.Rows.RowHeight = 40

    '######################################
    Application.GoTo Pesquisa.Range("G12")
    Pesquisa.Range("G12").Activate

    Application.CutCopyMode = False

    Application.ScreenUpdating = True

    Application.EnableEvents = True

End Sub

Public Sub limpeme()
    
    
    ' Limpa Filtro
    
    ' Atalho do teclado: Ctrl+Shift+C
    
   
    Range("tbPesquisaMapaAtual").Clear
End Sub









