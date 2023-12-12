Attribute VB_Name = "MODImprimePesquisa"
Option Explicit
'@Folder("SGES2020")

Public Sub exibeop()
    With Pesquisa.Shapes("btnpot")

        If .Visible = False Then
            .Visible = True
           
    
        End If

    End With

    With Pesquisa.Shapes("btnpesquisa")

        If .Visible = False Then
            .Visible = True
           
    
        End If

    End With

    With Pesquisa.Shapes("btnreposicao")

        If .Visible = False Then
            .Visible = True
            
        End If

    End With
    With Pesquisa.Shapes("btnpesagem")

        If .Visible = False Then
            .Visible = True
            
        End If

    End With

End Sub

Public Sub ocultaop()

    With Pesquisa.Shapes("btnpot")

        If .Visible = True Then
            .Visible = False
    
        End If

    End With

    With Pesquisa.Shapes("btnpesquisa")

        If .Visible = True Then
            .Visible = False
    
        End If

    End With

    With Pesquisa.Shapes("btnreposicao")

        If .Visible = True Then
            .Visible = False
    
        End If

    End With
    With Pesquisa.Shapes("btnpesagem")

        If .Visible = True Then
            .Visible = False
    
        End If

    End With

End Sub

Public Sub limpaPesquisa()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Range("campospesquisa" _
          ).ClearContents
   
    Range("cpcriterios").Calculate
  
    '    Range("I3").Activate
    filtrapesquisa
    'filtraimpressao
    'filtraformpes
    'filtraformtroca
    ' filtreme
    ocultaop
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub ImprimePesquisa()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    filtraimpressao
    
    Impressao.Visible = xlSheetVisible
    Pesquisa.Activate
    Pesquisa.Range("K3").Activate
    Impressao.ListObjects("tbImpressao").Range.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Pesquisa.pdf", Quality:=xlQualityStandard, _
                                                                   IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Impressao.Visible = xlSheetHidden
    Application.EnableEvents = True
    Application.EnableEvents = True

End Sub


Public Sub ImprimePot()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    filtraPOT
    
   
    
    
    Impressao1.Visible = xlSheetVisible
    Pesquisa.Activate
    Pesquisa.Range("K3").Activate
    Impressao1.ListObjects("tbImpressaopot").Range.AutoFilter Field:=6, Criteria1:="<>" & "45K", Operator:=xlFilterValues  ' Adiciona um filtro para excluir linhas onde a capacidade é "45k"
    Impressao1.ListObjects("tbImpressaopot").Range.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Pesquisa.pdf", Quality:=xlQualityStandard, _
                                                                       IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Impressao1.Visible = xlSheetHidden
    Application.EnableEvents = True
    Application.EnableEvents = True

    ' Remover o filtro após a impressão
    Impressao1.ListObjects("tbImpressaopot").AutoFilter.ShowAllData

End Sub

Public Sub ativaPesquisa()
    Pesquisa.Activate
End Sub

Public Sub Imprformcampotroca()
    On Error GoTo TError

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Impressaotroca.ListObjects("tbImpressaotroca").Range.PrintPreview
    'filtraimpressao
    'filtraformpes
    filtraformtroca
    Impressaotroca.Visible = xlSheetVisible
    Impressaotroca.ListObjects("tbImpressaotroca").Range.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Formulario_para_Reposição.pdf", Quality:=xlQualityStandard, _
                                                                             IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Impressaotroca.Visible = xlSheetHidden
    Pesquisa.Activate
    Pesquisa.Range("K3").Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True

fim:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
     
    GoTo fim
End Sub


'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub filtraformtroca()
    ' Data......: 26/11/2020
    ' Descricao.: Gera Form de Reposição
    '---------------------------------------------------------------------------------------
    'Sub filtraformtroca()
    'Dim ultlinhaP                          As Long
    'Dim ultlinhaI                          As Long
    'Dim rngpesquisa                                As Range
    'Dim tblpesquisa                                As ListObject
    'Dim rng                                As Range
    'Dim tbl                                As ListObject
    'Pesquisa.Unprotect "brigada"
    ''ultlinhaP = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    'ultlinhaP = Pesquisa.Cells(Pesquisa.Rows.Count, "G").End(xlUp).Row
    'Set tblpesquisa = Pesquisa.ListObjects("tbPesquisaMapaAtual")
    'Set rngpesquisa = Range("tbPesquisaMapaAtual[#All]").Resize(ultlinhaP - 10, 23)
    'tblpesquisa.Resize rngpesquisa
    ''    Pesquisa.Range("E3:W1048576").ClearContents
    'ultlinhaP = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    'Set rngpesquisa = Range("tbPesquisaMapaAtual[#All]").Resize(ultlinhaP - 10, 23)
    'tblpesquisa.Resize rngpesquisa
    '
    '    Application.CutCopyMode = False
    '    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
    '    xlFilterCopy, CopyToRange:=Range _
    '    ("tbImpressaotroca[[#Headers],[Sup]:[Série]]"), Unique:=False
    '    Sheets("Pesquisa").Range("tbPesquisaMapaAtual[#All]").AdvancedFilter Action:= _
    '    xlFilterCopy, CopyToRange:=Range _
    '    ("tbImpressaotroca[[#Headers],[Observação]:[Observação]]"), Unique:=False
    '    ultlinhaP = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    '    ultlinhaI = Pesquisa.Range("tbPesquisaMapaAtual[[#Headers],[Série]]").End(xlDown).Row
    '
    '    Set tbl = Impressaotroca.ListObjects("tbImpressaotroca")
    '
    '
    '    Set rng = Range("tbImpressaotroca[#All]").Resize(ultlinhaI - 10, 19)
    '    tbl.Resize rng
    '    tbl.DataBodyRange.Rows.RowHeight = 40
    '    Impressaotroca.Range("E12:E" & ultlinhaP).Columns.AutoFit
    '    Impressaotroca.Range("G12:G" & ultlinhaP).Columns.AutoFit
    '    Impressaotroca.Range("H12:H" & ultlinhaP).Columns.AutoFit
    '
    'End Sub

    '---------------------------------------------------------------------------------------
    ' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub Imprformcampopesagem()
    ' Data......: 06/01/2021
    ' Descricao.: Gera formulário de campo - pesagem para impressão filtrado através da tela Pesquisa
    '---------------------------------------------------------------------------------------
Public Sub Imprformcampopesagem()
    On Error GoTo TError


    Application.EnableEvents = False
    Application.ScreenUpdating = False
    'Impressaopes.ListObjects("tbImpressaopes").Range.PrintPreview
    filtraformpes
    Impressaopes.Visible = xlSheetVisible
    Impressaopes.ListObjects("tbImpressaopes").Range.ExportAsFixedFormat _
    Type:=xlTypePDF, Filename:="Formulario_para_Pesagem.pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Impressaopes.Visible = xlSheetHidden
    Pesquisa.Activate
    Pesquisa.Range("K3").Activate
    Application.EnableEvents = True
    
    Application.ScreenUpdating = True


fim:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub Imprforminconformidades()
    On Error GoTo TError


    Application.EnableEvents = False
    Application.ScreenUpdating = False
    'Impressaopes.ListObjects("tbImpressaopes").Range.PrintPreview
    filtraformpes
    Impressaopes.Visible = xlSheetVisible
    Impressaopes.ListObjects("tbImpressaopes").Range.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Formulario_de_Inconformidades.pdf", Quality:=xlQualityStandard, _
                                                                         IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Impressaopes.Visible = xlSheetHidden
    Pesquisa.Activate
    Pesquisa.Range("K3").Activate
    Application.EnableEvents = True
    
    Application.ScreenUpdating = True


fim:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub




