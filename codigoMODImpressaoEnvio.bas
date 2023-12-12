Attribute VB_Name = "MODImpressaoEnvio"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub ImprimeEnvio()
    ' Data......: 20/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub ImprimeEnvio()
    On Error GoTo TError

    Dim areaimpressao As Range
    '    Application.EnableEvents = False
    '    Application.ScreenUpdating = False
    
    Set areaimpressao = formenvio.Range("G8:O" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)
    formenvio.Visible = xlSheetVisible
    formenvio.ResetAllPageBreaks
    areaimpressao.CurrentRegion.Columns.HorizontalAlignment = xlCenter
    With formenvio.PageSetup
    
        .PrintArea = areaimpressao.Address
        .FitToPagesWide = 1
        '    .FitToPagesTall = 0
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
    

    areaimpressao.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Formulario de Envio " & Format$(Date, "dd-mm-yyyy") & ".pdf", Quality:=xlQualityStandard, _
                                                                                                                      IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    '    formenvio.Visible = xlSheetHidden
    Set areaimpressao = Nothing
    '    Application.EnableEvents = True
    '    Application.EnableEvents = True



fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "ImprimeEnvio()"
    GoTo fim
End Sub



Public Sub ImprimeRetorno()
    On Error GoTo TError

    Dim areaimpressao As Range
    '    Application.EnableEvents = False
    '    Application.ScreenUpdating = False
    
    Set areaimpressao = formenvio.Range("AO8:AW" & formenvio.Cells(Rows.Count, "AO").End(xlUp).Row)
    formenvio.Visible = xlSheetVisible
    formenvio.ResetAllPageBreaks
    areaimpressao.CurrentRegion.Columns.HorizontalAlignment = xlCenter
    With formenvio.PageSetup
    
        .PrintArea = areaimpressao.Address
        .FitToPagesWide = 1
        '    .FitToPagesTall = 0
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
    

    areaimpressao.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Formulario de Envio " & Format$(Date, "dd-mm-yyyy") & ".pdf", Quality:=xlQualityStandard, _
                                                                                                                      IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    '    formenvio.Visible = xlSheetHidden
    Set areaimpressao = Nothing
    '    Application.EnableEvents = True
    '    Application.EnableEvents = True



fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "ImprimeEnvio()"
    GoTo fim
End Sub


