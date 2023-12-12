Attribute VB_Name = "MODRibbon"
'@Folder("SGES2020")
'Public Sub salvarcomo()
'
'
'    'Updated by Extendoffice 20191223
'    Dim xWb   As Workbook
'    Dim xStr  As String
'    Dim xStrSP As String
'    Dim xStrOldName As String
'    Dim xStrDate As String
'    Dim xFileName As Variant
'    Dim xFileNameSP As Variant
'    Dim xStrPathSP As String
'    Dim xFileDlg As FileDialog
'    Dim i     As Variant
'    Application.DisplayAlerts = False
'    Set xWb = ActiveWorkbook
'    xStrPathSP = "https://eletrobrastermonuclear.sharepoint.com/sites/SGES/Documentos Compartilhados/Backup SGES/"
'    xStrOldName = ThisWorkbook.Path & "\" & xWb.Name
'    xStrSP = "SGES_MASTER"
'    xStr = Left$(xStrOldName, Len(xStrOldName) - 5)
'    xStrDate = Format$(Now, "yyyy-mm-dd hh-mm-ss")
'    checkInternetConnection 'verifica conexão com internet
'    xWb.SaveAs "https://eletrobrastermonuclear.sharepoint.com/sites/SGES/Documentos Compartilhados/Backup SGES/SGES_MASTER.xlsm", 52 ' xFileNameSP, 52
'    xWb.SaveAs "https://eletrobrastermonuclear.sharepoint.com/sites/SGES/Documentos Compartilhados/Backup SGES/" & "SGES2021" & " " & xStrDate & ".xlsm", 52 '"Excel Macro-Enabled Workbook (*.xlsm),*.xlsm"
'    MsgBox "Backup Concluido!", vbInformation, "Backup Concluído"
'    If xFileName = False Then
'        Application.DisplayAlerts = True
'        Exit Sub
'
'    Else
'
'    End If
'    Application.DisplayAlerts = True
'End Sub

Public Sub salvarcomo()


'    'Updated by Extendoffice 20191223
'    Dim xWb   As Workbook
'    Dim xStr  As String
'    Dim xStrOldName As String
'    Dim xStrDate As String
'    Dim xFileName As Variant
'    Dim xFileDlg As FileDialog
'    Dim i     As Variant
'    Application.DisplayAlerts = False
'    Set xWb = ActiveWorkbook
'    xStrOldName = ThisWorkbook.Path & "\" & xWb.Name
'    xStr = Left$(xStrOldName, Len(xStrOldName) - 5)
'    xStrDate = Format$(Now, "yyyy-mm-dd hh-mm-ss")
'    If Right$(xStrOldName, 4) = "xlsm" Then
'        xFileName = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\" & "SGES2020", "Excel Macro-Enabled Workbook (*.xlsm),*.xlsm")
'    Else
'        xFileName = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\" & "SGES2020", "Excel Workbook (*.xlsx),*.xlsx")
'    End If
'    If xFileName = False Then
'        Application.DisplayAlerts = True
'        Exit Sub
'
'    Else
'        xWb.SaveAs (xFileName)
'    End If
'    Application.DisplayAlerts = True

vazio = 0
serievazio
If vazio > 0 Then

resultado = MsgBox("Você deveria preencher o número de série! Deseja Deseja salvar mesmo assim?", vbYesNo, "Número de série indefinido")
If resultado = vbYes Then
    ActiveWorkbook.Save
Else
    MsgBox "Salvamento interrompido!"
    GoTo fim:
    
End If

Else
ActiveWorkbook.Save
fim:
End If


End Sub
Sub chamaformalteralocal(control As IRibbonControl)
Attribute chamaformalteralocal.VB_ProcData.VB_Invoke_Func = "L\n14"
    chamaformEditalocal
End Sub
Public Sub chamaPesquisar(control As IRibbonControl)
    ativaPesquisa
End Sub

'Callback for btnSalvar onAction
Public Sub chamasalvarcomo(control As IRibbonControl)

    salvarcomo
End Sub

'Callback for btnSair onAction
Public Sub chamasalvardireto(control As IRibbonControl)
    salvardireto

End Sub

'Callback for btnAddExt onAction
Public Sub chamafrmNovo(control As IRibbonControl)
    SetOnkey (True)
    frmNovo
End Sub

'Callback for btnAddLocal onAction
Public Sub chamafrmLocalNovo(control As IRibbonControl)
    SetOnkey (True)
    frmLocalNovo


End Sub


'Callback for btnRemessa onAction
Sub chamaformenvio(control As IRibbonControl)
    chamaformenviomanut
End Sub

Public Sub chamaatualizaserv(control As IRibbonControl)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.Speech.Speak ("Deseja ATUALIZAR o Mapa de Extintores?"), speakasync:=True
    atualiza = MsgBox("Deseja ATUALIZAR o Mapa de Extintores?", vbOKCancel, "Atualizar Mapa")
    If atualiza = vbOK Then
    
    
        PreviServ
        AtualizamapaMOV
        AtualizamapaExt
        Atualizamapaserv
        statusservico
    
        contvencido
        '    frmAtualiza
    Else
        Exit Sub:
    End If
    Application.Speech.Speak "Atualização concluída!", speakasync:=True
    MsgBox "Atualização Concluida!", vbOKOnly, "SGES"
    'If Info.Range("i8").Rows.Hidden = False Then populafrmAtualExt
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
'Callback for btnAtualExt onAction
Public Sub chamafrmatual(control As IRibbonControl)
    frmAtualiza

End Sub
Public Sub chamaexcluiservmapaatual(control As IRibbonControl)

    Application.Speech.Speak ("Deseja EXCLUIR estes serviços?"), speakasync:=True
    atualiza = MsgBox("Deseja EXCLUIR estes serviços?", vbOKCancel, "Atualizar Mapa")
    If atualiza = vbOK Then
        excluiservmapaatual
    Else
        Exit Sub:
    End If

End Sub



Public Sub salvardireto()


    'Updated by Extendoffice 20191223
    Dim xWb   As Workbook
    Dim xStr  As String
    Dim xStrOldName As String
    Dim xStrDate As String
    Dim xFileName As Variant
    Dim xFileDlg As FileDialog
    Dim i     As Variant
    Application.DisplayAlerts = False
    Set xWb = ActiveWorkbook
    xStrOldName = xWb.Name
    xStr = Left$(xStrOldName, Len(xStrOldName) - 5)
    xStrDate = Format$(Now, "yyyy-mm-dd hh-mm-ss")
    If Right$(xStrOldName, 4) = "xlsm" Then
        xFileName = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\" & "SGES2020" & " " & xStrDate, "Excel Macro-Enabled Workbook (*.xlsm),*.xlsm")
    Else
        xFileName = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\" & "SGES2020" & " " & xStrDate, "Excel Workbook (*.xlsx),*.xlsx")
    End If
    If xFileName = False Then
        Application.DisplayAlerts = True
        Exit Sub
    Else
        xWb.SaveAs (xFileName)
    End If
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
End Sub

