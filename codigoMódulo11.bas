Attribute VB_Name = "M�dulo11"
Public Sub aTUALIZARSERVICOS()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    MapaAtual.ListObjects(1).DataBodyRange.RemoveDuplicates Columns:=8, Header:=xlYes
'    MovReservaManutLote
'    AtualizamapaMOV
    Atualizamapaserv
    statusservico
    AtualizamapaExt
'    contvencido
    PreviServ
    Application.Speech.Speak "Atualiza��o conclu�da!", speakasync:=True
    MsgBox "Atualiza��o Concluida!", vbOKOnly, "SGES"
    Servi�os.Calculate
    Info.Calculate
    'If Info.Range("i8").Rows.Hidden = False Then populafrmAtualExt
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
End Sub
