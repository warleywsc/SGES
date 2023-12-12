Attribute VB_Name = "MODAtualizamapaatual"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub atualizamapaatual()
    ' Data......: 04/12/2020
    ' Descricao.: Copia dados de mapa atual para mapa antigo e atualiza Mapa Atual
    '---------------------------------------------------------------------------------------
Public Sub atualizamapaatual()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    MapaAtual.ListObjects(1).DataBodyRange.RemoveDuplicates Columns:=8, Header:=xlYes
    MovReservaManutLote
    AtualizamapaMOV
    Atualizamapaserv
    statusservico
    AtualizamapaExt
    contvencido
    PreviServ
    Application.Speech.Speak "Atualização concluída!", speakasync:=True
    MsgBox "Atualização Concluida!", vbOKOnly, "SGES"
    Serviços.Calculate
    Info.Calculate
    'If Info.Range("i8").Rows.Hidden = False Then populafrmAtualExt
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
End Sub
