Attribute VB_Name = "MODretornaextdestino"
Option Explicit

Public Sub retornaextdestino()

    Dim ultlinmapa As Long
    Dim linmapa As Long
    Dim seriepermuta As String
    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    With MapaAtual.ListObjects(1).DataBodyRange
        linmapa = 1
        Do Until linmapa > ultlinmapa  'busca status geral do extintor

            If Info.Cells(12, 13) & " " & Info.Cells(14, 9) = .Cells(linmapa, 4) & " " & .Cells(linmapa, 2) _
        And (.Cells(linmapa, 4) <> "RESERVA TÉCNICA" Or .Cells(linmapa, 4) <> "MANUTENÇÃO - BRIGADA" Or _
             .Cells(linmapa, 4) <> "MANUTENÇÃO - MAREFIRE") Then


                seriepermuta = .Cells(linmapa, 8)
                Info.Cells(6, 14) = seriepermuta
                Exit Do
            End If
            linmapa = linmapa + 1
        Loop

    End With
    Info.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub




