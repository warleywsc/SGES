Attribute VB_Name = "MODContvencido"
Option Explicit




'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub contvencido()
    ' Data......: 12/01/2021
    ' Descricao.: Insere status geral do extintor no Mapa Atual
    '---------------------------------------------------------------------------------------
Public Sub contvencido()
    On Error GoTo TError
  
    Dim formula As String
    Dim arr() As Variant
    Dim i     As Long
    Dim largura As Long
    Dim percentual As Double



    arr = MapaAtual.Range("tbMapaAtual[[Local]:[STATUS GERAL]]")
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width
        For i = LBound(arr, 1) To UBound(arr, 1)
    
            DoEvents
            percentual = i / UBound(arr, 1)
            .lblBarraEvolucao.TextAlign = fmTextAlignRight
            .lblBarraEvolucao.Caption = "Atualizando Status Geral..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"

            If InStr(arr(i, 1), "BUS") > 0 Or InStr(arr(i, 1), "EMPIL") Or InStr(arr(i, 1), "TRAT") > 0 Then

                arr(i, 19) = "Veículo"
            Else
                arr(i, 19) = "Habitação"

            End If
            If InStr(arr(i, 1), "MANUTENÇÃO - BRIGADA") > 0 Or InStr(arr(i, 1), "MANUTENÇÃO - MAREFIRE") > 0 Then

                arr(i, 20) = "Em Manutenção"
            ElseIf InStr(arr(i, 8), "VENCID") > 0 Or InStr(arr(i, 10), "VENCID") > 0 _
                                                                                 Or InStr(arr(i, 12), "VENCID") > 0 Or InStr(arr(i, 14), "VENCID") > 0 Or InStr(arr(i, 16), "VENCID") > 0 Then

                arr(i, 20) = "Vencido"
            ElseIf InStr(arr(i, 8), "SUBS") > 0 Or InStr(arr(i, 10), "SUBS") > 0 _
                                                                             Or InStr(arr(i, 12), "SUBS") > 0 Or InStr(arr(i, 14), "SUBS") > 0 Or InStr(arr(i, 16), "SUBS") > 0 Then

                arr(i, 20) = "Vencido"

            ElseIf InStr(arr(i, 8), "ATEN") > 0 Or InStr(arr(i, 10), "ATEN") > 0 _
                                                                             Or InStr(arr(i, 12), "ATEN") > 0 Or InStr(arr(i, 14), "ATEN") > 0 Or InStr(arr(i, 16), "ATEN") > 0 Then

                arr(i, 20) = "Vencendo"

            ElseIf InStr(arr(i, 8), "DIA") > 0 Or InStr(arr(i, 10), "DIA") > 0 _
                                                                           Or InStr(arr(i, 12), "DIA") > 0 Or InStr(arr(i, 14), "DIA") > 0 Or InStr(arr(i, 16), "DIA") > 0 Then

                arr(i, 20) = "Em dia"

            Else

                arr(i, 20) = "Conferir"

            End If
        Next i
    End With
    If i >= UBound(arr) Then Unload frmEvolucao
    MapaAtual.Range("tbMapaAtual[[Local]:[STATUS GERAL]]") = arr

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub




