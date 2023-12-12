Attribute VB_Name = "MODexcluiservmapaatual"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub excluiservmapaatual()
    ' Data......: 15/01/2021
    ' Descricao.: Exclui serviços através do menu de contexto em Info
    '---------------------------------------------------------------------------------------
Public Sub excluiservmapaatual()
    Dim LINMAPAATUAL As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Info

        LINMAPAATUAL = 8
        Do Until MapaAtual.Cells(LINMAPAATUAL, 14).Row > MapaAtual.Cells(MapaAtual.Rows.Count, "G").End(xlUp).Row

            If MapaAtual.Cells(LINMAPAATUAL, 14) = .Cells.Item(8, 9) Then

                MapaAtual.Range("P" & LINMAPAATUAL).Value = vbNullString

                MapaAtual.Range("R" & LINMAPAATUAL).Value = vbNullString
                MapaAtual.Range("T" & LINMAPAATUAL).Value = vbNullString
                MapaAtual.Range("V" & LINMAPAATUAL).Value = vbNullString
                MapaAtual.Range("X" & LINMAPAATUAL).Value = vbNullString
                MapaAtual.Range("Z" & LINMAPAATUAL).Value = vbNullString
                populafrmAtualExt
                ' MapaAtual.Range("AA" & LINMAPAATUAL).Value = .Range("G23").Value 'Observação
                
                
                MsgBox "Serviço Excluido!"


                Exit Do
            End If
            LINMAPAATUAL = LINMAPAATUAL + 1
        Loop
        UPDATESTATUSGERAL
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub excluiservIndividualmapaatual()
    ' Data......: 15/01/2021
    ' Descricao.: Exclui serviços "individualmente" através do menu de contexto em Info
    '---------------------------------------------------------------------------------------
Public Sub excluiservIndividualmapaatual()
    On Error GoTo TError
    Dim LINMAPAATUAL As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Info

        LINMAPAATUAL = 8
        Do Until MapaAtual.Cells(LINMAPAATUAL, 14).Row > MapaAtual.Cells(MapaAtual.Rows.Count, "G").End(xlUp).Row

            If MapaAtual.Cells(LINMAPAATUAL, 14) = .Cells.Item(8, 9) Then

                If ActiveCell.Address = "$I$16" Then


                    MapaAtual.Range("P" & LINMAPAATUAL).Value = vbNullString 'teste

                ElseIf ActiveCell.Address = "$M$16" Then

                    MapaAtual.Range("R" & LINMAPAATUAL).Value = vbNullString 'recarga

                ElseIf ActiveCell.Address = "$I$18" Then


                    MapaAtual.Range("T" & LINMAPAATUAL).Value = vbNullString 'PESAGEM

                ElseIf ActiveCell.Address = "$M$18" Then


                    MapaAtual.Range("V" & LINMAPAATUAL).Value = vbNullString 'SELO

                ElseIf ActiveCell.Address = "$I$20" Then
                    MapaAtual.Range("X" & LINMAPAATUAL).Value = vbNullString 'INSPECAO

                ElseIf ActiveCell.Address = "$M$20" Then
                    MapaAtual.Range("Z" & LINMAPAATUAL).Value = vbNullString ' PINTURA

                End If
                populafrmAtualExt
                
                MsgBox "Serviço Excluido!"

                Exit Do
            End If
            LINMAPAATUAL = LINMAPAATUAL + 1
        Loop

    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

