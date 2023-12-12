Attribute VB_Name = "MODPreviServ"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Rotina: Sub PreviServ()
    ' Data......: 16/11/2020
    ' Descricao.: calcula a data do proximo servico e insere nas colunas correspondentes
    '---------------------------------------------------------------------------------------
Public Sub PreviServ()


    Dim arr   As Variant
    Dim i     As Long
    Dim rng   As Range
    
    Dim largura As Long
    Dim percentual As Double
    'arr = Serviços.Range("tbServicos").CurrentRegion
    arr = Serviços.ListObjects(1).DataBodyRange
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width
        For i = LBound(arr, 1) To UBound(arr, 1)
            DoEvents
            percentual = i / UBound(arr, 1)
            .lblBarraEvolucao.Caption = "Atualizando previsões..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"
            
'            If i = 15122 Then
'
'Stop
'
'End If
'If InStr(arr(i, 2), "1K") > 0 Then Stop
   If InStr(arr(i, 2), "1K") > 0 Then
                    arr(i, 10) = ""
                    arr(i, 6) = ""
                    End If
            If arr(i, 4) <> vbNullString Then

                arr(i, 5) = DateAdd("yyyy", 5, arr(i, 4)) 'TESTE


            End If


            If arr(i, 6) <> vbNullString Then

'If InStr(arr(i, 2), "1K") > 0 Then Stop
                If arr(i, 3) = "CO" Then
            
                    '            If DateAdd("yyyy", 5, Info.Range("M16")) > arr(i, 5) Then 'compara recarga com teste. Se forma maior, iguala a teste.
                    '                arr(i, 8) = arr(i, 6)
                    '                Else
                    arr(i, 7) = DateAdd("yyyy", 5, arr(i, 6)) 'RECARGA CO
                    '                End If
                ElseIf arr(i, 3) = "FM" Then 'RECARGA FM
                    arr(i, 7) = DateAdd("yyyy", 5, arr(i, 4))
                    ElseIf InStr(arr(i, 2), "1K") > 0 Then
                    arr(i, 7) = ""
                Else
                    'If DateAdd("yyyy", 1, Info.Range("M16")) > arr(i, 5) Then 'compara recarga com teste. Se forma maior, iguala a teste.
                    'arr(i, 8) = arr(i, 6)
                    'Else
                    arr(i, 7) = DateAdd("yyyy", 1, arr(i, 6)) 'RECARGA OUTROS
                    ' End If
                End If
            End If

            If arr(i, 8) <> vbNullString Then

                arr(i, 9) = DateAdd("m", 6, arr(i, 8)) ' PESAGEM



            End If

             If InStr(arr(i, 2), "1K") > 0 Then
                    arr(i, 11) = ""
                    End If

            If arr(i, 10) <> vbNullString Then

                arr(i, 11) = DateAdd("yyyy", 1, arr(i, 10)) 'SELO
               


            End If

            If arr(i, 12) <> vbNullString Then

                If arr(i, 3) = "CO" Then

                    arr(i, 13) = DateAdd("m", 6, arr(i, 12)) 'INSPECAO CO

                ElseIf arr(i, 3) = "FM" Then

                    arr(i, 13) = DateAdd("m", 1, arr(i, 12)) 'INSPECAO FM
                Else
                    arr(i, 13) = DateAdd("yyyy", 1, arr(i, 12)) 'INSPECAO OUTROS

                End If
            End If

            If arr(i, 4) <> vbNullString Then

                arr(i, 15) = DateAdd("yyyy", 5, arr(i, 4)) 'PINTURA

            End If

            '        If arr(i, 15) <> vbNullString Then
            '
            '            arr(i, 16) = DateAdd("yyyy", 5, arr(i, 15))
            '
            '        End If
        Next i

        With Serviços

            Serviços.ListObjects(1).DataBodyRange = arr

            '        .Range("tbServicos[Próximo Teste],tbServicos[Próxima Recarga],tbServicos[Próxima Pesagem],tbServicos[Próxima Selagem],tbServicos[Próxima Inspeção],tbServicos[Próxima Pintura]").ClearContents
            '        .Range("K8:k" & UBound(arr, 1) + 7).Value = Application.Index(arr, , 6)
            '        .Range("m8:m" & UBound(arr, 1) + 7).Value = Application.Index(arr, , 8)
            '        .Range("o8:o" & UBound(arr, 1) + 7).Value = Application.Index(arr, , 10)
            '        .Range("q8:q" & UBound(arr, 1) + 7).Value = Application.Index(arr, , 12)
            '        .Range("s8:s" & UBound(arr, 1) + 7).Value = Application.Index(arr, , 14)
            '        .Range("u8:u" & UBound(arr, 1) + 7).Value = Application.Index(arr, , 16)
            '    formata as colunas para data
            '        For i = 11 To 21 Step 2
            '            Set rng = .Cells(9, i)
            '
            '            .Range(rng, .Cells(.Rows.Count, rng.Column).End(xlUp)).TextToColumns Destination:=Range(rng.Address), DataType:=xlDelimited, _
            '            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            '            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            '            :=Array(1, 4), TrailingMinusNumbers:=True
            '
            '        Next i


        End With
    End With
    
    Unload frmEvolucao
    Set rng = Nothing
    Set arr = Nothing
End Sub




