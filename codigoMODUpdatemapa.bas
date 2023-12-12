Attribute VB_Name = "MODUpdatemapa"
'@Folder("SGES2020")
Option Explicit

Public Sub updateservmapa()
    Dim ultlinmapa As Long
    Dim linmapa As Long

    limpafiltrosmapaatual
    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    linmapa = 1

    With MapaAtual.ListObjects(1).DataBodyRange
        Do Until linmapa > ultlinmapa


            If UCase(.Cells(linmapa, 8)) = Info.Cells.Item(8, 9) Then

                '##### SE EXTINTOR FOI MODIFICADO #####

                If Info.Range("F11").Value = "Sim" Then
                    .Cells(linmapa, 8) = CStr(Info.Range("K6").Value) & CStr(Info.Range("m8").Value) & CStr(Info.Range("m10").Value) 'Série

                End If

                '######### TESTE #########

            

                .Cells(linmapa, 10) = DateAdd("yyyy", 5, Info.Range("I16").Value) 'TESTE

                If (.Cells(linmapa, 4) = "MANUTENÇÃO - BRIGADA" Or .Cells(linmapa, 4) = "MANUTENÇÃO - MAREFIRE") Then

                    .Cells(linmapa, 11) = "EM MANUTENÇÃO"

                ElseIf .Cells(linmapa, 10) = vbNullString Then

                    .Cells(linmapa, 11) = "PREENCHER DATA DE TESTE"


            

                    '            ElseIf Date - .Cells(linmapa, 10) < 30 And Date - .Cells(linmapa, 10) >= 1 Then
                ElseIf DateDiff("m", .Cells(linmapa, 10), Date) = 0 Then

                    .Cells(linmapa, 11) = "ATENÇÃO"
            
                ElseIf DateDiff("m", .Cells(linmapa, 10), Date) > 0 Then

                    .Cells(linmapa, 11) = "TESTE VENCIDO"

                Else

                    .Cells(linmapa, 11) = "TESTE EM DIA"


                End If


                '######### RECARGA #########

                '
                If Info.Range("M8").Value = "CO" Then


                    If DateAdd("yyyy", 5, Info.Range("M16").Value) > DateAdd("yyyy", 5, Info.Range("I16").Value) Then
                        .Cells(linmapa, 12) = DateAdd("yyyy", 5, Info.Range("I16").Value) 'Recarga CO
                    Else

                        .Cells(linmapa, 12) = DateAdd("yyyy", 5, Info.Range("M16").Value) 'Recarga CO


                    End If



                ElseIf Info.Range("M8").Value = "FM" Then
                    If DateAdd("yyyy", 5, Info.Range("M16").Value) > DateAdd("yyyy", 5, Info.Range("I16").Value) Then
                        .Cells(linmapa, 12) = DateAdd("yyyy", 5, Info.Range("I16").Value) 'Recarga FM
                    Else

                        .Cells(linmapa, 12) = DateAdd("yyyy", 5, Info.Range("M16").Value) 'Recarga FM


                    End If
                    
                     ElseIf Info.Range("M10").Value = "1K" Then
'                    If DateAdd("yyyy", 5, Info.Range("M16").Value) > DateAdd("yyyy", 5, Info.Range("I16").Value) Then
'                        .Cells(linmapa, 12) = DateAdd("yyyy", 5, Info.Range("I16").Value) 'Recarga 1k
'                    Else

                        .Cells(linmapa, 12) = "" 'Recarga 1k


'                    End If
                Else
                    If DateAdd("yyyy", 5, Info.Range("M16").Value) > DateAdd("yyyy", 5, Info.Range("M16").Value) Then
                        .Cells(linmapa, 12) = DateAdd("yyyy", 1, Info.Range("M16").Value)
                    Else
                        .Cells(linmapa, 12) = DateAdd("yyyy", 1, Info.Range("M16").Value) 'Recarga OUTROS
                    End If
                End If


                If (.Cells(linmapa, 9) = "BRIGADA" Or .Cells(linmapa, 9) = "MAREFIRE") And .Cells(linmapa, 12) <> vbNullString And .Cells(linmapa, 3) = "MANUTENÇÃO - BRIGADA" Then

                    .Cells(linmapa, 13) = "EM MANUTENÇÃO"

                ElseIf .Cells(linmapa, 12) = vbNullString And .Cells(linmapa, 6) <> "1K" Then

                    .Cells(linmapa, 13) = "PREENCHER DATA DE RECARGA"


            

                ElseIf DateDiff("m", .Cells(linmapa, 12), Date) = 0 And .Cells(linmapa, 6) <> "1K" Then

                    .Cells(linmapa, 13) = "ATENÇÃO"

                ElseIf DateDiff("m", .Cells(linmapa, 12), Date) > 0 And .Cells(linmapa, 6) <> "1K" Then

                    .Cells(linmapa, 13) = "RECARGA VENCIDA"
                    
                    ElseIf .Cells(linmapa, 6) = "1K" Then

                    .Cells(linmapa, 13) = "NÃO APLICÁVEL"

                Else

                    .Cells(linmapa, 13) = "RECARGA EM DIA"


                End If

                '######### PESAGEM #########

                If Info.Range("M8").Value = "CO" Then
                    .Cells(linmapa, 14) = DateAdd("m", 6, Info.Range("I18").Value)

                ElseIf .Range("M8").Value = "FM" Then
                    .Cells(linmapa, 14) = DateAdd("yyyy", 5, Info.Range("I18").Value)

                Else
                    .Cells(linmapa, 14) = vbNullString

                End If
                '######### STATUS PESAGEM CO #########
                If Info.Range("M8").Value = "CO" And (.Cells(linmapa, 9) = "BRIGADA" Or .Cells(linmapa, 9) = "MAREFIRE") And _
                                                                                                                           .Cells(linmapa, 3) = "MANUTENÇÃO" Then

                    .Cells(linmapa, 15) = "EM MANUTENÇÃO"
                

                ElseIf .Cells(linmapa, 10) = vbNullString And Info.Range("M8").Value = "CO" Then

                    .Cells(linmapa, 15) = "PREENCHER DATA DE PESAGEM"

            

                ElseIf .Cells(linmapa, 5) = "CO" And DateDiff("m", .Cells(linmapa, 14), Date) = 0 Then

                    .Cells(linmapa, 15) = "ATENÇÃO"
                ElseIf .Cells(linmapa, 5) = "CO" And DateDiff("m", .Cells(linmapa, 14), Date) > 0 Then

                    .Cells(linmapa, 15) = "PESAGEM VENCIDA"

                ElseIf .Cells(linmapa, 5) = "CO" And DateDiff("m", .Cells(linmapa, 14), Date) < 0 Then
                    .Cells(linmapa, 15) = "PESAGEM EM DIA"



                ElseIf .Cells(linmapa, 5) <> "CO" Then
                    .Cells(linmapa, 15) = "NÃO APLICÁVEL"

                End If



                '######### SELAGEM #########
                
                If Info.Range("M10").Value = "1K" Then
                .Cells(linmapa, 16) = "" 'Selo 1k
                Else
                .Cells(linmapa, 16) = DateAdd("yyyy", 1, Info.Range("M18").Value) 'Selo
            End If
                If .Cells(linmapa, 5) = "CO" And (.Cells(linmapa, 6) = "34K" Or .Cells(linmapa, 6) = "45K") Then
                    .Cells(linmapa, 17) = "NÃO APLICÁVEL" ' CILINDROS

                ElseIf (.Cells(linmapa, 9) = "BRIGADA" Or .Cells(linmapa, 9) = "MAREFIRE") And .Cells(linmapa, 16) <> vbNullString And .Cells(linmapa, 3) = "MANUTENÇÃO - BRIGADA" Then

                    .Cells(linmapa, 17) = "EM MANUTENÇÃO"
                ElseIf .Cells(linmapa, 16) = vbNullString And .Cells(linmapa, 6) <> "1K" Then

                    .Cells(linmapa, 17) = "PREENCHER DATA DE SELAGEM"



                ElseIf DateDiff("m", .Cells(linmapa, 16), Date) = 0 And .Cells(linmapa, 6) <> "1K" Then

                    .Cells(linmapa, 17) = "ATENÇÃO"

                ElseIf DateDiff("m", .Cells(linmapa, 16), Date) > 0 And .Cells(linmapa, 6) <> "1K" Then

                    .Cells(linmapa, 17) = "SELO VENCIDO"
                    
                     ElseIf .Cells(linmapa, 6) = "1K" Then

                    .Cells(linmapa, 17) = "NÃO APLICÁVEL"

                Else

                    .Cells(linmapa, 17) = "SELO EM DIA"

                End If


                '######### INSPEÇÃO #########



                If Info.Range("M8").Value = "CO" Then

                    .Cells(linmapa, 18) = DateAdd("m", 6, Info.Range("I20").Value) 'Inspeção CO

                ElseIf Info.Range("M8").Value = "FM" Then
                    .Cells(linmapa, 18) = DateAdd("m", 1, Info.Range("I20").Value) 'Inspeção FM


                Else
                    .Cells(linmapa, 18) = DateAdd("yyyy", 1, Info.Range("I20").Value) 'Inspeção


                End If

                If (.Cells(linmapa, 9) = "BRIGADA" Or .Cells(linmapa, 9) = "MAREFIRE") And .Cells(linmapa, 18) <> vbNullString And .Cells(linmapa, 3) = "MANUTENÇÃO - BRIGADA" Then

                    .Cells(linmapa, 19) = "EM MANUTENÇÃO"

                ElseIf .Cells(linmapa, 18) = vbNullString Then

                    .Cells(linmapa, 19) = "PREENCHER DATA DE INSPEÇÃO"

           

                ElseIf DateDiff("m", .Cells(linmapa, 18), Date) = 0 Then

                    .Cells(linmapa, 19) = "ATENÇÃO"
            
                ElseIf DateDiff("m", .Cells(linmapa, 18), Date) = 0 Then

                    .Cells(linmapa, 19) = "INSPEÇÃO VENCIDA"

                Else

                    .Cells(linmapa, 19) = "INSPEÇÃO EM DIA"

                End If

                '######### PINTURA #########



                .Cells(linmapa, 20) = DateAdd("yyyy", 5, Info.Range("I16").Value) 'Pintura
            
                '######### STATUS GERAL #######
            
                If InStr(.Cells(linmapa, 11), "VENCID") > 0 Or InStr(.Cells(linmapa, 13), "VENCID") > 0 _
                                                                                                    Or InStr(.Cells(linmapa, 15), "VENCID") > 0 Or InStr(.Cells(linmapa, 17), "VENCID") > 0 Or InStr(.Cells(linmapa, 19), "VENCID") > 0 Then

                    .Cells(linmapa, 23) = "Vencido"
                ElseIf InStr(.Cells(linmapa, 11), "SUBS") > 0 Or InStr(.Cells(linmapa, 13), "SUBS") > 0 _
                                                                                                    Or InStr(.Cells(linmapa, 15), "SUBS") > 0 Or InStr(.Cells(linmapa, 17), "SUBS") > 0 Or InStr(.Cells(linmapa, 19), "SUBS") > 0 Then

                    .Cells(linmapa, 23) = "Substituir"

                ElseIf InStr(.Cells(linmapa, 11), "ATEN") > 0 Or InStr(.Cells(linmapa, 13), "ATEN") > 0 _
                                                                                                    Or InStr(.Cells(linmapa, 15), "ATEN") > 0 Or InStr(.Cells(linmapa, 17), "ATEN") > 0 Or InStr(.Cells(linmapa, 19), "ATEN") > 0 Then

                    .Cells(linmapa, 23) = "Vencendo"

                ElseIf InStr(.Cells(linmapa, 11), "DIA") > 0 Or InStr(.Cells(linmapa, 13), "DIA") > 0 _
                                                                                                  Or InStr(.Cells(linmapa, 15), "DIA") > 0 Or InStr(.Cells(linmapa, 17), "DIA") > 0 Or InStr(.Cells(linmapa, 19), "DIA") > 0 Then

                    .Cells(linmapa, 23) = "Em dia"

                Else

                    .Cells(linmapa, 23) = "Conferir"

                End If
            
            
                Exit Do
            End If

            linmapa = linmapa + 1
        Loop
        
        'UPDATESTATUSGERAL
    End With



End Sub


Public Sub UPDATESTATUSGERAL()
    Dim ultlinmapa As Long
    Dim linmapa As Long

    limpafiltrosmapaatual
    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    linmapa = 1

    With MapaAtual.ListObjects(1).DataBodyRange
        Do Until linmapa > ultlinmapa

            If InStr(.Cells(linmapa, 11), "VENCID") > 0 Or InStr(.Cells(linmapa, 13), "VENCID") > 0 _
                                                                                                Or InStr(.Cells(linmapa, 15), "VENCID") > 0 Or InStr(.Cells(linmapa, 17), "VENCID") > 0 Or InStr(.Cells(linmapa, 19), "VENCID") > 0 Then

                .Cells(linmapa, 23) = "Vencido"
            ElseIf InStr(.Cells(linmapa, 11), "SUBS") > 0 Or InStr(.Cells(linmapa, 13), "SUBS") > 0 _
                                                                                                Or InStr(.Cells(linmapa, 15), "SUBS") > 0 Or InStr(.Cells(linmapa, 17), "SUBS") > 0 Or InStr(.Cells(linmapa, 19), "SUBS") > 0 Then

                .Cells(linmapa, 23) = "Substituir"

            ElseIf InStr(.Cells(linmapa, 11), "ATEN") > 0 Or InStr(.Cells(linmapa, 13), "ATEN") > 0 _
                                                                                                Or InStr(.Cells(linmapa, 15), "ATEN") > 0 Or InStr(.Cells(linmapa, 17), "ATEN") > 0 Or InStr(.Cells(linmapa, 19), "ATEN") > 0 Then

                .Cells(linmapa, 23) = "Vencendo"

            ElseIf InStr(.Cells(linmapa, 11), "DIA") > 0 Or InStr(.Cells(linmapa, 13), "DIA") > 0 _
                                                                                              Or InStr(.Cells(linmapa, 15), "DIA") > 0 Or InStr(.Cells(linmapa, 17), "DIA") > 0 Or InStr(.Cells(linmapa, 19), "DIA") > 0 Then

                .Cells(linmapa, 23) = "Em dia"

            Else

                .Cells(linmapa, 23) = "Conferir"

            End If
            linmapa = linmapa + 1
        Loop
    End With

End Sub


Public Sub UPDATETBLSERV()
    Dim ultlinSERV As Long


    limpafiltrosservico
    ultlinSERV = Serviços.ListObjects(1).DataBodyRange.Rows.Count + 1


    With Serviços.ListObjects(1).DataBodyRange
        .Cells(ultlinSERV, 1) = Now    'Data
        .Cells(ultlinSERV, 3) = Info.Range("M8").Value 'Tipo


        If Info.Range("F11").Value = "Sim" And Info.Range("I18").Value <> vbNullString Then
            .Cells(ultlinSERV, 2) = UCase$(CStr(Info.Range("K6").Value)) & UCase$(CStr(Info.Range("m8").Value)) & UCase$(CStr(Info.Range("m10").Value)) 'Série
        Else
            .Cells(ultlinSERV, 2) = UCase$(CStr(Info.Range("I8").Value)) 'Série
        End If



        If Info.Range("F13").Value = "Sim" Then
            .Cells(ultlinSERV, 4) = Info.Range("I16").Value 'Teste
        End If

        If Info.Range("F14").Value = "Sim" Then
        
        If Info.Range("M10").Value = "1K" Then
        .Cells(ultlinSERV, 6) = ""
        
            ElseIf Info.Range("M8").Value = "FM" Then
                .Cells(ultlinSERV, 6) = Info.Range("I16").Value 'Sugerido pelo Sup. David Honório em Reunião de trabalho
            Else
                .Cells(ultlinSERV, 6) = Info.Range("M16").Value 'Recarga
            End If
        End If

        If Info.Range("F15").Value = "Sim" Then
            .Cells(ultlinSERV, 8) = Info.Range("I18").Value 'Pesagem
        End If

        If Info.Range("F16").Value = "Sim" Then 'And (Info.Range("$M$8").Value = "CO" And (Info.Range("$M$10").Value <> "34K" And Info.Range("$M$10").Value <> "45K")) Then
            .Cells(ultlinSERV, 10) = Info.Range("M18").Value 'Selo normal
        End If

        If Info.Range("F17").Value = "Sim" Then
            .Cells(ultlinSERV, 12) = Info.Range("I20").Value 'Inspeção
        End If

        If Info.Range("F18").Value = "Sim" Then
            .Cells(ultlinSERV, 14) = Info.Range("M20").Value 'Pintura
        End If

        If Info.Range("$F$20").Value = "Sim" Then
            PreviServ
            Serviços.ListObjects("tbHistServ13").DataBodyRange.Calculate
            Info.ListObjects("tbHistServ").DataBodyRange.Calculate
        End If

    End With

End Sub



Public Sub updatetblext()
    Dim cell  As Range
    Dim ultlinSERV As Long
    Dim LINEXTINTOR As Long
    Dim ultlinext As Long
    Dim ultlinmapa As Long
    Dim linmapa As Long
    Dim SERIEADPT As String
    Dim SERIEANTIGO As String
    ultlinext = Extintores.ListObjects(1).DataBodyRange.Rows.Count

    limpafiltrosext
    With Extintores.ListObjects(1).DataBodyRange
        LINEXTINTOR = 1
        Do Until LINEXTINTOR > ultlinext

            If .Cells(LINEXTINTOR, 9) = UCase(Info.Cells.Item(8, 9)) Then
                .Cells(LINEXTINTOR, 1).Value = UCase$(Info.Range("K6").Value) 'serie
                .Cells(LINEXTINTOR, 4).Value = Info.Range("I10").Value 'Fabricação
                .Cells(LINEXTINTOR, 2).Value = Info.Range("M8").Value 'tipo
                .Cells(LINEXTINTOR, 3).Value = Info.Range("M10").Value 'capacidade
                .Cells(LINEXTINTOR, 5).Value = Info.Range("I12").Value 'suporte

                .Cells(LINEXTINTOR, 8).Value = Now 'Data de atualizaçao
                .ListObject.ListRows(LINEXTINTOR).Range.Calculate
                SERIEADPT = .Cells(LINEXTINTOR, 9).Value 'ARMAZENA SERIE MODIFICADO
                Exit Do
            End If
            LINEXTINTOR = LINEXTINTOR + 1

        Loop
        limpafiltrosmapaatual
        
        '        Dim sht As Worksheet
        'Dim fnd As Variant
        'Dim rplc As Variant

        SERIEANTIGO = Info.Cells.Item(8, 9)


        For Each cell In Movimentacao.ListObjects(1).DataBodyRange.Columns(2) 'substitui numero de serie antigo pelo numero de serie modificado na planilha movimentacao
            Movimentacao.ListObjects(1).DataBodyRange.Cells.Replace what:=SERIEANTIGO, Replacement:=SERIEADPT, _
                                                                    lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                                                                    SearchFormat:=False, ReplaceFormat:=False
        Next cell
       
        For Each cell In Serviços.ListObjects(1).DataBodyRange.Columns(2)
            Serviços.ListObjects(1).DataBodyRange.Cells.Replace what:=SERIEANTIGO, Replacement:=SERIEADPT, _
                                                                lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                                                                SearchFormat:=False, ReplaceFormat:=False
        Next cell

        ultlinSERV = Serviços.ListObjects(1).DataBodyRange.Rows.Count

        limpafiltrosservico
        With Serviços.ListObjects(1).DataBodyRange
            LINEXTINTOR = 1
            Do Until LINEXTINTOR > ultlinSERV

                If .Cells(LINEXTINTOR, 2) = SERIEADPT Then
                    '                .Cells(LINEXTINTOR, 1).Value = UCase(Info.Range("K6").Value) 'serie
                    '                .Cells(LINEXTINTOR, 4).Value = Info.Range("I10").Value 'Fabricação
                    .Cells(LINEXTINTOR, 3).Value = Info.Range("M8").Value 'tipo
                    '                .Cells(LINEXTINTOR, 3).Value = Info.Range("M10").Value 'capacidade
                    '                .Cells(LINEXTINTOR, 5).Value = Info.Range("I12").Value 'suporte
                    '
                    '                .Cells(LINEXTINTOR, 8).Value =
                
                    '                SERIEADPT = .Cells(LINEXTINTOR, 9).Value 'ARMAZENA SERIE MODIFICADO
                    'Exit Do
                End If
                LINEXTINTOR = LINEXTINTOR + 1

            Loop
            .Calculate
        End With



        
        ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
        linmapa = 1
        Do Until linmapa > ultlinmapa
        
            With MapaAtual.ListObjects(1).DataBodyRange

                 If UCase(.Cells(linmapa, 8)) = Info.Cells.Item(8, 9) Then
                    '.Cells(linmapa, 8) = SERIEAPT
                    .Cells(linmapa, 7) = Info.Range("I10").Value 'Fabricação

                    .Cells(linmapa, 5) = Info.Range("M8").Value 'tipo
                    .Cells(linmapa, 6) = Info.Range("M10").Value 'capacidade
                    .Cells(linmapa, 1) = Info.Range("I12").Value 'suporte
                
                    For Each cell In MapaAtual.ListObjects(1).DataBodyRange.Columns(2)
                        MapaAtual.ListObjects(1).DataBodyRange.Cells.Replace what:=SERIEANTIGO, Replacement:=SERIEADPT, _
                                                                             lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                                                                             SearchFormat:=False, ReplaceFormat:=False
                    Next cell

                
                    Info.Cells.Item(8, 9) = SERIEADPT
                    Exit Do
                End If
                linmapa = linmapa + 1
            End With
        Loop

    End With
End Sub




