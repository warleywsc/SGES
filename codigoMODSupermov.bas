Attribute VB_Name = "MODSupermov"
'@Folder("SGES2020")
Option Explicit

Public Sub SUPERMOV()
    Dim datahora As Date
    Dim resultado As VbMsgBoxResult
    Dim ultlinmapa As Long
    Dim linmapa As Long
    Dim linmapaseriereserva As Long
    Dim linmapaseriepermuta As Long
    Dim ultlinmov As Long
    Dim linmov As Long
    Dim serieorigem As String
    Dim seriereserva As String
    Dim seriepermuta As String
    Dim localorigem As String
    Dim zonaorigem As String
    Dim edforigem As String
    Dim areaorigem As String
    Dim localseriereserva As String
    Dim edfseriereserva As String
    Dim areaseriereserva As String
    Dim zonaseriereserva As String
    Dim statusseriereserva As String
    Dim localseriepermuta As String
    Dim localconcatpermuta As String
    Dim edfseriepermuta As String
    Dim areaseriepermuta As String
    Dim zonaseriepermuta As String
    Dim statusseriepermuta As String
    Dim localdest As String
    Dim zonadest As String
    Dim edfdest As String
    Dim areadest As String
    Dim statusserieorigem As String
    Dim cap   As String
    '    Application.EnableEvents = False
    '    Application.ScreenUpdating = False
    ultlinmov = Movimentacao.ListObjects(1).DataBodyRange.Rows.Count
    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    With Info
        areaorigem = .Cells(7, 1)
        zonaorigem = .Cells(8, 1)
        localdest = .Cells(12, 13)
        areadest = .Cells(14, 9)
        zonadest = .Cells(14, 13)
        serieorigem = .Cells(8, 9)
        localorigem = .Cells(6, 1)
        cap = .Cells(10, 13)
    End With
    If InStr(localorigem, " - ") > 0 Then
        edforigem = Left$(localorigem, InStr(localorigem, " - ") - 1)

    Else
        edforigem = localorigem
    End If
    If InStr(localdest, " - ") > 0 Then
        edfdest = Left$(localdest, InStr(localdest, " - ") - 1)

    Else
        edfdest = localdest
    End If
    linmapa = 1
    With MapaAtual.ListObjects(1).DataBodyRange
        Do Until linmapa > ultlinmapa  'busca status geral do extintor

            If serieorigem = UCase(.Cells(linmapa, 8)) Then

                statusserieorigem = .Cells(linmapa, 23)
                Exit Do
            End If

            linmapa = linmapa + 1
        Loop
    End With
    'MANUTENÇÃO >> RESERVA

    If Info.Cells(6, 1) = "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) = "RESERVA TÉCNICA" Then
        If statusserieorigem <> "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para a Reserva Técnica ", speakasync:=True
            MsgBox "Movimentação imprópria. Apenas Extintores 'EM DIA' poderão ser movidos para a Reserva Técnica.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest
            End With
        End If
    End If
    'MANUTENÇÃO >> TERCEIRIZADA


    If Info.Cells(6, 1) = "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) = "MANUTENÇÃO - MAREFIRE" Then
        If statusserieorigem = "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a Empresa de Manutenção", speakasync:=True
            MsgBox "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a Empresa de Manutenção.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest
            End With
        End If
    End If

    'TERCEIRIZADA >> MANUTENÇÃO

    If Info.Cells(6, 1) = "MANUTENÇÃO - MAREFIRE" And Info.Cells(12, 13) = "MANUTENÇÃO - BRIGADA" Then
        If statusserieorigem <> "Em dia" And statusserieorigem <> "Em Manutenção" Then
            Application.Speech.Speak "Movimentação imprópria. Apenas Extintores EM DIA poderão ter origem da empresa de manutenção.", speakasync:=True
            MsgBox "Movimentação imprópria. Apenas Extintores EM DIA poderão ter origem da empresa de manutenção.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest
            End With
        End If
    End If

    'RESERVA >> MANUTENÇÃO

    If Info.Cells(6, 1) = "RESERVA TÉCNICA" And Info.Cells(12, 13) = "MANUTENÇÃO - BRIGADA" Then
        If statusserieorigem = "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a Manutenção Brigada", speakasync:=True
            MsgBox "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a MANUTENÇÃO - BRIGADA.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest
            End With
        End If
    End If
    'TERCEIRIZADA >> RESERVA

    If Info.Cells(6, 1) = "MANUTENÇÃO - MAREFIRE" And Info.Cells(12, 13) = "RESERVA TÉCNICA" Then
        If statusserieorigem <> "Em dia" And statusserieorigem <> "Em Manutenção" Then
            Application.Speech.Speak "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para a RESERVA TÉCNICA ", speakasync:=True
            MsgBox "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para a RESERVA TÉCNICA.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 7) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 5) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = DateAdd("s", 4, Now)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest
            End With
        End If

    End If

    'RESERVA >> TERCEIRIZADA

    If Info.Cells(6, 1) = "RESERVA TÉCNICA" And Info.Cells(12, 13) = "MANUTENÇÃO - MAREFIRE" Then
        If statusserieorigem = "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a Empresa de Manutenção", speakasync:=True
            MsgBox "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a Empresa de Manutenção.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 7) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 5) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest
            End With
        End If
    End If
    
    
    
    'CAMPO >> RESERVA

    If Info.Cells(6, 1) <> "RESERVA TÉCNICA" And Info.Cells(6, 1) <> "MANUTENÇÃO - MAREFIRE" _
                                                                  And Info.Cells(6, 1) <> "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) = "RESERVA TÉCNICA" Then
        If statusserieorigem <> "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para a RESERVA TÉCNICA.", speakasync:=True
            MsgBox "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para a RESERVA TÉCNICA.", , "SGES"
            GoTo fim:
        Else
voltainput:
            seriereserva = UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            If seriereserva = vbNullString Then

                MsgBox "Movimentação cancelada"
                GoTo fim:

            End If

            '##### CHECAGEM EXT RESERVA  #####
            With MapaAtual.ListObjects(1).DataBodyRange
                linmapaseriereserva = 1
                Do Until linmapaseriereserva > ultlinmapa 'busca status geral do extintor

                    If seriereserva = .Cells(linmapaseriereserva, 8) Then

                        localseriereserva = .Cells(linmapaseriereserva, 4)
                        edfseriereserva = .Cells(linmapaseriereserva, 3)
                        areaseriereserva = .Cells(linmapaseriereserva, 2)
                        zonaseriereserva = .Cells(linmapaseriereserva, 9)
                        statusseriereserva = UCase$(.Cells(linmapaseriereserva, 23))
                        Exit Do
                    End If

                    linmapaseriereserva = linmapaseriereserva + 1
                Loop
                If linmapaseriereserva > ultlinmapa And localseriereserva = vbNullString Then
                    resultado = MsgBox("Extintor não encontrado! Deseja escolher outro extintor?", vbYesNo, "SGES")
                    If resultado = vbYes Then
                        GoTo voltainput:
                    Else
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                    End If
                End If
                If statusseriereserva <> "EM DIA" Then

                    Application.Speech.Speak "Este extintor não está em dia. Deseja inserir um novo extintor?", speakasync:=True
                    resultado = MsgBox("Movimentação imprópria. Apenas Extintores EM DIA poderão substituir extintores no campo. Deseja escolher outro extintor?", vbYesNo, "SGES")
                    If resultado = vbYes Then
                        GoTo voltainput:
                    Else
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                    End If
                Else

                    '                If localseriereserva <> "RESERVA TÉCNICA" Then
                    '                    Application.Speech.Speak "Este extintor não está na reserva técnica. Deseja inserir um novo extintor?", speakasync:=True
                    '                    resultado = MsgBox("Movimentação imprópria. Apenas Extintores da reserva técnica poderão substituir extintores no CAMPO.", vbYesNoCancel, "SGES")
                    '                    If resultado = vbYes Then
                    '                        GoTo voltainput:
                    '                    Else
                    '                        MsgBox "Movimentação cancelada"
                    '                        GoTo fim:
                    '                    End If
                    '
                    '                    '##### FIM CHECAGEM EXT RESERVA  #####
                    '                Else
                    .Cells(linmapaseriereserva, 4) = localorigem
                    .Cells(linmapaseriereserva, 3) = edforigem
                    .Cells(linmapaseriereserva, 2) = areaorigem
                    .Cells(linmapaseriereserva, 9) = zonaorigem
                    '                End If
                End If
            End With
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "RESERVA TÉCNICA"
                .Cells(ultlinmov, 7) = "1111"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = seriereserva
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = "RESERVA TÉCNICA"
                .Cells(ultlinmov, 5) = "1111"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                .Cells(ultlinmov, 2) = seriereserva
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localorigem
                .Cells(ultlinmov, 7) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest


            End With
        End If
    End If

    'CAMPO >> MANUTENÇÃO

    If Info.Cells(6, 1) <> "RESERVA TÉCNICA" And Info.Cells(6, 1) <> "MANUTENÇÃO - MAREFIRE" _
                                                                  And Info.Cells(6, 1) <> "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) = "MANUTENÇÃO - BRIGADA" Then
        If statusserieorigem = "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a Manutenção Brigada.", speakasync:=True
            MsgBox "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a MANUTENÇÃO - BRIGADA.", , "SGES"
            GoTo fim:
        Else
            If localorigem = "APOIO A PARADA - ANGRA 1" Or InStr(localorigem, "BANCO RESERVA") > 0 Or InStr(localorigem, "BANCO PRINCIPAL") > 0 Or localorigem = "APOIO A PARADA - 2P18" _
            Or localorigem = "COMBOIO DO TRANSPORTE DE COMBUSTÍVEL" Or localorigem = "MANUTENÇÃO - IDEAL FIRE" Then
                GoTo apoioparada2:
            End If
            If cap = "34K" Or cap = "45K" Or cap = "100K" Or cap = "250K" _
            Or cap = "37L" Or cap = "40L" Or serieorigem = "2189NL10L" Then 'cilindros co2, CILINDROS PQS CAMINHOES E CILINDROS N2 CAMINHOES
                GoTo bancoreserva1:
            End If
voltainput2:
            seriereserva = serie 'UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            
            serie = ""
            If seriereserva = vbNullString Then

                MsgBox "Movimentação cancelada"
                GoTo fim:

            End If

            '##### CHECAGEM EXT RESERVA  #####
            With MapaAtual.ListObjects(1).DataBodyRange
                linmapaseriereserva = 1
                Do Until linmapaseriereserva > ultlinmapa 'busca status geral do extintor

                    If seriereserva = .Cells(linmapaseriereserva, 8) Then

                        localseriereserva = .Cells(linmapaseriereserva, 4)
                        edfseriereserva = .Cells(linmapaseriereserva, 3)
                        areaseriereserva = .Cells(linmapaseriereserva, 2)
                        zonaseriereserva = .Cells(linmapaseriereserva, 9)
                        statusseriereserva = UCase$(.Cells(linmapaseriereserva, 23))
                        Exit Do
                    End If

                    linmapaseriereserva = linmapaseriereserva + 1
                Loop
                If statusseriereserva <> "EM DIA" And statusseriereserva <> "VENCENDO" Then

                    Application.Speech.Speak "Este extintor não está em dia. Deseja inserir um novo extintor?", speakasync:=True
                    resultado = MsgBox("Movimentação imprópria. Apenas Extintores EM DIA poderão substituir extintores no campo. Deseja escolher outro extintor?", vbYesNo, "SGES")
                    If resultado = vbYes Then
                        GoTo voltainput2:
                    Else
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                    End If
                Else
                    '
                    '                If localseriereserva <> "RESERVA TÉCNICA" Then
                    '                    Application.Speech.Speak "Este extintor não está na reserva técnica. Deseja inserir um novo extintor?", speakasync:=True
                    '                    resultado = MsgBox("Movimentação imprópria. Apenas Extintores da reserva técnica poderão substituir extintores no CAMPO.", vbYesNoCancel, "SGES")
                    '                    If resultado = vbYes Then
                    '                        GoTo voltainput:
                    '                    Else
                    '                        MsgBox "Movimentação cancelada"
                    '                        GoTo fim:
                    '                    End If
                    '
                    '                    '##### FIM CHECAGEM EXT RESERVA  #####
                    '                Else
                    .Cells(linmapaseriereserva, 4) = localorigem
                    .Cells(linmapaseriereserva, 3) = edforigem
                    .Cells(linmapaseriereserva, 2) = areaorigem
                    .Cells(linmapaseriereserva, 9) = zonaorigem
                    '                End If
                End If
            End With
apoioparada2:
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 7) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
            
                If localorigem = localseriereserva Or InStr(localorigem, "APOIO A PARADA") > 0 Or _
                localorigem = "APOIO A PARADA - 2P18" Or localorigem = "COMBOIO DO TRANSPORTE DE COMBUSTÍVEL" Or _
                localorigem = "MANUTENÇÃO - IDEAL FIRE" Or InStr(localorigem, "BANCO RESERVA") > 0 Or InStr(localorigem, "BANCO PRINCIPAL") > 0 Then
            
                    GoTo pula3:
                Else
            
            
                    ultlinmov = ultlinmov + 1
                    .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                    .Cells(ultlinmov, 2) = seriereserva
                    .Cells(ultlinmov, 3) = "Saída"
                    .Cells(ultlinmov, 4) = "RESERVA TÉCNICA"
                    .Cells(ultlinmov, 5) = "1111"
                    .Cells(ultlinmov, 8) = "BRIGADA"
                    ultlinmov = ultlinmov + 1
                    .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                    .Cells(ultlinmov, 2) = seriereserva
                    .Cells(ultlinmov, 3) = "Entrada"
                    .Cells(ultlinmov, 6) = localorigem
                    .Cells(ultlinmov, 7) = areaorigem
                    .Cells(ultlinmov, 8) = zonaorigem
                    GoTo pula3:
                End If
            End With
bancoreserva1:
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 7) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
            
                If localorigem = localseriereserva Or localorigem = "CASA DE CILINDROS - BANCO RESERVA" Or _
                localorigem = "CASA DE CILINDROS - BANCO PRINCIPAL" Or localorigem = "CAMINHÃO DE BOMBEIRO - AHQ02" _
                Or localorigem = "CAMINHÃO DE BOMBEIRO - AHQ01" Or localorigem = "CAMINHÃO DE BOMBEIRO - ABT01" Then
            
                    GoTo pula3:
                Else
            
            
                    ultlinmov = ultlinmov + 1
                    .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                    .Cells(ultlinmov, 2) = seriereserva
                    .Cells(ultlinmov, 3) = "Saída"
                    .Cells(ultlinmov, 4) = "RESERVA TÉCNICA"
                    .Cells(ultlinmov, 5) = "1111"
                    .Cells(ultlinmov, 8) = "BRIGADA"
                    ultlinmov = ultlinmov + 1
                    .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                    .Cells(ultlinmov, 2) = seriereserva
                    .Cells(ultlinmov, 3) = "Entrada"
                    .Cells(ultlinmov, 6) = localorigem
                    .Cells(ultlinmov, 7) = areaorigem
                    .Cells(ultlinmov, 8) = zonaorigem
                End If
            End With

pula3:
           
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest


            End With
        End If
    End If

    'MANUTENÇÃO >> CAMPO
    If Info.Cells(6, 1) = "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) <> "RESERVA TÉCNICA" And Info.Cells(12, 13) <> "MANUTENÇÃO - MAREFIRE" _
                                                                                                                    And Info.Cells(12, 13) <> "MANUTENÇÃO - BRIGADA" Then
        If statusserieorigem <> "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para o CAMPO ", speakasync:=True
            MsgBox "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para o CAMPO.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "RESERVA TÉCNICA"
                .Cells(ultlinmov, 7) = "1111"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = "RESERVA TÉCNICA"
                .Cells(ultlinmov, 5) = "1111"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
        End If
    End If

    'RESERVA >> CAMPO

    If Info.Cells(6, 1) = "RESERVA TÉCNICA" And Info.Cells(12, 13) <> "RESERVA TÉCNICA" And Info.Cells(12, 13) <> "MANUTENÇÃO - MAREFIRE" _
                                                                                                               And Info.Cells(12, 13) <> "MANUTENÇÃO - BRIGADA" Then
''        If UCase$(statusserieorigem) <> UCase$("Em dia") Then
'            Application.Speech.Speak "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para o CAMPO ", speakasync:=True
'            MsgBox "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para o CAMPO.", , "SGES"
'            GoTo Fim:
'        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que será substituido.", "Repondo Extintor", vbOKCancel))

            With MapaAtual.ListObjects(1).DataBodyRange


                '##### CHECAGEM EXT permuta no campo  #####

                linmapaseriepermuta = 1
                Do Until linmapaseriepermuta > ultlinmapa 'busca status geral do extintor

                    If Info.Cells(12, 13) & " - " & Info.Cells(14, 9) = .Cells(linmapaseriepermuta, 4) & " " & .Cells(linmapaseriepermuta, 2) Then


                        seriepermuta = .Cells(linmapaseriepermuta, 8)
                        localseriepermuta = .Cells(linmapaseriepermuta, 4)
                        edfseriepermuta = .Cells(linmapaseriepermuta, 3)
                        areaseriepermuta = .Cells(linmapaseriepermuta, 2)
                        zonaseriepermuta = .Cells(linmapaseriepermuta, 9)
                        statusseriepermuta = .Cells(linmapaseriepermuta, 23)
                        localconcatpermuta = localseriepermuta & " " & areaseriepermuta
                        Exit Do
                    End If

                    linmapaseriepermuta = linmapaseriepermuta + 1
                Loop
            
                'se local não estiver no mapa
            
            
                If linmapaseriepermuta > ultlinmapa And seriepermuta = vbNullString Then
                    linmapaseriepermuta = 1
                    Do Until linmapaseriepermuta > ultlinmapa
                        If serieorigem = .Cells(linmapaseriepermuta, 8) Then
                            '            .Cells(linmapaseriepermuta, 8) = serieorigem
                            .Cells(linmapaseriepermuta, 4) = localdest
                            .Cells(linmapaseriepermuta, 3) = edfdest
                            .Cells(linmapaseriepermuta, 2) = areadest
                            .Cells(linmapaseriepermuta, 9) = zonadest
                            Exit Do
                        End If
                        linmapaseriepermuta = linmapaseriepermuta + 1
                    Loop
                    updateservmapa
            
                    '                MsgBox "Local não encontrado!", , "SGES"
                    GoTo restante:
                End If
                If statusseriepermuta = "Em dia" Then
          

                    .Cells(linmapaseriepermuta, 4) = localorigem
                    .Cells(linmapaseriepermuta, 3) = edforigem
                    .Cells(linmapaseriepermuta, 2) = areaorigem
                    .Cells(linmapaseriepermuta, 9) = zonaorigem

                    With Movimentacao.ListObjects(1).DataBodyRange
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = Now
                        datahora = .Cells(ultlinmov, 1)
                        .Cells(ultlinmov, 2) = serieorigem
                        .Cells(ultlinmov, 3) = "Saída"
                        .Cells(ultlinmov, 4) = localorigem
                        .Cells(ultlinmov, 5) = areaorigem
                        .Cells(ultlinmov, 8) = zonaorigem
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                        .Cells(ultlinmov, 2) = serieorigem
                        .Cells(ultlinmov, 3) = "Entrada"
                        .Cells(ultlinmov, 6) = localdest
                        .Cells(ultlinmov, 7) = areadest
                        .Cells(ultlinmov, 8) = zonadest
                        'restante:
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = Now
                        datahora = .Cells(ultlinmov, 1)
                        .Cells(ultlinmov, 2) = seriepermuta
                        .Cells(ultlinmov, 3) = "Saída"
                        .Cells(ultlinmov, 4) = localseriepermuta
                        .Cells(ultlinmov, 5) = areaseriepermuta
                        .Cells(ultlinmov, 8) = zonaseriepermuta
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                        .Cells(ultlinmov, 2) = seriepermuta
                        .Cells(ultlinmov, 3) = "Entrada"
                        .Cells(ultlinmov, 6) = localorigem
                        .Cells(ultlinmov, 7) = areaorigem
                        .Cells(ultlinmov, 8) = zonaorigem

                    End With
                
restante:
                    With Movimentacao.ListObjects(1).DataBodyRange
                        ultlinmov = ultlinmov + 1
                    
                        .Cells(ultlinmov, 1) = Now
                        datahora = .Cells(ultlinmov, 1)
                        .Cells(ultlinmov, 2) = serieorigem
                        .Cells(ultlinmov, 3) = "Saída"
                        .Cells(ultlinmov, 4) = localorigem
                        .Cells(ultlinmov, 5) = areaorigem
                        .Cells(ultlinmov, 8) = zonaorigem
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                        .Cells(ultlinmov, 2) = serieorigem
                        .Cells(ultlinmov, 3) = "Entrada"
                        .Cells(ultlinmov, 6) = localdest
                        .Cells(ultlinmov, 7) = areadest
                        .Cells(ultlinmov, 8) = zonadest

                    End With
                
                
                    '                .Cells(linmapa, 4) = localdest
                    '                .Cells(linmapa, 3) = edfdest
                    '                .Cells(linmapa, 2) = areadest
                    '                .Cells(linmapa, 9) = zonadest
                Else


                    .Cells(linmapaseriepermuta, 4) = "MANUTENÇÃO - BRIGADA"
                    .Cells(linmapaseriepermuta, 3) = "MANUTENÇÃO"
                    .Cells(linmapaseriepermuta, 2) = "0000"
                    .Cells(linmapaseriepermuta, 9) = "BRIGADA"
                    With Movimentacao.ListObjects(1).DataBodyRange
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = Now
                        datahora = .Cells(ultlinmov, 1)
                        .Cells(ultlinmov, 2) = serieorigem
                        .Cells(ultlinmov, 3) = "Saída"
                        .Cells(ultlinmov, 4) = localorigem
                        .Cells(ultlinmov, 5) = areaorigem
                        .Cells(ultlinmov, 8) = zonaorigem
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                        .Cells(ultlinmov, 2) = serieorigem
                        .Cells(ultlinmov, 3) = "Entrada"
                        .Cells(ultlinmov, 6) = localdest
                        .Cells(ultlinmov, 7) = areadest
                        .Cells(ultlinmov, 8) = zonadest
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = Now
                        datahora = .Cells(ultlinmov, 1)
                        .Cells(ultlinmov, 2) = seriepermuta
                        .Cells(ultlinmov, 3) = "Saída"
                        .Cells(ultlinmov, 4) = localseriepermuta
                        .Cells(ultlinmov, 5) = areaseriepermuta
                        .Cells(ultlinmov, 8) = zonaseriepermuta
                        ultlinmov = ultlinmov + 1
                        .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                        .Cells(ultlinmov, 2) = seriepermuta
                        .Cells(ultlinmov, 3) = "Entrada"
                        .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
                        .Cells(ultlinmov, 7) = "0000"
                        .Cells(ultlinmov, 8) = "BRIGADA"
                    End With
                    .Cells(linmapa, 4) = localdest
                    .Cells(linmapa, 3) = edfdest
                    .Cells(linmapa, 2) = areadest
                    .Cells(linmapa, 9) = zonadest
                End If
            End With
        End If
'    End If

    '    'CAMPO >> TERCEIRIZADA
    '
    '    If Info.Cells(6, 1) <> "RESERVA TÉCNICA" And Info.Cells(6, 1) <> "MANUTENÇÃO - MAREFIRE" _
    '    And Info.Cells(6, 1) <> "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) = "MANUTENÇÃO - MAREFIRE" Then
    '    If statusserieorigem = "Em dia" Then
    '        Application.Speech.Speak "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a empresa de manutenção", speakasync:=True
    '        MsgBox "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a empresa de manutenção.", , "SGES"
    '        GoTo fim:
    '    Else
    '        'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
    '        With Movimentacao.ListObjects(1).DataBodyRange
    '            ultlinmov = ultlinmov + 1
    '            .Cells(ultlinmov, 1) = Now
    '            datahora = .Cells(ultlinmov, 1)
    '            .Cells(ultlinmov, 2) = serieorigem
    '            .Cells(ultlinmov, 3) = "Saída"
    '            .Cells(ultlinmov, 4) = localorigem
    '            .Cells(ultlinmov, 5) = areaorigem
    '            .Cells(ultlinmov, 8) = zonaorigem
    '            ultlinmov = ultlinmov + 1
    '            .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
    '            .Cells(ultlinmov, 2) = serieorigem
    '            .Cells(ultlinmov, 3) = "Entrada"
    '            .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
    '            .Cells(ultlinmov, 7) = "0000"
    '            .Cells(ultlinmov, 8) = "BRIGADA"
    '            ultlinmov = ultlinmov + 1
    '            .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
    '            .Cells(ultlinmov, 2) = serieorigem
    '            .Cells(ultlinmov, 3) = "Saída"
    '            .Cells(ultlinmov, 4) = "MANUTENÇÃO - BRIGADA"
    '            .Cells(ultlinmov, 5) = "0000"
    '            .Cells(ultlinmov, 8) = "BRIGADA"
    '            ultlinmov = ultlinmov + 1
    '            .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
    '            .Cells(ultlinmov, 2) = serieorigem
    '            .Cells(ultlinmov, 3) = "Entrada"
    '            .Cells(ultlinmov, 6) = localdest
    '            .Cells(ultlinmov, 7) = areadest
    '            .Cells(ultlinmov, 8) = zonadest
    '        End With
    '        With MapaAtual.ListObjects(1).DataBodyRange
    '            .Cells(linmapa, 4) = localdest
    '            .Cells(linmapa, 3) = edfdest
    '            .Cells(linmapa, 2) = areadest
    '            .Cells(linmapa, 9) = zonadest
    '
    '            '##### CHECAGEM EXT RESERVA  #####
    '
    '            linmapaseriereserva = 1
    '            Do Until linmapaseriereserva > ultlinmapa 'busca status geral do extintor
    '
    '                If seriereserva = .Cells(linmapaseriereserva, 8) Then
    '
    '                    localseriereserva = .Cells(linmapaseriereserva, 4)
    '                    edfseriereserva = .Cells(linmapaseriereserva, 3)
    '                    areaseriereserva = .Cells(linmapaseriereserva, 2)
    '                    zonaseriereserva = .Cells(linmapaseriereserva, 9)
    '                    statusseriereserva = UCase(.Cells(linmapaseriereserva, 23))
    '                    Exit Do
    '                End If
    '
    '                linmapaseriereserva = linmapaseriereserva + 1
    '            Loop
    ''            If statusseriereserva <> "EM DIA" Then
    ''
    ''                Application.Speech.Speak "Este extintor não está em dia. Deseja inserir um novo extintor?", speakasync:=True
    ''                resultado = MsgBox("Movimentação imprópria. Apenas Extintores EM DIA poderão substituir extintores no campo.", vbYesNoCancel, "SGES")
    ''                If resultado = vbYes Then
    ''                    GoTo voltainput:
    ''                Else
    ''                    MsgBox "Movimentação cancelada"
    ''                    GoTo fim:
    ''                End If
    ''            Else
    ''
    '''                If localseriereserva <> "RESERVA TÉCNICA" Then
    '''                    Application.Speech.Speak "Este extintor não está na reserva técnica. Deseja inserir um novo extintor?", speakasync:=True
    '''                    resultado = MsgBox("Movimentação imprópria. Apenas Extintores da reserva técnica poderão substituir extintores no CAMPO.", vbYesNoCancel, "SGES")
    '''                    If resultado = vbYes Then
    '''                        GoTo voltainput:
    '''                    Else
    '''                        MsgBox "Movimentação cancelada"
    '''                        GoTo fim:
    '''                    End If
    '''
    '''                    '##### FIM CHECAGEM EXT RESERVA  #####
    '''                Else
    '                    .Cells(linmapaseriereserva, 4) = localorigem
    '                    .Cells(linmapaseriereserva, 3) = edforigem
    '                    .Cells(linmapaseriereserva, 2) = areaorigem
    '                    .Cells(linmapaseriereserva, 9) = zonaorigem
    '''                End If
    ''            End If
    '        End With
    '    End If
    '    End If

    'CAMPO >> terceirizada

    If Info.Cells(6, 1) <> "RESERVA TÉCNICA" And Info.Cells(6, 1) <> "MANUTENÇÃO - MAREFIRE" _
    And Info.Cells(6, 1) <> "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) = "MANUTENÇÃO - MAREFIRE" Then
        If statusserieorigem = "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a empresa de manutenção", speakasync:=True
            MsgBox "Movimentação imprópria. Extintores EM DIA não poderão ser movidos para a empresa de manutenção.", , "SGES"
            GoTo fim:
'        End If
        ElseIf localorigem = "APOIO A PARADA - ANGRA 1" Or cap = "100K" Or cap _
            = "250K" Or cap = "37L" Or cap = "40L" Or serieorigem = "2189NL10L" Then
            'cilindros co2, CILINDROS PQS CAMINHOES E CILINDROS N2 CAMINHOES
                GoTo apoioparada1 'SERVE TBM PARA CILINDROS CAMINHOES
                
        End If
'            End If
voltainput3:
            seriereserva = UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            If seriereserva = vbNullString Then

                MsgBox "Movimentação cancelada"
                GoTo fim:

            End If

            '##### CHECAGEM EXT RESERVA  #####
            With MapaAtual.ListObjects(1).DataBodyRange
                linmapaseriereserva = 1
                Do Until linmapaseriereserva > ultlinmapa 'busca status geral do extintor

                    If seriereserva = .Cells(linmapaseriereserva, 8) Then

                        localseriereserva = .Cells(linmapaseriereserva, 4)
                        edfseriereserva = .Cells(linmapaseriereserva, 3)
                        areaseriereserva = .Cells(linmapaseriereserva, 2)
                        zonaseriereserva = .Cells(linmapaseriereserva, 9)
                        statusseriereserva = UCase$(.Cells(linmapaseriereserva, 23))
                        Exit Do
                    End If

                    linmapaseriereserva = linmapaseriereserva + 1
                Loop
                If statusseriereserva <> "EM DIA" And statusseriereserva <> "VENCENDO" Then

                    Application.Speech.Speak "Este extintor não está em dia. Deseja inserir um novo extintor?", speakasync:=True
                    resultado = MsgBox("Movimentação imprópria. Apenas Extintores EM DIA poderão substituir extintores no campo. Deseja escolher outro extintor?", vbYesNo, "SGES")
                    If resultado = vbYes Then
                        GoTo voltainput3:
                    Else
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                    End If
                Else
                    '
                    '                If localseriereserva <> "RESERVA TÉCNICA" Then
                    '                    Application.Speech.Speak "Este extintor não está na reserva técnica. Deseja inserir um novo extintor?", speakasync:=True
                    '                    resultado = MsgBox("Movimentação imprópria. Apenas Extintores da reserva técnica poderão substituir extintores no CAMPO.", vbYesNoCancel, "SGES")
                    '                    If resultado = vbYes Then
                    '                        GoTo voltainput:
                    '                    Else
                    '                        MsgBox "Movimentação cancelada"
                    '                        GoTo fim:
                    '                    End If
                    '
                    '                    '##### FIM CHECAGEM EXT RESERVA  #####
                    '                Else
                    .Cells(linmapaseriereserva, 4) = localorigem
                    .Cells(linmapaseriereserva, 3) = edforigem
                    .Cells(linmapaseriereserva, 2) = areaorigem
                    .Cells(linmapaseriereserva, 9) = zonaorigem
                    '                End If
                End If
            End With
apoioparada1:
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 7) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 5) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            
                If localorigem = localseriereserva Or localorigem = "APOIO A PARADA - ANGRA 1" _
                Or localorigem = "CAMINHÃO DE BOMBEIRO - AHQ02" _
                Or localorigem = "CAMINHÃO DE BOMBEIRO - AHQ01" Or localorigem = "CAMINHÃO DE BOMBEIRO - ABT01" Then
            
                    GoTo pula2:
                Else
            
            
                    ultlinmov = ultlinmov + 1
                    .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 4)
                    .Cells(ultlinmov, 2) = seriereserva
                    .Cells(ultlinmov, 3) = "Saída"
                    .Cells(ultlinmov, 4) = "RESERVA TÉCNICA"
                    .Cells(ultlinmov, 5) = "1111"
                    .Cells(ultlinmov, 8) = "BRIGADA"
                    ultlinmov = ultlinmov + 1
                    .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 5)
                    .Cells(ultlinmov, 2) = seriereserva
                    .Cells(ultlinmov, 3) = "Entrada"
                    .Cells(ultlinmov, 6) = localorigem
                    .Cells(ultlinmov, 7) = areaorigem
                    .Cells(ultlinmov, 8) = zonaorigem
                End If
pula2:
            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest


            End With
        End If
'    End If


    ' TERCEIRIZADA >> CAMPO

    If Info.Cells(6, 1) = "MANUTENÇÃO - MAREFIRE" And Info.Cells(12, 13) <> "RESERVA TÉCNICA" And Info.Cells(12, 13) <> "MANUTENÇÃO - MAREFIRE" _
                                                                                                                       And Info.Cells(12, 13) <> "MANUTENÇÃO - BRIGADA" Then
        If statusserieorigem <> "Em dia" Then
            Application.Speech.Speak "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para o CAMPO ", speakasync:=True
            MsgBox "Movimentação imprópria. Apenas Extintores EM DIA poderão ser movidos para o CAMPO.", , "SGES"
            GoTo fim:
        Else
            'seriereserva = UCase(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 7) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 2)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = "MANUTENÇÃO - BRIGADA"
                .Cells(ultlinmov, 5) = "0000"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 3)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = "RESERVA TÉCNICA"
                .Cells(ultlinmov, 7) = "1111"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = DateAdd("s", 4, Now)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = "RESERVA TÉCNICA"
                .Cells(ultlinmov, 5) = "1111"
                .Cells(ultlinmov, 8) = "BRIGADA"
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = DateAdd("s", 5, Now)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest
            End With
        End If
    End If


    'CAMPO >> CAMPO

    If Info.Cells(6, 1) <> "RESERVA TÉCNICA" And Info.Cells(6, 1) <> "MANUTENÇÃO - MAREFIRE" _
                                                                  And Info.Cells(6, 1) <> "MANUTENÇÃO - BRIGADA" And Info.Cells(12, 13) <> "RESERVA TÉCNICA" _
                                                                  And Info.Cells(12, 13) <> "MANUTENÇÃO - MAREFIRE" _
                                                                  And Info.Cells(12, 13) <> "MANUTENÇÃO - BRIGADA" Then
        If statusserieorigem <> "Em dia" And statusserieorigem <> "Vencendo" Then
            Application.Speech.Speak "Movimentação imprópria. Extintores Vencidos só podem ser movidos para a manutenção ", speakasync:=True
            MsgBox "Movimentação imprópria. Extintores Vencidos só podem ser movidos para a manutenção.", , "SGES"
            GoTo fim:
        Else
            'seriepermuta = UCase(InputBox("Digite o Número de Série do Extintor que será substituido.", "Repondo Extintor", vbOKCancel))
            With Movimentacao.ListObjects(1).DataBodyRange
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = Now
                datahora = .Cells(ultlinmov, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Saída"
                .Cells(ultlinmov, 4) = localorigem
                .Cells(ultlinmov, 5) = areaorigem
                .Cells(ultlinmov, 8) = zonaorigem
                ultlinmov = ultlinmov + 1
                .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                .Cells(ultlinmov, 2) = serieorigem
                .Cells(ultlinmov, 3) = "Entrada"
                .Cells(ultlinmov, 6) = localdest
                .Cells(ultlinmov, 7) = areadest
                .Cells(ultlinmov, 8) = zonadest

            End With
            With MapaAtual.ListObjects(1).DataBodyRange
                .Cells(linmapa, 4) = localdest
                .Cells(linmapa, 3) = edfdest
                .Cells(linmapa, 2) = areadest
                .Cells(linmapa, 9) = zonadest

                '##### CHECAGEM EXT permuta no campo  #####

                linmapaseriepermuta = 1
                Do Until linmapaseriepermuta > ultlinmapa 'busca status geral do extintor

                    If Info.Cells(12, 13) & " " & Info.Cells(14, 9) = .Cells(linmapaseriepermuta, 4) & " " & .Cells(linmapaseriepermuta, 2) Then


                        seriepermuta = .Cells(linmapaseriepermuta, 8)
                        localseriepermuta = .Cells(linmapaseriepermuta, 4)
                        edfseriepermuta = .Cells(linmapaseriepermuta, 3)
                        areaseriepermuta = .Cells(linmapaseriepermuta, 2)
                        zonaseriepermuta = .Cells(linmapaseriepermuta, 9)
                        statusseriepermuta = .Cells(linmapaseriepermuta, 23)
                        localconcatpermuta = localseriepermuta & " " & areaseriepermuta
                        Exit Do
                    End If

                    linmapaseriepermuta = linmapaseriepermuta + 1
                Loop
            
                If seriepermuta <> serieorigem Then
                    If statusseriepermuta <> "Em dia" And statusseriepermuta <> "Vencendo" Then

                        Application.Speech.Speak "Movimentação imprópria. Extintores Vencidos só podem ser movidos para a manutenção ", speakasync:=True
                        MsgBox "Movimentação imprópria. Extintores Vencidos só podem ser movidos para a manutenção.", , "SGES"
                    Else


                        .Cells(linmapaseriepermuta, 4) = localorigem
                        .Cells(linmapaseriepermuta, 3) = edforigem
                        .Cells(linmapaseriepermuta, 2) = areaorigem
                        .Cells(linmapaseriepermuta, 9) = zonaorigem
                        With Movimentacao.ListObjects(1).DataBodyRange
                            ultlinmov = ultlinmov + 1
                            .Cells(ultlinmov, 1) = Now
                            datahora = .Cells(ultlinmov, 1)
                            .Cells(ultlinmov, 2) = seriepermuta
                            .Cells(ultlinmov, 3) = "Saída"
                            .Cells(ultlinmov, 4) = localdest
                            .Cells(ultlinmov, 5) = areadest
                            .Cells(ultlinmov, 8) = zonadest
                            ultlinmov = ultlinmov + 1
                            .Cells(ultlinmov, 1) = datahora + TimeSerial(0, 0, 1)
                            .Cells(ultlinmov, 2) = seriepermuta
                            .Cells(ultlinmov, 3) = "Entrada"
                            .Cells(ultlinmov, 6) = localorigem
                            .Cells(ultlinmov, 7) = areaorigem
                            .Cells(ultlinmov, 8) = zonaorigem
                        End With
                    End If
                End If
            End With


        End If
    End If
    Movimentacao.ListObjects("tbHistMov14").DataBodyRange.Calculate
    Info.ListObjects("tbHistMov").DataBodyRange.Calculate
fim:
    '    Application.EnableEvents = True
    '    Application.ScreenUpdating = True
End Sub






