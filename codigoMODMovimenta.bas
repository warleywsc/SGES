Attribute VB_Name = "MODMovimenta"
'@Folder("SGES2020")
Option Explicit

Public Sub movimentaext()
    Dim resultado As VbMsgBoxResult
    Dim localorigem As String, localdestino As String, status As String
    Dim areadestino As String, zonadestino As String, edfdestino As String
    Dim areaorigem As String, edforigem As String, zonaorigem As String
    Dim serieentrada As String, seriesaida As String
    Dim i     As Long
    Dim ultlinmapa As Long, ultlinext As Long, ultlinmov As Long
    Dim localxarea As Variant
    Dim lin   As Long
    Dim encontrado As String
    ultlinmapa = MapaAtual.Cells(MapaAtual.Rows.Count, "G").End(xlUp).Row
    ultlinext = Extintores.Cells(Extintores.Rows.Count, "G").End(xlUp).Row
    ultlinmov = Sheets("Movimentação").Cells(Sheets("Movimentação").Rows.Count, "G").End(xlUp).Row
    localorigem = Info.Cells(6, 1)
    areaorigem = Info.Cells(7, 1)
    If InStr(localorigem, " - ") > 0 Then
        edforigem = Left$(localorigem, InStr(localorigem, " - ") - 1)
    
    Else
        edforigem = localorigem
    End If
    zonaorigem = Info.Cells(8, 1)
    localdestino = Info.Cells(12, 13)
    areadestino = Info.Cells(14, 9)
    zonadestino = Info.Cells(14, 13)
    
    'edificio entrada
    If InStr(localdestino, " - ") > 0 Then
        edfdestino = Left$(localdestino, InStr(localdestino, " - ") - 1)
    
    Else
        edfdestino = localdestino
    End If
    
   
    serieentrada = vbNullString
    seriesaida = Info.Cells(8, 9)
    With Info
        If .Cells(12, 6) = "Sim" Then  'verifica se houve movimentação
            limpafiltrosmov

            On Error GoTo fim:


            localxarea = .Cells(12, 13).Value & " - " & .Cells(14, 9).Value

            lin = 9
            encontrado = vbNullString


            Do Until locais.Cells(lin, 10) = vbNullString
                If locais.Cells(lin, 8).Value & " - " & locais.Cells(lin, 9).Value = localxarea Then
                            
                    'SE ENCONTRAR LOCAL CONTINUA CADASTRO
                    GoTo ContinuaMAPA:


                Else
                    encontrado = "Não"

                End If
                lin = lin + 1
            Loop
            If encontrado = "Não" Then
                MsgBox "Local não encontrado! Talvez seja necessário modificar a área.", , "Erro de Local"
                GoTo fim:
            End If
ContinuaMAPA:

            If encontrado = "Não" Then
                localxarea = Info.Cells(6, 1).Value & " - " & Info.Cells(7, 1).Value
                lin = 9
                Do Until MapaAtual.Cells(lin, 14).Row = MapaAtual.Cells(MapaAtual.Rows.Count, "G").End(xlUp).Row
                    If MapaAtual.Cells(lin, 10).Value & " - " & MapaAtual.Cells(lin, 8).Value = localxarea Then
                            
                            
                        '### LIMPA EXTINTOR DO LOCAL ANTIGO NO MAPA
                        MapaAtual.Cells(lin, 7).Value = vbNullString
                        MapaAtual.Cells(lin, 12).Value = vbNullString
                        MapaAtual.Cells(lin, 13).Value = vbNullString
                        MapaAtual.Cells(lin, 14).Value = vbNullString

                        'GoTo ContinuaMOV:
                        Exit Do

                    End If
                    lin = lin + 1
                Loop
                    
                '### INSERE LOCAL NOVO NA ULTIMA LINHA DO MAPA JUNTAMENTE COM O RESTANTE DAS INFO
                    
                    
                lin = MapaAtual.Cells(MapaAtual.Rows.Count, "G").End(xlUp).Offset(1, 0).Row
                MapaAtual.Cells(lin, 7).Value = Info.Range("I12").Value 'SUPORTE
                MapaAtual.Cells(lin, 8).Value = Info.Range("I14").Value 'ÁREA
                If InStr(.Range("M12").Value, " - ") - 1 = -1 Then 'EDIFICIO
                    MapaAtual.Cells(lin, 9).Value = .Range("M12").Value
                Else
                    MapaAtual.Cells(lin, 9).Value = Left$(.Range("M12").Value, InStr(.Range("M12").Value, " - ") - 1) 'edificio entrada
                End If
                MapaAtual.Cells(lin, 10).Value = Info.Range("M12").Value 'LOCAL
                MapaAtual.Cells(lin, 11).Value = Info.Range("M8").Value 'TIPO
                MapaAtual.Cells(lin, 12).Value = Info.Range("M10").Value 'CAP
                MapaAtual.Cells(lin, 13).Value = Info.Range("I10").Value 'FAB
                MapaAtual.Cells(lin, 14).Value = Info.Range("I8").Value 'SERIE
                MapaAtual.Cells(lin, 15).Value = Info.Range("M14").Value 'ZONA
                   
                MapaAtual.Cells(lin, 27).Value = Info.Range("G23").Value 'OBS
                        
                        
            End If
            'ORIGEM = QUALQUER <> DE "MANUTENÇÃO - BRIGADA e STATUS GERAL = EM DIA"
            'DESTINO = QUALQUER <> DE "MANUTENÇÃO - MAREFIRE"
        
            If localorigem <> "MANUTENÇÃO - BRIGADA" And localorigem <> "MANUTENÇÃO - MAREFIRE" _
               And localdestino <> "MANUTENÇÃO - BRIGADA" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput:
                If status <> UCase$("Em dia") Then 'confere se status geral está em dia
                
                    MsgBox "Extintor não elegível para este tipo de movimentação. Cancelando operação de Movimentação...", , "Movimentação cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                        If MapaAtual.Cells(i, 10) = UCase$("RESERVA TÉCNICA") Then ' verifica se extintor está na reserva
                                            'saída em mov
                                            ultlinmov = ultlinmov + 1
                                            Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                            Sheets("Movimentação").Cells(ultlinmov, 10) = "RESERVA TÉCNICA"
                                            Sheets("Movimentação").Cells(ultlinmov, 11) = "1111"
                                            Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                            ultlinmov = ultlinmov + 1
                                            '                                         entrada em mov
                                            Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                            Sheets("Movimentação").Cells(ultlinmov, 12) = localorigem
                                            Sheets("Movimentação").Cells(ultlinmov, 13) = areaorigem
                                            Sheets("Movimentação").Cells(ultlinmov, 14) = zonaorigem
                                         
                                         
                                            MapaAtual.Cells(i, 10) = localorigem
                                            MapaAtual.Cells(i, 9) = edforigem
                                            MapaAtual.Cells(i, 8) = areaorigem
                                            MapaAtual.Cells(i, 15) = zonaorigem
                                            GoTo regsaida:
                                        Else
                                            resultado = MsgBox("Este extintor não se encontra na reserva técnica! Deseja excolher outro extintor?", vbYesNo, "Extintor impróprio")
                                            If resultado = vbYes Then
                                                GoTo voltainput:
                                            Else
                                                MsgBox "Movimentação cancelada"
                                                GoTo fim:
                                            End If
                                        
                                        End If
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor saída
regsaida:
                                i = 9
                                Do Until i > ultlinmapa
                        
                                    If seriesaida = MapaAtual.Cells(i, 14) Then
                                        ultlinmov = ultlinmov + 1
                                        'saída em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                        Sheets("Movimentação").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimentação").Cells(ultlinmov, 12) = "MANUTENÇÃO - BRIGADA"
                                        Sheets("Movimentação").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        MapaAtual.Cells(i, 10) = "MANUTENÇÃO - BRIGADA"
                                        MapaAtual.Cells(i, 9) = edfdestino
                                        MapaAtual.Cells(i, 8) = "0000"
                                        MapaAtual.Cells(i, 15) = "BRIGADA"
                                        Exit Do
                                    End If
                                    i = i + 1
                                Loop
                        
                                MsgBox "Movimentação concluída!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor não encontrado! Deseja tentar novamente?", vbYesNo, "Extintor não cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput:
                                Else
                                    MsgBox "Movimentação cancelada"
                                    GoTo fim:
                                End If
                        
                        
                            End If
                            i = i + 1
                        Loop
                
                    End If
                
                
                
                
                End If
                
                'OUTROS LOCAIS PARA MANUTENÇÃO - BRIGADA & status <> de Em dia
            ElseIf localorigem <> "MANUTENÇÃO - BRIGADA" And localdestino = "MANUTENÇÃO - BRIGADA" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput2:
                If status = "Em dia" Then 'confere se status geral está em dia
                
                    MsgBox "Status do extintor " & "Em Dia" & ". Cancelando operação de Movimentação...", , "Movimentação cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                        If MapaAtual.Cells(i, 10) = UCase$("RESERVA TÉCNICA") Then ' verifica se extintor está na reserva
                                            'saída em mov
                                            ultlinmov = ultlinmov + 1
                                            Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                            Sheets("Movimentação").Cells(ultlinmov, 10) = "RESERVA TÉCNICA"
                                            Sheets("Movimentação").Cells(ultlinmov, 11) = "1111"
                                            Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                            ultlinmov = ultlinmov + 1
                                            '                                         entrada em mov
                                            Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                            Sheets("Movimentação").Cells(ultlinmov, 12) = localorigem
                                            Sheets("Movimentação").Cells(ultlinmov, 13) = areaorigem
                                            Sheets("Movimentação").Cells(ultlinmov, 14) = zonaorigem
                                         
                                         
                                            MapaAtual.Cells(i, 10) = localorigem
                                            MapaAtual.Cells(i, 9) = edforigem
                                            MapaAtual.Cells(i, 8) = areaorigem
                                            MapaAtual.Cells(i, 15) = zonaorigem
                                            GoTo regsaida2:
                                        Else
                                            resultado = MsgBox("Este extintor não se encontra na reserva técnica! Deseja excolher outro extintor?", vbYesNo, "Extintor impróprio")
                                            If resultado = vbYes Then
                                                GoTo voltainput2:
                                            Else
                                                MsgBox "Movimentação cancelada"
                                                GoTo fim:
                                            End If
                                        
                                        End If
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor saída
regsaida2:
                                i = 9
                                Do Until i > ultlinmapa
                        
                                    If seriesaida = MapaAtual.Cells(i, 14) Then
                                        ultlinmov = ultlinmov + 1
                                        'saída em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieentrada
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                        Sheets("Movimentação").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieentrada
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimentação").Cells(ultlinmov, 12) = "MANUTENÇÃO - BRIGADA"
                                        Sheets("Movimentação").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        MapaAtual.Cells(i, 10) = "MANUTENÇÃO - BRIGADA"
                                        MapaAtual.Cells(i, 9) = edfdestino
                                        MapaAtual.Cells(i, 8) = "0000"
                                        MapaAtual.Cells(i, 15) = "BRIGADA"
                                        MsgBox "Movimentação concluída!"
                                        GoTo fim:
                                    End If
                                    i = i + 1
                                Loop
                        
                                    
                        
                            End If
                                
                            i = i + 1
                        Loop
                        If i > ultlinext Then
                            resultado = MsgBox("Extintor não encontrado! Deseja tentar novamente?", vbYesNo, "Extintor não cadastrado")
                            If resultado = vbYes Then
                                GoTo voltainput2:
                            Else
                                MsgBox "Movimentação cancelada"
                                GoTo fim:
                            End If
                        
                        
                        End If
                    End If
          
                End If
          
                'ORIGEM = QUALQUER <> DE "MANUTENÇÃO - BRIGADA e STATUS GERAL = EM DIA"
                'DESTINO = "MANUTENÇÃO - MAREFIRE"
        
            ElseIf localorigem <> "MANUTENÇÃO - BRIGADA" And localdestino = "MANUTENÇÃO - MAREFIRE" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput4:
                If status = UCase$("Em dia") Then 'confere se status geral está em dia
                
                    MsgBox "Extintor não elegível para este tipo de movimentação. Cancelando operação de Movimentação...", , "Movimentação cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                        If MapaAtual.Cells(i, 10) = UCase$("RESERVA TÉCNICA") Then ' verifica se extintor está na reserva
                                            'saída em mov
                                            ultlinmov = ultlinmov + 1
                                            Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                            Sheets("Movimentação").Cells(ultlinmov, 10) = "RESERVA TÉCNICA"
                                            Sheets("Movimentação").Cells(ultlinmov, 11) = "1111"
                                            Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                            ultlinmov = ultlinmov + 1
                                            '                                         entrada em mov
                                            Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                            Sheets("Movimentação").Cells(ultlinmov, 12) = localorigem
                                            Sheets("Movimentação").Cells(ultlinmov, 13) = areaorigem
                                            Sheets("Movimentação").Cells(ultlinmov, 14) = zonaorigem
                                         
                                         
                                            MapaAtual.Cells(i, 10) = localorigem
                                            MapaAtual.Cells(i, 9) = edforigem
                                            MapaAtual.Cells(i, 8) = areaorigem
                                            MapaAtual.Cells(i, 15) = zonaorigem
                                            GoTo regsaida4:
                                        Else
                                            resultado = MsgBox("Este extintor não se encontra na reserva técnica! Deseja excolher outro extintor?", vbYesNo, "Extintor impróprio")
                                            If resultado = vbYes Then
                                                GoTo voltainput4:
                                            Else
                                                MsgBox "Movimentação cancelada"
                                                GoTo fim:
                                            End If
                                        
                                        End If
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor saída
regsaida4:
                                i = 9
                                Do Until i > ultlinmapa
                        
                                    If seriesaida = MapaAtual.Cells(i, 14) Then
                                    
                                        '### primeira mov
                                        ultlinmov = ultlinmov + 1
                                        'saída em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                        Sheets("Movimentação").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimentação").Cells(ultlinmov, 12) = "MANUTENÇÃO - BRIGADA"
                                        Sheets("Movimentação").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        ultlinmov = ultlinmov + 1
                                         
                                        '### segunda mov
                                        'saída em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                        Sheets("Movimentação").Cells(ultlinmov, 10) = "MANUTENÇÃO - BRIGADA"
                                        Sheets("Movimentação").Cells(ultlinmov, 11) = "0000"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimentação").Cells(ultlinmov, 12) = "MANUTENÇÃO - MAREFIRE"
                                        Sheets("Movimentação").Cells(ultlinmov, 13) = "9999"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "MAREFIRE"
                                         
                                        MapaAtual.Cells(i, 10) = "MANUTENÇÃO - MAREFIRE"
                                        MapaAtual.Cells(i, 9) = "MANUTENÇÃO"
                                        MapaAtual.Cells(i, 8) = "9999"
                                        MapaAtual.Cells(i, 15) = "MAREFIRE"
                                        Exit Do
                                    End If
                                    i = i + 1
                                Loop
                        
                                MsgBox "Movimentação concluída!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor não encontrado! Deseja tentar novamente?", vbYesNo, "Extintor não cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput:
                                Else
                                    MsgBox "Movimentação cancelada"
                                    GoTo fim:
                                End If
                        
                        
                            End If
                            i = i + 1
                        Loop
                
                    End If
                
                End If
                '#######################################################################################################
       
                'ORIGEM = "MANUTENÇÃO - MAREFIRE"
                'DESTINO = "RESERVA TÉCNICA"
        
            ElseIf localorigem = "MANUTENÇÃO - MAREFIRE" And localdestino = "RESERVA TÉCNICA" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput5:
                If status <> UCase$("Em dia") Then 'confere se status geral está em dia
                
                    MsgBox "Extintor não elegível para este tipo de movimentação. Cancelando operação de Movimentação...", , "Movimentação cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                

regsaida5:
                                i = 9
                                Do Until i > ultlinmapa
                        
                                    If seriesaida = MapaAtual.Cells(i, 14) Then
                                    
                                        '### primeira mov
                                        ultlinmov = ultlinmov + 1
                                        'saída em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                        Sheets("Movimentação").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimentação").Cells(ultlinmov, 12) = "MANUTENÇÃO - BRIGADA"
                                        Sheets("Movimentação").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        ultlinmov = ultlinmov + 1
                                         
                                        '### segunda mov
                                        'saída em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                        Sheets("Movimentação").Cells(ultlinmov, 10) = "MANUTENÇÃO - BRIGADA"
                                        Sheets("Movimentação").Cells(ultlinmov, 11) = "0000"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimentação").Cells(ultlinmov, 12) = "RESERVA TÉCNICA"
                                        Sheets("Movimentação").Cells(ultlinmov, 13) = "1111"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                        MapaAtual.Cells(i, 10) = "RESERVA TÉCNICA"
                                        MapaAtual.Cells(i, 9) = "RESERVA TÉCNICA"
                                        MapaAtual.Cells(i, 8) = "1111"
                                        MapaAtual.Cells(i, 15) = "BRIGADA"
                                        Exit Do
                                    End If
                                    i = i + 1
                                Loop
                        
                                MsgBox "Movimentação concluída!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor não encontrado! Deseja tentar novamente?", vbYesNo, "Extintor não cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput5:
                                Else
                                    MsgBox "Movimentação cancelada"
                                    GoTo fim:
                                End If
                        
                        
                            End If
                            i = i + 1
                        Loop
                
                    End If
                
                End If
       
       
                '########################################################################################
       
       
                'ELSEIF
                'ORIGEM = MANUTENÇÃO - BRIGADA & STATUS GERAL <> EM DIA
                'DESTINO = MANUTENÇÃO - MAREFIRE
            
            ElseIf localorigem = "MANUTENÇÃO - BRIGADA" And localdestino = "MANUTENÇÃO - MAREFIRE" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput3:
                If status = UCase$("Em dia") Then 'confere se status geral está em dia
                
                    MsgBox "Status do extintor " & "Em Dia" & ". Cancelando operação de Movimentação...", , "Movimentação cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o Número de Série do Extintor que substituirá o Extintor que está saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimentação cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                            
                                        'saída em mov
                                        ultlinmov = ultlinmov + 1
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Saída"
                                        Sheets("Movimentação").Cells(ultlinmov, 10) = "MANUTENÇÃO - BRIGADA"
                                        Sheets("Movimentação").Cells(ultlinmov, 11) = "0000"
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = "BRIGADA"
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimentação").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimentação").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimentação").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimentação").Cells(ultlinmov, 12) = localdestino
                                        Sheets("Movimentação").Cells(ultlinmov, 13) = areadestino
                                        Sheets("Movimentação").Cells(ultlinmov, 14) = zonadestino
                                         
                                         
                                        MapaAtual.Cells(i, 10) = localdestino
                                        MapaAtual.Cells(i, 9) = edfdestino
                                        MapaAtual.Cells(i, 8) = areadestino
                                        MapaAtual.Cells(i, 15) = zonadestino
                                        GoTo regsaida3:
                                         
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor saída
regsaida3:
                                
                        
                                MsgBox "Movimentação concluída!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor não encontrado! Deseja tentar novamente?", vbYesNo, "Extintor não cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput3:
                                Else
                                    MsgBox "Movimentação cancelada"
                                    GoTo fim:
                                End If
                        
                        
                            End If
                            i = i + 1
                        Loop
                
                    End If
  
                End If
      
            End If
        
        End If
        

    End With
fim:
    Sheets("Movimentação").ListObjects("tbHistMov14").DataBodyRange.Calculate
    Info.ListObjects("tbHistMov").DataBodyRange.Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub




