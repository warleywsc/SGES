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
    ultlinmov = Sheets("Movimenta��o").Cells(Sheets("Movimenta��o").Rows.Count, "G").End(xlUp).Row
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
        If .Cells(12, 6) = "Sim" Then  'verifica se houve movimenta��o
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
                    encontrado = "N�o"

                End If
                lin = lin + 1
            Loop
            If encontrado = "N�o" Then
                MsgBox "Local n�o encontrado! Talvez seja necess�rio modificar a �rea.", , "Erro de Local"
                GoTo fim:
            End If
ContinuaMAPA:

            If encontrado = "N�o" Then
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
                MapaAtual.Cells(lin, 8).Value = Info.Range("I14").Value '�REA
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
            'ORIGEM = QUALQUER <> DE "MANUTEN��O - BRIGADA e STATUS GERAL = EM DIA"
            'DESTINO = QUALQUER <> DE "MANUTEN��O - MAREFIRE"
        
            If localorigem <> "MANUTEN��O - BRIGADA" And localorigem <> "MANUTEN��O - MAREFIRE" _
               And localdestino <> "MANUTEN��O - BRIGADA" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput:
                If status <> UCase$("Em dia") Then 'confere se status geral est� em dia
                
                    MsgBox "Extintor n�o eleg�vel para este tipo de movimenta��o. Cancelando opera��o de Movimenta��o...", , "Movimenta��o cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o N�mero de S�rie do Extintor que substituir� o Extintor que est� saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimenta��o cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                        If MapaAtual.Cells(i, 10) = UCase$("RESERVA T�CNICA") Then ' verifica se extintor est� na reserva
                                            'sa�da em mov
                                            ultlinmov = ultlinmov + 1
                                            Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 10) = "RESERVA T�CNICA"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 11) = "1111"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                            ultlinmov = ultlinmov + 1
                                            '                                         entrada em mov
                                            Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 12) = localorigem
                                            Sheets("Movimenta��o").Cells(ultlinmov, 13) = areaorigem
                                            Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonaorigem
                                         
                                         
                                            MapaAtual.Cells(i, 10) = localorigem
                                            MapaAtual.Cells(i, 9) = edforigem
                                            MapaAtual.Cells(i, 8) = areaorigem
                                            MapaAtual.Cells(i, 15) = zonaorigem
                                            GoTo regsaida:
                                        Else
                                            resultado = MsgBox("Este extintor n�o se encontra na reserva t�cnica! Deseja excolher outro extintor?", vbYesNo, "Extintor impr�prio")
                                            If resultado = vbYes Then
                                                GoTo voltainput:
                                            Else
                                                MsgBox "Movimenta��o cancelada"
                                                GoTo fim:
                                            End If
                                        
                                        End If
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor sa�da
regsaida:
                                i = 9
                                Do Until i > ultlinmapa
                        
                                    If seriesaida = MapaAtual.Cells(i, 14) Then
                                        ultlinmov = ultlinmov + 1
                                        'sa�da em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 12) = "MANUTEN��O - BRIGADA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        MapaAtual.Cells(i, 10) = "MANUTEN��O - BRIGADA"
                                        MapaAtual.Cells(i, 9) = edfdestino
                                        MapaAtual.Cells(i, 8) = "0000"
                                        MapaAtual.Cells(i, 15) = "BRIGADA"
                                        Exit Do
                                    End If
                                    i = i + 1
                                Loop
                        
                                MsgBox "Movimenta��o conclu�da!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor n�o encontrado! Deseja tentar novamente?", vbYesNo, "Extintor n�o cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput:
                                Else
                                    MsgBox "Movimenta��o cancelada"
                                    GoTo fim:
                                End If
                        
                        
                            End If
                            i = i + 1
                        Loop
                
                    End If
                
                
                
                
                End If
                
                'OUTROS LOCAIS PARA MANUTEN��O - BRIGADA & status <> de Em dia
            ElseIf localorigem <> "MANUTEN��O - BRIGADA" And localdestino = "MANUTEN��O - BRIGADA" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput2:
                If status = "Em dia" Then 'confere se status geral est� em dia
                
                    MsgBox "Status do extintor " & "Em Dia" & ". Cancelando opera��o de Movimenta��o...", , "Movimenta��o cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o N�mero de S�rie do Extintor que substituir� o Extintor que est� saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimenta��o cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                        If MapaAtual.Cells(i, 10) = UCase$("RESERVA T�CNICA") Then ' verifica se extintor est� na reserva
                                            'sa�da em mov
                                            ultlinmov = ultlinmov + 1
                                            Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 10) = "RESERVA T�CNICA"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 11) = "1111"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                            ultlinmov = ultlinmov + 1
                                            '                                         entrada em mov
                                            Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 12) = localorigem
                                            Sheets("Movimenta��o").Cells(ultlinmov, 13) = areaorigem
                                            Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonaorigem
                                         
                                         
                                            MapaAtual.Cells(i, 10) = localorigem
                                            MapaAtual.Cells(i, 9) = edforigem
                                            MapaAtual.Cells(i, 8) = areaorigem
                                            MapaAtual.Cells(i, 15) = zonaorigem
                                            GoTo regsaida2:
                                        Else
                                            resultado = MsgBox("Este extintor n�o se encontra na reserva t�cnica! Deseja excolher outro extintor?", vbYesNo, "Extintor impr�prio")
                                            If resultado = vbYes Then
                                                GoTo voltainput2:
                                            Else
                                                MsgBox "Movimenta��o cancelada"
                                                GoTo fim:
                                            End If
                                        
                                        End If
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor sa�da
regsaida2:
                                i = 9
                                Do Until i > ultlinmapa
                        
                                    If seriesaida = MapaAtual.Cells(i, 14) Then
                                        ultlinmov = ultlinmov + 1
                                        'sa�da em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieentrada
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieentrada
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 12) = "MANUTEN��O - BRIGADA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        MapaAtual.Cells(i, 10) = "MANUTEN��O - BRIGADA"
                                        MapaAtual.Cells(i, 9) = edfdestino
                                        MapaAtual.Cells(i, 8) = "0000"
                                        MapaAtual.Cells(i, 15) = "BRIGADA"
                                        MsgBox "Movimenta��o conclu�da!"
                                        GoTo fim:
                                    End If
                                    i = i + 1
                                Loop
                        
                                    
                        
                            End If
                                
                            i = i + 1
                        Loop
                        If i > ultlinext Then
                            resultado = MsgBox("Extintor n�o encontrado! Deseja tentar novamente?", vbYesNo, "Extintor n�o cadastrado")
                            If resultado = vbYes Then
                                GoTo voltainput2:
                            Else
                                MsgBox "Movimenta��o cancelada"
                                GoTo fim:
                            End If
                        
                        
                        End If
                    End If
          
                End If
          
                'ORIGEM = QUALQUER <> DE "MANUTEN��O - BRIGADA e STATUS GERAL = EM DIA"
                'DESTINO = "MANUTEN��O - MAREFIRE"
        
            ElseIf localorigem <> "MANUTEN��O - BRIGADA" And localdestino = "MANUTEN��O - MAREFIRE" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput4:
                If status = UCase$("Em dia") Then 'confere se status geral est� em dia
                
                    MsgBox "Extintor n�o eleg�vel para este tipo de movimenta��o. Cancelando opera��o de Movimenta��o...", , "Movimenta��o cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o N�mero de S�rie do Extintor que substituir� o Extintor que est� saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimenta��o cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                        If MapaAtual.Cells(i, 10) = UCase$("RESERVA T�CNICA") Then ' verifica se extintor est� na reserva
                                            'sa�da em mov
                                            ultlinmov = ultlinmov + 1
                                            Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 10) = "RESERVA T�CNICA"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 11) = "1111"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                            ultlinmov = ultlinmov + 1
                                            '                                         entrada em mov
                                            Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                            Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                            Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                            Sheets("Movimenta��o").Cells(ultlinmov, 12) = localorigem
                                            Sheets("Movimenta��o").Cells(ultlinmov, 13) = areaorigem
                                            Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonaorigem
                                         
                                         
                                            MapaAtual.Cells(i, 10) = localorigem
                                            MapaAtual.Cells(i, 9) = edforigem
                                            MapaAtual.Cells(i, 8) = areaorigem
                                            MapaAtual.Cells(i, 15) = zonaorigem
                                            GoTo regsaida4:
                                        Else
                                            resultado = MsgBox("Este extintor n�o se encontra na reserva t�cnica! Deseja excolher outro extintor?", vbYesNo, "Extintor impr�prio")
                                            If resultado = vbYes Then
                                                GoTo voltainput4:
                                            Else
                                                MsgBox "Movimenta��o cancelada"
                                                GoTo fim:
                                            End If
                                        
                                        End If
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor sa�da
regsaida4:
                                i = 9
                                Do Until i > ultlinmapa
                        
                                    If seriesaida = MapaAtual.Cells(i, 14) Then
                                    
                                        '### primeira mov
                                        ultlinmov = ultlinmov + 1
                                        'sa�da em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 12) = "MANUTEN��O - BRIGADA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        ultlinmov = ultlinmov + 1
                                         
                                        '### segunda mov
                                        'sa�da em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 10) = "MANUTEN��O - BRIGADA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 11) = "0000"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 12) = "MANUTEN��O - MAREFIRE"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 13) = "9999"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "MAREFIRE"
                                         
                                        MapaAtual.Cells(i, 10) = "MANUTEN��O - MAREFIRE"
                                        MapaAtual.Cells(i, 9) = "MANUTEN��O"
                                        MapaAtual.Cells(i, 8) = "9999"
                                        MapaAtual.Cells(i, 15) = "MAREFIRE"
                                        Exit Do
                                    End If
                                    i = i + 1
                                Loop
                        
                                MsgBox "Movimenta��o conclu�da!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor n�o encontrado! Deseja tentar novamente?", vbYesNo, "Extintor n�o cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput:
                                Else
                                    MsgBox "Movimenta��o cancelada"
                                    GoTo fim:
                                End If
                        
                        
                            End If
                            i = i + 1
                        Loop
                
                    End If
                
                End If
                '#######################################################################################################
       
                'ORIGEM = "MANUTEN��O - MAREFIRE"
                'DESTINO = "RESERVA T�CNICA"
        
            ElseIf localorigem = "MANUTEN��O - MAREFIRE" And localdestino = "RESERVA T�CNICA" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput5:
                If status <> UCase$("Em dia") Then 'confere se status geral est� em dia
                
                    MsgBox "Extintor n�o eleg�vel para este tipo de movimenta��o. Cancelando opera��o de Movimenta��o...", , "Movimenta��o cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o N�mero de S�rie do Extintor que substituir� o Extintor que est� saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimenta��o cancelada"
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
                                        'sa�da em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 10) = localorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 11) = areaorigem
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonaorigem
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 12) = "MANUTEN��O - BRIGADA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 13) = "0000"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                         
                                        ultlinmov = ultlinmov + 1
                                         
                                        '### segunda mov
                                        'sa�da em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 10) = "MANUTEN��O - BRIGADA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 11) = "0000"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = seriesaida 'serieSAIDA
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 12) = "RESERVA T�CNICA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 13) = "1111"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                         
                                        MapaAtual.Cells(i, 10) = "RESERVA T�CNICA"
                                        MapaAtual.Cells(i, 9) = "RESERVA T�CNICA"
                                        MapaAtual.Cells(i, 8) = "1111"
                                        MapaAtual.Cells(i, 15) = "BRIGADA"
                                        Exit Do
                                    End If
                                    i = i + 1
                                Loop
                        
                                MsgBox "Movimenta��o conclu�da!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor n�o encontrado! Deseja tentar novamente?", vbYesNo, "Extintor n�o cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput5:
                                Else
                                    MsgBox "Movimenta��o cancelada"
                                    GoTo fim:
                                End If
                        
                        
                            End If
                            i = i + 1
                        Loop
                
                    End If
                
                End If
       
       
                '########################################################################################
       
       
                'ELSEIF
                'ORIGEM = MANUTEN��O - BRIGADA & STATUS GERAL <> EM DIA
                'DESTINO = MANUTEN��O - MAREFIRE
            
            ElseIf localorigem = "MANUTEN��O - BRIGADA" And localdestino = "MANUTEN��O - MAREFIRE" Then
                i = 9
                Do Until i > ultlinmapa 'busca status geral do extintor
             
                    If .Cells(8, 9) = MapaAtual.Cells(i, 14) Then
                    
                        status = UCase$(MapaAtual.Cells(i, 29))
                    
                    End If
 
                    i = i + 1
                Loop
voltainput3:
                If status = UCase$("Em dia") Then 'confere se status geral est� em dia
                
                    MsgBox "Status do extintor " & "Em Dia" & ". Cancelando opera��o de Movimenta��o...", , "Movimenta��o cancelada"
                    GoTo fim:
                
                Else

                    serieentrada = UCase$(InputBox("Digite o N�mero de S�rie do Extintor que substituir� o Extintor que est� saindo", "Repondo Extintor", vbOKCancel))
                    If serieentrada = vbNullString Then
                
                        MsgBox "Movimenta��o cancelada"
                        GoTo fim:
                        
                    Else
                        
                        i = 9
                        Do Until i > ultlinext
                            If serieentrada = Extintores.Cells(i, 15) Then ' verifica se extintor existe na tbl extintores
                                
                                'altera extintor entrada
                                i = 9
                                Do Until i > ultlinmapa
                                
                                    If serieentrada = MapaAtual.Cells(i, 14) Then 'verifica se extintor existe no mapa
                                        
                                            
                                        'sa�da em mov
                                        ultlinmov = ultlinmov + 1
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Sa�da"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 10) = "MANUTEN��O - BRIGADA"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 11) = "0000"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = "BRIGADA"
                                        ultlinmov = ultlinmov + 1
                                        '                                         entrada em mov
                                        Sheets("Movimenta��o").Cells(ultlinmov, 7) = Now 'data
                                        Sheets("Movimenta��o").Cells(ultlinmov, 8) = serieentrada 'serieentrada
                                        Sheets("Movimenta��o").Cells(ultlinmov, 9) = "Entrada"
                                        Sheets("Movimenta��o").Cells(ultlinmov, 12) = localdestino
                                        Sheets("Movimenta��o").Cells(ultlinmov, 13) = areadestino
                                        Sheets("Movimenta��o").Cells(ultlinmov, 14) = zonadestino
                                         
                                         
                                        MapaAtual.Cells(i, 10) = localdestino
                                        MapaAtual.Cells(i, 9) = edfdestino
                                        MapaAtual.Cells(i, 8) = areadestino
                                        MapaAtual.Cells(i, 15) = zonadestino
                                        GoTo regsaida3:
                                         
                                        
                                
                                        
                                    End If
                                    i = i + 1
                                Loop
                                ' altera extintor sa�da
regsaida3:
                                
                        
                                MsgBox "Movimenta��o conclu�da!"
                                GoTo fim:
                        
                                
                            ElseIf i > ultlinmapa Then
                                resultado = MsgBox("Extintor n�o encontrado! Deseja tentar novamente?", vbYesNo, "Extintor n�o cadastrado")
                                If resultado = vbYes Then
                                    GoTo voltainput3:
                                Else
                                    MsgBox "Movimenta��o cancelada"
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
    Sheets("Movimenta��o").ListObjects("tbHistMov14").DataBodyRange.Calculate
    Info.ListObjects("tbHistMov").DataBodyRange.Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub




