Attribute VB_Name = "MODRestaurastatusserv"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub restaurastatusserv()
    ' Data......: 18/11/2020
    ' Descricao.: Restaura Status do Serviço anterior a data excluida.
    '---------------------------------------------------------------------------------------
Public Sub restaurastatusserv()

    Dim LINMAPAATUAL As Long
    With Info

        LINMAPAATUAL = 9
        Do Until MapaAtual.Cells(LINMAPAATUAL, 14).Row > MapaAtual.Cells(MapaAtual.Rows.Count, "G").End(xlUp).Row
                            
            If MapaAtual.Cells(LINMAPAATUAL, 14) = Info.Cells.Item(8, 9) Then
                            
                '######### TESTE #########
                                
                                                               
                                   
                If (MapaAtual.Cells(LINMAPAATUAL, 15) = "Brigada" Or MapaAtual.Cells(LINMAPAATUAL, 15) = "MAREFIRE") And MapaAtual.Cells(LINMAPAATUAL, 16) <> vbNullString And MapaAtual.Cells(LINMAPAATUAL, 10) = "Manutenção" Then

                    MapaAtual.Cells(LINMAPAATUAL, 17) = "Em Manutenção"
                                        
                ElseIf MapaAtual.Cells(LINMAPAATUAL, 16) = vbNullString Then

                    MapaAtual.Cells(LINMAPAATUAL, 17) = "PREENCHER DATA DE TESTE"
                                        

                ElseIf DateAdd("yyyy", 5, MapaAtual.Cells(LINMAPAATUAL, 16)) < Date Then

                    MapaAtual.Cells(LINMAPAATUAL, 17) = "SUBSTITUIR"
                                        
                ElseIf MapaAtual.Cells(LINMAPAATUAL, 16) - Date < 30 Then

                    MapaAtual.Cells(LINMAPAATUAL, 17) = "ATENÇÃO"
                                        
                Else

                    MapaAtual.Cells(LINMAPAATUAL, 17) = "TESTE EM DIA"
                                        
                End If
                              
                '######### RECARGA #########
                              
                If (MapaAtual.Cells(LINMAPAATUAL, 15) = "Brigada" Or MapaAtual.Cells(LINMAPAATUAL, 15) = "MAREFIRE") And MapaAtual.Cells(LINMAPAATUAL, 18) <> vbNullString And MapaAtual.Cells(LINMAPAATUAL, 10) = "Manutenção - Brigada" Then
    
                    MapaAtual.Cells(LINMAPAATUAL, 19) = "Em Manutenção"
                    
                    
                    ElseIf MapaAtual.Cells(LINMAPAATUAL, 12) = "1K" Then
                    MapaAtual.Cells(LINMAPAATUAL, 19) = "NÃO APLICÁVEL"
                                    
                ElseIf MapaAtual.Cells(LINMAPAATUAL, 18) = vbNullString Then
 
                    MapaAtual.Cells(LINMAPAATUAL, 19) = "PREENCHER DATA DE RECARGA"
                                    

                ElseIf DateAdd("yyyy", 5, MapaAtual.Cells(LINMAPAATUAL, 18)) < Date Then

                    MapaAtual.Cells(LINMAPAATUAL, 19) = "SUBSTITUIR"
                                    
                ElseIf MapaAtual.Cells(LINMAPAATUAL, 18) - Date < 30 Then

                    MapaAtual.Cells(LINMAPAATUAL, 19) = "ATENÇÃO"
                                    
                Else

                    MapaAtual.Cells(LINMAPAATUAL, 19) = "RECARGA EM DIA"
                                   
    
                End If
                                
                '######### PESAGEM #########
                           
                If (MapaAtual.Cells(LINMAPAATUAL, 15) = "Brigada" Or MapaAtual.Cells(LINMAPAATUAL, 15) = "MAREFIRE") And MapaAtual.Cells(LINMAPAATUAL, 20) <> vbNullString And _
                                                                                                                                                             MapaAtual.Cells(LINMAPAATUAL, 10) = "Manutenção" Then

                    MapaAtual.Cells(LINMAPAATUAL, 21) = "Em Manutenção"
                                   

                ElseIf MapaAtual.Cells(LINMAPAATUAL, 11) = "PQ" Or MapaAtual.Cells(LINMAPAATUAL, 11) = "AP" Or MapaAtual.Cells(LINMAPAATUAL, 11) = "EM" Or MapaAtual.Cells(LINMAPAATUAL, 11) = "EM" Or _
                                                                                                                                                                                               MapaAtual.Cells(LINMAPAATUAL, 11) = "AP" Or MapaAtual.Cells(LINMAPAATUAL, 11) = "PQ" Or MapaAtual.Cells(LINMAPAATUAL, 11) = "ES" Or MapaAtual.Cells(LINMAPAATUAL, 11) = "FM" Then

                    MapaAtual.Cells(LINMAPAATUAL, 21) = "NÃO APLICÁVEL"
                ElseIf MapaAtual.Cells(LINMAPAATUAL, 11) = "CO" Then
                                
                    If MapaAtual.Cells(LINMAPAATUAL, 20) = vbNullString Then

                        MapaAtual.Cells(LINMAPAATUAL, 21) = "PREENCHER DATA DE PESAGEM"
                        GoTo fim:
                    ElseIf MapaAtual.Cells(LINMAPAATUAL, 20) < Date Then

                        MapaAtual.Cells(LINMAPAATUAL, 21) = "PESAGEM VENCIDA"
                                    
                    ElseIf MapaAtual.Cells(LINMAPAATUAL, 20) - Date < 30 Then

                        MapaAtual.Cells(LINMAPAATUAL, 21) = "ATENÇÃO"
                                    

                    ElseIf MapaAtual.Cells(LINMAPAATUAL, 20) >= Date Then
                        MapaAtual.Cells(LINMAPAATUAL, 21) = "PESAGEM EM DIA"
fim:
                    End If

                End If
                                
                '######### SELO #########
                           
                If (MapaAtual.Cells(LINMAPAATUAL, 15) = "Brigada" Or MapaAtual.Cells(LINMAPAATUAL, 15) = "MAREFIRE") And MapaAtual.Cells(LINMAPAATUAL, 22) <> vbNullString And MapaAtual.Cells(LINMAPAATUAL, 10) = "Manutenção - Brigada" Then

                    MapaAtual.Cells(LINMAPAATUAL, 23) = "Em manutenção"
                ElseIf MapaAtual.Cells(LINMAPAATUAL, 22) = vbNullString Then

                    MapaAtual.Cells(LINMAPAATUAL, 23) = "PREENCHER DATA DE SELAGEM"

ElseIf MapaAtual.Cells(LINMAPAATUAL, 12) = "1K" Then
                    MapaAtual.Cells(LINMAPAATUAL, 23) = "NÃO APLICÁVEL"

                ElseIf MapaAtual.Cells(LINMAPAATUAL, 22) >= Date And MapaAtual.Cells(LINMAPAATUAL, 22) - Date < 30 Then

                    MapaAtual.Cells(LINMAPAATUAL, 23) = "ATENÇÃO"

                ElseIf MapaAtual.Cells(LINMAPAATUAL, 22) < Date Then

                    MapaAtual.Cells(LINMAPAATUAL, 23) = "SELO VENCIDO"

                Else

                    MapaAtual.Cells(LINMAPAATUAL, 23) = "SELO EM DIA"

                End If
                            
                '######### INSPEÇÃO #########
                            
                If (MapaAtual.Cells(LINMAPAATUAL, 15) = "Brigada" Or MapaAtual.Cells(LINMAPAATUAL, 15) = "MAREFIRE") And MapaAtual.Cells(LINMAPAATUAL, 24) <> vbNullString And MapaAtual.Cells(LINMAPAATUAL, 10) = "Manutenção - Brigada" Then

                    MapaAtual.Cells(LINMAPAATUAL, 27) = "Em Manutenção"

                ElseIf MapaAtual.Cells(LINMAPAATUAL, 24) = vbNullString Then

                    MapaAtual.Cells(LINMAPAATUAL, 25) = "PREENCHER DATA DE INSPEÇÃO"

                ElseIf MapaAtual.Cells(LINMAPAATUAL, 24) < Date Then

                    MapaAtual.Cells(LINMAPAATUAL, 25) = "INSPEÇÃO VENCIDA"

                ElseIf MapaAtual.Cells(LINMAPAATUAL, 24) >= Date And MapaAtual.Cells(LINMAPAATUAL, 24) - Date < 30 Then

                    MapaAtual.Cells(LINMAPAATUAL, 25) = "ATENÇÃO"

                Else

                    MapaAtual.Cells(LINMAPAATUAL, 25) = "INSPEÇÃO EM DIA"

                End If
                            
                '######### STATUS GERAL #######
            
                If InStr(MapaAtual.Cells(LINMAPAATUAL, 17), "VENCID") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "VENCID") > 0 _
                                                                                                                                Or InStr(MapaAtual.Cells(LINMAPAATUAL, 21), "VENCID") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 23), "VENCID") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "VENCID") > 0 Then

                    MapaAtual.Cells(LINMAPAATUAL, 29) = "Vencido"
                ElseIf InStr(MapaAtual.Cells(LINMAPAATUAL, 17), "SUBS") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "SUBS") > 0 _
                                                                                                                                Or InStr(MapaAtual.Cells(LINMAPAATUAL, 21), "SUBS") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 23), "SUBS") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "SUBS") > 0 Then

                    MapaAtual.Cells(LINMAPAATUAL, 29) = "Substituir"

                ElseIf InStr(MapaAtual.Cells(LINMAPAATUAL, 17), "ATEN") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "ATEN") > 0 _
                                                                                                                                Or InStr(MapaAtual.Cells(LINMAPAATUAL, 21), "ATEN") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 23), "ATEN") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "ATEN") > 0 Then

                    MapaAtual.Cells(LINMAPAATUAL, 29) = "Vencendo"

                ElseIf InStr(MapaAtual.Cells(LINMAPAATUAL, 17), "DIA") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "DIA") > 0 _
                                                                                                                              Or InStr(MapaAtual.Cells(LINMAPAATUAL, 21), "DIA") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 23), "DIA") > 0 Or InStr(MapaAtual.Cells(LINMAPAATUAL, 19), "DIA") > 0 Then

                    MapaAtual.Cells(LINMAPAATUAL, 29) = "Em dia"

                Else

                    MapaAtual.Cells(LINMAPAATUAL, 29) = "Conferir"

                End If
                              
                            
            End If
                            
            LINMAPAATUAL = LINMAPAATUAL + 1
        Loop
        'UPDATESTATUSGERAL
    End With


  
End Sub




