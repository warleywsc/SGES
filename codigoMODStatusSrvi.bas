Attribute VB_Name = "MODStatusSrvi"
'@Folder("SGES2020")
Option Explicit

'COLA AS COLUNAS DO ARRAY NA PLANILHA
'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Empresa: R&W Soluções em TI - Rotina: Sub statusservico()
    ' Data......: 17/01/2021
    ' Descricao.: Atualiza Status dos serviços no Mapaatual
    '---------------------------------------------------------------------------------------
Public Sub statusservico()
    StatusTeste
    StatusRecarga
    StatusPesagem
    statusselo
    statusinspecao

End Sub

Public Sub statusinspecao()
    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edifício]="Manutenção");"Em Manutenção";
    'SE([@[Próxima Inspeção]]="";"PREENCHER DATA DE INSPEÇÃO";
    'SE([@[Próxima Inspeção]]<$AB$2;"INSPEÇÃO VENCIDA";
    'SE(E([@[Próxima Inspeção]]>=$AB$2;[@[Próxima Inspeção]]-$AB$2<30);"ATENÇÃO";"INSPEÇÃO EM DIA"))))
    Dim arr   As Variant, i As Long
    
    Dim largura As Long
    Dim percentual As Double

    arr = MapaAtual.Range("N8").CurrentRegion
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width
        For i = LBound(arr, 1) + 1 To UBound(arr, 1)
    
            DoEvents
            percentual = i / UBound(arr, 1)
            .lblBarraEvolucao.TextAlign = fmTextAlignRight
            .lblBarraEvolucao = "Atualizando Status de Inspeção..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"
            If arr(i, 7) = "1K" And arr(i, 11) < Date Then
       
                arr(i, 20) = "SUBSTITUIR"
        

            ElseIf (arr(i, 10) = "Brigada" Or arr(i, 10) = "Manutenção") And arr(i, 4) = "Manutenção" Then
    
                arr(i, 20) = "Em Manutenção"
    
            ElseIf arr(i, 19) = vbNullString Then
    
                arr(i, 20) = "PREENCHER DATA DE INSPEÇÃO"
    
        
    
            ElseIf DateDiff("m", arr(i, 19), Date) = 0 Then
    
                arr(i, 20) = "ATENÇÃO"
            
            ElseIf DateDiff("m", arr(i, 19), Date) > 0 Then
    
                arr(i, 20) = "INSPEÇÃO VENCIDA"
    
            Else
    
                arr(i, 20) = "INSPEÇÃO EM DIA"
    
            End If
        Next i
    End With
    If i >= UBound(arr) Then Unload frmEvolucao
    MapaAtual.Range("y8:y" & UBound(arr, 1) + 7) = Application.Index(arr, , 20)
    Set arr = Nothing
End Sub

Public Sub statusselo()

    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edifício]="Manutenção");"Em Manutenção";
    'SE([@[Próxima Selagem]]="";"PREENCHER DATA DE SELAGEM";
    'SE([@[Próxima Selagem]]<$AB$2;"SELO VENCIDO";
    'SE(E([@[Próxima Selagem]]>=$AB$2;[@[Próxima Selagem]]-$AB$2<30);"ATENÇÃO";"SELO EM DIA"))))

    Dim arr   As Variant
    Dim i     As Long

    Dim largura As Long
    Dim percentual As Double

    arr = MapaAtual.Range("N8").CurrentRegion
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width

        For i = LBound(arr, 1) + 1 To UBound(arr, 1)
            DoEvents
            percentual = i / UBound(arr, 1)
            .lblBarraEvolucao.TextAlign = fmTextAlignRight
            .lblBarraEvolucao = "Atualizando Status do Selo..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"
            If arr(i, 7) = "1K" And DateAdd("yyyy", 5, arr(i, 8)) < Date Then
       
                arr(i, 18) = "SUBSTITUIR"
                'arr(i, 22) = "SUBSTITUIR"

            ElseIf arr(i, 6) = "CO" And (arr(i, 7) = "34K" Or arr(i, 7) = "45K") Then
                arr(i, 18) = "NÃO APLICÁVEL" ' CILINDROS

            ElseIf arr(i, 7) <> "1K" And (arr(i, 10) = "Brigada" Or arr(i, 10) = "MAREFIRE") And arr(i, 4) = "Manutenção" And arr(i, 17) <> vbNullString Then
    
                arr(i, 18) = "Em manutenção"
            ElseIf arr(i, 7) <> "1K" And arr(i, 17) = vbNullString Then

                arr(i, 18) = "PREENCHER DATA DE SELAGEM"



            ElseIf arr(i, 7) <> "1K" And DateDiff("m", arr(i, 17), Date) = 0 Then

                arr(i, 18) = "ATENÇÃO"

            ElseIf arr(i, 7) <> "1K" And arr(i, 17) <> vbNullString And DateDiff("m", arr(i, 17), Date) > 0 Then

                arr(i, 18) = "SELO VENCIDO"
            ElseIf arr(i, 7) <> "1K" Then
            
            arr(i, 18) = "NÃO APLICÁVEL"

            Else

                arr(i, 18) = "SELO EM DIA"

            End If

        Next i
    End With
    If i >= UBound(arr) Then Unload frmEvolucao
    MapaAtual.Range("w8:w" & UBound(arr, 1) + 7) = Application.Index(arr, , 18)
    Set arr = Nothing
End Sub

Public Sub StatusTeste()
    ''=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edifício]="Manutenção");"Em Manutenção";
    ''SE([@[Próximo Teste]]="";"PREENCHER DATA DE TESTE";
    ''SE([@[Próximo Teste]]<$AB$2;"TESTE VENCIDO";
    ''SE([@[Próximo Teste]]-$AB$2<30;"ATENÇÃO";"TESTE EM DIA"))))
    Dim arr   As Variant, i As Long

    Dim largura As Long
    Dim percentual As Double

    arr = MapaAtual.Range("N8").CurrentRegion
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width

        For i = LBound(arr, 1) + 1 To UBound(arr, 1)
    
            DoEvents
            percentual = i / UBound(arr, 1)
            .lblBarraEvolucao.TextAlign = fmTextAlignRight
            .lblBarraEvolucao = "Atualizando Status de Teste..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"
            ' STATUS TESTE
            If arr(i, 7) = "1K" And DateAdd("yyyy", 5, arr(i, 8)) < Date Then
       
                arr(i, 12) = "SUBSTITUIR"
                'arr(i, 22) = "SUBSTITUIR"

            ElseIf (arr(i, 10) = "Brigada" Or arr(i, 10) = "MAREFIRE") And arr(i, 4) = "Manutenção" Then
    
                arr(i, 12) = "Em Manutenção"
                'MapaAtual.Cells(i, 12) = arr(i, 12)
            ElseIf arr(i, 11) = vbNullString Then
 
                arr(i, 12) = "PREENCHER DATA DE TESTE"
                'MapaAtual.Cells(i, 12) = arr(i, 12)

            ElseIf DateDiff("m", arr(i, 11), Date) = 0 Then

                arr(i, 12) = "ATENÇÃO"
                ' MapaAtual.Cells(i, 12) = arr(i, 12)
            
    
            ElseIf DateDiff("m", arr(i, 11), Date) > 0 Then

                arr(i, 12) = "TESTE VENCIDO"
                'MapaAtual.Cells(i, 12) = arr(i, 12)
            Else

                arr(i, 12) = "TESTE EM DIA"
                'MapaAtual.Cells(i, 12) = arr(i, 12)
    
            End If
        Next i
    End With
    If i >= UBound(arr) Then Unload frmEvolucao
    MapaAtual.Range("q8:q" & UBound(arr, 1) + 7) = Application.Index(arr, , 12)
    Set arr = Nothing
End Sub

Public Sub StatusRecarga()
    'STATUS RECARGA
    '
    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edifício]="Manutenção");"Em Manutenção";
    'SE([@[Próxima Recarga]]="";"PREENCHER DATA DE RECARGA";
    'SE([@[Próxima Recarga]]<$AB$2;"RECARGA VENCIDA";
    'SE(E([@[Próxima Recarga]]>=$AB$2;[@[Próxima Recarga]]-$AB$2<30);"ATENÇÃO";"RECARGA EM DIA"))))
    Dim arr   As Variant, i As Long

    Dim largura As Long
    Dim percentual As Double

    arr = MapaAtual.Range("N8").CurrentRegion
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width

        For i = LBound(arr, 1) + 1 To UBound(arr, 1)
            DoEvents
            percentual = i / UBound(arr, 1)
            .lblBarraEvolucao.TextAlign = fmTextAlignRight
            .lblBarraEvolucao = "Atualizando Status de Recarga..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"
            
'            If arr(i, 9) = "15209PQ1K" Then Stop
            If arr(i, 7) = "1K" And DateAdd("yyyy", 5, arr(i, 8)) < Date Then
       
                arr(i, 14) = "SUBSTITUIR"
                'arr(i, 22) = "SUBSTITUIR"

            ElseIf arr(i, 7) <> "1K" And (arr(i, 10) = "Brigada" Or arr(i, 10) = "MAREFIRE") And _
                                                                         arr(i, 4) = "Manutenção" Then
    
                arr(i, 14) = "Em Manutenção"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
            ElseIf arr(i, 7) <> "1K" And arr(i, 13) = vbNullString Then
 
                arr(i, 14) = "PREENCHER DATA DE RECARGA"
                'MapaAtual.Cells(i, 14) = arr(i, 14)

       
            ElseIf arr(i, 7) <> "1K" And DateDiff("m", arr(i, 13), Date) = 0 Then

                arr(i, 14) = "ATENÇÃO"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
            
            ElseIf arr(i, 7) <> "1K" And DateDiff("m", arr(i, 13), Date) > 0 Then

                arr(i, 14) = "RECARGA VENCIDA"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
            Else

                arr(i, 14) = "RECARGA EM DIA"
                ' MapaAtual.Cells(i, 14) = arr(i, 14)
    
            End If
    
    
        Next i
    
    
    End With
    If i >= UBound(arr) Then Unload frmEvolucao
    'COLA AS COLUNAS DO ARRAY NA PLANILHA

    MapaAtual.Range("s8:s" & UBound(arr, 1) + 7) = Application.Index(arr, , 14)
    Set arr = Nothing

End Sub

Public Sub StatusPesagem()
    '=SE(([@Zona]="Manutenção")+([@Zona]="Brigada")+([@Zona]="MAREFIRE")=1;"Em Manutenção";
    'SE(([@Tipo]="PM")+([@Tipo]="AP")+([@Tipo]="EP")+([@Tipo]="EM")+([@Tipo]="AG")+([@Tipo]="PQ")+([@Tipo]="ES")=1;"NÃO APLICÁVEL";
    'SE(([@Tipo]="FM")+([@Tipo]="CO")*([@[Próxima Pesagem]]="")=1;"CADASTRAR DATA DE PESAGEM";
    'SE(([@Tipo]="CO")*([@[Próxima Pesagem]]<$AB$2)=1;"PESAGEM VENCIDA";
    'SE(([@Tipo]="CO")*([@[Próxima Pesagem]]-$AB$2<30)=1;"ATENÇÃO";
    'SE(([@Tipo]="CO")*([@[Próxima Pesagem]]>=$AB$2)=1;"PESAGEM EM DIA";""))))))


    Dim arr   As Variant, i As Long

    Dim largura As Long
    Dim percentual As Double

    arr = MapaAtual.Range("N8").CurrentRegion
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width

        For i = LBound(arr, 1) + 1 To UBound(arr, 1)
            DoEvents
            percentual = i / UBound(arr, 1)
            .lblBarraEvolucao.TextAlign = fmTextAlignRight
            .lblBarraEvolucao = "Atualizando Status de Pesagem..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"
            If arr(i, 7) = "1K" And DateAdd("yyyy", 5, arr(i, 8)) < Date Then
       
                arr(i, 16) = "SUBSTITUIR"
                'arr(i, 22) = "SUBSTITUIR"

            ElseIf (arr(i, 10) = "Brigada" Or arr(i, 10) = "MAREFIRE") And _
                                                                         arr(i, 4) = "Manutenção" Then
    
                arr(i, 16) = "Em Manutenção"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
          
            ElseIf arr(i, 6) = "PQ" Or arr(i, 6) = "AP" Or arr(i, 6) = "EM" Or arr(i, 6) = "AP" Then
        
                arr(i, 16) = "NÃO APLICÁVEL"
            ElseIf arr(i, 6) = "CO" And arr(i, 15) = vbNullString Then
        
                arr(i, 16) = "PREENCHER DATA DE PESAGEM"
                '         ElseIf arr(i, 6) = "CO" And arr(i, 15) < Date Then
            ElseIf arr(i, 6) = "CO" And DateDiff("m", arr(i, 15), Date) > 0 Then

                arr(i, 16) = "PESAGEM VENCIDA"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
        
       
        
                '        ElseIf arr(i, 6) = "CO" And (Date - arr(i, 15) < 30 And Date - arr(i, 15) >= 1) Then
            ElseIf arr(i, 6) = "CO" And (DateDiff("m", arr(i, 15), Date) = 0) Then

                arr(i, 16) = "ATENÇÃO"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
            
                '    ElseIf arr(i, 6) = "CO" And arr(i, 15) >= Date Then
            ElseIf arr(i, 6) = "CO" And DateDiff("m", arr(i, 15), Date) < 0 Then
                arr(i, 16) = "PESAGEM EM DIA"
                ' MapaAtual.Cells(i, 14) = arr(i, 14)
    
            End If
    
    
        Next i
    End With
    If i >= UBound(arr) Then Unload frmEvolucao

    MapaAtual.Range("u8:u" & UBound(arr, 1) + 7) = Application.Index(arr, , 16)
    Set arr = Nothing
End Sub

'Sub STATUSGERAL()
'    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edifício]="Manutenção");"Em Manutenção";
'    'SE([@[Próxima Inspeção]]="";"PREENCHER DATA DE INSPEÇÃO";
'    'SE([@[Próxima Inspeção]]<$AB$2;"INSPEÇÃO VENCIDA";
'    'SE(E([@[Próxima Inspeção]]>=$AB$2;[@[Próxima Inspeção]]-$AB$2<30);"ATENÇÃO";"INSPEÇÃO EM DIA"))))
'    Dim arr   As Variant, i As Long
'     Dim largura As Long
'Dim percentual As Double
'
'arr = MapaAtual.Range("N8").CurrentRegion
'With frmEvolucao
'.Show
'largura = .lblBarraEvolucao.Width
'
'    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
'
'        If arr(i, 7) = "1K" And arr(i, 11) < Date Then
'       DoEvents
'percentual = i / UBound(arr, 1)
'.lblBarraEvolucao = "Atualizando Status de Inspeção..."
'.lblBarraEvolucao.Width = percentual * largura
'.lblValor = Round(percentual * 100, 1) & "%"
'        arr(i, 20) = "SUBSTITUIR"
'        'arr(i, 22) = "SUBSTITUIR"
'
'        ElseIf (arr(i, 10) = "Brigada" Or arr(i, 10) = "Manutenção") And arr(i, 4) = "Manutenção" Then
'
'            arr(i, 20) = "Em Manutenção"
'
'        ElseIf arr(i, 19) = "" Then
'
'            arr(i, 20) = "PREENCHER DATA DE INSPEÇÃO"
'
'        ElseIf arr(i, 19) < Date Then
'
'            arr(i, 20) = "INSPEÇÃO VENCIDA"
'
'        ElseIf arr(i, 19) >= Date And arr(i, 19) - Date < 30 Then
'
'            arr(i, 20) = "ATENÇÃO"
'
'        Else
'
'            arr(i, 20) = "INSPEÇÃO EM DIA"
'
'        End If
'    Next i
'
'    MapaAtual.Range("y8:y" & UBound(arr, 1) + 7) = Application.Index(arr, , 20)
'Set arr = Nothing
'End Sub




