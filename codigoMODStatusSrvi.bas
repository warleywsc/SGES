Attribute VB_Name = "MODStatusSrvi"
'@Folder("SGES2020")
Option Explicit

'COLA AS COLUNAS DO ARRAY NA PLANILHA
'---------------------------------------------------------------------------------------
' Programador.....: Warley S Concei��o
' Contato...: warleywsc@gmail.com - Empresa: R&W Solu��es em TI - Rotina: Sub statusservico()
    ' Data......: 17/01/2021
    ' Descricao.: Atualiza Status dos servi�os no Mapaatual
    '---------------------------------------------------------------------------------------
Public Sub statusservico()
    StatusTeste
    StatusRecarga
    StatusPesagem
    statusselo
    statusinspecao

End Sub

Public Sub statusinspecao()
    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edif�cio]="Manuten��o");"Em Manuten��o";
    'SE([@[Pr�xima Inspe��o]]="";"PREENCHER DATA DE INSPE��O";
    'SE([@[Pr�xima Inspe��o]]<$AB$2;"INSPE��O VENCIDA";
    'SE(E([@[Pr�xima Inspe��o]]>=$AB$2;[@[Pr�xima Inspe��o]]-$AB$2<30);"ATEN��O";"INSPE��O EM DIA"))))
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
            .lblBarraEvolucao = "Atualizando Status de Inspe��o..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"
            If arr(i, 7) = "1K" And arr(i, 11) < Date Then
       
                arr(i, 20) = "SUBSTITUIR"
        

            ElseIf (arr(i, 10) = "Brigada" Or arr(i, 10) = "Manuten��o") And arr(i, 4) = "Manuten��o" Then
    
                arr(i, 20) = "Em Manuten��o"
    
            ElseIf arr(i, 19) = vbNullString Then
    
                arr(i, 20) = "PREENCHER DATA DE INSPE��O"
    
        
    
            ElseIf DateDiff("m", arr(i, 19), Date) = 0 Then
    
                arr(i, 20) = "ATEN��O"
            
            ElseIf DateDiff("m", arr(i, 19), Date) > 0 Then
    
                arr(i, 20) = "INSPE��O VENCIDA"
    
            Else
    
                arr(i, 20) = "INSPE��O EM DIA"
    
            End If
        Next i
    End With
    If i >= UBound(arr) Then Unload frmEvolucao
    MapaAtual.Range("y8:y" & UBound(arr, 1) + 7) = Application.Index(arr, , 20)
    Set arr = Nothing
End Sub

Public Sub statusselo()

    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edif�cio]="Manuten��o");"Em Manuten��o";
    'SE([@[Pr�xima Selagem]]="";"PREENCHER DATA DE SELAGEM";
    'SE([@[Pr�xima Selagem]]<$AB$2;"SELO VENCIDO";
    'SE(E([@[Pr�xima Selagem]]>=$AB$2;[@[Pr�xima Selagem]]-$AB$2<30);"ATEN��O";"SELO EM DIA"))))

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
                arr(i, 18) = "N�O APLIC�VEL" ' CILINDROS

            ElseIf arr(i, 7) <> "1K" And (arr(i, 10) = "Brigada" Or arr(i, 10) = "MAREFIRE") And arr(i, 4) = "Manuten��o" And arr(i, 17) <> vbNullString Then
    
                arr(i, 18) = "Em manuten��o"
            ElseIf arr(i, 7) <> "1K" And arr(i, 17) = vbNullString Then

                arr(i, 18) = "PREENCHER DATA DE SELAGEM"



            ElseIf arr(i, 7) <> "1K" And DateDiff("m", arr(i, 17), Date) = 0 Then

                arr(i, 18) = "ATEN��O"

            ElseIf arr(i, 7) <> "1K" And arr(i, 17) <> vbNullString And DateDiff("m", arr(i, 17), Date) > 0 Then

                arr(i, 18) = "SELO VENCIDO"
            ElseIf arr(i, 7) <> "1K" Then
            
            arr(i, 18) = "N�O APLIC�VEL"

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
    ''=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edif�cio]="Manuten��o");"Em Manuten��o";
    ''SE([@[Pr�ximo Teste]]="";"PREENCHER DATA DE TESTE";
    ''SE([@[Pr�ximo Teste]]<$AB$2;"TESTE VENCIDO";
    ''SE([@[Pr�ximo Teste]]-$AB$2<30;"ATEN��O";"TESTE EM DIA"))))
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

            ElseIf (arr(i, 10) = "Brigada" Or arr(i, 10) = "MAREFIRE") And arr(i, 4) = "Manuten��o" Then
    
                arr(i, 12) = "Em Manuten��o"
                'MapaAtual.Cells(i, 12) = arr(i, 12)
            ElseIf arr(i, 11) = vbNullString Then
 
                arr(i, 12) = "PREENCHER DATA DE TESTE"
                'MapaAtual.Cells(i, 12) = arr(i, 12)

            ElseIf DateDiff("m", arr(i, 11), Date) = 0 Then

                arr(i, 12) = "ATEN��O"
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
    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edif�cio]="Manuten��o");"Em Manuten��o";
    'SE([@[Pr�xima Recarga]]="";"PREENCHER DATA DE RECARGA";
    'SE([@[Pr�xima Recarga]]<$AB$2;"RECARGA VENCIDA";
    'SE(E([@[Pr�xima Recarga]]>=$AB$2;[@[Pr�xima Recarga]]-$AB$2<30);"ATEN��O";"RECARGA EM DIA"))))
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
                                                                         arr(i, 4) = "Manuten��o" Then
    
                arr(i, 14) = "Em Manuten��o"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
            ElseIf arr(i, 7) <> "1K" And arr(i, 13) = vbNullString Then
 
                arr(i, 14) = "PREENCHER DATA DE RECARGA"
                'MapaAtual.Cells(i, 14) = arr(i, 14)

       
            ElseIf arr(i, 7) <> "1K" And DateDiff("m", arr(i, 13), Date) = 0 Then

                arr(i, 14) = "ATEN��O"
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
    '=SE(([@Zona]="Manuten��o")+([@Zona]="Brigada")+([@Zona]="MAREFIRE")=1;"Em Manuten��o";
    'SE(([@Tipo]="PM")+([@Tipo]="AP")+([@Tipo]="EP")+([@Tipo]="EM")+([@Tipo]="AG")+([@Tipo]="PQ")+([@Tipo]="ES")=1;"N�O APLIC�VEL";
    'SE(([@Tipo]="FM")+([@Tipo]="CO")*([@[Pr�xima Pesagem]]="")=1;"CADASTRAR DATA DE PESAGEM";
    'SE(([@Tipo]="CO")*([@[Pr�xima Pesagem]]<$AB$2)=1;"PESAGEM VENCIDA";
    'SE(([@Tipo]="CO")*([@[Pr�xima Pesagem]]-$AB$2<30)=1;"ATEN��O";
    'SE(([@Tipo]="CO")*([@[Pr�xima Pesagem]]>=$AB$2)=1;"PESAGEM EM DIA";""))))))


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
                                                                         arr(i, 4) = "Manuten��o" Then
    
                arr(i, 16) = "Em Manuten��o"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
          
            ElseIf arr(i, 6) = "PQ" Or arr(i, 6) = "AP" Or arr(i, 6) = "EM" Or arr(i, 6) = "AP" Then
        
                arr(i, 16) = "N�O APLIC�VEL"
            ElseIf arr(i, 6) = "CO" And arr(i, 15) = vbNullString Then
        
                arr(i, 16) = "PREENCHER DATA DE PESAGEM"
                '         ElseIf arr(i, 6) = "CO" And arr(i, 15) < Date Then
            ElseIf arr(i, 6) = "CO" And DateDiff("m", arr(i, 15), Date) > 0 Then

                arr(i, 16) = "PESAGEM VENCIDA"
                'MapaAtual.Cells(i, 14) = arr(i, 14)
        
       
        
                '        ElseIf arr(i, 6) = "CO" And (Date - arr(i, 15) < 30 And Date - arr(i, 15) >= 1) Then
            ElseIf arr(i, 6) = "CO" And (DateDiff("m", arr(i, 15), Date) = 0) Then

                arr(i, 16) = "ATEN��O"
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
'    '=SE(E(OU([@Zona]="Brigada";[@Zona]="MAREFIRE");[@Edif�cio]="Manuten��o");"Em Manuten��o";
'    'SE([@[Pr�xima Inspe��o]]="";"PREENCHER DATA DE INSPE��O";
'    'SE([@[Pr�xima Inspe��o]]<$AB$2;"INSPE��O VENCIDA";
'    'SE(E([@[Pr�xima Inspe��o]]>=$AB$2;[@[Pr�xima Inspe��o]]-$AB$2<30);"ATEN��O";"INSPE��O EM DIA"))))
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
'.lblBarraEvolucao = "Atualizando Status de Inspe��o..."
'.lblBarraEvolucao.Width = percentual * largura
'.lblValor = Round(percentual * 100, 1) & "%"
'        arr(i, 20) = "SUBSTITUIR"
'        'arr(i, 22) = "SUBSTITUIR"
'
'        ElseIf (arr(i, 10) = "Brigada" Or arr(i, 10) = "Manuten��o") And arr(i, 4) = "Manuten��o" Then
'
'            arr(i, 20) = "Em Manuten��o"
'
'        ElseIf arr(i, 19) = "" Then
'
'            arr(i, 20) = "PREENCHER DATA DE INSPE��O"
'
'        ElseIf arr(i, 19) < Date Then
'
'            arr(i, 20) = "INSPE��O VENCIDA"
'
'        ElseIf arr(i, 19) >= Date And arr(i, 19) - Date < 30 Then
'
'            arr(i, 20) = "ATEN��O"
'
'        Else
'
'            arr(i, 20) = "INSPE��O EM DIA"
'
'        End If
'    Next i
'
'    MapaAtual.Range("y8:y" & UBound(arr, 1) + 7) = Application.Index(arr, , 20)
'Set arr = Nothing
'End Sub




