Attribute VB_Name = "MODATUALIZAMAPA"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub AtualizamapaMOV()
    ' Data......: 03/12/2020
    ' Descricao.: Atualiza mapa atual com dados de movimentaçao
    '---------------------------------------------------------------------------------------
Public Sub AtualizamapaMOV()
    On Error GoTo TError



    Dim tbBase As Range
    Dim tbPesquisa As Range
    Set tbBase = Movimentacao.ListObjects("tbCadastroMovimentacao").DataBodyRange
    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim movarray() As Variant
    Dim mapaatualArray() As Variant
    Dim a     As Long
    Dim b     As Long
    ReDim Preserve movarray(1 To tbBase.Rows.Count, 1 To 8)
    ReDim Preserve mapaatualArray(1 To tbPesquisa.Rows.Count, 1 To 23)

    For a = 1 To UBound(movarray)
        For b = 1 To 8
            movarray(a, b) = tbBase.Cells(a, b)
        Next b
    Next a

    For a = 1 To tbPesquisa.Rows.Count
        For b = 1 To 23
            mapaatualArray(a, b) = tbPesquisa.Cells(a, b)
        Next b
    Next a

    Dim serie As String
    Dim Data  As Date
    Dim largura As Long
    Dim percentual As Double

    With frmEvolucao

        .Show
        largura = .lblBarraEvolucao.Width
        For a = 1 To UBound(mapaatualArray)
    
            DoEvents
            percentual = a / (UBound(mapaatualArray) + 10)
            .lblBarraEvolucao.Caption = "Atualizando Movimentação..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"

            Data = 0

            serie = mapaatualArray(a, 8)


            For b = 1 To UBound(movarray)


                If movarray(b, 2) = serie Then

                    If movarray(b, 1) > Data Then
                        If movarray(b, 3) = "Entrada" Then ' verifica se a movimentação é de entrada
                            Data = movarray(b, 1)

                            mapaatualArray(a, 2) = movarray(b, 7) 'área entrada
                            mapaatualArray(a, 4) = movarray(b, 6) 'local entrada
                            mapaatualArray(a, 9) = movarray(b, 8) 'zona entrada

                            'insere edifício

                            If InStr(movarray(b, 6), " - ") - 1 = -1 Then
                                mapaatualArray(a, 3) = movarray(b, 6)
                            Else
                                mapaatualArray(a, 3) = Left$(movarray(b, 6), InStr(movarray(b, 6), " - ") - 1) 'edificio entrada
                            End If

                            'fim da movimentação



                        End If
                    End If

                End If

            Next b

            Data = 0
        Next a
        If b >= UBound(movarray) Then Unload frmEvolucao
    End With

    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange = mapaatualArray
    
    'Set mapaatualArray = Nothing
    'Set movarray = Nothing
    Set tbBase = Nothing
    Set tbPesquisa = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub Atualizamapaserv()
    ' Data......: 03/12/2020
    ' Descricao.: Atualiza mapa atual com dados de serviços
    '---------------------------------------------------------------------------------------
Public Sub Atualizamapaserv()
    On Error GoTo TError

    Dim tbBase As Range
    Dim tbPesquisa As Range
    Set tbBase = Serviços.ListObjects("tbServicos").DataBodyRange
    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim servicosArray() As Variant
    Dim mapaatualArray() As Variant
    Dim a     As Long
    Dim b     As Long
    ReDim Preserve servicosArray(1 To tbBase.Rows.Count, 1 To 15)
    ReDim Preserve mapaatualArray(1 To tbPesquisa.Rows.Count, 1 To 23)

    For a = 1 To tbBase.Rows.Count
        For b = 1 To 15
            servicosArray(a, b) = tbBase.Cells(a, b)
        Next b
    Next a

    For a = 1 To tbPesquisa.Rows.Count
        For b = 1 To 23
            mapaatualArray(a, b) = tbPesquisa.Cells(a, b)
        Next b
    Next a

    Dim serie As String
    Dim dataT As Date
    Dim dataR As Date
    Dim dataP As Date
    Dim dataS As Date
    Dim dataI As Date
    Dim largura As Long
    Dim percentual As Double

    With frmEvolucao

        .Show
        largura = .lblBarraEvolucao.Width
        
        For a = 1 To UBound(mapaatualArray)
    
            DoEvents
            DoEvents
            percentual = a / (UBound(mapaatualArray) + 10)
            .lblBarraEvolucao.Caption = "Atualizando Serviços..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"

            dataT = 0
            dataR = 0
            dataP = 0
            dataS = 0
            dataI = 0

            serie = mapaatualArray(a, 8)


            For b = 1 To UBound(servicosArray)

                If mapaatualArray(a, 5) = "FM" Then
                    
                    mapaatualArray(a, 12) = mapaatualArray(a, 10) 'orientado pelo Sup. David Honório em reunião de trabalho
                End If
            
            
                If servicosArray(b, 2) = serie Then
'If servicosArray(b, 2) = "15209PQ1K" Then Stop
                    If servicosArray(b, 1) > dataT And servicosArray(b, 5) <> vbNullString Then

                        dataT = servicosArray(b, 1)
                        'If servicosArray(b, 5) <> "" And servicosArray(b, 5) > mapaatualArray(a, 10) Then
                        If servicosArray(b, 5) <> vbNullString Then
                            mapaatualArray(a, 10) = servicosArray(b, 5) 'PROX TESTE
                            mapaatualArray(a, 20) = servicosArray(b, 5) 'PROX PINTURA
                        End If
                    End If
                    
                    If InStr(servicosArray(b, 2), "1K") > 0 Then ' REGARGA 1K
                    
                    mapaatualArray(a, 12) = ""
                    
                    End If
                    
                    If servicosArray(b, 1) > dataR And servicosArray(b, 7) <> vbNullString Then

                        dataR = servicosArray(b, 1)
                        ' If servicosArray(b, 7) <> "" And servicosArray(b, 7) > mapaatualArray(a, 12) Then
                    
                        mapaatualArray(a, 12) = servicosArray(b, 7) 'PROX RECARGA
                    
                    End If
                    
                   
                    If servicosArray(b, 1) > dataP And (servicosArray(b, 3) = "CO") Then

                        dataP = servicosArray(b, 1)
                        '                    If servicosArray(b, 9) <> "" And servicosArray(b, 9) > mapaatualArray(a, 14) Then
                        If servicosArray(b, 9) <> vbNullString Then
                            mapaatualArray(a, 14) = servicosArray(b, 9) 'PROX PESAGEM
                        End If
                    End If
                    
                     If InStr(servicosArray(b, 2), "1K") > 0 Then ' SELO 1K
                    
                    mapaatualArray(a, 16) = ""
                    
                    End If

                    If servicosArray(b, 1) > dataS And servicosArray(b, 11) <> vbNullString Then

                        dataS = servicosArray(b, 1)
                        '                    If servicosArray(b, 11) <> "" And servicosArray(b, 11) > mapaatualArray(a, 16) Then
                        If servicosArray(b, 11) <> vbNullString Then
                            mapaatualArray(a, 16) = servicosArray(b, 11) 'PROX SELO
                        End If
                    End If

                    If servicosArray(b, 1) > dataI And servicosArray(b, 13) <> vbNullString Then

                        dataI = servicosArray(b, 1)
                        '                    If servicosArray(b, 13) <> "" And servicosArray(b, 13) > mapaatualArray(a, 18) Then
                        If servicosArray(b, 13) <> vbNullString Then
                            mapaatualArray(a, 18) = servicosArray(b, 13) 'PROX INSPECAO

                        End If
                    End If
               
                End If

            Next b

        Next a

        If b >= UBound(servicosArray) Then Unload frmEvolucao
    End With
    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange = mapaatualArray
   
    Set tbBase = Nothing
    Set tbPesquisa = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub AtualizamapaExt()
    On Error GoTo TError



    Dim tbBase As Range
    Dim tbPesquisa As Range
    Set tbBase = Extintores.ListObjects("tbExtintores").DataBodyRange
    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim extarray() As Variant
    Dim mapaatualArray() As Variant
    Dim a     As Long
    Dim b     As Long
    ReDim Preserve extarray(1 To tbBase.Rows.Count, 1 To 9)
    ReDim Preserve mapaatualArray(1 To tbPesquisa.Rows.Count, 1 To 23)

    For a = 1 To tbBase.Rows.Count
        For b = 1 To 9
            extarray(a, b) = tbBase.Cells(a, b)
        Next b
    Next a

    For a = 1 To tbPesquisa.Rows.Count
        For b = 1 To 23
            mapaatualArray(a, b) = tbPesquisa.Cells(a, b)
        Next b
    Next a

    Dim serie As String
    
    Dim largura As Long
    Dim percentual As Double
    With frmEvolucao
        .Show
        largura = .lblBarraEvolucao.Width
        For a = 1 To UBound(mapaatualArray)
            DoEvents
            percentual = a / (UBound(mapaatualArray) + 10)
            .lblBarraEvolucao.Caption = "Atualizando Extintores..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"

            serie = mapaatualArray(a, 8)


            For b = 1 To UBound(extarray)


                If extarray(b, 9) = serie Then


                    mapaatualArray(a, 1) = extarray(b, 5) 'Suporte
                    mapaatualArray(a, 5) = extarray(b, 2) 'Tipo
                    mapaatualArray(a, 6) = extarray(b, 3) 'CAPACIDADE
                    mapaatualArray(a, 7) = extarray(b, 4) 'FABRICAÇÃO
                    mapaatualArray(a, 21) = extarray(b, 6) 'OBS
                
                End If

            Next b

        Next a
    End With
    If a >= UBound(mapaatualArray) Then Unload frmEvolucao
    
    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange = mapaatualArray
    'Unload frmEvolucao
    'Set mapaatualArray = Nothing
    'Set extarray = Nothing
    Set tbBase = Nothing
    Set tbPesquisa = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub Atualizamapaserv2()
    On Error GoTo TError

    Dim tbBase As Range
    Dim tbPesquisa As Range
    Set tbBase = Serviços.ListObjects("tbServicos").DataBodyRange
    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim baseArray() As Variant
    Dim pesquisaArray() As Variant
    Dim a     As Long
    Dim b     As Long
    ReDim Preserve baseArray(1 To tbBase.Rows.Count, 1 To 15)
    ReDim Preserve pesquisaArray(1 To tbPesquisa.Rows.Count, 1 To 23)

    For a = 1 To tbBase.Rows.Count
        For b = 1 To 15
            baseArray(a, b) = tbBase.Cells(a, b)
        Next b
    Next a

    For a = 1 To tbPesquisa.Rows.Count
        For b = 1 To 23
            pesquisaArray(a, b) = tbPesquisa.Cells(a, b)
        Next b
    Next a

    Dim serie As String
    Dim dataT As Date
    Dim dataR As Date
    Dim dataP As Date
    Dim dataS As Date
    Dim dataI As Date

    For a = 1 To UBound(pesquisaArray)
        dataT = 0
        dataR = 0
        dataP = 0
        dataS = 0
        dataI = 0

        serie = pesquisaArray(a, 8)


        For b = 1 To UBound(baseArray)


            If baseArray(b, 2) = serie And baseArray(b, 2) = Info.Cells(8, 9) Then

                If baseArray(b, 1) > dataT And baseArray(b, 5) <> vbNullString Then

                    dataT = baseArray(b, 1)
                    If baseArray(b, 5) <> vbNullString And baseArray(b, 5) > pesquisaArray(a, 10) Then
                        pesquisaArray(a, 10) = baseArray(b, 5) 'PROX TESTE
                        pesquisaArray(a, 20) = baseArray(b, 5) 'PROX PINTURA
                    End If
                End If
                If baseArray(b, 1) > dataR And baseArray(b, 7) <> vbNullString Then

                    dataR = baseArray(b, 1)
                    If baseArray(b, 7) <> vbNullString And baseArray(b, 7) > pesquisaArray(a, 12) Then
                        pesquisaArray(a, 12) = baseArray(b, 7) 'PROX RECARGA
                    End If
                End If
                If baseArray(b, 1) > dataP And (baseArray(b, 3) = "CO" Or baseArray(b, 3) = "FM") Then

                    dataP = baseArray(b, 1)
                    If baseArray(b, 9) <> vbNullString And baseArray(b, 9) > pesquisaArray(a, 14) Then

                        pesquisaArray(a, 14) = baseArray(b, 9) 'PROX PESAGEM
                    End If
                End If

                If baseArray(b, 1) > dataS And baseArray(b, 11) <> vbNullString Then

                    dataS = baseArray(b, 1)
                    If baseArray(b, 11) <> vbNullString And baseArray(b, 11) > pesquisaArray(a, 16) Then
                        pesquisaArray(a, 16) = baseArray(b, 11) 'PROX SELO
                    End If
                End If

                If baseArray(b, 1) > dataI And baseArray(b, 13) <> vbNullString Then

                    dataI = baseArray(b, 1)
                    If baseArray(b, 13) <> vbNullString And baseArray(b, 13) > pesquisaArray(a, 18) Then
                        pesquisaArray(a, 18) = baseArray(b, 13) 'PROX INSPECAO

                    End If
                End If
                '                     If baseArray(b, 1) > data And baseArray(b, 15) <> "" Then
                '
                '                    data = baseArray(b, 1)
                '                    pesquisaArray(a, 20) = baseArray(b, 15) 'PROX PINTURA
                '
                '
                '                End If

            End If

        Next b

        'data = 0
    Next a


    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange = pesquisaArray
    'Set pesquisaArray = Nothing
    'Set baseArray = Nothing
    Set tbBase = Nothing
    Set tbPesquisa = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub AtualizamapaMOV2()
    On Error GoTo TError



    Dim tbBase As Range
    Dim tbPesquisa As Range
    Set tbBase = Movimentacao.ListObjects("tbCadastroMovimentacao").DataBodyRange
    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim movarray() As Variant
    Dim mapaatualArray() As Variant
    Dim a     As Long
    Dim b     As Long
    ReDim Preserve movarray(1 To tbBase.Rows.Count, 1 To 8)
    ReDim Preserve mapaatualArray(1 To tbPesquisa.Rows.Count, 1 To 23)

    For a = 1 To tbBase.Rows.Count
        For b = 1 To 8
            movarray(a, b) = tbBase.Cells(a, b)
        Next b
    Next a

    For a = 1 To tbPesquisa.Rows.Count
        For b = 1 To 23
            mapaatualArray(a, b) = tbPesquisa.Cells(a, b)
        Next b
    Next a

    Dim serie As String
    Dim serieinfo As String
    Dim Data  As Date

    For a = 1 To UBound(mapaatualArray)
        Data = 0

        serie = mapaatualArray(a, 8)
        serieinfo = Info.Range("I8").Value

        For b = 1 To UBound(movarray)


            If serieinfo = serie Then

                If movarray(b, 1) > Data Then
                    If movarray(b, 3) = "Entrada" Then ' verifica se a movimentação é de entrada
                        Data = movarray(b, 1)

                        mapaatualArray(a, 2) = movarray(b, 7) 'área entrada
                        mapaatualArray(a, 4) = movarray(b, 6) 'local entrada
                        mapaatualArray(a, 9) = movarray(b, 8) 'zona entrada

                        'insere edifício

                        If InStr(movarray(b, 6), " - ") - 1 = -1 Then
                            mapaatualArray(a, 3) = movarray(b, 6)
                        Else
                            mapaatualArray(a, 3) = Left$(movarray(b, 6), InStr(movarray(b, 6), " - ") - 1) 'edificio entrada
                        End If

                        'fim da movimentação



                    End If
                End If

            End If

        Next b

        Data = 0
    Next a


    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange = mapaatualArray
    'Set mapaatualArray = Nothing
    'Set movarray = Nothing
    Set tbBase = Nothing
    Set tbPesquisa = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub
























Public Sub Atualizamapaservindividual()
    On Error GoTo TError

    Dim tbBase As Range
    Dim tbPesquisa As Range
    Set tbBase = Serviços.ListObjects("tbServicos").DataBodyRange
    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim servicosArray() As Variant
    Dim mapaatualArray() As Variant
    Dim a     As Long
    Dim b     As Long
    ReDim Preserve servicosArray(1 To tbBase.Rows.Count, 1 To 15)
    ReDim Preserve mapaatualArray(1 To tbPesquisa.Rows.Count, 1 To 23)

    For a = 1 To tbBase.Rows.Count
        For b = 1 To 15
            servicosArray(a, b) = tbBase.Cells(a, b)
        Next b
    Next a

    For a = 1 To tbPesquisa.Rows.Count
        For b = 1 To 23
            mapaatualArray(a, b) = tbPesquisa.Cells(a, b)
        Next b
    Next a

    Dim serie As String
    Dim dataT As Date
    Dim dataR As Date
    Dim dataP As Date
    Dim dataS As Date
    Dim dataI As Date
    Dim largura As Long
    Dim percentual As Double
    Dim linserieb As Long
    Dim linseriea As Long
    With frmEvolucao

        .Show
        largura = .lblBarraEvolucao.Width
        
        For a = 1 To UBound(mapaatualArray)
    
            DoEvents
            DoEvents
            percentual = a / (UBound(mapaatualArray) + 10)
            .lblBarraEvolucao.Caption = "Atualizando Serviços..."
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"

            dataT = 0
            dataR = 0
            dataP = 0
            dataS = 0
            dataI = 0

            '        serie = mapaatualArray(a, 8)
            serie = Info.Cells(8, 9).Value


            For b = 1 To UBound(servicosArray)
        

           
            
                If servicosArray(b, 2) = serie Then
                    linserieb = b
                    linseriea = a
                    If mapaatualArray(a, 5) = "FM" Then
                    
                        mapaatualArray(a, 12) = mapaatualArray(a, 10) 'orientado pelo Sup. David Honório em reunião de trabalho
                    End If
            

                    If servicosArray(b, 1) > dataT And servicosArray(b, 5) <> vbNullString Then

                        dataT = servicosArray(b, 1)
                        'If servicosArray(b, 5) <> "" And servicosArray(b, 5) > mapaatualArray(a, 10) Then
                        If servicosArray(b, 5) <> vbNullString Then
                            mapaatualArray(a, 10) = servicosArray(b, 5) 'PROX TESTE
                            mapaatualArray(a, 20) = servicosArray(b, 5) 'PROX PINTURA
                        End If
                    End If
                    If servicosArray(b, 1) > dataR And servicosArray(b, 7) <> vbNullString Then

                        dataR = servicosArray(b, 1)
                        '                    If servicosArray(b, 7) <> "" And servicosArray(b, 7) > mapaatualArray(a, 12) Then
                    
                        mapaatualArray(a, 12) = servicosArray(b, 7) 'PROX RECARGA
                    
                    End If
                    If servicosArray(b, 1) > dataP And (servicosArray(b, 3) = "CO") Then

                        dataP = servicosArray(b, 1)
                        '                    If servicosArray(b, 9) <> "" And servicosArray(b, 9) > mapaatualArray(a, 14) Then
                        If servicosArray(b, 9) <> vbNullString Then
                            mapaatualArray(a, 14) = servicosArray(b, 9) 'PROX PESAGEM
                        End If
                    End If

                    If servicosArray(b, 1) > dataS And servicosArray(b, 11) <> vbNullString Then

                        dataS = servicosArray(b, 1)
                        '                    If servicosArray(b, 11) <> "" And servicosArray(b, 11) > mapaatualArray(a, 16) Then
                        If servicosArray(b, 11) <> vbNullString Then
                            mapaatualArray(a, 16) = servicosArray(b, 11) 'PROX SELO
                        End If
                    End If

                    If servicosArray(b, 1) > dataI And servicosArray(b, 13) <> vbNullString Then

                        dataI = servicosArray(b, 1)
                        '                    If servicosArray(b, 13) <> "" And servicosArray(b, 13) > mapaatualArray(a, 18) Then
                        If servicosArray(b, 13) <> vbNullString Then
                            mapaatualArray(a, 18) = servicosArray(b, 13) 'PROX INSPECAO

                        End If
                    End If
                    '                     If ServicosArray(b, 1) > data And ServicosArray(b, 15) <> "" Then
                    '
                    '                    data = ServicosArray(b, 1)
                    '                    mapaatualArray(a, 20) = ServicosArray(b, 15) 'PROX PINTURA
                    '
                    '
                    '                End If

                End If

            Next b

            'data = 0
        Next a

        If b >= UBound(servicosArray) Then Unload frmEvolucao
    End With
    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange = mapaatualArray
    'Set mapaatualArray = Nothing
    'Set ServicosArray = Nothing
    Set tbBase = Nothing
    Set tbPesquisa = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

