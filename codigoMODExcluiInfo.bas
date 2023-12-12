Attribute VB_Name = "MODExcluiInfo"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Autor: Warley da Silva Conceiçao - Rotina: Sub excluiserv()
    ' Data......: 15/11/2020
    ' Descricao.: Exclui Serviço a partir da tela info clicando com o botão direito
    ' na célula e excluir serviço
    '---------------------------------------------------------------------------------------
Public Sub excluiserv()

    Dim arrserv As Variant
    Dim arrservinfo As Variant
    Dim i     As Long
    Dim d     As String
    Dim a     As String

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    arrserv = Serviços.ListObjects("tbServicos").DataBodyRange
    arrservinfo = Info.ListObjects("tbHistServ").DataBodyRange

    i = 1
    With Info

        .Unprotect
        d = .Range("I8").Value & _
                               .Range("Q" & Selection.Row).Value & _
                               .Range("R" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("S" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("T" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("U" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("V" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("W" & .Range("Q" & Selection.Row).Row).Value

        Do Until i = UBound(arrserv) + 1
            a = arrserv(i, 2) & arrserv(i, 1) _
      & arrserv(i, 5) & arrserv(i, 7) & arrserv(i, 9) & arrserv(i, 11) _
      & arrserv(i, 13) & arrserv(i, 15)
            'With Serviços
            If d = a Then

                Serviços.ListObjects("tbServicos").ListRows(i).Range.Delete
                'Info.Range("tbHistServ").Calculate
                Calculate
                Atualizamapaserv
                'populafrmAtualExt
                restaurastatusserv
                'RestauraServ
            End If


            ' End With
            i = i + 1
        Loop
        
        formatatbhistmov
        dimbarra
        populafrmAtualExt
        .Protect
    End With
    
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Autor: Warley da Silva Conceiçao - Rotina: Sub excluimov()
    ' Data......: 16/11/2020
    ' Descricao.: Exclui movimentaçao a partir da tela info, clicando com o botão direito
    ' na célula e excluir movimentação
    '---------------------------------------------------------------------------------------
Public Sub excluimov()

    Dim arrMov As Variant
    Dim arrmapa As Variant
    Dim i     As Long
    Dim d     As String
    Dim a     As String
    Dim EDIF  As String
    Dim POSICAO As Long
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    arrMov = Movimentacao.ListObjects("tbCadastroMovimentacao").DataBodyRange
    arrmapa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange


    With Info


        d = .Range("I8").Value & _
                               .Range("Q" & Selection.Row).Value & _
                               .Range("R" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("S" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("T" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("U" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("V" & .Range("Q" & Selection.Row).Row).Value & _
                               .Range("W" & .Range("Q" & Selection.Row).Row).Value
        i = 1
        Do Until i = UBound(arrmapa) + 1
            If .Cells(8, 9) = arrmapa(i, 8) Then
                If .ListObjects("tbHistMov").DataBodyRange.Cells(Selection.Row - Selection.ListObject.Range.Row, 2) = "Entrada" And Selection.Row - Selection.ListObject.Range.Row < 3 Then

                    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 9) = "Brigada" 'Zona
                    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 4) = "Reserva Técnica"

                    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 2) = "1111"
                    'Movimentacao.ListObjects("tbCadastroMovimentacao").DataBodyRange.Cells(Movimentacao.ListObjects("tbCadastroMovimentacao").ListRows.Count + 1, 1) = Date
                    ' Movimentacao.ListObjects("tbCadastroMovimentacao").DataBodyRange.Cells(Movimentacao.ListObjects("tbCadastroMovimentacao").ListRows.Count + 1, 2) = "Saída"
                    
                    GoTo sair:

                ElseIf .Range("R" & Selection.Row).Value = "Entrada" And .Range("R" & Selection.Row).Row > 3 Then

                    '##############
                    arrmapa(i, 4) = .Range("R" & Selection.Row).Offset(-2, 3).Value 'restaura local anterior
                    arrmapa(i, 2) = .Range("R" & Selection.Row).Offset(-2, 4).Value 'restaura área anterior

                    arrmapa(i, 9) = .Range("R" & Selection.Row).Offset(-2, 5).Value 'restaura zona anterior

                    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 9) = arrmapa(i, 9)
                    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 4) = arrmapa(i, 4)
                    MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 2) = arrmapa(i, 2)

                    'restaura EDIFICIO

                    EDIF = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 4)
                    POSICAO = InStr(EDIF, " - ") - 1
                    If POSICAO = -1 Then
                        MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 3) = EDIF
                    Else
                        MapaAtual.ListObjects("tbMapaAtual").DataBodyRange.Cells(i, 3) = Left$(EDIF, POSICAO)
                    End If


                    GoTo sair:
                ElseIf .Range("R" & Selection.Row).Value = "Saída" Then
                    MsgBox "Por favor, selecione o último registro de " & "Entrada.", vbCritical, "Seleção Incorreta"
                    Exit Sub:


                End If
            End If

            i = i + 1
        Loop
sair:

        'EXCLUI MOVIMENTACAO
        i = 1
        Do Until i = UBound(arrMov) + 1
            a = arrMov(i, 2) & arrMov(i, 1) & arrMov(i, 3) _
      & arrMov(i, 4) & arrMov(i, 5) & arrMov(i, 6) & arrMov(i, 7) _
      & arrMov(i, 8)
            With Movimentacao
                If d = a Then
                   
                    .ListObjects("tbCadastroMovimentacao").ListRows(i).Range.Delete
                    .ListObjects("tbCadastroMovimentacao").ListRows(i - 1).Range.Delete
                    '                    .ListObjects("tbCadastroMovimentacao").DataBodyRange.Calculate
                    Exit Do
                End If


            End With
            i = i + 1
        Loop
       
        formatatbhistmov
        dimbarra
        Calculate
        populafrmAtualExt
        dimbtnsalvaext
    End With
    '     Info.Range("I8").Value = Info.Range("I8").Value
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub




