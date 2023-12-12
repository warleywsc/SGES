Attribute VB_Name = "MODRstauraserv"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Autor: Warley da Silva Conceiçao - Rotina: Sub RestauraServ()
    ' Data......: 10/11/2020
    ' Descricao.: Restaura data de serviço imediatamente anterior a data excluida
    '---------------------------------------------------------------------------------------
Public Sub RestauraServ()

    Dim lo1   As ListObject
    Dim lo2   As ListObject
    Dim linhalo1 As Long
    Dim linhalo2 As Long
    Dim i     As Integer
    Dim j     As Integer
    Dim LINHA As Long
    Dim LINMAPAATUAL As Long


    Set lo1 = Info.ListObjects("tbHistServ")
    Set lo2 = MapaAtual.ListObjects("tbMapaAtual")
    For LINMAPAATUAL = 1 To lo2.DataBodyRange.Rows.Count

        If lo2.DataBodyRange(LINMAPAATUAL, 8) = Info.Cells(8, 9) Then

            Exit For
        End If
    Next


    j = 10
    For i = 2 To 7

        For linhalo1 = lo1.DataBodyRange.Rows.Count To 1 Step -1

            If lo1.DataBodyRange(linhalo1, i) <> vbNullString Then

                Exit For

            End If

        Next linhalo1


        Select Case i

            Case Is = 2                'TESTE

                If linhalo1 = 0 Then
                    lo2.DataBodyRange(LINMAPAATUAL, j) = lo1.DataBodyRange(linhalo1 + 1, i)

                    If j >= 21 Then Exit Sub:
                    j = j + 2

                Else
                    lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("yyyy", 5, lo1.DataBodyRange(linhalo1, i))

                    If j >= 21 Then Exit Sub:
                    j = j + 2
                End If
            Case Is = 3                'RECARGA

                If linhalo1 = 0 Then
                    lo2.DataBodyRange(LINMAPAATUAL, j) = lo1.DataBodyRange(linhalo1 + 1, i)

                    If j >= 21 Then Exit Sub:
                    j = j + 2

                Else
                    lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("yyyy", 5, lo1.DataBodyRange(linhalo1, i))

                    If j >= 21 Then Exit Sub:
                    j = j + 2
                End If
            Case Is = 4                'PESAGEM

                If linhalo1 = 0 Then
                    lo2.DataBodyRange(LINMAPAATUAL, j) = lo1.DataBodyRange(linhalo1 + 1, i)

                    If j >= 21 Then Exit Sub:
                    j = j + 2

                Else
                    lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("m", 6, lo1.DataBodyRange(linhalo1, i))

                    If j >= 21 Then Exit Sub:
                    j = j + 2
                End If
            Case Is = 5                ' SELO

                If linhalo1 = 0 Then
                    lo2.DataBodyRange(LINMAPAATUAL, j) = lo1.DataBodyRange(linhalo1 + 1, i)

                    If j >= 21 Then Exit Sub:
                    j = j + 2

                Else
                    lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("yyyy", 1, lo1.DataBodyRange(linhalo1, i))

                    If j >= 21 Then Exit Sub:
                    j = j + 2
                End If
            Case Is = 6                ' INSPEÇÃO

                If linhalo1 = 0 Then
                    lo2.DataBodyRange(LINMAPAATUAL, j) = lo1.DataBodyRange(linhalo1 + 1, i)

                    If j >= 21 Then Exit Sub:
                    j = j + 2

                Else
                    If Info.Range("m8") = "CO" Then
                        lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("m", 6, lo1.DataBodyRange(linhalo1, i))

                        If j >= 21 Then Exit Sub:
                        j = j + 2

                    ElseIf Info.Range("m8") = "FM" Then
                        lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("m", 1, lo1.DataBodyRange(linhalo1, i))
                        If j >= 21 Then Exit Sub:
                        j = j + 2
                    Else
                        lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("yyyy", 1, lo1.DataBodyRange(linhalo1, i))
                        If j >= 21 Then Exit Sub:
                        j = j + 2

                    End If
                End If
            Case Is = 7                ' pintura

                If linhalo1 = 0 Then
                    lo2.DataBodyRange(LINMAPAATUAL, j) = lo1.DataBodyRange(linhalo1 + 1, i)

                    If j >= 21 Then Exit Sub:
                    j = j + 2

                Else
                    lo2.DataBodyRange(LINMAPAATUAL, j) = DateAdd("yyyy", 5, lo1.DataBodyRange(linhalo1, i))

                    If j >= 21 Then Exit Sub:
                    j = j + 2
                End If

        End Select



    Next i
    lo2.DataBodyRange(LINMAPAATUAL, 20) = lo2.DataBodyRange(LINMAPAATUAL, 10)
    Set lo1 = Nothing
    Set lo2 = Nothing
    populafrmAtualExt
    restaurastatusserv
End Sub

