Attribute VB_Name = "MODLimpafiltro"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub limpafiltrosmapaatual()
    ' Data......: 12/01/2021
    ' Descricao.:Desfaz filtros (Se houver). Ref.: http://dailydoseofexcel.com/archives/2014/03/02/how-do-you-know-if-a-listobject-is-filtered/
    '---------------------------------------------------------------------------------------
Public Sub limpafiltrosmapaatual()
    On Error GoTo TError

    Dim lo    As ListObject
    Dim af    As AutoFilter

    Set lo = MapaAtual.ListObjects(1)
    Set af = lo.AutoFilter

    If af Is Nothing Then
        Exit Sub
    Else
        af.ShowAllData
        lo.Sort. _
        SortFields.Clear
        lo.Sort. _
        SortFields.Add Key:=Range("tbmapaatual[[#All],[Série]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        lo.Sort.Apply
    End If

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub limpafiltrosmapaantigo()
    On Error GoTo TError

    Dim lo    As ListObject
    Dim af    As AutoFilter

    Set lo = MapaAntigo.ListObjects(1)
    Set af = lo.AutoFilter

    If af Is Nothing Then
        Exit Sub
    Else
        af.ShowAllData
        lo.Sort. _
        SortFields.Clear
        lo.Sort. _
        SortFields.Add Key:=Range("tbmapaantigo[[#All],[Série]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        lo.Sort.Apply
    End If
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub limpafiltrosservico()
    On Error GoTo TError

    Dim lo    As ListObject
    Dim af    As AutoFilter

    Set lo = Serviços.ListObjects("tbServicos")
    Set af = lo.AutoFilter

    If af Is Nothing Then
        Exit Sub
    Else
        af.ShowAllData
        lo.Sort. _
        SortFields.Clear
        lo.Sort. _
        SortFields.Add Key:=Range("tbServicos[[#All],[Data]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        lo.Sort.Apply
    End If

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub limpafiltrosmov()

    On Error GoTo TError

    Dim lo    As ListObject
    Dim af    As AutoFilter

    Set lo = Movimentacao.ListObjects("tbCadastroMovimentacao")
    Set af = lo.AutoFilter

    If af Is Nothing Then
        Exit Sub
    Else
        af.ShowAllData
        lo.Sort. _
        SortFields.Clear
        lo.Sort. _
        SortFields.Add Key:=Range("tbCadastroMovimentacao[[#All],[Data]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        lo.Sort.Apply
    End If

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub limpafiltrosext()
    On Error GoTo TError

    Dim lo    As ListObject
    Dim af    As AutoFilter

    Set lo = Extintores.ListObjects("tbExtintores")
    Set af = lo.AutoFilter

    If af Is Nothing Then
        Exit Sub
    Else
        af.ShowAllData
        lo.Sort. _
        SortFields.Clear
        lo.Sort. _
        SortFields.Add Key:=Range("tbExtintores[[#All],[Série Adaptado]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        lo.Sort.Apply
    End If

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub limpafiltrolocal()

    On Error GoTo TError

    Dim lo    As ListObject
    Dim af    As AutoFilter

    Set lo = locais.ListObjects("tbLocalNovo")
    Set af = lo.AutoFilter

    If af Is Nothing Then
        Exit Sub
    Else
        af.ShowAllData
        lo.Sort. _
        SortFields.Clear
        lo.Sort. _
        SortFields.Add Key:=Range("tbLocalNovo[[#All],[LocalxÁrea]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        lo.Sort.Apply
    End If

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub




