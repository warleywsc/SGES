Attribute VB_Name = "ModVerificaSerieVazio"
Global vazio As Long
Global dataserv1 As String, dataserv2 As String
Global serie As String




Sub serievazio()

Dim tbmov As Range, tbmapa As Range, tbserv As Range, cell As Range, rng As Range

Set tbmapa = MapaAtual.ListObjects(1).DataBodyRange
Set tbmov = Movimentacao.ListObjects(1).DataBodyRange
Set tbserv = Serviços.ListObjects(1).DataBodyRange

'vazio = 0

'verifica mapa

With tbmapa

    For Each cell In .ListObject.ListColumns(8).DataBodyRange.Cells
    
        If cell = "" Then
        vazio = vazio + 1
        Set rng = Range(cell.Address)
'        Debug.Print "Favor preencher o número de série em: " & cell.Address
'        MsgBox ("Favor preencher o número de série em: " & cell.Address)
        
'        MapaAtual.Activate
'        Selection.ListObject.ListRows(4470).Delete
        MapaAtual.Range(rng.Address).EntireRow.Delete
        vazio = 0
        GoTo fim:
        End If
    
    
    Next cell

End With

'verifica mov

With tbmov

    For Each cell In .ListObject.ListColumns(2).DataBodyRange.Cells
    
        If cell = "" Then
        vazio = vazio + 1
        Set rng = Range(cell.Address & ":" & cell.Address)
        Debug.Print "Favor preencher o número de série em: " & cell.Address
        MsgBox ("Favor preencher o número de série em: " & cell.Address)
        Movimentacao.Activate
        Movimentacao.Range(rng.Address).Select

        GoTo fim:
        
        End If
    
    
    Next cell

End With


'verifica serv

With tbserv

    For Each cell In .ListObject.ListColumns(2).DataBodyRange.Cells
    
        If cell = "" Then
        vazio = vazio + 1
        
        Set rng = Range(cell.Address & ":" & cell.Address)
        Debug.Print "Favor preencher o número de série em: " & cell.Address
        MsgBox ("Favor preencher o número de série em: " & cell.Address)
        Serviços.Activate
        Serviços.Range(rng.Address).Select

        GoTo fim:
        
        End If
    
    
    Next cell

End With


fim:
Set tbmapa = Nothing
Set tbmov = Nothing
Set tbserv = Nothing



End Sub
