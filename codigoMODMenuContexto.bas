Attribute VB_Name = "MODMenuContexto"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Autor: Warley da Silva Conceiçao - Rotina: Sub addmenu()
    ' Data......: 14/11/2020
    ' Descricao.: Adiciona ou exclui Menus de Contexto em Info // parte do código está em Info (selection change) e em Thisworkbook
    '---------------------------------------------------------------------------------------
Public Sub addmenuserv()
    On Error GoTo TError
    
    Dim menucustom As CommandBarControl
    With Info.ListObjects("tbHistServ")
        Set menucustom = Application.CommandBars("List Range Popup").Controls.Add(msoControlButton, , , 1)
 
        With menucustom
    
            .Caption = "Excluir Serviço"
            .OnAction = "'" & ThisWorkbook.Name & "'!excluiServ"
            .BeginGroup = True
        End With
   
        Set menucustom = Nothing
    End With
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub addmenumov()
    On Error GoTo TError
    
    Dim menucustom As CommandBarControl
    With Info.ListObjects("tbHistMov")
        Set menucustom = Application.CommandBars("List Range Popup").Controls.Add(msoControlButton, , , 1)
 
        With menucustom
    
            .Caption = "Excluir Movimentação"
            .OnAction = "'" & ThisWorkbook.Name & "'!excluiMov"
            .BeginGroup = True
        End With
   
        Set menucustom = Nothing
    End With
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub deletemenumov()
    On Error GoTo TError
    Dim control As CommandBarControl
    With Info.Application.CommandBars("List Range Popup")
        For Each control In .Controls

            If .Controls(1).Caption = "Excluir Movimentação" Then
                .Controls(1).Delete
            Else
                Exit For:
            End If

        Next control
    End With
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub deletemenuserv()
    On Error GoTo TError
    Dim control As CommandBarControl
    With Info.Application.CommandBars("List Range Popup")
        For Each control In .Controls

            If .Controls(1).Caption = "Excluir Serviço" Then
                .Controls(1).Delete
            Else
                Exit For:
            End If

        Next control
    End With


fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub ShowShortcutMenuNames()
    Dim Row   As Long
    Dim cbar  As CommandBar
    Row = 1
    For Each cbar In Application.CommandBars
        If cbar.Type = 2 Then          'msoBarTypePopUp
            Cells(Row, 1) = cbar.Index
            Cells(Row, 2) = cbar.Name
            Row = Row + 1
        End If
    Next cbar
    Debug.Print
End Sub



Public Sub addmenuexcluiservmapa()
    On Error GoTo TError
    
    Dim menucustom As CommandBarControl
    With Info.Range("SERVICOSMAPA")
        Set menucustom = Application.CommandBars("Cell").Controls.Add(msoControlButton, , , 1)
 
        With menucustom
    
            .Caption = "Excluir Serviço"
            .OnAction = "'" & ThisWorkbook.Name & "'!excluiservIndividualmapaatual"
            .BeginGroup = True
        End With
    
        Set menucustom = Nothing
    End With
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub


Public Sub deletemenuexcluiservmapa()
    On Error GoTo TError
    Dim control As CommandBarControl
    With Info.Application.CommandBars("Cell")
        For Each control In .Controls

            If .Controls(1).Caption = "Excluir Serviço" Then
                .Controls(1).Delete
            Else
                Exit For:
            End If

        Next control
    End With


fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub
