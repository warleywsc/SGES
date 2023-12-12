Attribute VB_Name = "MODDimenBtnFrmAtual"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub dimBotoesFormAtualiza()
    ' Data......: 06/01/2021
    ' Descricao.: dimensiona botoes do form atualizaçaõ de extintores
    '---------------------------------------------------------------------------------------
Public Sub dimBotoesFormAtualiza()
    On Error GoTo TError
    With Info
        If .Range(" M12").RowHeight > 20 Then
            Info.Shapes("btnLocalAdd").Top = 163.53
        Else
            Info.Shapes("btnLocalAdd").Top = .Shapes("btnExtAdd").Top + 96.33
        End If
        .Shapes("btnLocalAdd").Height = 31.12
        .Shapes("btnLocalAdd").Width = 34.91
        .Shapes("btnextadd").Width = 37.38
        .Shapes("btnextadd").Height = 39.7
        '.Shapes("btnLocalAdd").Height = 36.6
    End With
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub
