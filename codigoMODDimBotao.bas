Attribute VB_Name = "MODDimBotao"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Rotina: Sub dimbtnsalvaext()
    ' Data......: 16/11/2020
    ' Descricao.: Dimensiona botões do form info
    '---------------------------------------------------------------------------------------
Public Sub dimbtnsalvaext()
    With Info

        .Shapes("btnSalvaAtualExt").Visible = msoCTrue
        .Shapes("btnSalvaAtualExt").Top = .Range(" M27").Top
        .Shapes("btnSalvaAtualExt").Height = 26.02
        .Shapes("btnSalvaAtualExt").Width = 93.5
    End With
End Sub

Public Sub dimbtnsalvaextnovo()
    With Info

        .Shapes("btnSalvaNovoExt").Visible = msoCTrue

        .Shapes("btnSalvaNovoExt").Height = 26.02
        .Shapes("btnSalvaNovoExt").Width = 93.5
    End With
End Sub

Public Sub dimbtncancelnovoext()
    With Info

        .Shapes("btnCancelarNovoExt").Visible = msoCTrue

        .Shapes("btnCancelarNovoExt").Height = 26.02
        .Shapes("btnCancelarNovoExt").Width = 93.5
    End With
End Sub

Public Sub dimbtnsalvalocalatual()
    With Info

        .Shapes("btnSalvaLocalAtual").Visible = msoCTrue
        .Shapes("btnSalvaLocalAtual").Top = Cells(107, 13).Top
        .Shapes("btnSalvaLocalAtual").Height = 26.02
        .Shapes("btnSalvaLocalAtual").Width = 93.5
    End With
End Sub

Public Sub dimbtnCancelarLocalAtual()
    With Info
        
        .Shapes("btnCancelarLocalAtual").Visible = msoCTrue
        
        .Shapes("btnCancelarLocalAtual").Height = 26.02
        .Shapes("btnCancelarLocalAtual").Width = 93.5
        .Shapes("btnCancelarLocalAtual").Top = Cells(107, 13).Top
    End With
End Sub

Public Sub dimbtnSalvaLocalNovo()
    With Info

        .Shapes("btnSalvaLocalNovo").Visible = msoCTrue
        
        .Shapes("btnSalvaLocalNovo").Height = 26.02
        .Shapes("btnSalvaLocalNovo").Width = 93.5
        .Shapes("btnSalvaLocalNovo").Top = Cells(71, 13).Top
    End With
End Sub

Public Sub dimbtnCancelarLocalNovo()
    With Info

        .Shapes("btnCancelarLocalNovo").Visible = msoCTrue
        
        .Shapes("btnCancelarLocalNovo").Height = 26.02
        .Shapes("btnCancelarLocalNovo").Width = 93.5
        .Shapes("btnCancelarLocalNovo").Top = Cells(71, 13).Top
    End With
End Sub

