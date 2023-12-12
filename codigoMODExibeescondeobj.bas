Attribute VB_Name = "MODExibeescondeobj"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub escondeobjfrmatualiza()
    ' Data......: 13/01/2021
    ' Descricao.: exibe botões dos forms em info
    '---------------------------------------------------------------------------------------
Public Sub escondeobjfrmatualiza()
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
    With Info
        .Shapes("btnExtAdd").Visible = msoFalse
        .Shapes("Scroll Bar 26").Visible = msoFalse
        .Shapes("Scroll Bar 48").Visible = msoFalse
        .Shapes("btnCancelarLocalAtual").Visible = msoFalse
        .Shapes("btnSalvaLocalAtual").Visible = msoFalse
        .Shapes("btnCancelarLocalNovo").Visible = msoFalse
        .Shapes("btnSalvaLocalNovo").Visible = msoFalse
        .Shapes("btnImprime").Visible = msoFalse
        .Shapes("btnCancelarNovoExt").Visible = msoFalse
        .Shapes("btnLocalAdd2").Visible = msoFalse
        .Shapes("btnSalvaAtualExt").Visible = msoFalse
        .Shapes("btnLocalAdd").Visible = msoFalse
        .Shapes("btnImprime").Height = 28.07
        .Shapes("btnextadd").Width = 37.38
        .Shapes("btnextadd").Height = 39.7
        
    End With
End Sub

Public Sub exibeobjfrmatualiza()
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
    With Info
        .Shapes("btnExtAdd").Visible = msoTrue
        .Shapes("Scroll Bar 26").Visible = msoTrue
        .Shapes("Scroll Bar 48").Visible = msoTrue
       
        .Shapes("btnImprime").Visible = msoTrue
        .Shapes("btnImprime").Height = 28.07
        .Shapes("btnextadd").Width = 37.38
        .Shapes("btnextadd").Height = 39.7

        .Shapes("btnSalvaAtualExt").Visible = msoTrue
        .Shapes("btnLocalAdd").Visible = msoTrue
    End With
End Sub

Public Sub exibeobjfrmnovoext()
    
    With Info

        .Shapes("btnSalvaNovoExt").Visible = msoTrue

        .Shapes("btnCancelarNovoExt").Visible = msoTrue
        .Shapes("btnLocalAdd2").Visible = msoTrue
        .Shapes("btnLocalAdd2").Top = 143.3667
    End With
End Sub

Public Sub exibeobjfrmnovolocal()
    
    With Info

        .Shapes("btnCancelarLocalNovo").Visible = msoTrue

        .Shapes("btnSalvaLocalNovo").Visible = msoTrue

    End With
End Sub

