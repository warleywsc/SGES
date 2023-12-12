Attribute VB_Name = "MODBarradeprogresso"
'@Folder("SGES2020")
Option Explicit

Public Sub barraevolucao()
    Dim total As Long
    Dim X     As Long
    Dim largura As Long
    Dim percentual As Double

    total = 100000
    With frmEvolucao

        .Show
        largura = .lblBarraEvolucao.Width
        For X = 1 To total
            DoEvents
            percentual = X / total
            .lblBarraEvolucao.Width = percentual * largura
            .lblValor = Round(percentual * 100, 1) & "%"


        Next X
    End With
    Unload frmEvolucao
End Sub
