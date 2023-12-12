Attribute VB_Name = "MODBarrainfo"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Sub Barraderolagem26_Alteração()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub Barraderolagem26_Alteração()

    Info.Unprotect

    Range("Q19").CurrentRegion.Calculate

    Application.ScreenUpdating = False
    
    Application.EnableEvents = False
    With Info
        ' .Shapes("btnSalvaAtualExt").Top = Info.Range("tbHistMov").Height + Info.Range("tbHistMov").Top + 5
        '.Shapes("Scroll Bar 26").Left = 650.4 + ActiveSheet.Shapes.Range(Array("btnocultarmenu")).Left
        '.Shapes("Scroll Bar 48").Left = .Shapes("Scroll Bar 26").Left
        ' .Shapes("Scroll Bar 26").Height = Info.Range("tbHistMov").Height
        '.Shapes("Scroll Bar 26").Top = Info.Range("tbHistMov").Top
    End With
    dimbarra
    formatatbhistmov
    dimbtnsalvaext
    Info.Protect
    Application.ScreenUpdating = True
    
    Application.EnableEvents = True
End Sub

Public Sub Barraderolagem48_Alteração()
    Info.Unprotect
    Range("q9").CurrentRegion.Calculate
    Info.Protect
End Sub

