Attribute VB_Name = "MODFormatatbhistmov"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Autor: Warley da Silva Conceiçao - Rotina: Sub formatatbhistmov()
    ' Data......: 10/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub formatatbhistmov()
    On Error GoTo TError
    Dim n     As Integer
    Info.Unprotect
    For n = 18 To 26
        If Info.Range("R" & n).Value <> vbNullString Then
            With Info
                .Unprotect
                .Range("S" & n).WrapText = True
                .Range("U" & n).WrapText = True
                .Range("u" & n).Rows.AutoFit
                .Range("s" & n).Rows.AutoFit
            
                .Protect
            End With
        End If
    Next
    Info.Protect
fim:
    Info.Protect
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

