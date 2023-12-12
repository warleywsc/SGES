Attribute VB_Name = "MODEncontraLinha"
Option Explicit


'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub ultlinha()
    ' Data......: 28/11/2020
    ' Descricao.: Seleciona a última linha da tabela
    '---------------------------------------------------------------------------------------
Public Sub ultlinhapesq()
    On Error GoTo TError

    '   Dim sht As Worksheet
    'Dim LastRow As Long
    '
    'Set sht = ActiveSheet
    '    sht.Cells(sht.Rows.Count, "N").End(xlDown).Select
    Range("n12").End(xlDown).Offset(0, -7).Select

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub ultlinha()
    On Error GoTo TError
    Dim sht   As Worksheet
    Dim lastRow As Long

    Set sht = ActiveSheet
    sht.Cells(sht.Rows.Count, "G").End(xlUp).Select

fim:
    Set sht = Nothing
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
    Set sht = Nothing
End Sub
Public Sub ultlinhaEXT()
    On Error GoTo TError
    Dim sht   As Worksheet
    Dim lastRow As Long

    Set sht = ActiveSheet
    sht.Cells(sht.Rows.Count, "G").End(xlUp).Select

fim:
    Set sht = Nothing
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
    Set sht = Nothing
End Sub
Public Sub ultlinhaSERV()
    On Error GoTo TError


    Dim sht   As Worksheet
    Dim lastRow As Long

    Set sht = ActiveSheet
    sht.Cells(sht.Rows.Count, "G").End(xlUp).Select


fim:
    Set sht = Nothing
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
    Set sht = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub primlinha()
    ' Data......: 28/11/2020
    ' Descricao.: Seleciona o primeiro registro da tabela
    '---------------------------------------------------------------------------------------
Public Sub primlinha()
    On Error GoTo TError

    Range("G8").Offset(1, 0).Select
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub


'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub primlinhaPESQUISA()
    ' Data......: 28/11/2020
    ' Descricao.:Seleciona o primeiro registro da tabela
    '---------------------------------------------------------------------------------------
Public Sub primlinhaPESQUISA()
    On Error GoTo TError


    Range("G11").Offset(1, 0).Select


fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

