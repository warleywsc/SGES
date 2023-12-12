Attribute VB_Name = "MODLogUndo"
Option Explicit
'@Folder("SGES2020")

Public Sub AddToLogInfo()
    Dim ActRow As Variant
    Dim LogRow As Long
    With Info
        ActRow = .Range("F10").Value   'You can also use ActiveCell.Row
        LogRow = Logs.Range("G1000").End(xlUp).Row + 1 'First Available Log Row
        Logs.Range("G" & LogRow).Value = Now 'Add Current Date & Time
        
        Logs.Range("H" & LogRow).Value = Environ$("UserName") 'Excel Application User Name
        Logs.Range("I" & LogRow).Value = "Célula alterada" 'Change Type
    
        Logs.Range("J" & LogRow).Value = .Range("A5").Value 'nome da planilha
        Logs.Range("K" & LogRow).Value = .Range("A2").Value 'Cell Address
        Logs.Range("L" & LogRow).Value = .Range("A4").Value 'Previous Value
        Logs.Range("M" & LogRow).Value = .Range(.Range("A2").Value).Value
    
        
    End With

End Sub

Public Sub Log_PlaceUndoButton()
    With Logs.Shapes("UndoBtn")
        .Visible = msoCTrue
        .Left = Logs.Range("K" & ActiveCell.Row).Left
        .Top = Logs.Range("K" & ActiveCell.Row).Top
    End With
End Sub

Public Sub Log_Undo()
    Dim nomePlanilha As String
    Dim LogRow As Long
    Dim cell  As String
    Dim resp  As Integer
    resp = MsgBox("Are you sure you want to undo this change?", vbYesNo, "Undo Célula alterada")
    If resp = vbNo Then Exit Sub
    With Logs
    
        LogRow = .Range("B3").Value    'Set Row
    
        nomePlanilha = .Range("G" & LogRow).Value 'Pega nome da planilha
        cell = .Range("H" & LogRow).Value 'Get Cell Address
        ActiveWorkbook.Worksheets(nomePlanilha).Range(cell).Value = .Range("I" & LogRow).Value 'Add Back Previous Value
        .Range(LogRow & ":" & LogRow).EntireRow.Delete
    End With
End Sub

Public Sub Log_Show()
    Dim Pass  As String
    Pass = InputBox("Enter password to access log", "Enter Password")
    If Pass = "demo" Then              'Correct Password
        Logs.Visible = xlSheetVisible
        Logs.Activate
    Else:                              'Incorrect Password
        MsgBox "Incorrect password entered"
    End If
End Sub

