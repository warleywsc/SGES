Attribute VB_Name = "MODAtualizatblocal"
'@Folder("SGES2020")
Option Explicit

Public Sub verificalocalxarea()

    Dim localxarea As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    
    With Info
        localxarea = .Cells(12, 13).Value & " - " & .Cells(14, 9).Value
       
        lin = 9
      
      
        
        Do Until MapaAtual.Cells(lin, 14) = vbNullString
            If MapaAtual.Cells(lin, 10).Value & " - " & MapaAtual.Cells(lin, 8).Value = localxarea Then
                GoTo Continua:
               
                Exit Sub
            Else
                MsgBox "Local não encontrado! Talvez seja necessário modificar a área.", , "Erro de Local"
                GoTo ErrorHandler:
            End If
            lin = lin + 1
        Loop

ErrorHandler:
        Exit Sub

    End With
 
Continua:

End Sub
