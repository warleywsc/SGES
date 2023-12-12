Attribute VB_Name = "MODPINTACPVAZIOS"
Option Explicit

Public Sub PINTAVAZIOCILINDRO()
    Dim cell  As Range
    Dim contvazio As Long
    With Info
        For Each cell In .Range("I8,M8,M10,I12,M12,I14,M14,I16,M16,I18,I20,M20")
            cell.Interior.Color = &HF9F9F9
    
            cell.ClearComments
            If cell.Value = vbNullString Then
                cell.Interior.Color = &HC0C0FF 'pinta campos vazios
                cell.AddComment
                cell.Comment.Visible = True
                cell.Comment.Shape.TextFrame.Characters.Font.Bold = True
                cell.Comment.Shape.TextFrame.Characters.Font.Size = 12
                cell.Comment.Shape.TextFrame.Characters.Font.Color = &HCC
                cell.Comment.Text Text:="SGES:" & Chr$(10) & "Preencha todos os campos!!!"
            
                contvazio = contvazio + 1
    
            End If
        Next cell
        If contvazio > 0 Then
    
            Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
            MsgBox "Há campos vazios no formulário! Preencha todos os campos!", , "Atenção!"
            Application.EnableEvents = True
            Application.ScreenUpdating = True
    
    
    
            Exit Sub
        End If

    End With

End Sub

Public Sub PINTAVAZIONORMAL()
    Dim cell  As Range
    Dim contvazio As Long
    With Info

        For Each cell In .Range("I8,M8,M10,I12,M12,I14,M14,I16,M16,I18,M18,I20,M20")
            cell.Interior.Color = &HF9F9F9
    
            cell.ClearComments
            If cell.Value = vbNullString Then
                cell.Interior.Color = &HC0C0FF 'pinta campos vazios
                cell.AddComment
                cell.Comment.Visible = True
                cell.Comment.Shape.TextFrame.Characters.Font.Bold = True
                cell.Comment.Shape.TextFrame.Characters.Font.Size = 12
                cell.Comment.Shape.TextFrame.Characters.Font.Color = &HCC
                cell.Comment.Text Text:="SGES:" & Chr$(10) & "Preencha todos os campos!!!"
            
                contvazio = contvazio + 1
    
            End If
        Next cell
        If contvazio > 0 Then
    
            Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
            MsgBox "Há campos vazios no formulário! Preencha todos os campos!", , "Atenção!"
            Application.EnableEvents = True
            Application.ScreenUpdating = True
    
    
    
            Exit Sub
        End If


    End With
End Sub



Public Sub PINTAVAZIOIK()
    Dim cell  As Range
    Dim contvazio As Long
    With Info

        For Each cell In .Range("I8,M8,M10,I12,M12,I14,M14,I16,I20,M20")
            cell.Interior.Color = &HF9F9F9
    
            cell.ClearComments
            If cell.Value = vbNullString Then
                cell.Interior.Color = &HC0C0FF 'pinta campos vazios
                cell.AddComment
                cell.Comment.Visible = True
                cell.Comment.Shape.TextFrame.Characters.Font.Bold = True
                cell.Comment.Shape.TextFrame.Characters.Font.Size = 12
                cell.Comment.Shape.TextFrame.Characters.Font.Color = &HCC
                cell.Comment.Text Text:="SGES:" & Chr$(10) & "Preencha todos os campos!!!"
            
                contvazio = contvazio + 1
    
            End If
        Next cell
        If contvazio > 0 Then
    
            Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
            MsgBox "Há campos vazios no formulário! Preencha todos os campos!", , "Atenção!"
            Application.EnableEvents = True
            Application.ScreenUpdating = True
    
    
    
            Exit Sub
        End If


    End With
End Sub

