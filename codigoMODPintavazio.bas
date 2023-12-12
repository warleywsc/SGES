Attribute VB_Name = "MODPintavazio"
'@Folder("SGES2020")
Option Explicit

Public Sub PINTAVAZIO()

    Dim contvazio As Long
    Dim cell  As Range

    With Info
        

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        If Range("M8").Value = "CO" Or Range("M8").Value = "FM" Then

            'VERIFICA SE ALGUM CAMPO ESTÁ VAZIO ANTES DE SALVAR  INFO

            .Unprotect
            For Each cell In .Range("I8,M8,M10,I12,M12,I14,M14,I16,M16,I18,M18,I20,M20")
                cell.Interior.Color = &HF9F9F9

                If cell.Value = vbNullString Then
                    cell.Interior.Color = &HC0C0FF 'pinta campos vazios

                    contvazio = contvazio + 1

                End If
            Next cell
            If contvazio > 0 Then

                MsgBox "Há campos vazios no formulário! Preencha todos os campos!"
                .Protect
            End If
        End If
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
              
End Sub
