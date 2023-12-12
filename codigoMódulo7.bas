Attribute VB_Name = "Módulo7"
Option Explicit

Sub updateExtReserva()

    Dim resp  As String
    Application.Speech.Speak "Deseja atualizar o ÉCIGÉS ante de iniciar??", speakasync:=True
    resp = MsgBox("Deseja atualizar o SGES ante de iniciar?", vbYesNo, "Atualização")

    If resp = vbYes Then
        atualizamapaatual
        Info.Unprotect
        Info.Range("N28").Value = DateAdd("d", 1, Info.Range("m28").Value)
        Info.Protect
    Else
        Exit Sub
    End If
End Sub
