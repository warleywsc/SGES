Attribute VB_Name = "M�dulo7"
Option Explicit

Sub updateExtReserva()

    Dim resp  As String
    Application.Speech.Speak "Deseja atualizar o �CIG�S ante de iniciar??", speakasync:=True
    resp = MsgBox("Deseja atualizar o SGES ante de iniciar?", vbYesNo, "Atualiza��o")

    If resp = vbYes Then
        atualizamapaatual
        Info.Unprotect
        Info.Range("N28").Value = DateAdd("d", 1, Info.Range("m28").Value)
        Info.Protect
    Else
        Exit Sub
    End If
End Sub
