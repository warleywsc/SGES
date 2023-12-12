Attribute VB_Name = "MODInsereInsp"
Option Explicit

Sub insereinsp()
Attribute insereinsp.VB_ProcData.VB_Invoke_Func = "P\n14"
If Info.Range("$M$8").Value = "CO" Then
Info.Range("$I$18") = "8/1/23"
End If
Info.Range("$I$20") = "8/1/23"
End Sub


Sub insereselo()
Attribute insereselo.VB_ProcData.VB_Invoke_Func = "O\n14"
Info.Range("m12") = "CALDEIRA AUXILIAR"
Info.Range("$I$14") = "169"
'Info.Range("$I$20") = "1/1/2022"
End Sub
