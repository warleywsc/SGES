Attribute VB_Name = "MODInsereManut"


Public Sub movmanut()
Attribute movmanut.VB_ProcData.VB_Invoke_Func = "m\n14"

    With Info

        .Range("m12").Value = "MANUTENÇÃO - BRIGADA"
        .Range("i14").Value = "0000"
        .Range("i16").Select
    End With
End Sub

Public Sub movmanutEXTERNA()
Attribute movmanutEXTERNA.VB_ProcData.VB_Invoke_Func = "e\n14"

    With Info

        .Range("m12").Value = "MANUTENÇÃO - MAREFIRE"
        .Range("i14").Value = "9999"
        .Range("i16").Activate
    End With
End Sub

Public Sub movmanutreserva()
Attribute movmanutreserva.VB_ProcData.VB_Invoke_Func = "r\n14"

    With Info

        .Range("m12").Value = "RESERVA TÉCNICA"
        .Range("i14").Value = "1111"
        .Range("i16").Select
    End With
End Sub
