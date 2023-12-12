Attribute VB_Name = "MODcontvencidos"
Option Explicit






Public Sub contvencidos()
    Dim vencidos As Long, lin As Long
    Dim vencendo As Long
    Dim emdia As Long

    Dim arr   As Variant, i As Long
    vencidos = 0
    vencendo = 0
    emdia = 0
    'i = 1
    arr = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    For i = LBound(arr, 1) To UBound(arr, 1)
    
        If InStr(arr(i, 11), UCase$("Vencid")) Or _
                                               InStr(arr(i, 11), UCase$("subs")) Or _
                                               InStr(arr(i, 13), UCase$("Vencid")) Or _
                                               InStr(arr(i, 13), UCase$("Vencid")) Or _
                                               InStr(arr(i, 15), UCase$("Vencid")) Or _
                                               InStr(arr(i, 17), UCase$("Vencid")) Or _
                                               InStr(arr(i, 19), UCase$("Vencid")) Then
        
            vencidos = vencidos + 1
      
        ElseIf InStr(arr(i, 11), UCase$("Atenção")) Or _
                                                    InStr(arr(i, 13), UCase$("Atenção")) Or _
                                                    InStr(arr(i, 15), UCase$("Atenção")) Or _
                                                    InStr(arr(i, 17), UCase$("Atenção")) Or _
                                                    InStr(arr(i, 19), UCase$("Atenção")) Then
        
            vencendo = vencendo + 1
      
        ElseIf InStr(arr(i, 11), UCase$("Em Dia")) Or _
                                                   InStr(arr(i, 13), UCase$("Em Dia")) Or _
                                                   InStr(arr(i, 15), UCase$("Em Dia")) Or _
                                                   InStr(arr(i, 17), UCase$("Em Dia")) Or _
                                                   InStr(arr(i, 19), UCase$("Em Dia")) Then
        
            emdia = emdia + 1
            lin = lin + 1
        End If
    Next
    'MapaAtual.ListObjects("tbMapaAtual").DataBodyRange = arr
    MsgBox vencidos & " Vencidos, " & vencendo & " vencendo " & " e " & emdia & " em dia." & " Total: " & vencidos + vencendo + emdia
      
End Sub




