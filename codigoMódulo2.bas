Attribute VB_Name = "Módulo2"
Option Explicit


Public Sub MovEnvioEmBloco2()
    Dim tbEnvio As Range
    Dim tbmov As Range
    Set tbEnvio = formenvio.Range("J9:J" & formenvio.Cells(Rows.Count, "J").End(xlUp).Row)
    '    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim tbmovarray() As Variant
    Dim tbenvioArray() As Variant
    Dim a     As Integer
    Dim b     As Integer
    Dim c     As Integer
    tbenvioArray = tbEnvio
    ReDim Preserve tbenvioArray(1 To tbEnvio.Rows.Count, 1 To 1)
    ReDim Preserve tbmovarray(1 To tbEnvio.Rows.Count * 2, 1 To 8)
    c = 1
    a = 2
    For b = 1 To tbEnvio.Rows.Count

        tbmovarray(c, 1) = DateAdd("s", b, Now)
        tbmovarray(c, 2) = tbenvioArray(b, 1)
        tbmovarray(c, 3) = "Saída"
        tbmovarray(c, 4) = "MANUTENÇÃO - BRIGADA"
        tbmovarray(c, 5) = "1111"
        tbmovarray(c, 8) = "BRIGADA"

        tbmovarray(a, 1) = DateAdd("s", b + 1, Now)
        tbmovarray(a, 2) = tbenvioArray(b, 1)
        tbmovarray(a, 3) = "Entrada"
        tbmovarray(a, 6) = "MANUTENÇÃO - MAREFIRE"
        tbmovarray(a, 7) = "9999"
        tbmovarray(a, 8) = "BRIGADA"
        c = c + 2
        a = a + 2

    Next b
    
    Set tbmov = Movimentacao.Range("G" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row & ":N" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
    tbmov = tbmovarray
    Set tbEnvio = Nothing
    Set tbmov = Nothing
    
End Sub
