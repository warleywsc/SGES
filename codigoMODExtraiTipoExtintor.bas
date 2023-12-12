Attribute VB_Name = "MODExtraiTipoExtintor"
Sub DefinirTipoExtintor(Target As Range)
If IsEmpty(Target) Then Exit Sub
    Dim serie As String
    serie = Target.Value

    Dim tipo As String
    tipo = ""

    ' Lista de tipos possíveis
    Dim tiposPossiveis As Variant
    tiposPossiveis = Array("AP", "CO", "EM", "PQ", "FM", "NL", "PE")

    Dim encontrado As Boolean
    encontrado = False

    Dim j As Integer
    For j = LBound(tiposPossiveis) To UBound(tiposPossiveis)
        If InStr(serie, tiposPossiveis(j)) > 0 Then
            tipo = tiposPossiveis(j)
            encontrado = True
            Exit For
        End If
    Next j

    If Not encontrado Then
        tipo = ""
    End If

    ' Desabilita eventos para evitar chamadas recursivas
    Application.EnableEvents = False
    Target.Offset(0, 1).Value = tipo ' Coloca o tipo na coluna seguinte
    Application.EnableEvents = True
End Sub

