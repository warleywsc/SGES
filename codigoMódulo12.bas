Attribute VB_Name = "Módulo12"
Sub ExtrairFormatacaoCondicional()
    Dim rng As Range
    Dim condFormatRules As FormatCondition
    Dim rule As FormatCondition
    Dim i As Integer
    
    Set rng = Selection
    
    If rng.FormatConditions.Count > 0 Then
        For i = 1 To rng.FormatConditions.Count
            Set condFormatRules = rng.FormatConditions(i)
            
            Debug.Print "Regra " & i & ":"
            Debug.Print "Tipo: " & condFormatRules.Type
            Debug.Print "Fórmula: " & condFormatRules.Formula1
            Debug.Print "Valor: " & condFormatRules.Formula1
            Debug.Print "Formato: " & condFormatRules.Interior.Color
            Debug.Print ""
        Next i
    Else
        Debug.Print "Nenhuma formatação condicional encontrada na célula selecionada."
    End If
End Sub

Sub InserirFormatacaoCondicional()
    Dim rng As Range
    Dim condFormatRules As FormatCondition
    
    Set rng = Selection
    
    ' Limpa as regras de formatação condicional existentes na célula selecionada
    rng.FormatConditions.Delete
    
    ' Regra 1
    Set condFormatRules = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=CELL(""address"")=CELL(""address"",$M$16)")
    condFormatRules.Interior.Color = 13564414
    
    ' Regra 2
    Set condFormatRules = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(AND(DATE(YEAR(M16)+5,MONTH(M16),DAY(M16))-$M$28<1,$M$8=""CO""),AND(DATE(YEAR(M16)+5,MONTH(M16),DAY(M16))-$M$28<1,$M$8=""FM""),AND(DATE(YEAR(M16)+1,MONTH(M16),DAY(M16))-$M$28<1,$M$8<>""CO"",$M$8<>""FM""))")
    condFormatRules.Interior.Color = 192
    
    ' Regra 3
    Set condFormatRules = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(AND(DATE(YEAR(M16)+5,MONTH(M16),DAY(M16))-$M$28<30,$M$8=""CO""),AND(DATE(YEAR(M16)+5,MONTH(M16),DAY(M16))-$M$28<30,$M$8=""FM""),AND(DATE(YEAR(M16)+1,MONTH(M16),DAY(M16))-$M$28<30,$M$8<>""CO"",$M$8<>""FM""))")
    condFormatRules.Interior.Color = 65535
    
    ' Regra 4
    Set condFormatRules = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(AND(DATE(YEAR(M16)+5,MONTH(M16),DAY(M16))-$M$28>30,$M$8=""CO""),AND(DATE(YEAR(M16)+5,MONTH(M16),DAY(M16))-$M$28>30,$M$8=""FM""),AND(DATE(YEAR(M16)+1,MONTH(M16),DAY(M16))-$M$28>30,$M$8<>""CO"",$M$8<>""FM""))")
    condFormatRules.Interior.Color = 5287936
    
    ' Regra 5
    Set condFormatRules = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($M$10=""1K"", $M$16 > DATE(YEAR(NOW())+5, MONTH(NOW()), DAY(NOW())))")
    condFormatRules.Interior.Color = RGB(255, 0, 0) ' Vermelho
    
    ' Regra 6
    Set condFormatRules = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($M$10=""1K"", $M$16 > DATE(YEAR(NOW())+4, MONTH(NOW())+11, DAY(NOW()))))")
    condFormatRules.Interior.Color = RGB(255, 255, 0) ' Amarelo
    
    ' Regra 7
    Set condFormatRules = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($M$10=""1K"", $M$16 < DATE(YEAR(NOW())+4, MONTH(NOW())+11, DAY(NOW()))))")
    condFormatRules.Interior.Color = RGB(0, 255, 0) ' Verde
End Sub


