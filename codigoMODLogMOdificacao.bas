Attribute VB_Name = "MODLogMOdificacao"
Dim oldValues() As Variant
Dim tableValues As Variant
Dim tableRange As Range

Sub HandleSheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim cell As Range
    Dim userName As String
    Dim sheetName As String
    Dim lastRow As ListRow
    Dim oldValue As Variant
    Dim newValue As Variant
    
    ' Defina a planilha de log onde as alterações serão registradas (no seu caso, "Logs")
    Set wsLog = ThisWorkbook.Sheets("Logs")
    
    ' Defina a tabela onde serão registradas as alterações (use o nome correto da sua tabela)
    Set tblLog = wsLog.ListObjects("Tblog")
    
    ' Verifique se a célula de destino não está vazia
    If Not Intersect(Target, Sh.UsedRange) Is Nothing Then
        ' Obtenha o nome de usuário (você pode personalizar essa parte)
        userName = Environ("Username")
        
        ' Obtenha o nome da planilha onde as células foram modificadas
        sheetName = Sh.Name
        
        ' Carregue os valores da tabela de log em um array
        Set tableRange = tblLog.DataBodyRange
        tableValues = tableRange.Value
        
        ' Redimensione o array de valores antigos para o mesmo tamanho que as células modificadas
        ReDim oldValues(1 To Target.Cells.Count)
        
        ' Percorra todas as células alteradas
        Dim i As Long
        i = 1
        For Each cell In Target
            ' Armazene o valor atual da célula antes de ser modificado
            oldValue = cell.Value
            ' Realize a modificação (isso ativará o evento Change novamente)
            cell.Value = oldValue
            ' Obtenha o novo valor da célula após a modificação
            newValue = cell.Value
            
            ' Compare o valor antigo com o valor na tabela
            If oldValue <> tableValues(cell.Row, cell.Column) Then
                ' Se forem diferentes, registre a alteração
                
                ' Encontre a última linha preenchida na tabela de log
                Set lastRow = tblLog.ListRows(tblLog.ListRows.Count)
                
                ' Adicione uma nova linha à tabela de log após a última linha preenchida
                Set newRow = tblLog.ListRows.Add
                
                ' Registre a data e hora da modificação
                newRow.Range(1) = Now
                ' Registre o nome de usuário
                newRow.Range(2) = userName
                ' Registre o nome da planilha
                newRow.Range(3) = sheetName
                ' Registre o nome da tabela à qual a célula pertence
                newRow.Range(4) = tblLog.Name
                ' Registre o intervalo (formate como um link para a célula modificada)
                newRow.Range(5).Hyperlinks.Add Anchor:=newRow.Range(5), Address:="", SubAddress:=sheetName & "!" & cell.Address, TextToDisplay:=cell.Address
                ' Registre os valores antigos e novos (apenas os valores, sem formatação)
                newRow.Range(6) = Trim(oldValue)
                newRow.Range(7) = Trim(newValue)
                
                ' Próxima célula
                i = i + 1
            End If
        Next cell
    End If
End Sub











Sub MostrarTodasPlanilhasOcultas()
    Dim ws As Worksheet
    
    ' Loop através de todas as planilhas na pasta de trabalho
    For Each ws In ThisWorkbook.Worksheets
        ' Verifique se a planilha está oculta
        If ws.Visible = xlSheetHidden Then
            ' Se estiver oculta, torne-a visível
            ws.Visible = xlSheetVisible
        End If
    Next ws
End Sub


