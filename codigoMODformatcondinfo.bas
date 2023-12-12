Attribute VB_Name = "MODformatcondinfo"
Sub formatcondinfo()
Info.Unprotect
Dim teste As Range, recarga As Range, pesagem As Range
Dim selo As Range, inspecao As Range, pintura As Range
Dim tipo As Range, capacidade As Range

Set tipo = Info.Range("$M$8")
Set capacidade = Info.Range("$M$10")
Set teste = Info.Range("$I$16")
Set recarga = Info.Range("$M$16")
Set pesagem = Info.Range("$I$18")
Set selo = Info.Range("$M$18")
Set inspecao = Info.Range("$I$20")
Set pintura = Info.Range("$M$20")

' VERIFICA CAPACIDADE

If capacidade <> "1K" Then
    If UCase(tipo) <> UCase("co") Then
    
    
     If DateDiff("m", teste, Date) = 0 Then
    
                teste.Cells.Interior.ColorIndex = 6
            
            ElseIf DateDiff("m", teste, Date) > 0 Then
    
                Info.Range(teste.Address).Interior.ColorIndex = 6
    
            Else
    
                teste.Cells.Interior.ColorIndex = 10
    End If
    
    Else
    
    End If


    
    Else
    
    




End If




End Sub
