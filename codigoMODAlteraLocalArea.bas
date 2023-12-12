Attribute VB_Name = "MODAlteraLocalArea"
Public Sub chamaformEditalocal()
Attribute chamaformEditalocal.VB_ProcData.VB_Invoke_Func = "L\n14"
    frmAlteraLocal.Show
End Sub
Public Sub inseredataserv()
Attribute inseredataserv.VB_Description = "Insere a data dos serviços executados no formulário Cdastro - Atualização"
Attribute inseredataserv.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim dataserv1 As Date, dataserv2 As Date, resposta As VbMsgBoxResult

'    resposta = MsgBox("As datas dos serviços são diferentes?", vbYesNo, "Serviços")
'    If resposta = vbYes Then
'    dataserv1 = CDate("1/2/23") 'recarga
'    dataserv2 = CDate("1/2/23") 'teste

'    dataserv1 = CDate("1/12/22") 'recarga
'    dataserv2 = CDate("1/12/22") 'teste
'
    
   
    
    dataserv1 = Info.Range("$T$3").Value 'CDate("1/12/22") 'recarga
    dataserv2 = Info.Range("$S$3").Value 'CDate("1/12/21") 'teste



    With Info
If .Range("M8").FormulaR1C1 <> "CO" Then
'
Range("i16").FormulaR1C1 = dataserv2
        .Range("$M$16").FormulaR1C1 = dataserv1

        .Range("$M$18").FormulaR1C1 = dataserv1
        .Range("$M$20").FormulaR1C1 = dataserv2
        .Range("$I$20").FormulaR1C1 = dataserv1
'
Else
        .Range("i16").FormulaR1C1 = dataserv2
        .Range("$M$16").FormulaR1C1 = dataserv1
        .Range("$I$18").FormulaR1C1 = dataserv1
        .Range("$M$18").FormulaR1C1 = dataserv1
        .Range("$M$20").FormulaR1C1 = dataserv2
        .Range("$I$20").FormulaR1C1 = dataserv1
        End If
'
        serie = Info.Range("I8").Value
'
    End With
End Sub

Public Sub modificaLocalArea()
Attribute modificaLocalArea.VB_ProcData.VB_Invoke_Func = "L\n14"
    Dim ultlinmapa As Long, ultlinlocal As Long, ultlinmov As Long
    Dim lin   As Long
    Dim iControl As control
    Dim localmapaSaida As String, areamapaSaida As String, enderecoEntrada As String, zonamapasaida As String, enderecoSaida As String
    Dim localmapaEntrada As String, zonamapaentrada As String, areamapaEntrada As String

    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    ultlinmov = Movimentacao.ListObjects(1).DataBodyRange.Rows.Count
    ultlinlocal = locais.ListObjects(1).DataBodyRange.Rows.Count

    
    '    localmapaSaida = InputBox("Qual local deseja alterar?")
    '    localmapaEntrada = InputBox("Digite o novo local")
    '    areamapaSaida = InputBox("Qual a area a ser alterada?")
    '    areamapaEntrada = InputBox("Digite a nova area")
  
    
    
    localmapaSaida = Trim$(AlphaNumericOnly(frmAlteraLocal.txtLAntigo.Value))
    localmapaEntrada = Trim$(AlphaNumericOnly(frmAlteraLocal.txtLNovo.Value))
    areamapaSaida = Trim$(AlphaNumericOnly(frmAlteraLocal.txtAAtual.Value))
    areamapaEntrada = Trim$(AlphaNumericOnly(frmAlteraLocal.txtAnova.Value))
    zonamapasaida = Trim$(AlphaNumericOnly(frmAlteraLocal.txtZonaAtual.Value))
    zonamapaentrada = Trim$(AlphaNumericOnly(frmAlteraLocal.txtZonaNova.Value))
    
    enderecoSaida = localmapaSaida & " - " & areamapaSaida
    enderecoEntrada = localmapaEntrada & " - " & areamapaEntrada
    'tblocais
    For Each iControl In frmAlteraLocal.Controls
        If TypeOf iControl Is MSForms.TextBox Then
            If iControl.Value = vbNullString Then
                MsgBox "Preencha todos os campos", vbCritical, "Campos vazios"
                Exit Sub
            End If
        End If
    Next iControl
    
    'altera locais em locais
    With locais.ListObjects(1).DataBodyRange
        lin = 1
        
        Do Until lin > ultlinlocal     'busca status geral do extintor

            If enderecoSaida = .Cells(lin, 4) Then
                .Cells(lin, 4) = enderecoEntrada
                .Cells(lin, 3) = areamapaEntrada
                .Cells(lin, 1) = zonamapaentrada
                locais.PivotTables("tbdimLocal").PivotCache.Refresh
                locais.PivotTables("tbdimLocal").PivotFields("Local").AutoSort _
        xlAscending, "Local"
                Exit Do
            End If
            lin = lin + 1
        Loop

    End With
    'altera locais em mov
    With Movimentacao.ListObjects(1).DataBodyRange
        lin = 1
        Do Until lin > ultlinmov

            If enderecoSaida = .Cells(lin, 4) & " - " & .Cells(lin, 5) Then
                .Cells(lin, 4) = localmapaEntrada
                .Cells(lin, 5) = areamapaEntrada
                 .Cells(lin, 8) = zonamapaentrada
            End If
            If enderecoSaida = .Cells(lin, 6) & " - " & .Cells(lin, 7) Then
                .Cells(lin, 6) = localmapaEntrada
                .Cells(lin, 7) = areamapaEntrada
                .Cells(lin, 8) = zonamapaentrada
                '                Exit Do
            End If
            lin = lin + 1
        Loop

    End With
    
    'altera locais em mapa
    
    With MapaAtual.ListObjects(1).DataBodyRange
        lin = 1
        Do Until lin > ultlinmapa

            If enderecoSaida = .Cells(lin, 4) & " - " & .Cells(lin, 2) Then
                .Cells(lin, 4) = localmapaEntrada
                .Cells(lin, 2) = areamapaEntrada
                 .Cells(lin, 9) = zonamapaentrada
            End If
           
            lin = lin + 1
        Loop
        Info.Unprotect
Info.Range("$I$14").Value = areamapaEntrada
Info.Unprotect
Info.Range("$A$7").Value = areamapaEntrada
Info.Unprotect
Info.Range("$M$12").Value = localmapaEntrada
Info.Unprotect
Info.Range("$A$6").Value = localmapaEntrada
If ActiveSheet Is Info Then
Info.Range("$I$8").Select
End If
Application.Speech.Speak "Local alterado"
Info.Protect
    End With
    
    MsgBox "Alteração bem sucedida!", vbExclamation, "Concluído!"
End Sub


Public Function Isalphanumeric(cadena As String) As Boolean

    Select Case Asc(UCase$(cadena))
        Case 65 To 90                  'letras
            Isalphanumeric = True
        Case 48 To 57                  'numeros
            Isalphanumeric = True
        Case 45                        'HIFEN
            Isalphanumeric = True
        Case 32                        'ESPAÇO
            Isalphanumeric = True
        Case Else
            Isalphanumeric = False

    End Select

End Function


Public Function AlphaNumericOnly(strSource As String) As String
    Dim i     As Long
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid$(strSource, i, 1))
            Case 32, 45, 48 To 57, 65 To 90, _
                 97 To 122, 170, 176, 180, 186, _
                 192 To 195, 199 To 206, 210 To 213, _
                 217 To 218, 224 To 227, 231 To 238, _
                 242 To 245, 249 To 250: 'include 32 if you want to include space
                strResult = strResult & Mid$(strSource, i, 1)
        End Select
    Next
    AlphaNumericOnly = strResult
End Function




