Attribute VB_Name = "MODMovimentaManutencao"
'@Folder("SGES2020")
Option Explicit
Public Sub chamaformenviomanut()
    On Error GoTo TError
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    formenvio.Activate
    frmMovimentaManutencao.Show (False)
    frmMovimentaManutencao.cbEnvioSerie.SetFocus
    Application.EnableEvents = True
    Application.ScreenUpdating = True

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "chamaformenvio()"
    GoTo fim
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub populalboxenvio()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub populalboxenvio()
    On Error GoTo TError

    '    Dim lo    As ListObject
    Dim lin   As Long
    Dim baseformEnvio As Range
    '    Set baseformEnvio = formenvio.Range("G9:o" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row).Rows
    
    lin = formenvio.Range("g8").CurrentRegion.Rows.Count + 8
    If lin = 9 Then lin = 10
    Set baseformEnvio = formenvio.Range(formenvio.Cells(9, 7), _
                                        formenvio.Cells(lin, 15))
    
    With frmMovimentaManutencao.ListBoxenvio
    
        '    .RowSource = vbNullString
        .ColumnCount = baseformEnvio.Columns.Count
        .ColumnHeads = True
        .Font.Size = 10
        .BackColor = RGB(52, 108, 135)
        .ForeColor = vbWhite
        .Font.Bold = True
        .RowSource = baseformEnvio.Address(external:=True)
    
    
    End With
    
    
    

    Set baseformEnvio = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "populalboxenvio()"
    GoTo fim
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub EnviarManutenção()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub EnviarManutenção()
    On Error GoTo TError


    Dim ultlinmapa As Long, lin As Long, ultilinformenvio As Long
    Dim serie As String
    Dim basemapa As Range, baseformEnvio As Range, colunaseriebaseenvio As Range
    Dim cell  As Range

    
    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    serie = frmMovimentaManutencao.cbEnvioSerie.Text
    Set basemapa = MapaAtual.ListObjects(1).DataBodyRange
    
    Set baseformEnvio = formenvio.Range("G8:N" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)
    '    Set colunaseriebaseenvio = baseformEnvio.Range(Cells(2, 4), Cells(baseformEnvio.Rows.Count, 4))
    Set colunaseriebaseenvio = formenvio.Range("J8:J" & formenvio.Cells(Rows.Count, "J").End(xlUp).Row)
    For Each cell In colunaseriebaseenvio.Cells
    
        If cell = serie Then
        
            MsgBox "Este extintor já está na lista de envio", vbExclamation, "Série duplicado"
            frmMovimentaManutencao.cbEnvioSerie.SetFocus
            GoTo fim:
  
        End If
    Next cell
    
    populalboxenvio
    For lin = 1 To basemapa.Rows.Count
    
        With frmMovimentaManutencao
        
            ultilinformenvio = baseformEnvio.Rows.Count
            If serie = basemapa.Cells(lin, 8) Then
               
                If baseformEnvio.Cells(2, 1) = "ID" Then
                    'ID  Data do Registro    Série   Tipo    Capacidade  Último Teste    Última Recarga  Data de Envio
                    On Error Resume Next
                    baseformEnvio.Cells(ultilinformenvio + 1, 1) = ultilinformenvio
                    baseformEnvio.Cells(ultilinformenvio + 1, 2) = Date
                    baseformEnvio.Cells(ultilinformenvio + 1, 3) = .cbEnvioSerie.Value
                    baseformEnvio.Cells(ultilinformenvio + 1, 4) = .cbEnvioSerie.Text
                    baseformEnvio.Cells(ultilinformenvio + 1, 5) = basemapa.Cells(lin, 5)
                    baseformEnvio.Cells(ultilinformenvio + 1, 6) = basemapa.Cells(lin, 6)
                    baseformEnvio.Cells(ultilinformenvio + 1, 7) = Format$(basemapa.Cells(lin, 10), "dd/mm/yyyy")
                    'recarga
                    If basemapa.Cells(lin, 5) = "CO" Then
                        baseformEnvio.Cells(ultilinformenvio + 1, 8) = Format$(DateAdd("yyyy", -5, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
                    
                    Else
                        baseformEnvio.Cells(ultilinformenvio + 1, 8) = Format$(DateAdd("yyyy", -1, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
            
                    End If
                    baseformEnvio.Cells(ultilinformenvio + 1, 9) = Format$(Date, "dd/mm/yyyy")
            
                    Exit Sub
                
                
                Else
                    On Error Resume Next
                    baseformEnvio.Cells(ultilinformenvio + 1, 1) = ultilinformenvio
                    baseformEnvio.Cells(ultilinformenvio + 1, 2) = Date
                    baseformEnvio.Cells(ultilinformenvio + 1, 3) = .cbEnvioSerie.Value
                    baseformEnvio.Cells(ultilinformenvio + 1, 4) = .cbEnvioSerie.Text
                    baseformEnvio.Cells(ultilinformenvio + 1, 5) = basemapa.Cells(lin, 5)
                    baseformEnvio.Cells(ultilinformenvio + 1, 6) = basemapa.Cells(lin, 6)
                    baseformEnvio.Cells(ultilinformenvio + 1, 7) = Format$(basemapa.Cells(lin, 10), "dd/mm/yyyy")
                    'recarga
                    If basemapa.Cells(lin, 5) = "CO" Then
                        baseformEnvio.Cells(ultilinformenvio + 1, 8) = Format$(DateAdd("yyyy", -5, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
                    
                    Else
                        baseformEnvio.Cells(ultilinformenvio + 1, 8) = Format$(DateAdd("yyyy", -1, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
            
                    End If
                    baseformEnvio.Cells(ultilinformenvio + 1, 9) = Format$(Date, "dd/mm/yyyy")
                
                End If
            End If
    
        End With
    
    Next lin
    populalboxenvio
    frmMovimentaManutencao.cbEnvioSerie.SetFocus
   
    Set basemapa = Nothing
    Set baseformEnvio = Nothing
    Set colunaseriebaseenvio = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "EnviarManutenção()"
    GoTo fim
End Sub
Public Sub ReceberManutenção()
    On Error GoTo TError


    Dim ultlinmapa As Long, lin As Long, ultilinformRetorno As Long
    Dim serie As String
    Dim basemapa As Range, baseformRetorno As Range, colunaseriebaseRetorno As Range
    Dim cell  As Range

    
    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    serie = frmMovimentaManutencao.cbRetornoSerie.Text
    Set basemapa = MapaAtual.ListObjects(1).DataBodyRange
    
    Set baseformRetorno = formenvio.Range("AO8:aw" & formenvio.Cells(Rows.Count, "AO").End(xlUp).Row)
    '    Set colunaseriebaseRetorno = baseformRetorno.Range(Cells(2, 4), Cells(baseformRetorno.Rows.Count, 4))
    Set colunaseriebaseRetorno = formenvio.Range("AR8:AR" & formenvio.Cells(Rows.Count, "AR").End(xlUp).Row)
    For Each cell In colunaseriebaseRetorno.Cells
    
        If cell = serie Then
        
            MsgBox "Este extintor já está na lista de Recebimento", vbExclamation, "Série duplicado"
            frmMovimentaManutencao.cbRetornoSerie.SetFocus
            GoTo fim:
  
        End If
    Next cell
    
    populalbxRetorno
    For lin = 1 To basemapa.Rows.Count
    
        With frmMovimentaManutencao
        
            ultilinformRetorno = baseformRetorno.Rows.Count
            If serie = basemapa.Cells(lin, 8) Then
               
                If baseformRetorno.Cells(2, 1) = "ID" Then
                    'ID  Data do Registro    Série   Tipo    Capacidade  Último Teste    Última Recarga  Data de Envio
                    On Error Resume Next
                    baseformRetorno.Cells(ultilinformRetorno + 1, 1) = ultilinformRetorno
                    baseformRetorno.Cells(ultilinformRetorno + 1, 2) = Date
                    baseformRetorno.Cells(ultilinformRetorno + 1, 3) = .cbRetornoSerie.Value
                    baseformRetorno.Cells(ultilinformRetorno + 1, 4) = .cbRetornoSerie.Text
                    baseformRetorno.Cells(ultilinformRetorno + 1, 5) = basemapa.Cells(lin, 5)
                    baseformRetorno.Cells(ultilinformRetorno + 1, 6) = basemapa.Cells(lin, 6)
                    baseformRetorno.Cells(ultilinformRetorno + 1, 7) = Format$(basemapa.Cells(lin, 10), "dd/mm/yyyy")
                    'recarga
                    If basemapa.Cells(lin, 5) = "CO" Then
                        baseformRetorno.Cells(ultilinformRetorno + 1, 8) = Format$(DateAdd("yyyy", -5, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
                    
                    Else
                        baseformRetorno.Cells(ultilinformRetorno + 1, 8) = Format$(DateAdd("yyyy", -1, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
            
                    End If
                    baseformRetorno.Cells(ultilinformRetorno + 1, 9) = Format$(Date, "dd/mm/yyyy")
            
                    Exit Sub
                
                
                Else
                    On Error Resume Next
                    baseformRetorno.Cells(ultilinformRetorno + 1, 1) = ultilinformRetorno
                    baseformRetorno.Cells(ultilinformRetorno + 1, 2) = Date
                    baseformRetorno.Cells(ultilinformRetorno + 1, 3) = .cbRetornoSerie.Value
                    baseformRetorno.Cells(ultilinformRetorno + 1, 4) = .cbRetornoSerie.Text
                    baseformRetorno.Cells(ultilinformRetorno + 1, 5) = basemapa.Cells(lin, 5)
                    baseformRetorno.Cells(ultilinformRetorno + 1, 6) = basemapa.Cells(lin, 6)
                    baseformRetorno.Cells(ultilinformRetorno + 1, 7) = Format$(basemapa.Cells(lin, 10), "dd/mm/yyyy")
                    'recarga
                    If basemapa.Cells(lin, 5) = "CO" Then
                        baseformRetorno.Cells(ultilinformRetorno + 1, 8) = Format$(DateAdd("yyyy", -5, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
                    
                    Else
                        baseformRetorno.Cells(ultilinformRetorno + 1, 8) = Format$(DateAdd("yyyy", -1, basemapa.Cells(lin, 12)), "dd/mm/yyyy")
            
                    End If
                    baseformRetorno.Cells(ultilinformRetorno + 1, 9) = Format$(Date, "dd/mm/yyyy")
                
                End If
            End If
    
        End With
    
    Next lin
    populalbxRetorno
    frmMovimentaManutencao.cbRetornoSerie.SetFocus
   
    Set basemapa = Nothing
    Set baseformRetorno = Nothing
    Set colunaseriebaseRetorno = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "ReceberManutenção()"
    GoTo fim
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub deletaEnvios()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub deletaEnvios()
    On Error GoTo TError
    Dim lin   As Object
    Dim baseformEnvio As Range
    Dim i     As Long
    Dim ultlin As Long
    Set baseformEnvio = formenvio.Range("G9:o" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)

    For i = baseformEnvio.Rows.Count To 1 Step -1
        With baseformEnvio
        
            If .CurrentRegion.Rows.Count = 1 Then Exit Sub
            '            ultlin = .Rows.Count

            .Range(Cells(i, 1), Cells(i, 9)).ClearContents
            '            .Range("G" & i & ":O" & i).ClearContents
            '            .EntireRow.ClearContents
            

        End With

    Next
    frmMovimentaManutencao.cbEnvioSerie.ListIndex = -1
   
    frmMovimentaManutencao.ListBoxenvio.RowSource = vbNullString
    frmMovimentaManutencao.lbEnvioLocalAtual.Caption = vbNullString
    frmMovimentaManutencao.lbEnvioStatusResult.Caption = vbNullString
    '    Exit Sub
    populalboxenvio
    Set baseformEnvio = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "deletaEnvios()"
    GoTo fim
End Sub

Public Sub deletaRetornos()
    On Error GoTo TError
    Dim lin   As Object
    Dim baseformRetorno As Range
    Dim i     As Long
    Dim ultlin As Long
    Set baseformRetorno = formenvio.Range("AO9:aw" & formenvio.Cells(Rows.Count, "AO").End(xlUp).Row)

    For i = baseformRetorno.Rows.Count To 1 Step -1
        With baseformRetorno
        
            If .CurrentRegion.Rows.Count = 1 Then Exit Sub
            ultlin = .Rows.Count
            '            Debug.Print formenvio.Range("G" & i + 8 & ":O" & i + 8).Address
            '            .Range("G" & i + 8 & ":O" & i + 8).ClearContents
            .Range(Cells(i, 1), Cells(i, 9)).ClearContents
            '            .Range("G" & i & ":O" & i).ClearContents
            '            .EntireRow.ClearContents
            

        End With

    Next
    frmMovimentaManutencao.cbRetornoSerie.ListIndex = -1
   
    frmMovimentaManutencao.ListBoxRetorno.RowSource = vbNullString
    frmMovimentaManutencao.lbRetornoLocalAtual.Caption = vbNullString
    frmMovimentaManutencao.lbRetornoStatusResult.Caption = vbNullString
    '    Exit Sub
    populalbxRetorno
    Set baseformRetorno = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "deletaRetornos()"
    GoTo fim
End Sub
'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub PopulalbEnvioLocalAtual()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub PopulalbEnvioLocalAtual()
    On Error GoTo TError


    Dim ultlinmapa As Long, lin As Long, ultilinformenvio As Long
    Dim serie As String
    Dim basemapa As Range, baseformEnvio As Range


    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    serie = frmMovimentaManutencao.cbEnvioSerie.Text
    Set basemapa = MapaAtual.ListObjects(1).DataBodyRange
    
    Set baseformEnvio = formenvio.Range("G8:N" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)
    For lin = 1 To basemapa.Rows.Count
    
        With frmMovimentaManutencao
        
            ultilinformenvio = baseformEnvio.Rows.Count
            If serie = basemapa.Cells(lin, 8) Then
               
                
                'ID  Data do Registro    Série   Tipo    Capacidade  Último Teste    Última Recarga  Data de Envio
                On Error Resume Next
                   
                If UCase$(basemapa.Cells(lin, 4)) <> UCase$("Manutenção - Brigada") Then
                    frmMovimentaManutencao.lbEnvioLocalAtual.ForeColor = 255
                Else
                    frmMovimentaManutencao.lbEnvioLocalAtual.ForeColor = 32768
                
                End If
                
                If UCase$(basemapa.Cells(lin, 23)) <> UCase$("Vencido") And _
                                                                        UCase$(basemapa.Cells(lin, 23)) <> UCase$("Em Manutenção") Then
                    frmMovimentaManutencao.lbEnvioStatusResult.ForeColor = 255
                Else
                    frmMovimentaManutencao.lbEnvioStatusResult.ForeColor = 32768
                
                End If
                frmMovimentaManutencao.lbEnvioLocalAtual.Caption = UCase$(basemapa.Cells(lin, 4))
                frmMovimentaManutencao.lbEnvioStatusResult.Caption = UCase$(basemapa.Cells(lin, 23))
      
                GoTo sair:

            End If
    
        End With
    
    Next lin
sair:
    '    Application.EnableEvents = True
    '    Application.ScreenUpdating = True
    Set basemapa = Nothing
    Set baseformEnvio = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "PopulalbEnvioLocalAtual()"
    GoTo fim
End Sub

Public Sub PopulalbRetornoLocalAtual()
    On Error GoTo TError


    Dim ultlinmapa As Long, lin As Long, ultilinformRetorno As Long
    Dim serie As String
    Dim basemapa As Range, baseformRetorno As Range

    '    Application.EnableEvents = False
    '    Application.ScreenUpdating = False

    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    serie = frmMovimentaManutencao.cbRetornoSerie.Text
    Set basemapa = MapaAtual.ListObjects(1).DataBodyRange
    
    Set baseformRetorno = formenvio.Range("AO8:AW" & formenvio.Cells(Rows.Count, "AO").End(xlUp).Row)
    For lin = 1 To basemapa.Rows.Count
    
        With frmMovimentaManutencao
        
            ultilinformRetorno = baseformRetorno.Rows.Count
            If serie = basemapa.Cells(lin, 8) Then
               
                
                'ID  Data do Registro    Série   Tipo    Capacidade  Último Teste    Última Recarga  Data de Envio
                On Error Resume Next
                   
                If UCase$(basemapa.Cells(lin, 4)) <> UCase$("Manutenção - MAREFIRE") Then
                    frmMovimentaManutencao.lbRetornoLocalAtual.ForeColor = 255
                Else
                    frmMovimentaManutencao.lbRetornoLocalAtual.ForeColor = 32768
                
                End If
                
                If UCase$(basemapa.Cells(lin, 23)) <> UCase$("Em Manutenção") Then
                    frmMovimentaManutencao.lbRetornoStatusResult.ForeColor = 255
                Else
                    frmMovimentaManutencao.lbRetornoStatusResult.ForeColor = 32768
                
                End If
                frmMovimentaManutencao.lbRetornoLocalAtual.Caption = UCase$(basemapa.Cells(lin, 4))
                frmMovimentaManutencao.lbRetornoStatusResult.Caption = UCase$(basemapa.Cells(lin, 23))
      
                GoTo sair:

            End If
    
        End With
    
    Next lin
sair:
    '    Application.EnableEvents = True
    '    Application.ScreenUpdating = True
    Set basemapa = Nothing
    Set baseformRetorno = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "PopulalbEnvioLocalAtual()"
    GoTo fim
End Sub

'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub enviarecebeext()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub enviarecebeext()
    On Error GoTo TError

    Dim arr   As Variant
    Dim i     As Long
    Dim serie As String
    
    'arr = Serviços.Range("tbServicos").CurrentRegion
    arr = MapaAtual.ListObjects(1).DataBodyRange
    serie = frmMovimentaManutencao.cbEnvioSerie.Value
    For i = LBound(arr, 1) To UBound(arr, 1)
        If serie = arr(i, 8) Then
               
                
            'ID  Data do Registro    Série   Tipo    Capacidade  Último Teste    Última Recarga  Data de Envio
            On Error Resume Next
                   
                
            frmMovimentaManutencao.lbEnvioLocalAtual.Caption = arr(i, 4)
            frmMovimentaManutencao.lbEnvioStatusResult.Caption = arr(i, 23)
            Exit Sub
            '                 GoTo sair:

        End If
    Next i


fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "enviarecebeext()"
    GoTo fim
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub enviarext()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub enviarext()
    On Error GoTo TError



    If frmMovimentaManutencao.cbEnvioSerie.Value <> vbNullString Then
       

        EnviarManutenção
        populalboxenvio
      
    Else
        
        MsgBox "Insira o número de série", vbCritical, "Atenção"
        frmMovimentaManutencao.cbEnvioSerie.SetFocus
        Exit Sub
    End If

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "enviarext()"
    GoTo fim
End Sub
        
'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub selecaolistbox()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub selecaolistbox()
    On Error GoTo TError
    Dim strng As String
    Dim lCol  As Long, lRow As Long

    With frmMovimentaManutencao.ListBoxenvio '<--| refer to your listbox: change "ListBox1" with your actual listbox name
        For lRow = 0 To .ListCount - 1 '<--| loop through listbox rows
            If .Selected(lRow) Then    '<--| if current row selected
                For lCol = 0 To .ColumnCount - 1 '<--| loop through listbox columns
                    strng = strng & .List(lRow, lCol) & " | " '<--| build your output string
                Next lCol
                MsgBox "you selected" & vbCrLf & Left$(strng, (Len(strng) - 1)) '<--| show output string (after removing its last character ("|"))
                Exit For               '<-_| exit loop
            End If
        Next lRow
    End With
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "selecaolistbox()"
    GoTo fim
End Sub


Public Sub RemoverdalistaRetorno()
    On Error GoTo TError

    Dim rng   As Range
    Dim retornoArray As Variant
    Dim varTemp As Variant
    Dim serie As String
    Dim i     As Long, j As Long, c As Long
    serie = frmMovimentaManutencao.cbRetornoSerie.Text


    Set rng = formenvio.Range("AO9:AW" & formenvio.Cells(Rows.Count, "AO").End(xlUp).Row)
    retornoArray = rng
    
    For i = 1 To UBound(retornoArray)
        
        If retornoArray(i, 4) = serie Then
        
            For j = 1 To 9             'UBound(retornoArray, 1)
            
                retornoArray(i, j) = vbNullString
            Next
  
        End If
    Next
    
    rng = retornoArray
    rng.Sort rng.Columns(1), xlAscending
    

    populalbxRetorno
    frmMovimentaManutencao.cbRetornoSerie.SetFocus
    Set rng = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "RemoverdalistaRetorno()"
    GoTo fim
End Sub




'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub RemoverdalistaEnviar()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub RemoverdalistaEnviar()
    On Error GoTo TError

    Dim rng   As Range
    Dim varArray As Variant
    Dim varTemp As Variant
    Dim serie As String
    Dim i     As Long, j As Long, c As Long
    serie = frmMovimentaManutencao.cbEnvioSerie.Text


    Set rng = formenvio.Range("G9:O" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)
    varArray = rng
    
    For i = 1 To UBound(varArray)
        
        If varArray(i, 4) = serie Then
        
            For j = 1 To 9             'UBound(varArray, 1)
            
                varArray(i, j) = vbNullString
            Next
  
        End If
    Next
    
    rng = varArray
    rng.Sort rng.Columns(1), xlAscending
    
    '    Set rng = formenvio.Range("G9:O" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)
    '    varArray = rng
    '
    '    For c = 1 To UBound(varArray, 1)
    '
    '        varArray(c, 1) = c
    '
    '    Next
    '    rng = varArray
    populalboxenvio
    frmMovimentaManutencao.cbEnvioSerie.SetFocus
    Set rng = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "RemoverdalistaEnviar()"
    GoTo fim
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub bloqueiaserieduplicado()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub bloqueiaserieduplicado()
    On Error GoTo TError

    Dim rng   As Range
    Dim varArray As Variant
    Dim varTemp As Variant
    Dim serie As String
    Dim i     As Long, j As Long, c As Long
    serie = frmMovimentaManutencao.cbEnvioSerie.Value


    Set rng = formenvio.Range("G9:N" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)
    varArray = rng
    
    For i = 1 To UBound(varArray)
        
        If varArray(i, 3) = serie Then
        
            MsgBox "Este extintor já está na lista de envio", vbExclamation, "Série duplicado"
            frmMovimentaManutencao.cbEnvioSerie.SetFocus
            Exit Sub
  
        End If
    Next
    Set rng = Nothing
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "bloqueiaserieduplicado()"
    GoTo fim
End Sub




'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub PermiteInserirListaEnvio()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub PermiteInserirListaEnvio()
    On Error GoTo TError


    With frmMovimentaManutencao
        Dim resp As String
        If .lbEnvioLocalAtual.ForeColor = 255 Then
 
            resp = MsgBox("Este extintor não está na MANUTENÇÃO - BRIGADA!" _
                        & " Deseja fazer a movimentação?", vbYesNo, "Movimentação Imprória!")
  
            If resp = vbYes Then
            
                '                Application.EnableEvents = True
                
                
                Info.Activate
                Application.Wait (Now() + TimeValue("00:00:01"))
                Info.Range("frmCadastroSerie") = .cbEnvioSerie.Text
                Unload frmMovimentaManutencao
                Application.ScreenUpdating = True
                Info.Calculate
                Application.ScreenUpdating = False
                movmanut
                cadastraAtualExt
                resp = MsgBox("Movimentação Concluída! Deseja" & _
                              " retornar ao Formulário de Envios e Recebimentos?" _
                              , vbYesNo, "Continuar?")
                If resp = vbYes Then
                    '                    chamaformenvio
                    frmMovimentaManutencao.Show
                    frmMovimentaManutencao.cbEnvioSerie.SetFocus
                    DoEvents
                    If Userform_Check("frmMovimentaManutencao") = 2 Then

                        .cbEnvioSerie.Text = Info.Range("frmCadastroSerie").Value
                    Else
                        DoEvents
                        .cbEnvioSerie.Text = Info.Range("frmCadastroSerie").Value
                    End If
                    '                Application.Wait (Now() + TimeValue("00:00:01"))
                    '                .cbEnvioSerie.Text = Info.Range("frmCadastroSerie")
                Else
                
                    End
                End If
            Else
                MsgBox "Movimentação cancelada"
                Exit Sub
                '                GoTo Fim:
            
            End If
        Else
            EnviarManutenção
            
            
        End If
        '.cbEnvioSerie.Text = Info.Range("frmCadastroSerie")
    End With

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "PermiteInserirListaEnvio()"
    GoTo fim
End Sub

Public Sub MovEnvioEmBloco()
    On Error GoTo TError
    Dim tbEnvio As Range
    Dim lin   As Long
    Dim tbmov As Range, tbArmazenaEnvio As Range
    Dim resp  As String
    resp = MsgBox("Deseja efetivar a movimentação dos extintores da lista?", _
                  vbYesNo, "Cadastrar movimentação")
    If resp = vbNo Then
        frmMovimentaManutencao.cbEnvioSerie.SetFocus
    
        Exit Sub
    Else
        lin = formenvio.Range("G9:O" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row).Rows.Count
        
        Set tbEnvio = formenvio.Range("G9:O" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row)
        '    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

        Dim tbmovarray() As Variant
        Dim tbenvioArray() As Variant
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        tbenvioArray = tbEnvio
        ReDim Preserve tbenvioArray(1 To tbEnvio.Rows.Count, 1 To 9)
        
        
        '        If lin = 1 Then
        '         ReDim Preserve tbmovarray(1 To tbEnvio.Rows.Count * 2, 1 To 8)
        '        lin = 2
        '        Else
        ReDim Preserve tbmovarray(1 To tbEnvio.Rows.Count * 2, 1 To 8)
        '        End If
        c = 1
        a = 2
        For b = 1 To lin

            tbmovarray(c, 1) = DateAdd("s", b, Format(Now, "d/m/yyyy hh:mm:ss"))
            tbmovarray(c, 2) = tbenvioArray(b, 4)
            tbmovarray(c, 3) = "Saída"
            tbmovarray(c, 4) = "MANUTENÇÃO - BRIGADA"
            tbmovarray(c, 5) = "0"
            tbmovarray(c, 8) = "BRIGADA"

            tbmovarray(a, 1) = DateAdd("s", b + 1, Format(Now, "d/m/yyyy hh:mm:ss"))
            tbmovarray(a, 2) = tbenvioArray(b, 4)
            tbmovarray(a, 3) = "Entrada"
            tbmovarray(a, 6) = "MANUTENÇÃO - MAREFIRE"
            tbmovarray(a, 7) = "9999"
            tbmovarray(a, 8) = "BRIGADA"
            c = c + 2
            a = a + 2

        Next b
    
        '        Set tbArmazenaEnvio = formenvio.Range("S" & formenvio.Cells(Rows.Count, "S").End(xlUp).Offset(1, 0).Row & ":AB" & formenvio.Cells(Rows.Count, "T").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
        Set tbmov = Movimentacao.Range("G" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row & ":n" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
        '        tbArmazenaEnvio = tbmovarray
        SalvaMovEnvio
        tbmov = tbmovarray
        Set tbEnvio = Nothing
        Set tbmov = Nothing
        Set tbArmazenaEnvio = Nothing
        frmMovimentaManutencao.Hide
        
        AtualizamapaMOV
    End If
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "MovEnvioEmBloco()"
    GoTo fim
End Sub
'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub MovEnvioEmBloco()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub MovRetornoEmBloco()
    On Error GoTo TError
    Dim tbRetorno As Range
    Dim lin   As Long
    Dim tbmov As Range, tbArmazenaRetorno As Range
    Dim resp  As String
    resp = MsgBox("Deseja efetivar a movimentação dos extintores da lista?", _
                  vbYesNo, "Cadastrar movimentação")
    If resp = vbNo Then
        frmMovimentaManutencao.cbRetornoSerie.SetFocus
    
        Exit Sub
    Else
        lin = formenvio.Range("AO9:O" & formenvio.Cells(Rows.Count, "AO").End(xlUp).Row).Rows.Count
        
        Set tbRetorno = formenvio.Range("AO9:O" & formenvio.Cells(Rows.Count, "AO").End(xlUp).Row)
        '    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

        Dim tbmovarray() As Variant
        Dim tbRetornoArray() As Variant
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim d As Integer, e As Integer
        
        tbRetornoArray = tbRetorno
        ReDim Preserve tbRetornoArray(1 To tbRetorno.Rows.Count, 1 To 9)
        
        
        '        If lin = 1 Then
        '         ReDim Preserve tbmovarray(1 To tbRetorno.Rows.Count * 2, 1 To 8)
        '        lin = 2
        '        Else
        ReDim Preserve tbmovarray(1 To tbRetorno.Rows.Count * 4, 1 To 8)
        '        End If
        d = 3
        c = 1
        a = 2
        e = 4
        For b = 1 To lin

            tbmovarray(c, 1) = DateAdd("s", b, Format(Now, "d/m/yyyy hh:mm:ss"))
            tbmovarray(c, 2) = tbRetornoArray(b, 4)
            tbmovarray(c, 3) = "Saída"
            tbmovarray(c, 4) = "MANUTENÇÃO - MAREFIRE"
            tbmovarray(c, 5) = "9999"
            tbmovarray(c, 8) = "BRIGADA"

            tbmovarray(a, 1) = DateAdd("s", b + 1, Format(Now, "d/m/yyyy hh:mm:ss"))
            tbmovarray(a, 2) = tbRetornoArray(b, 4)
            tbmovarray(a, 3) = "Entrada"
            tbmovarray(a, 6) = "MANUTENÇÃO - BRIGADA"
            tbmovarray(a, 7) = "0"
            tbmovarray(a, 8) = "BRIGADA"
            
            tbmovarray(e, 1) = DateAdd("s", b, Format(Now, "d/m/yyyy hh:mm:ss"))
            tbmovarray(e, 2) = tbRetornoArray(b, 4)
            tbmovarray(e, 3) = "Saída"
            tbmovarray(e, 4) = "MANUTENÇÃO - BRIGADA"
            tbmovarray(e, 5) = "0"
            tbmovarray(e, 8) = "BRIGADA"
            
            tbmovarray(d, 1) = DateAdd("s", b + 1, Format(Now, "d/m/yyyy hh:mm:ss"))
            tbmovarray(d, 2) = tbRetornoArray(b, 4)
            tbmovarray(d, 3) = "Entrada"
            tbmovarray(d, 6) = "RESERVA TÉCNICA"
            tbmovarray(d, 7) = "1111"
            tbmovarray(d, 8) = "BRIGADA"
            
            c = c + 4
            a = a + 4
            d = d + 4
            e = e + 4

        Next b
    
        '        Set tbArmazenaRetorno = formenvio.Range("S" & formenvio.Cells(Rows.Count, "S").End(xlUp).Offset(1, 0).Row & ":AB" & formenvio.Cells(Rows.Count, "T").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
        Set tbmov = Movimentacao.Range("G" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row & ":N" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
        '        tbArmazenaRetorno = tbmovarray
        SalvaMovRetorno
        tbmov = tbmovarray
        Set tbRetorno = Nothing
        Set tbmov = Nothing
        Set tbArmazenaRetorno = Nothing
        frmMovimentaManutencao.Hide
        
        '        AtualizamapaMOV
    End If
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "MovRetornoEmBloco()"
    GoTo fim
End Sub


Public Sub SalvaMovEnvio()
    On Error GoTo TError
    Dim tbEnvio As Range
    Dim tbmov As Range, tbArmazenaEnvio As Range, valorultcel As Long
    Dim contaenvios As Long
    Dim resp  As String

    
    
    Set tbEnvio = formenvio.Range("H9:O" & formenvio.Cells(Rows.Count, "H").End(xlUp).Row)
    '    Set tbPesquisa = MapaAtual.ListObjects("tbMapaAtual").DataBodyRange

    Dim tbmovarray() As Variant
    Dim tbenvioArray() As Variant
    Dim a     As Integer
    Dim b     As Integer
    Dim c     As Integer
    tbenvioArray = tbEnvio
    ReDim Preserve tbenvioArray(1 To tbEnvio.Rows.Count, 1 To 10)
    ReDim Preserve tbmovarray(1 To tbEnvio.Rows.Count, 1 To 10)
    If formenvio.ListObjects(1).ListRows.Count = 0 Then _
                                                 formenvio.ListObjects(1).ListRows.Add
   
    
    contaenvios = formenvio.Range("U4").Value
    valorultcel = formenvio.ListObjects(1).DataBodyRange.Rows.Count
    If valorultcel = 1 Then
        valorultcel = 0
        Set tbArmazenaEnvio = formenvio.Range("S" & formenvio.Cells _
                                              (Rows.Count, "S").End(xlUp).Row & ":AB" & formenvio.Cells _
                                              (Rows.Count, "S").End(xlUp).Row + UBound(tbmovarray) - 1)

    Else
        valorultcel = formenvio.ListObjects(1).DataBodyRange.Rows.Count
        Set tbArmazenaEnvio = formenvio.Range("S" & formenvio.Cells _
                                              (Rows.Count, "S").End(xlUp).Offset(1, 0).Row & ":AB" & formenvio.Cells _
                                              (Rows.Count, "S").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
    End If

    For b = 1 To tbEnvio.Rows.Count
        tbmovarray(b, 1) = contaenvios + 1
        tbmovarray(b, 2) = b           '+ valorultcel
        tbmovarray(b, 3) = tbenvioArray(b, 1)
        tbmovarray(b, 4) = tbenvioArray(b, 2)
        tbmovarray(b, 5) = tbenvioArray(b, 3)
        tbmovarray(b, 6) = tbenvioArray(b, 4)
        tbmovarray(b, 7) = tbenvioArray(b, 5)

        tbmovarray(b, 8) = tbenvioArray(b, 6)
        tbmovarray(b, 9) = tbenvioArray(b, 7)
        tbmovarray(b, 10) = tbenvioArray(b, 8)


    Next b
    '      Set tbArmazenaEnvio = formenvio.Range("T" & formenvio.ListObjects(1).DataBodyRange.Rows.Count _
    & ":AB" & formenvio.ListObjects(1).DataBodyRange.Rows.Count + UBound(tbmovarray) - 1)
    '        Set tbmov = Movimentacao.Range("G" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row & ":N" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
    tbArmazenaEnvio = tbmovarray
    '        tbmov = tbmovarray
    Set tbEnvio = Nothing
    Set tbmov = Nothing
    Set tbArmazenaEnvio = Nothing
    With formenvio.ListObjects(1).DataBodyRange
    
        .ClearFormats
        .Columns(1).NumberFormat = "0"
        .Columns(2).NumberFormat = "0"
        .Columns(3).NumberFormat = "dd/mm/yyyy"
        .Columns(4).NumberFormat = "@"
        .Columns(5).NumberFormat = "@"
        .Columns(6).NumberFormat = "@"
        .Columns(7).NumberFormat = "@"
        .Columns(8).NumberFormat = "dd/mm/yyyy"
        .Columns(9).NumberFormat = "dd/mm/yyyy"
        .Columns(10).NumberFormat = "dd/mm/yyyy"
        .Columns.HorizontalAlignment = xlCenter
    End With
    With formenvio
'        .Range("U4").Value = contaenvios + 1
.Range("U4").Value = formenvio.ListObjects(1).DataBodyRange.Cells(formenvio.ListObjects(1).DataBodyRange.Rows.Count, 1).Value
        '    .PageSetup.RightHeader = formenvio.Name & vbNewLine & " Número: " & contaenvios
        .PageSetup.RightHeader = "&""-,Negrito""&12&K04-049Formulário de Envio" & Chr(10) & "" & Chr(10) & " Número: " & (contaenvios + 1)
        .PageSetup.CenterFooter = "&""-,Negrito""&12&K04-049" & "Carimbo/Assinatura" & Chr(10) & "" & Chr(10) & "___________________"
    End With
    
    
    
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "MovEnvioEmBloco()"
    GoTo fim
End Sub

Public Sub SalvaMovRetorno()
    On Error GoTo TError
    Dim tbRetorno As Range
    Dim tbmov As Range, tbArmazenaRetorno As Range, valorultcel As Long
    Dim contaretornos As Long
    Dim resp  As String

    
    
    Set tbRetorno = formenvio.Range("AP9:AW" & formenvio.Cells(Rows.Count, "AP").End(xlUp).Row)
  

    Dim tbmovarray() As Variant
    Dim tbRetornoArray() As Variant
    Dim a     As Integer
    Dim b     As Integer
    Dim c     As Integer
    tbRetornoArray = tbRetorno
    ReDim Preserve tbRetornoArray(1 To tbRetorno.Rows.Count, 1 To 10)
    ReDim Preserve tbmovarray(1 To tbRetorno.Rows.Count, 1 To 10)
    If formenvio.ListObjects(1).ListRows.Count = 0 Then _
                                                 formenvio.ListObjects(1).ListRows.Add
   
    
    contaretornos = formenvio.Range("U4").Value
    valorultcel = formenvio.ListObjects(1).DataBodyRange.Rows.Count
    If valorultcel = 1 Then
        valorultcel = 0
        Set tbArmazenaRetorno = formenvio.Range("S" & formenvio.Cells _
                                                (Rows.Count, "S").End(xlUp).Row & ":AB" & formenvio.Cells _
                                                (Rows.Count, "S").End(xlUp).Row + UBound(tbmovarray) - 1)

    Else
        valorultcel = formenvio.ListObjects(1).DataBodyRange.Rows.Count
        Set tbArmazenaRetorno = formenvio.Range("S" & formenvio.Cells _
                                                (Rows.Count, "S").End(xlUp).Offset(1, 0).Row & ":AB" & formenvio.Cells _
                                                (Rows.Count, "S").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
    End If

    For b = 1 To tbRetorno.Rows.Count
        tbmovarray(b, 1) = contaretornos + 1
        tbmovarray(b, 2) = b           '+ valorultcel
        tbmovarray(b, 3) = tbRetornoArray(b, 1)
        tbmovarray(b, 4) = tbRetornoArray(b, 2)
        tbmovarray(b, 5) = tbRetornoArray(b, 3)
        tbmovarray(b, 6) = tbRetornoArray(b, 4)
        tbmovarray(b, 7) = tbRetornoArray(b, 5)

        tbmovarray(b, 8) = tbRetornoArray(b, 6)
        tbmovarray(b, 9) = tbRetornoArray(b, 7)
        tbmovarray(b, 10) = tbRetornoArray(b, 8)


    Next b
    '      Set tbArmazenaRetorno = formenvio.Range("T" & formenvio.ListObjects(1).DataBodyRange.Rows.Count _
    & ":AB" & formenvio.ListObjects(1).DataBodyRange.Rows.Count + UBound(tbmovarray) - 1)
    '        Set tbmov = Movimentacao.Range("G" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row & ":N" & Movimentacao.Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Row + UBound(tbmovarray) - 1)
    tbArmazenaRetorno = tbmovarray
    '        tbmov = tbmovarray
    Set tbRetorno = Nothing
    Set tbmov = Nothing
    Set tbArmazenaRetorno = Nothing
    With formenvio.ListObjects(1).DataBodyRange
    
        .ClearFormats
        .Columns(1).NumberFormat = "0"
        .Columns(2).NumberFormat = "0"
        .Columns(3).NumberFormat = "dd/mm/yyyy"
        .Columns(4).NumberFormat = "@"
        .Columns(5).NumberFormat = "@"
        .Columns(6).NumberFormat = "@"
        .Columns(7).NumberFormat = "@"
        .Columns(8).NumberFormat = "dd/mm/yyyy"
        .Columns(9).NumberFormat = "dd/mm/yyyy"
        .Columns(10).NumberFormat = "dd/mm/yyyy"
        .Columns.HorizontalAlignment = xlCenter
    End With
    With formenvio
        .Range("U4").Value = contaretornos + 1
        '    .PageSetup.RightHeader = formenvio.Name & vbNewLine & " Número: " & contaretornos
        .PageSetup.RightHeader = "&""-,Negrito""&12&K04-049Formulário de Envio" & Chr(10) & "" & Chr(10) & " Número: " & (contaretornos + 1)
        .PageSetup.CenterFooter = "&""-,Negrito""&12&K04-049" & "Carimbo/Assinatura" & Chr(10) & "" & Chr(10) & "___________________"
    End With
    
    
    
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "MovEnvioEmBloco()"
    GoTo fim
End Sub

Public Sub populalbxRetorno()
    On Error GoTo TError

    '    Dim lo    As ListObject
    Dim lin   As Long
    Dim baseformRetorno As Range
    '    Set baseformEnvio = formenvio.Range("G9:o" & formenvio.Cells(Rows.Count, "G").End(xlUp).Row).Rows
    
    lin = formenvio.Range("AO8").CurrentRegion.Rows.Count + 8
    If lin = 9 Then lin = 10
    Set baseformRetorno = formenvio.Range(formenvio.Cells(9, 41), _
                                          formenvio.Cells(lin, 49))
    
    With frmMovimentaManutencao.ListBoxRetorno
    
        '    .RowSource = vbNullString
        .ColumnCount = baseformRetorno.Columns.Count
        .ColumnHeads = True
        .Font.Size = 10
        .BackColor = RGB(52, 108, 135)
        .ForeColor = vbWhite
        .Font.Bold = True
        .RowSource = baseformRetorno.Address(external:=True)
    
    
    End With
    
    
    

    Set baseformRetorno = Nothing

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "populalbRetorno()"
    GoTo fim
End Sub

Public Sub PermiteInserirListaRetorno()
    On Error GoTo TError


    With frmMovimentaManutencao
        Dim resp As String
        If .lbRetornoLocalAtual.ForeColor = 255 Then 'verifica se a fonte é vermelha
 
            resp = MsgBox("Este extintor não está na MANUTENÇÃO - MAREFIRE!" _
                        & " Deseja fazer a movimentação?", vbYesNo, "Movimentação Imprória!")
  
            If resp = vbYes Then
            
                '                Application.EnableEvents = True
                
                
                Info.Activate
                Application.Wait (Now() + TimeValue("00:00:01"))
                Info.Range("frmCadastroSerie") = .cbRetornoSerie.Text 'chama form cadastro ext
                Unload frmMovimentaManutencao 'descarrega form envio Retorno
                Application.ScreenUpdating = True
                Info.Calculate
                Application.ScreenUpdating = False
                movmanutEXTERNA        ' insere dados de movimentação para MANUTENÇÃO - MAREFIRE
                cadastraAtualExt       ' efetiva movimentação
                resp = MsgBox("Movimentação Concluída! Deseja" & _
                              " retornar ao Formulário de Envios e Recebimentos?" _
                              , vbYesNo, "Continuar?")
                If resp = vbYes Then
                    '                    chamaformenvio
                    frmMovimentaManutencao.Show ' carrega form envio recebimento

                    frmMovimentaManutencao.btMenuRetorno_Click ' chama form/frame recebimento
                    frmMovimentaManutencao.cbRetornoSerie.SetFocus
                    DoEvents
                    If Userform_Check("frmMovimentaManutencao") = 2 Then

                        .cbRetornoSerie.Text = Info.Range("frmCadastroSerie").Value
                    Else
                        DoEvents
                        .cbRetornoSerie.Text = Info.Range("frmCadastroSerie").Value
                    End If

                Else
                
                    End
                End If
            Else
                MsgBox "Movimentação cancelada"
                Exit Sub
                '                GoTo Fim:
            
            End If
        Else
            ReceberManutenção
            
            
        End If

    End With

fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "PermiteInserirListaRetorno()"
    GoTo fim
End Sub

Public Sub MovReservaManutLote()
    Dim ultlinmapa As Long, lin As Long, ultilinmov As Long, i As Long
    Dim serie As String
    Dim basemapa As Range, baseMov As Range, colunaseriebasemapa As Range
    Dim cell  As Variant
    Dim maparray() As Variant, movarray() As Variant
    ultilinmov = Movimentacao.ListObjects(1).DataBodyRange.Rows.Count + 9
    
    ultlinmapa = MapaAtual.ListObjects(1).DataBodyRange.Rows.Count
    Set basemapa = MapaAtual.ListObjects(1).DataBodyRange

    ReDim Preserve maparray(1 To ultlinmapa, 1 To basemapa.Columns.Count)
    
    
    maparray = basemapa
    Dim a     As Long, b As Long, j As Long, k As Long, z As Long
    a = 1
    b = 2
    j = 2
    k = 8
    z = 1
    For i = 1 To ultlinmapa
        If maparray(i, 3) = "RESERVA TÉCNICA" _
                            And (maparray(i, 23) = "Vencido" Or _
                                 maparray(i, 23) = "Vencendo") Then
            z = 1
            ReDim Preserve movarray(1 To j, 1 To k) As Variant
            movarray(a, 1) = Format(DateAdd("s", z, Now), "d/m/yy hh:mm:ss")
            movarray(a, 2) = maparray(i, 8)
            movarray(a, 3) = "Saída"
            movarray(a, 4) = "RESERVA TÉCNICA"
            movarray(a, 5) = "1111"
            movarray(a, 8) = "BRIGADA"
            z = z + 1
            movarray(b, 1) = Format(DateAdd("s", z, Now), "d/m/yy hh:mm:ss")
            movarray(b, 2) = maparray(i, 8)
            movarray(b, 3) = "Entrada"
            movarray(b, 6) = "MANUTENÇÃO - BRIGADA"
            movarray(b, 7) = "0"
            movarray(b, 8) = "BRIGADA"

            movarray = Application.Transpose(movarray)
     
        
            j = j + 2
        
            ReDim Preserve movarray(1 To k, 1 To j)
        

            movarray = Application.Transpose(movarray)

        
            a = a + 2
            b = b + 2

        End If
    Next i
    If i > 1 And a = 1 Then Exit Sub
    movarray = Application.Transpose(movarray)
      
    j = j - 2
        
    ReDim Preserve movarray(1 To k, 1 To j)
    movarray = Application.Transpose(movarray)
    Set baseMov = Movimentacao.Range("G" & ultilinmov & ":N" & UBound(movarray) + ultilinmov)
    baseMov = movarray
    ultilinmov = Movimentacao.ListObjects(1).DataBodyRange.Rows.Count + 9
    baseMov.ListObject.ListRows(ultilinmov - 9).Delete
   
    '    Stop


    Set basemapa = Nothing
    Set baseMov = Nothing
    
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "MovReservaManutLote()"
    GoTo fim
End Sub




Sub testtranspose()

    Dim n, m  As Integer
    n = 2
    m = 1
    Dim arrCity() As Variant
    ReDim arrCity(1 To n, 1 To m)

    m = m + 1
    ReDim Preserve arrCity(1 To n, 1 To m)
    arrCity = Application.Transpose(arrCity)
    n = n + 1
    ReDim Preserve arrCity(1 To m, 1 To n)
    arrCity = Application.Transpose(arrCity)
End Sub

