Attribute VB_Name = "MODAtualixaext"
Option Explicit
'@Folder("SGES2020")

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub cadastraAtualExt()
    ' Data......: 18/11/2020
    ' Descricao.: Atualiza informações sobre o extintor de incêndio
    '---------------------------------------------------------------------------------------
Public Sub cadastraAtualExt()
Attribute cadastraAtualExt.VB_ProcData.VB_Invoke_Func = "s\n14"
        On Error GoTo TError
    Dim ultlinha As Long
    Dim ultlinhaLog As Long
    Dim endSaida As Variant
    Dim endEntrada As Variant
    Dim cell  As Range
    Dim i     As Integer
    Dim contvazio As Long
    Dim LINMAPAATUAL As Long
    Dim LINEXTINTOR As Long
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    With Info
        For Each cell In .Range("$F$9:$F$21")
            If cell.Value = "Sim" Then

                GoTo exec:
              
            End If
        Next
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
exec:
        Application.EnableEvents = False
        Application.ScreenUpdating = False
If .Range("M10").Value = "1K" Then
.Unprotect
PINTAVAZIOIK
  If .Range("F20").Value = "Sim" Then
                    UPDATETBLSERV
                    updateservmapa
'                      Atualizamapaserv
                    
                End If
                
            

                '   UPDATESTATUSGERAL
                
                If .Range("F12").Value = "Sim" Then ' se movimentação foi alterada
                    SUPERMOV
                End If
               
                If .Range("F11").Value = "Sim" Then 'se ext foi alterado
                    updatetblext
                End If
               
                'Local OBS
                
                If .Range("F19").Value = "Sim" Or .Range("F21").Value = "Sim" Then
                    cadastraObsext
                    cadastraObslocal
                    ATUALIZAmAPAObs
                End If
                GoTo FINAL:
                

End If
        If .Range("M8").Value = "CO" Then
            
            
            If .Range("$M$10").Value <> "45K" And .Range("$M$10").Value <> "34K" Then
                .Unprotect

                'pintavazio
                'VERIFICA SE ALGUM CAMPO ESTÁ VAZIO ANTES DE SALVAR  INFO
                PINTAVAZIONORMAL
                
                If .Range("F20").Value = "Sim" Then
                    UPDATETBLSERV
                    updateservmapa
'                      Atualizamapaserv
                    
                End If
                
            

                '   UPDATESTATUSGERAL
                
                If .Range("F12").Value = "Sim" Then ' se movimentação foi alterada
                    SUPERMOV
                End If
               
                If .Range("F11").Value = "Sim" Then 'se ext foi alterado
                    updatetblext
                End If
               
                'Local OBS
                
                If .Range("F19").Value = "Sim" Or .Range("F21").Value = "Sim" Then
                    cadastraObsext
                    cadastraObslocal
                    ATUALIZAmAPAObs
                End If
                

                GoTo FINAL:

            Else
            
            
                'CILINDROS

                .Unprotect
    
                'pintavazio
                'VERIFICA SE ALGUM CAMPO ESTÁ VAZIO ANTES DE SALVAR  INFO
                PINTAVAZIOCILINDRO
                            
    
                If .Range("F20").Value = "Sim" Then
                    UPDATETBLSERV
                    updateservmapa
                
                        
    
                End If
                
                
    
                ' UPDATESTATUSGERAL
    
               
                If .Range("F12").Value = "Sim" Then ' se movimentação foi alterada
                    SUPERMOV
                End If
              
    
               
                If .Range("F11").Value = "Sim" Then 'se ext foi alterado
                    updatetblext
                End If
               
    
    
    
                'Local OBS
            
                If .Range("F19").Value = "Sim" Or .Range("F21").Value = "Sim" Then
                    cadastraObsext
                    cadastraObslocal
                    ATUALIZAmAPAObs
                End If
              
            End If
            GoTo FINAL:
            

            
        Else                           '######## DIFERENTE DE CO OU FM



            Application.EnableEvents = False
            Application.ScreenUpdating = False
            .Unprotect
            'VERIFICA SE ALGUM CAMPO , EXCETO PESAGEM, ESTÁ VAZIO ANTES DE SALVAR  INFO
            For Each cell In .Range("I8,M8,I12,M10,M12,I14,M14,I16,M16,M18,M20,I20")

                cell.Interior.Color = &HF9F9F9
                cell.ClearComments
                If cell.Value = vbNullString Then
                    cell.Interior.Color = &HC0C0FF 'pinta campos vazios
                    cell.AddComment
                    cell.Comment.Visible = True
                    cell.Comment.Shape.TextFrame.Characters.Font.Bold = True
                    cell.Comment.Shape.TextFrame.Characters.Font.Size = 12
                    cell.Comment.Shape.TextFrame.Characters.Font.Color = &HCC
                    cell.Comment.Text Text:="SGES:" & Chr$(10) & "Preencha todos os campos!!!"
                    'CELL.comment.Text "Preencha os campos"
                    contvazio = contvazio + 1
                End If
            Next cell
            If contvazio > 0 Then
                Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
                MsgBox "Há campos vazios no formulário! Preencha todos os campos!"
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                Exit Sub
            Else
        
        
                If .Range("F20").Value = "Sim" Then
                    UPDATETBLSERV
                    updateservmapa
                    
'                      Atualizamapaserv
                    
    
                End If
          
                '  UPDATESTATUSGERAL
                If .Range("F12").Value = "Sim" Then ' se movimentação foi alterada
                    SUPERMOV
                End If
        
        
                'Extintores
        
                If .Range("F11").Value = "Sim" Then 'se ext foi alterado
                    updatetblext
        
                End If
        
        
        
                'Local OBS
                If .Range("F19").Value = "Sim" Or .Range("F21").Value = "Sim" Then
                    cadastraObsext
                    cadastraObslocal
                    ATUALIZAmAPAObs
                End If
    
                ' GoTo FINAL:
            End If
        End If
FINAL:
        .Range("F11:F21").ClearContents
        
    End With


    Info.Protect

    Application.EnableEvents = True
    Application.ScreenUpdating = True


    formatatbhistmov
    dimbarra
    dimbtnsalvaext
    Application.Speech.Speak "Atualização concluída", speakasync:=True

    MsgBox "Atualização concluída!", , "Concluído"
    Info.Range("i8").Select
    Info.Range("i8").Value = Info.Range("i8").Value


    

fim:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub cadastraExtNovo()
    ' Data......: 31/12/2020
    ' Descricao.: Cadastrar novo extintor
    '---------------------------------------------------------------------------------------
Public Sub cadastraExtNovo()
    On Error GoTo TError
    Dim ultlinha As Long
    Dim LINMAPAATUAL As Long
    Dim LINEXT As Long
    Dim SerieEXT As Range
    Dim serieinfo As Range
    Dim resultado As VbMsgBoxResult

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim serieadapt As String
    With Info
        .Unprotect
        'VERIFICA SE ALGUM CAMPO ESTÁ VAZIO ANTES DE SALVAR  INFO
        Dim cell As Range
        Dim contvazio As Long
        For Each cell In .Range("I37,M37,I39,M39,I41,M41,I43,M43,I45,I47,M45,M47,I49,M49")
            cell.Interior.Color = &HF9F9F9


        Next cell
        If .Range("m37").Value = "CO" Or .Range("m37").Value = "FM" Then
            For Each cell In .Range("I37,M37,I39,M39,I41,M41,I43,M43,I45,I47,M45,M47,I49,M49")
                'CELL.Interior.Color = &HF9F9F9

                If cell.Value = vbNullString Then
                    cell.Interior.Color = &HC0C0FF 'pinta campos vazios
                    contvazio = contvazio + 1
                End If
            Next cell
        Else
            For Each cell In .Range("I37,M37,I39,M39,I41,M41,I43,M43,I45,M45,M47,I49,M49")
                'CELL.Interior.Color = &HF9F9F9

                If cell.Value = vbNullString Then
                    cell.Interior.Color = &HC0C0FF 'pinta campos vazios
                    contvazio = contvazio + 1
                End If
            Next cell
        End If
        If contvazio > 0 Then

            Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
            MsgBox "Há campos vazios no formulário! Preencha todos os campos!", , "Atenção!": Exit Sub


        End If
        serieadapt = Info.Range("I35").Value

        LINEXT = 9
        Info.Cells(35, 9) = Info.Cells(37, 9) & Info.Cells(37, 13) & Info.Cells(39, 13)
        Set serieinfo = Info.Cells(35, 9)
        Do Until Extintores.Cells(LINEXT, 15) = vbNullString


            For Each SerieEXT In serieinfo
                Do Until Extintores.Cells(LINEXT, 15) = vbNullString
                    Set SerieEXT = Extintores.Cells(LINEXT, 15)


                    If serieinfo.Value = SerieEXT.Value Then

                        Application.Speech.Speak "Este número de série já existe. Favor inserir um novo número de série", speakasync:=True
                        MsgBox "Este número de série já existe. Favor inserir um novo número de série."

                        Info.Range("I37").ClearContents
                        Info.Range("i37").Activate
                        Info.Calculate: Exit Sub

                    Else

                        LINEXT = LINEXT + 1
                    End If


                Loop

            Next SerieEXT

        Loop

        'movimentação
        limpafiltrosmov
        ultlinha = Movimentacao.Range("tbCadastroMovimentacao[[#Headers],[Data]]").End(xlDown).Offset(1, 0).Row
        Movimentacao.Range("G" & ultlinha).Value = Now
        Movimentacao.Range("H" & ultlinha).Value = .Range("I35").Value 'Série
        Movimentacao.Range("I" & ultlinha).Value = "Entrada"
        Movimentacao.Range("J" & ultlinha).Value = vbNullString
        Movimentacao.Range("K" & ultlinha).Value = vbNullString
        Movimentacao.Range("L" & ultlinha).Value = .Range("M41").Value 'Local entrada
        Movimentacao.Range("M" & ultlinha).Value = .Range("I43").Value 'área
        Movimentacao.Range("N" & ultlinha).Value = .Range("M43").Value 'Zona
        Movimentacao.Range("tbHistMov14").Calculate
        'Extintores
        limpafiltrosext
        ultlinha = Extintores.Range("tbExtintores[[#Headers],[Série]]").End(xlDown).Offset(1, 0).Row
        Extintores.Range("G" & ultlinha).Value = .Range("I37").Value 'serie
        Extintores.Range("J" & ultlinha).Value = .Range("I39").Value 'Fabricação
        Extintores.Range("H" & ultlinha).Value = .Range("M37").Value 'tipo
        Extintores.Range("I" & ultlinha).Value = .Range("M39").Value 'capacidade
        Extintores.Range("K" & ultlinha).Value = .Range("I41").Value 'suporte
        Extintores.Range("L" & ultlinha).Value = .Range("G52").Value 'Observação
        Extintores.Range("N" & ultlinha).Value = Now
        Extintores.Range("G" & ultlinha).EntireRow.Calculate

        'SERVIÇOS
        limpafiltrosservico
        ultlinha = Serviços.Range("tbServicos[[#Headers],[Data]]").End(xlDown).Offset(1, 0).Row
        Serviços.Range("G" & ultlinha).Value = Now 'Data
        Serviços.Range("H" & ultlinha).Value = CStr(.Range("I35").Value) 'Série
        Serviços.Range("I" & ultlinha).Value = .Range("M37").Value 'Tipo
        Serviços.Range("J" & ultlinha).Value = .Range("I45").Value 'Teste
        Serviços.Range("L" & ultlinha).Value = .Range("M45").Value 'Recarga
        If .Range("M37").Value = "CO" Or .Range("M37").Value = "FM" Then
            Serviços.Range("N" & ultlinha).Value = .Range("I47").Value 'Pesagem
        End If
        Serviços.Range("P" & ultlinha).Value = .Range("M47").Value 'Selo
        Serviços.Range("R" & ultlinha).Value = .Range("I49").Value 'Inspeção
        Serviços.Range("T" & ultlinha).Value = .Range("M49").Value 'Pintura
        PreviServ
        Serviços.Range("tbHistServ13").Calculate
        Serviços.Range("tbHistServ1327").Calculate

        'Mapa atual
        limpafiltrosmapaatual
        ultlinha = MapaAtual.Range("tbMapaAtual[[#Headers],[Série]]").End(xlDown).Offset(1, 0).Row
        MapaAtual.Range("N" & ultlinha).Value = .Range("I35").Value 'SÉRIE
        MapaAtual.Range("G" & ultlinha).Value = .Range("I41").Value 'SUPORTE
        MapaAtual.Range("H" & ultlinha).Value = .Range("I43").Value 'ÁREA
        MapaAtual.Range("I" & ultlinha).Value = .Range("M41").Value 'EDIFICIO
        MapaAtual.Range("J" & ultlinha).Value = .Range("M41").Value 'LOCAL
        MapaAtual.Range("K" & ultlinha).Value = .Range("M37").Value 'TIPO
        MapaAtual.Range("L" & ultlinha).Value = .Range("M39").Value 'CAP
        MapaAtual.Range("M" & ultlinha).Value = .Range("I39").Value 'FAB
        MapaAtual.Range("O" & ultlinha).Value = .Range("M43").Value 'ZONA
        MapaAtual.Range("P" & ultlinha).Value = DateAdd("yyyy", 5, .Range("I45").Value) 'TESTE
        If .Range("M37").Value = "CO" Or .Range("M37").Value = "FM" Then
            MapaAtual.Range("R" & ultlinha).Value = DateAdd("yyyy", 5, .Range("M45").Value) 'Recarga
        Else
            MapaAtual.Range("R" & ultlinha).Value = DateAdd("yyyy", 1, .Range("M45").Value) 'Recarga outros
        End If
        If .Range("M37").Value = "CO" Then
            MapaAtual.Range("T" & ultlinha).Value = DateAdd("m", 6, .Range("I47").Value) 'Pesagem
        End If
        If .Range("M37").Value = "FM" Then
            MapaAtual.Range("T" & ultlinha).Value = DateAdd("yyyy", 5, .Range("I47").Value) 'Pesagem FM
        End If
        MapaAtual.Range("V" & ultlinha).Value = DateAdd("yyyy", 1, .Range("M47").Value) 'Selo
        MapaAtual.Range("X" & ultlinha).Value = DateAdd("m", 6, .Range("I49").Value) 'Inspeção
        MapaAtual.Range("Z" & ultlinha).Value = DateAdd("yyyy", 5, .Range("M49").Value) 'Pintura
        MapaAtual.Range("AA" & ultlinha).Value = .Range("G52").Value 'OBS
        
        MapaAtual.Range("Q" & ultlinha).Value = "TESTE EM DIA" 'STATUS TESTE
        MapaAtual.Range("S" & ultlinha).Value = "RECARGA EM DIA"
        If .Range("M37").Value = "CO" Then
            MapaAtual.Range("U" & ultlinha).Value = "PESAGEM EM DIA"
        Else
            MapaAtual.Range("U" & ultlinha).Value = "NÃO APLICÁVEL"
        End If
        MapaAtual.Range("W" & ultlinha).Value = "SELO EM DIA"
        MapaAtual.Range("Y" & ultlinha).Value = "INSPEÇÃO EM DIA"
        MapaAtual.Range("AC" & ultlinha).Value = "Em Dia"
     

        .Range("E57").Value = .Range("I35").Value
        .Calculate
        Application.Speech.Speak "Extintor cadastrado com sucesso! Deseja cadastrar um novo extintor?", speakasync:=True
        resultado = MsgBox(" Extintor cadastrado com sucesso! Deseja cadastrar um novo extintor?", vbYesNo, "Cadastro conclúido")
        If resultado = vbYes Then
            .Range("I37").ClearContents
            .Range("I37").Activate
        Else
            frmAtualiza
            .Range("I8").Value = serieadapt
            .Range("M12:N12").Select
        End If
        .Protect


    End With


    Application.EnableEvents = True
    Application.ScreenUpdating = True
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub salvaNovoLocalnovo()
    ' Data......: 31/12/2020
    ' Descricao.: cadastra novo local a partir do form Cadastro - Novo Extintor
    '---------------------------------------------------------------------------------------
Public Sub salvaNovoLocalnovo()
    On Error GoTo TError



    ' VERIFICA SE ALGUM CAMPO ESTÁ VAZIO
    Dim cell  As Range
    Dim contvazio As Long
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each cell In Info.Range("I69,N67,I67")
        On Error GoTo ERRO
        If cell.Value = Empty Then
            contvazio = contvazio + 1
        End If
    Next cell
    With Info
        .Unprotect
        If contvazio > 0 Then
            Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
            MsgBox "Há campos vazios no formulário! Preencha todos os campos!", , "Atenção!"
        Else
            .Shapes("btnCancelarLocalAtual").Visible = msoFalse
            .Shapes("btnSalvaLocalAtual").Visible = msoFalse
            .Shapes("btnCancelarLocalNovo").Visible = msoFalse
            .Shapes("btnSalvaLocalNovo").Visible = msoFalse
            .Shapes("btnSalvaNovoExt").Visible = msoCTrue
            .Shapes("btnCancelarNovoExt").Visible = msoCTrue
            .Shapes("btnSalvaAtualExt").Visible = msoFalse
            .Shapes("btnLocalAdd").Visible = msoFalse
            .Shapes("btnLocalAdd2").Visible = msoCTrue
            .Shapes("btnLocalAdd2").Top = 143.3667
            .Shapes("btnExtAdd").Visible = msoFalse
            Dim ultlinha As Variant

            limpafiltrolocal
            ultlinha = locais.Range("tbLocalNovo[[#Headers],[Zona]]").End(xlDown).Offset(1, 0).Row

            locais.Range("G" & ultlinha).Value = .Range("I69").Value 'zona
            locais.Range("H" & ultlinha).Value = .Range("I67").Value 'local
            locais.Range("I" & ultlinha).Value = .Range("N67").Value 'área
            locais.Range("J" & ultlinha).Value = .Range("I67").Value & " - " & .Range("N67").Value

            atualizatbdinlocal         ' ATUALIZA ATABELA DINAMICA LOCAIS NA PLANILHA LOCAIS

            '   ADICIONA LOCAL NOVO NA TBL MAPAATUAL

            ultlinha = MapaAtual.Range("tbMapaAtual[[#Headers],[Série]]").End(xlDown).Offset(1, 0).Row
            MapaAtual.Range("O" & ultlinha).Value = .Range("I69").Value 'zona
            MapaAtual.Range("J" & ultlinha).Value = .Range("I67").Value 'local
            MapaAtual.Range("H" & ultlinha).Value = .Range("N67").Value 'área
            
            
            .Range("M41:N41,I43,M43:N43").ClearContents
            .Range("M43:N43").Value = .Range("I69:K69").Value 'zona
            .Range("M41:N41").Value = .Range("I67:K67").Value 'local
            .Range("I43").Value = .Range("N67").Value 'área
            .Range("I67:K67,I69:K69,N67").ClearContents
            .Range("C31:C58").EntireRow.Hidden = False 'esconde frmnovo
            .Range("C2:C30").EntireRow.Hidden = True 'exibe frmatualiza
            .Range("C90:C120").EntireRow.Hidden = True 'esconde frmnovolocalnovo
            .Range("C59:C89").EntireRow.Hidden = True 'esconde frmnovolocalatual
            .Range("I43").Activate

        End If
        .Protect
    End With

    Exit Sub
ERRO:
    Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
    MsgBox "Prencha todos os campos do formulário!"

    Application.EnableEvents = True
    Application.ScreenUpdating = True
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub salvaNovoLocalAtualiza()
    ' Data......: 31/12/2020
    ' Descricao.: cadastra novo local a partir do form Cadastro - Atualização
    '---------------------------------------------------------------------------------------
Public Sub salvaNovoLocalAtualiza()
    On Error GoTo TError
    Application.EnableEvents = False
    Application.ScreenUpdating = False


    ' On Error GoTo ERRO
    ' VERIFICA SE ALGUM CAMPO ESTÁ VAZIO
    Dim cell  As Range
    Dim contvazio As Long
    For Each cell In Info.Range("I103,N103,I105")
        If cell.Value = Empty Then
            contvazio = contvazio + 1
        End If
    Next cell
    With Info
        .Unprotect
        If contvazio > 0 Then
            Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
            MsgBox "Há campos vazios no formulário! Preencha todos os campos!"
        Else
            .Shapes("btnCancelarLocalAtual").Visible = msoFalse
            .Shapes("btnSalvaLocalAtual").Visible = msoFalse
            .Shapes("btnCancelarLocalNovo").Visible = msoFalse
            .Shapes("btnSalvaLocalNovo").Visible = msoFalse
            .Shapes("btnSalvaNovoExt").Visible = msoFalse
            .Shapes("btnCancelarNovoExt").Visible = msoFalse
            .Shapes("btnSalvaAtualExt").Visible = msoCTrue
            .Shapes("btnLocalAdd").Visible = msoCTrue
            .Shapes("btnExtAdd").Visible = msoCTrue
            '            .Shapes("btnImprime").Height = 28.07
            .Shapes("btnextadd").Width = 37.38
            .Shapes("btnextadd").Height = 39.7
            Dim ultlinha As Variant
            limpafiltrolocal
            ultlinha = locais.Range("tbLocalNovo[[#Headers],[Zona]]").End(xlDown).Offset(1, 0).Row

            locais.Range("G" & ultlinha).Value = UCase$(.Range("I105").Value) 'zona
            locais.Range("H" & ultlinha).Value = UCase$(.Range("I103").Value) 'local
            locais.Range("I" & ultlinha).Value = UCase$(.Range("N103").Value) 'área
            locais.Range("J" & ultlinha).Value = UCase$(Info.Range("i103").Value) & " - " & UCase$(Info.Range("n103"))

            atualizatbdinlocal         ' ATUALIZA ATABELA DINAMICA LOCAIS NA PLANILHA LOCAIS
 
            ultlinha = MapaAtual.Range("tbMapaAtual[[#Headers],[Série]]").End(xlDown).Offset(1, 0).Row
            MapaAtual.Range("O" & ultlinha).Value = UCase$(.Range("I105").Value) 'zona
            MapaAtual.Range("J" & ultlinha).Value = UCase$(.Range("I103").Value) 'local
            MapaAtual.Range("H" & ultlinha).Value = UCase$(.Range("N103").Value) 'área
            
            .Range("I14,M12:N12,M14:N14").ClearContents
            .Range("M14").Value = .Range("I105:K105").Value 'zona
            .Range("M12:N12").Value = .Range("I103:K103").Value 'local
            .Range("I14").Value = .Range("N103").Value 'área


            .Range("I103:K103,N103,I105:K105").ClearContents

            cancelaNovoLocalAtualiza

            .Range("frmNovoExtintorSerie").Activate

        End If
        .Protect
    End With
    Exit Sub
    'ERRO:
    'Application.Speech.Speak "Há campos vazios no formulário! Preencha todos os campos!", speakasync:=True
    '    MsgBox "Ocorreu um erro! Contacte o administrador do sistema"
    '    Exit Sub

    Application.EnableEvents = True
    Application.ScreenUpdating = True
fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

