Attribute VB_Name = "MODExibeforms"
Option Explicit
'@Folder("SGES2020")

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Autor: Warley da Silva Conceição - Rotina: Public Sub frmNovo()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub frmNovo()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    Dim cell  As Range
    With Info
        '.Activate
       
        '        .Shapes("btnExtAdd").Visible = msoFalse
        '        .Shapes("Scroll Bar 26").Visible = msoFalse
        '        .Shapes("Scroll Bar 48").Visible = msoFalse
        '        .Shapes("btnCancelarLocalAtual").Visible = msoFalse
        '        .Shapes("btnSalvaLocalAtual").Visible = msoFalse
        '        .Shapes("btnCancelarLocalNovo").Visible = msoFalse
        '        .Shapes("btnSalvaLocalNovo").Visible = msoFalse
        '        .Shapes("btnImprime").Visible = msoFalse
        '
        '        .Shapes("btnSalvaAtualExt").Visible = msoFalse
        '        .Shapes("btnLocalAdd").Visible = msoFalse
        escondeobjfrmatualiza
        exibeobjfrmnovoext
        .Activate
        .Range("frmNovoExtintorSerie").Activate
        .Range("C31:C58").EntireRow.Hidden = False
        .Range("C2:C30").EntireRow.Hidden = True
        .Range("C59:C89").EntireRow.Hidden = True
        .Range("C90:C129").EntireRow.Hidden = True
        dimbtnsalvaextnovo


        .Shapes("btnSalvaNovoExt").Top = .Range("P55").Top + .Range("OBSNOVO").Height + 5

        dimbtncancelnovoext
        .Shapes("btnCancelarNovoExt").Top = .Range("P55").Top + .Range("OBSNOVO").Height + 5

        .Shapes("btnLocalAdd2").Visible = msoCTrue

        .Shapes("btnLocalAdd2").Top = 143.3667
        .Shapes("btnLocalAdd2").Width = 30
        .Shapes("btnLocalAdd2").Height = 30
        For Each cell In .Range("I37,M37,I39,M39,I41,M41,I43,M43,I45,I47,M45,M47,I49,M49")
            cell.Interior.Color = &HF9F9F9
            cell.Value = vbNullString

        Next cell
        '.Range("I37,M37,I39,M39,I41,M41:N41,I43,M43:N43,I45,M45,I47,M47,M49,I49,G52:N55").ClearContents
        .PageSetup.PrintArea = vbNullString
        .PageSetup.PrintArea = "$F$32:$W$57"
        .Range("M41").Value = "RESERVA TÉCNICA"
        .Range("I43").Value = "1111"
        .Range("M43").Value = "BRIGADA"
        .Range("e57").Value = vbNullString
        .Range("i35").Value = vbNullString
        '.Range("frmNovoExtintorSerie").Value = Info.Range("k6").Value
        '.Range("I35").ClearContents
        '.Range("frmNovoExtintorSerie").ClearContents
        .Range("I37").Activate
        .Calculate
    End With
    Info.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub:

End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub frmAtualiza()
    ' Data......: 13/01/2021
    ' Descricao.: chama form atualização de extintores
    '---------------------------------------------------------------------------------------
Public Sub frmAtualiza()
    On Error GoTo TError


    Application.EnableEvents = False
    Application.ScreenUpdating = False
   
    SetOnkey (True)
    Info.Unprotect
     
    With Info
        .Activate
        
        escondeobjfrmatualiza


        .Range("C2:C30").EntireRow.Hidden = False
        .Range("C90:C129").EntireRow.Hidden = True
        .Range("C59:C89").EntireRow.Hidden = True
        .Range("C31:C58").EntireRow.Hidden = True
        
        exibeobjfrmatualiza

        .Shapes("btnExtAdd").Top = 67.2
        
        .Range("I8").Select
        '        .Shapes("btnImprime").Height = 28.07
        

        dimBotoesFormAtualiza
        

        Info.Shapes("btnSalvaAtualExt").Top = Info.Range("P26").Top + Info.Range("OBS").Height + 5

        .Shapes("Scroll Bar 26").Left = 650.4 + .Shapes.Range(Array("btnocultarmenu")).Left
        .Shapes("Scroll Bar 48").Left = .Shapes("Scroll Bar 26").Left
        .Shapes("Scroll Bar 26").Height = Info.Range("tbHistMov").Height
        dimbarra
        .PageSetup.PrintArea = vbNullString
        .PageSetup.PrintArea = "$F$3:$W$28"
        populafrmAtualExt
        dimbtnsalvaext
    End With
    Info.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

Public Sub cancelaNovoLocalNovo()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    On Error GoTo ERRO

    With Info
        .Activate
        frmNovo

        .Range("frmCadastroLocal").Activate

    End With

ERRO:
    Exit Sub
    Info.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub cancelaNovoLocalAtualiza()
    ' Data......: 07/12/2020
    ' Descricao.: Retorna ao formulário de atualização de extintor a partir do formulário de
    '             cadastro de novo local
    '---------------------------------------------------------------------------------------
Public Sub cancelaNovoLocalAtualiza()
    On Error GoTo TError
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
   

    With Info

        .Activate
        .Range("I103:K103,I105:K105,N103").ClearContents
        frmAtualiza
        .Range("M12:N12").Select

    End With

    Info.Protect
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
' Contato...: warleywsc@gmail.com - Rotina: Public Sub frmLocalAtualiza()
    ' Data......: 07/12/2020
    ' Descricao.: Exibe o formulário de atulaização de extintor
    '---------------------------------------------------------------------------------------
Public Sub frmLocalAtualiza()
    On Error GoTo TError

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    With Info
        .Activate
       
        escondeobjfrmatualiza

        .Range("C90:C129").EntireRow.Hidden = False
        .Range("C2:C30").EntireRow.Hidden = True
        .Range("C31:C58").EntireRow.Hidden = True
        .Range("C59:C89").EntireRow.Hidden = True
        

        dimbtnsalvalocalatual
        dimbtnCancelarLocalAtual
        .Range("frmNovoLocal").Activate

    End With
    Info.Protect
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
' Contato...: warleywsc@gmail.com - Rotina: Public Sub frmLocalNovo()
    ' Data......: 07/12/2020
    ' Descricao.: Exibe o formulário de cadastro de novo extintor
    '---------------------------------------------------------------------------------------
Public Sub frmLocalNovo()
    On Error GoTo TError

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    With Info
        .Activate
       
        escondeobjfrmatualiza
        .Range("C59:C89").EntireRow.Hidden = False
        .Range("C2:C30").EntireRow.Hidden = True
        .Range("C31:C58").EntireRow.Hidden = True

        .Range("C90:C129").EntireRow.Hidden = True
        .Shapes("btnLocalAdd2").Visible = msoFalse
        dimbtnCancelarLocalNovo
        dimbtnSalvaLocalNovo
        .Range("I67").Activate

    End With
    Info.Protect
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
' Contato...: warleywsc@gmail.com - Rotina: Public Sub populafrmAtualExt()
    ' Data......: 28/11/2020
    ' Descricao.: Popula os campos do Formulário Atualizar extintor
    '---------------------------------------------------------------------------------------
Public Sub populafrmAtualExt()
    On Error GoTo TError

    Info.Unprotect
    On Error Resume Next
    Dim serieB As Variant
    Dim serieA As Variant
    Dim cell  As Range
    Dim lin   As Long
    For Each cell In Info.Range("I8,I10,M8,M10,I12,M12,I14,M14,I16,M16,I18,M18,I20,M20")
        'Retorna a cor da célula ao normal
        With cell.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
               
    Next cell
    With Info
        .Activate
        serieA = UCase$(.Range("frmCadastroSerie").Value)
    
    

        lin = 9
        Info.Range("$M$8,$I$10,$I$12,$M$10,$I$14,$M$14:$N$14,$I$16,$M$16,$I$18:$J$18,$I$20:$J$20,$M$18,$M$20").ClearContents
        Do Until Extintores.Cells(lin, 15) = vbNullString
            If Extintores.Cells(lin, 15) = serieA Then
                
                .Range("K6").Value = Extintores.Cells(lin, 7).Value 'TIPO
                .Range("M8").Value = Extintores.Cells(lin, 8).Value 'TIPO
                .Range("M10").Value = Extintores.Cells(lin, 9).Value 'CAP
                .Range("i10").Value = Extintores.Cells(lin, 10).Value 'CAP
                .Range("I12").Value = Extintores.Cells(lin, 11).Value 'SUP
               
            End If
            lin = lin + 1
        Loop
        lin = 9
        Do Until MapaAtual.Cells(lin, 14) = vbNullString
            If UCase$(MapaAtual.Cells(lin, 14)) = serieA Then

                .Range("M12").Value = UCase$(MapaAtual.Cells(lin, 10).Value) 'Local
               
                .Range("I14").Value = MapaAtual.Cells(lin, 8).Value 'Área
                .Range("M14").Value = UCase$(MapaAtual.Cells(lin, 15).Value) 'Zona
                If MapaAtual.Cells(lin, 16).Value <> vbNullString Then
                    .Range("I16").Value = DateAdd("yyyy", -5, MapaAtual.Cells(lin, 16).Value) 'TESTE

                Else
                    .Range("I16").Value = vbNullString
                End If
                If MapaAtual.Cells(lin, 18).Value <> vbNullString Then
                    If .Range("M8").Value = "CO" Or .Range("M8").Value = "FM" Then
                        .Range("M16").ClearContents
                        .Range("M16").Value = DateAdd("yyyy", -5, MapaAtual.Cells(lin, 18).Value) 'RECARGA CO, FM
                   
                    Else
                        .Range("M16").ClearContents
                        .Range("M16").Value = DateAdd("yyyy", -1, MapaAtual.Cells(lin, 18).Value) 'RECARGA OUTRO
                    End If
                Else
                    .Range("M16").ClearContents
                    .Range("M16").Value = vbNullString
                End If


                If MapaAtual.Cells(lin, 20).Value <> vbNullString Then
                    .Range("I18").Value = DateAdd("m", -6, MapaAtual.Cells(lin, 20).Value) 'PESAGEM
                Else
                    .Range("I18").Value = vbNullString
                End If

                If MapaAtual.Cells(lin, 22).Value <> vbNullString Then
                    .Range("M18").Value = DateAdd("yyyy", -1, MapaAtual.Cells(lin, 22).Value) 'SELAGEM
                Else
                    .Range("M18").Value = vbNullString
                End If

                If MapaAtual.Cells(lin, 24).Value <> vbNullString Then
                    If .Range("M8").Value = "CO" Then
                        .Range("$I$20:$J$20").ClearContents
                        .Range("$I$20:$J$20").Value = DateAdd("m", -6, MapaAtual.Cells(lin, 24).Value) 'INSPEÇÃO CO
                    ElseIf .Range("M8").Value = "FM" Then
                        .Range("I20").ClearContents
                        .Range("I20").Value = DateAdd("m", -1, MapaAtual.Cells(lin, 24).Value) 'INSPEÇÃO FM
                    Else
                        .Range("$I$20:$J$20").ClearContents
                        .Range("I20").Value = DateAdd("yyyy", -1, MapaAtual.Cells(lin, 24).Value) 'INSPEÇÃO OUTROS
                
                   
                    End If
                Else
                    .Range("$I$20:$J$20").ClearContents
                    .Range("I20").Value = vbNullString 'INSPEÇÃO OUTRO
                End If
              

                If MapaAtual.Cells(lin, 26).Value <> vbNullString Then
                    .Range("M20").Value = DateAdd("yyyy", -5, MapaAtual.Cells(lin, 26).Value) 'PINTURA
                Else
                    .Range("M20").Value = .Range("i16").Value
                End If

          
                'GoTo sair:
            End If

            lin = lin + 1

        Loop
        'sair:
        PopulaInfoOBS
        Range("A6").Value = Range("M12:N12").Value 'ARMAZENA LOCAL INICIAL (ANTES DE QUALQUER MUDANÇA)
        .Range("M12:n12").RowHeight = .Range("A6").RowHeight

        Range("A7").Value = Range("I14").Value 'ARMAZENA ÁREA INICIAL (ANTES DE QUALQUER MUDANÇA)
        Range("A8").Value = Range("M14:N14").Value 'ARMAZENA ZONA INICIAL (ANTES DE QUALQUER MUDANÇA)
        Range("A9").Value = Range("M8").Value 'ARMAZENA TIPO INICIAL (ANTES DE QUALQUER MUDANÇA)
        Range("A10").Value = Range("I10").Value 'ARMAZENA FABRICACAO INICIAL (ANTES DE QUALQUER MUDANÇA)
        Range("A11").Value = Range("M10").Value 'ARMAZENA CAPACIDADE INICIAL (ANTES DE QUALQUER MUDANÇA)
        Range("A14").Value = Range("I16").Value 'teste
        Range("A16").Value = Range("M16").Value
        Range("A18").Value = Range("I18").Value
        Range("A20").Value = Range("M18").Value
        Range("A22").Value = Range("I20").Value
        Range("A24").Value = Range("M20").Value
        .Calculate
        If .Cells(8, 13) = "CO" Or .Cells(8, 13) = "FM" Then
            .Range("I18:J18").Locked = False
        Else
            .Range("I18:J18").Locked = True

        End If
      

        dimBotoesFormAtualiza
        dimbtnsalvaext
        
    End With
 
    Info.Protect

fim:
    Exit Sub
TError:
    MsgBox Err.Number & Err.Description
    GoTo fim
End Sub

