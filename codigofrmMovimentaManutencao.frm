VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMovimentaManutencao 
   Caption         =   "Controle de Remessas"
   ClientHeight    =   7305
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12465
   OleObjectBlob   =   "codigofrmMovimentaManutencao.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMovimentaManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btEnvioInserir_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If KeyCode = 13 Then

        btEnvioInserir_Click

    
    End If
    cbEnvioSerie.SetFocus
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub btEnvioExcluir_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    deletaEnvios
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub btEnvioInserir_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If cbEnvioSerie.Value <> vbNullString Then
        PermiteInserirListaEnvio

    Else
        MsgBox "Insira o número de série", vbCritical, "Atenção"
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub btMenuEnvio_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Rolagem FrameEnvio
    CriarControlesFormMovimentacao
    '    Call btPesquisaLimpar_Click
    With Me
        .cbEnvioSerie.SetFocus
        '        .cbEnvioSerie.BackColor = COR_CXTEXTO
        '        .cbEnvioSerie.ForeColor = FONTE_BRANCA
        .btEnvioInserir.Enabled = True
        '        .btVendasEditar.Enabled = False
        '        .btVendasExcluir.Enabled = False
        .btEnvioExcluir.Enabled = True
    End With
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub btMenuRetorno_Click()
    '    Application.EnableEvents = False
    '    Application.ScreenUpdating = False
    '    Rolagem FrameRetorno
    '    populalbxRetorno
    '    CriarControlesFormMovimentacao
    '    '    Call btPesquisaLimpar_Click
    '    With Me
    '        .cbRetornoSerie.SetFocus
    '        '        .cbRetornoSerie.BackColor = COR_CXTEXTO
    '        '        .cbRetornoSerie.ForeColor = FONTE_BRANCA
    '        .btRetornoInserir.Enabled = True
    '        '        .btVendasEditar.Enabled = False
    '        '        .btVendasExcluir.Enabled = False
    '        .btRetornoExcluir.Enabled = True
    '    End With
    '    Application.EnableEvents = True
    '    Application.ScreenUpdating = True
End Sub


Private Sub btMenuSair_Click()

    Unload Me
    End
   
End Sub



Private Sub btnRemoverRetorno_Click()
    RemoverdalistaRetorno
End Sub

Private Sub btRetornoExcluir_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    deletaRetornos
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub btRetornoInserir_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If cbRetornoSerie.Value <> vbNullString Then
        PermiteInserirListaRetorno

        
        
  
    Else
        MsgBox "Insira o número de série", vbCritical, "Atenção"
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub btRetornoInserir_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If KeyCode = 13 Then

        btRetornoInserir_Click

    
    End If
    cbRetornoSerie.SetFocus
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub cbEnvioSerie_Change()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    PopulalbEnvioLocalAtual
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



Private Sub btnRemoverEnvio_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    RemoverdalistaEnviar
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub cbRetornoSerie_Change()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    PopulalbRetornoLocalAtual
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub lbRetornoImprimir_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    
    MovRetornoEmBloco
    ImprimeRetorno
    If frmMovimentaManutencao.Visible = False Then frmMovimentaManutencao.Show
    
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub lbEnviarImprimir_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    
    MovEnvioEmBloco
    ImprimeEnvio
    If frmMovimentaManutencao.Visible = False Then frmMovimentaManutencao.Show
    
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Private Sub UserForm_Initialize()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    CriarControlesFormMovimentacao
    CriarMenuFormMovimentacao
    ConfiguracaoLayoutFormMovimentacao
    CriarBotaoAcaoFormMovimentacao
    Rolagem FrameEnvio
    populalboxenvio
    
    autodimlistbox
    btMenuEnvio_Click
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Private Sub IfrmMovimentaManutencao_Rolagem(FrameDestino As MSForms.Frame)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Rolagem FrameDestino
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub Rolagem(FrameDestino As MSForms.Frame)

    Dim PosicaoReferencia As Single
    Dim PosicaoFinal As Single
    Dim PosicaoBarra As Single
    Dim i     As Single
    Dim f     As Long
    '    Application.EnableEvents = False
    '    Application.ScreenUpdating = False
    PosicaoReferencia = frmMovimentaManutencao.FrameEnvio.Top
    PosicaoFinal = FrameDestino.Top
    
    PosicaoBarra = PosicaoFinal - PosicaoReferencia
    If frmMovimentaManutencao.FrameCorpo.ScrollTop > PosicaoBarra Then f = -1 Else f = 1
    
    For i = frmMovimentaManutencao.FrameCorpo.ScrollTop To PosicaoBarra Step 8 * f
        DoEvents
        frmMovimentaManutencao.FrameCorpo.ScrollTop = i
        
    Next i
    frmMovimentaManutencao.FrameCorpo.ScrollTop = PosicaoBarra
    
    '    Application.EnableEvents = True
    '    Application.ScreenUpdating = True
End Sub



Private Sub UserForm_Terminate()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Me.ListBoxenvio.RowSource = vbNullString
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'
'Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'
'    On Error GoTo ErrIt
'
'    Me.Controls(ListBox1.Value).Value = True
'    Exit Sub
'
'ErrIt:
'    CallByName Me, ListBox1.Value & "_Click", VbMethod
'    Exit Sub
'
'End Sub
