VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Public WithEvents btMenu As MSForms.Label
Attribute btMenu.VB_VarHelpID = -1
Public Index  As Long
Public Focado As Boolean

Public WithEvents btMenuMov As MSForms.Label
Attribute btMenuMov.VB_VarHelpID = -1
Public IndexMov As Long
Public FocadoMov As Boolean

Public Sub PropriedadePadrao()
    Dim espacamento As Long, altura As Single
    espacamento = 3
    altura = Me.btMenu.Height
    With Me.btMenu
        .BackColor = FUNDO_CINZA_ESCURO
        .ForeColor = FONTE_BRANCA
        .SpecialEffect = fmSpecialEffectFlat
        '    .BorderColor = FUNDO_CINZA_MEDIO
        .TextAlign = fmTextAlignCenter
        .Left = ESP_HORIZONTAL
        .Top = espacamento * Index + altura * (Index - 1)
    End With

    Me.Focado = False

End Sub

Public Sub PropriedadeCustom()
    With Me.btMenu
        .SpecialEffect = fmSpecialEffectBump

    End With

    Me.Focado = True


End Sub

Private Sub btMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim iMenu As clsMenu

    If Me.Focado = True Then Exit Sub

    For Each iMenu In ColMenu

        iMenu.PropriedadePadrao

    Next

    Me.PropriedadeCustom


End Sub







'--------------------------------------------------------





Public Sub PropriedadePadraoMov()
    Dim espacamento As Long, altura As Single
    espacamento = 3
    altura = Me.btMenuMov.Height
    With Me.btMenuMov
        .BackColor = FUNDO_CINZA_ESCURO_MOV
        .ForeColor = FONTE_BRANCA
        .SpecialEffect = fmSpecialEffectFlat
        '    .BorderColor = FUNDO_CINZA_MEDIO
        .TextAlign = fmTextAlignCenter
        .Left = ESP_HORIZONTAL
        .Top = espacamento * IndexMov + altura * (IndexMov - 1)
    End With

    Me.FocadoMov = False

End Sub

Public Sub PropriedadeCustomMov()
    With Me.btMenuMov
        .SpecialEffect = fmSpecialEffectBump

    End With

    Me.FocadoMov = True


End Sub

Private Sub btMenuMov_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim iMenu As clsMenu

    If Me.FocadoMov = True Then Exit Sub

    For Each iMenu In ColMenu

        iMenu.PropriedadePadraoMov

    Next

    Me.PropriedadeCustomMov


End Sub

