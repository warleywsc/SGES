VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBotaoAcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Public WithEvents CommandButton As MSForms.CommandButton
Attribute CommandButton.VB_VarHelpID = -1

Public Sub PropriedadePadraoBotao()

    With CommandButton
        
        .BackColor = FUNDO_CINZA_MEDIO
        .ForeColor = FONTE_BRANCA
        
    End With

End Sub

Public Sub PropriedadeCustom()

    With CommandButton
        .BackColor = vbBlack
        .ForeColor = FONTE_BRANCA
  
    End With
 
End Sub

Private Sub CommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim iBotaoAcao As clsBotaoAcao
    For Each iBotaoAcao In ColBotaoAcao
        iBotaoAcao.PropriedadePadraoBotao
    Next
    Me.PropriedadeCustom
End Sub

