VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private WithEvents pTextBox As MSForms.TextBox
Attribute pTextBox.VB_VarHelpID = -1
Private WithEvents pComboBox As MSForms.ComboBox
Attribute pComboBox.VB_VarHelpID = -1
Public pControl As MSForms.control
Public Focado As Boolean

Public Property Set campo(iControl As MSForms.control)

    Set pControl = iControl

    If TypeOf iControl Is MSForms.TextBox Then
    
        Set pTextBox = iControl
    
    ElseIf TypeOf iControl Is MSForms.ComboBox Then
    
        Set pComboBox = iControl
        
    End If
    
End Property

Private Sub iControlKeyUp(KeyCode As MSForms.ReturnInteger)
    
    Dim iControl As clsControles
    
    If KeyCode <> vbKeyTab And KeyCode <> vbKeyReturn _
       And KeyCode <> vbKeyDown And KeyCode <> vbKeyUp Then Exit Sub
    
        
    
    
    If Me.Focado = True Then Exit Sub
    
    For Each iControl In ColControle
        iControl.PropriedadePadrao
    Next iControl
    
    Me.PropriedadeCustom
    
    

End Sub

Public Sub icontrolMouseMove()
    Dim iControl As clsControles
    If Me.Focado = True Then Exit Sub
    
    For Each iControl In ColControle
        On Error Resume Next
        If iControl.pControl.Name <> Me.pControl.Parent.ActiveControl.Name Then
    
    
            iControl.PropriedadePadrao
        End If
    Next iControl
    Me.PropriedadeCustom
End Sub

Public Sub PropriedadePadrao()

    With pControl
        .ForeColor = COR_FONTE_SECUNDARIA
        .BackColor = FONTE_BRANCA
        .SpecialEffect = fmSpecialEffectEtched
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = vbBlack
        
    
    End With
    Me.Focado = False

End Sub

Public Sub PropriedadeCustom()
    With pControl
   
        .BackColor = 13564414          'amarelo claro
        .BorderColor = vbBlack
    End With
    Me.Focado = True

End Sub


Private Sub pComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    iControlKeyUp KeyCode
End Sub

Private Sub pComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    icontrolMouseMove
End Sub

Private Sub pTextBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    iControlKeyUp KeyCode
End Sub


Private Sub pTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    icontrolMouseMove
End Sub




