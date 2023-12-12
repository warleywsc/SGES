VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlteraLocal 
   Caption         =   "Edição de locais e áreas"
   ClientHeight    =   4440
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6960
   OleObjectBlob   =   "codigofrmAlteraLocal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAlteraLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SGES2020")
Option Explicit

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btnFechar_Click()
Unload Me
End Sub

Private Sub btSalvar_Click()
    modificaLocalArea
End Sub









Private Sub txtLAntigo_Change()

End Sub

Private Sub UserForm_Initialize()
    Dim iControl As MSForms.control

    CriarControles
   
    CriarBotaoAcao
    For Each iControl In frmAlteraLocal.Controls

   

        If TypeOf iControl Is MSForms.Label Then
            With iControl
                .ForeColor = vbWhite
                .BackColor = 5855577
            End With

        End If
     
    Next
    With Me
    
        .BackColor = 15921906
        .frmMenuLateral.BackColor = 5855577
        .frmMenuLateral.Left = 0
        .frmMenuLateral.Top = 0
        .frmMenuLateral.Height = Me.Height
        .FrameCorpo.BackColor = 15921906 '11711154
        .FrameCorpo.BorderColor = 15921906
        .txtLAntigo.SetFocus
        .txtLAntigo.BackColor = 13564414
    
    
    End With

  Dim localatual As String
    Dim areaatual As String
    Dim zonaatual As String
'    Dim zonanova As String
'    Dim areanova As String
'    Dim localnovo As String
    
    
    
    localatual = Info.Range("$M$12").Value
    areaatual = Info.Range("$I$14").Value
    zonaatual = Info.Range("$M$14").Value
'    zonanova = zonaatual
'    areanova = areaatual
'    localnovo = localatual
    
     frmAlteraLocal.txtLAntigo.Value = localatual
    frmAlteraLocal.txtAAtual.Value = areaatual
    frmAlteraLocal.txtZonaAtual.Value = zonaatual
      frmAlteraLocal.txtLNovo.Value = localatual
    frmAlteraLocal.txtAnova.Value = areaatual
    frmAlteraLocal.txtZonaNova.Value = zonaatual
frmAlteraLocal.txtLNovo.SetFocus
End Sub
