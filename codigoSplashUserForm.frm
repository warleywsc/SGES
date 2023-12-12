VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashUserForm 
   Caption         =   "teste"
   ClientHeight    =   8640.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15345
   OleObjectBlob   =   "codigoSplashUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplashUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SGES2020")




Option Explicit

Private mblnExitMe As Boolean
Private mstrTitle As String
Private mstrSubTitle As String

Private Const mbytcSpeedSeconds As Byte = 4

Public Property Let FormTitle(ByVal strText As String)
    mstrTitle = strText
End Property

Public Property Let FormSubTitle(ByVal strText As String)
    mstrSubTitle = strText
End Property

Private Sub UserForm_QueryClose(ByRef Cancel As Integer, _
                                ByRef CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        'close properly
        mblnExitMe = True
    End If
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mblnExitMe = True
    End If
End Sub

Private Sub UserForm_Activate()
    Dim dtmWait As Date
    Dim intProgress As Integer

    On Error GoTo ErrHandler
    Application.EnableCancelKey = xlErrorHandler

    With Me
        'form position
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)

        'set up labels
        '        .lblTitle.Caption = mstrTitle
        '        With .lblSubTitle
        '            .Caption = mstrSubTitle
        '            .AutoSize = True
        '        End With
        .lbl_Data.ForeColor = RGB(255, 192, 0)
        .lbl_Data.Caption = Format$(Now, "dddd d mmmm yyyy hh:mm")
        
    End With

    'to kill splash after displaying for a short time
    'Application.OnTime EarliestTime:=Now + TimeValue("00:00:03"), Procedure:="EndSplash"


    'run progress bar
    '    For intProgress = 0 To 100 Step (100 / mbytcSpeedSeconds)
    '
    '        dtmWait = Now + TimeValue("0:00:01")
    '
    '        Do While Now < dtmWait
    '            DoEvents
    '            'user closed form
    '            If mblnExitMe = True Then
    '                Exit For
    '            End If
    '
    '            Me.ProgressBar1.Value = intProgress
    '        Loop
    '    Next intProgress

    HideTitleBarAndBorder Me           'Esconde barra de títulos e bordas
    Application.Wait (Now + TimeValue("00:00:03"))
    SplashUserForm.lblcarregando.ForeColor = RGB(255, 192, 0)
    SplashUserForm.lblcarregando.Caption = "Carregando Extintores..."
    SplashUserForm.lblcarregando.Font.Size = "11"
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:03"))
    SplashUserForm.lblcarregando.Caption = "Aduchando mangueiras..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:03"))
    SplashUserForm.lblcarregando.Caption = "Prevenindo incêndios..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:03"))
    
    MakeUserFormTransparent Me         'Torna algumas cores transparentes




ExitProcedure:
    Unload Me
    Exit Sub

ErrHandler:
    Debug.Print Now() & vbTab & Err.Number & vbTab & Err.Description
    Err.Clear
    mblnExitMe = True
    Resume Next
End Sub

'
'Private Sub UserForm_Activate()
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'
'    HideTitleBarAndBorder Me           'Esconde barra de títulos e bordas
'    Application.Wait (Now + TimeValue("00:00:04"))
'    SplashUserForm.lblcarregando.Caption = "Carregando Extintores..."
'    SplashUserForm.lblcarregando.Font.Size = "11"
'    SplashUserForm.Repaint
'    Application.Wait (Now + TimeValue("00:00:04"))
'    SplashUserForm.lblcarregando.Caption = "Aduchando mangueiras..."
'    SplashUserForm.Repaint
'    Application.Wait (Now + TimeValue("00:00:04"))
'    SplashUserForm.lblcarregando.Caption = "Prevenindo incêndios..."
'    SplashUserForm.Repaint
'    Application.Wait (Now + TimeValue("00:00:04"))
'
'    MakeUserFormTransparent Me         'Torna algumas cores transparentes
'    Unload SplashUserForm
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'End Sub




