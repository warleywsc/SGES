Attribute VB_Name = "MODEscondeTitulo"


Option Explicit
'@Folder("SGES2020")


Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
ByVal hwnd As LongPtr, _
ByVal crKey As Long, _
ByVal bAlpha As Long, _
ByVal dwFlags As LongPtr) As Long

Private Declare PtrSafe Function SetWindowText _
Lib "user32" Alias "SetWindowTextA" _
(ByVal hwnd As Long, _
ByVal lpString As String) As Long
'Constants for title bar
Private Const GWL_STYLE As Long = (-16) 'The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20) 'The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000 'Style to add a titlebar
Private Const WS_EX_DLGMODALFRAME As Long = &H1 'Controls if the window has an icon
 
'Constants for transparency
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1       'Chroma key for fading a certain color on your Form
Private Const LWA_ALPHA = &H2          'Only needed if you want to fade the entire userform
'Only needed if you want to fade the entire userform

Public Sub MakeUserFormTransparent(frm As Object, Optional Color As Variant)
    'set transparencies on userform
    Dim formhandle As Long
    Dim bytOpacity As Byte
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    formhandle = FindWindow(vbNullString, frm.Caption)
    If IsMissing(Color) Then Color = vbBlack 'default to vbwhite
    bytOpacity = 100                   ' variable keeping opacity setting
 
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    'The following line makes only a certain color transparent so the
    ' background of the form and any object whose BackColor you've set to match
    ' vbColor (default vbWhite) will be transparent.
    frm.BackColor = Color
    SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY
End Sub

Public Sub HideTitleBarAndBorder(frm As Object)
    'Hide title bar and border around userform
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
    'Build window and set window until you remove the caption, title bar and frame around the window
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
   
End Sub

Public Sub AlterExcelCaption(strCaption As String)
    SetWindowText FindWindow("Sistema de Gestão de Equipamentos e Serviços - SGES2020", Application.Caption), strCaption
End Sub

Public Sub Test()
    AlterExcelCaption "Sistema de Gestão de Equipamentos e Serviços - SGES2021"
End Sub




