Attribute VB_Name = "TabMacros"
Option Explicit
'@Folder("SGES2020_dev")

Public Sub SetOnkey(ByVal state As Boolean)
    'Compiled By Randy Austin
    'Workbook Provided By www.ExcelForFreelancers.com
    If state Then
        With Application
        On Error Resume Next
            .OnKey "{TAB}", "'TabRange xlNext'"
            .OnKey "~", "'TabRange xlNext'"
            .OnKey "{ENTER}", "'TabRange xlNext'"
            .OnKey "{RIGHT}", "'TabRange xlNext'"
            .OnKey "{LEFT}", "'TabRange xlPrevious'"
            .OnKey "+{TAB}", "'TabRange xlPrevious'"
            .OnKey "{DOWN}", "'UpOrDownArrow xlDown'"
            .OnKey "{UP}", "'UpOrDownArrow xlUp'"
        End With
    Else
        'reset keys
        With Application
            .OnKey "{ENTER}"
            .OnKey "{TAB}"
            .OnKey "~"
            .OnKey "+{TAB}"
            .OnKey "{RIGHT}"
            .OnKey "{LEFT}"
            .OnKey "{DOWN}"
            .OnKey "{UP}"
        End With
    End If
End Sub

Public Function GetTabOrder() As Variant
    '--set the tab order of input cells - change ranges as required
    '  don't include "$" in these cell references
    If ActiveSheet.Name = "Info" Then
       
        If ActiveCell.Row > 1 And ActiveCell.Row < 31 Then 'frm atualizar
            GetTabOrder = Array("I8", "M8", "I10", "M10", "I12", "M12", "I14", "M14", "I16", "M16", "I18", "M18", "I20", "M20", "G23", "M23")
        End If
    
        If ActiveCell.Row > 30 And ActiveCell.Row < 58 Then 'frm novo
            GetTabOrder = Array("I37", "M37", "I39", "M39", "I41", "M41", "I43", "M43", "I45", "M45", "I47", "M47", "I49", "M49", "G52")
    
        End If
    
        If ActiveCell.Row > 58 And ActiveCell.Row < 102 Then 'frm cadastro local/frm atualizar
            GetTabOrder = Array("I67", "N67", "I69")
    
        End If
        If ActiveCell.Row > 89 And ActiveCell.Row < 130 Then 'frm cadastro local/frm atualizar
            GetTabOrder = Array("I103", "N103", "I105")
    
        End If
     
    End If
    
    
    If ActiveSheet.Name = "Pesquisa" Then
   
        If ActiveCell.Row >= 2 And ActiveCell.Row < 5000 Then 'frm atualizar
            GetTabOrder = Array("I2", "I3", "I4", "I5", "I6", "I7", "K3", "K4", "K5", "K6", "M3", "N3", "O3", "M4", "N4", "O4", "M5", "N5", "O5", "M6", "N6", "O6", "M7", "N7", "O7")
        End If
        
    End If
  
End Function

Public Sub TabRange(Optional ByRef iDirection As Long = xlNext)
    
    Dim vTabOrder As Variant
    Dim m     As Variant

    Dim lItems As Long
    Dim iAdjust As Long

    On Error GoTo ExitSub
    '--get the tab order from shared function
    vTabOrder = GetTabOrder
    lItems = UBound(vTabOrder) - LBound(vTabOrder) + 1

    On Error GoTo ErrorHandler
    m = Application.Match(ActiveCell.Address(0, 0), vTabOrder, False)
    

    '--if activecell is not in Tab Order return to the first cell
    If IsError(m) Then
        m = 1
    Else
        '--get adjustment to index
        iAdjust = IIf(iDirection = xlPrevious, -1, 1)

        '--calculate new index wrapping around list
        m = (m + lItems + iAdjust - 1) Mod lItems + 1
    End If

    '--select cell adjusting for Option Base 0 or 1
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Range(vTabOrder(m + (LBound(vTabOrder) = 0))).Select
   
    Selection.Calculate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
ExitSub:


    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then             'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub UpOrDownArrow(Optional ByRef iDirection As Long = xlUp)

    Dim vTabOrder As Variant
    Dim lRowClosest As Long
    Dim lRowTest As Long

    Dim i     As Long
    Dim iSign As Long


    Dim sActiveCol As String
    Dim bFound As Boolean

    '--get the tab order from shared function
    vTabOrder = GetTabOrder

    '--find TabCells in same column as ActiveCell in iDirection
    '--  rTest will include ActiveCell

    sActiveCol = GetColLtr(ActiveCell.Address(0, 0))

    iSign = IIf(iDirection = xlDown, -1, 1)
    lRowClosest = IIf(iDirection = xlDown, ActiveSheet.Rows.Count + 1, 0)

    For i = LBound(vTabOrder) To UBound(vTabOrder)
        If GetColLtr(CStr(vTabOrder(i))) = sActiveCol Then
            lRowTest = Range(CStr(vTabOrder(i))).Row

            '--find closest cell to ActiveCell in rTest
            If iSign * lRowTest > iSign * lRowClosest And _
               iSign * lRowTest < iSign * ActiveCell.Row Then
                '--at least one cell in iDirection of same columnn
                bFound = True
                lRowClosest = lRowTest
            End If
        End If
    Next i

    If bFound Then
     
        ActiveSheet.Cells.Item(lRowClosest, ActiveCell.Column).Select
       
    End If
End Sub

Private Function GetColLtr(ByRef sAddr As String) As String
    Dim iPos  As Long
    Dim sTest As String

    Do While iPos < 3
        iPos = iPos + 1
        If IsNumeric(Mid$(sAddr, iPos, 1)) Then
            Exit Do
        Else
            sTest = sTest & Mid$(sAddr, iPos, 1)
        End If
    Loop
    GetColLtr = sTest
End Function









