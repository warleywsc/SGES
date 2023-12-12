Attribute VB_Name = "MODRemovebarraformula"
Option Explicit

Public Sub EXIBEBARRAFORMULA()
    '
    ' EXIBEBARRAFORMULA Macro
    '

    '
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
End Sub
