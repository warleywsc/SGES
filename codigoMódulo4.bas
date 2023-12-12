Attribute VB_Name = "Módulo4"
Option Explicit

Sub Macro5()
    '
    ' Macro5 Macro
    '

    '
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "&G"
        .CenterHeader = _
                      "&""+,Negrito""&24&K44546ASistema de Gestão de Equipamentos e Serviços"
        .RightHeader = "&""-,Negrito""&12&K04-049Formulário de Envio" & Chr(10) & "" & Chr(10) & " Número: 3"
        .LeftFooter = "&D & - &T"
        .CenterFooter = ""
        .RightFooter = "&P"
        .LeftMargin = Application.InchesToPoints(0.511811023622047)
        .RightMargin = Application.InchesToPoints(0.511811023622047)
        .TopMargin = Application.InchesToPoints(0.866141732283465)
        .BottomMargin = Application.InchesToPoints(0.748031496062992)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 60
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintPreview
End Sub
Sub Macro6()
    '
    ' Macro6 Macro
    '

    '
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
End Sub
Sub Macro7()
    '
    ' Macro7 Macro
    '

    '
    ActiveSheet.PageSetup.PrintArea = "$G$8:$O$21"
    ActiveSheet.ResetAllPageBreaks
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$8:$8"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$G$8:$O$21"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "&G"
        .CenterHeader = _
                      "&""+,Negrito""&24&K44546ASistema de Gestão de Equipamentos e Serviços"
        .RightHeader = "&""-,Negrito""&12&K04-049Formulário de Envio" & Chr(10) & "" & Chr(10) & " Número: 1"
        .LeftFooter = "&D & - &T"
        .CenterFooter = ""
        .RightFooter = "&P"
        .LeftMargin = Application.InchesToPoints(0.511811023622047)
        .RightMargin = Application.InchesToPoints(0.511811023622047)
        .TopMargin = Application.InchesToPoints(0.866141732283465)
        .BottomMargin = Application.InchesToPoints(0.748031496062992)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
End Sub

