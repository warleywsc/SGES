Attribute VB_Name = "Módulo6"
Option Explicit

Sub Macro8()
    '
    ' Macro8 Macro
    '

    '
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    Range("AO8:AW8").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Preview:=True, Collate:= _
                                         True, IgnorePrintAreas:=False
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$8:$8"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$AO$8:$AW$8"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "&G"
        .CenterHeader = _
                      "&""+,Negrito""&22&K44546ASistema de Gestão de Equipamentos e Serviços"
        .RightHeader = "&""-,Negrito""&12&K04-048Formulário de Envio" & Chr(10) & "" & Chr(10) & " Número: 1"
        .LeftFooter = "&D & - &T"
        .CenterFooter = _
                      "&""-,Negrito""&12&K04-049Carimbo/Assinatura" & Chr(10) & "" & Chr(10) & "___________________"
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
        .FitToPagesTall = False
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

