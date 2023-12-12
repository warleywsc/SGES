Attribute VB_Name = "MODCriacabecalho"

Option Explicit
'@Folder("SGES2020")

'Public Sub cabecalhorodape()
'
'    Range("F3:W28").Select
'    ActiveSheet.PageSetup.PrintArea = "$F$3:$W$28"
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .PrintTitleRows = vbNullString
'        .PrintTitleColumns = vbNullString
'    End With
'    Application.PrintCommunication = True
'    ActiveSheet.PageSetup.PrintArea = "$F$3:$W$28"
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .LeftHeader = vbNullString
'        .CenterHeader = vbNullString
'        .RightHeader = vbNullString
'        .LeftFooter = vbNullString
'        .CenterFooter = vbNullString
'        .RightFooter = vbNullString
'        .LeftMargin = Application.InchesToPoints(0.25)
'        .RightMargin = Application.InchesToPoints(0.25)
'        .TopMargin = Application.InchesToPoints(0.75)
'        .BottomMargin = Application.InchesToPoints(0.75)
'        .HeaderMargin = Application.InchesToPoints(0.3)
'        .FooterMargin = Application.InchesToPoints(0.3)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .CenterHorizontally = False
'        .CenterVertically = False
'        .Orientation = xlLandscape
'        .Draft = False
'        .PaperSize = xlPaperA4
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = False
'        .FitToPagesWide = 1
'        .FitToPagesTall = False
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = True
'        .EvenPage.LeftHeader.Text = vbNullString
'        .EvenPage.CenterHeader.Text = vbNullString
'        .EvenPage.RightHeader.Text = vbNullString
'        .EvenPage.LeftFooter.Text = vbNullString
'        .EvenPage.CenterFooter.Text = vbNullString
'        .EvenPage.RightFooter.Text = vbNullString
'        .FirstPage.LeftHeader.Text = vbNullString
'        .FirstPage.CenterHeader.Text = vbNullString
'        .FirstPage.RightHeader.Text = vbNullString
'        .FirstPage.LeftFooter.Text = vbNullString
'        .FirstPage.CenterFooter.Text = vbNullString
'        .FirstPage.RightFooter.Text = vbNullString
'    End With
'    Application.PrintCommunication = True
'    ActiveWindow.View = xlPageLayoutView
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .LeftHeader = vbNullString
'        .CenterHeader = vbNullString
'        .RightHeader = vbNullString
'        .LeftFooter = vbNullString
'        .CenterFooter = vbNullString
'        .RightFooter = vbNullString
'        .LeftMargin = Application.InchesToPoints(0.25)
'        .RightMargin = Application.InchesToPoints(0.25)
'        .TopMargin = Application.InchesToPoints(0.75)
'        .BottomMargin = Application.InchesToPoints(0.75)
'        .HeaderMargin = Application.InchesToPoints(0.3)
'        .FooterMargin = Application.InchesToPoints(0.3)
'        .Zoom = 63
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = True
'        .EvenPage.LeftHeader.Text = vbNullString
'        .EvenPage.CenterHeader.Text = vbNullString
'        .EvenPage.RightHeader.Text = vbNullString
'        .EvenPage.LeftFooter.Text = vbNullString
'        .EvenPage.CenterFooter.Text = vbNullString
'        .EvenPage.RightFooter.Text = vbNullString
'        .FirstPage.LeftHeader.Text = vbNullString
'        .FirstPage.CenterHeader.Text = vbNullString
'        .FirstPage.RightHeader.Text = vbNullString
'        .FirstPage.LeftFooter.Text = vbNullString
'        .FirstPage.CenterFooter.Text = vbNullString
'        .FirstPage.RightFooter.Text = vbNullString
'    End With
'    Application.PrintCommunication = True
'    ActiveSheet.PageSetup.LeftHeaderPicture.Filename = _
'                                                     "C:\Users\warle\Documents\BRIGADA\ícones\logotipo.png"
'    With ActiveSheet.PageSetup.LeftHeaderPicture
'        .Height = 31.8
'        .Width = 82.8
'    End With
'    Selection.Font.Bold = True
'    With Selection.Font
'        .ThemeColor = xlThemeColorAccent1
'        .TintAndShade = -0.499984740745262
'    End With
'    Selection.Font.Bold = True
'    With Selection.Font
'        .ThemeColor = xlThemeColorAccent1
'        .TintAndShade = -0.499984740745262
'    End With
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .LeftHeader = "&G"
'        .CenterHeader = _
'                      "&""+,Negrito""&13&K04-049" & Chr$(10) & "Sistema de Gestão de Equipamentos e Serviços"
'        .RightHeader = vbNullString & Chr$(10) & "&""-,Negrito""&12&K04-049&A"
'        .LeftFooter = vbNullString
'        .CenterFooter = vbNullString
'        .RightFooter = vbNullString
'        .LeftMargin = Application.InchesToPoints(0.25)
'        .RightMargin = Application.InchesToPoints(0.25)
'        .TopMargin = Application.InchesToPoints(0.75)
'        .BottomMargin = Application.InchesToPoints(0.75)
'        .HeaderMargin = Application.InchesToPoints(0.3)
'        .FooterMargin = Application.InchesToPoints(0.3)
'        .Zoom = 63
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = True
'        .EvenPage.LeftHeader.Text = vbNullString
'        .EvenPage.CenterHeader.Text = vbNullString
'        .EvenPage.RightHeader.Text = vbNullString
'        .EvenPage.LeftFooter.Text = vbNullString
'        .EvenPage.CenterFooter.Text = vbNullString
'        .EvenPage.RightFooter.Text = vbNullString
'        .FirstPage.LeftHeader.Text = vbNullString
'        .FirstPage.CenterHeader.Text = vbNullString
'        .FirstPage.RightHeader.Text = vbNullString
'        .FirstPage.LeftFooter.Text = vbNullString
'        .FirstPage.CenterFooter.Text = vbNullString
'        .FirstPage.RightFooter.Text = vbNullString
'    End With
'    Application.PrintCommunication = True
'    ActiveWindow.ScrollRow = 1
'    ActiveWindow.SmallScroll Down:=15
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .LeftHeader = "&G"
'        .CenterHeader = _
'                      "&""+,Negrito""&13&K04-048" & Chr$(10) & "Sistema de Gestão de Equipamentos e Serviços"
'        .RightHeader = vbNullString & Chr$(10) & "&""-,Negrito""&12&K04-048&A"
'        .LeftFooter = vbNullString
'        .CenterFooter = vbNullString
'        .RightFooter = vbNullString
'        .LeftMargin = Application.InchesToPoints(0.25)
'        .RightMargin = Application.InchesToPoints(0.25)
'        .TopMargin = Application.InchesToPoints(0.75)
'        .BottomMargin = Application.InchesToPoints(0.75)
'        .HeaderMargin = Application.InchesToPoints(0.3)
'        .FooterMargin = Application.InchesToPoints(0.3)
'        .Zoom = 63
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = True
'        .EvenPage.LeftHeader.Text = vbNullString
'        .EvenPage.CenterHeader.Text = vbNullString
'        .EvenPage.RightHeader.Text = vbNullString
'        .EvenPage.LeftFooter.Text = vbNullString
'        .EvenPage.CenterFooter.Text = vbNullString
'        .EvenPage.RightFooter.Text = vbNullString
'        .FirstPage.LeftHeader.Text = vbNullString
'        .FirstPage.CenterHeader.Text = vbNullString
'        .FirstPage.RightHeader.Text = vbNullString
'        .FirstPage.LeftFooter.Text = vbNullString
'        .FirstPage.CenterFooter.Text = vbNullString
'        .FirstPage.RightFooter.Text = vbNullString
'    End With
'    Application.PrintCommunication = True
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .LeftHeader = "&G"
'        .CenterHeader = _
'                      "&""+,Negrito""&13&K04-047" & Chr$(10) & "Sistema de Gestão de Equipamentos e Serviços"
'        .RightHeader = vbNullString & Chr$(10) & "&""-,Negrito""&12&K04-047&A"
'        .LeftFooter = "&D - &T" & Chr$(10) & vbNullString
'        .CenterFooter = vbNullString
'        .RightFooter = vbNullString & Chr$(10) & "&P"
'        .LeftMargin = Application.InchesToPoints(0.25)
'        .RightMargin = Application.InchesToPoints(0.25)
'        .TopMargin = Application.InchesToPoints(0.75)
'        .BottomMargin = Application.InchesToPoints(0.75)
'        .HeaderMargin = Application.InchesToPoints(0.3)
'        .FooterMargin = Application.InchesToPoints(0.3)
'        .Zoom = 63
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = True
'        .EvenPage.LeftHeader.Text = vbNullString
'        .EvenPage.CenterHeader.Text = vbNullString
'        .EvenPage.RightHeader.Text = vbNullString
'        .EvenPage.LeftFooter.Text = vbNullString
'        .EvenPage.CenterFooter.Text = vbNullString
'        .EvenPage.RightFooter.Text = vbNullString
'        .FirstPage.LeftHeader.Text = vbNullString
'        .FirstPage.CenterHeader.Text = vbNullString
'        .FirstPage.RightHeader.Text = vbNullString
'        .FirstPage.LeftFooter.Text = vbNullString
'        .FirstPage.CenterFooter.Text = vbNullString
'        .FirstPage.RightFooter.Text = vbNullString
'    End With
'    Application.PrintCommunication = True
'    ActiveWindow.LargeScroll Down:=-1
'    Range("G7").Select
'    ActiveWindow.View = xlNormalView
'    ActiveWindow.SmallScroll Down:=-18
'End Sub









