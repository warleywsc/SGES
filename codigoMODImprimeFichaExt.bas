Attribute VB_Name = "MODImprimeFichaExt"
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley
' Contato...: warleywsc@gmail.com - Rotina: Sub FichaExt()
    ' Data......: 16/11/2020
    ' Descricao.: Gera impressão de info do extintor
    '---------------------------------------------------------------------------------------
    #If VBA7 Then
        Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    #Else
        Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #End If





Public Sub FichaExt()
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim clm   As Range
    Dim i     As Integer
    Dim j     As Variant
    Dim k     As Variant
    Dim cli   As Range
    Dim cls   As Range
    Dim ultlinha As Range
    Dim LINHA As Long
    Application.GoTo ImprimeFichaExt.Range("b15")
    ImprimeFichaExt.Range("b15:H5000").Clear

    With ImprimeFichaExt

        .Range("c4").Value = Info.Range("k6").Value
        .Range("c6").Value = Info.Range("i10").Value
        .Range("c8").Value = Info.Range("i12").Value
        .Range("c10").Value = Info.Range("i14").Value
        .Range("G4").Value = Info.Range("m8").Value
        .Range("G6").Value = Info.Range("m10").Value
        .Range("G8").Value = Info.Range("m12").Value
        .Range("G10").Value = Info.Range("m14").Value
    End With
    Set clm = Movimentacao.Range("y9:y" & Movimentacao.Cells(Rows.Count, "y").End(xlUp).Row)
    If clm.Cells(1).Value = vbNullString Then

        ImprimeFichaExt.Range("b15") = "Não houve movimentação"
        With ImprimeFichaExt.Range("b15:h15")
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 20
        End With


        GoTo servico:
    Else
        i = Movimentacao.ListObjects("tbHistMov14").DataBodyRange.Rows.Count - WorksheetFunction.CountBlank(Movimentacao.Range("y8:y" & Movimentacao.ListObjects("tbHistMov14").DataBodyRange.Rows.Count))

        Movimentacao.Range("X9:AD" & i).Copy

        With ImprimeFichaExt

            .Range("b1048576").End(xlUp).CurrentRegion.Offset(1, 0).PasteSpecial xlPasteValues
            .Range("b1048576").End(xlUp).CurrentRegion.Offset(1, 0).PasteSpecial xlPasteFormats
            ImprimeFichaExt.Range(Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Address & ":H1048576").Clear

servico:
            Application.CutCopyMode = True
            Serviços.Range("bc10").CurrentRegion.Copy

            ImprimeFichaExt.Range("b1048576").End(xlUp).Offset(2, 0).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            ImprimeFichaExt.Range(Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Address & ":H1048576").ClearContents
            
            k = Serviços.ListObjects("tbHistServ13").DataBodyRange.Rows.Count - WorksheetFunction.CountBlank(Serviços.Range("x8:x" & Serviços.ListObjects("tbHistServ13").DataBodyRange.Rows.Count))
            Serviços.Range("X9:AD" & k).Copy
            
            LINHA = ImprimeFichaExt.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
            ImprimeFichaExt.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            ImprimeFichaExt.Cells(LINHA, 2).PasteSpecial xlPasteFormats
            ' k = Serviços.ListObjects("tbHistServ13").DataBodyRange.Rows.Count - WorksheetFunction.CountBlank(Serviços.Range("y8:y" & Serviços.ListObjects("tbHistServ13").DataBodyRange.Rows.Count))
            LINHA = ImprimeFichaExt.Cells(Rows.Count, 2).End(xlUp).Row
        End With

    End If
    With ImprimeFichaExt.Range("D:D,F:F").Columns
        .AutoFit
    End With
    ImprimeFichaExt.Range(Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Address & ":H1048576").ClearContents
    ImprimeFichaExt.PageSetup.PrintArea = "$B$2:$H$" & LINHA
    ImprimeFichaExt.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Extintor_numero_" & Info.Cells(8, 9).Value & "_" & ".pdf", Quality:=xlQualityStandard, _
                                                                                                               IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True

    Application.GoTo Info.Range("frmCadastroSerie")
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub




