Attribute VB_Name = "Módulo5"
'@Folder("Tests")
Option Explicit

'Sub prencherListBoxPlanilha(Consulta As ADODB.Recordset)
'    Dim montastr As String
'    Dim i     As Long, nlin As Long
'    Dim tamanhocol As Integer
'    Dim base  As Range
'    Dim colhead As Collection
'    PConsulta.Cells.Clear
'
'    '----------Nome das colunas ----------
'
'    With PConsulta
'
'        For i = 0 To Consulta.Fields.Count - 1
'            .Cells(1, i + 1) = Consulta.Fields(i).Name
'        Next i
'
'    End With
'    '----------Corpo da consulta ----------
'    PConsulta.Range("A2").CopyFromRecordset Consulta
'    nlin = PConsulta.Range("A1").CurrentRegion.Rows.Count
'    If nlin = 1 Then nlin = 2
'    Set base = PConsulta.Range(PConsulta.Cells(2, 1), PConsulta.Cells(nlin, Consulta.Fields.Count))
'
'    With FormPrincipal.lbPesquisa
'        .ColumnHeads = True
'        .ColumnCount = Consulta.Fields.Count
'        .RowSource = base.Address(external:=True) 'Fixa a fonte de dados na planilha definida em base
'        Set colhead = New Collection
'        '----------Autodimensionar largura das colunas da listbox ----------
'        PConsulta.Range("a1").CurrentRegion.Columns.AutoFit
'        montastr = ""
'        '----------monta a string que define o ColumnWidths ----------
'        For i = 1 To Consulta.Fields.Count
'            tamanhocol = PConsulta.Columns(i).Width
'            colhead.Add tamanhocol
'            montastr = montastr & (colhead.Item(i) * 1.3) & ";"
'
'        Next i
'
'        montastr = Left(montastr, Len(montastr) - 1)
'        .ColumnWidths = montastr
'        .TextAlign = fmTextAlignCenter
'
'    End With
'
'End Sub




Sub autodimlistbox()
    Dim montastr As String
    Dim i     As Long, nlin As Long
    Dim tamanhocol As Integer
    Dim base  As Range
    Dim colhead As Collection

    

    With frmMovimentaManutencao.ListBoxenvio
        Set colhead = New Collection
        '----------Autodimensionar largura das colunas da listbox ----------
        formenvio.Range("G8").CurrentRegion.Columns.AutoFit
        montastr = ""
        '----------monta a string que define o ColumnWidths ----------
        For i = 1 To 9
            tamanhocol = formenvio.Range("G8").CurrentRegion.Columns(i).Width
            colhead.Add tamanhocol
            montastr = montastr & (colhead.Item(i)) & ";"
         
        Next i
 
        montastr = Left(montastr, Len(montastr) - 1)
        
        .ColumnWidths = montastr
        .TextAlign = fmTextAlignCenter
        
    End With

End Sub
