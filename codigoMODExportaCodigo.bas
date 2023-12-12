Attribute VB_Name = "MODExportaCodigo"
' Criado por Warley da Silva Conceição em 23/10/2023

Sub ExportAllCode()
    Dim VBComp As VBIDE.VBComponent
    Dim SavePath As String
    Dim i As Integer
    
    ' Define o diretório onde os arquivos .bas serão salvos
    SavePath = "D:\brigada\SGES\codigo"
    
    ' Loop através de todos os componentes VBA (módulos, userforms, etc.)
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        ' Verifica o tipo de componente
        Select Case VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                VBComp.Export SavePath & VBComp.Name & ".bas"
            Case vbext_ct_MSForm
                VBComp.Export SavePath & VBComp.Name & ".frm"
            Case vbext_ct_Document
                ' Esta é uma planilha ou workbook. Você não pode exportar diretamente,
                ' mas você pode copiar o código para um módulo padrão e então exportar.
        End Select
    Next VBComp
End Sub

