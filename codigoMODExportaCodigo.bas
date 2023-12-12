Attribute VB_Name = "MODExportaCodigo"
' Criado por Warley da Silva Concei��o em 23/10/2023

Sub ExportAllCode()
    Dim VBComp As VBIDE.VBComponent
    Dim SavePath As String
    Dim i As Integer
    
    ' Define o diret�rio onde os arquivos .bas ser�o salvos
    SavePath = "D:\brigada\SGES\codigo"
    
    ' Loop atrav�s de todos os componentes VBA (m�dulos, userforms, etc.)
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        ' Verifica o tipo de componente
        Select Case VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                VBComp.Export SavePath & VBComp.Name & ".bas"
            Case vbext_ct_MSForm
                VBComp.Export SavePath & VBComp.Name & ".frm"
            Case vbext_ct_Document
                ' Esta � uma planilha ou workbook. Voc� n�o pode exportar diretamente,
                ' mas voc� pode copiar o c�digo para um m�dulo padr�o e ent�o exportar.
        End Select
    Next VBComp
End Sub

