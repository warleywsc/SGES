Attribute VB_Name = "MODOcultaMenu"
Option Explicit
'@Folder("SGES2020")

Public Sub ocultaMenuInfo()
   
    If ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = False Then
        Info.Unprotect
        ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = True

        ActiveSheet.Shapes.Range(Array("btnocultarmenu")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Exibir Menu"
        ActiveSheet.Shapes.Range(Array("btnSalvaAtualExt")).Left = 506.29
        If Range("e37").EntireRow.Hidden = False Then
            Range("frmNovoExtintorSerie").Select
        ElseIf Range("e66").EntireRow.Hidden = False Then
            Range("I67").Select
        ElseIf Range("e103").EntireRow.Hidden = False Then
            Range("I103").Select
        
        ElseIf Range("e8").EntireRow.Hidden = False Then
            Range("frmCadastroSerie").Select
        End If
        Info.Protect
    Else
        Info.Unprotect
        ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = False
        ActiveSheet.Shapes.Range(Array("btnocultarmenu")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Ocultar Menu"
        ActiveSheet.Shapes.Range(Array("btnSalvaAtualExt")).Left = 638.89
        Range("frmCadastroSerie").Select
        If Range("e37").EntireRow.Hidden = False Then
            Range("frmNovoExtintorSerie").Select
        ElseIf Range("e66").EntireRow.Hidden = False Then
            Range("I67").Select
        ElseIf Range("e103").EntireRow.Hidden = False Then
            Range("I103").Select
        
        ElseIf Range("e8").EntireRow.Hidden = False Then
            Range("frmCadastroSerie").Select
        End If
    
        Info.Protect
    End If
End Sub

Public Sub ocultaMenuOutros()
    
    If ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = False Then
        



        
        ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = True

        ActiveSheet.Shapes.Range(Array("btnocultarmenu")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Exibir Menu"
        primlinha
        
    Else
        
        ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = False
        ActiveSheet.Shapes.Range(Array("btnocultarmenu")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Ocultar Menu"
        primlinha
    
        
    End If
    
End Sub

Public Sub ocultaMenuPesquisa()
 
    If ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = False Then
        


        Pesquisa.Unprotect "brigada"
        
        ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = True

        ActiveSheet.Shapes.Range(Array("btnocultarmenu")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Exibir Menu"
        Pesquisa.Range("I3").Activate
       
        Pesquisa.Protect "brigada"
    Else
        Pesquisa.Unprotect "brigada"
        ActiveSheet.Columns.Item("B:D").EntireColumn.Hidden = False
        ActiveSheet.Shapes.Range(Array("btnocultarmenu")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Ocultar Menu"
        Pesquisa.Range("I3").Activate
        Pesquisa.Protect "brigada"
       
    
        
    End If
End Sub

