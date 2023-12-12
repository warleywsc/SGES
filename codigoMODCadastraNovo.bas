Attribute VB_Name = "MODCadastraNovo"
Option Explicit
'@Folder("SGES2020")

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub oldvalue()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub oldValue()
    Application.ScreenUpdating = False
    If Info.Range("I8").Value = "sinnaoMessageBox" Then
        ExtintorInexistente
    Else
        Exit Sub
    End If
    Application.ScreenUpdating = True

End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub ExtintorInexistente()
    ' Data......: 16/11/2020
    ' Descricao.: verifica se o capo série no form atualizar possui um valor não existente na tabela extintores
    ' e oferece a opção de cadastrar um novo extintor
    '---------------------------------------------------------------------------------------
Public Sub ExtintorInexistente()

    Dim Answer As String
    Dim MyNote As String
  

    With Info
        MyNote = "Extintor não encontrado. Deseja cadastrar um novo extintor?"

        'Pergunta se deseja cadastrar novo extintor.
        Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Extintor inexistente!")

        If Answer = vbNo Then
            'Se resposta for "Não", limpa e seleciona o campo "série"
            
            'Info.Range("I8,M8,I10,M10,I12,M12:N12,I14,M14:N14,I16:J16,M16,I18:J18,M18,I20:J20,M20,G23:N26").ClearContents
            Info.Range("E28").ClearContents
            Info.Range("frmCadastroSerie").ClearContents
            Info.Range("frmCadastroSerie").Activate
            Exit Sub:
        
        Else
            'Se resposta for "Sim", exibe form "frmNovo", preenche o campo
            '"série" e seleciona o campo "Tipo"
            Info.Range("E28") = "NOVO"
            frmNovo
            
            
        End If
                
        Exit Sub:
        
    End With
   

End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub LocalInexistentefrmNovo()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub LocalInexistentefrmNovo()
    Dim Answer As String
    Dim MyNote As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    
    MyNote = "Este local não existe! Deseja cadastrar novo local?"

    
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Local inexistente!")

    If Answer = vbNo Then
      
        Info.Range("M41:N41").ClearContents
        Info.Range("M41:N41").Activate
        
    Else
        
        Info.Range("frmNovoLocal").Value = Info.Range("frmCadastroLocal").Value
        Info.Range("M41:N41").ClearContents
        frmLocalAtualiza
        
        Info.Range("frmNovoLocal").Activate
    End If
    Info.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub LocalInexistentefrmAtual()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub LocalInexistentefrmAtual()
    Dim Answer As String
    Dim MyNote As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    
    MyNote = "Este local não existe! Deseja cadastrar novo local?"

    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Local inexistente!")

    If Answer = vbNo Then
     
        Info.Range("M12:N12").ClearContents
        Info.Range("M12:N12").Activate
        
    Else
   
    
        frmLocalAtualiza
        Info.Range("frmNovoLocal").Value = Info.Range("M12:N12").Value
        Info.Range("frmNovoLocal").Activate
    End If
    Info.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub sinnaoMessageBox()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub sinnaoMessageBox()
    Dim Answer As String
    Dim MyNote As String
    Application.ScreenUpdating = False

 
    MyNote = "Deseja cadastrar um novo extintor?"

  
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Extintor inexistente!")

    If Answer = vbNo Then
       
        old
    Else
       
    
        frmNovo
        Info.Range("E69").Value = Info.Range("E36").Value
        Info.Range("L69").Select
    End If
    Application.ScreenUpdating = True

End Sub

Public Sub old()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Info.Unprotect
    Range("E36").Value = Range("B37").Value
    Info.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

