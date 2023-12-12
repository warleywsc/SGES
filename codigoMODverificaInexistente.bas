Attribute VB_Name = "MODverificaInexistente"
Option Explicit
'@Folder("SGES2020")

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub verificaextinexistente()
    ' Data......: 18/11/2020
    ' Descricao.:Verifica se o Número do Extintor não existe e sugere cadastro novo
    '---------------------------------------------------------------------------------------
Public Sub verificaextinexistente()

    Dim novoext As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
   
    With Info
        novoext = UCase$(.Range("frmCadastroSerie").Value)
       
        lin = 9
      
        Do Until Extintores.Cells(lin, 15) = vbNullString
            If Extintores.Cells(lin, 15) = novoext Then
                Info.Range("E28").ClearContents
                Range("frmCadastroSerie").Select
                
                Exit Sub

            End If
            lin = lin + 1
        Loop
        ExtintorInexistente
ErrorHandler:
        Exit Sub

    End With



End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub verificaLocalinexistentefrmatual()
    ' Data......: 18/11/2020
    ' Descricao.:Verifica se local não existe e sugere cadastro novo
    '---------------------------------------------------------------------------------------
Public Sub verificaLocalinexistentefrmatual()

    Dim novolocal As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    
    With Info
        novolocal = .Range("M12").Value
       
        lin = 9
      
        Do Until locais.Cells(lin, 13) = vbNullString
            If locais.Cells(lin, 13) = novolocal Then
           
                Range("I14").Select
                Exit Sub

            End If
            lin = lin + 1
        Loop
        Range("I67").Value = Range("M12").Value
      
        LocalInexistentefrmAtual
ErrorHandler:
        Exit Sub

    End With
 


End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub verificaLocalinexistentefrmNovo()
    ' Data......: 18/11/2020
    ' Descricao.: Verifica se local não existe e sugere cadastro novo
    '---------------------------------------------------------------------------------------
Public Sub verificaLocalinexistentefrmNovo()
    Dim novolocal As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Info
        novolocal = .Range("frmCadastroLocal").Value
       
        lin = 9
     
        Do Until locais.Cells(lin, 8) = vbNullString
            If locais.Cells(lin, 8) = novolocal Then
           
                Range("I43").Select
                Exit Sub

            End If
            lin = lin + 1
        Loop
        
        LocalInexistentefrmNovo
ErrorHandler:
        Exit Sub

    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True


End Sub

