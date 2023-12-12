Attribute VB_Name = "MODCadastraObs"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub cadastraObs()
    ' Data......: 04/01/2021
    ' Descricao.: Cadastra observação na tabela locais
    '---------------------------------------------------------------------------------------
Public Sub cadastraObslocal()

    Dim localxarea As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    
    With Info
        localxarea = .Cells(12, 13).Value & " - " & .Cells(14, 9).Value
       
        lin = 9
      
        Do Until locais.Cells(lin, 10) = vbNullString
            If locais.Cells(lin, 8).Value & " - " & locais.Cells(lin, 9).Value = localxarea Then
                locais.Cells(lin, 11) = .Range("OBS").Value
             
                Exit Sub

            End If
            lin = lin + 1
        Loop
        
'        Do Until MapaAtual.Cells(lin, 14) = vbNullString
'            If MapaAtual.Cells(lin, 10) & " - " & MapaAtual.Cells(lin, 8) = localxarea Then
'                MapaAtual.Cells(lin, 27) = .Range("OBS").Value
'
'                Exit Sub
'
'            End If
'            lin = lin + 1
'        Loop
        .Range("F19").ClearContents
ErrorHandler:
        Exit Sub

    End With
 


End Sub
Public Sub cadastraObsext()

    Dim serie As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    
    With Info
        serie = .Range("i8")
       
        lin = 9
      
        Do Until Extintores.Cells(lin, 15) = vbNullString
            If Extintores.Cells(lin, 15) = serie Then
                Extintores.Cells(lin, 12) = .Range("m23").Value
             
                Exit Sub

            End If
            lin = lin + 1
        Loop
        
'        Do Until MapaAtual.Cells(lin, 14) = vbNullString
'            If MapaAtual.Cells(lin, 14) = Info Then
'                MapaAtual.Cells(lin, 27) = .Range("m23").Value
'
'                Exit Sub
'
'            End If
'            lin = lin + 1
'        Loop
        .Range("F21").ClearContents
ErrorHandler:
        Exit Sub

    End With
 


End Sub

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub ATUALIZAmAPAObs()
    ' Data......: 04/01/2021
    ' Descricao.: Cadastra observação na tabela MapaAtual
    '---------------------------------------------------------------------------------------
Public Sub ATUALIZAmAPAObs()

    Dim localxarea As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    Dim serie As String, seriemapa As String, obslocal As String, obsext As String
    Dim localmapa As String
    
    
    With Info
        localxarea = .Cells(12, 13).Value & " - " & .Cells(14, 9).Value
       serie = .Range("i8").Value
       obslocal = .Range("g23").Value
        obsext = .Range("M23").Value
        lin = 9
      
        
        Do Until MapaAtual.Cells(lin, 14) = vbNullString
        seriemapa = MapaAtual.Cells(lin, 14).Value
      localmapa = MapaAtual.Cells(lin, 10).Value & " - " & MapaAtual.Cells(lin, 8).Value
            If localmapa & seriemapa = localxarea & serie Then
            If obslocal <> "" Then
                MapaAtual.Cells(lin, 27) = "Observação Local: " & obslocal
                Else
                MapaAtual.Cells(lin, 27) = ""
                End If
                If obsext <> "" Then GoTo ext
                Exit Sub
            
            End If
            lin = lin + 1
        Loop
ext:
'                Do Until Extintores.Cells(lin, 15) = vbNullString
'            If seriemapaseriemapa = serie Then
            If obslocal <> "" Then
                MapaAtual.Cells(lin, 27) = "Observação Local: " & obslocal & vbNewLine _
                 & "Observação Extintor: " & obsext
                ElseIf obsext <> "" Then
                 MapaAtual.Cells(lin, 27) = "Observação Extintor: " & obsext
                
                Exit Sub
'End If
            End If
'            lin = lin + 1
'        Loop
  
ErrorHandler:
        Exit Sub

    End With
 


End Sub


'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub PopulaInfoOBS()
    ' Data......: 04/01/2021
    ' Descricao.: Popula o campo OBS no form Cadastro - Atualização em Info
    '---------------------------------------------------------------------------------------
Public Sub PopulaInfoOBS()

    Dim localxarea As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    Dim obslocal As String, obsext As String
    
    With Info
        localxarea = .Cells(12, 13).Value & " - " & .Cells(14, 9).Value
       
        lin = 9
      
      
        
        Do Until locais.Cells(lin, 10) = vbNullString
            If locais.Cells(lin, 10).Value = localxarea Then
           .Range("G23") = vbNullString
            obslocal = locais.Cells(lin, 11).Value
            If obslocal = vbNullString Then GoTo ext 'Exit Do
            
            
            .Range("G23") = obslocal
'            ATUALIZAmAPAObs
            Exit Do
            

            End If
            lin = lin + 1
        Loop
ext:
         lin = 9
 Do Until Extintores.Cells(lin, 15) = vbNullString
            If Extintores.Cells(lin, 15).Value = .Range("I8") Then
           .Range("m23") = vbNullString
            obsext = Extintores.Cells(lin, 12).Value
            If obsext = vbNullString Then
            ATUALIZAmAPAObs
            Exit Sub
            End If
            .Range("m23") = obsext
            ATUALIZAmAPAObs
            Exit Sub
            End If
            lin = lin + 1
            Loop
            Exit Sub
ErrorHandler:
        Exit Sub

    End With
 


End Sub
