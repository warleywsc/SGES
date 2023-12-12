Attribute VB_Name = "MODbuscaZona"
Option Explicit
'@Folder("SGES2020")

'---------------------------------------------------------------------------------------
' Programador.....: Warley S Conceição
' Contato...: warleywsc@gmail.com - Rotina: Public Sub buscazona()
    ' Data......: 16/11/2020
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub buscazona()
    Dim novolocal As Variant
    Dim lin   As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Info
        novolocal = .Cells.Range("M12")
      
        lin = 9
        
        Do Until locais.Cells(lin, 13) = vbNullString
            If locais.Cells(lin, 13) = novolocal Then

                Info.Range("M14").Value = locais.Cells(lin, 14)
                
                Exit Sub

            End If
            lin = lin + 1
        Loop

    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Public Sub buscazonaEXTnOVO()
    Dim novolocal As Variant
    Dim lin   As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Info
        novolocal = .Cells.Range("M41")
      
        lin = 9
        
        Do Until locais.Cells(lin, 12) = vbNullString
            If locais.Cells(lin, 12) = novolocal Then

                Info.Range("M43").Value = locais.Cells(lin, 13)
                
                Exit Sub

            End If
            lin = lin + 1
        Loop

    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

