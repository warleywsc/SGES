Attribute VB_Name = "MODprocuraduplicado"
Option Explicit

'@Folder("SGES2020")
Public Sub procuralocalduplicado()
    Dim novolocal As Variant
    Dim lin   As Long
    limpafiltrolocal
    With Info
        novolocal = .Cells.Range("frmNovoLocal")
       
        lin = 9
        
        Do Until locais.Cells(lin, 12) = vbNullString
            If locais.Cells(lin, 12) = novolocal Then

                Info.Cells.Range("frmNovoLocal").MergeArea.ClearContents
                Info.Select
                Info.Range("frmNovoLocal").Activate
                MsgBox "O local já existe. Digite um local novo."
                Exit Sub

            End If
            lin = lin + 1
        Loop

    End With



End Sub

Public Sub procuraextduplicado()
    Dim novoext As Variant
    Dim lin   As Long
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Info
        novoext = .Range("frmNovoExtintorSerie").Value
       
        lin = 9
        
        Do Until Extintores.Cells(lin, 7) = vbNullString
            If Extintores.Cells(lin, 7) = novoext Then





              
                Info.Range("frmNovoExtintorSerie").ClearContents
                Info.Range("frmNovoExtintorSerie").Activate
                
                MsgBox "O extintor já existe. Digite um novo número de série."
                




            End If
            lin = lin + 1
        Loop
ErrorHandler:
        Exit Sub

    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True


End Sub

Public Sub verificaSerieDuplicado()
    Dim LINEXT As Long
    Dim SerieEXT As Range
    Dim serieinfo As Range

    
    LINEXT = 9
     
         
    Do Until Extintores.Cells(LINEXT, 15) = vbNullString
        Set serieinfo = Info.Cells(35, 9)
        For Each SerieEXT In serieinfo
            Do Until Extintores.Cells(LINEXT, 15) = vbNullString
                Set SerieEXT = Extintores.Cells(LINEXT, 15)
           
                
                If serieinfo.Value = SerieEXT.Value Then
                    MsgBox "O Extintor já existe. Favor inserir um novo número de série."
                    
                    Info.Range("I37").ClearContents
                    Info.Range("i37").Activate
                    Info.Calculate: Exit Sub
                    
                    
                        
                Else
                        
                    LINEXT = LINEXT + 1
                End If
                    
                   
            Loop
              
          
         
            
        Next SerieEXT
       
    Loop
    'Exit Sub
End Sub

