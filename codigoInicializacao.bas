Attribute VB_Name = "Inicializacao"
'@Folder("VBAProject")
Option Explicit

Public Sub ConfiguracaoLayout()
    Dim LarguraLabel As Single
    Dim iControl As MSForms.control
    LarguraLabel = frmAlteraLocal.btSalvar.Width

    'formulário
    

    With frmAlteraLocal.frmMenuLateral
        .BackColor = FUNDO_CINZA_ESCURO
        .SpecialEffect = fmSpecialEffectFlat
        .Height = frmAlteraLocal.Height
        .Top = 0
        .Left = 0
        .Width = LarguraLabel + ESP_HORIZONTAL * 2
    
    End With

 

    'Percorre todos os controles do frame corpo

    For Each iControl In frmAlteraLocal.Controls

       
        
        
        '--------------Procura Listboxes-----------------
        
        
        If TypeOf iControl Is MSForms.ListBox Then
    
            With iControl
                .BackColor = frmAlteraLocal.BackColor
                .ForeColor = COR_FONTE_SECUNDARIA
                .BorderStyle = fmBorderStyleSingle
                .SpecialEffect = fmSpecialEffectFlat
            End With

        End If
        
        If TypeOf iControl Is MSForms.TextBox Or TypeOf iControl Is MSForms.ComboBox Then
    
            With iControl
                .SelectionMargin = True
            End With

        End If
        
    Next iControl
    
    'alimenta listbox
    
    

End Sub

Public Sub CriarMenu()
    Dim iBotao As MSForms.Label
    Set ColMenu = New Collection


    For Each iBotao In frmAlteraLocal.frmMenuLateral.Controls
        Set Menu = New clsMenu
        Set Menu.btMenu = iBotao

        ColMenu.Add Menu
        Menu.Index = ColMenu.Count
        Menu.PropriedadePadrao

    Next


End Sub

Public Sub CriarControles()

    Dim iControl As MSForms.control

    Set ColControle = New Collection

    For Each iControl In frmAlteraLocal.FrameCorpo.Controls

        If iControl.Tag = "texto" Then
    
            Set controle = New clsControles
            Set controle.campo = iControl
        
            ColControle.Add controle
            controle.PropriedadePadrao
            
       
        End If

    Next iControl

End Sub

Public Sub CriarBotaoAcao()

    Dim iControl As MSForms.control

    Set ColBotaoAcao = New Collection
    
    For Each iControl In frmAlteraLocal.FrameCorpo.Controls
    
        If TypeOf iControl Is MSForms.CommandButton Then
            
            Set BotaoAcao = New clsBotaoAcao
            Set BotaoAcao.CommandButton = iControl
            ColBotaoAcao.Add BotaoAcao
            BotaoAcao.PropriedadePadraoBotao
        End If
       
    Next iControl

End Sub


Public Sub ConfiguracaoLayoutFormMovimentacao()
    Dim LarguraLabel As Single
    Dim iControl As MSForms.control
    LarguraLabel = frmMovimentaManutencao.btMenuEnvio.Width

    'formulário
    

    With frmMovimentaManutencao.FrameMenu
        .BackColor = FUNDO_CINZA_ESCURO_MOV
        .SpecialEffect = fmSpecialEffectFlat
        .Height = frmMovimentaManutencao.Height
        .Top = 0
        .Left = 0
        .Width = LarguraLabel + ESP_HORIZONTAL * 2
    
    End With

 

    'Percorre todos os controles do frame corpo

    For Each iControl In frmMovimentaManutencao.Controls

       
        
        
        '--------------Procura Listboxes-----------------
        
        
        If TypeOf iControl Is MSForms.ListBox Then
    
            With iControl
                .BackColor = frmMovimentaManutencao.BackColor
                .ForeColor = COR_FONTE_SECUNDARIA
                .BorderStyle = fmBorderStyleSingle
                .SpecialEffect = fmSpecialEffectFlat
            End With

        End If
        
        If TypeOf iControl Is MSForms.TextBox Or TypeOf iControl Is MSForms.ComboBox Then
    
            With iControl
                .SelectionMargin = True
            End With

        End If
        
    Next iControl
    
    'alimenta listbox
    
    

End Sub

Public Sub CriarControlesFormMovimentacao()

    Dim iControl As MSForms.control

    Set ColControle = New Collection

    For Each iControl In frmMovimentaManutencao.FrameCorpo.Controls

        If iControl.Tag = "texto" Then
    
            Set controle = New clsControles
            Set controle.campo = iControl
        
            ColControle.Add controle
            controle.PropriedadeCustom
            
       
        End If

    Next iControl

End Sub
Public Sub CriarMenuFormMovimentacao()
    Dim iBotao As MSForms.Label
    Set ColMenu = New Collection


    For Each iBotao In frmMovimentaManutencao.FrameMenu.Controls
        Set Menu = New clsMenu
        Set Menu.btMenuMov = iBotao

        ColMenu.Add Menu
        Menu.IndexMov = ColMenu.Count
        Menu.PropriedadePadraoMov

    Next


End Sub
Public Sub CriarBotaoAcaoFormMovimentacao()

    Dim iControl As MSForms.control

    Set ColBotaoAcao = New Collection
    
    For Each iControl In frmMovimentaManutencao.FrameCorpo.Controls
    
        If TypeOf iControl Is MSForms.CommandButton Then
            
            Set BotaoAcao = New clsBotaoAcao
            Set BotaoAcao.CommandButton = iControl
            ColBotaoAcao.Add BotaoAcao
            BotaoAcao.PropriedadePadraoBotao
        End If
       
    Next iControl

End Sub

