VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Public WithEvents LabelDia As MSForms.Label
Attribute LabelDia.VB_VarHelpID = -1
Public labelRotulo As MSForms.Label
Public WithEvents labelproximo As MSForms.Label
Attribute labelproximo.VB_VarHelpID = -1
Public WithEvents labelanterior As MSForms.Label
Attribute labelanterior.VB_VarHelpID = -1
Public CalendarioFilho As clsCalendario
Public DiaCalendario As Date           'O dia que cada lbl dia representa

Public Property Set FormCalendario(MeuFormCalendario As MSForms.UserForm)

    Dim iLabel As MSForms.Label        'contador de lbs "Dias"
    Dim irow  As Long, icol As Long    'contadores dos "For"
    Dim largura As Single, altura As Single 'Largura e altura de cada lbl do calendario
    Dim esp   As Single                ' espaçamento entre os lbl do calendario
    MeuFormCalendario.BackColor = FUNDO_CINZA_ESCURO
    esp = 1
    '# Pega a largura interna do form calendario e subtrai do produto do espaçamento
    '# pelo numero de lbls que cabem na largura do calendario + 1 e divide tudo pelo
    '# pelo numero de lbls que cabem na largura do calendario.
    largura = (MeuFormCalendario.InsideWidth - esp * 8) / 7
    '# Pega a altura interna do form calendario e subtrai do produto do espaçamento
    '# pelo numero de lbls que cabem na altura do calendario + 1 e divide tudo pelo
    '# pelo numero de lbls que cabem na altura do calendario.
    altura = (MeuFormCalendario.InsideHeight - esp * 9) / 8
    Set ColCalendario = New Collection 'cria nova coleção p/armazenar os itens calendario filho
    ColCalendario.Add Me               'adiciona a classe mãe a coleção ColCalendario Video 6 1:21
    Me.DiaCalendario = Date            'atribui a data atual a propriedade diacalendario da classe mãe
    
    
    For irow = 3 To 8                  'Percorre as colunas do calendario a partir da terceira linha
    
        For icol = 1 To 7              ' percorre as linhas do formulario
            Set CalendarioFilho = New clsCalendario ' cria a classe calendariofilho
            Set iLabel = MeuFormCalendario.Controls.Add("Forms.Label.1") 'armazena o controle na valriavel
            formatlabeldia iLabel, largura, altura, irow, icol, esp 'chama função que formata os lbls
            Set CalendarioFilho.LabelDia = iLabel 'adiciona os novos lbls a cls calendariodia como tipo LabelDia
            ColCalendario.Add CalendarioFilho 'alimenta a coleção
        Next icol
    
    Next irow
    
    'Criar dias da semana
    
    For icol = 1 To 7                  ' percorre as linhas do formulario
        Set CalendarioFilho = New clsCalendario ' cria a classe calendariofilho
        Set iLabel = MeuFormCalendario.Controls.Add("Forms.Label.1") 'armazena o controle na valriavel
        formatlabeldia iLabel, largura, altura, 2, icol, esp 'chama função que formata os lbls
        iLabel.Caption = WeekdayName(icol, True)
        iLabel.BackStyle = fmBackStyleTransparent
        iLabel.BorderStyle = fmBorderStyleNone
    Next icol
    
    'criar label rotulo
    
    Set labelRotulo = MeuFormCalendario.Controls.Add("Forms.label.1")
    formatlabeldia labelRotulo, largura * 5 + (esp * 5), altura, 1, 1, esp
    
    With labelRotulo
    
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .BorderStyle = fmBorderStyleNone
        .Caption = Format$(Date, "mmmm" & " \d\e " & "yyyy")
        
    End With
    'criar label próximo
    Set labelproximo = MeuFormCalendario.Controls.Add("Forms.Label.1")
    formatlabeldia labelproximo, largura, altura, 1, 7, esp
    
    With labelproximo
    
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .Caption = ChrW$(&H25BC)       'https://stackoverflow.com/questions/23559158/how-do-i-assign-a-cell-formula-containing-symbols-using-vba/23559562
    
    End With
    
    Set labelanterior = MeuFormCalendario.Controls.Add("Forms.Label.1")
    formatlabeldia labelanterior, largura, altura, 1, 6, esp
    
    With labelanterior
    
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .Caption = ChrW$(&H25B2)
    
    End With
    
    
    AtualizarDias Date
    
    
    
    
End Property

' função para formatar os lablesdia do calendario
Public Sub formatlabeldia(ilabledia As MSForms.Label, largura As Single, altura As Single, irow As Long, icol As Long, esp As Single)

    'definindo tamanho da fonte adequado aos lbls
    
    Dim menordimensao As Single
    
    If altura > largura Then
        menordimensao = largura
    Else
        menordimensao = altura
    End If
    

    With ilabledia

        .BorderStyle = fmBorderStyleSingle
        .Width = largura
        .Height = altura
        .Top = irow * esp + altura * (irow - 1) 'explicado na aula 6
        .Left = icol * esp + largura * (icol - 1) 'explicado na aula 6
        .Font.Size = menordimensao * 0.7 '70% do tamanho do lbl
        .TextAlign = fmTextAlignCenter
        .BackColor = FUNDO_CINZA_MEDIO
        .ForeColor = FONTE_BRANCA
        
    End With



End Sub

Private Sub AtualizarDias(DataReferencia As Date) ' data utilizada para calculos
    'Percorre todos os lbls dia e poe uma data em cada um.
    Dim iCalendario As clsCalendario
    Dim i     As Integer
    Dim primDiaMes As Date, PrimDiaCalendario As Date
    
    primDiaMes = DateSerial(Year(DataReferencia), Month(DataReferencia), 1)
    PrimDiaCalendario = primDiaMes - (Weekday(primDiaMes) - 1)
    
    For i = 2 To ColCalendario.Count
    
        Set iCalendario = ColCalendario.Item(i)
        
        iCalendario.DiaCalendario = PrimDiaCalendario + i - 2
        iCalendario.LabelDia.Caption = Day(PrimDiaCalendario + i - 2)
        If Month(iCalendario.DiaCalendario) <> Month(DataReferencia) Then
            iCalendario.LabelDia.ForeColor = 10066329
        Else
            iCalendario.LabelDia.ForeColor = rgbWhite
        End If
        
        If iCalendario.DiaCalendario = Date Then
            iCalendario.LabelDia.BackColor = FUNDO_CINZA_ESCURO
        Else
            iCalendario.LabelDia.BackColor = FUNDO_CINZA_MEDIO
        End If
        
    
    Next i
    
    Set iCalendario = ColCalendario.Item(1)
    iCalendario.DiaCalendario = DataReferencia
    iCalendario.labelRotulo.Caption = Format$(DataReferencia, "mmmm" & " \d\e " & "yyyy")

End Sub

Private Sub labelanterior_Click()
    Dim novoMes As Date
    novoMes = DateAdd("m", -1, Me.DiaCalendario)
    AtualizarDias novoMes
End Sub

Private Sub LabelDia_Click()
    Dim calendarioMae As clsCalendario
    Set calendarioMae = ColCalendario.Item(1)
    calendarioMae.DiaCalendario = Me.DiaCalendario
    Unload Me.LabelDia.Parent

End Sub

Private Sub labelproximo_Click()
    Dim novoMes As Date

    novoMes = DateAdd("m", 1, Me.DiaCalendario)
    AtualizarDias novoMes

End Sub

