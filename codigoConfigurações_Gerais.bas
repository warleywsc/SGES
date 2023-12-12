Attribute VB_Name = "Configurações_Gerais"
'@Folder("VBAProject")
Option Explicit


Public Const FUNDO_CINZA_CLARO_MOV As Long = 15921906 'DimGray #696969 (105,105,105)
Public Const FUNDO_CINZA_ESCURO_MOV As Long = 5855577 '11119017
'Public Const FUNDO_CINZA_MEDIO As Long = 8355711 'grey31  #4F4F4F (79,79,79)
'Public Const FONTE_BRANCA As Long = rgbWhite
'Public Const COR_FONTE_SECUNDARIA As Long = 5197647
'Public Const NOME_FONTE_PRINCIPAL As String = "Tw Cen MT"
'Public Const TAMANHO_FONTE_PRINCIPAL As Integer = 12
'Public Const ESP_HORIZONTAL As Single = 3
'Public Const COR_BORDACXTEXTO As Long = 8519755 'CadetBlue   #5F9EA0 (95,158,160)
'Public Const COR_CXTEXTO As Long = 5855577


'------------------------------------------------------

Public Const FUNDO_CINZA_CLARO As Long = 5855577 'DimGray #696969 (105,105,105)
Public Const FUNDO_CINZA_ESCURO As Long = 15921906 '11119017
Public Const FUNDO_CINZA_MEDIO As Long = 8355711 'grey31  #4F4F4F (79,79,79)
Public Const FONTE_BRANCA As Long = rgbWhite
Public Const COR_FONTE_SECUNDARIA As Long = 5197647
Public Const NOME_FONTE_PRINCIPAL As String = "Tw Cen MT"
Public Const TAMANHO_FONTE_PRINCIPAL As Integer = 12
Public Const ESP_HORIZONTAL As Single = 3
Public Const COR_BORDACXTEXTO As Long = 8519755 'CadetBlue   #5F9EA0 (95,158,160)
Public Const COR_CXTEXTO As Long = 5855577

Public Menu   As clsMenu
Public ColMenu As Collection
Public controle As clsControles
Public ColControle As Collection
Public BotaoAcao As clsBotaoAcao
Public ColBotaoAcao As Collection
Public calendario As clsCalendario
Public ColCalendario As Collection

Public Enum TipoValidacao
    
    ValObrigatorio
    ValData
    ValNumero
    ValTexto
        
End Enum

