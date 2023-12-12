Attribute VB_Name = "Funcoes"
'@Folder("SGES2020")
Option Explicit

'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Function Userform_Check(  form_name As String)  As Integer
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Function Userform_Check( _
       form_name As String) _
        As Integer
    On Error GoTo TError

    ' Returns:
    '   0 - Userform is not loaded
    '   1 - Loaded but not visible
    '   2 - Loaded and visible

    ' mUtilities.Userform_Check()

    Dim frm   As Object

    Userform_Check = 0

    For Each frm In VBA.UserForms
        If frm.Name = form_name Then
            Userform_Check = 1

            If frm.Visible Then Userform_Check = 2

            Exit For
        End If
    Next frm

    ' Function Userform_Check( _
    form_name As String) _
    As Integer
fim:
    Exit Function
TError:
    MsgBox Err.Description, Err.Number, "Userform_Check()"
    GoTo fim
End Function
'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Sub testaform()
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Sub testaform()
    On Error GoTo TError
    If Userform_Check("frmMovimentaManutencao") = 2 Then

        MsgBox "Carregado"
    Else
        MsgBox "fechado"

    End If
fim:
    Exit Sub
TError:
    MsgBox Err.Description, Err.Number, "testaform()"
    GoTo fim
End Sub



'---------------------------------------------------------------------------------------
' Autor.....: WARLEY SC
' Contato...: warleywsc@gmail.com - Empresa: RW SOLUÇÕES - Rotina: Public Function IsLinhaMatch(ByVal LINHA As String, ParamArray Padroes() As Variant) As Boolean
    ' Data......: 19/07/2021
    ' Descricao.:
    '---------------------------------------------------------------------------------------
Public Function IsLinhaMatch(ByVal LINHA As String, ParamArray Padroes() As Variant) As Boolean
    On Error GoTo TError
    Dim resultado As Boolean
    Dim Contador As Byte
    Dim RegExp As Object
 
    'New VBScript_RegExp_55.RegExp
    If RegExp Is Nothing Then Set RegExp = VBA.CreateObject("VBScript.RegExp")
    With RegExp
        For Contador = 0 To UBound(Padroes) Step 1
            If Not Padroes(Contador) = VBA.vbNullString Then
                .Pattern = Padroes(Contador)
                If .Test(LINHA) Then
                    resultado = True
                    Exit For
                End If
            End If
        Next Contador
    End With
    IsLinhaMatch = resultado
    ' Exit Function
    'tratarerro:
fim:
    Exit Function
TError:
    MsgBox Err.Description, Err.Number, "IsLinhaMatch()"
    GoTo fim
End Function

'https://stackoverflow.com/a/1963263/10246702

Function checkInternetConnection() As Integer
    'code to check for internet connection
    'by Daniel Isoje
    On Error Resume Next
    checkInternetConnection = False
    Dim objSvrHTTP As ServerXMLHTTP
    Dim varProjectID, varCatID, strT As String
    Set objSvrHTTP = New ServerXMLHTTP
    objSvrHTTP.Open "GET", "http://www.google.com"
    objSvrHTTP.setRequestHeader "Accept", "application/xml"
    objSvrHTTP.setRequestHeader "Content-Type", "application/xml"
    objSvrHTTP.Send strT
    If Err = 0 Then
        checkInternetConnection = True
    Else
        MsgBox "Sem conexão de internet: " & Err.Description & "", 64, "Additt !"
    End If
End Function

