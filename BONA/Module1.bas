Attribute VB_Name = "Module1"
Option Explicit
Public xusuario As String
Public xsenha As String
Public Str_oper As String
Public Function limpa_tela(Form As Integer)
Dim vobj As Control

If Form = 1 Then
    For Each vobj In frm_Pagamentos
        
        If TypeOf vobj Is TextBox Then
            vobj.Text = Empty
         
        ElseIf TypeOf vobj Is ComboBox Then
            vobj.ListIndex = -1
        
        ElseIf TypeOf vobj Is MaskEdBox Then
            vobj.Mask = ""
            vobj.Text = ""
            vobj.Mask = "99/99/9999"
               
        End If
    Next
    frm_Pagamentos.mask_recebimento.SetFocus
End If

End Function


Public Function branco_separa(xstring As String)

Dim x As Integer

    For x = 1 To Len(xstring)
        branco_separa = Mid(xstring, x, 1)
            If branco_separa = " " Then
                branco_separa = Mid(xstring, 1, x - 1)
                Exit For
            End If
    Next
    
End Function

Public Function Xbusca(ByVal xArquivo As String, Optional xDiretorio As Boolean)

If Trim$(Len(xArquivo)) > 0 Then
    If IsMissing(xDiretorio) Or xDiretorio = False Then
        Xbusca = (Dir$(xArquivo) <> "")
    Else
        Xbusca = (Dir$(xArquivo, vbDirectory) <> "")
    End If
 End If
End Function

Public Function Carac(xstring As String, xproib As String)
Dim x As Integer
Dim xqtd As Integer
Dim xaux As String

xqtd = Len(xstring)
xaux = Mid(xstring, xqtd, 1)

If IsNumeric(xaux) = False Then
        
    If xaux <> xproib Then
        
        MsgBox "Caracter Inválido"
        Carac = Mid(xstring, 1, xqtd - 1)
    Else
        
        Carac = xstring
        
    End If
Else
    
    Carac = xstring
    
End If


Public Function EMAIL(xenvio As String, xanexo1 As String, xassunto As String, xmensagem1 As String)




Dim xMail As Outlook.Application
Dim xmensagem As Outlook.MailItem
Dim xanexo As Outlook.Items



Set xMail = CreateObject("Outlook.Application")
Set xmensagem = CreateItem(olMailItem)


xmensagem.To = xenvio

xmensagem.Subject = xassunto

    
xmensagem.Body = xmensagem1
                
xmensagem.Importance = olImportanceHigh

xmensagem.Attachments.Add xanexo1 & ".xls"

xmensagem.Send


Set xmensagem = Nothing
Set xMail = Nothing

End Function

