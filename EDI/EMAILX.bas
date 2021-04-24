Attribute VB_Name = "EMAILX"
Public Function xmail(xremet As String, xemail As String, xassunto As String, xcorpoemail As String, xarquivo As String)

Dim Xemail1 As Outlook.Application
Dim xmensagem As Outlook.MailItem
Dim x As Integer
Dim xAux As String
Dim xend As String
Dim xmais As Integer


Set Xemail1 = CreateObject("Outlook.Application")
Set xmensagem = Xemail1.CreateItem(olMailItem)
'xemail = "anderson_andrade@videolar.com.br;franklin@luftexpress.com.br"

xmensagem.To = xemail
xmensagem.Subject = xassunto
xmensagem.Body = xcorpoemail


xmensagem.Attachments.Add xarquivo

DoEvents
xmensagem.Send

Set xmensagem = Nothing
Set Xemail1 = Nothing




End Function
