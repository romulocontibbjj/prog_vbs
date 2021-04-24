Attribute VB_Name = "Module1"
Option Explicit

Public Function EMAIL(xtrans As String)
Dim X As Integer

Dim xMail As Outlook.Application
Dim xMensagem As Outlook.MailItem
Dim xanexo As Outlook.Items
Dim xtexto As String


Set xMail = CreateObject("Outlook.Application")
Set xMensagem = CreateItem(olMailItem)

xtexto = "Arquivo de Edi Importado. Dia: " & Date & " Hora: " & Time & Chr$(13) & "Total de Registros no Arquivo: " & frmEdiImport.lblTotReg.Caption & Chr$(13) & "Total de Baixas: " & frmEdiImport.lblBxGrav.Caption & Chr$(13) & "Total de Ocorrências Gravadas: " & frmEdiImport.lblOcorrGrav.Caption & Chr$(13) & "Total de Registros Criticados: " & frmEdiImport.lblCrit.Caption

xMensagem.To = "sheyla@intec.com.br; romulo@intec.com.br; silvia@intec.com.br"
xMensagem.Subject = "EDIS - " & UCase(xtrans)

If frmEdiImport.List1.ListCount = 0 Then
    
    xMensagem.Body = "Arquivo de Edi Importado. Dia: " & Date & " Hora: " & Time & Chr$(13) & "Total de Registros no Arquivo: " & frmEdiImport.lblTotReg.Caption & Chr$(13) & "Total de Baixas: " & frmEdiImport.lblBxGrav.Caption & Chr$(13) & "Total de Ocorrências Gravadas: " & frmEdiImport.lblOcorrGrav.Caption & Chr$(13) & "Total de Registros Criticados: " & frmEdiImport.lblCrit.Caption & Chr$(13) & Chr$(13) & "Romulo Conti"
                   
Else
          
    For X = 0 To frmEdiImport.List1.ListCount - 1
    
    xtexto = xtexto & Chr$(13) & Chr$(13) & frmEdiImport.List1.List(X)
    
    Next
    
    xMensagem.Body = xtexto & Chr$(13) & Chr$(13) & "Romulo Conti"
    
End If
                
                
                
xMensagem.Importance = olImportanceHigh

xMensagem.Send


Set xMensagem = Nothing
Set xMail = Nothing



End Function
