VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents crExport As vbSendMail.clsSendMail
Attribute crExport.VB_VarHelpID = -1

Private Sub Command1_Click()
With crExport

.SMTPHost = "smtp.primeti.com"
   .SMTPPort = 25
   .Username = "rconti@primeti.com"
   .Password = "conti601515"
   .UseAuthentication = True
   .Connect
   If Err Then MsgBox Err.Description
   .FromDisplayName = "Atendimento - Omnia Saúde"
   .Subject = "Guia de Funcionário"
   .From = "rconti@primeti.com"
   .Recipient = "rconti@primeti.com"
   .Message = "TESTANDO"
   .AsHTML = True
   
   .Send
   
End With

If Err Then MsgBox Err.Description

End Sub

Private Sub Form_Load()

Set crExport = New vbSendMail.clsSendMail

End Sub
