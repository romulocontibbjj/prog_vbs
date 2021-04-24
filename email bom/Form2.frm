VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   1215
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   555
      Left            =   4800
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Text            =   "romulo@intec.com.br"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim objSession As Object
Dim objMessage As Object

Set objSession = CreateObject("MAPI.SESSION")
objSession.Logon "Your Profile Name", , False, False

Set objMessage = objSession.Inbox.Messages.Add

'ASSUNTO
objMessage.Subject = UCase(Text2.Text)

'MENSAGEM
objMessage.Text = Text3.Text

'CAMINHO E ANEXO
objMessage.Attachments.Add Text4.Text

'DESTINATARIO
objMessage.Recipients.Add LCase(Text1.Text)
objMessage.Recipients.Resolve
objMessage.Send


End Sub

Private Sub Command2_Click()
Unload Me

End Sub
