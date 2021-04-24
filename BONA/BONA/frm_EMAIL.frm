VERSION 5.00
Begin VB.Form frm_EMAIL 
   Caption         =   "ENVIO DE EMAILS"
   ClientHeight    =   7200
   ClientLeft      =   2715
   ClientTop       =   2190
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8535
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   5280
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENVIO DE EMAILS"
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "ANEXOS:"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frm_EMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   For x = 0 To List1.ListCount - 1
   
   Dim xdest As String
   Dim xnome As String
   Dim xarq As String
   Dim xMail As Outlook.Application
    Dim xMensagem As Outlook.MailItem
    Set xMail = CreateObject("outlook.application")
    Set xMensagem = xMail.CreateItem(olMailItem)
    
    
        xarq = List1.List(x)
      
    xMensagem.To = "romulo@intec.com.br"
    xMensagem.Subject = "ROMULO CONTI"
    xMensagem.Body = "Segue arquivo anexo."
    
    xMensagem.Attachments.Add xarq
   
    DoEvents
    xMensagem.Send
    
    DoEvents
    
    'xMail.Quit
    Set xMensagem = Nothing
    Set xMail = Nothing
    
    
    Next
    
    
    
End Sub

Private Sub Command2_Click()

List1.AddItem File1.Path & "\" & File1.FileName




End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
Dim xextensao As String

xextensao = Mid(File1.FileName, Len(File1.FileName) - 2, 3)

If xextensao = "jpg" Then

    Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
    
End If



End Sub

Private Sub Form_Load()
Drive1.Drive = "D:\"
Dir1.Path = Drive1.Drive
End Sub

