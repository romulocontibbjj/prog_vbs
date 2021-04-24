VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   2520
   ClientTop       =   3285
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9660
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   8880
         Top             =   4440
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   3240
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3480
         Width           =   2295
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   3015
      End
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   3240
         TabIndex        =   8
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   240
         Top             =   1920
      End
      Begin MSMAPI.MAPIMessages MAPIMessages1 
         Left            =   8040
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession MAPISession1 
         Left            =   7320
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin VB.CommandButton cmd_enviar 
         Caption         =   "&Enviar"
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   2175
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   8415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "romulo@intec.com.br"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lab_time 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8520
         TabIndex        =   14
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label lab_anexo 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Anexo:"
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Mensagem:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Assunto:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Destinatário:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xassunto As Integer

Private Sub cmd_enviar_Click()

MAPISession1.Action = 1
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Compose

'DESTINATÁRIO
MAPIMessages1.RecipAddress = Text1.Text
MAPIMessages1.AddressResolveUI = True
MAPIMessages1.ResolveName

'TITULO
MAPIMessages1.MsgSubject = UCase(Text2.Text)
'MENSAGEM
MAPIMessages1.MsgNoteText = UCase(Text3.Text)
MAPIMessages1.Send False

'anexa no final da mensagem
MAPIMessages1.AttachmentPosition = Len(MAPIMessages1.MsgNoteText)
'define o tipo de dados do anexo
MAPIMessages1.AttachmentType = mapData
'da um nome ao anexo
MAPIMessages1.AttachmentName = File1.FileName
'define o caminho e nome do arquivo a anexar
MAPIMessages1.AttachmentPathName = lab_anexo


MAPISession1.SignOff

End Sub

Private Sub Combo1_Click()

If Combo1.ListIndex = 0 Then
    File1.Pattern = "*.txt"
Else
    File1.Pattern = "*.*"
End If



End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path


End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive


End Sub

Private Sub File1_Click()
lab_anexo.Caption = Dir1.Path & "\" & File1.FileName
End Sub

Private Sub Form_Load()
Drive1.Drive = "C:"
Dir1.Path = Drive1.Drive & "\informa"
lab_time.Caption = Time


End Sub

Private Sub Timer1_Timer()

xassunto = xassunto + 1

Text2.Text = "ASSUNTO: " & xassunto

Text3.Text = "NÚMERO DE TENTATIVAS: " & xassunto

cmd_enviar_Click


If xassunto = 10 Then
    MsgBox "10 emails"
    Timer1.Enabled = False
End If





End Sub

Private Sub Timer2_Timer()

lab_time.Caption = Time


End Sub
