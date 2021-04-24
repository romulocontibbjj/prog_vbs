VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Mensagem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MailToCompleto"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   4635
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   4575
      TabIndex        =   7
      Top             =   4350
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Form1.frx":5324
      Left            =   5310
      List            =   "Form1.frx":5331
      TabIndex        =   5
      Text            =   "3 - Normal"
      Top             =   1020
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Anexar Arquivo"
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2730
      Width           =   6465
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   225
      Left            =   1260
      TabIndex        =   3
      Text            =   "Assunto"
      Top             =   1860
      Width           =   3675
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   225
      Left            =   1260
      TabIndex        =   2
      Text            =   "vb10000@ieg.com.br"
      Top             =   1560
      Width           =   3675
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   225
      Left            =   1260
      TabIndex        =   1
      Text            =   "vb10000@ieg.com.br"
      Top             =   1260
      Width           =   3675
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   225
      Left            =   1260
      TabIndex        =   0
      Text            =   "vb10000@ieg.com.br"
      Top             =   960
      Width           =   3675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Prioridade"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   810
      Width           =   1275
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   3750
      MouseIcon       =   "Form1.frx":5356
      MousePointer    =   99  'Custom
      ToolTipText     =   "Anexar Arquivo"
      Top             =   300
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   90
      MouseIcon       =   "Form1.frx":5660
      MousePointer    =   99  'Custom
      ToolTipText     =   "Enviar Mensagem"
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "Mensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    GeraMensagem "C:\WINDOWS\Desktop\assunto.eml", _
                 "luciochaves@quick.com.br", _
                 Text1, _
                 Left(Combo1, 1), _
                 Text2, _
                 Text3, _
                 Text4, _
                 Text5, _
                 CommonDialog1.filename
    ShellExecute Me.hwnd, "OPEN", "C:\WINDOWS\Desktop\assunto.eml", 0&, "", 0
End Sub
Private Sub Image2_Click()
    CommonDialog1.Action = 1
End Sub
Private Sub Text4_Change()
    Me.Caption = Text4.Text
End Sub

