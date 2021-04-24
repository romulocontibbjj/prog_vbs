VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   840
      TabIndex        =   0
      Top             =   630
      Width           =   4740
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Left            =   315
         TabIndex        =   2
         Top             =   315
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   330
         Left            =   2310
         TabIndex        =   1
         Top             =   2415
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2415
         TabIndex        =   3
         Top             =   1050
         Width           =   1275
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xhoras As String
Dim xminutos As Integer
Dim xhora1 As Integer

xhoras = MaskEdBox1.Text


xminutos = Xminutes(xhoras)

Label1.Caption = xminutos


End Sub

