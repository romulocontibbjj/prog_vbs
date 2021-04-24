VERSION 5.00
Begin VB.Form frmLogo 
   Caption         =   "Logo Intec"
   ClientHeight    =   7095
   ClientLeft      =   1065
   ClientTop       =   2310
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   12060
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   6360
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   3960
      ScaleHeight     =   1185
      ScaleWidth      =   3825
      TabIndex        =   7
      Top             =   5640
      Width           =   3852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   6
      Text            =   "123456789"
      Top             =   5520
      Width           =   3732
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   612
      Left            =   9360
      TabIndex        =   5
      Top             =   5520
      Width           =   975
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Small"
      Height          =   192
      Index           =   0
      Left            =   8040
      TabIndex        =   4
      Top             =   5400
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Medium"
      Height          =   192
      Index           =   1
      Left            =   8040
      TabIndex        =   3
      Top             =   5640
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Large"
      Height          =   192
      Index           =   2
      Left            =   8040
      TabIndex        =   2
      Top             =   5880
      Width           =   972
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   615
      Left            =   10560
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.PictureBox piclogo 
      Height          =   5055
      Left            =   120
      Picture         =   "frmLogo.frx":0000
      ScaleHeight     =   4995
      ScaleWidth      =   12915
      TabIndex        =   0
      Top             =   120
      Width           =   12975
   End
   Begin VB.Label Label1 
      Caption         =   "Enter text for barcode"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "This barcode is copied to clipboard"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   5400
      Width           =   3615
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Printer.PaintPicture frmLogo.Picture1, 2000, 2000, frmLogo.Picture1.Picture.Width * 0.5, frmLogo.Picture1.Picture.Height * 0.35        'BARCODE

End Sub

Private Sub Command2_Click()

End Sub

Private Sub optSize_Click(Index As Integer)
    Picture1.ScaleMode = 3
    
    Select Case Index
    Case 0
        Picture1.Height = Picture1.Height * (1.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 8
    Case 1
        Picture1.Height = Picture1.Height * (2.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 10
    Case 2
        Picture1.Height = Picture1.Height * (3 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 14
    End Select


    Call Text1_Change

End Sub
Private Sub Text1_Change()
    
    Call DrawBarcode(Text1, Picture1)
    
    MinWidth = 2 * Text1.Left + Text1.Width
    pw = 2 * Picture1.Left + Picture1.Width
    fw = MinWidth
    If pw > fw Then fw = pw
'    Form1.Width = fw

End Sub

