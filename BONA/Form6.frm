VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7995
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   6585
   LinkTopic       =   "Form6"
   ScaleHeight     =   7995
   ScaleWidth      =   6585
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin BONAGURA.SComboBox SComboBox2 
      Height          =   300
      Left            =   2160
      TabIndex        =   2
      Top             =   4560
      Width           =   1800
      _extentx        =   3175
      _extenty        =   529
      font            =   "Form6.frx":1CDA
      maxlistlength   =   -1
      numberitemstoshow=   -1
      shadowcolortext =   6582129
   End
   Begin BONAGURA.SComboBox SComboBox1 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2055
      _extentx        =   3625
      _extenty        =   556
      appearancecombo =   14
      backcolor       =   12648447
      gradientcolor1  =   32768
      gradientcolor2  =   12648384
      font            =   "Form6.frx":1D06
      highlightbordercolor=   12648384
      listcolor       =   16777152
      listgradient    =   -1  'True
      maxlistlength   =   -1
      numberitemstoshow=   -1
      shadowcolortext =   6582129
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000008&
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim x As Integer
For x = 1 To 10
SComboBox1.AddItem x, &HFF&



Next

End Sub


