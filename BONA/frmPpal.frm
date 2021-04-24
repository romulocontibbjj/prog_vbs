VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPpal 
   BackColor       =   &H00E9DEDB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SComboBox 1.0.4 - fred.cpp & HACKPRO TM 2004 @ Colombia - México"
   ClientHeight    =   5535
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "A&bout"
      Height          =   375
      Left            =   3450
      MouseIcon       =   "frmPpal.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1680
      Width           =   1125
   End
   Begin VB.TextBox txtMaxListLenght 
      ForeColor       =   &H00C56A31&
      Height          =   285
      Left            =   3585
      TabIndex        =   3
      Top             =   330
      Width           =   1590
   End
   Begin VB.TextBox txtAddItem 
      ForeColor       =   &H00C56A31&
      Height          =   285
      Left            =   1410
      TabIndex        =   14
      Text            =   "HACKPRO TM"
      Top             =   1935
      Width           =   1860
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "&Add Item"
      Height          =   405
      Left            =   120
      MouseIcon       =   "frmPpal.frx":0894
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1875
      Width           =   1110
   End
   Begin VB.TextBox txtSearchItem 
      ForeColor       =   &H00C56A31&
      Height          =   285
      Left            =   1410
      TabIndex        =   12
      Text            =   "fred_cpp"
      Top             =   1365
      Width           =   1860
   End
   Begin VB.CommandButton cmdTextItem 
      Caption         =   "Text Ite&m"
      Height          =   405
      Left            =   2190
      MouseIcon       =   "frmPpal.frx":0B9E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   105
      Width           =   915
   End
   Begin VB.CommandButton cmdSearchItem 
      Caption         =   "&Search Item"
      Height          =   405
      Left            =   120
      MouseIcon       =   "frmPpal.frx":0EA8
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1290
      Width           =   1110
   End
   Begin VB.CommandButton cmdDisabled 
      Caption         =   "&Enabled"
      Height          =   375
      Left            =   3450
      MouseIcon       =   "frmPpal.frx":11B2
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   720
      Width           =   1125
   End
   Begin VB.ComboBox cmbAlign 
      ForeColor       =   &H00C56A31&
      Height          =   315
      ItemData        =   "frmPpal.frx":14BC
      Left            =   5370
      List            =   "frmPpal.frx":14C9
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   315
      Width           =   1860
   End
   Begin VB.CommandButton cmdListIndex 
      Caption         =   "ListIn&dex"
      Height          =   405
      Left            =   1290
      MouseIcon       =   "frmPpal.frx":14FD
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   105
      Width           =   915
   End
   Begin VB.CommandButton cmdListCount 
      Caption         =   "&ListCount"
      Height          =   405
      Left            =   390
      MouseIcon       =   "frmPpal.frx":1807
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   105
      Width           =   915
   End
   Begin VB.ComboBox cmbStyle 
      ForeColor       =   &H00C56A31&
      Height          =   315
      ItemData        =   "frmPpal.frx":1B11
      Left            =   240
      List            =   "frmPpal.frx":1B54
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   855
      Width           =   1860
   End
   Begin MSComctlLib.ImageList imgListIcon 
      Left            =   -765
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2255
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":27EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3201
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3C13
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4625
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4BBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4F59
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5C33
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":61CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6327
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6D39
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":774F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7AE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":84FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":91D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":956F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A0A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AAB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AE4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B1E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B583
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":BB1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C52F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CAC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D4DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DEED
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E8FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EC9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F037
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F3D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F76F
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FBBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10007
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10453
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1089F
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10CEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11137
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11583
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":119CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11CEB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCombo 
      BackColor       =   &H00E9DEDB&
      Height          =   1590
      Left            =   4665
      ScaleHeight     =   1530
      ScaleWidth      =   2490
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   2550
      Begin VB.ComboBox ComboBox1 
         ForeColor       =   &H00C56A31&
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Text            =   "ComboBox1"
         Top             =   1125
         Width           =   2220
      End
      Begin VB.PictureBox XpComboBox1 
         Height          =   315
         Left            =   165
         ScaleHeight     =   255
         ScaleWidth      =   2115
         TabIndex        =   20
         Top             =   375
         Width           =   2180
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Combo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C56A31&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   17
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SComboBox Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C56A31&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   16
         Top             =   120
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdRemoveItem 
      Caption         =   "&RemoveItem"
      Height          =   375
      Left            =   3450
      MouseIcon       =   "frmPpal.frx":12249
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1200
      Width           =   1125
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   20
      Left            =   240
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   21
      Top             =   2880
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   4
      Left            =   240
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   22
      Top             =   3240
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   8
      Left            =   240
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   23
      Top             =   3615
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   12
      Left            =   240
      MousePointer    =   99  'Custom
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   24
      Top             =   3975
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   1
      Left            =   2010
      MousePointer    =   99  'Custom
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   25
      Top             =   2880
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   5
      Left            =   2010
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   26
      Top             =   3240
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      BackColor       =   &H00F5F9FA&
      Height          =   300
      Index           =   9
      Left            =   2010
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   27
      Top             =   3615
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   13
      Left            =   2010
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   28
      Top             =   3975
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   2
      Left            =   3795
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   29
      Top             =   2880
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   315
      Index           =   6
      Left            =   3795
      ScaleHeight     =   255
      ScaleWidth      =   1440
      TabIndex        =   30
      Top             =   3240
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      BackColor       =   &H00F5F9FA&
      Height          =   300
      Index           =   10
      Left            =   3795
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   31
      Top             =   3615
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   11
      Left            =   5595
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   32
      Top             =   3615
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   315
      Index           =   16
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   1440
      TabIndex        =   33
      Top             =   4350
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   14
      Left            =   3795
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   34
      Top             =   3975
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   15
      Left            =   5595
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   35
      Top             =   3975
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      BackColor       =   &H00F5F9FA&
      Height          =   300
      Index           =   3
      Left            =   5595
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   36
      Top             =   2880
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   7
      Left            =   5595
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   37
      Top             =   3240
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   0
      Left            =   2010
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   38
      Top             =   4350
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   17
      Left            =   3795
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   39
      Top             =   4350
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   18
      Left            =   5595
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   40
      Top             =   4350
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   19
      Left            =   240
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   41
      Top             =   4725
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   21
      Left            =   2010
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   42
      Top             =   4725
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   22
      Left            =   3795
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   43
      Top             =   4725
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   23
      Left            =   5595
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   44
      Top             =   4725
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   24
      Left            =   240
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   45
      Top             =   5115
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   25
      Left            =   2010
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   46
      Top             =   5115
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   26
      Left            =   3795
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   47
      Top             =   5115
      Width           =   1500
   End
   Begin VB.PictureBox XpComboBox2 
      Height          =   300
      Index           =   27
      Left            =   5595
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   48
      Top             =   5115
      Width           =   1500
   End
   Begin VB.Image img1 
      Height          =   405
      Index           =   1
      Left            =   0
      Picture         =   "frmPpal.frx":12553
      Top             =   -585
      Width           =   1155
   End
   Begin VB.Image img1 
      Height          =   405
      Index           =   0
      Left            =   1740
      Picture         =   "frmPpal.frx":13E0D
      Top             =   -585
      Width           =   1155
   End
   Begin VB.Image imgIsButton 
      Height          =   405
      Left            =   2190
      MouseIcon       =   "frmPpal.frx":156C7
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":159D1
      ToolTipText     =   "Visit this spectacular control"
      Top             =   720
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   195
      X2              =   7380
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   195
      X2              =   7380
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goto Planet-Source-Code to download and Vote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00805F4F&
      Height          =   195
      Index           =   6
      Left            =   3120
      MouseIcon       =   "frmPpal.frx":1728B
      MousePointer    =   99  'Custom
      TabIndex        =   50
      Top             =   2370
      Width           =   4095
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combo Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   195
      Index           =   5
      Left            =   255
      TabIndex        =   49
      Top             =   2565
      Width           =   1065
   End
   Begin VB.Image imgMaxListLenght 
      Height          =   240
      Left            =   4935
      MouseIcon       =   "frmPpal.frx":17595
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":1789F
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MaxListLength"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   195
      Index           =   4
      Left            =   3585
      TabIndex        =   19
      Top             =   75
      Width           =   1245
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment Text List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   195
      Index           =   3
      Left            =   5385
      TabIndex        =   18
      Top             =   75
      Width           =   1635
   End
   Begin VB.Image imgAlign 
      Height          =   240
      Left            =   6990
      MouseIcon       =   "frmPpal.frx":17C29
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":17F33
      Top             =   75
      Width           =   240
   End
   Begin VB.Image imgSetStyle 
      Height          =   240
      Left            =   1365
      MouseIcon       =   "frmPpal.frx":182BD
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":185C7
      Top             =   630
      Width           =   240
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style Combo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   195
      Index           =   0
      Left            =   255
      TabIndex        =   15
      Top             =   630
      Width           =   1065
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

          '***********************************'
          '* Copyright (C) 2004 - HACKPRO TM *'
          '*  Heriberto Mantilla Santamaría  *'
          '*        Barrancabermeja          *'
          '*            Colombia             *'
          '***********************************'
Option Explicit

 Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (Iccex As tagInitCommonControlsEx) As Boolean
 Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
 End Type
 Private Const ICC_USEREX_CLASSES = &H200

 Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
 Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
 Private m_hMod As Long

 Private Const SW_SHOWMAXIMIZED = 3
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 Private i As Integer

Private Sub cmdAbout_Click()
 frmAbout.Show 1
End Sub

Private Sub cmdAddItem_Click()
 XpComboBox1.AddItem txtAddItem.Text, &HFE0099
End Sub

Private Sub cmdDisabled_Click()
 If (cmdDisabled.Caption = "&Disabled") Then
  XpComboBox1.Enabled = False
  cmdDisabled.Caption = "&Enabled"
 Else
  XpComboBox1.Enabled = True
  cmdDisabled.Caption = "&Disabled"
 End If
End Sub

Private Sub cmdListCount_Click()
 MsgBox "ListCount: " & XpComboBox1.ListCount, vbInformation + vbOKOnly, "SComboBox"
End Sub

Private Sub cmdListIndex_Click()
 MsgBox "ListIndex: " & XpComboBox1.ListIndex, vbInformation + vbOKOnly, "SComboBox"
End Sub

Private Sub cmdRemoveItem_Click()
 XpComboBox1.RemoveItem 3
End Sub

Private Sub cmdSearchItem_Click()
 MsgBox "FindItem: " & XpComboBox1.FindItemText(txtSearchItem.Text, None), vbOKOnly + vbInformation, "SComboBox"
End Sub

Private Sub cmdTextItem_Click()
 MsgBox "ItemText: " & XpComboBox1.List(XpComboBox1.ListIndex), vbInformation + vbOKOnly, "SComboBox"
End Sub

Private Sub Form_Initialize()
 Dim Iccex As tagInitCommonControlsEx

 Iccex.lngSize = LenB(Iccex)
 Iccex.lngICC = ICC_USEREX_CLASSES
 InitCommonControlsEx Iccex
 m_hMod = LoadLibrary("shell32.dll")
End Sub

Private Sub Form_Load()
 Me.Show
 For i = 1 To 18
  ComboBox1.AddItem "Hackpro" & i
  If (i = 4) Or (i = 9) Or (i = 15) Or (i = 10) Then
   XpComboBox1.AddItem "Hackpro" & i, , imgListIcon.ListImages(i).Picture, False
  ElseIf (i = 8) Or (i = 12) Then
   XpComboBox1.AddItem "Hackpro" & i, &HFE0099, , , "Hola" & i, , , , , True
  ElseIf (i = 5) Or (i = 1) Or (i = 13) Or (i = 18) Then
   XpComboBox1.AddItem "Hackpro" & i, &HFE0099, imgListIcon.ListImages(i).Picture, , , , , imgListIcon.ListImages(41).Picture, True
  Else
   XpComboBox1.AddItem "Hackpro" & i, , imgListIcon.ListImages(i).Picture
  End If
 Next
 XpComboBox1.AddItem "Download and vote", vbRed, imgListIcon.ListImages(42).Picture, , "Developed by HACKPRO TM", , , , True, True
 Set XpComboBox1.MouseIcon = imgListIcon.ListImages(41).Picture
 XpComboBox1.MousePointer = vbCustom
 Set XpComboBox1.NormalPictureUser = imgListIcon.ListImages(39).Picture
 Set XpComboBox1.DisabledPictureUser = imgListIcon.ListImages(40).Picture
 Set XpComboBox1.FocusPictureUser = imgListIcon.ListImages(39).Picture
 Set XpComboBox1.HighLightPictureUser = imgListIcon.ListImages(39).Picture
 For i = 1 To 3
  XpComboBox2(12).AddItem "Picture 0" & i
 Next
 XpComboBox2(12).ListIndex = 2
 Set XpComboBox2(12).MouseIcon = imgListIcon.ListImages(41).Picture
 XpComboBox2(12).MousePointer = vbCustom
 Call XpComboBox2_SelectionMade(12, "Picture 02", 2)
 XpComboBox1.MaxListLength = 19
 XpComboBox1.NumberItemsToShow = 8
 'XpComboBox1.Text = XpComboBox1.List(19)
 XpComboBox1.Text = XpComboBox1.List(1)
 cmbStyle.ListIndex = XpComboBox1.AppearanceCombo - 1
 cmbAlign.ListIndex = XpComboBox1.Alignment
 txtMaxListLenght.Text = XpComboBox1.MaxListLength
 imgSetStyle_Click
 If (XpComboBox1.Enabled = True) Then cmdDisabled.Caption = "&Disabled"
End Sub

Private Sub Form_Terminate()
 FreeLibrary m_hMod
End Sub

Private Sub imgAlign_Click()
 XpComboBox1.Alignment = cmbAlign.ListIndex
End Sub

Private Sub imgIsButton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Set imgIsButton.Picture = img1(1).Picture
 Call Espera(0.5)
 '* Call the isButton from the web www.planet-source-code.com.
 Call ShellExecute(frmPpal.hWnd, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56053&lngWId=1", vbNullString, vbNullString, SW_SHOWMAXIMIZED)
 Set imgIsButton.Picture = img1(0).Picture
End Sub

Private Sub imgMaxListLenght_Click()
 Dim isValue As Long
 
 isValue = CLng(txtMaxListLenght.Text)
 If (isValue > 0) And (IsNumeric(isValue) = True) Then XpComboBox1.MaxListLength = isValue
End Sub

Private Sub imgSetStyle_Click()
 XpComboBox1.AppearanceCombo = cmbStyle.ListIndex + 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 8) Then Exit Sub
 If (IsNumeric(Chr$(KeyAscii)) = False) Then KeyAscii = 0: Beep
End Sub

Private Sub lblMessage_Click(Index As Integer)
 If (Index = 6) Then Call ShellExecute(frmPpal.hWnd, "open", "http://www.planet-source-code.com/vb/", vbNullString, vbNullString, SW_SHOWMAXIMIZED)
End Sub

Private Sub XpComboBox2_SelectionMade(Index As Integer, ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
 If (Index = 12) Then
  Select Case SelectedItem
   Case "Picture 01"
    Set XpComboBox2(12).FocusPictureUser = imgListIcon.ListImages(32).Picture
    Set XpComboBox2(12).HighLightPictureUser = imgListIcon.ListImages(29).Picture
    Set XpComboBox2(12).NormalPictureUser = imgListIcon.ListImages(31).Picture
    Set XpComboBox2(12).DisabledPictureUser = imgListIcon.ListImages(30).Picture
    XpComboBox2(12).NormalBorderColor = &HB99D7F
    XpComboBox2(12).SelectBorderColor = &HC56A31
    XpComboBox2(12).HighLightBorderColor = &HC56A31
   Case "Picture 02"
    Set XpComboBox2(12).FocusPictureUser = imgListIcon.ListImages(36).Picture
    Set XpComboBox2(12).HighLightPictureUser = imgListIcon.ListImages(34).Picture
    Set XpComboBox2(12).NormalPictureUser = imgListIcon.ListImages(35).Picture
    Set XpComboBox2(12).DisabledPictureUser = imgListIcon.ListImages(33).Picture
    XpComboBox2(12).NormalBorderColor = &H9F989F
    XpComboBox2(12).SelectBorderColor = &H406790
    XpComboBox2(12).HighLightBorderColor = &H90887F
   Case "Picture 03"
    Set XpComboBox2(12).FocusPictureUser = imgListIcon.ListImages(37).Picture
    Set XpComboBox2(12).HighLightPictureUser = imgListIcon.ListImages(37).Picture
    Set XpComboBox2(12).NormalPictureUser = imgListIcon.ListImages(37).Picture
    Set XpComboBox2(12).DisabledPictureUser = imgListIcon.ListImages(38).Picture
    XpComboBox2(12).NormalBorderColor = &H103030
    XpComboBox2(12).SelectBorderColor = &H103030
    XpComboBox2(12).HighLightBorderColor = &H103030
  End Select
 End If
End Sub

Private Sub Espera(ByVal Segundos As Single)
 Dim ComienzoSeg As Single, FinSeg As Single
 
 '* English: Wait a certain time.
 '* Español: Esperar un determinado tiempo.
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 Do While FinSeg > Timer
  DoEvents
  If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
 Loop
End Sub
