VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   2970
   ClientTop       =   1530
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   9300
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFLEX 
      Height          =   1455
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
      _Version        =   393216
      BackColorFixed  =   16777152
      BackColorBkg    =   12648447
      BackColorUnpopulated=   12632319
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   4080
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

MSFLEX.Rows = 1
DoEvents
MSFLEX.Rows = 2
MSFLEX.FixedRows = 1

MSFLEX.TextMatrix(MSFLEX.Row, 1) = "rrc"

MSFLEX.Rows = MSFLEX.Rows + 1

MSFLEX.TextMatrix(MSFLEX.Rows - 1, 1) = "TESTE"

MSFLEX.Rows = MSFLEX.Rows + 1

MSFLEX.Cols = MSFLEX.Cols + 1

MSFLEX.TextMatrix(0, 1) = "TR"

MSFLEX.TextMatrix(1, 0) = "TR2"



MSFLEX.TextMatrix(1, 2) = "TESTE"

MSFLEX.TextMatrix(1, 2) = "TESTE"

Label1.Caption = MSFLEX.Rows - 1 & " / " & MSFLEX.Cols - 1





End Sub

Private Sub Command2_Click()
Dim X As Integer

For X = 1 To 7

    MSFLEX.Rows = MSFLEX.Rows + X
    MSFLEX.Cols = MSFLEX.Cols + X
Next






End Sub

