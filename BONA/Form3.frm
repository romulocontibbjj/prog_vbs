VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8235
   ClientLeft      =   2895
   ClientTop       =   1440
   ClientWidth     =   8130
   LinkTopic       =   "Form3"
   ScaleHeight     =   8235
   ScaleWidth      =   8130
   Begin MSChart20Lib.MSChart graf 
      Height          =   2055
      Left            =   2520
      OleObjectBlob   =   "Form3.frx":0000
      TabIndex        =   2
      Top             =   3360
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   1920
      TabIndex        =   3
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _Version        =   524288
      _ExtentX        =   5530
      _ExtentY        =   4683
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2005
      Month           =   6
      Day             =   15
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()

MsgBox Calendar1.Day & "/" & Calendar1.Month

End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Command2_Click()



graf.Column = 1
graf.RowLabel = "ROM"
'MsgBox graf.RowLabel
graf.Data = 80 - 70
graf.Column = 2
graf.RowLabel = "ROM2"


End Sub

