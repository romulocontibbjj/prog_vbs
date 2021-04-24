VERSION 5.00
Begin VB.Form frm_excel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   3120
   ClientTop       =   1650
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.ListBox List2 
         Height          =   3765
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Registro Atual:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lab_atual 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Total de Registros:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lab_max 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
End
Attribute VB_Name = "frm_excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Excel As Excel.Application
Dim ExcelA1 As Excel.Worksheet
Dim xlinhamax As Integer, xlinhaatual As Integer

Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
Excel.Interactive = False
Excel.Workbooks.Open FileName:="C:\EMAIL.XLS"
Set ExcelA1 = Excel.Worksheets(1)

xlinhaatual = 0

Do While True

    xlinhaatual = xlinhaatual + 1
    xlinhamax = xlinhamax + 1
    
    If Len(Trim$(ExcelA1.Cells(xlinhaatual, 1))) = 0 And Len(Trim$(ExcelA1.Cells(xlinhaatual, 2))) = 0 And xlinhaatual >= 2 Then
    
        Exit Do
    End If
    
    lab_max.Caption = xlinhamax - 1
    Me.Refresh
Loop
    


Label1.Caption = UCase(ExcelA1.Cells(1, 1))
Label2.Caption = UCase(ExcelA1.Cells(1, 2))

For xlinhaatual = 1 To xlinhamax - 1

    List1.AddItem UCase(ExcelA1.Cells(xlinhaatual + 1, 1))
    List2.AddItem LCase(ExcelA1.Cells(xlinhaatual + 1, 2))
    
Next

MsgBox "OK"


End Sub
