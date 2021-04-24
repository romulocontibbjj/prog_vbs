VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   3360
   ClientTop       =   2160
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Excel As Excel.Application
Dim ExcelA1 As Excel.Worksheet
Dim ExcelWb As Excel.Workbook

Set Excel = CreateObject("Excel.Application")
Set Excel = GetObject(, "Excel.Application")

Excel.Visible = True

Set ExcelWb = Excel.Workbooks.Add
Set ExcelA1 = Excel.Worksheets(1)

Excel.Cells(1, 1) = "ROMULO"


Excel.Quit

ExcelWb = Nothing
ExcelA1 = Nothing





End Sub
