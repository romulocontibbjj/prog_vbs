VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Font"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Foreground Color"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   960
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Background Color"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tool Tip Text"
      Height          =   3495
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   6495
      Begin VB.OptionButton Option6 
         Caption         =   "Error Icon"
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Warning Icon"
         Height          =   255
         Left            =   4920
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Info Icon"
         Height          =   255
         Left            =   4920
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "No Icon"
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Update Title"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Update Text"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Top             =   840
         Width           =   4695
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Hide Tool Tip"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Graphics Type Of Tool Tip"
      Height          =   1575
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox Check4 
         Caption         =   "Always Tip"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2895
      End
      Begin VB.CheckBox Check3 
         Caption         =   "No Animate"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox Check2 
         Caption         =   "No Fade"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ballon"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ex Type Of Tool Tip"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "Stand Always [Extended] with events Click,DblClick,Mouse Move,Mouse Leave (Require Show Tool Tip After Create!!!!!!!!)"
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Simple"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Destroy ToolTip"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Move At Position"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7800
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Tool Tip"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create ToolTip"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Event:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   8535
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   6840
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Y"
      Height          =   255
      Left            =   7800
      TabIndex        =   27
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Left            =   6960
      TabIndex        =   26
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents TT As ToolTipEx
Attribute TT.VB_VarHelpID = -1
Private Sub Command1_Click()
Dim PT As ToolTipExStyle

If Option1.Value = True Then
PT = Simple
ElseIf Option2.Value = True Then
PT = StandAlways
End If

Dim TS As TTS_Styles
If Check1.Value = 1 Then TS = TS Or TTS_BALLOON
If Check2.Value = 1 Then TS = TS Or TTS_NOFADE
If Check3.Value = 1 Then TS = TS Or TTS_NOANIMATE
If Check4.Value = 1 Then TS = TS Or TTS_ALWAYSTIP




TT.CreateToolTip Form1.hWnd, TS, "Author:Vanja Fuckar, EMAIL:INGA@VIP.HR" & vbCrLf & "BTW:Dont forget check that memory location!", "Warning:Error At Address 'Ch034' Offset '13h'", 3, PT
TT.SetToolTipFont "Courier New", 7, 238, False, False, False, False
End Sub

Private Sub Command10_Click()
cd1.Flags = FontsConstants.cdlCFEffects Or FontsConstants.cdlCFScreenFonts
cd1.ShowFont
TT.SetToolTipFont cd1.FontName, cd1.FontSize, 0, cd1.FontItalic, cd1.FontUnderline, cd1.FontStrikethru, cd1.FontBold

End Sub

Private Sub Command2_Click()
TT.ShowToolTip Me.hWnd, True
End Sub

Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Then Exit Sub
If Not IsNumeric(Text1) Or Not IsNumeric(Text2) Then Exit Sub
TT.AtPosition Text1, Text2
End Sub



Private Sub Command4_Click()
TT.DestroyToolTip
End Sub

Private Sub Command5_Click()
TT.Text = Text3
End Sub



Private Sub Command6_Click()
TT.ShowToolTip Me.hWnd, False
End Sub

Private Sub Command7_Click()
cd1.ShowColor
TT.BackgroundColor = cd1.Color

End Sub

Private Sub Command8_Click()
cd1.ShowColor
TT.ForegroundColor = cd1.Color

End Sub

Private Sub Command9_Click()
Dim L1 As Long
If Option3.Value = True Then
L1 = 0
ElseIf Option4.Value = True Then
L1 = 1
ElseIf Option5.Value = True Then
L1 = 2
ElseIf Option6.Value = True Then
L1 = 3
End If

TT.Title(L1) = Text3
End Sub

Private Sub Form_Load()
Option2.Value = True
Option3.Value = True
Set TT = New ToolTipEx
End Sub


Private Sub TT_BeginShow()
Label4 = "Event: Begin Show"
End Sub

Private Sub TT_Click()
Label4 = "Event: Click"
End Sub

Private Sub TT_DblClick()
Label4 = "Event: Double Click"
End Sub

Private Sub TT_MouseLeave()
Label4 = "Event: Mouse Leave"
End Sub

Private Sub TT_MouseMove(ByVal X As Long, ByVal Y As Long)
Label4 = "Event: Mouse Move X:" & X & "," & "Y:" & Y
End Sub

Private Sub TT_RightClick()
Label4 = "Event: Right Click"
End Sub

Private Sub TT_RightDblClick()
Label4 = "Event: Right Double Click"
End Sub
