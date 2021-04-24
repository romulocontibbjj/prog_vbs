VERSION 5.00
Begin VB.Form frm_Binario 
   Caption         =   "BINÁRIO PARA DECIMAL"
   ClientHeight    =   3270
   ClientLeft      =   1395
   ClientTop       =   1755
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   3720
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frm_Binario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As Integer
Dim xtotal As Double
Dim xvalor As String
Dim xaux As Double
Dim xcont As Integer

xvalor = Int(Val(Text1.Text))

xcont = Len(Trim$(xvalor))

For x = 1 To Len(Trim$(xvalor))
    
    xaux = Int(Val(Mid(xvalor, xcont, 1)))
    xaux = xaux * 2 ^ (x - 1)
    xtotal = xtotal + xaux
    xcont = xcont - 1

Next

Label1.Caption = xtotal


End Sub

Private Sub Text1_LostFocus()

Dim x As Integer
Dim xtotal As Double
Dim xvalor As String
Dim xaux As Double
Dim xcont As Integer

xvalor = Int(Val(Text1.Text))

xcont = Len(Trim$(xvalor))

If xcont < 8 Then
    
    xvalor = String(8 - xcont, "0") & xvalor
End If

Text1.Text = xvalor


End Sub
