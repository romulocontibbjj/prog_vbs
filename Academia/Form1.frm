VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   4080
   ClientTop       =   2505
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7560
   Begin VB.CommandButton cmd_Limpa 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   2160
      TabIndex        =   22
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmd_Seleciona 
      Caption         =   "Seleciona"
      Height          =   495
      Left            =   2160
      TabIndex        =   21
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   4320
      TabIndex        =   20
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Text            =   "ABDOMINAL"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Text            =   "BIKE"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   480
      TabIndex        =   15
      Text            =   "CORRIDA"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Text            =   "PERNA"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Text            =   "OMBRO"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   12
      Text            =   "ANTEBRAÇO"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Text            =   "BICEPS"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   10
      Text            =   "TRICEPS"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Text            =   "COSTAS"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Text            =   "PEITO"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "10-"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "9-"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "8-"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "7-"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "6-"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "3-"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "2-"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "4-"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "5-"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1-"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Seleciona_Click()
Dim x As Integer
Dim xexercicio As String
Dim y As Integer
Dim z As Integer
Dim w As Integer


For x = 1 To 10

Randomize

z = Int(Rnd * 10)

Select Case z

Case 1:
xexercicio = Text1.Text
Case 2:
xexercicio = Text2.Text
Case 3:
xexercicio = Text3.Text
Case 4:
xexercicio = Text4.Text
Case 5:
xexercicio = Text5.Text
Case 6:
xexercicio = Text6.Text
Case 7:
xexercicio = Text7.Text
Case 8:
xexercicio = Text8.Text
Case 9:
xexercicio = Text9.Text
Case 10:
xexercicio = Text10.Text
End Select

'If List1.ListCount = 0 Then

'    List1.AddItem xexercicio

'End If




y = List1.ListCount

Do While y <> -1
If Trim(xexercicio) = Trim$(List1.List(x)) Then
cmd_Seleciona_Click
Else
y = y - 1
End If
Loop




w = w + 1

If w <> 10 Then
    List1.AddItem xexercicio
Else
    Exit Sub
End If

Next

End Sub
