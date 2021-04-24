VERSION 5.00
Begin VB.Form frm_decimal 
   Caption         =   "DECIMAL PARA BINARIO"
   ClientHeight    =   6150
   ClientLeft      =   3615
   ClientTop       =   2520
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8280
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frm_decimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xresto As Integer
Dim xquociente As Integer
Dim xresultado As String
Dim xdividendo As Integer
Dim x As Integer

'For x = 0 To Int(Val(Text1.Text))

xdividendo = Int(Val(Text1.Text))

'xdividendo = x

Do Until xdividendo = 0

    xresto = xdividendo Mod 2
    xdividendo = Int(xdividendo / 2)
    xresultado = xresto & xresultado


Loop


If Len(Trim$(xresultado)) < 8 Then

    List1.AddItem String(8 - Len(Trim$(xresultado)), "0") & xresultado

Else
    
    List1.AddItem xresultado

End If

'xresultado = Empty

'Next


End Sub
