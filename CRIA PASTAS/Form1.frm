VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   900
   ClientTop       =   1485
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11625
   Begin VB.ListBox List7 
      Height          =   4350
      Left            =   9840
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox List6 
      Height          =   4350
      Left            =   8160
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox List5 
      Height          =   4350
      Left            =   6480
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox List4 
      Height          =   4350
      Left            =   4920
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   4350
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   4350
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmd_ler 
      Caption         =   "Ler"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   6000
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "REGISTRO"
      Height          =   255
      Left            =   9840
      TabIndex        =   14
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "LOCAL"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "STATUS"
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "COD_STATUS"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "HORA"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "DATA"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "OBSERVAÇÕES"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ler_Click()
Dim xlin As Integer
Dim xlinha As String

Open "C:\SAIDA.TXT" For Input As #1

    Do Until EOF(1)
    xlin = xlin + 1
    Line Input #1, xlinha
    
   ' If Mid(xlinha, 39, 2) = "40" Then
    List1.AddItem Mid(xlinha, 1, 20)
    Label8.Caption = List1.ListCount
    List2.AddItem Mid(xlinha, 21, 10)
    List3.AddItem Mid(xlinha, 31, 8)
    List4.AddItem Mid(xlinha, 39, 2)
    List5.AddItem Mid(xlinha, 41, 20)
    List6.AddItem Mid(xlinha, 61, 130)
    List7.AddItem Mid(xlinha, 191, 13)
   ' End If
   
    
    Loop
    

End Sub
