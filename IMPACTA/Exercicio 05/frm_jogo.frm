VERSION 5.00
Begin VB.Form frm_jogo 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caixa Forte - TIO PATINHAS"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_jogo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_start 
      Caption         =   "Star&t"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmd_stop 
      Caption         =   "&Pare"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.Timer tmr_sorteio 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   1680
   End
   Begin VB.CommandButton cmd_iniciar 
      Caption         =   "&Inicio"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "Sai&r"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmd_Sorteio 
      Caption         =   "&Sorteio"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image img_Sorteio 
      Height          =   1380
      Left            =   2040
      Picture         =   "frm_jogo.frx":0442
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label lab3 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lab2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lab1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frm_jogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_iniciar_Click()
lab1.Caption = 0
lab2.Caption = 0
lab3.Caption = 0
img_Sorteio.Visible = False
cmd_start.Enabled = True
cmd_Sorteio.Enabled = True
cmd_Sorteio.SetFocus

End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub cmd_Sorteio_Click()
Randomize
    lab1.Caption = Int(Rnd * 10)
    lab2.Caption = Int(Rnd * 10)
    lab3.Caption = Int(Rnd * 10)
    If lab1.Caption Mod 2 = 0 And lab2.Caption Mod 2 = 0 _
        And lab3.Caption Mod 2 = 0 Then
        cmd_start.Enabled = False
        img_Sorteio.Visible = True
        cmd_Sorteio.Enabled = False
        cmd_iniciar.SetFocus
        MsgBox "O Número Vencedor foi: " & lab1 & lab2 & lab3, vbInformation, "Campeão"
        
    End If
    

End Sub

Private Sub cmd_start_Click()
tmr_sorteio.Enabled = True

End Sub

Private Sub cmd_stop_Click()
tmr_sorteio.Enabled = False
 If lab1.Caption Mod 2 = 0 And lab2.Caption Mod 2 = 0 _
        And lab3.Caption Mod 2 = 0 Then
        img_Sorteio.Visible = True
        cmd_Sorteio.Enabled = False
        cmd_iniciar.SetFocus
        MsgBox "O Número Vencedor foi: " & lab1 & lab2 & lab3, vbInformation, "Campeão"
        cmd_start.Enabled = False
        
    End If

End Sub

Private Sub Form_Load()
lab1.Caption = 0
lab2.Caption = 0
lab3.Caption = 0
End Sub

Private Sub tmr_sorteio_Timer()
Randomize
    lab1.Caption = Int(Rnd * 10)
    lab2.Caption = Int(Rnd * 10)
    lab3.Caption = Int(Rnd * 10)

    
   End Sub
