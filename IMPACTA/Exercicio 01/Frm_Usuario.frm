VERSION 5.00
Begin VB.Form Frm_Usuario 
   BackColor       =   &H00C0E0FF&
   Caption         =   "EXERCICIO 01"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2310
   Icon            =   "Frm_Usuario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Sair 
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Mensagem 
      Caption         =   "MENSAGEM"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Usuario 
      BackColor       =   &H00C0C0C0&
      Caption         =   "USUÁRIO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Mensagem_Click()
    MsgBox "Treina VBEssentials", vbInformation, "AVISO"
    'Msgbox Exercicio 01


End Sub

Private Sub Cmd_Sair_Click()
    If MsgBox("Deseja Sair?", vbQuestion + vbOKCancel, "AVISO") = vbOK Then
    Unload Frm_Usuario
    End If
    
    
End Sub

Private Sub Cmd_Usuario_Click()
    Frm_Usuario.Caption = InputBox("Digite seu Nome", "Usuário", "Digite Aqui")

End Sub
