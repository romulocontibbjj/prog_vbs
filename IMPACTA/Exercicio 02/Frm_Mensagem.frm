VERSION 5.00
Begin VB.Form Frm_Mensagem 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exercicio 02"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "Frm_Mensagem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5040
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tmr_Hora 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   3240
   End
   Begin VB.CommandButton Cmd_hora 
      Caption         =   "Hora"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Escreve 
      Caption         =   "Es&creve"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_transfere 
      Caption         =   "Transfere"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "Sai&r"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmd_limpar 
      Caption         =   "&Limpar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_mensagem 
      Caption         =   "&Mensagem"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lab_hora 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lab_Mensagem 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Lab_Escreve 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "Frm_Mensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Escreve_Click()
    Lab_Escreve.Visible = True
    Cmd_transfere.Visible = True
    Lab_Escreve.Caption = "TREINAMENTO DE VB ESSENTIALS"
    cmd_limpar.Enabled = True
    
End Sub

Private Sub Cmd_hora_Click()
    lab_hora.Caption = Time
    Tmr_Hora.Enabled = True

End Sub

Private Sub cmd_limpar_Click()
    Lab_Escreve.Visible = False
    lab_Mensagem.Visible = False
    Cmd_Escreve.Visible = False
    Cmd_transfere.Visible = False
    Cmd_hora.Visible = False
    Tmr_Hora.Enabled = False
    lab_hora.Caption = Empty
    Cmd_mensagem.SetFocus
    cmd_limpar.Enabled = False
    
End Sub

Private Sub Cmd_mensagem_Click()
    lab_Mensagem.Visible = True
    'Mostrar Campo Mensagem
    Cmd_Escreve.Visible = True
    'Mostrar CMD Escreve
    lab_Mensagem.Caption = "ROMULO"
    'Mostra Mensagem

End Sub

Private Sub cmd_sair_Click()
    If MsgBox("Deseja sair????", vbQuestion + vbOKCancel, "SAIR") = vbOK Then
        Unload Frm_Mensagem
    End If
        
End Sub

Private Sub Cmd_transfere_Click()
    Dim vnome As String
    vnome = Lab_Escreve.Caption
    Lab_Escreve = lab_Mensagem.Caption
    lab_Mensagem = vnome
    Cmd_hora.Visible = True
    
    

End Sub

Private Sub Tmr_Hora_Timer()
    lab_hora.Caption = Time
    
End Sub
