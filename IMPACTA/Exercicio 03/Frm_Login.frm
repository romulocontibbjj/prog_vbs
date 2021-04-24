VERSION 5.00
Begin VB.Form Frm_Login 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Frm_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_OK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      MousePointer    =   99  'Custom
      Picture         =   "Frm_Login.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Sair 
      Caption         =   "&SAIR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Picture         =   "Frm_Login.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txt_Senha 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txt_usuario 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Senha.......:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Usuário....:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_OK_Click()
    'USUARIOS E SENHAS IGUAIS
    If UCase(txt_usuario.Text) = "ROMULO" And txt_Senha.Text = "5754" Then
        frm_Misturador.Show
        Unload Me
Else
      'SENHA ERRADA
      If UCase(txt_usuario.Text) = "ROMULO" And txt_Senha.Text <> "5754" Then
            MsgBox "Senha Inválida", vbExclamation, "Erro de Senha"
            txt_Senha.Text = Empty
            txt_Senha.SetFocus
        
        End If
    
    'USUARIO INCORRETO
    If UCase(txt_usuario.Text) <> "ROMULO" And txt_Senha = "5754" Then
        MsgBox "Usuário Inválido", vbExclamation, "Erro de Usuário"
        txt_usuario.SelStart = 0
        txt_usuario.SelLength = Len(txt_usuario.Text)
        txt_usuario.SetFocus
        
        End If
        
     'USUARIO E SENHA INCORRETOS
    If UCase(txt_usuario.Text) <> "ROMULO" And txt_Senha.Text <> "5754" Then
        MsgBox "Usuário e Senha Inválidos", vbInformation, "ERRO"
    
    End If
        
    
    End If
        

       
End Sub

Private Sub cmd_Sair_Click()
If MsgBox("Você Deseja Sair???", vbYesNo + vbQuestion, "SAIR?") = vbYes Then
    Unload Me
Else
txt_usuario.Text = Empty
txt_Senha.Text = Empty
txt_usuario.SetFocus

End If

End Sub

Private Sub txt_Senha_Change()

If Trim(txt_usuario.Text) = Empty Or Trim(txt_Senha.Text) = Empty Then
    cmd_OK.Enabled = False
Else
    cmd_OK.Enabled = True
End If

End Sub
Private Sub txt_usuario_Change()

If Trim(txt_usuario.Text) = Empty Or Trim(txt_Senha.Text) = Empty Then
    cmd_OK.Enabled = False
Else
    cmd_OK.Enabled = True
End If

End Sub
