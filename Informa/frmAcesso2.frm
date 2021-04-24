VERSION 5.00
Begin VB.Form frmAcesso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Acesso"
   ClientHeight    =   3885
   ClientLeft      =   3030
   ClientTop       =   1545
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "frmAcesso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4305
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1575
      Width           =   1275
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4305
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2100
      Width           =   1275
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   330
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   435
   End
   Begin VB.CommandButton cmdConfSenha 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirma..."
      Height          =   330
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblSair 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   195
      Left            =   5145
      TabIndex        =   8
      Top             =   3465
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4305
      TabIndex        =   7
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4305
      TabIndex        =   6
      Top             =   1890
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   3885
      Left            =   -1800
      Picture         =   "frmAcesso.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "de  3"
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   2100
      Width           =   360
   End
   Begin VB.Label lbltenta 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   1470
      TabIndex        =   4
      Top             =   2100
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tentativas: "
      Height          =   195
      Left            =   525
      TabIndex        =   2
      Top             =   2100
      Width           =   840
   End
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload mdiInforma
    Unload Me
End Sub

Private Sub cmdConfSenha_Click()
    If TxtUsuario <> txtSenha Then
        MsgBox "ERRO ! Senhas Diferentes. Verifique Novamente."
        TxtUsuario.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtSenha)) < 6 Then
        MsgBox "A Senha deve ter no mínimo 6 caracteres !"
        TxtUsuario.SetFocus
        Exit Sub
    End If
    If de_informa.rsSel_Usuario.Fields("senha") = txtSenha Then
        MsgBox "Senha Inválida ! Escolha Outra."
        TxtUsuario.SetFocus
        Exit Sub
    End If
    de_informa.alt_senha txtSenha, de_informa.rsSel_Usuario("usuario")
    MsgBox "Senha Alterada com Sucesso ! Entre Novamente no Sistema."
    
    Unload mdiInforma
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Me.MousePointer = 11
    If de_informa.rsSel_Usuario.State = 1 Then de_informa.rsSel_Usuario.Close
    de_informa.Sel_Usuario TxtUsuario.Text   'PROCURA USUÁRIO
    If de_informa.rsSel_Usuario.RecordCount > 0 Then
        If txtSenha.Text <> de_informa.rsSel_Usuario.Fields("senha") Then
            Me.MousePointer = 0
            MsgBox "Senha Incorreta !", vbCritical + vbExclamation, "Erro"
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "LOGIN", TxtUsuario.Text, "FALHA: SENHA INCORRETA"
            
            txtSenha.SetFocus
            lbltenta = lbltenta + 1
            If lbltenta = 4 Then
                Unload mdiInforma
                Unload Me
            End If
            Exit Sub
        Else
            If de_informa.rsSel_Usuario.Fields("expirada") = "S" Then
                Me.MousePointer = 0
                MsgBox "Atenção. Sua SENHA expirou ! Cadastre uma Nova Senha."
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "LOGIN", TxtUsuario.Text, "FALHA: SENHA EXPIRADA"
             
                Label1 = "NOVA SENHA:"
                Label2 = "CONFIRME"
                TxtUsuario = ""
                txtSenha = ""
                TxtUsuario.PasswordChar = "*"
                CmdOk.Visible = False
                cmdConfSenha.Visible = True
            ElseIf de_informa.rsSel_Usuario.Fields("status") = "0" Then
                Me.MousePointer = 0
                MsgBox "Usuário com Acesso BLOQUEADO ao Sistema !"
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "LOGIN", TxtUsuario.Text, "FALHA: ACESSO BLOQUEADO"
                
                Unload mdiInforma
                Unload Me
            Else
                Me.MousePointer = 0
                xusuario = de_informa.rsSel_Usuario.Fields("usuario")
                xdireitos = de_informa.rsSel_Usuario.Fields("stringdireitos")
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "LOGIN", TxtUsuario.Text, "OK"
                
                If Mid$(de_informa.rsSel_Usuario.Fields("stringdireitos"), 24, 1) = "1" Then
                    mdiInforma.tmAlarmeUrgencia.Interval = 15000
                    xtempoalarme = 0
                Else
                    mdiInforma.tmAlarmeUrgencia.Interval = 0
                End If
                
                Unload Me
            End If
        End If
    Else
        Me.MousePointer = 0
        MsgBox "Usuário Não Cadastrado !", vbCritical + vbExclamation, "Erro"
        TxtUsuario.SetFocus
        lbltenta = lbltenta + 1
        If lbltenta = 4 Then  'SOMENTE 3 TENTATIVAS
            Unload mdiInforma
            Unload Me
        End If
        Exit Sub
    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set frmAcesso = Nothing
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSair.Font.Size = 8
    lblSair.ForeColor = &H8000000A
    DoEvents
End Sub

Private Sub lblSair_Click()
    Unload mdiInforma
    Unload Me
End Sub

Private Sub lblSair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSair.ForeColor = &HFFFF&
    lblSair.Font.Size = 10
    DoEvents
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = 10
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtSenha_LostFocus()
    txtSenha.Text = UCase(txtSenha.Text)
End Sub

Private Sub txtUsuario_GotFocus()
    TxtUsuario.SelStart = 0
    TxtUsuario.SelLength = 10
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUsuario_LostFocus()
    TxtUsuario.Text = UCase(TxtUsuario.Text)
End Sub



