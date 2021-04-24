VERSION 5.00
Begin VB.Form frmAcesso 
   BackColor       =   &H00800000&
   Caption         =   "Informa - Acesso"
   ClientHeight    =   3165
   ClientLeft      =   2685
   ClientTop       =   1920
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdConfigConexao 
      BackColor       =   &H00808080&
      Caption         =   "..."
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   315
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   330
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   915
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2520
      Width           =   1275
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   10
      TabIndex        =   0
      Top             =   2520
      Width           =   1275
   End
   Begin VB.CommandButton cmdConfSenha 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirma..."
      Height          =   330
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Controle de Acesso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Módulo de Faturamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   150
      TabIndex        =   12
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema Informa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1290
      TabIndex        =   11
      Top             =   240
      Width           =   3285
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
      Left            =   5280
      TabIndex        =   4
      Top             =   2520
      Width           =   345
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   615
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "de  3"
      Height          =   195
      Left            =   2355
      TabIndex        =   7
      Top             =   4980
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lbltenta 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   2145
      TabIndex        =   8
      Top             =   4980
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tentativas: "
      Height          =   195
      Left            =   1200
      TabIndex        =   9
      Top             =   4980
      Visible         =   0   'False
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
    Unload mdiFatura
    Unload Me
End Sub

Private Sub cmdConfigConexao_Click()
    frmConexao.Show 1
End Sub

Private Sub cmdConfSenha_Click()
    If txtUsuario <> txtSenha Then
        MsgBox "ERRO ! Senhas Diferentes. Verifique Novamente."
        txtUsuario.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtSenha)) < 6 Then
        MsgBox "A Senha deve ter no mínimo 6 caracteres !"
        txtUsuario.SetFocus
        Exit Sub
    End If
    If de_informa.rsSel_Usuario.Fields("senha") = txtSenha Then
        MsgBox "Senha Inválida ! Escolha Outra."
        txtUsuario.SetFocus
        Exit Sub
    End If
    de_informa.alt_senha txtSenha, de_informa.rsSel_Usuario("usuario")
    MsgBox "Senha Alterada com Sucesso ! Entre Novamente no Sistema."
    
    Unload mdiFatura
    Unload Me
End Sub
Private Sub cmdOk_Click()
    Dim xlinha As String

    Me.MousePointer = 11
    
    xstrcon = de_informa.cn_informa.ConnectionString
    xstrcon2 = "driver={SQL Server};server=192.9.205.3;database=intec;user id=sa;password=zx11bbb7"
    
    
    If Dir("C:\informa.cnx") <> "" Then
    
        Open "C:\informa.cnx" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "CNX" Then
                xCnx = Trim$(Mid$(xlinha, 5))
            ElseIf Mid$(xlinha, 1, 3) = "BCO" Then
                xBco = Trim$(Mid$(xlinha, 5))
            End If
        Loop
        
        If Len(Trim$(xCnx)) > 0 And Len(Trim$(xBco)) = 0 Then
            xBco = "INTEC"
        End If
        
        Close #1

        xstrcon = "Provider=SQLOLEDB.1;Password=zx11bbb7;Persist Security Info=True;User ID=sa;Initial Catalog=" & xBco & ";Data Source=" & xCnx
        xstrcon2 = "driver={SQL Server};server=" & xCnx & ";database=" & xBco & ";user id=sa;password=zx11bbb7"
        
        If de_informa.cn_informa.State = 1 Then
            de_informa.cn_informa.Close
            de_informa.cn_informa.ConnectionString = xstrcon
            de_informa.cn_informa.Open xstrcon
        Else
            de_informa.cn_informa.ConnectionString = xstrcon
        End If
        
    End If
    
    If de_informa.rsSel_Usuario.State = 1 Then de_informa.rsSel_Usuario.Close
    de_informa.Sel_Usuario txtUsuario.Text   'PROCURA USUÁRIO
    
        If de_informa.rsSel_Usuario.RecordCount > 0 Then
            
            If txtSenha.Text <> de_informa.rsSel_Usuario.Fields("senha") Then
                
                Me.MousePointer = 0
                MsgBox "Senha Incorreta !", vbCritical + vbExclamation, "Erro"
                txtSenha.SetFocus
                lbltenta = lbltenta + 1
                
                If lbltenta = 4 Then
                    MsgBox "Suas tentativas de acesso expirou, Consulte o Administrador do Sistema !", vbCritical + vbExclamation, "Acesso Negado"
                    Unload mdiFatura
                    Unload Me
                End If
                Exit Sub
            Else
                
                xdireitos = de_informa.rsSel_Usuario.Fields("stringdireitos")
                
                If de_informa.rsSel_Usuario.Fields("expirada") = "S" Then
                    
                    Me.MousePointer = 0
                    
                    MsgBox "Atenção. Sua SENHA expirou ! Cadastre uma Nova Senha."
                    
                    Label1 = "NOVA SENHA:"
                    Label2 = "CONFIRME"
                    txtUsuario = ""
                    txtSenha = ""
                    txtUsuario.PasswordChar = "*"
                    
                    cmdOk.Visible = False
                    cmdConfSenha.Visible = True
                
                ElseIf de_informa.rsSel_Usuario.Fields("status") = "0" Then
                    
                    Me.MousePointer = 0
                    
                    MsgBox "Usuário com Acesso BLOQUEADO ao Sistema !"
                    
                    Unload mdiFatura
                    Unload Me
                
                
                'Franklin Inicio
                ElseIf Mid$(xdireitos, 86, 1) = "0" Then
                    Me.MousePointer = 0
                    MsgBox "Acesso não permitido, consulte seu Coordenador ou Gerente !", vbCritical + vbExclamation, "Acesso Não Permitido"
                    Unload mdiFatura
                    Unload Me
                'Franklin Fim
                Else
                    Me.MousePointer = 0
                    xusuario = de_informa.rsSel_Usuario.Fields("usuario")
                    xdireitos = de_informa.rsSel_Usuario.Fields("stringdireitos")
                    Unload Me
            End If
        End If
        
    Else
        Me.MousePointer = 0
        MsgBox "Usuário Não Cadastrado !", vbCritical + vbExclamation, "Erro"
        txtUsuario.SetFocus
        lbltenta = lbltenta + 1
        If lbltenta = 4 Then  'SOMENTE 3 TENTATIVAS
            MsgBox "Suas tentativas de acesso expirou, Consulte o Administrador do Sistema !", vbCritical + vbExclamation, "Acesso Negado"
            Unload mdiFatura
            Unload Me
        End If
        Exit Sub
    End If
End Sub
Private Sub Command1_Click()

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
    Unload mdiFatura
    Unload Me
    End
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
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = 10
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUsuario_LostFocus()
    txtUsuario.Text = UCase(txtUsuario.Text)
End Sub



