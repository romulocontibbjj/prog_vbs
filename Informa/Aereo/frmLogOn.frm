VERSION 5.00
Begin VB.Form frmLogOn 
   Caption         =   "Informa Aéreo"
   ClientHeight    =   3420
   ClientLeft      =   4620
   ClientTop       =   2415
   ClientWidth     =   2685
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogOn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   2685
   Begin VB.CommandButton CmdProsseguir 
      Caption         =   "Prosseguir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   375
      TabIndex        =   7
      Top             =   1860
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox ComboCON 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   375
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1380
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   525
      TabIndex        =   5
      Top             =   2580
      Width           =   1635
   End
   Begin VB.CommandButton cmdEntrar 
      BackColor       =   &H00404080&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   525
      MaskColor       =   &H0000C0C0&
      TabIndex        =   4
      Top             =   1875
      Width           =   1635
   End
   Begin VB.TextBox txtPassWord 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   525
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1245
      Width           =   1635
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   525
      TabIndex        =   2
      Top             =   600
      Width           =   1635
   End
   Begin VB.Label LblLogOn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Escola a Forma de Log On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label LblSenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   525
      TabIndex        =   1
      Top             =   1020
      Width           =   555
   End
   Begin VB.Label LblUsuario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   525
      TabIndex        =   0
      Top             =   360
      Width           =   660
   End
   Begin VB.Image Image1 
      Height          =   3435
      Left            =   0
      Picture         =   "frmLogOn.frx":49E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xTry As Integer


Private Sub cmdEntrar_Click()
    If Len(Trim(txtUser.Text)) = 0 Then
    MsgBox "Usuário não Existente! Tente novamente.", vbCritical, "ERRO!"
    xTry = xTry + 1
        If xTry >= 3 Then
        cmdSair_Click
        End If
    Exit Sub
    ElseIf Len(Trim(txtPassWord.Text)) = 0 Then
    MsgBox "Senha incorreta! Tente novamente.", vbCritical, "ERRO!"
    xTry = xTry + 1
        If xTry >= 3 Then
        cmdSair_Click
        End If
    Exit Sub
    End If

Me.MousePointer = 11
If de_informa.rsConfUser.State = 1 Then de_informa.rsConfUser.Close
de_informa.ConfUser Trim(txtUser.Text)

    If de_informa.rsConfUser.RecordCount = 0 Then
    MsgBox "Usuário não Existente! Tente novamente.", vbCritical, "ERRO!"
    xTry = xTry + 1
        If xTry >= 3 Then
        cmdSair_Click
        End If
    Me.MousePointer = 0
    Exit Sub
    Else
        If de_informa.rsConfUser.Fields("senha") <> UCase(Trim(txtPassWord.Text)) Then
        MsgBox "Senha incorreta! Tente novamente.", vbCritical, "ERRO!"
        xTry = xTry + 1
            If xTry >= 3 Then
            cmdSair_Click
            End If
        Me.MousePointer = 0
        Exit Sub
        End If
    End If
    
    'If Mid(de_informa.rsConfUser.Fields("stringdireitos"), 32, 1) <> "1" Then
    'MsgBox "Você não tem permissão para utilizar o AW Informa! Contate o administrador do sistema.", vbCritical, "ACESSO NEGADO!"
    'cmdSair_Click
    'Exit Sub
    'End If
    
    
xUsuario = Trim(txtUser.Text)
StringDireitos = de_informa.rsConfUser.Fields("stringdireitos")
Me.MousePointer = 0
Unload Me
mdiAereo.Show
End Sub

Private Sub CmdProsseguir_Click()
xTry = 0



    If Len(Trim(ComboCON.Text)) = 0 Then
    Exit Sub
    End If

Dim xPos As Integer
Dim xConnectionString As String
Dim StringSemIP As String
Dim IPAtual As String
Dim NovoIP As String
Dim xNovaConnectionString As String

xPos = InStr(1, de_informa.cn_informa.ConnectionString, "Data Source=", vbTextCompare)
xConnectionString = de_informa.cn_informa.ConnectionString
StringSemIP = Mid(xConnectionString, 1, xPos - 1)
IPAtual = Mid(xConnectionString, xPos + 12)

    If ComboCON.Text = "Internet" Then
    xNovaConnectionString = StringSemIP & "Data Source=" & "200.160.204.10"
    ElseIf ComboCON.Text = "Local" Then
    xNovaConnectionString = StringSemIP & "Data Source=" & "192.9.205.3"
    End If

de_informa.cn_informa = xNovaConnectionString

LblLogOn.Visible = False
ComboCON.Visible = False
CmdProsseguir.Visible = False

LblUsuario.Visible = True
LblSenha.Visible = True
txtUser.Visible = True
txtPassWord.Visible = True
cmdEntrar.Visible = True
cmdSair.Visible = True
txtUser.SetFocus
DoEvents
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub ComboCON_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
xTry = 0

ComboCON.AddItem "Internet"
ComboCON.AddItem "Local"

LblLogOn.Visible = True
ComboCON.Visible = True
CmdProsseguir.Visible = True

LblUsuario.Visible = False
LblSenha.Visible = False
txtUser.Visible = False
txtPassWord.Visible = False
cmdEntrar.Visible = False
cmdSair.Visible = False
DoEvents
End Sub

Private Sub txtPassWord_GotFocus()
txtPassWord.SelStart = 0
txtPassWord.SelLength = 200
End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtUser_GotFocus()
txtUser.SelStart = 0
txtUser.SelLength = 200
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtUser_LostFocus()
txtUser = UCase(txtUser)
End Sub
