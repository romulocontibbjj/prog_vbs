VERSION 5.00
Begin VB.Form frm_Acesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Coty - Acesso ao Sistema"
   ClientHeight    =   2100
   ClientLeft      =   5970
   ClientTop       =   4425
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdsair 
         Caption         =   "Sair"
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "OK"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txt_login 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt_senha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frm_Acesso.frx":0000
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Senha"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_Acesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()

Dim xlinha As String
Dim xstrcon As String
Dim TAM As Integer

If Xbusca("C:\Coty.cnx", True) = False Then
    MsgBox ("Crie o arquivo com a ConectionString ..."), vbCritical + vbInformation + vbOKOnly
Else
    Open "C:\Coty.Cnx" For Input As #1
        Line Input #1, xlinha
        If Mid(xlinha, 1, 3) = "CNX" Then
            TAM = Len(xlinha)
            xstrcon = Mid(xlinha, 5, TAM)
            If deb_coty.Connection1.State = 1 Then deb_coty.Connection1.Close
                deb_coty.Connection1.ConnectionString = xstrcon
                deb_coty.Connection1.Open xstrcon
            Else
                deb_coty.Connection1.ConnectionString = xstrcon
            End If
        End If
    Close #1
    
    If deb_coty.rsSel_userLogin.State = 1 Then deb_coty.rsSel_userLogin.Close
    deb_coty.Sel_userLogin txt_login, txt_senha
    
    If deb_coty.rsSel_userLogin.EOF Then
        MsgBox ("Usuário ou senha inválida"), vbCritical + vbOKOnly
        txt_login.SetFocus
        Exit Sub
    Else
        MDIForm1.Show
    End If

Unload Me

End Sub

Private Sub cmdsair_Click()
If deb_coty.Connection1.State = 1 Then
    deb_coty.Connection1.Close
End If
End
End Sub

