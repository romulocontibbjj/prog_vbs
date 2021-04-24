VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Configurações Informa"
   ClientHeight    =   4020
   ClientLeft      =   1995
   ClientTop       =   2145
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7785
   Begin VB.Frame Frame1 
      Caption         =   "Opções de Configurações"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   495
         Left            =   6240
         TabIndex        =   3
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Height          =   495
         Left            =   4800
         TabIndex        =   2
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox chkEnvEmail 
         Caption         =   "Envio de Emails por Este Micro Instalado"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "01"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdGravar_Click()
    'atualmente somente 1 linha então não precisa de do until
    Open App.Path & "\INFORMA.INI" For Output As #1
    If chkEnvEmail.Value = 1 Then
        xlinha = "01=SIM"
    Else
        xlinha = "01=NAO"
    End If
    Print #1, xlinha
    Close #1
    MsgBox "OK ! Configuração Gravada."
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Dir(App.Path & "\INFORMA.INI") = "" Then
    Else
        Open App.Path & "\INFORMA.INI" For Input As #1
        Line Input #1, xlinha
        If Mid(xlinha, 4, 3) = "SIM" Then
            chkEnvEmail.Value = 1
        Else
            chkEnvEmail.Value = 0
        End If
        Close #1
    End If
End Sub
