VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRecalcPrazos 
   Caption         =   "Recalcular Prazos de Entrega"
   ClientHeight    =   1950
   ClientLeft      =   2880
   ClientTop       =   2100
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   6285
   Begin VB.Frame fraDados 
      Caption         =   "Seleção do Período (data de emissão)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6045
      Begin VB.TextBox txtCgc 
         Height          =   285
         Left            =   1150
         TabIndex        =   8
         Text            =   "%"
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processa..."
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2625
         TabIndex        =   2
         Top             =   420
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPer1 
         Height          =   285
         Left            =   1155
         TabIndex        =   1
         Top             =   420
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CGC Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblaguarde 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2415
         TabIndex        =   6
         Top             =   420
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Período:  De"
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   420
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmRecalcPrazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSair_Click()
    Unload Me
End Sub

Private Sub cmdProcessa_Click()
    If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
        MsgBox "Período Inválido !"
        mskPer1.SetFocus
        Exit Sub
    End If
    If CDate(mskPer1) > CDate(mskPer2) Then
        MsgBox "Período Inválido !"
        mskPer1.SetFocus
        Exit Sub
    End If
    
    lblaguarde.Caption = "Aguarde Processamento..."
    DoEvents
    de_informa.alt_RecalcPrazo mskPer1, mskPer2, txtCgc
    frmAtualPrazos.Show 1
    DoEvents
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "RECALCULAR PRAZOS DE ENTREGA"
    
    lblaguarde.Caption = "OK. Processo Finalizado !"
    cmdSair.SetFocus
End Sub

Private Sub cmdProcessa2_Click()

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmRecalcPrazos = Nothing
End Sub

Private Sub mskPer1_GotFocus()
    mskPer1.SelStart = 0
    mskPer1.SelLength = 10
End Sub

Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer1_LostFocus()
    mskPer1.Text = century(mskPer1.Text)
End Sub

Private Sub mskPer2_GotFocus()
    mskPer2.SelStart = 0
    mskPer2.SelLength = 10
End Sub

Private Sub mskPer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer2_LostFocus()
    mskPer2.Text = century(mskPer2.Text)
End Sub
