VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnOperacao 
   Caption         =   "Análise de Operações"
   ClientHeight    =   6555
   ClientLeft      =   1575
   ClientTop       =   1260
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8805
   Begin VB.Frame fraMensagem 
      Caption         =   "Mensagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Visible         =   0   'False
      Width           =   8565
      Begin MSComCtl2.Animation Animation1 
         Height          =   645
         Left            =   3240
         TabIndex        =   37
         Top             =   240
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   1138
         _Version        =   393216
         AutoPlay        =   -1  'True
         Center          =   -1  'True
         FullWidth       =   327
         FullHeight      =   43
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Processando dados selecionados. Por favor aguarde ..."
         Height          =   480
         Left            =   105
         TabIndex        =   40
         Top             =   315
         Width           =   2565
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   210
         TabIndex        =   39
         Top             =   945
         Width           =   495
      End
      Begin VB.Label lblStatus 
         Caption         =   "Processo X/Y. Processando Qtde de CTCs entregues..."
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   960
         Width           =   6375
      End
   End
   Begin VB.Frame fraDados 
      Caption         =   "Seleção dos Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5040
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.Frame Frame5 
         Caption         =   "** CONSIGNATÁRIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   30
         Top             =   3480
         Width           =   4815
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   14
            TabIndex        =   32
            Top             =   600
            Width           =   1590
         End
         Begin VB.CommandButton Command1 
            Caption         =   "?"
            Height          =   375
            Left            =   4080
            TabIndex        =   31
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "14 Caracteres"
            Height          =   195
            Left            =   2400
            TabIndex        =   35
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   33
            Top             =   960
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "** Modal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5040
         TabIndex        =   26
         Top             =   3480
         Width           =   1695
         Begin VB.CheckBox chkModal 
            Caption         =   "Todos Modais"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.OptionButton optRodo 
            Caption         =   "Rodoviário"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optAir 
            Caption         =   "Aéreo"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "** Cliente REMETENTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtCGCRem 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   8
            TabIndex        =   22
            Top             =   600
            Width           =   1590
         End
         Begin VB.CheckBox chkTodosEstab 
            Caption         =   "Todos os Estabelecimentos"
            Height          =   225
            Left            =   2280
            TabIndex        =   21
            Top             =   240
            Value           =   1  'Checked
            Width           =   2325
         End
         Begin VB.CommandButton cmdBuscaREM 
            Caption         =   "?"
            Height          =   375
            Left            =   4080
            TabIndex        =   20
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblNomeRem 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   24
            Top             =   960
            Width           =   3375
         End
         Begin VB.Label lbl8caract 
            AutoSize        =   -1  'True
            Caption         =   "Primeiros 8 Caracteres"
            Height          =   195
            Left            =   2400
            TabIndex        =   23
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "** Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5040
         TabIndex        =   14
         Top             =   360
         Width           =   3375
         Begin VB.Frame Frame6 
            Caption         =   "No Período de ...  (máximo de 30 dias)"
            Height          =   855
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   3135
            Begin MSMask.MaskEdBox mskPer2 
               Height          =   285
               Left            =   1680
               TabIndex        =   16
               Top             =   360
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
               Left            =   120
               TabIndex        =   17
               Top             =   360
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "à"
               Height          =   195
               Left            =   1440
               TabIndex        =   18
               Top             =   360
               Width           =   90
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "** Cliente DESTINATÁRIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   4815
         Begin VB.CommandButton cmdBuscaDES 
            Caption         =   "?"
            Height          =   375
            Left            =   4080
            TabIndex        =   10
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtCGCDes 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   8
            TabIndex        =   9
            Top             =   600
            Width           =   1590
         End
         Begin VB.Label lblNomeDes 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   13
            Top             =   960
            Width           =   3375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Primeiros 8 Caracteres"
            Height          =   195
            Left            =   2400
            TabIndex        =   11
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "** Localidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5040
         TabIndex        =   3
         Top             =   1920
         Width           =   3375
         Begin VB.Frame Frame8 
            Caption         =   "UF"
            Height          =   855
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1095
            Begin VB.ComboBox cmbUF 
               Height          =   315
               Left            =   120
               TabIndex        =   7
               Text            =   "Todos"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Cidade comece com ..."
            Height          =   855
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   1935
            Begin VB.TextBox txtCidade 
               Height          =   285
               Left            =   120
               TabIndex        =   5
               Top             =   360
               Width           =   1695
            End
         End
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processa..."
         Height          =   450
         Left            =   6840
         TabIndex        =   2
         Top             =   3720
         Width           =   1515
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Canc/Sair"
         Height          =   450
         Left            =   6840
         TabIndex        =   1
         Top             =   4320
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmAnOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkModal_Click()
    If chkModal = 1 Then
        optAir = False
        optRodo = False
        optAir.Enabled = False
        optRodo.Enabled = False
    Else
        optAir.Enabled = True
        optRodo.Enabled = True
    End If
End Sub

Private Sub chkModal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub chkTodosEstab_Click()
    If chkTodosEstab.Value = 1 Then
        txtCGCRem.MaxLength = 8
        lbl8caract.Visible = True
    Else
        txtCGCRem.MaxLength = 14
        lbl8caract.Visible = False
    End If
    txtCGCRem.SetFocus
End Sub

Private Sub cmbUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdBuscaDES_Click()
    frmBuscaCLI.Caption = "Busca Cliente DESTINATÁRIO"
    frmBuscaCLI.Show 1
End Sub

Private Sub cmdBuscaREM_Click()
    frmBuscaCLI.Caption = "Busca Cliente REMETENTE"
    frmBuscaCLI.Show 1
End Sub

Private Sub cmdProcessa_Click()
    Dim xmodal As String, xuf As String, xdata1 As Date, xdata2 As Date
    If chkModal = 1 Then
        xmodal = "%"
    Else
        If optAir = True Then
            xmodal = "AEREO%"
        Else
            xmodal = "RODOVIARIO%"
        End If
    End If
    If cmbUF.Text = "Todos" Then
        xuf = "%"
    Else
        xuf = cmbUF.Text & "%"
    End If
    If optPer5d.Value = True Then
        xdata1 = Date - 5
        xdata2 = Date
    ElseIf opt20d.Value = True Then
        xdata1 = Date - 20
        xdata2 = Date
    ElseIf opt30d.Value = True Then
        xdata1 = Date - 30
        xdata2 = Date
    Else
        xdata1 = CDate(mskPer1)
        xdata2 = CDate(mskPer2)
    End If
    If xdata2 - xdata1 > 31 Then
        MsgBox "Período Escolhido Maior que 30 Dias ! Escolha um Período Menor."
        mskPer1.SetFocus
        Exit Sub
    End If
    If de_informa.rsSel_PesqCTC.State = 1 Then de_informa.rsSel_PesqCTC.Close
    de_informa.Sel_PesqCTC RTrim(txtCGCRem) & "%", RTrim(txtCGCDes) & "%", xmodal, xdata1, xdata2, xuf, RTrim(txtCidade) & "%"
    fraGrid.Caption = "Resultado - CTCs Selecionados:" & CVar(de_informa.rsSel_PesqCTC.RecordCount) & " registros"
    If de_informa.rsSel_PesqCTC.RecordCount > 0 Then
        cmdTransporta.Enabled = True
    Else
        cmdTransporta.Enabled = False
    End If
    GridPesqCTC.DataMember = "Sel_PesqCTC"
    GridPesqCTC.Refresh
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdTransporta_Click()
    Me.Hide
    frmSac.txtFilial = Mid(GridPesqCTC.Columns(0), 1, 2)
    frmSac.txtCtc = Mid(GridPesqCTC.Columns(0), 3, 8)
    DoEvents
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
    Unload frmPesquisaCTC
End Sub

Private Sub Form_Load()
    GridPesqCTC.DataMember = ""
    GridPesqCTC.Refresh
    If de_informa.rsSel_Ufs.State = 1 Then de_informa.rsSel_Ufs.Close
    de_informa.Sel_Ufs
    de_informa.rsSel_Ufs.MoveFirst
    cmbUF.AddItem "Todos"
    Do Until de_informa.rsSel_Ufs.EOF
        cmbUF.AddItem de_informa.rsSel_Ufs.Fields("uf")
        de_informa.rsSel_Ufs.MoveNext
    Loop
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
    If mskPer1.Text <> "__/__/____" Then
        mskPer1.Text = century(mskPer1.Text)
        If IsDate(mskPer1.Text) = False Or Mid(mskPer1.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
        If CDate(mskPer1.Text) > Date Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
        If IsDate(mskPer2.Text) Then
            If CDate(mskPer2.Text) < CDate(mskPer1.Text) Then
                MsgBox "Período Inválido !", vbCritical, "Erro"
                mskPer1.SetFocus
                Exit Sub
            Else
                opt20d.Value = False
                opt30d.Value = False
                optPer5d.Value = False
            End If
        End If
    Else
        If mskPer2.Text = "__/__/____" Then
            optPer5d.Value = True
        End If
    End If
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
    If mskPer2.Text <> "__/__/____" Then
        mskPer2.Text = century(mskPer2.Text)
        If IsDate(mskPer2.Text) = False Or Mid(mskPer2.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
        If CDate(mskPer2.Text) > Date Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
        If IsDate(mskPer1.Text) Then
            If CDate(mskPer2.Text) < CDate(mskPer1.Text) Then
                MsgBox "Período Inválido !", vbCritical, "Erro"
                mskPer2.SetFocus
                Exit Sub
            Else
                opt20d.Value = False
                opt30d.Value = False
                optPer5d.Value = False
            End If
        End If
    Else
        If mskPer1.Text = "__/__/____" Then
            optPer5d.Value = True
        End If
    End If
End Sub

Private Sub opt20d_Click()
    mskPer1.Mask = ""
    mskPer1.Text = ""
    mskPer1.Mask = "##/##/####"
    mskPer2.Mask = ""
    mskPer2.Text = ""
    mskPer2.Mask = "##/##/####"
End Sub

Private Sub opt20d_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub opt30d_Click()
    mskPer1.Mask = ""
    mskPer1.Text = ""
    mskPer1.Mask = "##/##/####"
    mskPer2.Mask = ""
    mskPer2.Text = ""
    mskPer2.Mask = "##/##/####"
End Sub

Private Sub opt30d_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optPer5d_Click()
    mskPer1.Mask = ""
    mskPer1.Text = ""
    mskPer1.Mask = "##/##/####"
    mskPer2.Mask = ""
    mskPer2.Text = ""
    mskPer2.Mask = "##/##/####"
End Sub

Private Sub optPer5d_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCGCDes_Change()
    If Len(txtCGCDes) = txtCGCDes.MaxLength Then chkModal.SetFocus
End Sub

Private Sub txtCGCDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCGCDes_LostFocus()
    If txtCGCDes.Text <> "" Then
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli Trim(txtCGCDes) & "%"
        If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
            lblNomeDes.Caption = de_informa.rsSel_ConsCadCli.Fields("nome")
        Else
            txtCGCDes.SetFocus
        End If
    Else
        lblNomeDes.Caption = ""
    End If
End Sub
Private Sub txtCGCRem_Change()
    If Len(txtCGCRem) = txtCGCRem.MaxLength Then txtCGCDes.SetFocus
End Sub

Private Sub txtCGCRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCGCRem_LostFocus()
    If txtCGCRem.Text <> "" Then
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli Trim(txtCGCRem) & "%"
        If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
            lblNomeRem.Caption = de_informa.rsSel_ConsCadCli.Fields("nome")
        Else
            txtCGCRem.SetFocus
        End If
    Else
        lblNomeRem.Caption = ""
    End If
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCidade_LostFocus()
    txtCidade = UCase(txtCidade)
End Sub

