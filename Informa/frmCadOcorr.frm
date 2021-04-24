VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadOcorr 
   Caption         =   "Cadastro de Ocorrências"
   ClientHeight    =   7635
   ClientLeft      =   1575
   ClientTop       =   585
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   8745
   Begin VB.Frame Frame4 
      Caption         =   "Ordenar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   15
      Top             =   210
      Width           =   3165
      Begin VB.OptionButton optOrdemDes 
         Caption         =   "Por Descrição"
         Height          =   225
         Left            =   1575
         TabIndex        =   9
         Top             =   315
         Width           =   1380
      End
      Begin VB.OptionButton optOrdemCod 
         Caption         =   "Por Código"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame fraGravar 
      Caption         =   "Gravação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   2745
      Begin VB.CommandButton cmdCanc 
         Caption         =   "Cancelar"
         Height          =   435
         Left            =   1470
         TabIndex        =   5
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   420
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Width           =   1170
      End
   End
   Begin MSDataGridLib.DataGrid gridCadOcorr 
      Bindings        =   "frmCadOcorr.frx":0000
      Height          =   2190
      Left            =   105
      TabIndex        =   10
      Top             =   945
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3863
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Sel_CadOcorrCod"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "cod_ocorr"
         Caption         =   "Cod."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "descricao"
         Caption         =   "Descrição Ocorrência"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "abonaSN"
         Caption         =   "Abona Atraso Entrega"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "env_email"
         Caption         =   "env_email"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "email1"
         Caption         =   "email1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "email2"
         Caption         =   "email2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "email3"
         Caption         =   "email3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "email4"
         Caption         =   "email4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "email_cliente"
         Caption         =   "email_cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5894,929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1679,811
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   989,858
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraComandos 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   5415
      Begin VB.CommandButton cmdAltera 
         Caption         =   "Alteração"
         Height          =   375
         Left            =   1440
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpr 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdIncl 
         Caption         =   "Inclusão"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraDados 
      Caption         =   "Ocorrências"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   8535
      Begin VB.TextBox txtEmailInt4 
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox txtEmailInt3 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   4095
      End
      Begin VB.CheckBox chkEnvEmailCli 
         Caption         =   $"frmCadOcorr.frx":0019
         Height          =   615
         Left            =   2400
         TabIndex        =   25
         Top             =   1200
         Width           =   6015
      End
      Begin VB.CheckBox chkEnvEmailInt 
         Caption         =   "Email Automático Para os Endereços Abaixo (INTEC)."
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtEmailInt1 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox txtEmailInt2 
         Height          =   285
         Left            =   4320
         TabIndex        =   20
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Frame fraAbona 
         Caption         =   "Abona Atraso ?"
         Height          =   615
         Left            =   6720
         TabIndex        =   17
         Top             =   360
         Width           =   1695
         Begin VB.OptionButton optAbonaNao 
            Caption         =   "Não"
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optAbonaSim 
            Caption         =   "Sim"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Descrição"
         Height          =   615
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   5775
         Begin VB.TextBox txtDescOcorr 
            Height          =   285
            Left            =   105
            MaxLength       =   50
            TabIndex        =   7
            Top             =   210
            Width           =   5535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cod"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
         Begin VB.TextBox txtCodOcorr 
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   6
            Top             =   210
            Width           =   375
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Email 3:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Email 4:"
         Height          =   195
         Left            =   4320
         TabIndex        =   28
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Email 2:"
         Height          =   195
         Left            =   4320
         TabIndex        =   24
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   555
      End
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   5565
      TabIndex        =   16
      Top             =   420
      Width           =   75
   End
End
Attribute VB_Name = "frmCadOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAltera_Click()
    If Mid$(xdireitos, 3, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        fraComandos.Visible = False
        fraGravar.Visible = True
        fraDados.Enabled = True
        txtCodOcorr.Enabled = False
        txtDescOcorr.Enabled = False
        txtEmailInt1.BackColor = &HC0FFFF     'AMARELO
        txtEmailInt2.BackColor = &HC0FFFF     'AMARELO
        txtEmailInt3.BackColor = &HC0FFFF     'AMARELO
        txtEmailInt4.BackColor = &HC0FFFF     'AMARELO
        lblModo.Caption = "Alteração"
    End If
End Sub

Private Sub cmdCanc_Click()
    txtCodOcorr.Text = ""
    txtDescOcorr.Text = ""
    txtEmailInt1 = ""
    txtEmailInt2 = ""
    txtEmailInt3 = ""
    txtEmailInt4 = ""
    chkEnvEmailCli.Value = 0
    chkEnvEmailInt.Value = 0
    txtCodOcorr.Enabled = True
    txtDescOcorr.Enabled = True
    fraDados.Enabled = False
    fraGravar.Visible = False
    fraComandos.Visible = True
    lblModo.Caption = ""
    txtCodOcorr.BackColor = &H8000000E       'BRANCO
    txtDescOcorr.BackColor = &H8000000E       'BRANCO
    txtEmailInt1.BackColor = &H8000000E       'BRANCO
    txtEmailInt2.BackColor = &H8000000E       'BRANCO
    txtEmailInt3.BackColor = &H8000000E       'BRANCO
    txtEmailInt4.BackColor = &H8000000E       'BRANCO
    optOrdemCod_Click
    gridCadOcorr_Click
End Sub
Private Sub cmdIncl_Click()
    If Mid$(xdireitos, 3, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        fraComandos.Visible = False
        fraGravar.Visible = True
        fraDados.Enabled = True
        txtCodOcorr.SetFocus
        txtCodOcorr.Text = ""
        txtDescOcorr.Text = ""
        txtEmailInt1 = ""
        txtEmailInt2 = ""
        txtEmailInt3 = ""
        txtEmailInt4 = ""
        chkEnvEmailCli.Value = 0
        chkEnvEmailInt.Value = 0
        txtCodOcorr.BackColor = &HC0FFFF      'AMARELO
        txtDescOcorr.BackColor = &HC0FFFF      'AMARELO
        txtEmailInt1.BackColor = &HC0FFFF      'AMARELO
        txtEmailInt2.BackColor = &HC0FFFF      'AMARELO
        txtEmailInt3.BackColor = &HC0FFFF      'AMARELO
        txtEmailInt4.BackColor = &HC0FFFF      'AMARELO
        lblModo.Caption = "Inclusão"
    End If
End Sub
Private Sub cmdOk_Click()
    Dim xabona As String, xenv_email As String, xenv_emailCLI As String
        If Len(txtCodOcorr.Text) = 0 Then
            MsgBox "Código de Ocorrência Inválido !"
            txtCodOcorr.SetFocus
            Exit Sub
        ElseIf Len(txtDescOcorr.Text) < 5 Then
            MsgBox "Descrição de Ocorrência Inválido !"
            txtDescOcorr.SetFocus
            Exit Sub
        End If
        If Len(Trim$((txtEmailInt1 & txtEmailInt2 & txtEmailInt3 & txtEmailInt4))) = 0 Then
            chkEnvEmailInt.Value = 0
        End If
        DoEvents
        If optAbonaSim = True Then
            xabona = "S"
        Else
            xabona = "N'"
        End If
        If chkEnvEmailInt.Value = 1 Then
            xenv_email = "S"
        Else
            xenv_email = "N"
        End If
        If chkEnvEmailCli.Value = 1 Then
            xenv_emailCLI = "S"
        Else
            xenv_emailCLI = "N"
        End If
        
        If lblModo = "Inclusão" Then
            'Inclue no banco de dados
            de_informa.ins_cadocorr txtCodOcorr.Text, txtDescOcorr.Text, xabona, xenv_email, txtEmailInt1, txtEmailInt2, txtEmailInt3, txtEmailInt4, xenv_emailCLI
        Else
            'Altera no banco de dados
            de_informa.alt_cadocorr txtCodOcorr.Text, xabona, xenv_email, txtEmailInt1, txtEmailInt2, txtEmailInt3, txtEmailInt4, xenv_emailCLI
        End If
        
'atualiza os recordsets
        If de_informa.rsSel_CadOcorrCod.State = 1 Then de_informa.rsSel_CadOcorrCod.Close
        de_informa.Sel_CadOcorrCod   'rs por ordem de codigo
        If de_informa.rsSel_CadOcorrDes.State = 1 Then de_informa.rsSel_CadOcorrDes.Close
        de_informa.Sel_CadOcorrDes   'rs por ordem de descricao
'Atualiza o GRID
        If optOrdemCod.Value = True Then
            optOrdemCod_Click
        Else
            optOrdemDes_Click
        End If
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "CAD.DE CÓD.OCORRÊNCIA: " & txtDescOcorr
        
        cmdCanc_Click
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
End Sub
Private Sub Form_Activate()
    mdiInforma.StatusBar1.Panels(2).Text = "Para ALTERAR ou EXCLUIR Ocorrencias consulte o Gestor do Sistema"
    gridCadOcorr.DataMember = "sel_cadocorrcod"
    gridCadOcorr.Refresh
End Sub
Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    gridCadOcorr_Click
    'txtCodOcorr.Enabled = False
    ' txtDescOcorr.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmCadOcorr = Nothing
    mdiInforma.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub gridCadOcorr_Click()
    txtCodOcorr = gridCadOcorr.Columns(0)
    txtDescOcorr = gridCadOcorr.Columns(1)
    If gridCadOcorr.Columns(2) = "S" Then
        optAbonaSim.Value = True
    Else
        optAbonaNao.Value = True
    End If
    If gridCadOcorr.Columns(3) = "S" Then
        chkEnvEmailInt.Value = 1
    Else
        chkEnvEmailInt.Value = 0
    End If
    If gridCadOcorr.Columns(8) = "S" Then
        chkEnvEmailCli.Value = 1
    Else
        chkEnvEmailCli.Value = 0
    End If
    txtEmailInt1 = gridCadOcorr.Columns(4)
    txtEmailInt2 = gridCadOcorr.Columns(5)
    txtEmailInt3 = gridCadOcorr.Columns(6)
    txtEmailInt4 = gridCadOcorr.Columns(7)
    
End Sub

Private Sub optOrdemCod_Click()
    gridCadOcorr.DataMember = "sel_cadocorrcod"
    gridCadOcorr.Refresh
End Sub
Private Sub optOrdemDes_Click()
    gridCadOcorr.DataMember = "sel_cadocorrdes"
    gridCadOcorr.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCodOcorr_GotFocus()
    txtCodOcorr.SelStart = 0
    txtCodOcorr.SelLength = 2
End Sub

Private Sub txtCodOcorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCodOcorr_LostFocus()
If txtCodOcorr.Text <> "" Then
    If Len(txtCodOcorr.Text) = 1 Or Not IsNumeric(txtCodOcorr.Text) Then
        MsgBox "Código de Ocorrência Inválido !", vbOKOnly + vbCritical, "ERRO"
        txtCodOcorr.SetFocus
        Exit Sub
    End If
    If Len(txtCodOcorr.Text) = 2 Then
        If de_informa.rsSel_ConsCadOcor.State = 1 Then de_informa.rsSel_ConsCadOcor.Close
        de_informa.Sel_ConsCadOcor txtCodOcorr.Text
        If de_informa.rsSel_ConsCadOcor.RecordCount > 0 Then
            txtDescOcorr.Text = de_informa.rsSel_ConsCadOcor.Fields("descricao")
            txtCodOcorr.SetFocus
        Else
            txtDescOcorr.Enabled = True
            txtDescOcorr.Text = ""
            txtDescOcorr.SetFocus
        End If
    End If
End If
End Sub
Private Sub txtDescOcorr_GotFocus()
    txtDescOcorr.SelStart = 0
    txtDescOcorr.SelLength = 40
End Sub

Private Sub txtDescOcorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtDescOcorr_LostFocus()
    txtDescOcorr.Text = UCase(txtDescOcorr.Text)
End Sub
