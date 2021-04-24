VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPesquisaCTC 
   Caption         =   "Pesquisa CTC"
   ClientHeight    =   6825
   ClientLeft      =   855
   ClientTop       =   1725
   ClientWidth     =   12810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrid 
      Caption         =   "Resultado - CTCs Selecionados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   33
      Top             =   3480
      Width           =   11775
      Begin VB.CommandButton cmdImprPesq 
         Caption         =   "Imprimir Resultado da Pesquisa ..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   6000
         TabIndex        =   38
         ToolTipText     =   "Imprime os CTCs Resultantes da Pesquina."
         Top             =   240
         Width           =   2715
      End
      Begin VB.CommandButton cmdTransporta 
         Caption         =   "Transporta CTC Selecionado ..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   8880
         TabIndex        =   37
         ToolTipText     =   "Selecione o CTC clicando na Borda Esquerda da Grade e Clique Aqui."
         Top             =   240
         Width           =   2715
      End
      Begin MSDataGridLib.DataGrid GridPesqCTC 
         Bindings        =   "frmPesquisaCTC.frx":0000
         Height          =   3975
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   7011
         _Version        =   393216
         BackColor       =   8388608
         ForeColor       =   8454143
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
         DataMember      =   "Sel_PesqCTC"
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "filialctc"
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
            DataField       =   "data"
            Caption         =   "data"
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
            DataField       =   "remet_nome"
            Caption         =   "remet_nome"
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
            DataField       =   "dest_nome"
            Caption         =   "dest_nome"
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
            DataField       =   "cidade_dest"
            Caption         =   "cidade_dest"
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
            DataField       =   "uf_dest"
            Caption         =   "uf_dest"
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
            DataField       =   "nfs"
            Caption         =   "nfs"
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
            DataField       =   "valmerc"
            Caption         =   "valmerc"
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
            DataField       =   "peso"
            Caption         =   "peso"
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
         BeginProperty Column09 
            DataField       =   "volumes"
            Caption         =   "volumes"
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
         BeginProperty Column10 
            DataField       =   "natureza"
            Caption         =   "natureza"
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
         BeginProperty Column11 
            DataField       =   "modal"
            Caption         =   "modal"
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
         BeginProperty Column12 
            DataField       =   "transp_sub"
            Caption         =   "transp_sub"
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
            BeginProperty Column00 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3360,189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3374,929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2145,26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
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
      Height          =   3240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CommandButton cmdSair 
         Caption         =   "Canc/Sair"
         Height          =   450
         Left            =   10320
         TabIndex        =   16
         Top             =   2640
         Width           =   1275
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processa..."
         Height          =   450
         Left            =   10320
         TabIndex        =   15
         Top             =   2040
         Width           =   1275
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
         Height          =   1215
         Left            =   6840
         TabIndex        =   30
         Top             =   1920
         Width           =   3375
         Begin VB.Frame Frame9 
            Caption         =   "Cidade comece com ..."
            Height          =   735
            Left            =   1320
            TabIndex        =   32
            Top             =   360
            Width           =   1935
            Begin VB.TextBox txtCidade 
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "UF"
            Height          =   735
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1095
            Begin VB.ComboBox cmbUF 
               Height          =   315
               Left            =   120
               TabIndex        =   13
               Text            =   "Todos"
               Top             =   360
               Width           =   855
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
         Left            =   5040
         TabIndex        =   23
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtCGCDes 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   8
            TabIndex        =   3
            Top             =   600
            Width           =   1590
         End
         Begin VB.CommandButton cmdBuscaDES 
            Caption         =   "?"
            Height          =   375
            Left            =   4080
            TabIndex        =   4
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Primeiros 8 Caracteres"
            Height          =   195
            Left            =   2400
            TabIndex        =   35
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblNomeDes 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   24
            Top             =   960
            Width           =   3375
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
         Height          =   1215
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   6615
         Begin VB.Frame Frame6 
            Caption         =   "No Período de ...  (máximo de 30 dias)"
            Height          =   735
            Left            =   3480
            TabIndex        =   27
            Top             =   360
            Width           =   3015
            Begin MSMask.MaskEdBox mskPer2 
               Height          =   285
               Left            =   1680
               TabIndex        =   12
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
               TabIndex        =   11
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
               TabIndex        =   28
               Top             =   360
               Width           =   90
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Emissão nos Últimos ..."
            Height          =   735
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   2895
            Begin VB.OptionButton opt30d 
               Caption         =   "30 dias"
               Height          =   195
               Left            =   1920
               TabIndex        =   10
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton opt20d 
               Caption         =   "20 dias"
               Height          =   195
               Left            =   960
               TabIndex        =   9
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton opt5d 
               Caption         =   "5 dias"
               Height          =   195
               Left            =   120
               TabIndex        =   8
               Top             =   360
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "OU"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   29
            Top             =   600
            Width           =   495
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
         TabIndex        =   18
         Top             =   360
         Width           =   4815
         Begin VB.CommandButton cmdBuscaREM 
            Caption         =   "?"
            Height          =   375
            Left            =   4080
            TabIndex        =   2
            Top             =   840
            Width           =   495
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
         Begin VB.TextBox txtCGCRem 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   8
            TabIndex        =   1
            Top             =   600
            Width           =   1590
         End
         Begin VB.Label lbl8caract 
            AutoSize        =   -1  'True
            Caption         =   "Primeiros 8 Caracteres"
            Height          =   195
            Left            =   2400
            TabIndex        =   36
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblNomeRem 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   20
            Top             =   960
            Width           =   3375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   375
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
         Left            =   9960
         TabIndex        =   17
         Top             =   360
         Width           =   1695
         Begin VB.OptionButton optAir 
            Caption         =   "Aéreo"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   735
         End
         Begin VB.OptionButton optRodo 
            Caption         =   "Rodoviário"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   1080
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CheckBox chkModal 
            Caption         =   "Todos Modais"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Value           =   1  'Checked
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmPesquisaCTC"
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
        TxtCGCRem.MaxLength = 8
        lbl8caract.Visible = True
    Else
        TxtCGCRem.MaxLength = 14
        lbl8caract.Visible = False
    End If
    TxtCGCRem.SetFocus
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
    
    frmPesquisaCTC.MousePointer = 11
    
    If opt5d.Value = True Then
        mskPer1.Mask = ""
        mskPer1.Text = ""
        mskPer1.Mask = "##/##/####"
        mskPer2.Mask = ""
        mskPer2.Text = ""
        mskPer2.Mask = "##/##/####"
        xdata1 = datahora("data") - 5
        xdata2 = datahora("data")
    ElseIf opt20d.Value = True Then
        mskPer1.Mask = ""
        mskPer1.Text = ""
        mskPer1.Mask = "##/##/####"
        mskPer2.Mask = ""
        mskPer2.Text = ""
        mskPer2.Mask = "##/##/####"
        xdata1 = datahora("data") - 20
        xdata2 = datahora("data")
    ElseIf opt30d.Value = True Then
        mskPer1.Mask = ""
        mskPer1.Text = ""
        mskPer1.Mask = "##/##/####"
        mskPer2.Mask = ""
        mskPer2.Text = ""
        mskPer2.Mask = "##/##/####"
        xdata1 = datahora("data") - 30
        xdata2 = datahora("data")
    Else
        If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
            MsgBox "Período Escolhido Inválido !"
            mskPer1.SetFocus
            Me.MousePointer = 0
            Exit Sub
        End If
        If CDate(mskPer1) > CDate(mskPer2) Then
            MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
            mskPer1.SetFocus
            Me.MousePointer = 0
            Exit Sub
        End If
        xdata1 = CDate(mskPer1)
        xdata2 = CDate(mskPer2)
    End If
    
    If xdata2 - xdata1 > 32 Then
        MsgBox "Período Escolhido Maior que 30 Dias ! Escolha um Período Menor."
        mskPer1.SetFocus
        Me.MousePointer = 0
        Exit Sub
    End If

    If chkModal = 1 Then
        xmodal = "%"
    Else
        If optAir = True Then
            xmodal = "AEREO%"
        Else
            xmodal = "RODOVIARIO%"
        End If
    End If
    If cmbUf.Text = "Todos" Then
        xuf = "%"
    Else
        xuf = cmbUf.Text & "%"
    End If
    
    cmdProcessa.Caption = "Aguarde ..."
    cmdProcessa.Enabled = False
    CmdSair.Enabled = False
    fraDados.Enabled = False
    fraGrid.Enabled = False
    If de_informa.rsSel_PesqCTC.State = 1 Then de_informa.rsSel_PesqCTC.Close
    de_informa.Sel_PesqCTC RTrim(TxtCGCRem) & "%", RTrim(TxtCGCDES) & "%", xmodal, xdata1, xdata2, xuf, RTrim(txtCidade) & "%"
    fraGrid.Caption = "Resultado - CTCs Selecionados:" & CVar(de_informa.rsSel_PesqCTC.RecordCount) & " registros"
    If de_informa.rsSel_PesqCTC.RecordCount > 0 Then
        cmdTransporta.Enabled = True
    Else
        cmdTransporta.Enabled = False
    End If
    GridPesqCTC.DataMember = "Sel_PesqCTC"
    GridPesqCTC.Refresh
    fraDados.Enabled = True
    fraGrid.Enabled = True
    cmdProcessa.Caption = "Processa..."
    cmdProcessa.Enabled = True
    CmdSair.Enabled = True
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "PESQUISA DE CTCs (INFORM. SAC)"
    
    frmPesquisaCTC.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdTransporta_Click()
    Me.Hide
    xultimofilial = Mid(GridPesqCTC.Columns(0), 1, 2)
    xultimoctc = Mid(GridPesqCTC.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridPesqCTC.Columns(0), 1, 2)
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
    cmbUf.AddItem "Todos"
    Do Until de_informa.rsSel_Ufs.EOF
        cmbUf.AddItem de_informa.rsSel_Ufs.Fields("uf")
        de_informa.rsSel_Ufs.MoveNext
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPesquisaCTC = Nothing
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
        If CDate(mskPer1.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
        If IsDate(mskPer2.Text) Then
'            If CDate(mskPer2.Text) < CDate(mskPer1.Text) Then
'                MsgBox "Período Inválido !", vbCritical, "Erro"
'                mskPer1.SetFocus
'                Exit Sub
'            Else
                opt5d.Value = False
                opt20d.Value = False
                opt30d.Value = False
'            End If
        End If
    Else
        If mskPer2.Text = "__/__/____" And opt30d = False And opt60d = False And optPer15d = False Then
            opt5d.Value = True
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
        If CDate(mskPer2.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
        If IsDate(mskPer1.Text) Then
'            If CDate(mskPer2.Text) < CDate(mskPer1.Text) Then
'                MsgBox "Período Inválido !", vbCritical, "Erro"
'                'mskPer2.SetFocus
'                Exit Sub
'            Else
                opt5d.Value = False
                opt20d.Value = False
                opt30d.Value = False
'            End If
        End If
    Else
        If mskPer1.Text = "__/__/____" And opt30d = False And opt60d = False And optPer15d = False Then
            opt5d.Value = True
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
    If Len(TxtCGCDES) = TxtCGCDES.MaxLength Then chkModal.SetFocus
End Sub

Private Sub txtCGCDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCGCDes_LostFocus()
    If TxtCGCDES.Text <> "" Then
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli Trim(TxtCGCDES) & "%"
        If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
            lblNomeDes.Caption = de_informa.rsSel_ConsCadCli.Fields("nome")
        Else
            TxtCGCDES.SetFocus
        End If
    Else
        lblNomeDes.Caption = ""
    End If
End Sub
Private Sub txtCGCRem_Change()
    If Len(TxtCGCRem) = TxtCGCRem.MaxLength Then TxtCGCDES.SetFocus
End Sub

Private Sub txtCGCRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCGCRem_LostFocus()
    If TxtCGCRem.Text <> "" Then
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli Trim(TxtCGCRem) & "%"
        If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
            lblNomeRem.Caption = de_informa.rsSel_ConsCadCli.Fields("nome")
        Else
            TxtCGCRem.SetFocus
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
