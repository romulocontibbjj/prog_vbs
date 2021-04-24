VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAcompanha 
   Caption         =   "Informa - Acompanhamento de Cliente"
   ClientHeight    =   8775
   ClientLeft      =   600
   ClientTop       =   1065
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame fraComandos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   33
      Top             =   1560
      Width           =   11775
      Begin VB.Frame fraDadosPor 
         Caption         =   "Dados Por..."
         Height          =   495
         Left            =   4320
         TabIndex        =   71
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton optPorNF 
            Caption         =   "NF"
            Height          =   195
            Left            =   1080
            TabIndex        =   73
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optPorCTC 
            Caption         =   "CTC"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CheckBox chkProcEmTransito 
         Caption         =   "Em Trânsito"
         Height          =   195
         Left            =   2520
         TabIndex        =   20
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   18
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "S A I R"
         Height          =   375
         Left            =   10680
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkProcEntregue 
         Caption         =   "Entregue"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdNovaSel 
         Caption         =   "Nova Seleção"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9120
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processar / Atualizar Dados"
         Height          =   375
         Left            =   6480
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Processar:"
         Height          =   195
         Left            =   1560
         TabIndex        =   69
         Top             =   310
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Somente da Filial:"
         Height          =   435
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame FraPeriodo 
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
      Height          =   1335
      Left            =   6720
      TabIndex        =   29
      Top             =   120
      Width           =   5175
      Begin VB.Frame fraPorEmissao 
         Height          =   855
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   3015
         Begin VB.OptionButton opt60d 
            Caption         =   "Últimos 60 dias"
            Height          =   435
            Left            =   1920
            TabIndex        =   63
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optPer15d 
            Caption         =   "Últimos 15 dias"
            Height          =   195
            Left            =   600
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton opt30d 
            Caption         =   "Últimos 30 dias"
            Height          =   195
            Left            =   600
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame fraPorPeriodo 
         Height          =   855
         Left            =   2040
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   1680
            TabIndex        =   17
            Top             =   480
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
            TabIndex        =   16
            Top             =   480
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "( Intervalo Máximo de 30 dias )"
            Height          =   195
            Left            =   390
            TabIndex        =   64
            Top             =   120
            Width           =   2160
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1440
            TabIndex        =   31
            Top             =   480
            Width           =   90
         End
      End
      Begin VB.OptionButton optPorMes 
         Caption         =   "Por Mês..."
         Height          =   315
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optPorPeriodo 
         Caption         =   "Por Período..."
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optPorEmissao 
         Caption         =   "Emissão nos..."
         Height          =   315
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Frame fraPorMesAno 
         Height          =   855
         Left            =   2040
         TabIndex        =   65
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.ComboBox comboMesAnoAcomp 
            Height          =   315
            Left            =   480
            TabIndex        =   66
            Text            =   "Mes/Ano"
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.Frame fraRemetente 
      Caption         =   "Opção de Seleção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraRegiao 
         Height          =   960
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   4575
         Begin VB.TextBox txtregiaosac 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   2
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Atendimento:"
            Height          =   195
            Left            =   1560
            TabIndex        =   26
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "UFs:"
            Height          =   255
            Left            =   4080
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblUfs 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label lblAtendSac 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2520
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Região SAC:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame fraCliente 
         Height          =   960
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   4575
         Begin VB.OptionButton optRemetente 
            Caption         =   "Remetente"
            Height          =   195
            Left            =   3240
            TabIndex        =   5
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optDestinatario 
            Caption         =   "Destinatário"
            Height          =   195
            Left            =   3240
            TabIndex        =   6
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtCGCRem 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   8
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdBuscaREM 
            Caption         =   "?"
            Height          =   255
            Left            =   2280
            TabIndex        =   36
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox chkTodosEstab 
            Caption         =   "Todos Estabelecim."
            Height          =   225
            Left            =   2760
            TabIndex        =   4
            Top             =   240
            Value           =   1  'Checked
            Width           =   1725
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblNomeRem 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Height          =   960
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optSelReg 
            Caption         =   "Por Região SAC"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optSelCli 
            Caption         =   "Por Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   1455
         End
      End
   End
   Begin TabDlg.SSTab TabGeral 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "SEM POSIÇÃO"
      TabPicture(0)   =   "frmAcompanha.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNfsSemPos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCTCsSemPos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GridSemPos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGerarArqSemPos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GridSemPosCtc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSACSemPos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGerarArqSemPosCtc"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "EM OCORRÊNCIA"
      TabPicture(1)   =   "frmAcompanha.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdGerarArqOcorrCtc"
      Tab(1).Control(1)=   "cmdSACOcorr"
      Tab(1).Control(2)=   "cmdGerarArqOcorr"
      Tab(1).Control(3)=   "GridOcorrCTC"
      Tab(1).Control(4)=   "GridOcorr"
      Tab(1).Control(5)=   "lblCtcsEmOcorr"
      Tab(1).Control(6)=   "lblNfsEmOcorr"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "EM TRÂNSITO"
      TabPicture(2)   =   "frmAcompanha.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdGerarArqTransitoCtc"
      Tab(2).Control(1)=   "cmdSACTransito"
      Tab(2).Control(2)=   "GridTransitoCTC"
      Tab(2).Control(3)=   "GridTransito"
      Tab(2).Control(4)=   "cmdGerarArqTransito"
      Tab(2).Control(5)=   "lblAviso2"
      Tab(2).Control(6)=   "lblCtcsTransito"
      Tab(2).Control(7)=   "lblNfsTransito"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "ENTREGUE"
      TabPicture(3)   =   "frmAcompanha.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdGerarArqEntregueCtc"
      Tab(3).Control(1)=   "cmdSACEntregue"
      Tab(3).Control(2)=   "GridEntregueCtc"
      Tab(3).Control(3)=   "GridEntregue"
      Tab(3).Control(4)=   "cmdGerarArqEntregue"
      Tab(3).Control(5)=   "lblAviso"
      Tab(3).Control(6)=   "lblCtcsEntregue"
      Tab(3).Control(7)=   "lblNfsEntregue"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "BAIXA s/Entrega"
      TabPicture(4)   =   "frmAcompanha.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdProcessarBxSemEntr"
      Tab(4).Control(1)=   "cmdSACBxSemEntr"
      Tab(4).Control(2)=   "GridBaixa"
      Tab(4).Control(3)=   "lblNfsBaixa"
      Tab(4).ControlCount=   4
      Begin VB.CommandButton cmdGerarArqOcorrCtc 
         Caption         =   "Geração de Arquivo - EM OCORRÊNCIA - POR CTC ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67680
         TabIndex        =   80
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdSACOcorr 
         Caption         =   "Consulta SAC..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69480
         TabIndex        =   77
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdGerarArqOcorr 
         Caption         =   "Geração de Arquivo - EM OCORRÊNCIA - POR NF ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67680
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CommandButton cmdProcessarBxSemEntr 
         Caption         =   "Processar Bx. sem Entrega"
         Height          =   330
         Left            =   -67680
         TabIndex        =   70
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdGerarArqEntregueCtc 
         Caption         =   "Geração de Arquivo - ENTREGUE - POR CTC ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67680
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CommandButton cmdSACEntregue 
         Caption         =   "Consulta SAC..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69480
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdGerarArqTransitoCtc 
         Caption         =   "Geração de Arquivo - EM TRÂNSITO - POR CTC ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67680
         TabIndex        =   46
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdSACTransito 
         Caption         =   "Consulta SAC..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69480
         TabIndex        =   45
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdGerarArqSemPosCtc 
         Caption         =   "Geração de Arquivo - SEM POSIÇÃO - POR CTC ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   39
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdSACSemPos 
         Caption         =   "Consulta SAC..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdSACBxSemEntr 
         Caption         =   "Consulta SAC..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -65160
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid GridBaixa 
         Bindings        =   "frmAcompanha.frx":008C
         Height          =   4335
         Left            =   -74880
         TabIndex        =   28
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompBaixa"
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "Filial CTC"
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
            Caption         =   "Emissão"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "dtentr"
            Caption         =   "Entrega"
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratada"
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
            DataField       =   "nfs"
            Caption         =   "Notas Fiscais"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column13 
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column14 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column15 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
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
         BeginProperty Column19 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column20 
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
         BeginProperty Column21 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3479,811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3240
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3089,764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1500,095
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   8384,882
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   10500,09
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridSemPosCtc 
         Bindings        =   "frmAcompanha.frx":00A5
         Height          =   4455
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompSemPosCTC"
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "nfs"
            Caption         =   "Notas Fiscais"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column13 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column14 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "obs_emissao"
            Caption         =   "obs_emissao"
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
         BeginProperty Column21 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3390,236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3525,166
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2399,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1574,929
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   6329,764
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   10695,12
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdGerarArqSemPos 
         Caption         =   "Geração de Arquivo - SEM POSIÇÃO - POR NF ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
      End
      Begin MSDataGridLib.DataGrid GridSemPos 
         Bindings        =   "frmAcompanha.frx":00BE
         Height          =   4455
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompSemPos"
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "numnf"
            Caption         =   "NF"
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
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            DataField       =   "data"
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column13 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column14 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "obs_emissao"
            Caption         =   "obs_emissao"
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
         BeginProperty Column21 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3270,047
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3209,953
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2475,213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   10695,12
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridTransitoCTC 
         Bindings        =   "frmAcompanha.frx":00D7
         Height          =   4335
         Left            =   -74880
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompTransitoCTC"
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "nfs"
            Caption         =   "Notas Fiscais"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column13 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column14 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "obs_emissao"
            Caption         =   "obs_emissao"
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
         BeginProperty Column21 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3390,236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3525,166
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2399,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1574,929
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   6329,764
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   10695,12
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridTransito 
         Bindings        =   "frmAcompanha.frx":00F0
         Height          =   4335
         Left            =   -74880
         TabIndex        =   50
         Top             =   960
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompTransito"
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "numnf"
            Caption         =   "NF"
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
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            DataField       =   "data"
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column13 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column14 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "obs_emissao"
            Caption         =   "obs_emissao"
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
         BeginProperty Column21 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3270,047
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3209,953
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2475,213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   10695,12
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridEntregueCtc 
         Bindings        =   "frmAcompanha.frx":0109
         Height          =   4335
         Left            =   -74880
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompEntregueCTC"
         ColumnCount     =   24
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "dtentr"
            Caption         =   "Entrega"
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "nfs"
            Caption         =   "Notas Fiscais"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
         BeginProperty Column13 
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column14 
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column15 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column16 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column17 
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
         BeginProperty Column18 
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column21 
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
         BeginProperty Column22 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
         BeginProperty Column23 
            DataField       =   "usu_datapre"
            Caption         =   "usu_datapre"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154,835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3509,858
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3614,74
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3060,284
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   420,095
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   5850,142
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   10709,86
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column22 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column23 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridEntregue 
         Bindings        =   "frmAcompanha.frx":0122
         Height          =   4335
         Left            =   -74880
         TabIndex        =   58
         Top             =   960
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompEntregue"
         ColumnCount     =   25
         BeginProperty Column00 
            DataField       =   "numnf"
            Caption         =   "NF"
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
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            DataField       =   "data"
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "dtentr"
            Caption         =   "Entrega"
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
         BeginProperty Column13 
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column14 
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column15 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column16 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column17 
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
         BeginProperty Column18 
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column21 
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
         BeginProperty Column22 
            DataField       =   "obs_emissao"
            Caption         =   "obs_emissao"
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
         BeginProperty Column23 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
         BeginProperty Column24 
            DataField       =   "usu_datapre"
            Caption         =   "usu_datapre"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3360,189
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3420,284
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2894,74
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   404,787
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1349,858
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1635,024
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   10725,17
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column22 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column23 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column24 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdGerarArqTransito 
         Caption         =   "Geração de Arquivo - EM TRÂNSITO - POR NF ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67680
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CommandButton cmdGerarArqEntregue 
         Caption         =   "Geração de Arquivo - ENTREGUE - POR NF ..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67680
         TabIndex        =   57
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
      End
      Begin MSDataGridLib.DataGrid GridOcorrCTC 
         Bindings        =   "frmAcompanha.frx":013B
         Height          =   4455
         Left            =   -74880
         TabIndex        =   74
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompOcorrCTC"
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "nfs"
            Caption         =   "Notas Fiscais"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column13 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column14 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "obs_emissao"
            Caption         =   "obs_emissao"
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
         BeginProperty Column21 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3390,236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3525,166
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2399,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1574,929
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   6329,764
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   10695,12
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridOcorr 
         Bindings        =   "frmAcompanha.frx":0154
         Height          =   4455
         Left            =   -74880
         TabIndex        =   75
         Top             =   960
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_AcompOcorr"
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "numnf"
            Caption         =   "NF"
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
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            DataField       =   "data"
            Caption         =   "Emissão"
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
            DataField       =   "prioridade"
            Caption         =   "Prioridade"
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
            DataField       =   "prev_entrega"
            Caption         =   "Prev.Entr."
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
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
            DataField       =   "uf_dest"
            Caption         =   "UF"
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
            DataField       =   "modal"
            Caption         =   "Modal"
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
            DataField       =   "transp_sub"
            Caption         =   "T.SubContratado"
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
            DataField       =   "obs_emissao"
            Caption         =   "Observação de Emissão"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
         BeginProperty Column13 
            DataField       =   "emissor"
            Caption         =   "emissor"
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
         BeginProperty Column14 
            DataField       =   "cidade_orig"
            Caption         =   "cidade_orig"
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "especie"
            Caption         =   "especie"
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
         BeginProperty Column19 
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
         BeginProperty Column20 
            DataField       =   "obs_emissao"
            Caption         =   "obs_emissao"
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
         BeginProperty Column21 
            DataField       =   "tem_ocorr"
            Caption         =   "tem_ocorr"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3270,047
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3209,953
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2475,213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   10695,12
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCtcsEmOcorr 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de CTCs EM OCORRÊNCIA: 0 CTC(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74760
         TabIndex        =   79
         Top             =   480
         Width           =   4005
      End
      Begin VB.Label lblNfsEmOcorr 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de NFs EM OCORRÊNCIA: 0 NF(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74760
         TabIndex        =   78
         Top             =   480
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.Label lblAviso2 
         AutoSize        =   -1  'True
         Caption         =   "Opção EM TRÂNSITO não processado ! Escolha acima - Processar:  Em Trânsito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73320
         TabIndex        =   68
         Top             =   2160
         Width           =   8460
      End
      Begin VB.Label lblAviso 
         AutoSize        =   -1  'True
         Caption         =   "Opção ENTREGUE não processado ! Escolha acima - Processar:  Entregue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73080
         TabIndex        =   56
         Top             =   2160
         Width           =   7860
      End
      Begin VB.Label lblCtcsEntregue 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de CTCs ENTREGUES: 0 CTC(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label lblCtcsTransito 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de CTCs EM TRÂNSITO: 0 CTC(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label lblCTCsSemPos 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de CTCs SEM POSIÇÃO: 0 CTC(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   3720
      End
      Begin VB.Label lblNfsSemPos 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de NFs SEM POSIÇÃO: 0 NF(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Label lblNfsBaixa 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de CTCs BAIXADOS (Sem Entrega): 0 CTC(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   4650
      End
      Begin VB.Label lblNfsTransito 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de NFs EM TRÂNSITO: 0 NF(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblNfsEntregue 
         AutoSize        =   -1  'True
         Caption         =   "Listagem de NFs ENTREGUES: 0 NF(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   3360
      End
   End
End
Attribute VB_Name = "frmAcompanha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkProcEmTransito_Click()
    If chkProcEmTransito.Value = 1 Then
        lblAviso2.Visible = False
        If optPorCTC = True Then
            GridTransitoCTC.Visible = True
            lblCtcsTransito.Visible = True
            cmdGerarArqTransitoCtc.Visible = True
            cmdSACTransito.Visible = True
        Else
            GridTransito.Visible = True
            lblNfsTransito.Visible = True
            cmdGerarArqTransito.Visible = True
            cmdSACTransito.Visible = True
        End If
    Else
        lblAviso2.Visible = True
        If optPorCTC = True Then
            GridTransitoCTC.Visible = False
            lblCtcsTransito.Visible = False
            cmdGerarArqTransitoCtc.Visible = False
            cmdSACTransito.Visible = False
        Else
            GridTransito.Visible = False
            lblNfsTransito.Visible = False
            cmdGerarArqTransito.Visible = False
            cmdSACTransito.Visible = False
        End If
    End If

End Sub

Private Sub chkProcEmTransito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub chkProcEmTransito_LostFocus()
    If chkProcEmTransito.Value = 1 Then
        lblAviso2.Visible = False
        GridTransito.Visible = True
        GridTransitoCTC.Visible = True
    Else
        lblAviso2.Visible = True
        GridTransito.Visible = False
        GridTransitoCTC.Visible = False
    End If
End Sub

Private Sub chkProcEntregue_Click()
    If chkProcEntregue.Value = 1 Then
        lblAviso.Visible = False
        If optPorCTC = True Then
            GridEntregueCtc.Visible = True
            lblCtcsEntregue.Visible = True
            cmdGerarArqEntregueCtc.Visible = True
            cmdSACEntregue.Visible = True
        Else
            GridEntregue.Visible = True
            lblNfsEntregue.Visible = True
            cmdGerarArqEntregue.Visible = True
            cmdSACEntregue.Visible = True
        End If
    Else
        lblAviso.Visible = True
        If optPorCTC = True Then
            GridEntregueCtc.Visible = False
            lblCtcsEntregue.Visible = False
            cmdGerarArqEntregueCtc.Visible = False
            cmdSACEntregue.Visible = False
        Else
            GridEntregue.Visible = False
            lblNfsEntregue.Visible = False
            cmdGerarArqEntregue.Visible = False
            cmdSACEntregue.Visible = False
        End If
    End If
End Sub

Private Sub chkProcEntregue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub chkProcEntregue_LostFocus()
    If chkProcEntregue.Value = 1 Then
        lblAviso.Visible = False
        GridEntregue.Visible = True
        GridEntregueCtc.Visible = True
    Else
        lblAviso.Visible = True
        GridEntregue.Visible = False
        GridEntregueCtc.Visible = False
    End If
End Sub

Private Sub chkTodosEstab_Click()
    If chkTodosEstab.Value = 1 Then
        txtCGCRem.MaxLength = 8
    Else
        txtCGCRem.MaxLength = 14
    End If
    txtCGCRem.SetFocus
End Sub

Private Sub chkTodosEstab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdBuscaREM_Click()
    frmBuscaCLI.Caption = "Busca Cliente REMETENTE - (Acompanhamento)"
    frmBuscaCLI.Show 1
End Sub
Private Sub cmdGerarArqEntregue_Click()
    Dim xlinha As String, xrecebedor As String, xdigitado As String
            
    Me.MousePointer = 11
    DoEvents
    
    Open "C:\ENTREGUE.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Data Entrega;Hora Entrega;Recebedor;Transp.Sub;Destinatario;Obs. de Emissao;Entrega Digitada Em;"
    Print #1, xlinha
    Do Until de_informa.rsSel_AcompEntregue.EOF
        If IsNull(de_informa.rsSel_AcompEntregue.Fields("receb")) Then
            xrecebedor = de_informa.rsSel_AcompEntregue.Fields("recebpre")
        Else
            xrecebedor = de_informa.rsSel_AcompEntregue.Fields("receb")
        End If
        xlinha = ""
        xdigitado = CDate(Year(de_informa.rsSel_AcompEntregue.Fields("usu_datapre")) & "/" & _
                    Month(de_informa.rsSel_AcompEntregue.Fields("usu_datapre")) & "/" & _
                    Day(de_informa.rsSel_AcompEntregue.Fields("usu_datapre")))
        xlinha = de_informa.rsSel_AcompEntregue.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompEntregue.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompEntregue.Fields("remet_cgc"), 13, 2) & ";" & _
                de_informa.rsSel_AcompEntregue.Fields("numnf") & ";" & Mid$(de_informa.rsSel_AcompEntregue.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompEntregue.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompEntregue.Fields("data") & ";" & de_informa.rsSel_AcompEntregue.Fields("prioridade") & ";" & de_informa.rsSel_AcompEntregue.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompEntregue.Fields("uf_dest") & ";" & de_informa.rsSel_AcompEntregue.Fields("modal") & ";" & _
                de_informa.rsSel_AcompEntregue.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompEntregue.Fields("dtentr") & ";" & de_informa.rsSel_AcompEntregue.Fields("hsentr") & ";" & _
                xrecebedor & ";" & de_informa.rsSel_AcompEntregue.Fields("transp_sub") & ";" & de_informa.rsSel_AcompEntregue.Fields("dest_nome") & ";" & de_informa.rsSel_AcompEntregue.Fields("obs_emissao") & ";" _
                & xdigitado & ";"
        Print #1, xlinha
        de_informa.rsSel_AcompEntregue.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-ENTREGUE"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo ENTREGUE.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompEntregue.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompEntregue.MoveFirst
            
End Sub

Private Sub cmdGerarArqEntregueCtc_Click()
    Dim xlinha As String, xrecebedor As String, xdigitado As String
            
    Me.MousePointer = 11
    DoEvents
    
    Open "C:\ENTREGUE.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Data Entrega;Hora Entrega;Recebedor;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;Entrega Digitada Em"
    Print #1, xlinha
    Do Until de_informa.rsSel_AcompEntregueCTC.EOF
        If IsNull(de_informa.rsSel_AcompEntregueCTC.Fields("receb")) Then
            xrecebedor = de_informa.rsSel_AcompEntregueCTC.Fields("recebpre")
        Else
            xrecebedor = de_informa.rsSel_AcompEntregueCTC.Fields("receb")
        End If
        xdigitado = CDate(Year(de_informa.rsSel_AcompEntregueCTC.Fields("usu_datapre")) & "/" & _
                    Month(de_informa.rsSel_AcompEntregueCTC.Fields("usu_datapre")) & "/" & _
                    Day(de_informa.rsSel_AcompEntregueCTC.Fields("usu_datapre")))
        xlinha = ""
        xlinha = de_informa.rsSel_AcompEntregueCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompEntregueCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompEntregueCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                Mid$(de_informa.rsSel_AcompEntregueCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompEntregueCTC.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompEntregueCTC.Fields("data") & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("prioridade") & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompEntregueCTC.Fields("uf_dest") & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("modal") & ";" & _
                de_informa.rsSel_AcompEntregueCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("dtentr") & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("hsentr") & ";" & _
                xrecebedor & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("transp_sub") & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("dest_nome") & ";" & de_informa.rsSel_AcompEntregueCTC.Fields("nfs") & ";" & _
                de_informa.rsSel_AcompEntregueCTC.Fields("obs_emissao") & ";" & xdigitado & ";"
        Print #1, xlinha
        de_informa.rsSel_AcompEntregueCTC.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-ENTREGUE"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo ENTREGUE.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompEntregueCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompEntregueCTC.MoveFirst
    
End Sub

Private Sub cmdGerarArqOcorr_Click()
    Dim xlinha As String, xocorr As String
    
    Me.MousePointer = 11
    DoEvents
    
    Open "C:\EM_OCORRENCIA.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Obs. de Emissao;Ocorrências"
    Print #1, xlinha
    Do Until de_informa.rsSel_AcompOcorr.EOF
        'busca ocorrências deste ctc
        xocorr = ""
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
        de_informa.Sel_ConsOcorr2 de_informa.rsSel_AcompOcorr.Fields("filialctc"), "01"
        Do Until de_informa.rsSel_ConsOcorr2.EOF
            xdataocorr = de_informa.rsSel_ConsOcorr2.Fields("data")
            xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
            xocorr = xocorr & "(" & xdataocorr & "-" & Trim$(de_informa.rsSel_ConsOcorr2.Fields("descr_ocorr")) & ") - "
            de_informa.rsSel_ConsOcorr2.MoveNext
        Loop
        xocorr = Mid$(xocorr, 1, Len(xocorr) - 3)
        xlinha = ""
        xlinha = de_informa.rsSel_AcompOcorr.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompOcorr.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompOcorr.Fields("remet_cgc"), 13, 2) & ";" & _
                de_informa.rsSel_AcompOcorr.Fields("numnf") & ";" & Mid$(de_informa.rsSel_AcompOcorr.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompOcorr.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompOcorr.Fields("data") & ";" & de_informa.rsSel_AcompOcorr.Fields("prioridade") & ";" & de_informa.rsSel_AcompOcorr.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompOcorr.Fields("uf_dest") & ";" & de_informa.rsSel_AcompOcorr.Fields("modal") & ";" & _
                de_informa.rsSel_AcompOcorr.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompOcorr.Fields("transp_sub") & ";" & _
                de_informa.rsSel_AcompOcorr.Fields("dest_nome") & ";" & de_informa.rsSel_AcompOcorr.Fields("obs_emissao") & ";" & xocorr
        Print #1, xlinha
        de_informa.rsSel_AcompOcorr.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM OCORRÊNCIA"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo EM_OCORRENCIA.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompOcorr.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompOcorr.MoveFirst
            
End Sub
Private Sub cmdGerarArqOcorrCtc_Click()
    Dim xlinha As String, xocorr As String
    
    Me.MousePointer = 11
    DoEvents
    
    Open "C:\EM_OCORRENCIA.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;Ocorrências"
    Print #1, xlinha
    Do Until de_informa.rsSel_AcompOcorrCTC.EOF
        'busca ocorrências deste ctc
        xocorr = ""
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
        de_informa.Sel_ConsOcorr2 de_informa.rsSel_AcompOcorrCTC.Fields("filialctc"), "01"
        Do Until de_informa.rsSel_ConsOcorr2.EOF
            xdataocorr = de_informa.rsSel_ConsOcorr2.Fields("data")
            xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
            xocorr = xocorr & "(" & xdataocorr & "-" & Trim$(de_informa.rsSel_ConsOcorr2.Fields("descr_ocorr")) & ") - "
            de_informa.rsSel_ConsOcorr2.MoveNext
        Loop
        xocorr = Mid$(xocorr, 1, Len(xocorr) - 3)
        xlinha = ""
        
        xlinha = de_informa.rsSel_AcompOcorrCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompOcorrCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompOcorrCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                Mid$(de_informa.rsSel_AcompOcorrCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompOcorrCTC.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompOcorrCTC.Fields("data") & ";" & de_informa.rsSel_AcompOcorrCTC.Fields("prioridade") & ";" & de_informa.rsSel_AcompOcorrCTC.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompOcorrCTC.Fields("uf_dest") & ";" & de_informa.rsSel_AcompOcorrCTC.Fields("modal") & ";" & _
                de_informa.rsSel_AcompOcorrCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompOcorrCTC.Fields("transp_sub") & ";" & _
                de_informa.rsSel_AcompOcorrCTC.Fields("dest_nome") & ";" & de_informa.rsSel_AcompOcorrCTC.Fields("nfs") & ";" & _
                de_informa.rsSel_AcompOcorrCTC.Fields("obs_emissao") & ";" & xocorr

        Print #1, xlinha
        de_informa.rsSel_AcompOcorrCTC.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM OCORRÊNCIA"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo EM_OCORRENCIA.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompOcorrCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompOcorrCTC.MoveFirst
    
    
    
    
    
    
    
End Sub
Private Sub cmdGerarArqSemPos_Click()
    Dim xlinha As String
            
    Me.MousePointer = 11
    DoEvents
            
    Open "C:\SEMPOSICAO.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Obs. de Emissao;"
    Print #1, xlinha
    Do Until de_informa.rsSel_AcompSemPos.EOF
        xlinha = ""
        xlinha = de_informa.rsSel_AcompSemPos.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompSemPos.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompSemPos.Fields("remet_cgc"), 13, 2) & ";" & _
                de_informa.rsSel_AcompSemPos.Fields("numnf") & ";" & Mid$(de_informa.rsSel_AcompSemPos.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompSemPos.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompSemPos.Fields("data") & ";" & de_informa.rsSel_AcompSemPos.Fields("prioridade") & ";" & de_informa.rsSel_AcompSemPos.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompSemPos.Fields("uf_dest") & ";" & de_informa.rsSel_AcompSemPos.Fields("modal") & ";" & _
                de_informa.rsSel_AcompSemPos.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompSemPos.Fields("transp_sub") & ";" & _
                de_informa.rsSel_AcompSemPos.Fields("dest_nome") & ";" & de_informa.rsSel_AcompSemPos.Fields("obs_emissao") & ";"
        Print #1, xlinha
        de_informa.rsSel_AcompSemPos.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-SEM POSIÇÃO"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo SEMPOSICAO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompSemPos.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompSemPos.MoveFirst
    
End Sub
Private Sub cmdGerarArqSemPosCtc_Click()
    Dim xlinha As String
    
    Me.MousePointer = 11
    DoEvents
    
    Open "C:\SEMPOSICAO.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;"
    Print #1, xlinha
    de_informa.rsSel_AcompSemPosCTC.MoveFirst
    Do Until de_informa.rsSel_AcompSemPosCTC.EOF
        xlinha = ""
        xlinha = de_informa.rsSel_AcompSemPosCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompSemPosCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompSemPosCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                Mid$(de_informa.rsSel_AcompSemPosCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompSemPosCTC.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompSemPosCTC.Fields("data") & ";" & de_informa.rsSel_AcompSemPosCTC.Fields("prioridade") & ";" & de_informa.rsSel_AcompSemPosCTC.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompSemPosCTC.Fields("uf_dest") & ";" & de_informa.rsSel_AcompSemPosCTC.Fields("modal") & ";" & _
                de_informa.rsSel_AcompSemPosCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompSemPosCTC.Fields("transp_sub") & ";" & _
                de_informa.rsSel_AcompSemPosCTC.Fields("dest_nome") & ";" & de_informa.rsSel_AcompSemPosCTC.Fields("nfs") & ";" & _
                de_informa.rsSel_AcompSemPosCTC.Fields("obs_emissao") & ";"
        Print #1, xlinha
        de_informa.rsSel_AcompSemPosCTC.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-SEM POSIÇÃO"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo SEMPOSICAO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompSemPosCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompSemPosCTC.MoveFirst
            
            
End Sub
Private Sub cmdGerarArqTransito_Click()
    Dim xlinha As String

    Me.MousePointer = 11
    DoEvents

    Open "C:\EM_TRANSITO.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Obs. de Emissao;"
    Print #1, xlinha
    Do Until de_informa.rsSel_AcompTransito.EOF
        xlinha = ""
        xlinha = de_informa.rsSel_AcompTransito.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompTransito.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompTransito.Fields("remet_cgc"), 13, 2) & ";" & _
                de_informa.rsSel_AcompTransito.Fields("numnf") & ";" & Mid$(de_informa.rsSel_AcompTransito.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompTransito.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompTransito.Fields("data") & ";" & de_informa.rsSel_AcompTransito.Fields("prioridade") & ";" & de_informa.rsSel_AcompTransito.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompTransito.Fields("uf_dest") & ";" & de_informa.rsSel_AcompTransito.Fields("modal") & ";" & _
                de_informa.rsSel_AcompTransito.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompTransito.Fields("transp_sub") & ";" & _
                de_informa.rsSel_AcompTransito.Fields("dest_nome") & ";" & de_informa.rsSel_AcompTransito.Fields("obs_emissao") & ";"
        Print #1, xlinha
        de_informa.rsSel_AcompTransito.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM TRÂNSITO"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo EM_TRANSITO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompTransito.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompTransito.MoveFirst
    
End Sub
Private Sub cmdGerarArqTransitoCtc_Click()
    Dim xlinha As String
    
    Me.MousePointer = 11
    DoEvents
    
    Open "C:\EM_TRANSITO.TXT" For Output As #1
    'cria cabeçário do arquivo (campos)
    xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;"
    Print #1, xlinha
    Do Until de_informa.rsSel_AcompTransitoCTC.EOF
        xlinha = ""
        xlinha = de_informa.rsSel_AcompTransitoCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_AcompTransitoCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_AcompTransitoCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                Mid$(de_informa.rsSel_AcompTransitoCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_AcompTransitoCTC.Fields("filialctc"), 3, 8) & ";" & _
                de_informa.rsSel_AcompTransitoCTC.Fields("data") & ";" & de_informa.rsSel_AcompTransitoCTC.Fields("prioridade") & ";" & de_informa.rsSel_AcompTransitoCTC.Fields("cidade_dest") & ";" & _
                de_informa.rsSel_AcompTransitoCTC.Fields("uf_dest") & ";" & de_informa.rsSel_AcompTransitoCTC.Fields("modal") & ";" & _
                de_informa.rsSel_AcompTransitoCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_AcompTransitoCTC.Fields("transp_sub") & ";" & _
                de_informa.rsSel_AcompTransitoCTC.Fields("dest_nome") & ";" & de_informa.rsSel_AcompTransitoCTC.Fields("nfs") & ";" & _
                de_informa.rsSel_AcompTransitoCTC.Fields("obs_emissao") & ";"
        Print #1, xlinha
        de_informa.rsSel_AcompTransitoCTC.MoveNext
    Loop
    Close #1
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM TRÂNSITO"
    
    Me.MousePointer = 0
    DoEvents
    
    MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
            "Arquivo EM_TRANSITO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
            "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_AcompTransitoCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
            "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
            "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
            
    de_informa.rsSel_AcompTransitoCTC.MoveFirst
            
End Sub

Private Sub cmdNovaSel_Click()
    If MsgBox("Confirma Nova Seleção de Dados ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        frmAcompanha.MousePointer = 11
        DoEvents
        fraComandos.Enabled = False
        TabGeral.Enabled = False
        If de_informa.rsSel_AcompSemPos.State = 1 Then de_informa.rsSel_AcompSemPos.Close
        If de_informa.rsSel_AcompOcorr.State = 1 Then de_informa.rsSel_AcompOcorr.Close
        If de_informa.rsSel_AcompTransito.State = 1 Then de_informa.rsSel_AcompTransito.Close
        If de_informa.rsSel_AcompEntregue.State = 1 Then de_informa.rsSel_AcompEntregue.Close
        If de_informa.rsSel_AcompBaixa.State = 1 Then de_informa.rsSel_AcompBaixa.Close
        If de_informa.rsSel_AcompSemPosCTC.State = 1 Then de_informa.rsSel_AcompSemPosCTC.Close
        If de_informa.rsSel_AcompOcorrCTC.State = 1 Then de_informa.rsSel_AcompOcorrCTC.Close
        If de_informa.rsSel_AcompTransitoCTC.State = 1 Then de_informa.rsSel_AcompTransitoCTC.Close
        If de_informa.rsSel_AcompEntregueCTC.State = 1 Then de_informa.rsSel_AcompEntregueCTC.Close
        GridSemPos.DataMember = "Sel_AcompSemPos"
        GridSemPos.Refresh
        GridOcorr.DataMember = "Sel_AcompOcorr"
        GridOcorr.Refresh
        GridTransito.DataMember = "Sel_AcompTransito"
        GridTransito.Refresh
        GridEntregue.DataMember = "Sel_AcompEntregue"
        GridEntregue.Refresh
        GridBaixa.DataMember = "Sel_AcompBaixa"
        GridBaixa.Refresh
        GridSemPosCtc.DataMember = "Sel_AcompSemPosCTC"
        GridSemPosCtc.Refresh
        GridOcorrCTC.DataMember = "Sel_AcompOcorrCTC"
        GridOcorrCTC.Refresh
        GridTransitoCTC.DataMember = "Sel_AcompTransitoCTC"
        GridTransitoCTC.Refresh
        GridEntregueCtc.DataMember = "Sel_AcompEntregueCTC"
        GridEntregueCtc.Refresh
        lblNfsSemPos.Caption = "Listagem de NFs SEM POSIÇÃO: 0 NF(s)"
        lblNfsEmOcorr.Caption = "Listagem de NFs EM OCORRÊNCIA: 0 NF(s)"
        lblNfsTransito.Caption = "Listagem de NFs EM TRÂNSITO: 0 NF(s)"
        lblNfsEntregue.Caption = "Listagem de NFs ENTREGUES: 0 NF(s)"
        lblNfsBaixa.Caption = "Listagem de CTCs BAIXADOS (Sem Entrega): 0 CTC(s)"
        lblCTCsSemPos.Caption = "Listagem de CTCs SEM POSIÇÃO: 0 CTC(s)"
        lblCtcsEmOcorr.Caption = "Listagem de CTCs EM OCORRÊNCIA: 0 CTC(s)"
        lblCtcsTransito.Caption = "Listagem de CTCs EM TRÂNSITO: 0 CTC(s)"
        lblCtcsEntregue.Caption = "Listagem de CTCs ENTREGUES: 0 CTC(s)"
        
        DoEvents
        fraRemetente.Enabled = True
        FraPeriodo.Enabled = True
        fraComandos.Enabled = True
        cmdNovaSel.Enabled = False
        If optSelReg = True Then
            txtregiaosac.SetFocus
        ElseIf optSelCli = True Then
            txtCGCRem.SetFocus
        End If
        fraDadosPor.Enabled = True
        cmdGerarArqSemPos.Enabled = False
        cmdGerarArqSemPosCtc.Enabled = False
        cmdGerarArqOcorr.Enabled = False
        cmdGerarArqOcorrCtc.Enabled = False
        cmdGerarArqTransito.Enabled = False
        cmdGerarArqTransitoCtc.Enabled = False
        cmdGerarArqEntregue.Enabled = False
        cmdGerarArqEntregueCtc.Enabled = False
        frmAcompanha.MousePointer = 0
    End If
End Sub
Private Sub cmdProcessa_Click()
    Dim xcgc As String, xcgcdest As String, xdata1 As Date, xdata2 As Date, xregiaosac As String, xprioridade As String
    
    If Len(Trim$(TxtFilial)) = 0 Then
        TxtFilial = "%"
    End If
    
    frmAcompanha.MousePointer = 11
    
    If optPorEmissao.Value = True Then  'por emissao
        If optPer15d.Value = True Then
            xdata1 = datahora("data") - 15
            xdata2 = datahora("data")
        ElseIf opt30d.Value = True Then
            xdata1 = datahora("data") - 30
            xdata2 = datahora("data")
        ElseIf opt60d.Value = True Then
            xdata1 = datahora("data") - 60
            xdata2 = datahora("data")
        Else
            MsgBox "Período Escolhido Inválido !"
            Exit Sub
        End If
    End If
    If optPorMes.Value = True Then   'por mes
        xdata1 = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 "01")
        xdata2 = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 UltDiaMes(Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2)), _
                           Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4))))
                           
        If xdata2 > datahora("DATA") Then xdata2 = datahora("DATA")
        
    End If
    If optPorPeriodo.Value = True Then   'por periodo
        If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
            MsgBox "Período Escolhido Inválido !"
            mskPer1.SetFocus
            frmAcompanha.MousePointer = 0
            Exit Sub
        End If
        If CDate(mskPer1) > CDate(mskPer2) Then
            MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
            mskPer1.SetFocus
            frmAcompanha.MousePointer = 0
            Exit Sub
        End If
        xdata1 = CDate(mskPer1)
        xdata2 = CDate(mskPer2)
        If xdata2 - xdata1 > 32 Then
            MsgBox "Período Escolhido Maior que 30 Dias ! Escolha um Período Menor."
            mskPer1.SetFocus
            frmAcompanha.MousePointer = 0
            Exit Sub
        End If
    End If
    
    
    fraRemetente.Enabled = False
    FraPeriodo.Enabled = False
    fraComandos.Enabled = False
    TabGeral.Enabled = False
    DoEvents

    If optSelReg = True Then
        If Trim$(txtregiaosac) = "" Then txtregiaosac = "%"
        xregiaosac = txtregiaosac
        xcgc = "%"
        xcgcdest = "%"
    Else
        If Trim(txtCGCRem.Text) = "" Then txtCGCRem = "%"
        If optRemetente = True Then
            xcgc = Trim$(txtCGCRem) & "%"
            If xcgc = "%%" Then xcgc = "%"
            xcgcdest = "%"
        Else
            xcgcdest = Trim$(txtCGCRem) & "%"
            If xcgcdest = "%%" Then xcgcdest = "%"
            xcgc = "%"
        End If
        xregiaosac = "%"
    End If
    
'    If optPrioriTodo = True Then
'        xprioridade = "%"
'    ElseIf optPrioriPrioridades = True Then
'        xprioridade = "PRIORIDADE%"
'    ElseIf optPrioriUrgencias = True Then
'        xprioridade = "URGÊNCIA%"
'     End If
        
    DoEvents
        
    'TAB SEM POSIÇÃO
    
    If optPorNF.Value = True Then
        'POR NFS
        If de_informa.rsSel_AcompSemPos.State = 1 Then de_informa.rsSel_AcompSemPos.Close
        de_informa.Sel_AcompSemPos xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
        GridSemPos.DataMember = "Sel_AcompSemPos"
        GridSemPos.Refresh
        lblNfsSemPos.Caption = "Listagem de NFs SEM POSIÇÃO: " & de_informa.rsSel_AcompSemPos.RecordCount & " NF(s)"
        DoEvents
    Else
        'POR CTCS
        If de_informa.rsSel_AcompSemPosCTC.State = 1 Then de_informa.rsSel_AcompSemPosCTC.Close
        de_informa.Sel_AcompSemPosCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
        GridSemPosCtc.DataMember = "Sel_AcompSemPosCTC"
        GridSemPosCtc.Refresh
        lblCTCsSemPos.Caption = "Listagem de CTCs SEM POSIÇÃO: " & de_informa.rsSel_AcompSemPosCTC.RecordCount & " CTC(s)"
        DoEvents
    End If


    'TAB OCORRÊNCIA
    
    If optPorNF.Value = True Then
        'POR NFS
        If de_informa.rsSel_AcompOcorr.State = 1 Then de_informa.rsSel_AcompOcorr.Close
        de_informa.Sel_AcompOcorr xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
        GridOcorr.DataMember = "Sel_AcompOcorr"
        GridOcorr.Refresh
        lblNfsEmOcorr.Caption = "Listagem de NFs EM OCORRÊNCIA: " & de_informa.rsSel_AcompOcorr.RecordCount & " NF(s)"
        DoEvents
    Else
        'POR CTCS
        If de_informa.rsSel_AcompOcorrCTC.State = 1 Then de_informa.rsSel_AcompOcorrCTC.Close
        de_informa.Sel_AcompOcorrCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
        GridOcorrCTC.DataMember = "Sel_AcompOcorrCTC"
        GridOcorrCTC.Refresh
        lblCtcsEmOcorr.Caption = "Listagem de CTCs EM OCORRÊNCIA: " & de_informa.rsSel_AcompOcorrCTC.RecordCount & " CTC(s)"
        DoEvents
    End If
    
    'TAB EM TRÂNSITO
    
    If chkProcEmTransito.Value = 1 Then
        If optPorNF.Value = True Then
            'POR NFS
            If de_informa.rsSel_AcompTransito.State = 1 Then de_informa.rsSel_AcompTransito.Close
            de_informa.Sel_Acomptransito xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
            GridTransito.DataMember = "Sel_AcompTransito"
            GridTransito.Refresh
            lblNfsTransito.Caption = "Listagem de NFs EM TRÂNSITO: " & de_informa.rsSel_AcompTransito.RecordCount & " NF(s)"
            DoEvents
        Else
            'POR CTCS
            If de_informa.rsSel_AcompTransitoCTC.State = 1 Then de_informa.rsSel_AcompTransitoCTC.Close
            de_informa.Sel_AcompTransitoCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
            GridTransitoCTC.DataMember = "Sel_AcompTransitoCTC"
            GridTransitoCTC.Refresh
            lblCtcsTransito.Caption = "Listagem de CTCs EM TRÂNSITO: " & de_informa.rsSel_AcompTransitoCTC.RecordCount & " CTC(s)"
            DoEvents
        End If
    End If

    'TAB ENTREGUE
    
    If chkProcEntregue.Value = 1 Then
        If optPorNF.Value = True Then
            'POR NFS
            de_informa.cn_informa.CommandTimeout = 0
            If de_informa.rsSel_AcompEntregue.State = 1 Then de_informa.rsSel_AcompEntregue.Close
            de_informa.Sel_Acompentregue xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
            GridEntregue.DataMember = "Sel_AcompEntregue"
            GridEntregue.Refresh
            lblNfsEntregue.Caption = "Listagem de NFs ENTREGUES: " & de_informa.rsSel_AcompEntregue.RecordCount & " NF(s)"
            DoEvents
        Else
            'POR CTCS
            If de_informa.rsSel_AcompEntregueCTC.State = 1 Then de_informa.rsSel_AcompEntregueCTC.Close
            de_informa.Sel_AcompEntregueCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
            GridEntregueCtc.DataMember = "Sel_AcompEntregueCTC"
            GridEntregueCtc.Refresh
            lblCtcsEntregue.Caption = "Listagem de CTCs ENTREGUES: " & de_informa.rsSel_AcompEntregueCTC.RecordCount & " CTC(s)"
            DoEvents
        End If
    End If
    
    If optPorCTC.Value = True Then
        If de_informa.rsSel_AcompSemPosCTC.RecordCount > 0 Then
            cmdSACSemPos.Enabled = True
        Else
            cmdSACSemPos.Enabled = False
        End If
    End If
    If optPorCTC.Value = True Then
        If de_informa.rsSel_AcompOcorrCTC.RecordCount > 0 Then
            cmdSACOcorr.Enabled = True
        Else
            cmdSACOcorr.Enabled = False
        End If
    End If
    If chkProcEmTransito.Value = 1 Then
        If optPorCTC.Value = True Then
            If de_informa.rsSel_AcompTransitoCTC.RecordCount > 0 Then
                cmdSACTransito.Enabled = True
            Else
                cmdSACTransito.Enabled = False
            End If
        End If
    End If
    If chkProcEntregue.Value = 1 Then
        If optPorCTC.Value = True Then
            If de_informa.rsSel_AcompEntregueCTC.RecordCount > 0 Then
                cmdSACEntregue.Enabled = True
            Else
                cmdSACEntregue.Enabled = False
            End If
        End If
    End If
    'fraRemetente.Enabled = True
    'FraPeriodo.Enabled = True
    fraComandos.Enabled = True
    cmdNovaSel.Enabled = True
    fraDadosPor.Enabled = False
    If optPorCTC.Value = True Then
        cmdGerarArqSemPosCtc.Enabled = True
        cmdGerarArqOcorrCtc.Enabled = True
        If chkProcEmTransito.Value = 1 Then
            cmdGerarArqTransitoCtc.Enabled = True
        Else
            cmdGerarArqTransitoCtc.Enabled = False
        End If
        If chkProcEntregue.Value = 1 Then
            cmdGerarArqEntregueCtc.Enabled = True
        Else
            cmdGerarArqEntregueCtc.Enabled = False
        End If
    Else
        cmdGerarArqSemPos.Enabled = True
        cmdGerarArqOcorr.Enabled = True
        If chkProcEmTransito.Value = 1 Then
            cmdGerarArqTransito.Enabled = True
        Else
            cmdGerarArqTransito.Enabled = False
        End If
        If chkProcEntregue.Value = 1 Then
            cmdGerarArqEntregue.Enabled = True
        Else
            cmdGerarArqEntregue.Enabled = False
        End If
    End If
    
    TabGeral.Enabled = True
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: " & lblNomeRem
        
    frmAcompanha.MousePointer = 0
    
End Sub

Private Sub cmdProcessa_GotFocus()
    If chkProcEntregue.Value = 1 Then
        If optPorEmissao = True And opt60d = True Then
            MsgBox "Para Processar CTC/NF Entregue, Escolha um Período de Até 30 dias ou um Mês Específico !"
            chkProcEntregue.Value = 0
            Exit Sub
        End If
        If optPorPeriodo = True And IsDate(mskPer2) And IsDate(mskPer1) Then
            If (CDate(mskPer2) - CDate(mskPer1)) > 32 Then
                MsgBox "Para Processar CTC/NF Entregue, Escolha um Período de Até 30 dias ou um Mês Específico !"
                chkProcEntregue.Value = 0
                Exit Sub
            End If
        End If
        lblAviso.Visible = False
        If optPorCTC = True Then
            GridEntregueCtc.Visible = True
            lblCtcsEntregue.Visible = True
            cmdGerarArqEntregueCtc.Visible = True
            cmdSACEntregue.Visible = True
        Else
            GridEntregue.Visible = True
            lblNfsEntregue.Visible = True
            cmdGerarArqEntregue.Visible = True
            cmdSACEntregue.Visible = True
        End If
    End If
End Sub

Private Sub cmdProcessarBxSemEntr_Click()
    
    If Len(Trim$(TxtFilial)) = 0 Then
        TxtFilial = "%"
    End If
    
    frmAcompanha.MousePointer = 11
    
    If optPorEmissao.Value = True Then  'por emissao
        If optPer15d.Value = True Then
            xdata1 = datahora("data") - 15
            xdata2 = datahora("data")
        ElseIf opt30d.Value = True Then
            xdata1 = datahora("data") - 30
            xdata2 = datahora("data")
        ElseIf opt60d.Value = True Then
            xdata1 = datahora("data") - 60
            xdata2 = datahora("data")
        Else
            MsgBox "Período Escolhido Inválido !"
            Exit Sub
        End If
    End If
    If optPorMes.Value = True Then   'por mes
        xdata1 = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 "01")
        xdata2 = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 UltDiaMes(Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2)), _
                           Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4))))
                           
        If xdata2 > datahora("DATA") Then xdata2 = datahora("DATA")
        
    End If
    If optPorPeriodo.Value = True Then   'por periodo
        If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
            MsgBox "Período Escolhido Inválido !"
            mskPer1.SetFocus
            frmAcompanha.MousePointer = 0
            Exit Sub
        End If
        If CDate(mskPer1) > CDate(mskPer2) Then
            MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
            mskPer1.SetFocus
            frmAcompanha.MousePointer = 0
            Exit Sub
        End If
        xdata1 = CDate(mskPer1)
        xdata2 = CDate(mskPer2)
        If xdata2 - xdata1 > 32 Then
            MsgBox "Período Escolhido Maior que 30 Dias ! Escolha um Período Menor."
            mskPer1.SetFocus
            frmAcompanha.MousePointer = 0
            Exit Sub
        End If
    End If
    
    
    fraRemetente.Enabled = False
    FraPeriodo.Enabled = False
    fraComandos.Enabled = False
    TabGeral.Enabled = False
    DoEvents

    If optSelReg = True Then
        If Trim$(txtregiaosac) = "" Then txtregiaosac = "%"
        xregiaosac = txtregiaosac
        xcgc = "%"
        xcgcdest = "%"
    Else
        If Trim(txtCGCRem.Text) = "" Then txtCGCRem = "%"
        If optRemetente = True Then
            xcgc = Trim$(txtCGCRem) & "%"
            If xcgc = "%%" Then xcgc = "%"
            xcgcdest = "%"
        Else
            xcgcdest = Trim$(txtCGCRem) & "%"
            If xcgcdest = "%%" Then xcgcdest = "%"
            xcgc = "%"
        End If
        xregiaosac = "%"
    End If
    'TAB BAIXA SEM ENTREGA
    If de_informa.rsSel_AcompBaixa.State = 1 Then de_informa.rsSel_AcompBaixa.Close
    de_informa.Sel_AcompBaixa xdata1, xdata2, xcgc, xcgcdest, xregiaosac, TxtFilial
    GridBaixa.DataMember = "Sel_AcompBaixa"
    GridBaixa.Refresh
    lblNfsBaixa.Caption = "Listagem de CTCs BAIXADOS (Sem Entrega): " & de_informa.rsSel_AcompBaixa.RecordCount & " CTC(s)"

    If optPorCTC.Value = True Then
        If de_informa.rsSel_AcompBaixa.RecordCount > 0 Then
            cmdSACBxSemEntr.Enabled = True
        Else
            cmdSACBxSemEntr.Enabled = False
        End If
    End If
    
    fraRemetente.Enabled = True
    FraPeriodo.Enabled = True
    fraComandos.Enabled = True
    TabGeral.Enabled = True
    DoEvents
    
    frmAcompanha.MousePointer = 0
    
    

End Sub

Private Sub cmdSACBxSemEntr_Click()
    xultimofilial = Mid(GridBaixa.Columns(0), 1, 2)
    xultimoctc = Mid(GridBaixa.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridBaixa.Columns(0), 1, 2)
    frmSac.txtCtc = Mid(GridBaixa.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSACEntregue_Click()
    xultimofilial = Mid(GridEntregueCtc.Columns(0), 1, 2)
    xultimoctc = Mid(GridEntregueCtc.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridEntregueCtc.Columns(0), 1, 2)
    frmSac.txtCtc = Mid(GridEntregueCtc.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSACOcorr_Click()
    xultimofilial = Mid(GridOcorrCTC.Columns(0), 1, 2)
    xultimoctc = Mid(GridOcorrCTC.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridOcorrCTC.Columns(0), 1, 2)
    frmSac.txtCtc = Mid(GridOcorrCTC.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSACSemPos_Click()
    xultimofilial = Mid(GridSemPosCtc.Columns(0), 1, 2)
    xultimoctc = Mid(GridSemPosCtc.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridSemPosCtc.Columns(0), 1, 2)
    frmSac.txtCtc = Mid(GridSemPosCtc.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSACTransito_Click()
    xultimofilial = Mid(GridTransitoCTC.Columns(0), 1, 2)
    xultimoctc = Mid(GridTransitoCTC.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridTransitoCTC.Columns(0), 1, 2)
    frmSac.txtCtc = Mid(GridTransitoCTC.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSair_Click()
        frmAcompanha.MousePointer = 11
        DoEvents
        fraComandos.Enabled = False
        TabGeral.Enabled = False
        If de_informa.rsSel_AcompSemPos.State = 1 Then de_informa.rsSel_AcompSemPos.Close
        If de_informa.rsSel_AcompOcorr.State = 1 Then de_informa.rsSel_AcompOcorr.Close
        If de_informa.rsSel_AcompTransito.State = 1 Then de_informa.rsSel_AcompTransito.Close
        If de_informa.rsSel_AcompEntregue.State = 1 Then de_informa.rsSel_AcompEntregue.Close
        If de_informa.rsSel_AcompBaixa.State = 1 Then de_informa.rsSel_AcompBaixa.Close
        If de_informa.rsSel_AcompSemPosCTC.State = 1 Then de_informa.rsSel_AcompSemPosCTC.Close
        If de_informa.rsSel_AcompOcorrCTC.State = 1 Then de_informa.rsSel_AcompOcorrCTC.Close
        If de_informa.rsSel_AcompTransitoCTC.State = 1 Then de_informa.rsSel_AcompTransitoCTC.Close
        If de_informa.rsSel_AcompEntregueCTC.State = 1 Then de_informa.rsSel_AcompEntregueCTC.Close
        GridSemPos.DataMember = "Sel_AcompSemPos"
        GridSemPos.Refresh
        GridOcorr.DataMember = "Sel_AcompOcorr"
        GridOcorr.Refresh
        GridTransito.DataMember = "Sel_AcompTransito"
        GridTransito.Refresh
        GridEntregue.DataMember = "Sel_AcompEntregue"
        GridEntregue.Refresh
        GridBaixa.DataMember = "Sel_AcompBaixa"
        GridBaixa.Refresh
        GridSemPosCtc.DataMember = "Sel_AcompSemPosCTC"
        GridSemPosCtc.Refresh
        GridOcorrCTC.DataMember = "Sel_AcompOcorrCTC"
        GridOcorrCTC.Refresh
        GridTransitoCTC.DataMember = "Sel_AcompTransitoCTC"
        GridTransitoCTC.Refresh
        GridEntregueCtc.DataMember = "Sel_AcompEntregueCTC"
        GridEntregueCtc.Refresh
        lblNfsSemPos.Caption = "Listagem de NFs SEM POSIÇÃO: 0 NF(s)"
        lblNfsEmOcorr.Caption = "Listagem de NFs EM OCORRÊNCIA: 0 NF(s)"
        lblNfsTransito.Caption = "Listagem de NFs EM TRÂNSITO: 0 NF(s)"
        lblNfsEntregue.Caption = "Listagem de NFs ENTREGUES: 0 NF(s)"
        lblNfsBaixa.Caption = "Listagem de CTCs BAIXADOS (Sem Entrega): 0 CTC(s)"
        lblCTCsSemPos.Caption = "Listagem de CTCs SEM POSIÇÃO: 0 CTC(s)"
        lblCtcsEmOcorr.Caption = "Listagem de CTCs EM OCORRÊNCIA: 0 CTC(s)"
        lblCtcsTransito.Caption = "Listagem de CTCs EM TRÂNSITO: 0 CTC(s)"
        lblCtcsEntregue.Caption = "Listagem de CTCs ENTREGUES: 0 CTC(s)"
        
        DoEvents
        optRemetente = True
        optSelReg = True
        fraRemetente.Enabled = True
        FraPeriodo.Enabled = True
        fraComandos.Enabled = True
        cmdNovaSel.Enabled = False
        txtCGCRem.Text = ""
        lblNomeRem.Caption = ""
        txtregiaosac = ""
        lblAtendSac = ""
        lblUfs = ""
        txtregiaosac.SetFocus
        cmdGerarArqSemPos.Enabled = False
        cmdGerarArqSemPosCtc.Enabled = False
        cmdGerarArqOcorr.Enabled = False
        cmdGerarArqOcorrCtc.Enabled = False
        cmdGerarArqTransito.Enabled = False
        cmdGerarArqTransitoCtc.Enabled = False
        cmdGerarArqEntregue.Enabled = False
        cmdGerarArqEntregueCtc.Enabled = False
        optPer15d.Value = True
        chkProcEntregue = 0
        frmAcompanha.MousePointer = 0

        Set frmAcompanha = Nothing
        Unload Me
End Sub

Private Sub cmdSelImprSemPos_Click()
    frmSelGerarArqAcomp.Show 1
End Sub



Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    Call combomesano(comboMesAnoAcomp)
    comboMesAnoAcomp.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.StatusBar1.Visible = True
    Set frmAcompanha = Nothing
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


Private Sub optDestinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub Option4_Click()

End Sub

Private Sub optPer15d_Click()
    mskPer1.Mask = ""
    mskPer1.Text = ""
    mskPer1.Mask = "##/##/####"
    mskPer2.Mask = ""
    mskPer2.Text = ""
    mskPer2.Mask = "##/##/####"
End Sub

Private Sub optPer15d_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optPorCTC_Click()
    lblCTCsSemPos.Visible = True
    lblNfsSemPos.Visible = False
    cmdGerarArqSemPosCtc.Visible = True
    cmdGerarArqSemPos.Visible = False
    GridSemPosCtc.Visible = True
    GridSemPos.Visible = False
    cmdSACSemPos.Visible = True
    
    lblCtcsEmOcorr.Visible = True
    lblNfsEmOcorr.Visible = False
    cmdGerarArqOcorrCtc.Visible = True
    cmdGerarArqOcorr.Visible = False
    GridOcorrCTC.Visible = True
    GridOcorr.Visible = False
    cmdSACOcorr.Visible = True
    
    lblCtcsTransito.Visible = True
    lblNfsTransito.Visible = False
    cmdGerarArqTransitoCtc.Visible = True
    cmdGerarArqTransito.Visible = False
    GridTransitoCTC.Visible = True
    GridTransito.Visible = False
    cmdSACTransito.Visible = True
    
    If chkProcEntregue.Value = 1 Then
        lblCtcsEntregue.Visible = True
        lblNfsEntregue.Visible = False
        cmdGerarArqEntregueCtc.Visible = True
        cmdGerarArqEntregue.Visible = False
        GridEntregueCtc.Visible = True
        GridEntregue.Visible = False
        cmdSACEntregue.Visible = True
    End If
    
End Sub

Private Sub optPorCTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optPorEmissao_Click()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If
End Sub

Private Sub optPorEmissao_GotFocus()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If
End Sub

Private Sub optPorMes_Click()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If

End Sub

Private Sub optPorNF_Click()
    lblCTCsSemPos.Visible = False
    lblNfsSemPos.Visible = True
    cmdGerarArqSemPosCtc.Visible = False
    cmdGerarArqSemPos.Visible = True
    GridSemPosCtc.Visible = False
    GridSemPos.Visible = True
    cmdSACSemPos.Visible = False
    
    lblCtcsEmOcorr.Visible = False
    lblNfsEmOcorr.Visible = True
    cmdGerarArqOcorrCtc.Visible = False
    cmdGerarArqOcorr.Visible = True
    GridOcorrCTC.Visible = False
    GridOcorr.Visible = True
    cmdSACOcorr.Visible = False
    
    lblCtcsTransito.Visible = False
    lblNfsTransito.Visible = True
    cmdGerarArqTransitoCtc.Visible = False
    cmdGerarArqTransito.Visible = True
    GridTransitoCTC.Visible = False
    GridTransito.Visible = True
    cmdSACTransito.Visible = False
    
    If chkProcEntregue.Value = 1 Then
        lblCtcsEntregue.Visible = False
        lblNfsEntregue.Visible = True
        cmdGerarArqEntregueCtc.Visible = False
        cmdGerarArqEntregue.Visible = True
        GridEntregueCtc.Visible = False
        GridEntregue.Visible = True
        cmdSACEntregue.Visible = False
    End If

End Sub

Private Sub optPorNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optPorPeriodo_Click()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If

End Sub

Private Sub optRemetente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optSelCli_Click()
    If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub optSelCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optSelCli_LostFocus()
     If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub optSelReg_Click()
    If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub optSelReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optSelReg_LostFocus()
    If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub SSTab5_DblClick()

End Sub

Private Sub txtCGCRem_GotFocus()
    txtCGCRem.SelStart = 0
    txtCGCRem.SelLength = 14
End Sub

Private Sub txtCGCRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtCGCRem_LostFocus()
    If txtCGCRem.Text = "%" Then
        lblNomeRem.Caption = "TODOS CLIENTES"
        Exit Sub
    End If
    If Len(Trim$(txtCGCRem.Text)) <> txtCGCRem.MaxLength And Len(Trim$(txtCGCRem.Text)) > 0 Then
        MsgBox "Quantidade de Caracteres Inválida para este Número de CGC !"
        txtCGCRem.SetFocus
        SendKeys "{END}"
    End If
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

Private Sub TxtFilial_GotFocus()
    TxtFilial.SelStart = 0
    TxtFilial.SelLength = 2
End Sub

Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtregiaosac_GotFocus()
    txtregiaosac.SelStart = 0
    txtregiaosac.SelLength = 2
    If txtregiaosac.ToolTipText = "" Then
        Dim xtip As String
        If de_informa.rsSel_RegiaoAtendGroup.State = 1 Then de_informa.rsSel_RegiaoAtendGroup.Close
        de_informa.Sel_RegiaoAtendGroup
        If de_informa.rsSel_RegiaoAtendGroup.RecordCount > 0 Then
            xtip = ""
            Do Until de_informa.rsSel_RegiaoAtendGroup.EOF
                xtip = xtip + de_informa.rsSel_RegiaoAtendGroup.Fields("regiaosac") & "-" & _
                de_informa.rsSel_RegiaoAtendGroup.Fields("atendsac") & " ; "
                de_informa.rsSel_RegiaoAtendGroup.MoveNext
            Loop
            txtregiaosac.ToolTipText = xtip
            lblAtendSac.ToolTipText = xtip
        End If
    End If
    If Trim$(txtregiaosac) = "%" Then txtregiaosac = ""
End Sub

Private Sub txtregiaosac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtregiaosac_LostFocus()
    If Trim$(Len(txtregiaosac.Text)) > 0 Then
        If Len(Trim$(txtregiaosac.Text)) = 1 And Trim$(txtregiaosac.Text) <> "%" Then txtregiaosac = "0" & Trim$(txtregiaosac)
        If de_informa.rsSel_BuscaRegiaoSac.State = 1 Then de_informa.rsSel_BuscaRegiaoSac.Close
        de_informa.Sel_BuscaRegiaoSac txtregiaosac
        If de_informa.rsSel_BuscaRegiaoSac.RecordCount > 0 Then
            lblAtendSac = de_informa.rsSel_BuscaRegiaoSac.Fields("atendsac")
            lblUfs = ""
            Do Until de_informa.rsSel_BuscaRegiaoSac.EOF
                lblUfs = lblUfs & de_informa.rsSel_BuscaRegiaoSac.Fields("uf") & "-"
                de_informa.rsSel_BuscaRegiaoSac.MoveNext
            Loop
        Else
            txtregiaosac.SetFocus
            If de_informa.rsSel_UF.State = 1 Then de_informa.rsSel_UF.Close
            de_informa.Sel_UF Trim$(txtregiaosac)
            If de_informa.rsSel_UF.RecordCount > 0 Then
                txtregiaosac = de_informa.rsSel_UF.Fields("regiaosac")
                SendKeys "{TAB}"
            Else
                lblAtendSac = ""
                lblUfs = ""
            End If
        End If
    Else
        lblUfs = ""
        lblAtendSac = ""
    End If
End Sub
