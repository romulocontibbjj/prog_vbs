VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsultaFatura 
   Caption         =   "Consulta Faturas Emitidas (Alteração)"
   ClientHeight    =   7920
   ClientLeft      =   420
   ClientTop       =   750
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Height          =   945
      Left            =   5520
      TabIndex        =   17
      Top             =   120
      Width           =   6255
      Begin VB.Frame Frame8 
         Caption         =   "Status"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   120
         Width           =   2415
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Caption         =   "STATUS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   2205
         End
      End
      Begin VB.Label lblPagamento 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   65
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Pagamento:"
         Height          =   195
         Left            =   2640
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame Frame5 
      Height          =   945
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdListaFat 
         Caption         =   "?"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtFatura 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   345
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Dados da Fatura"
      TabPicture(0)   =   "frmConsultaFatura.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDadosCob"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "CTCs/NFs da Fatura"
      TabPicture(1)   =   "frmConsultaFatura.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Histórico"
      TabPicture(2)   =   "frmConsultaFatura.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraHistoricoCmds"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(2)=   "fraHistorico"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraHistoricoCmds 
         Height          =   1215
         Left            =   -66960
         TabIndex        =   87
         Top             =   480
         Width           =   3495
         Begin VB.CommandButton cmdCancelarHist 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   90
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdGravarHist 
            Caption         =   "Gravar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   89
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdIncluirHist 
            Caption         =   "Incluir Nova Posição de Cobrança"
            Height          =   375
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74880
         TabIndex        =   80
         Top             =   1800
         Width           =   11415
         Begin MSDataGridLib.DataGrid gridHistorico 
            Bindings        =   "frmConsultaFatura.frx":0054
            Height          =   4335
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   7646
            _Version        =   393216
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
            DataMember      =   "Sel_FaturaHistorico"
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "id"
               Caption         =   "ID"
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
               DataField       =   "filialfatura"
               Caption         =   "Filial-Fatura"
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
               Caption         =   "Data"
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
               DataField       =   "usuario"
               Caption         =   "Usuário"
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
               DataField       =   "tipohist"
               Caption         =   "Tipo Hist."
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
               DataField       =   "historico"
               Caption         =   "Histórico"
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
               DataField       =   "contato"
               Caption         =   "Contato"
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
               DataField       =   "retornar"
               Caption         =   "Retornar"
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
               DataField       =   "data_retorno"
               Caption         =   "Data Retorno"
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
                  ColumnWidth     =   675,213
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1709,858
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1065,26
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   7529,953
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1244,976
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   734,74
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1080
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fraHistorico 
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
         Height          =   1215
         Left            =   -74880
         TabIndex        =   79
         Top             =   480
         Width           =   7815
         Begin VB.CheckBox chkAgendar 
            Caption         =   "Retorno/Agendar Para: "
            Height          =   255
            Left            =   4080
            TabIndex        =   85
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtContatoHistorico 
            Height          =   285
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   83
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txtHistorico 
            Height          =   495
            Left            =   120
            MaxLength       =   320
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   240
            Width           =   7575
         End
         Begin MSMask.MaskEdBox mskAgendar 
            Height          =   285
            Left            =   6360
            TabIndex        =   86
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label27 
            Caption         =   "Contato Com:"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   975
         Left            =   120
         TabIndex        =   56
         Top             =   5400
         Width           =   11415
         Begin VB.CommandButton Command5 
            Enabled         =   0   'False
            Height          =   375
            Left            =   9000
            TabIndex        =   12
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdQuita 
            Caption         =   "Quitar Fatura"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6240
            TabIndex        =   11
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdProrroga 
            Caption         =   "Prorrogar Vencimento"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3480
            TabIndex        =   10
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdDesconto 
            Caption         =   "Conceder Desconto/Abatimento"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame fraDadosCob 
         Caption         =   "Endereço de Cobrança"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   120
         TabIndex        =   45
         Top             =   3800
         Width           =   11415
         Begin VB.TextBox txtFoneCob 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   51
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtContatoCob 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3600
            MaxLength       =   20
            TabIndex        =   50
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CommandButton cmdCancCob 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   8880
            TabIndex        =   8
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmdGravarCob 
            Caption         =   "Gravar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   6600
            TabIndex        =   7
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmdAlterarCob 
            Caption         =   "Alterar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   6600
            TabIndex        =   5
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdEndPadraoCob 
            Caption         =   "End. Padrão"
            Enabled         =   0   'False
            Height          =   350
            Left            =   8880
            TabIndex        =   6
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtEndCob 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   49
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox txtCidCob 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   48
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtUFCob 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5400
            MaxLength       =   2
            TabIndex        =   47
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtCepCob 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   8
            TabIndex        =   46
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
            Height          =   195
            Left            =   2880
            TabIndex        =   55
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cep/Cidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   900
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados da Fatura"
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
         Left            =   6600
         TabIndex        =   37
         Top             =   480
         Width           =   4935
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Venc Original:"
            Height          =   195
            Left            =   2520
            TabIndex        =   78
            Top             =   2880
            Width           =   990
         End
         Begin VB.Label lblVencOrig 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3600
            TabIndex        =   77
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lblEnd 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   76
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label lblCidade 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   75
            Top             =   1440
            Width           =   3255
         End
         Begin VB.Label lblUf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4440
            TabIndex        =   74
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblCep 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   73
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblPreFatAvulsa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   67
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Pré-Fatura/Av:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   2880
            Width           =   1050
         End
         Begin VB.Label lblEmissor 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3600
            TabIndex        =   63
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Emissor:"
            Height          =   195
            Left            =   2880
            TabIndex        =   62
            Top             =   2520
            Width           =   585
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   1800
            Width           =   360
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cliente-Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   990
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cidade/UF:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Cliente-CNPJ:"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblClienteNome 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   43
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   2160
            Width           =   510
         End
         Begin VB.Label lblBanco 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Top             =   2160
            Width           =   3615
         End
         Begin VB.Label lblConta 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   40
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Conta Cobr:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   2520
            Width           =   840
         End
         Begin VB.Label lblClienteCNPJ 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   38
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Valores da Fatura"
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
         TabIndex        =   24
         Top             =   480
         Width           =   6375
         Begin VB.Label lblObsAcres 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3050
            TabIndex        =   72
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label lblObsAbat 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3050
            TabIndex        =   71
            Top             =   1710
            Width           =   3255
         End
         Begin VB.Label lblObsFatura 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   70
            Top             =   2805
            Width           =   5700
         End
         Begin VB.Label lblAbat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   69
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblAcrescimo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   68
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "OBS:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   2835
            Width           =   375
         End
         Begin VB.Label lblTipoAbat 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3050
            TabIndex        =   35
            Top             =   1440
            Width           =   3255
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "( - )  Abatimento:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "( + )  Acréscimos:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "( = )  Valor da Fatura:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label lblValorFatura 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   31
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblValorFaturaBruto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   30
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto da Fatura:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   1545
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bruto com ICMS:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1605
         End
         Begin VB.Label lblValorFaturaBrutoICMS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "( - ) Desc. ICMS:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label lblValorICMS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   25
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CTCs da Fatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   11415
         Begin MSDataGridLib.DataGrid GridFaturaItens 
            Bindings        =   "frmConsultaFatura.frx":006D
            Height          =   5535
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   9763
            _Version        =   393216
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
            DataMember      =   "Sel_CTCsDaFatura"
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "filialfatura"
               Caption         =   "Filial-Fatura"
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
               DataField       =   "tipodoc"
               Caption         =   "Doc"
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
            BeginProperty Column03 
               DataField       =   "data"
               Caption         =   "Data CTC"
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
               DataField       =   "frete"
               Caption         =   "Frete Cobr."
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
               DataField       =   "fretebruto"
               Caption         =   "Frete Bruto"
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
               DataField       =   "obs"
               Caption         =   "Observação"
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
            BeginProperty Column08 
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   945,071
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   404,787
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1049,953
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2459,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   2594,835
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   3044,977
               EndProperty
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmConsultaFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAgendar_Click()
    If chkAgendar.Value = 1 And fraHistorico.Enabled = True Then
        mskAgendar.Enabled = True
        mskAgendar.BackColor = xamarelo1
        mskAgendar.SetFocus
    Else
        mskAgendar.Enabled = False
        mskAgendar.BackColor = xbranco
        mskAgendar.Mask = ""
        mskAgendar.Text = ""
        mskAgendar.Mask = "##/##/####"
    End If
End Sub

Private Sub chkAgendar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdAlterarCob_Click()
    cmdAlterarCob.Enabled = False
    cmdGravarCob.Enabled = True
    cmdCancCob.Enabled = True
    cmdEndPadraoCob.Enabled = True
    txtEndCob.Enabled = True
    txtEndCob.BackColor = xamarelo1
    txtCepCob.Enabled = True
    txtCepCob.BackColor = xamarelo1
    txtCidCob.Enabled = True
    txtCidCob.BackColor = xamarelo1
    txtUFCob.Enabled = True
    txtUFCob.BackColor = xamarelo1
    txtFoneCob.Enabled = True
    txtFoneCob.BackColor = xamarelo1
    txtContatoCob.Enabled = True
    txtContatoCob.BackColor = xamarelo1
    txtEndCob.SetFocus
End Sub

Private Sub cmdCancCob_Click()
    cmdAlterarCob.Enabled = True
    cmdGravarCob.Enabled = False
    cmdCancCob.Enabled = False
    cmdEndPadraoCob.Enabled = False
    txtEndCob.Enabled = False
    txtEndCob.BackColor = xbranco
    txtCepCob.Enabled = False
    txtCepCob.BackColor = xbranco
    txtCidCob.Enabled = False
    txtCidCob.BackColor = xbranco
    txtUFCob.Enabled = False
    txtUFCob.BackColor = xbranco
    txtFoneCob.Enabled = False
    txtFoneCob.BackColor = xbranco
    txtContatoCob.Enabled = False
    txtContatoCob.BackColor = xbranco
    cmdBuscar_Click
End Sub
Private Sub cmdCancelarHist_Click()
    If de_informa.rsSel_FaturaHistorico.State = 1 Then de_informa.rsSel_FaturaHistorico.Close
    de_informa.Sel_FaturaHistorico TransFatur(txtFilial, txtFatura)
    gridHistorico.DataMember = "Sel_FaturaHistorico"
    gridHistorico.Refresh
    
    txtHistorico.BackColor = xbranco
    txtContatoHistorico.BackColor = xbranco
    mskAgendar.BackColor = xbranco
    
    txtHistorico.Text = ""
    txtContatoHistorico.Text = ""
    chkAgendar.Value = 0
    mskAgendar.Mask = ""
    mskAgendar.Text = ""
    mskAgendar.Mask = "##/##/####"
    
    cmdIncluirHist.Enabled = True
    cmdGravarHist.Enabled = False
    cmdCancelarHist.Enabled = False
    fraHistorico.Enabled = False
    
End Sub

Private Sub cmdDesconto_Click()
    frmDescontos.lblValorFaturaBrutoICMS = lblValorFaturaBrutoICMS
    frmDescontos.lblValorICMS = lblValorICMS
    frmDescontos.lblValorFaturaBruto = lblValorFaturaBruto
    frmDescontos.txtAbat = lblAbat
    frmDescontos.lblTipoAbat = lblTipoAbat
    frmDescontos.txtObsAbat = lblObsAbat
    frmDescontos.lblAcrescimo = lblAcrescimo
    frmDescontos.lblObsAcres = lblObsAcres
    frmDescontos.lblValorFatura = lblValorFatura
    frmDescontos.lblFilialFatura = TransFatur(txtFilial, txtFatura)
    frmDescontos.Show 1
    cmdBuscar_Click
    DoEvents
End Sub

Private Sub cmdEndPadraoCob_Click()
    If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
    de_informa.Sel_CadCliCGC Trim$(lblClienteCNPJ)
    txtEndCob = de_informa.rsSel_CadCliCGC.Fields("endereco")
    txtCepCob = de_informa.rsSel_CadCliCGC.Fields("cep")
    txtCidCob = de_informa.rsSel_CadCliCGC.Fields("cidade")
    txtUFCob = de_informa.rsSel_CadCliCGC.Fields("uf")
    txtEndCob.SetFocus
End Sub

Private Sub cmdGravarCob_Click()
    
    de_informa.Alt_CadCliDadosCob Trim$(txtEndCob), Trim$(txtCepCob), Trim$(txtCidCob), Trim$(txtUFCob), _
                                  Trim$(txtFoneCob), Trim$(txtContatoCob), zeros2(lblClienteCNPJ, 14)
                                  
    de_informa.Alt_DadosCobFatura Trim$(txtEndCob), Trim$(txtCepCob), Trim$(txtCidCob), Trim$(txtUFCob), _
                                  Trim$(txtFoneCob), Trim$(txtContatoCob), TransFatur(txtFilial, txtFatura)
                                      
    MsgBox "Dados Gravados !", vbInformation, "Gravação"

    cmdAlterarCob.Enabled = True
    cmdGravarCob.Enabled = False
    cmdCancCob.Enabled = False
    cmdEndPadraoCob.Enabled = False
    txtEndCob.Enabled = False
    txtEndCob.BackColor = xbranco
    txtCepCob.Enabled = False
    txtCepCob.BackColor = xbranco
    txtCidCob.Enabled = False
    txtCidCob.BackColor = xbranco
    txtUFCob.Enabled = False
    txtUFCob.BackColor = xbranco
    txtFoneCob.Enabled = False
    txtFoneCob.BackColor = xbranco
    txtContatoCob.Enabled = False
    txtContatoCob.BackColor = xbranco
    cmdBuscar_Click

End Sub

Private Sub cmdGravarHist_Click()

    If chkAgendar.Value = 1 Then
        If Not IsDate(mskAgendar) Then
            MsgBox "Data Inválida para Retorno / Agenda !", vbInformation
            mskAgendar.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(Trim$(txtHistorico)) < 3 Then
        MsgBox "Histórico Inválido ! ", vbInformation
        txtHistorico.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtContatoHistorico)) < 3 Then
        MsgBox "Contato Inválido ! ", vbInformation
        txtContatoHistorico.SetFocus
        Exit Sub
    End If
     
    If MsgBox("Você Confirma o Lançamento Desta Posição de Cobrança/Histórico ?", vbYesNo + vbQuestion, "Histórico") = vbYes Then
        If chkAgendar.Value = 1 Then
            de_informa.Ins_FaturaHistoricoRetorno TransFatur(txtFilial, txtFatura), xusuario, "COBRANCA", Trim$(txtHistorico), Trim$(txtContatoHistorico), CDate(mskAgendar)
        Else
            de_informa.Ins_FaturaHistorico TransFatur(txtFilial, txtFatura), xusuario, "COBRANCA", Trim$(txtHistorico), Trim$(txtContatoHistorico)
        End If
    End If
    
    cmdCancelarHist_Click
    
End Sub

Private Sub cmdIncluirHist_Click()
    txtHistorico.BackColor = xamarelo1
    txtContatoHistorico.BackColor = xamarelo1
    fraHistorico.Enabled = True
    cmdIncluirHist.Enabled = False
    cmdGravarHist.Enabled = True
    cmdCancelarHist.Enabled = True
    txtHistorico.SetFocus
End Sub

Private Sub cmdProrroga_Click()
    frmProrroga.lblEmissao = lblEmissao
    frmProrroga.lblVencto = lblVencto
    frmProrroga.lblFilialFatura = TransFatur(txtFilial, txtFatura)
    frmProrroga.Show 1
    cmdBuscar_Click
    DoEvents
End Sub

Private Sub cmdQuita_Click()
    frmQuitacao.lblEmissao = lblEmissao
    frmQuitacao.lblVencto = lblVencto
    frmQuitacao.lblValorFatura = lblValorFatura
    frmQuitacao.txtAcrescimo = 0
    frmQuitacao.lblValorRecebido = lblValorFatura
    frmQuitacao.lblFilialFatura = TransFatur(txtFilial, txtFatura)
    frmQuitacao.Show 1
    cmdBuscar_Click
    DoEvents
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub cmdBuscar_Click()
    Dim xFilialFatura As String
    
    xFilialFatura = TransFatur(txtFilial, txtFatura)
    
    cmdAlterarCob.Enabled = False
    cmdGravarCob.Enabled = False
    cmdEndPadraoCob.Enabled = False
    cmdCancCob.Enabled = False
    cmdProrroga.Enabled = False
    cmdDesconto.Enabled = False
    cmdQuita.Enabled = False
    
    lblEmissao = ""
    lblVencto = ""
    lblPagamento = ""

    lblValorFaturaBrutoICMS = ""
    lblValorICMS = ""
    lblValorFaturaBruto = ""
    lblAbat = ""
    lblTipoAbat = ""
    lblObsAbat = ""
    lblAcrescimo = ""
    lblObsAcres = ""
    lblValorFatura = ""
    lblObsFatura = ""
    
    lblClienteCNPJ = ""
    lblClienteNome = ""
    lblEnd = ""
    lblCidade = ""
    lblUf = ""
    lblCep = ""
    lblBanco = ""
    lblConta = ""
    lblEmissor = ""
    lblPreFatAvulsa = ""
    lblVencOrig.ToolTipText = ""
    txtEndCob = ""
    txtCidCob = ""
    txtCepCob = ""
    txtUFCob = ""
    txtFoneCob = ""
    txtContatoCob = ""
    
    lblStatus = ""
        
    If de_informa.rsSel_CTCsDaFatura.State = 1 Then de_informa.rsSel_CTCsDaFatura.Close
    GridFaturaItens.DataMember = "Sel_CTCsDaFatura"
    GridFaturaItens.Refresh
    
    If de_informa.rsSel_NFsdaFatura.State = 1 Then de_informa.rsSel_NFsdaFatura.Close
    
    If de_informa.rsSel_FaturaHistorico.State = 1 Then de_informa.rsSel_FaturaHistorico.Close
    gridHistorico.DataMember = "Sel_FaturaHistorico"
    gridHistorico.Refresh
    
    If de_informa.rsSel_Fatura.State = 1 Then de_informa.rsSel_Fatura.Close
    de_informa.Sel_Fatura xFilialFatura
    
    If de_informa.rsSel_Fatura.RecordCount > 0 Then
    
        lblEmissao = de_informa.rsSel_Fatura.Fields("emissao")
        lblVencto = de_informa.rsSel_Fatura.Fields("vencimento")
        If IsNull(de_informa.rsSel_Fatura.Fields("pagamento")) Then
            lblPagamento = ""
        Else
            lblPagamento = de_informa.rsSel_Fatura.Fields("pagamento")
        End If
    
        lblValorFaturaBrutoICMS = Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "##,###,##0.00")
        lblValorICMS = Format(de_informa.rsSel_Fatura.Fields("descicms"), "##,###,##0.00")
        lblValorFaturaBruto = Format(de_informa.rsSel_Fatura.Fields("valorbruto"), "##,###,##0.00")
        lblAbat = Format(de_informa.rsSel_Fatura.Fields("abatimento"), "##,###,##0.00")
        lblTipoAbat = de_informa.rsSel_Fatura.Fields("tipoabat")
        lblObsAbat = de_informa.rsSel_Fatura.Fields("obsabat")
        lblAcrescimo = Format(de_informa.rsSel_Fatura.Fields("acrescimo"), "##,###,##0.00")
        lblObsAcres = de_informa.rsSel_Fatura.Fields("obsacres")
        lblValorFatura = Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00")
        lblObsFatura = de_informa.rsSel_Fatura.Fields("obsfatura")
        
        lblClienteCNPJ = de_informa.rsSel_Fatura.Fields("cliente_cgc")
        lblClienteNome = de_informa.rsSel_Fatura.Fields("cliente_nome")
        lblEnd = de_informa.rsSel_Fatura.Fields("cliente_end")
        lblCidade = de_informa.rsSel_Fatura.Fields("cliente_cidade")
        lblUf = de_informa.rsSel_Fatura.Fields("cliente_uf")
        lblCep = de_informa.rsSel_Fatura.Fields("cliente_cep")
        lblBanco = de_informa.rsSel_Fatura.Fields("banconome")
        lblConta = de_informa.rsSel_Fatura.Fields("conta")
        lblEmissor = de_informa.rsSel_Fatura.Fields("emissor")
        lblPreFatAvulsa = de_informa.rsSel_Fatura.Fields("avulsa")
        lblVencOrig = de_informa.rsSel_Fatura.Fields("venc_orig")
        
        If de_informa.rsSel_Fatura.Fields("venc_orig") <> de_informa.rsSel_Fatura.Fields("vencimento") Then
            lblVencOrig.ToolTipText = "Usuário: " & de_informa.rsSel_Fatura.Fields("prorrog_usu") & "  " & _
                                      "Em: " & de_informa.rsSel_Fatura.Fields("prorrog_data") & "  " & _
                                      "OBS: " & de_informa.rsSel_Fatura.Fields("prorrog_obs")
        Else
            lblVencOrig.ToolTipText = ""
        End If
        
        txtEndCob = de_informa.rsSel_Fatura.Fields("endcob")
        txtCidCob = de_informa.rsSel_Fatura.Fields("cidadecob")
        txtCepCob = de_informa.rsSel_Fatura.Fields("cepcob")
        txtUFCob = de_informa.rsSel_Fatura.Fields("ufcob")
        txtFoneCob = de_informa.rsSel_Fatura.Fields("telefonecob")
        txtContatoCob = de_informa.rsSel_Fatura.Fields("contatocob")
        
        
        If de_informa.rsSel_Fatura.Fields("status") = "C" Then
            lblStatus = "CANCELADO"
            lblStatus.ToolTipText = "Em: " & de_informa.rsSel_Fatura.Fields("canc_data") & _
                                    " Por " & de_informa.rsSel_Fatura.Fields("canc_usu") & _
                                    ". Obs: " & de_informa.rsSel_Fatura.Fields("canc_obs")
        ElseIf de_informa.rsSel_Fatura.Fields("status") = "N" Then
            lblStatus = "EM ABERTO"
            cmdAlterarCob.Enabled = True
            cmdProrroga.Enabled = True
            cmdDesconto.Enabled = True
            cmdQuita.Enabled = True
        ElseIf de_informa.rsSel_Fatura.Fields("status") = "Q" Then
            lblStatus = "QUITADO"
             lblStatus.ToolTipText = "Em: " & de_informa.rsSel_Fatura.Fields("pag_data") & _
                                    " Por " & de_informa.rsSel_Fatura.Fields("pag_usu") & _
                                    ". Obs: " & de_informa.rsSel_Fatura.Fields("pag_obs")
        End If
        
        If de_informa.rsSel_CTCsDaFatura.State = 1 Then de_informa.rsSel_CTCsDaFatura.Close
        de_informa.Sel_CTCsDaFatura xFilialFatura
        GridFaturaItens.DataMember = "Sel_CTCsDaFatura"
        GridFaturaItens.Refresh
        
        If de_informa.rsSel_CTCsDaFatura.RecordCount < 1 Then
            If de_informa.rsSel_NFsdaFatura.State = 1 Then de_informa.rsSel_NFsdaFatura.Close
            de_informa.Sel_NFsdaFatura xFilialFatura
            GridFaturaItens.DataMember = "Sel_NFsdaFatura"
            GridFaturaItens.Refresh
        End If
        
        If de_informa.rsSel_FaturaHistorico.State = 1 Then de_informa.rsSel_FaturaHistorico.Close
        de_informa.Sel_FaturaHistorico xFilialFatura
        gridHistorico.DataMember = "Sel_FaturaHistorico"
        gridHistorico.Refresh
        
    Else
        
        MsgBox "Fatura Inexistente !", vbCritical, "Erro"
    
    End If
    
    txtFilial.SetFocus


End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    
    '73
    contador = 73
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmConsultaFatura.cmdAlterarCob.Enabled = False
    End If
    '74
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmConsultaFatura.cmdDesconto.Enabled = False
    End If
    '75
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmConsultaFatura.cmdProrroga.Enabled = False
    End If
                    
    '76
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmConsultaFatura.cmdQuita.Enabled = False
    End If
    
    mdiFatura.ToolFaturamento.Visible = False
    If de_informa.rsSel_CTCsDaFatura.State = 1 Then de_informa.rsSel_CTCsDaFatura.Close
    GridFaturaItens.DataMember = "Sel_CTCsDaFatura"
    GridFaturaItens.Refresh
    
    If de_informa.rsSel_FaturaHistorico.State = 1 Then de_informa.rsSel_FaturaHistorico.Close
    gridHistorico.DataMember = "Sel_FaturaHistorico"
    gridHistorico.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFatura.ToolFaturamento.Visible = True
End Sub

Private Sub gridHistorico_Click()
    txtHistorico = gridHistorico.Columns(5)
    txtContatoHistorico = gridHistorico.Columns(6)
    If gridHistorico.Columns(7) = "S" Then
        chkAgendar.Value = 1
    Else
        chkAgendar.Value = 0
    End If
    If IsDate(gridHistorico.Columns(8)) Then
        mskAgendar.Mask = ""
        mskAgendar.Text = gridHistorico.Columns(8)
        mskAgendar.Mask = "##/##/####"
    Else
        mskAgendar.Mask = ""
        mskAgendar.Text = ""
        mskAgendar.Mask = "##/##/####"
    End If
End Sub

Private Sub mskAgendar_GotFocus()
    mskAgendar.SelStart = 0
    mskAgendar.SelLength = 10
End Sub

Private Sub mskAgendar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub mskAgendar_LostFocus()
    If mskAgendar.Text <> "__/__/____" Then
        mskAgendar.Text = century(mskAgendar.Text)
        If IsDate(mskAgendar.Text) = False Or Mid(mskAgendar.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskAgendar.SetFocus
            Exit Sub
        End If
        If CDate(mskAgendar.Text) <= datahora("data") Then
            MsgBox "ATENÇÃO ! Confira a Data de Vencimento. Data Menor/Igual à Hoje ???", vbCritical, "Erro"
        End If
    End If
End Sub

Private Sub txtContatoHistorico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtContatoHistorico_LostFocus()
    txtContatoHistorico = UCase(txtContatoHistorico)
End Sub

Private Sub txtFatura_Change()
    If Len(Trim$(txtFatura)) > 0 Then
        If Not IsNumeric(txtFatura) Or Mid$(txtFatura, Len(txtFatura), 1) = "," Or Mid$(txtFatura, Len(txtFatura), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
End Sub
Private Sub txtFatura_GotFocus()
    txtFatura.SelStart = 0
    txtFatura.SelLength = 8
End Sub
Private Sub txtFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFatura)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtFilial_Change()
    If Len(Trim$(txtFilial)) > 0 Then
        If Not IsNumeric(txtFilial) Or Mid$(txtFilial, Len(txtFilial), 1) = "," Or Mid$(txtFilial, Len(txtFilial), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
    If Len(Trim$(txtFilial)) = 2 Then
        txtFatura.SetFocus
    End If
End Sub
Private Sub txtFilial_GotFocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub
Private Sub txtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFilial)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtFilial_LostFocus()
    If Len(Trim$(txtFilial)) = 1 Then
        txtFilial = "0" & Trim$(txtFilial)
    End If
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtHistorico_LostFocus()
    txtHistorico = UCase(txtHistorico)
End Sub
