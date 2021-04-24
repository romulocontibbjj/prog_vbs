VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSACNovo 
   Caption         =   "SAC NOVO"
   ClientHeight    =   8415
   ClientLeft      =   1215
   ClientTop       =   675
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   9465
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Ocorrências e Entrega"
      TabPicture(0)   =   "frmSACNovo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMensagem"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraOcorrencias"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame25"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame21"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraNfsOcorr"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Detalhe da Entrega"
      TabPicture(1)   =   "frmSACNovo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame19"
      Tab(1).Control(2)=   "Frame20"
      Tab(1).Control(3)=   "Frame22"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Dados do CTC"
      TabPicture(2)   =   "frmSACNovo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame27"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "Frame26"
      Tab(2).Control(3)=   "Frame7"
      Tab(2).Control(4)=   "Frame5"
      Tab(2).Control(5)=   "Frame4"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Obs SITLA e Manifesto"
      TabPicture(3)   =   "frmSACNovo.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame17"
      Tab(3).Control(2)=   "Frame14"
      Tab(3).Control(3)=   "Frame12"
      Tab(3).Control(4)=   "Frame11"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Atendimento"
      TabPicture(4)   =   "frmSACNovo.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(1)=   "fraDados"
      Tab(4).Control(2)=   "fraHistorico"
      Tab(4).Control(3)=   "fraComandos"
      Tab(4).Control(4)=   "fraSolicitacao"
      Tab(4).ControlCount=   5
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4335
         Left            =   -74880
         TabIndex        =   197
         Top             =   360
         Width           =   11415
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "EM DESENVOLVIMENTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3000
            TabIndex        =   198
            Top             =   1200
            Width           =   5655
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Últimas Consultas à este CTC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -66240
         TabIndex        =   192
         Top             =   480
         Width           =   2895
         Begin MSDataGridLib.DataGrid gridUltCons 
            Bindings        =   "frmSACNovo.frx":008C
            Height          =   2295
            Left            =   120
            TabIndex        =   193
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   4048
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
            DataMember      =   "Sel_Ultconssac"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "data"
               Caption         =   "Data"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/mm/yyyy HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
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
            SplitCount      =   1
            BeginProperty Split0 
               AllowRowSizing  =   0   'False
               Locked          =   -1  'True
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  ColumnAllowSizing=   -1  'True
                  Locked          =   -1  'True
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   -1  'True
                  Locked          =   -1  'True
                  ColumnWidth     =   1035,213
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Notas Fiscais do CTC"
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
         Left            =   -74880
         TabIndex        =   188
         Top             =   1560
         Width           =   5775
         Begin MSDataGridLib.DataGrid gridNfsdoCTC 
            Bindings        =   "frmSACNovo.frx":00A5
            Height          =   2895
            Left            =   120
            TabIndex        =   189
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   5106
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
            DataMember      =   "Sel_NFsdoCTC"
            ColumnCount     =   19
            BeginProperty Column00 
               DataField       =   "idcodigo"
               Caption         =   "idcodigo"
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
            BeginProperty Column02 
               DataField       =   "numnfnum"
               Caption         =   "numnfnum"
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
            BeginProperty Column04 
               DataField       =   "cliente_cgc"
               Caption         =   "cliente_cgc"
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
               DataField       =   "cliente_nome"
               Caption         =   "cliente_nome"
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
               DataField       =   "emissao_nf"
               Caption         =   "emissao_nf"
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
               DataField       =   "numpedido"
               Caption         =   "numpedido"
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
               DataField       =   "dtpedido"
               Caption         =   "dtpedido"
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
               DataField       =   "valornf"
               Caption         =   "valornf"
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
               DataField       =   "pesonf"
               Caption         =   "pesonf"
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
               DataField       =   "volumesnf"
               Caption         =   "volumesnf"
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
               DataField       =   "data_interface"
               Caption         =   "data_interface"
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
               DataField       =   "hora_interface"
               Caption         =   "hora_interface"
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
               DataField       =   "at_cliente"
               Caption         =   "at_cliente"
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
               DataField       =   "canhotonf"
               Caption         =   "Canhoto"
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
               DataField       =   "canhotonfprot"
               Caption         =   "Prot. Canhoto"
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
               DataField       =   "canhotonfdata"
               Caption         =   "Prot. Data"
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
               DataField       =   "tem_ocorr"
               Caption         =   "Status"
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
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   854,929
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1440
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
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column14 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   840,189
               EndProperty
               BeginProperty Column16 
                  ColumnWidth     =   1275,024
               EndProperty
               BeginProperty Column17 
                  ColumnWidth     =   1349,858
               EndProperty
               BeginProperty Column18 
                  ColumnWidth     =   645,165
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Calendário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74760
         TabIndex        =   184
         Top             =   480
         Width           =   7095
         Begin MSComCtl2.MonthView calend1 
            Height          =   2370
            Left            =   1080
            TabIndex        =   185
            Top             =   360
            Visible         =   0   'False
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            MaxSelCount     =   64
            MonthColumns    =   2
            MultiSelect     =   -1  'True
            StartOfWeek     =   24576001
            CurrentDate     =   37678
         End
      End
      Begin VB.Frame fraDados 
         Caption         =   "Dados do Atendimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -73320
         TabIndex        =   88
         Top             =   2280
         Width           =   8535
         Begin VB.TextBox txtEmail 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5280
            TabIndex        =   95
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox txtFone 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   94
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtObservacao 
            Height          =   285
            Left            =   1680
            TabIndex        =   93
            Top             =   1560
            Width           =   6735
         End
         Begin VB.CommandButton cmdGravaCancela 
            Caption         =   "Cancela"
            Height          =   375
            Left            =   6720
            TabIndex        =   92
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton cmdGrava 
            Caption         =   "Grava"
            Height          =   375
            Left            =   4920
            TabIndex        =   91
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtDescricao 
            Height          =   735
            Left            =   960
            MultiLine       =   -1  'True
            TabIndex        =   90
            Top             =   720
            Width           =   7455
         End
         Begin VB.TextBox txtSolicitante 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   89
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Left            =   4800
            TabIndex        =   100
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Fone:"
            Height          =   195
            Left            =   2640
            TabIndex        =   99
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Observação Interna:"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante:"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Características da Carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -69000
         TabIndex        =   169
         Top             =   2400
         Width           =   2655
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Volumes:"
            Height          =   195
            Left            =   120
            TabIndex        =   183
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "CIF/FOB:"
            Height          =   195
            Left            =   120
            TabIndex        =   178
            Top             =   2040
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
            Height          =   195
            Left            =   120
            TabIndex        =   177
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
            Height          =   195
            Left            =   120
            TabIndex        =   176
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Natureza:"
            Height          =   195
            Left            =   120
            TabIndex        =   175
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Dimensões:"
            Height          =   195
            Left            =   120
            TabIndex        =   174
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor NFs:"
            Height          =   195
            Left            =   120
            TabIndex        =   173
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblfpag 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   170
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblEspecie 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   180
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblNatureza 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   181
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblDimensoes 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   179
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblVolumes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   171
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblPeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   172
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblValmerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   182
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Composição do Frete / Fatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -66240
         TabIndex        =   154
         Top             =   2400
         Width           =   2895
         Begin VB.Label lblTxOutros 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   191
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label lblTxColeta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   190
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Frete Calculado Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   163
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Outras Taxas:"
            Height          =   195
            Left            =   120
            TabIndex        =   161
            Top             =   2040
            Width           =   990
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Frete Nacional:"
            Height          =   195
            Left            =   120
            TabIndex        =   160
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "AdValor:"
            Height          =   195
            Left            =   120
            TabIndex        =   159
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Redesp.:"
            Height          =   195
            Left            =   120
            TabIndex        =   158
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Destino:"
            Height          =   195
            Left            =   120
            TabIndex        =   157
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Coleta:"
            Height          =   195
            Left            =   120
            TabIndex        =   156
            Top             =   1800
            Width           =   1125
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Origem:"
            Height          =   195
            Left            =   120
            TabIndex        =   155
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label lblRedespacho 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   168
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblTxDestino 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   167
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblTxOrigem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   166
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblAdValorem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   165
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblFreteNacional 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   164
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblFreteTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   162
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Consignatário"
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
         Left            =   -69000
         TabIndex        =   151
         Top             =   1560
         Width           =   5655
         Begin VB.Label lblRespons_CGC 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   153
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label lblRespons_Nome 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   152
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dados do Destinatário"
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
         Left            =   -69000
         TabIndex        =   146
         Top             =   480
         Width           =   5655
         Begin VB.Label lblEndDest 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   147
            Top             =   600
            Width           =   4515
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "CGC Dest.:"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   149
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblDest_CGC 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   148
            Top             =   360
            Width           =   2325
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados do Remetente"
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
         Left            =   -74880
         TabIndex        =   141
         Top             =   480
         Width           =   5775
         Begin VB.Label lblEndRem 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   142
            Top             =   600
            Width           =   4605
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "CGC Cliente:"
            Height          =   195
            Left            =   120
            TabIndex        =   145
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   144
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblRemet_CGC 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   143
            Top             =   360
            Width           =   2220
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Estatística do Manifesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -69120
         TabIndex        =   122
         Top             =   3240
         Width           =   5775
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "NFs/CTCs:"
            Height          =   195
            Left            =   120
            TabIndex        =   140
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Frete/Peso:"
            Height          =   195
            Left            =   1800
            TabIndex        =   139
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Frete/Val.NF:"
            Height          =   195
            Left            =   3720
            TabIndex        =   138
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label lblManifEstatIndFreteNf 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4800
            TabIndex        =   137
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblManifEstatIndFretePeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2760
            TabIndex        =   136
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblManifEstatIndNfCtc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   135
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblManifEstatFRETE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4200
            TabIndex        =   134
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblManifEstatNF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   133
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblManifEstatPESO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   132
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblManifEstatVOL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   131
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblManifEstatVALMERC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4200
            TabIndex        =   130
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblManifEstatCTC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   129
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Valor Frete:"
            Height          =   195
            Left            =   3240
            TabIndex        =   128
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
            Height          =   195
            Left            =   1440
            TabIndex        =   127
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Valor NFs:"
            Height          =   195
            Left            =   3240
            TabIndex        =   126
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "NFs:"
            Height          =   195
            Left            =   120
            TabIndex        =   125
            Top             =   720
            Width           =   330
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Volumes:"
            Height          =   195
            Left            =   1440
            TabIndex        =   124
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CTCs:"
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "CTCs do Manifesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74880
         TabIndex        =   120
         Top             =   1680
         Width           =   5655
         Begin MSDataGridLib.DataGrid gridCTCManifesto 
            Bindings        =   "frmSACNovo.frx":00BE
            Height          =   2775
            Left            =   120
            TabIndex        =   121
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4895
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
            DataMember      =   "Sel_CTCsdoManifesto"
            ColumnCount     =   10
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
               DataField       =   "peso"
               Caption         =   "Peso"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#.##0,0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "volumes"
               Caption         =   "Vols"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#.##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "remet_nome"
               Caption         =   "Cliente Remetente"
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
               DataField       =   "dest_nome"
               Caption         =   "Cliente Destino"
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
            BeginProperty Column06 
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
            BeginProperty Column07 
               DataField       =   "valmerc"
               Caption         =   "Val. NF"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "fretetotal"
               Caption         =   "Frete"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "nfs"
               Caption         =   "NFs deste CTC"
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
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   629,858
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   434,835
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2670,236
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2355,024
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   2129,953
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   360
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   5715,213
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Manifesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   117
         Top             =   480
         Width           =   5655
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            Height          =   2175
            Left            =   3840
            TabIndex        =   118
            Top             =   2400
            Width           =   5895
         End
         Begin MSDataGridLib.DataGrid gridManifesto 
            Bindings        =   "frmSACNovo.frx":00D7
            Height          =   735
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1296
            _Version        =   393216
            BackColor       =   8388608
            Enabled         =   0   'False
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
            DataMember      =   "Sel_ManifestoPorCTC"
            ColumnCount     =   15
            BeginProperty Column00 
               DataField       =   "idcodigo"
               Caption         =   "idcodigo"
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
            BeginProperty Column02 
               DataField       =   "filialmanifesto"
               Caption         =   "filialmanifesto"
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
               DataField       =   "filial"
               Caption         =   "Filial"
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
               DataField       =   "manifesto"
               Caption         =   "Manif."
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
               DataField       =   "embarcador"
               Caption         =   "embarcador"
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
               DataField       =   "dtemissao"
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
            BeginProperty Column07 
               DataField       =   "hsemissao"
               Caption         =   "Hs."
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
               DataField       =   "dtsaida"
               Caption         =   "dtsaida"
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
               DataField       =   "hssaida"
               Caption         =   "hssaida"
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
               DataField       =   "placaveic"
               Caption         =   "Placa"
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
               DataField       =   "motorista"
               Caption         =   "Motorista"
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
            BeginProperty Column13 
               DataField       =   "at_manif_cif"
               Caption         =   "at_manif_cif"
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
               DataField       =   "at_manif_cif_data"
               Caption         =   "at_manif_cif_data"
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
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   390,047
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   675,213
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   464,882
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   929,764
               EndProperty
               BeginProperty Column14 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Observação do SITLA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -69120
         TabIndex        =   116
         Top             =   480
         Width           =   2775
         Begin TabDlg.SSTab SSTab2 
            Height          =   2175
            Left            =   120
            TabIndex        =   194
            Top             =   360
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   3836
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "De Emissão..."
            TabPicture(0)   =   "frmSACNovo.frx":00F0
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblObs_Emissao"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "De Ocorr..."
            TabPicture(1)   =   "frmSACNovo.frx":010C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblObs_OcorrSitla"
            Tab(1).ControlCount=   1
            Begin VB.Label lblObs_Emissao 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   120
               TabIndex        =   196
               Top             =   480
               Width           =   2295
            End
            Begin VB.Label lblObs_OcorrSitla 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   1575
               Left            =   -74880
               TabIndex        =   195
               Top             =   480
               Width           =   2295
            End
         End
      End
      Begin VB.Frame fraHistorico 
         Caption         =   "Histórico de Atendimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -68400
         TabIndex        =   114
         Top             =   1440
         Width           =   5055
         Begin MSDataGridLib.DataGrid gridHistorico 
            Height          =   3015
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   5318
            _Version        =   393216
            BackColor       =   8388608
            ForeColor       =   8454143
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
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
         Left            =   -74880
         TabIndex        =   103
         Top             =   480
         Width           =   11535
         Begin VB.CommandButton cmdPendentes 
            Caption         =   "Solicitações Pendentes"
            Height          =   495
            Left            =   7080
            TabIndex        =   113
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdX 
            Caption         =   "X"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11040
            TabIndex        =   112
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdf5 
            Caption         =   "F5"
            Enabled         =   0   'False
            Height          =   495
            Left            =   10560
            TabIndex        =   111
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdf4 
            Caption         =   "F4"
            Enabled         =   0   'False
            Height          =   495
            Left            =   10080
            TabIndex        =   110
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdf3 
            Caption         =   "F3"
            Enabled         =   0   'False
            Height          =   495
            Left            =   9600
            TabIndex        =   109
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdf2 
            Caption         =   "F2"
            Enabled         =   0   'False
            Height          =   495
            Left            =   9120
            TabIndex        =   108
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdf1 
            Caption         =   "F1"
            Enabled         =   0   'False
            Height          =   495
            Left            =   8640
            TabIndex        =   107
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdFinalizaSolic 
            Caption         =   "Finalizar Solicitação de Atendimento"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5280
            TabIndex        =   106
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdNovaOcorrInf 
            Caption         =   "Incluir Nova Ocorrência/Informação de Atendimento"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2400
            TabIndex        =   105
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton cmdNovasolic 
            Caption         =   "Abrir uma Nova Solicitação de Atendimento"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   104
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame fraSolicitacao 
         Caption         =   "Solicitação de Atendimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         TabIndex        =   101
         Top             =   1440
         Width           =   6375
         Begin MSDataGridLib.DataGrid GridSolicitacao 
            Height          =   3015
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   5318
            _Version        =   393216
            BackColor       =   8388608
            ForeColor       =   8454143
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Pré - Baixa (emails, relatórios, etc)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -67440
         TabIndex        =   77
         Top             =   480
         Width           =   3915
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Em: (data/hs)"
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label lblUsu_bxpre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   86
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Quem Baixou:"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   2400
            TabIndex        =   84
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   240
            TabIndex        =   83
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   240
            TabIndex        =   82
            Top             =   840
            Width           =   525
         End
         Begin VB.Label lblHsBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2880
            TabIndex        =   81
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblDtBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   80
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label lblRecebPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   79
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label lblDtUsuPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   78
            Top             =   1560
            Width           =   2295
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Baixa Física (pelo CTC físico)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -67440
         TabIndex        =   66
         Top             =   2640
         Width           =   3915
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Data Baixa:"
            Height          =   195
            Left            =   240
            TabIndex        =   76
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label lblUsu_bx 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   75
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Quem Baixou:"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   2400
            TabIndex        =   73
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   240
            TabIndex        =   72
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   240
            TabIndex        =   71
            Top             =   840
            Width           =   525
         End
         Begin VB.Label lblDtBaixa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   70
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label lblHsBaixa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2880
            TabIndex        =   69
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblReceb 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   68
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label lblDtUsu 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   67
            Top             =   1560
            Width           =   2295
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Observação de Entrega"
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
         Left            =   -74760
         TabIndex        =   64
         Top             =   3480
         Width           =   7095
         Begin VB.Label lblObsEntr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   765
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   6840
         End
      End
      Begin VB.Frame FraNfsOcorr 
         Caption         =   "Total de Notas Fiscais do CTC: 0 NF(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   6600
         TabIndex        =   58
         Top             =   1620
         Width           =   5055
         Begin VB.ListBox lstNfOcorrNAO 
            Height          =   1815
            ItemData        =   "frmSACNovo.frx":0128
            Left            =   3960
            List            =   "frmSACNovo.frx":012A
            TabIndex        =   63
            Top             =   360
            Width           =   975
         End
         Begin VB.ListBox lstNfOcorr 
            Height          =   1815
            ItemData        =   "frmSACNovo.frx":012C
            Left            =   120
            List            =   "frmSACNovo.frx":012E
            TabIndex        =   62
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblOcorrPorNF 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "0 Nfs tiveram a Ocorr. Abaixo"
            Height          =   195
            Left            =   1440
            TabIndex        =   61
            Top             =   480
            Width           =   2085
         End
         Begin VB.Label lblOcorrSelec 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   765
            Left            =   1200
            TabIndex        =   60
            Top             =   1440
            Width           =   2685
         End
         Begin VB.Label lblOcorrPorNFNao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "0 Nfs NÃO tiveram a Ocorr. Abaixo"
            Height          =   195
            Left            =   1320
            TabIndex        =   59
            Top             =   960
            Width           =   2475
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   3600
            X2              =   1080
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   3960
            X2              =   1200
            Y1              =   1200
            Y2              =   1200
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Protocolo POD: Núm / Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         TabIndex        =   21
         Top             =   900
         Width           =   2655
         Begin VB.Label lblArquivo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblNumProt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Observação de Ocorrência"
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
         TabIndex        =   19
         Top             =   3960
         Width           =   11535
         Begin VB.Label lblObs_Ocorr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   480
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   11295
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Tab. Prazos / Meta / Prev. de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   900
         Width           =   3615
         Begin VB.CommandButton cmdRastPrevEntr 
            Caption         =   "?"
            Height          =   255
            Left            =   3000
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblMetaPrazo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   187
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblTabPrazo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblPrevEntr 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Dias"
            Height          =   195
            Left            =   1440
            TabIndex        =   17
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Data de Entrega / Prazo / Recebedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   12
         Top             =   900
         Width           =   5055
         Begin VB.CommandButton cmdRastrPrazo 
            Caption         =   "?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4440
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblPrazoDias 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   186
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblRecebedor 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   57
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblEntrega 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Dias"
            Height          =   195
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame fraOcorrencias 
         Caption         =   "Ocorrências da NF/CTC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   1620
         Width           =   6375
         Begin MSDataGridLib.DataGrid GridConsOcorr 
            Bindings        =   "frmSACNovo.frx":0130
            Height          =   1935
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            ColumnCount     =   7
            BeginProperty Column00 
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
            BeginProperty Column01 
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
            BeginProperty Column02 
               DataField       =   "cod_ocorr"
               Caption         =   "cd"
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
               DataField       =   "descr_ocorr"
               Caption         =   "ocorrência / descrição"
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
               DataField       =   "usu_ocorr"
               Caption         =   "usuário"
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
               DataField       =   "usu_dataocorr"
               Caption         =   "data inclusão"
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
               DataField       =   "obs_ocorr"
               Caption         =   "obs_ocorr"
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
               Locked          =   -1  'True
               BeginProperty Column00 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   480,189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   269,858
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   3509,858
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1560,189
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   14,74
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblMensagem 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Mensagem: "
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
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Top             =   540
         Width           =   11535
      End
   End
   Begin VB.Frame Frame23 
      Caption         =   "CTC - Conhecimento de Transporte de Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   41
      Top             =   960
      Width           =   11775
      Begin VB.Label lblTranspSub 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10440
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Contratado:"
         Height          =   195
         Left            =   9240
         TabIndex        =   52
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Modal/Via:"
         Height          =   195
         Left            =   6600
         TabIndex        =   51
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblVia 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8640
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblModal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7440
         TabIndex        =   49
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblEmissor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5280
         TabIndex        =   48
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblHora 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   47
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblData 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         TabIndex        =   46
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "Emissor:"
         Height          =   195
         Left            =   4680
         TabIndex        =   45
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   2280
         TabIndex        =   44
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblFilialctc 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Filial CTC:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdPesq 
      Caption         =   "Pesquisa"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame xt 
      Caption         =   "Procura Por ..."
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
      TabIndex        =   40
      Top             =   120
      Width           =   3360
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtCtc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtNumNf 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.OptionButton optCTC 
         Caption         =   "Por Núm. de CTC"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optNf 
         Caption         =   "Por Núm de NF"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   " S T A T U S "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7320
      TabIndex        =   38
      Top             =   120
      Width           =   4575
      Begin VB.Label lblEntregueSN 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   4305
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Origem / Remetente"
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
      TabIndex        =   31
      Top             =   1680
      Width           =   5895
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   4800
         TabIndex        =   35
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblRemet_Nome 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   34
         Top             =   240
         Width           =   4740
      End
      Begin VB.Label lblCidade_orig 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   33
         Top             =   600
         Width           =   3690
      End
      Begin VB.Label lblUF_Orig 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5160
         TabIndex        =   32
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Destino / Destinatário"
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
      Left            =   6000
      TabIndex        =   24
      Top             =   1680
      Width           =   5895
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Destinatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Cidade: "
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   4800
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblDest_Nome 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   4635
      End
      Begin VB.Label lblCidade_Dest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   600
         Width           =   3570
      End
      Begin VB.Label lblUf_Dest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5160
         TabIndex        =   25
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.CommandButton cmbProcurar 
      Caption         =   "Procurar"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   495
      Left            =   6600
      Picture         =   "frmSACNovo.frx":0149
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmbSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmSACNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public xtemocorr As String
Private Sub cmbProcurar_Click()
Dim xentrega As String, xhora_ent As String, xrecebedor As String, xobs_ocorr As String, xdata As String
Dim xagora As Date, xtemocorrCTC As String
    
'BUSCA DATA DO SERVIDOR
If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
de_informa.Sel_DataServidor
xagora = de_informa.rsSel_DataServidor.Fields("agora")

'grava nas variaveis globais os ctcs / nfs consultados
    
xultimofilial = txtFilial
xultimoctc = txtCtc
xultimonf = txtNumNf

'Limpa os grids do form
GridConsOcorr.DataMember = ""
GridConsOcorr.Refresh
gridHistorico.DataMember = ""
gridHistorico.Refresh
gridCTCManifesto.DataMember = ""
gridCTCManifesto.Refresh
gridManifesto.DataMember = ""
gridManifesto.Refresh
gridNfsdoCTC.DataMember = ""
gridNfsdoCTC.Refresh
GridSolicitacao.DataMember = ""
GridSolicitacao.Refresh
gridUltCons.DataMember = ""
gridUltCons.Refresh

'limpa os Label/Outros Objetos do Form

Call limpatela(Me)

lstNfOcorr.Clear
lstNfOcorrNAO.Clear
lblEntregueSN = ""
lblOcorrPorNF = "0 Nfs tiveram a Ocorr. Abaixo"
lblOcorrPorNFNao = "0 Nfs NÃO tiveram a Ocorr. Abaixo"
FraNfsOcorr = "Total de Notas Fiscais do CTC: 0 NF(s)"
lblMensagem = " Mensagem:"
calend1.Visible = False

'volta os text filial/ctc/nf para os valores globais

txtFilial = xultimofilial
txtCtc = xultimoctc
txtNumNf = xultimonf

If optCTC.Value = True Then   'Se for procura por Filial / CTC
        
    lblFilialctc = transctc(txtFilial, txtCtc)
    
    'verifica consistência dos dados
        
    If txtFilial.Text = "" Or txtCtc.Text = "" Then
        MsgBox "Filial / CTC Inválidos !", vbCritical, "Erro"
        txtFilial.SetFocus
        Exit Sub
    End If
    If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
    de_informa.Sel_Ctc_SAC lblFilialctc 'Procura na Tabela a Filial/CTC
    If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then
        MsgBox "Número de Filial/CTC Não Encontrados !", vbCritical + vbOKOnly, "Erro"
        txtFilial.SetFocus
        Exit Sub
    End If

    'BUSCA AS NOTAS FISCAIS DESTE CTC
    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
    de_informa.Sel_NFsdoCTC lblFilialctc
    
ElseIf optNf.Value = True Then   'Se for procura por Nota Fiscal

    'verifica consistência dos dados
        
    If txtNumNf.Text = "" Then
        MsgBox "Número de Nota Fiscal Inválida !", vbCritical, "Erro"
        txtNumNf.SetFocus
        Exit Sub
    End If
    If de_informa.rsSel_NF_SAC.State = 1 Then de_informa.rsSel_NF_SAC.Close
    de_informa.Sel_NF_SAC Val(txtNumNf)   'Procura a NF na Tabela
    If de_informa.rsSel_NF_SAC.RecordCount = 0 Then
        MsgBox "Número de NF Não Encontrado !", vbCritical + vbOKOnly, "Erro"
        txtNumNf.SetFocus
        Exit Sub
    ElseIf de_informa.rsSel_NF_SAC.RecordCount > 1 Then  'achou mais de uma NF com o mesmo número
        frmDuplNF.Caption = "SAC - Número de NFs Duplicadas"
        frmDuplNF.Show 1  'direciona para o form que trata casos de NF duplicadas
    Else  'Caso seja encontrada somente uma NF
        lblFilialctc = de_informa.rsSel_NF_SAC.Fields("filialctc")
    End If
    
    'PROCURA O CTC NA TABELA
    If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
    de_informa.Sel_Ctc_SAC lblFilialctc  'Procura na Tabela a Filial/CTC

    If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then
        MsgBox "Erro de Consistência. Chame Suporte Técnico ! ", vbCritical + vbOKOnly, "Erro" 'erro de consistência
        txtNumNf.SetFocus
        Exit Sub
    End If
    
    'posiciona o registro na NF que está sendo consultada
    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
    de_informa.Sel_NFsdoCTC lblFilialctc
    Do Until de_informa.rsSel_NFsdoCTC.Fields("numnfnum") = Val(txtNumNf)
        de_informa.rsSel_NFsdoCTC.MoveNext
    Loop
    
End If

'DADOS SUPERIORES GERAIS *****************************************************************
        
'CTC - Conhecimento de Transporte de Carga
        
lblData = de_informa.rsSel_Ctc_SAC.Fields("data")
lblHora = de_informa.rsSel_Ctc_SAC.Fields("hora")
lblEmissor = de_informa.rsSel_Ctc_SAC.Fields("emissor")
lblModal = de_informa.rsSel_Ctc_SAC.Fields("modal")
lblVia = de_informa.rsSel_Ctc_SAC.Fields("via")
lblTranspSub = de_informa.rsSel_Ctc_SAC.Fields("transp_sub")
    
If optCTC = True Then
    xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
ElseIf optNf = True Then
    xtemocorr = de_informa.rsSel_NFsdoCTC.Fields("tem_ocorr")
    xtemocorrCTC = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
    If xtemocorr <> xtemocorrCTC Then
        lblMensagem = " Mensagem: O CTC desta NF Está Com Status Diferente. Para Maiores Informações Consulte o CTC."
    End If
End If
        
'Origem / Remetente
        
If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
de_informa.Sel_ConsCadCli de_informa.rsSel_Ctc_SAC.Fields("remet_cgc")
        
lblRemet_Nome = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
lblCidade_orig = de_informa.rsSel_Ctc_SAC.Fields("cidade_orig")
If Not IsNull(de_informa.rsSel_ConsCadCli.Fields("uf")) Then
    lblUF_Orig = de_informa.rsSel_ConsCadCli.Fields("uf")
End If
        
'Destino / Destinatário
        
If de_informa.rsSel_ConsCadCliDest.State = 1 Then de_informa.rsSel_ConsCadCliDest.Close
de_informa.Sel_ConsCadCliDest de_informa.rsSel_Ctc_SAC.Fields("dest_cgc")
        
lblDest_Nome = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
lblCidade_Dest = de_informa.rsSel_Ctc_SAC.Fields("cidade_dest")
lblUf_Dest = de_informa.rsSel_Ctc_SAC.Fields("uf_dest")
        
'atualiza a tela
DoEvents
    
'DADOS DAS TAB (ssTab) ******************************************************************

'Tab - Ocorrências e Entrega
'Tab. Prazos / Meta / Prev. de Entrega
        
lblTabPrazo = de_informa.rsSel_ConsCadCli.Fields("prazo")
lblMetaPrazo = buscaprazo(lblUF_Orig, lblCidade_orig, lblTabPrazo, Mid$(lblModal, 1, 1))
lblPrevEntr = de_informa.rsSel_Ctc_SAC.Fields("prev_entrega")
        
'Data de Entrega / Prazo / Recebedor  e  Protocolo POD: Núm / Data
        
If xtemocorr = "1" Then  'se já estiver entregue (ctc ou nf)
    If optCTC = True Then  'se for por CTC busca pelo número do CTC
        If de_informa.rsSel_ConsEntregaCTC.State = 1 Then de_informa.rsSel_ConsEntregaCTC.Close
        de_informa.Sel_ConsEntregaCTC lblFilialctc
            
        lblEntrega = de_informa.rsSel_ConsEntregaCTC.Fields("data")
        If Not IsNull(de_informa.rsSel_ConsEntregaCTC.Fields("diasuteis")) Then
            lblPrazoDias = de_informa.rsSel_ConsEntregaCTC.Fields("diasuteis")
        End If
        If de_informa.rsSel_ConsEntregaCTC.Fields("baixadofinal") = "S" Then
            lblRecebedor = de_informa.rsSel_ConsEntregaCTC.Fields("receb")
        Else
            lblRecebedor = de_informa.rsSel_ConsEntregaCTC.Fields("recebpre")
        End If
        If Not IsNull(de_informa.rsSel_ConsEntregaCTC.Fields("rel_arq_num")) Then
            lblNumProt = de_informa.rsSel_ConsEntregaCTC.Fields("rel_arq_num")
            lblArquivo = de_informa.rsSel_ConsEntregaCTC.Fields("rel_arq_data")
        End If
    ElseIf optNf = True Then  'se for por NF, busca a entrega da NF
        If de_informa.rsSel_ConsEntregaNF.State = 1 Then de_informa.rsSel_ConsEntregaNF.Close
        de_informa.Sel_ConsEntregaNF lblFilialctc, txtNumNf
            
        lblEntrega = de_informa.rsSel_ConsEntregaNF.Fields("data")
        If Not IsNull(de_informa.rsSel_ConsEntregaNF.Fields("diasuteis")) Then
            lblPrazoDias = de_informa.rsSel_ConsEntregaNF.Fields("diasuteis")
        End If
        If de_informa.rsSel_ConsEntregaNF.Fields("baixadofinal") = "S" Then
            lblRecebedor = de_informa.rsSel_ConsEntregaNF.Fields("receb")
        Else
            lblRecebedor = de_informa.rsSel_ConsEntregaNF.Fields("recebpre")
        End If
        If Not IsNull(de_informa.rsSel_ConsEntregaNF.Fields("rel_arq_num")) Then
            lblNumProt = de_informa.rsSel_ConsEntregaNF.Fields("rel_arq_num")
            lblArquivo = de_informa.rsSel_ConsEntregaNF.Fields("rel_arq_data")
        End If
    End If
End If

'se for ocorrência de baixa sem entrega o POD/CTC pode ter sido enviado para o Arquivo

If xtemocorr = "0" Then 'se for CTC baixado busca se foi para o Arquivo
    If optCTC = True Then
        If de_informa.rsSel_ConsOcorrCTC.State = 1 Then de_informa.rsSel_ConsOcorrCTC.Close
        de_informa.Sel_ConsOcorrCTC lblFilialctc, "00"
        If Not IsNull(de_informa.rsSel_ConsOcorrCTC.Fields("rel_arq_num")) Then
            lblNumProt = de_informa.rsSel_ConsOcorrCTC.Fields("rel_arq_num")
            lblArquivo = de_informa.rsSel_ConsOcorrCTC.Fields("rel_arq_data")
        End If
    ElseIf optNf = True Then
        If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
        de_informa.Sel_ConsOcorrNF lblFilialctc, txtNumNf, "00"
        If Not IsNull(de_informa.rsSel_ConsOcorrNF.Fields("rel_arq_num")) Then
            lblNumProt = de_informa.rsSel_ConsOcorrNF.Fields("rel_arq_num")
            lblArquivo = de_informa.rsSel_ConsOcorrNF.Fields("rel_arq_data")
        End If
    End If
End If
        
'Ocorrências (GRID)
        
If optCTC = True Then
    If de_informa.rsSel_ConsOcorrCTC.State = 1 Then de_informa.rsSel_ConsOcorrCTC.Close
    de_informa.Sel_ConsOcorrCTC lblFilialctc, "%"
    
    fraOcorrencias = "Ocorrências deste CTC"
    GridConsOcorr.DataMember = "sel_consocorrctc"
    GridConsOcorr.Refresh
    
    If de_informa.rsSel_ConsOcorrCTC.RecordCount > 0 Then
        'observação de ocorrência
        de_informa.rsSel_ConsOcorrCTC.MoveFirst
        If Not IsNull(de_informa.rsSel_ConsOcorrCTC.Fields("obs_ocorr")) Then
            lblObs_Ocorr = de_informa.rsSel_ConsOcorrCTC.Fields("obs_ocorr")
        End If
    End If
ElseIf optNf = True Then
    If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
    de_informa.Sel_ConsOcorrNF lblFilialctc, txtNumNf, "%"
    
    fraOcorrencias = "Ocorrências desta Nota Fiscal"
    GridConsOcorr.DataMember = "sel_consocorrnf"
    GridConsOcorr.Refresh
    
    If de_informa.rsSel_ConsOcorrNF.RecordCount > 0 Then
        'observação de ocorrência
        de_informa.rsSel_ConsOcorrNF.MoveFirst
        If Not IsNull(de_informa.rsSel_ConsOcorrNF.Fields("obs_ocorr")) Then
            lblObs_Ocorr = de_informa.rsSel_ConsOcorrNF.Fields("obs_ocorr")
        End If
    End If
End If
        
FraNfsOcorr = "Total de Notas Fiscais do CTC: " & Trim$(Str(de_informa.rsSel_NFsdoCTC.RecordCount)) & " NF(s)"

If xtemocorr <> "N" And xtemocorr <> "C" Then
    'Total de Notas Fiscais deste CTC:
    If de_informa.rsSel_NfsComOcorr.State = 1 Then de_informa.rsSel_NfsComOcorr.Close
    de_informa.Sel_NfsComOcorr lblFilialctc, GridConsOcorr.Columns(2)
            
    If de_informa.rsSel_NfsNaoOcorr.State = 1 Then de_informa.rsSel_NfsNaoOcorr.Close
    de_informa.Sel_NfsNaoOcorr lblFilialctc, GridConsOcorr.Columns(2)
            
    Do Until de_informa.rsSel_NfsComOcorr.EOF  'preenche a LST Nf com Ocorr
        lstNfOcorr.AddItem de_informa.rsSel_NfsComOcorr.Fields("numnf")
        de_informa.rsSel_NfsComOcorr.MoveNext
    Loop
            
    lblOcorrPorNF = Trim$(Str(de_informa.rsSel_NfsComOcorr.RecordCount)) & _
                    " Nfs tiveram a Ocorr. Abaixo"
            
    Do Until de_informa.rsSel_NfsNaoOcorr.EOF  'preenche a LST Nf com Ocorr
        lstNfOcorrNAO.AddItem de_informa.rsSel_NfsNaoOcorr.Fields("numnf")
        de_informa.rsSel_NfsNaoOcorr.MoveNext
    Loop
            
    lblOcorrPorNFNao = Trim$(Str(de_informa.rsSel_NfsNaoOcorr.RecordCount)) & _
                    " Nfs NÃO tiveram a Ocorr. Abaixo"
            
    lblOcorrSelec = GridConsOcorr.Columns(3)
            
End If
        
'S T A T U S   D O   C T C / N F
        
If optCTC = True Then
    fraStatus = "S T A T U S   D O   C T C"
ElseIf optNf = True Then
    fraStatus = "S T A T U S   D A   N F"
End If

lblEntregueSN.ToolTipText = ""
If xtemocorr = "0" Then
    lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
    lblEntregueSN.Caption = "OCORR/Baixado"
ElseIf xtemocorr = "1" Then
    lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
    If IsNumeric(lblPrazoDias) Then
        If Val(lblPrazoDias) = Val(lblMetaPrazo) Then
            lblEntregueSN.Caption = "OK. ENTREGUE - No Prazo"
        ElseIf Val(lblPrazoDias) > Val(lblMetaPrazo) Then
            lblEntregueSN.Caption = "OK. ENTREGUE - Atraso: " & Trim$(Str(Val(lblPrazoDias) - Val(lblMetaPrazo))) & " dia(s)"
        ElseIf Val(lblPrazoDias) < Val(lblMetaPrazo) Then
            lblEntregueSN.Caption = "OK. ENTREGUE - Antecipado: " & Trim$(Str(Val(lblMetaPrazo) - Val(lblPrazoDias))) & " dia(s)"
        End If
    Else
        lblEntregueSN.Caption = "OK. ENTREGUE"
    End If
ElseIf xtemocorr = "2" Then
    lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
    lblEntregueSN.Caption = "OCORR/Pendente"
ElseIf xtemocorr = "N" Then
    If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= xagora Then
        lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
        lblEntregueSN.Caption = "EM TRÂNSITO"
        lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
    Else
        lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
        lblEntregueSN.Caption = "SEM POSIÇÃO há " & Trim$(Str(Val(xagora - de_informa.rsSel_Ctc_SAC.Fields("prev_entrega")))) & " dia(s)"
        lblEntregueSN.ToolTipText = "SEM POSIÇÃO = Após a Previsão de Entrega"
    End If
ElseIf xtemocorr = "C" Then
    cmdRastPrevEntr.Enabled = False
    lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
    lblEntregueSN.Caption = "CTC CANCELADO"
    lblEntregueSN.ToolTipText = "Cancelado em:" & de_informa.rsSel_Ctc_SAC.Fields("canc_data") & _
                                "  Usuário:" & de_informa.rsSel_Ctc_SAC.Fields("canc_usu") & _
                                "  Motivo:" & de_informa.rsSel_Ctc_SAC.Fields("canc_obs")
End If
        
'FINALIZA A PRIMEIRA TELA E A PRIMEIRA TAB e atualiza a tela
DoEvents
            
'Tab - Detalhe da Entrega
        
'calendário
calend1.SelStart = CDate(lblData)
calend1.SelEnd = CDate(lblData)
               
If xtemocorr = "1" Then
        
    'Calendário
    If Abs(Month(CDate(lblData)) - Month(CDate(lblEntrega))) > 1 Then
    Else
        calend1.Visible = True
        calend1.SelStart = CDate(lblData)
        calend1.SelEnd = CDate(lblEntrega)
    End If
            
    If optCTC = True Then
    
        'Pré Baixa (emails, relatórios, etc)
        lblDtBaixaPre = de_informa.rsSel_ConsEntregaCTC.Fields("dtbaixapre")
        lblHsBaixaPre = de_informa.rsSel_ConsEntregaCTC.Fields("hsbaixapre")
        lblRecebPre = de_informa.rsSel_ConsEntregaCTC.Fields("recebpre")
        lblUsu_bxpre = de_informa.rsSel_ConsEntregaCTC.Fields("usu_bxpre")
        lblDtUsuPre = de_informa.rsSel_ConsEntregaCTC.Fields("usu_datapre")
        
        If de_informa.rsSel_ConsEntregaCTC.Fields("baixadofinal") = "S" Then
                
            'Baixa Física (pelo CTC Físico)
            lblDtBaixa = de_informa.rsSel_ConsEntregaCTC.Fields("dtbaixa")
            lblHsBaixa = de_informa.rsSel_ConsEntregaCTC.Fields("hsbaixa")
            lblReceb = de_informa.rsSel_ConsEntregaCTC.Fields("receb")
            lblUsu_bx = de_informa.rsSel_ConsEntregaCTC.Fields("usu_bx")
            lblDtUsu = de_informa.rsSel_ConsEntregaCTC.Fields("usu_databx")
                        
        End If
        
        'Observação de Entrega
        lblObsEntr = de_informa.rsSel_ConsEntregaCTC.Fields("obs_ocorr")
        
    ElseIf optNf = True Then
    
        'Pré Baixa (emails, relatórios, etc)
        lblDtBaixaPre = de_informa.rsSel_ConsEntregaNF.Fields("dtbaixapre")
        lblHsBaixaPre = de_informa.rsSel_ConsEntregaNF.Fields("hsbaixapre")
        lblRecebPre = de_informa.rsSel_ConsEntregaNF.Fields("recebpre")
        lblUsu_bxpre = de_informa.rsSel_ConsEntregaNF.Fields("usu_bxpre")
        lblDtUsuPre = de_informa.rsSel_ConsEntregaNF.Fields("usu_datapre")
        
        If de_informa.rsSel_ConsEntregaNF.Fields("baixadofinal") = "S" Then
                
            'Baixa Física (pelo CTC Físico)
            lblDtBaixa = de_informa.rsSel_ConsEntregaNF.Fields("dtbaixa")
            lblHsBaixa = de_informa.rsSel_ConsEntregaNF.Fields("hsbaixa")
            lblReceb = de_informa.rsSel_ConsEntregaNF.Fields("receb")
            lblUsu_bx = de_informa.rsSel_ConsEntregaNF.Fields("usu_bx")
            lblDtUsu = de_informa.rsSel_ConsEntregaNF.Fields("usu_data")
                        
        End If
        
        'Observação de Entrega
        lblObsEntr = de_informa.rsSel_ConsEntregaNF.Fields("obs_ocorr")
    
    End If
End If
        
'Tab - Dados do CTC
            
'Dados do Remetente
lblRemet_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), "@@.@@@.@@@/@@@@-@@")
If Not IsNull(de_informa.rsSel_ConsCadCli.Fields("endereco")) Then
    lblEndRem = de_informa.rsSel_ConsCadCli.Fields("endereco")
End If
        
'Dados do Destinatário
lblDest_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("dest_cgc"), "@@.@@@.@@@/@@@@-@@")
If Not IsNull(de_informa.rsSel_ConsCadCliDest.Fields("endereco")) Then
    lblEndDest = de_informa.rsSel_ConsCadCliDest.Fields("endereco")
End If
        
'Consignatário
lblRespons_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("respons_cgc"), "@@.@@@.@@@/@@@@-@@")
If de_informa.rsSel_BuscaConsig.State = 1 Then de_informa.rsSel_BuscaConsig.Close
de_informa.Sel_BuscaConsig de_informa.rsSel_Ctc_SAC.Fields("respons_cgc")
If de_informa.rsSel_BuscaConsig.RecordCount > 0 Then
    lblRespons_Nome = de_informa.rsSel_BuscaConsig.Fields("nome")
End If
        
'Composição do Frete / Fatura
lblValmerc = Format(de_informa.rsSel_Ctc_SAC.Fields("valmerc"), "##,###,##0.00")
lblPeso = Format(de_informa.rsSel_Ctc_SAC.Fields("peso"), "##,##0.0")
lblVolumes = Format(de_informa.rsSel_Ctc_SAC.Fields("volumes"), "##,##0")
lblEspecie = de_informa.rsSel_Ctc_SAC.Fields("especie")
lblNatureza = de_informa.rsSel_Ctc_SAC.Fields("natureza")
lblDimensoes = de_informa.rsSel_Ctc_SAC.Fields("dimensoes")
lblfpag = de_informa.rsSel_Ctc_SAC.Fields("fpag")
        
'Características da Carga
lblFreteNacional = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("fretenacional")), "#,###,##0.00")
lblAdValorem = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("advalorem")), "#,###,##0.00")
lblTxOrigem = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txorigem")), "#,###,##0.00")
lblTxDestino = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txdestino")), "#,###,##0.00")
lblRedespacho = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txredespacho")), "#,###,##0.00")
lblTxColeta = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txcoleta")), "#,###,##0.00")
lblTxOutros = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txoutros")), "#,###,##0.00")
lblFreteTotal = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("fretetotal")), "#,###,##0.00")
        
'Notas Fiscais do CTC

gridNfsdoCTC.DataMember = "sel_nfsdoctc"
gridNfsdoCTC.Refresh
        
'Tab - OBS SITLA e Manifesto
        
'Manifesto
If de_informa.rsSel_ManifestoPorCTC.State = 1 Then de_informa.rsSel_ManifestoPorCTC.Close
de_informa.Sel_ManifestoPorCTC lblFilialctc
gridManifesto.DataMember = "sel_manifestoporctc"
gridManifesto.Refresh
If de_informa.rsSel_ManifestoPorCTC.RecordCount > 0 Then
    gridManifesto.Enabled = True
Else
    gridManifesto.Enabled = False
End If
    
'Observação de Emissão e de Ocorr/Entrega do SITLA
        
lblObs_Emissao = de_informa.rsSel_Ctc_SAC.Fields("obs_emissao")
lblObs_OcorrSitla = de_informa.rsSel_Ctc_SAC.Fields("obs_ocorr")
        
DoEvents
     
'volta o foco para a Filial
If optCTC = True Then
    txtFilial.SetFocus
ElseIf optNf = True Then
    txtNumNf.SetFocus
End If
        
'ATUALIZA GRID DE ÚLTIMAS CONSULTAS DE USUÁRIOS

If de_informa.rsSel_Ultconssac.State = 1 Then de_informa.rsSel_Ultconssac.Close
de_informa.Sel_Ultconssac lblFilialctc
gridUltCons.DataMember = "sel_ultconssac"
gridUltCons.Refresh

'atualiza usuário e hora que consultou
xdata = CVar(xagora)
de_informa.ins_ultconssac lblFilialctc, xusuario, xdata
        
'LOG DE USUÁRIO
de_informa.ins_LogUsuario "CONSULTA", xusuario, "INFORMAÇÃO SAC - CONSULTA CTC: " & lblFilialctc
    
End Sub
Private Sub cmbSair_Click()
    Unload Me
End Sub

Private Sub cmdFinalizaSolic_Click()
    fraDados.Enabled = True
    fraSolicitacao.Enabled = False
    fraHistorico.Enabled = False
    cmdGravaCancela.Enabled = True
    txtSolicitante.Enabled = False
    txtFone.Enabled = False
    txtEmail.Enabled = False
    'txtSolicitante.BackColor = &HC0FFFF    'AMARELO
    'txtFone.BackColor = &HC0FFFF   'AMARELO
    'txtEmail.BackColor = &HC0FFFF   'AMARELO
    txtDescricao.BackColor = &HC0FFFF   'AMARELO
    txtObservacao.BackColor = &HC0FFFF 'AMARELO
    txtDescricao.SetFocus
End Sub

Private Sub cmdGravaCancela_Click()
    fraDados.Enabled = False
    fraSolicitacao.Enabled = True
    fraHistorico.Enabled = True
    cmdGravaCancela.Enabled = False
    txtSolicitante.BackColor = &H80000005     'branco
    txtFone.BackColor = &H80000005     'branco
    txtEmail.BackColor = &H80000005     'branco
    txtDescricao.BackColor = &H80000005     'branco
    txtObservacao.BackColor = &H80000005     'branco
End Sub

Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdNovaOcorrInf_Click()
    fraDados.Enabled = True
    fraSolicitacao.Enabled = False
    fraHistorico.Enabled = False
    cmdGravaCancela.Enabled = True
    txtSolicitante.Enabled = False
    txtFone.Enabled = False
    txtEmail.Enabled = False
    'txtSolicitante.BackColor = &HC0FFFF    'AMARELO
    'txtFone.BackColor = &HC0FFFF   'AMARELO
    'txtEmail.BackColor = &HC0FFFF   'AMARELO
    txtDescricao.BackColor = &HC0FFFF   'AMARELO
    txtObservacao.BackColor = &HC0FFFF 'AMARELO
    txtDescricao.SetFocus
End Sub

Private Sub cmdNovasolic_Click()
    fraDados.Enabled = True
    fraSolicitacao.Enabled = False
    fraHistorico.Enabled = False
    cmdGravaCancela.Enabled = True
    txtSolicitante.Enabled = True
    txtSolicitante.BackColor = &HC0FFFF     'AMARELO
    txtFone.BackColor = &HC0FFFF            'AMARELO
    txtEmail.BackColor = &HC0FFFF           'AMARELO
    txtDescricao.BackColor = &HC0FFFF       'AMARELO
    txtObservacao.BackColor = &HC0FFFF      'AMARELO
    txtSolicitante.SetFocus
End Sub

Private Sub cmdPesq_Click()
    optCTC.Value = True
    optCTC_Click
    DoEvents
    frmPesquisaCTC.Show 1
End Sub
Private Sub cmdRastPrevEntr_Click()
    frmRastrPrazo.lblFilialctc = transctc(frmSac.txtFilial, frmSac.txtCtc)
    frmRastrPrazo.lblModal = frmSac.lblModal
    frmRastrPrazo.lblCidadeDest = frmSac.lblCidade_Dest
    frmRastrPrazo.lblUFdest = frmSac.lblUf_Dest
    frmRastrPrazo.lblEmissao = frmSac.lblData
    frmRastrPrazo.lblHsEmiss = frmSac.lblHora
    frmRastrPrazo.lblEntrega = lblPrevEntr.Caption
    frmRastrPrazo.lblPrazo = lblPrazoDias.Caption
    frmRastrPrazo.Caption = "Rastrear Cálculo de Previsão de Entrega"
    frmRastrPrazo.Frame1.Caption = "Dados do CTC / Previsão de Entrega"
    frmRastrPrazo.Label7.Caption = "Previsão"
    frmRastrPrazo.Label18.Visible = False
    frmRastrPrazo.Label13.Visible = False
    frmRastrPrazo.lblHsEntr.Visible = False
    frmRastrPrazo.lblMeta.Visible = False
    DoEvents
    frmRastrPrazo.Show 1
End Sub
Private Sub cmdRastrPrazo_Click()
    de_informa.Alt_AtualPrazoSCTC transctc(frmSac.txtFilial, frmSac.txtCtc)
'    frmAtualPrazos.Show 1
    If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
    de_informa.Sel_ConsOcorr transctc(frmSac.txtFilial, frmSac.txtCtc), "01"
    frmRastrPrazo.lblFilialctc = transctc(frmSac.txtFilial, frmSac.txtCtc)
    frmRastrPrazo.lblModal = frmSac.lblModal
    frmRastrPrazo.lblCidadeDest = frmSac.lblCidade_Dest
    frmRastrPrazo.lblUFdest = frmSac.lblUf_Dest
    frmRastrPrazo.lblEmissao = frmSac.lblData
    frmRastrPrazo.lblHsEmiss = frmSac.lblHora
    frmRastrPrazo.lblEntrega = de_informa.rsSel_ConsOcorr.Fields("data")
    frmRastrPrazo.lblHsEntr = de_informa.rsSel_ConsOcorr.Fields("hora")
    frmRastrPrazo.lblMeta = de_informa.rsSel_ConsOcorr.Fields("prazoentr")
    frmRastrPrazo.lblPrazo = de_informa.rsSel_ConsOcorr.Fields("diasuteis")
'    frmConsOcorr.lblMetaPrazo = Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("prazoentr"))) & "/" & Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("diasuteis")))
    DoEvents
    frmRastrPrazo.Show 1
End Sub
Private Sub Form_Activate()
    txtFilial.Text = xultimofilial
    txtCtc.Text = xultimoctc
    txtNumNf.Text = xultimonf
    If optCTC = True Then
        txtFilial.SetFocus
    Else
        txtNumNf.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    GridConsOcorr.DataMember = ""
    gridUltCons.DataMember = ""
    gridCTCManifesto.DataMember = ""
    gridManifesto.DataMember = ""
    gridNfsdoCTC.DataMember = ""
    cmbProcurar.Enabled = True
    cmbSair.Enabled = True
    optNf_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmSACNovo = Nothing
End Sub

Private Sub GridConsOcorr_Click()
    'If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
    'atualiza o campo de obs de ocorrência quando clicado no grid
        lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
    'End If
End Sub

Private Sub gridManifesto_Click()

Dim xvolumes As Long, xpeso As Currency, xvalmerc As Currency, xfretetotal As Currency

Me.MousePointer = 11

'ATUALIZA RECORDSET QUE IRÁ ATUALIZAR O GRID DE CTCS DO MANIFESTO
    If de_informa.rsSel_CTCsdoManifesto.State = 1 Then de_informa.rsSel_CTCsdoManifesto.Close
    de_informa.Sel_CTCsdoManifesto gridManifesto.Columns(2)
    
'ATUALIZA DADOS ESTATÍSTICOS DE MANIFESTO
    
    de_informa.rsSel_CTCsdoManifesto.MoveFirst
    xvolumes = 0
    xpeso = 0
    xvalmerc = 0
    xfretetotal = 0
    Do Until de_informa.rsSel_CTCsdoManifesto.EOF
        xvolumes = xvolumes + de_informa.rsSel_CTCsdoManifesto.Fields("volumes")
        xpeso = xpeso + de_informa.rsSel_CTCsdoManifesto.Fields("peso")
        xvalmerc = xvalmerc + de_informa.rsSel_CTCsdoManifesto.Fields("valmerc")
        xfretetotal = xfretetotal + de_informa.rsSel_CTCsdoManifesto.Fields("fretetotal")
        de_informa.rsSel_CTCsdoManifesto.MoveNext
    Loop
    
    'Busca Qtdes. de NFs
    
    If de_informa.rsSel_NfsdoManifesto.State = 1 Then de_informa.rsSel_NfsdoManifesto.Close
    de_informa.Sel_NfsdoManifesto gridManifesto.Columns(2)
    
    'ATUALIZA LABEL DE ESTATÍSTICAS
    
    lblManifEstatCTC = de_informa.rsSel_CTCsdoManifesto.RecordCount
    lblManifEstatNF = de_informa.rsSel_NfsdoManifesto.Fields("qtd")
    lblManifEstatVOL = xvolumes
    lblManifEstatPESO = Format(xpeso, "##,##0.0")
    lblManifEstatVALMERC = Format(xvalmerc, "##,###,##0.00")
    lblManifEstatFRETE = Format(xfretetotal, "##,###,##0.00")
    lblManifEstatIndNfCtc = Format(Val(lblManifEstatNF) / Val(lblManifEstatCTC), "#,##0.0")
    lblManifEstatIndFretePeso = Format(xfretetotal / xpeso, "#,##0.00")
    lblManifEstatIndFreteNf = Format(xfretetotal / xvalmerc, "##0.00%")
    
    
    'ATUALIZA O GRID DE MANIFESTO
    
    gridCTCManifesto.DataMember = "Sel_CTCsdoManifesto"
    gridCTCManifesto.Refresh

    
Me.MousePointer = 0
End Sub

Private Sub txtCtc_Change()
    On Error Resume Next
    If Len(txtCtc.Text) >= 8 Then cmbProcurar.SetFocus
End Sub
Private Sub txtCTC_GotFocus()
    txtCtc.SelStart = 0
    txtCtc.SelLength = 8
End Sub
Private Sub txtCTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtCtc_LostFocus()
    If txtCtc.Text <> "" Then
        If Not IsNumeric(txtCtc.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtCtc.SetFocus
            Exit Sub
        End If
    End If
    txtCtc.Text = Trim(txtCtc.Text)
End Sub

Private Sub txtDescricao_Change()
    If Len(txtSolicitante) > 0 And Len(txtDescricao) > 0 Then
        cmdGrava.Enabled = True
    Else
        cmdGrava.Enabled = False
    End If
End Sub

Private Sub txtDescricao_GotFocus()
    cmdf1.Enabled = True
    cmdf2.Enabled = True
    cmdf3.Enabled = True
    cmdf4.Enabled = True
    cmdf5.Enabled = True
    cmdX.Enabled = True
End Sub

Private Sub txtDescricao_LostFocus()
    cmdf1.Enabled = False
    cmdf2.Enabled = False
    cmdf3.Enabled = False
    cmdf4.Enabled = False
    cmdf5.Enabled = False
    cmdX.Enabled = False
End Sub

Private Sub txtfilial_Change()
    On Error Resume Next
    If Len(txtFilial.Text) >= 2 Then txtCtc.SetFocus
End Sub
Private Sub txtfilial_gotfocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub
Private Sub optCTC_Click()
    On Error Resume Next
    'fraProcura.Caption = "Núm. Filial e CTC"
    txtFilial.Visible = True
    txtCtc.Visible = True
    txtNumNf.Visible = False
    txtFilial.SetFocus
End Sub
Private Sub optNf_Click()
    On Error Resume Next
    txtFilial.Visible = False
    txtCtc.Visible = False
    txtNumNf.Visible = True
    txtNumNf.SetFocus
End Sub
Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        If txtFilial.Text = "" Then
            KeyAscii = 0
            optNf.Value = True
            optNf_Click
        Else
            KeyAscii = 0
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub
Private Sub txtFilial_LostFocus()
    If txtFilial.Text <> "" Then
        If Not IsNumeric(txtFilial.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtFilial.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtNumNf_GotFocus()
    txtNumNf.SelStart = 0
    txtNumNf.SelLength = 12
End Sub
Private Sub txtNumNf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        If txtNumNf.Text = "" Then
            KeyAscii = 0
            optCTC.Value = True
            optCTC_Click
        Else
            KeyAscii = 0
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub
Private Sub txtNumNf_LostFocus()
    If txtNumNf.Text <> "" Then
        If Not IsNumeric(txtNumNf.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtNumNf.SetFocus
            Exit Sub
        End If
    End If
End Sub


