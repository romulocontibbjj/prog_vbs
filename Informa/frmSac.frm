VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSac 
   Caption         =   "SAC - Informação de Transporte"
   ClientHeight    =   7935
   ClientLeft      =   1635
   ClientTop       =   1440
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
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
      Height          =   720
      Left            =   120
      TabIndex        =   196
      Top             =   0
      Width           =   1800
      Begin VB.OptionButton optCTC 
         Caption         =   "Núm. de Ctc/Ctr"
         Height          =   270
         Left            =   105
         TabIndex        =   198
         Top             =   210
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optNf 
         Caption         =   "Núm. de N.Fiscal"
         Height          =   270
         Left            =   105
         TabIndex        =   197
         Top             =   420
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdAWB 
      Caption         =   "AWB Inf."
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox chkAutoScan 
      Caption         =   "Scanner Auto Load"
      Height          =   195
      Left            =   5040
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmbSair 
      Caption         =   "Sair"
      Height          =   315
      Left            =   6840
      TabIndex        =   8
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   375
      Left            =   6240
      Picture         =   "frmSac.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdPesq 
      Caption         =   "Pesquisa"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   50
      Width           =   1095
   End
   Begin VB.CommandButton cmbProcurar 
      Caption         =   "Procurar"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame fraProcura 
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2040
      TabIndex        =   52
      Top             =   0
      Width           =   1815
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   330
      End
      Begin VB.TextBox txtCtc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   600
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox txtNumNf 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10186
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Dados 1"
      TabPicture(0)   =   "frmSac.frx":0772
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TabOcorrencias"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraPODScan"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame22"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdImagem"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Dados 2"
      TabPicture(1)   =   "frmSac.frx":078E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame26"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "fraNfs"
      Tab(1).Control(4)=   "Frame11"
      Tab(1).Control(5)=   "Frame7"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Manifesto"
      TabPicture(2)   =   "frmSac.frx":07AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame17"
      Tab(2).Control(2)=   "Frame14"
      Tab(2).Control(3)=   "Frame12"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdImagem 
         Caption         =   "POD SCANNER ..."
         Height          =   1740
         Left            =   8760
         TabIndex        =   195
         Top             =   440
         Width           =   2895
      End
      Begin VB.Frame Frame26 
         Caption         =   "Composição do Frete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1700
         Left            =   -74880
         TabIndex        =   53
         Top             =   3960
         Width           =   8175
         Begin VB.Label lblPagtoFatura 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   207
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Pagto:"
            Height          =   195
            Left            =   5520
            TabIndex        =   206
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label lblEmissorFatura 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7080
            TabIndex        =   205
            Top             =   1320
            Width           =   1020
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Vencto:"
            Height          =   195
            Left            =   3840
            TabIndex        =   204
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label lblVencFatura 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4440
            TabIndex        =   203
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   2040
            TabIndex        =   202
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label lblDataFatura 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2760
            TabIndex        =   201
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label lblFatura 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   200
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Faturamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   199
            Top             =   1320
            Width           =   930
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   8040
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Frete CIF/FOB:"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label lblfpag 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   149
            Top             =   300
            Width           =   855
         End
         Begin VB.Label lblDescrOutros 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6480
            TabIndex        =   104
            Top             =   900
            Width           =   1620
         End
         Begin VB.Label lbltxUrgencia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   54
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label lblTxEntrega 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   55
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label lblTxColeta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5400
            TabIndex        =   56
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblAdValorem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3480
            TabIndex        =   57
            Top             =   600
            Width           =   1000
         End
         Begin VB.Label lblFreteNacional 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3480
            TabIndex        =   58
            Top             =   300
            Width           =   1000
         End
         Begin VB.Label lblPedagio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7200
            TabIndex        =   103
            Top             =   300
            Width           =   885
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Pedágio:"
            Height          =   195
            Left            =   6480
            TabIndex        =   102
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Frete Total Br:"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label lblFreteBr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   86
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tx.Coleta:"
            Height          =   195
            Left            =   4560
            TabIndex        =   69
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "GRIS:"
            Height          =   195
            Left            =   2640
            TabIndex        =   68
            Top             =   900
            Width           =   435
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Tx.Entrega:"
            Height          =   195
            Left            =   4560
            TabIndex        =   67
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Tx.Urgênc:"
            Height          =   195
            Left            =   4560
            TabIndex        =   66
            Top             =   900
            Width           =   795
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Frete Valor:"
            Height          =   195
            Left            =   2640
            TabIndex        =   65
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Frete Peso:"
            Height          =   195
            Left            =   2640
            TabIndex        =   64
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Outros:"
            Height          =   195
            Left            =   6480
            TabIndex        =   63
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lblGris 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3480
            TabIndex        =   62
            Top             =   900
            Width           =   1000
         End
         Begin VB.Label lblTxOutros 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7200
            TabIndex        =   61
            Top             =   600
            Width           =   885
         End
         Begin VB.Label lblFreteLiq 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   60
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Frete Total Líq:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   1095
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
         Height          =   2340
         Left            =   -66600
         TabIndex        =   70
         Top             =   1560
         Width           =   3255
         Begin VB.Label lblVolumes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   71
            Top             =   990
            Width           =   735
         End
         Begin VB.Label lblPeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   72
            Top             =   690
            Width           =   735
         End
         Begin VB.Label lblPesotax 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1920
            TabIndex        =   101
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Volumes:"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   990
            Width           =   840
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor NFs:"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   390
            Width           =   735
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Dimensões:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1290
            Width           =   825
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Natureza:"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   1890
            Width           =   690
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   1590
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Peso/Tax:"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   690
            Width           =   750
         End
         Begin VB.Label lblDimensoes 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   73
            Top             =   1290
            Width           =   1575
         End
         Begin VB.Label lblEspecie 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   74
            Top             =   1590
            Width           =   1575
         End
         Begin VB.Label lblNatureza 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   75
            Top             =   1890
            Width           =   2055
         End
         Begin VB.Label lblValmerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   76
            Top             =   390
            Width           =   1575
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
         Height          =   1695
         Left            =   -66600
         TabIndex        =   147
         Top             =   3960
         Width           =   3255
         Begin MSDataGridLib.DataGrid gridUltCons 
            Bindings        =   "frmSac.frx":07C6
            Height          =   1335
            Left            =   120
            TabIndex        =   148
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2355
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
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
                  ColumnWidth     =   1544,882
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   -1  'True
                  Locked          =   -1  'True
                  ColumnWidth     =   1230,236
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Detalhe do Manifesto (clique no manifesto)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   -74880
         TabIndex        =   146
         Top             =   2280
         Width           =   5775
         Begin VB.CheckBox chkMnfCamFria 
            Caption         =   "Cam.Fria"
            Height          =   195
            Left            =   4680
            TabIndex        =   163
            Top             =   1665
            Width           =   975
         End
         Begin VB.CheckBox chkMnfPlatHidr 
            Caption         =   "Plat.Hidr."
            Height          =   195
            Left            =   3720
            TabIndex        =   162
            Top             =   1665
            Width           =   975
         End
         Begin VB.CheckBox chkMnfSuspAr 
            Caption         =   "Susp. AR"
            Height          =   195
            Left            =   2640
            TabIndex        =   161
            Top             =   1665
            Width           =   975
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   5640
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   5640
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Label lblMnfMotorista 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   184
            Top             =   2040
            Width           =   4335
         End
         Begin VB.Label Label90 
            AutoSize        =   -1  'True
            Caption         =   "Motorista:"
            Height          =   195
            Left            =   120
            TabIndex        =   183
            Top             =   2040
            Width           =   690
         End
         Begin VB.Label lblMnfConferente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   182
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label88 
            AutoSize        =   -1  'True
            Caption         =   "Conferente:"
            Height          =   195
            Left            =   120
            TabIndex        =   181
            Top             =   600
            Width           =   825
         End
         Begin VB.Label lblMnfEmissor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4320
            TabIndex        =   180
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            Caption         =   "Emissor:"
            Height          =   195
            Left            =   3720
            TabIndex        =   179
            Top             =   600
            Width           =   585
         End
         Begin VB.Label lblMnfHora 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2880
            TabIndex        =   177
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblMnfEmissao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   176
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   2400
            TabIndex        =   175
            Top             =   300
            Width           =   390
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   174
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            Caption         =   "Ajudantes:"
            Height          =   195
            Left            =   120
            TabIndex        =   173
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Vencto. da CNH:"
            Height          =   195
            Left            =   3000
            TabIndex        =   172
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblMnfMotoristaCNH 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   171
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblMnfMotoristaCNHVencto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4440
            TabIndex        =   170
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblMnfMotoristaAdmissao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4440
            TabIndex        =   169
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label lblMnfAjudantes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   168
            Top             =   2940
            Width           =   4335
         End
         Begin VB.Label lblMnfMotoristaCateg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   167
            Top             =   2340
            Width           =   1575
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "Data de Admissão:"
            Height          =   195
            Left            =   3000
            TabIndex        =   166
            Top             =   2340
            Width           =   1335
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "Num. da CNH:"
            Height          =   195
            Left            =   120
            TabIndex        =   165
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Categoria:"
            Height          =   195
            Left            =   120
            TabIndex        =   164
            Top             =   2340
            Width           =   720
         End
         Begin VB.Label lblMnfCodVeic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   160
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label lblMnfRastream 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   159
            Top             =   1620
            Width           =   1215
         End
         Begin VB.Label lblMnfMarcaVeic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   158
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblMnfTipoVeic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4320
            TabIndex        =   157
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblMnfPropriet 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2880
            TabIndex        =   156
            Top             =   1020
            Width           =   2775
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Veic:"
            Height          =   195
            Left            =   3480
            TabIndex        =   155
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "Marca/Modelo:"
            Height          =   195
            Left            =   120
            TabIndex        =   154
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "Rastreamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   153
            Top             =   1620
            Width           =   1035
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário:"
            Height          =   195
            Left            =   1920
            TabIndex        =   152
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Veículo:"
            Height          =   195
            Left            =   120
            TabIndex        =   151
            Top             =   1020
            Width           =   975
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Estatística do Manifesto (clique no manifesto)"
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
         Left            =   -69000
         TabIndex        =   127
         Top             =   4320
         Width           =   5655
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "NFs/CTCs:"
            Height          =   195
            Left            =   120
            TabIndex        =   145
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Frete/Peso:"
            Height          =   195
            Left            =   1800
            TabIndex        =   144
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Frete/Val.NF:"
            Height          =   195
            Left            =   3720
            TabIndex        =   143
            Top             =   960
            Width           =   960
         End
         Begin VB.Label lblManifEstatIndFreteNf 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4800
            TabIndex        =   142
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblManifEstatIndFretePeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2760
            TabIndex        =   141
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblManifEstatIndNfCtc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   140
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblManifEstatFRETE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4200
            TabIndex        =   139
            Top             =   620
            Width           =   1335
         End
         Begin VB.Label lblManifEstatNF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   138
            Top             =   620
            Width           =   735
         End
         Begin VB.Label lblManifEstatPESO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   137
            Top             =   620
            Width           =   855
         End
         Begin VB.Label lblManifEstatVOL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   136
            Top             =   280
            Width           =   855
         End
         Begin VB.Label lblManifEstatVALMERC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4200
            TabIndex        =   135
            Top             =   280
            Width           =   1335
         End
         Begin VB.Label lblManifEstatCTC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   134
            Top             =   280
            Width           =   735
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Valor Frete:"
            Height          =   195
            Left            =   3240
            TabIndex        =   133
            Top             =   620
            Width           =   810
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
            Height          =   195
            Left            =   1440
            TabIndex        =   132
            Top             =   620
            Width           =   405
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Valor NFs:"
            Height          =   195
            Left            =   3240
            TabIndex        =   131
            Top             =   280
            Width           =   735
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "NFs:"
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   620
            Width           =   330
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Volumes:"
            Height          =   195
            Left            =   1440
            TabIndex        =   129
            Top             =   280
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CTCs:"
            Height          =   195
            Left            =   120
            TabIndex        =   128
            Top             =   280
            Width           =   435
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Documentos do Manifesto (clique no manifesto)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3880
         Left            =   -69000
         TabIndex        =   125
         Top             =   360
         Width           =   5655
         Begin MSDataGridLib.DataGrid gridCTCManifesto 
            Bindings        =   "frmSac.frx":07DF
            Height          =   3375
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   5953
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
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
         Height          =   1815
         Left            =   -74880
         TabIndex        =   123
         Top             =   360
         Width           =   5775
         Begin MSDataGridLib.DataGrid gridManifesto 
            Bindings        =   "frmSac.frx":07F8
            Height          =   1455
            Left            =   120
            TabIndex        =   178
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   2566
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
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
               Caption         =   "Manifesto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "filial"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "manifesto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "Data Manif."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "Hora"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "Data Saída"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "Hora"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   540,284
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   659,906
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   945,071
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   615,118
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   1665,071
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   870,236
               EndProperty
               BeginProperty Column14 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            Height          =   2175
            Left            =   3840
            TabIndex        =   124
            Top             =   2400
            Width           =   5895
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados do Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         TabIndex        =   115
         Top             =   330
         Width           =   8535
         Begin VB.Label lblEntregaUF 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8040
            TabIndex        =   215
            Top             =   870
            Width           =   375
         End
         Begin VB.Label lblColetaUF 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3840
            TabIndex        =   214
            Top             =   870
            Width           =   375
         End
         Begin VB.Label lblEntrega 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5040
            TabIndex        =   213
            Top             =   870
            Width           =   3015
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Entrega:"
            Height          =   195
            Left            =   4320
            TabIndex        =   212
            Top             =   870
            Width           =   600
         End
         Begin VB.Label lblColeta 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   211
            Top             =   870
            Width           =   3015
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Coleta:"
            Height          =   195
            Left            =   120
            TabIndex        =   210
            Top             =   870
            Width           =   495
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Emissor:"
            Height          =   195
            Left            =   3240
            TabIndex        =   209
            Top             =   270
            Width           =   585
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   2040
            TabIndex        =   208
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblVia 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7800
            TabIndex        =   194
            Top             =   570
            Width           =   615
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   5320
            TabIndex        =   193
            Top             =   570
            Width           =   480
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Via:"
            Height          =   195
            Left            =   7440
            TabIndex        =   192
            Top             =   570
            Width           =   270
         End
         Begin VB.Label lblModal 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5880
            TabIndex        =   191
            Top             =   570
            Width           =   1455
         End
         Begin VB.Label lblTranspsubRedesp 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   190
            Top             =   570
            Width           =   2955
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Redesp:"
            Height          =   195
            Left            =   120
            TabIndex        =   186
            Top             =   570
            Width           =   600
         End
         Begin VB.Label lblTranspSub 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   185
            Top             =   570
            Width           =   1455
         End
         Begin VB.Label lblTipodoc 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5880
            TabIndex        =   122
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   5320
            TabIndex        =   121
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lblData 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   840
            TabIndex        =   120
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label lblHora 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2520
            TabIndex        =   119
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   118
            Top             =   270
            Width           =   630
         End
         Begin VB.Label lblEmissor 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3960
            TabIndex        =   117
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label lblPrioridade 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NORMAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6960
            TabIndex        =   116
            Top             =   270
            Width           =   1455
         End
      End
      Begin VB.Frame fraNfs 
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
         Height          =   2400
         Left            =   -74880
         TabIndex        =   84
         Top             =   1560
         Width           =   8175
         Begin MSDataGridLib.DataGrid gridNfsdoCTC 
            Bindings        =   "frmSac.frx":0811
            Height          =   2055
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            Enabled         =   -1  'True
            ForeColor       =   0
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
            ColumnCount     =   7
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
               DataField       =   "serie"
               Caption         =   "Série"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "valornf"
               Caption         =   "Valor NF"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
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
            BeginProperty Column04 
               DataField       =   "canhotonfprot"
               Caption         =   "Protocolo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "canhotonfdata"
               Caption         =   "Data Prot."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "id_local"
               Caption         =   "ID_Local"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   540,284
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1319,811
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   854,929
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1844,787
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   929,764
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Observação de Emissão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   -74880
         TabIndex        =   82
         Top             =   360
         Width           =   7455
         Begin VB.Label lblObs_Emissao 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   7215
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
         Height          =   1185
         Left            =   -67320
         TabIndex        =   49
         Top             =   360
         Width           =   3975
         Begin VB.Label lblRespons_Nome 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label lblRespons_CGC 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   1755
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
         Height          =   1455
         Left            =   6840
         TabIndex        =   47
         Top             =   4200
         Width           =   4815
         Begin VB.Label lblObsEntr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   1125
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   4560
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
         Height          =   1455
         Left            =   3480
         TabIndex        =   34
         Top             =   4200
         Width           =   3315
         Begin VB.Label lblDtUsu 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   35
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblReceb 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   46
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblHsBaixa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblDtBaixa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   38
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   525
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   1920
            TabIndex        =   40
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Quem Baixou:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   990
         End
         Begin VB.Label lblUsu_bx 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   37
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Data Baixa:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   825
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
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   3315
         Begin VB.Label lblDtUsuPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   26
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblRecebPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   43
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblDtBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   28
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label lblHsBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   44
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   525
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   1920
            TabIndex        =   31
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Quem Baixou:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   990
         End
         Begin VB.Label lblUsu_bxpre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   29
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Em: (data/hs)"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2025
         Left            =   8760
         TabIndex        =   24
         Top             =   2140
         Width           =   2895
         Begin VB.CommandButton cmdMostraAbono 
            Caption         =   "?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   100
            Top             =   900
            Width           =   495
         End
         Begin VB.CommandButton cmdRastrPrazo 
            Caption         =   "?"
            Height          =   255
            Left            =   2520
            TabIndex        =   94
            Top             =   560
            Width           =   255
         End
         Begin VB.CommandButton cmdRastPrevEntr 
            Caption         =   "?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2520
            TabIndex        =   89
            Top             =   225
            Width           =   255
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   2760
            Y1              =   1250
            Y2              =   1250
         End
         Begin VB.Label lblNumProt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   108
            Top             =   1635
            Width           =   1335
         End
         Begin VB.Label lblArquivo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   107
            Top             =   1330
            Width           =   1335
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Núm. do Protoc.:"
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   1635
            Width           =   1200
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Físico p/ Arquivo:"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   1330
            Width           =   1275
         End
         Begin VB.Label lblAbonoDias 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1425
            TabIndex        =   99
            Top             =   900
            Width           =   375
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Abono de Atraso:"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   900
            Width           =   1230
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Prazo Real:"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   555
            Width           =   825
         End
         Begin VB.Label lblTabPrazo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   96
            Top             =   555
            Width           =   975
         End
         Begin VB.Label lblPrazoDias 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   95
            Top             =   555
            Width           =   375
         End
         Begin VB.Label lblLabelMeta 
            AutoSize        =   -1  'True
            Caption         =   "Meta:"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   225
            Width           =   405
         End
         Begin VB.Label lblPrevEntr 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   92
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Dias"
            Height          =   195
            Left            =   1845
            TabIndex        =   91
            Top             =   900
            Width           =   315
         End
         Begin VB.Label lblMetaPrazo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   90
            Top             =   225
            Width           =   375
         End
      End
      Begin VB.Frame fraPODScan 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   8760
         TabIndex        =   187
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
         Begin VB.Label Label93 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não Disponível"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   480
            TabIndex        =   189
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label92 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "        POD SCANNER         "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   0
            TabIndex        =   188
            Top             =   0
            Width           =   2850
         End
      End
      Begin TabDlg.SSTab TabOcorrencias 
         Height          =   2535
         Left            =   120
         TabIndex        =   216
         Top             =   1620
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4471
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Ocorrências"
         TabPicture(0)   =   "frmSac.frx":082A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblObs_Ocorr"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "GridConsOcorr"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Status Dispara"
         TabPicture(1)   =   "frmSac.frx":0846
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "gridDispara"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid GridConsOcorr 
            Bindings        =   "frmSac.frx":0862
            Height          =   1215
            Left            =   120
            TabIndex        =   217
            Top             =   360
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2143
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            ForeColor       =   0
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
         Begin MSDataGridLib.DataGrid gridDispara 
            Bindings        =   "frmSac.frx":087B
            Height          =   2055
            Left            =   -74880
            TabIndex        =   219
            Top             =   360
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            ForeColor       =   0
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
            DataMember      =   "Sel_StatusDispara"
            ColumnCount     =   5
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "hora"
               Caption         =   "Hora"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "tipostatus"
               Caption         =   "Tipo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "descricao"
               Caption         =   "Descrição"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "id_usuario"
               Caption         =   "Usuario (ID)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   1275,024
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   720
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   3270,047
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   1275,024
               EndProperty
            EndProperty
         End
         Begin VB.Label lblObs_Ocorr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   820
            Left            =   120
            TabIndex        =   218
            Top             =   1620
            Width           =   8295
         End
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1725
         Left            =   8760
         Picture         =   "frmSac.frx":0894
         Stretch         =   -1  'True
         Top             =   435
         Visible         =   0   'False
         Width           =   2865
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   6120
      TabIndex        =   15
      Top             =   720
      Width           =   5775
      Begin VB.Label lblDest_Nome 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   420
         Width           =   4575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   675
         Width           =   690
      End
      Begin VB.Label lblEndDest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   113
         Top             =   675
         Width           =   4575
      End
      Begin VB.Label lblDest_CGC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   112
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Localidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   930
         Width           =   825
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Destinatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblCidade_Dest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   930
         Width           =   4050
      End
      Begin VB.Label lblUf_Dest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5235
         TabIndex        =   23
         Top             =   930
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Origem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   5775
      Begin VB.Label lblRemet_Nome 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   420
         Width           =   4575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   111
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblEndRem 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   110
         Top             =   675
         Width           =   4575
      End
      Begin VB.Label lblRemet_CGC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   109
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Localidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblCidade_orig 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   930
         Width           =   4050
      End
      Begin VB.Label lblUF_Orig 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5235
         TabIndex        =   20
         Top             =   930
         Width           =   420
      End
   End
   Begin VB.Frame Frame15 
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
      Left            =   7800
      TabIndex        =   9
      Top             =   0
      Width           =   4095
      Begin VB.Label lblEntregueSN 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   70
         TabIndex        =   10
         Top             =   360
         Width           =   3930
      End
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   840
      Y2              =   1920
   End
End
Attribute VB_Name = "frmSac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public xtemocorr As String

Private Sub chkAutoScan_Click()
    If chkAutoScan.Value = 1 Then
        cmdImagem.Visible = False
        fraPODScan.Visible = True
        Image1.Visible = True
    Else
        cmdImagem.Visible = True
        fraPODScan.Visible = False
        Image1.Visible = False
    End If
    DoEvents
End Sub

Private Sub cmbProcurar_Click()
Dim xentrega As String, xhora_ent As String, xrecebedor As String, xobs_ocorr As String, xdata As String, xflexnf As Integer
Dim xvialink As String, xbuscaprazo As String, xprazo_TT As Integer, xfilialctc As String

    'limpa imagem
    Image1.Picture = LoadPicture(App.Path & "\semscan.jpg")
   
'    If de_informa.rsSel_NFBasecli.State = 1 Then de_informa.rsSel_NFBasecli.Close
'    gridClienteEspeciais.DataMember = "Sel_NFBasecli"
'    gridClienteEspeciais.Refresh
    
'    If de_informa.rsSel_NFVideolar.State = 1 Then de_informa.rsSel_NFVideolar.Close
'    gridVideolarEspeciais.DataMember = "Sel_NFVideolar"
'    gridVideolarEspeciais.Refresh
    
'    If de_informa.rsSel_CheckReceb.State = 1 Then de_informa.rsSel_CheckReceb.Close
'    gridCheckEspeciais.DataMember = "Sel_CheckReceb"
'    gridCheckEspeciais.Refresh
    

'    If SSTab1.Tab = 4 Then SSTab1.Tab = 0

    If TxtFilial.Visible = True Then
        TxtFilial.SetFocus
    End If
    
    
    If optCTC.Value = True Then   'Se for procura por Filial / CTC
        If TxtFilial.Text = "" Or txtCtc.Text = "" Then
            MsgBox "Filial / CTC Inválidos !", vbCritical, "Erro"
        End If
        If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
        de_informa.Sel_Ctc_SAC transctc(TxtFilial, txtCtc) 'Procura na Tabela a Filial/CTC
        If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then
            MsgBox "Número de Filial/CTC Não Encontrados !", vbCritical + vbOKOnly, "Erro"
            TxtFilial.SetFocus
            Exit Sub
        End If
    ElseIf optNf.Value = True Then  'Se a procura for por NF
        If txtNumNf.Text = "" Then
            MsgBox "Número de Nota Fiscal Inválida !", vbCritical, "Erro"
        End If
            If de_informa.rsSel_NF_SAC.State = 1 Then de_informa.rsSel_NF_SAC.Close
            de_informa.Sel_NF_SAC Val(txtNumNf)   'Procura a NF na Tabela
            If de_informa.rsSel_NF_SAC.RecordCount = 0 Then
                MsgBox "Número de NF Não Encontrado !", vbCritical + vbOKOnly, "Erro"
                txtNumNf.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_NF_SAC.RecordCount > 1 Then
                frmDuplNF.Caption = "SAC - Número de NFs Duplicadas"
                DoEvents
                frmDuplNF.Show 1  'direciona para o form que trata casos de NF duplicadas
                Exit Sub
            Else  'Caso seja encontrada somente uma NF
                optCTC_Click
                TxtFilial.Text = Mid(de_informa.rsSel_NF_SAC.Fields("filialctc"), 1, 2)
                txtCtc.Text = Mid(de_informa.rsSel_NF_SAC.Fields("filialctc"), 3, 8) 'Busca a Filial e o CTC com base na NF
                If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
                de_informa.Sel_Ctc_SAC transctc(TxtFilial, txtCtc)
                If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then 'Caso não encontre erro de consistência
                    MsgBox "Erro de Consistência ! Avise o Suporte de Informática !"
                    Exit Sub
                End If
            End If
    End If
    optCTC.Value = True
    
'busca prazo de entrega para este cliente
    If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
    de_informa.Sel_ConsCadCli de_informa.rsSel_Ctc_SAC.Fields("remet_cgc")
    
'pegar endereço do destinatário
    If de_informa.rsSel_ConsCadCliDest.State = 1 Then de_informa.rsSel_ConsCadCliDest.Close
    de_informa.Sel_ConsCadCliDest de_informa.rsSel_Ctc_SAC.Fields("dest_cgc")
    
'grava variável global com número da filial e ctc para utilizá-los em outros forms
    xultimofilial = TxtFilial.Text
    xultimoctc = txtCtc.Text
    
    xfilialctc = transctc(frmSac.TxtFilial.Text, frmSac.txtCtc.Text)

'limpa lbls de entrega

    cmbProcurar.Caption = "Aguarde..."
    cmbProcurar.Enabled = False
    DoEvents

    lblNumProt = ""
    lblArquivo = ""
    lblDtBaixaPre = ""
    lblHsBaixaPre = ""
    lblRecebPre = ""
    lblUsu_bxpre = ""
    lblDtBaixa = ""
    lblHsBaixa = ""
    lblReceb = ""
    lblUsu_bx = ""
    lblObsEntr = ""
    lblDtUsuPre = ""
    lblDtUsu = ""
    lblTabPrazo = ""
    lblMetaPrazo = ""
    lblPrazoDias = ""
    cmdRastrPrazo.Enabled = False
    
'registra as variáveis da tela com os dados buscados no recorset
    lblData = de_informa.rsSel_Ctc_SAC.Fields("data")
    lblHora = de_informa.rsSel_Ctc_SAC.Fields("hora")
    lblPrevEntr = de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") & ""
    lblRemet_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), "@@.@@@.@@@/@@@@-@@")
    lblRemet_Nome = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
    
    If IsNull(de_informa.rsSel_Ctc_SAC.Fields("remet_end")) Then
        lblEndRem = de_informa.rsSel_ConsCadCli.Fields("endereco")
    Else
        If Trim$(de_informa.rsSel_Ctc_SAC.Fields("remet_end")) = "" Then
            lblEndRem = de_informa.rsSel_ConsCadCli.Fields("endereco")
        Else
            lblEndRem = de_informa.rsSel_Ctc_SAC.Fields("remet_end")
        End If
    End If
    
    If IsNull(de_informa.rsSel_Ctc_SAC.Fields("remet_cidade")) Then
        lblCidade_orig = de_informa.rsSel_ConsCadCli.Fields("cidade")
    Else
        If Trim$(de_informa.rsSel_Ctc_SAC.Fields("remet_cidade")) = "" Then
            lblCidade_orig = de_informa.rsSel_ConsCadCli.Fields("cidade")
        Else
            lblCidade_orig = de_informa.rsSel_Ctc_SAC.Fields("remet_cidade")
        End If
    End If
    
    If IsNull(de_informa.rsSel_Ctc_SAC.Fields("remet_uf")) Then
        lblUF_Orig = de_informa.rsSel_ConsCadCli.Fields("uf")
    Else
        If Trim$(de_informa.rsSel_Ctc_SAC.Fields("remet_uf")) = "" Then
            lblUF_Orig = de_informa.rsSel_ConsCadCli.Fields("uf")
        Else
            lblUF_Orig = de_informa.rsSel_Ctc_SAC.Fields("remet_uf")
        End If
    End If
    
    lblDest_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("dest_cgc"), "@@.@@@.@@@/@@@@-@@")
    lblDest_Nome = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
    
    If IsNull(de_informa.rsSel_Ctc_SAC.Fields("dest_end")) Then
        lblEndDest = de_informa.rsSel_ConsCadCliDest.Fields("endereco")
    Else
        If Trim$(de_informa.rsSel_Ctc_SAC.Fields("dest_end")) = "" Then
            lblEndDest = de_informa.rsSel_ConsCadCliDest.Fields("endereco")
        Else
            lblEndDest = de_informa.rsSel_Ctc_SAC.Fields("dest_end")
        End If
    End If
    
    If IsNull(de_informa.rsSel_Ctc_SAC.Fields("dest_cidade")) Then
        lblCidade_Dest = de_informa.rsSel_ConsCadCliDest.Fields("cidade")
    Else
        If Trim$(de_informa.rsSel_Ctc_SAC.Fields("dest_cidade")) = "" Then
            lblCidade_Dest = de_informa.rsSel_ConsCadCliDest.Fields("cidade")
        Else
            lblCidade_Dest = de_informa.rsSel_Ctc_SAC.Fields("dest_cidade")
        End If
    End If
    
    If IsNull(de_informa.rsSel_Ctc_SAC.Fields("dest_uf")) Then
        lblUf_Dest = de_informa.rsSel_ConsCadCliDest.Fields("uf")
    Else
        If Trim$(de_informa.rsSel_Ctc_SAC.Fields("dest_uf")) = "" Then
            lblUf_Dest = de_informa.rsSel_ConsCadCliDest.Fields("uf")
        Else
            lblUf_Dest = de_informa.rsSel_Ctc_SAC.Fields("dest_uf")
        End If
    End If
    
    lblColeta = Trim$(de_informa.rsSel_Ctc_SAC.Fields("cidade_orig"))
    lblColetaUF = de_informa.rsSel_Ctc_SAC.Fields("uf_orig")
    
    lblEntrega = Trim$(de_informa.rsSel_Ctc_SAC.Fields("cidade_dest"))
    lblEntregaUF = de_informa.rsSel_Ctc_SAC.Fields("uf_dest")
    
    lblVia = de_informa.rsSel_Ctc_SAC.Fields("via")
    lblRespons_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("respons_cgc"), "@@.@@@.@@@/@@@@-@@")
    lblValmerc = Format(de_informa.rsSel_Ctc_SAC.Fields("valmerc"), "##,###,##0.00")
    lblPeso = Format(de_informa.rsSel_Ctc_SAC.Fields("peso"), "##,##0.0")
    lblPesotax = Format(de_informa.rsSel_Ctc_SAC.Fields("pesotax"), "##,##0.0")
    lblVolumes = Format(de_informa.rsSel_Ctc_SAC.Fields("volumes"), "##,##0")
    lblEspecie = de_informa.rsSel_Ctc_SAC.Fields("especie")
    lblNatureza = de_informa.rsSel_Ctc_SAC.Fields("natureza")
    lblDimensoes = de_informa.rsSel_Ctc_SAC.Fields("dimensoes")
    If de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "URGÊNCIA" Or _
        de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "PRIORIDADE" Then
        LblPrioridade.ForeColor = &HC0&
    Else
        LblPrioridade.ForeColor = &H80000012
    End If
    LblPrioridade = de_informa.rsSel_Ctc_SAC.Fields("prioridade")
    If de_informa.rsSel_Ctc_SAC.Fields("tipodoc") = "MC" Then
        lblTipodoc = "CTR" & "/" & de_informa.rsSel_Ctc_SAC.Fields("motivodoc")
    Else
        lblTipodoc = de_informa.rsSel_Ctc_SAC.Fields("tipodoc") & "/" & de_informa.rsSel_Ctc_SAC.Fields("motivodoc")
    End If
    
    lblFreteNacional = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("fretepeso")), "#,###,##0.00")
    lblAdValorem = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("fretevalor")), "#,###,##0.00")
    lblTxColeta = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txcoleta")), "#,###,##0.00")
    lblTxEntrega = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txentregared")), "#,###,##0.00")
    lblGris = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("gris")), "#,###,##0.00")
    lbltxUrgencia = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txurgencia")), "#,###,##0.00")
    lblPedagio = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("pedagio")), "#,###,##0.00")
    lblTxOutros = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("txoutros")), "#,###,##0.00")
    lblDescrOutros = de_informa.rsSel_Ctc_SAC.Fields("descrtxoutros")
    lblFreteLiq = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("fretetotal")), "#,###,##0.00")
    lblFreteBr = Format(CDbl(de_informa.rsSel_Ctc_SAC.Fields("fretetotalbruto")), "#,###,##0.00")
    
    lblEmissor = de_informa.rsSel_Ctc_SAC.Fields("emissor")
    lblfpag = de_informa.rsSel_Ctc_SAC.Fields("fpag")
    lblFatura = de_informa.rsSel_Ctc_SAC.Fields("faturanum")
    
    If Len(Trim$(lblFatura)) > 0 Then
        If de_informa.rsSel_Fatura.State = 1 Then de_informa.rsSel_Fatura.Close
        de_informa.Sel_Fatura lblFatura
        If de_informa.rsSel_Fatura.RecordCount > 0 Then
            lblDataFatura = de_informa.rsSel_Fatura.Fields("emissao")
            lblVencFatura = de_informa.rsSel_Fatura.Fields("vencimento")
            lblEmissorFatura = de_informa.rsSel_Fatura.Fields("emissor")
            If IsNull(de_informa.rsSel_Fatura.Fields("pagamento")) Then
                lblPagtoFatura = ""
            Else
                lblPagtoFatura = de_informa.rsSel_Fatura.Fields("pagamento")
            End If
        Else
            lblDataFatura = ""
            lblVencFatura = ""
            lblEmissorFatura = ""
            lblPagtoFatura = ""
        End If
    Else
        lblDataFatura = ""
        lblVencFatura = ""
        lblEmissorFatura = ""
        lblPagtoFatura = ""
    End If
    
    lblModal = de_informa.rsSel_Ctc_SAC.Fields("modal")
    xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("modal")
    lblObs_Emissao = de_informa.rsSel_Ctc_SAC.Fields("obs_emissao")
    lblTranspSub = de_informa.rsSel_Ctc_SAC.Fields("transp_sub")
    lblTranspsubRedesp = de_informa.rsSel_Ctc_SAC.Fields("redesp_nome")
    
    If de_informa.rsSel_Ctc_SAC.Fields("motivodoc") = "DEV" Then
        If de_informa.rsSel_Ctc_SAC.Fields("prev_entregatipo") = "I" Then
            lblLabelMeta = "Meta Inform."
            lblMetaPrazo = diasprazo(de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("prev_entrega"), _
                           de_informa.rsSel_Ctc_SAC.Fields("remet_uf"), de_informa.rsSel_Ctc_SAC.Fields("remet_cidade"), _
                           de_informa.rsSel_Ctc_SAC.Fields("hora"), de_informa.rsSel_Ctc_SAC.Fields("modal"), _
                           de_informa.rsSel_Ctc_SAC.Fields("filial"))
        Else
            lblLabelMeta = "Meta Calc."
            xbuscaprazo = buscaprazo2(lblUF_Orig, lblCidade_orig, "TAB000", Mid$(lblModal, 1, 1))
            xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rsSel_Ctc_SAC.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rsSel_Ctc_SAC.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                   Val(Mid$(de_informa.rsSel_Ctc_SAC.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("remet_uf"), _
                   de_informa.rsSel_Ctc_SAC.Fields("remet_cidade")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
            lblMetaPrazo = xprazo_TT
        End If
    Else
        If de_informa.rsSel_Ctc_SAC.Fields("prev_entregatipo") = "I" Then
            lblLabelMeta = "Meta Inform."
            lblMetaPrazo = diasprazo(de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("prev_entrega"), _
                           de_informa.rsSel_Ctc_SAC.Fields("uf_dest"), de_informa.rsSel_Ctc_SAC.Fields("cidade_dest"), _
                           de_informa.rsSel_Ctc_SAC.Fields("hora"), de_informa.rsSel_Ctc_SAC.Fields("modal"), _
                           de_informa.rsSel_Ctc_SAC.Fields("filial"))
        Else
            lblLabelMeta = "Meta Calc."
            xbuscaprazo = buscaprazo2(lblEntregaUF, lblEntrega, de_informa.rsSel_ConsCadCli.Fields("prazo"), Mid$(lblModal, 1, 1))
            xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rsSel_Ctc_SAC.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rsSel_Ctc_SAC.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                   Val(Mid$(de_informa.rsSel_Ctc_SAC.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("uf_dest"), _
                   de_informa.rsSel_Ctc_SAC.Fields("cidade_dest")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
            lblMetaPrazo = xprazo_TT
        End If
    End If

    lblPrevEntr = de_informa.rsSel_Ctc_SAC.Fields("prev_entrega")
    xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")

'ATUALIZA COM DADOS DE ENTREGA E OCORRÊNCIA E STATUS DISPARA

    'traz status de informação do DISPARA
    If de_informa.rsSel_StatusDispara.State = 1 Then de_informa.rsSel_StatusDispara.Close
    de_informa.Sel_StatusDispara xfilialctc
    
    If de_informa.rsSel_StatusDispara.RecordCount > 0 Then
        TabOcorrencias.TabEnabled(1) = True
    Else
        TabOcorrencias.TabEnabled(1) = False
    End If
    
    Set gridDispara.DataSource = de_informa
    gridDispara.DataMember = "Sel_StatusDispara"
    gridDispara.Refresh
    
    
    'consulta que traz os campos que são dados de ocorrência
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
        de_informa.Sel_ConsOcorr2 xfilialctc, "01"
        Set GridConsOcorr.DataSource = de_informa
        GridConsOcorr.DataMember = "Sel_ConsOcorr2"
        GridConsOcorr.Refresh
        If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr") <> "" Then
                lblObs_Ocorr = de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr")
            Else
                lblObs_Ocorr = ""
            End If
        Else
            lblObs_Ocorr = ""
        End If
        If xtemocorr = "0" Then  'se for recusa, busca registro "00" e verifica se há dados de Protocolo para Arquivo
            de_informa.rsSel_ConsOcorr2.MoveFirst
            Do Until de_informa.rsSel_ConsOcorr2.EOF
                If de_informa.rsSel_ConsOcorr2.Fields("cod_ocorr") = "00" Then
                    If IsNull(de_informa.rsSel_ConsOcorr2.Fields("rel_arq_data")) = False Then
                        lblArquivo = de_informa.rsSel_ConsOcorr2.Fields("rel_arq_data")
                    Else
                        lblArquivo = ""
                    End If
                    If IsNull(de_informa.rsSel_ConsOcorr2.Fields("rel_arq_num")) = False Then
                        lblNumProt = String(6 - Len(Trim$(Str(de_informa.rsSel_ConsOcorr2.Fields("rel_arq_num")))), "0") & Trim$(Str(de_informa.rsSel_ConsOcorr2.Fields("rel_arq_num")))
                    Else
                        lblNumProt = ""
                    End If
                End If
                de_informa.rsSel_ConsOcorr2.MoveNext
            Loop
        Else
            lblNumProt = ""
            lblArquivo = ""
        End If
    
        'consulta que traz os campos = 01 que é dado de entrega (ENTREGA REALIZADA)
        If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
        de_informa.Sel_ConsOcorr xfilialctc, "01"
        If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
        'atualiza os campos referente a dados de entrega
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")) = False Then
                lblDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            Else
                lblDtBaixaPre = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")) = False Then
                lblHsBaixaPre = de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")
            Else
                lblHsBaixaPre = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("recebpre")) = False Then
                lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            Else
                lblRecebPre = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")) = False Then
                lblUsu_bxpre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            Else
                lblUsu_bxpre = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) = False Then
                lblDtBaixa = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
            Else
                lblDtBaixa = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("hsbaixa")) = False Then
                lblHsBaixa = de_informa.rsSel_ConsOcorr.Fields("hsbaixa")
            Else
                lblHsBaixa = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("receb")) = False Then
                lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
            Else
                lblReceb = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_bx")) = False Then
                lblUsu_bx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
            Else
                lblUsu_bx = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")) = False Then
                lblObsEntr = de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")
            Else
                lblObsEntr = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_datapre")) = False Then
                lblDtUsuPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            Else
                lblDtUsuPre = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_databx")) = False Then
                lblDtUsu = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            Else
                lblDtUsu = ""
            End If
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("rel_arq_data")) = False And de_informa.rsSel_ConsOcorr.Fields("rel_arquivo") = "S" Then
                lblArquivo = de_informa.rsSel_ConsOcorr.Fields("rel_arq_data")
                If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("rel_arq_num")) Then
                    lblNumProt = String(6 - Len(Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("rel_arq_num")))), "0") & Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("rel_arq_num")))
                Else
                    lblNumProt = ""
                End If
            Else
                lblArquivo = ""
                lblNumProt = ""
            End If
            cmdMostraAbono.Enabled = False
            lblTabPrazo = de_informa.rsSel_ConsCadCli.Fields("prazo")
            'lblMetaPrazo = Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("prazoentr"))) & "/" & Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("diasuteis")))
            If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("diasuteis")) Then
                lblPrazoDias = Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("diasuteis")))
                lblAbonoDias = Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("abonodias")))
                If Val(lblAbonoDias) > 0 Then
                    cmdMostraAbono.Enabled = True
                End If
            Else
                lblPrazoDias = ""
                lblAbonoDias = ""
            End If
            cmdRastrPrazo.Enabled = True
        End If

'dados de status do ctc
    
    cmdRastPrevEntr.Enabled = True
    lblEntregueSN.ToolTipText = ""
    If xtemocorr = "0" Then
        lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
        lblEntregueSN.Caption = "OCORR/Baixado"
    ElseIf xtemocorr = "1" Then
        lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
        If IsNumeric(lblPrazoDias) Then
            If Val(lblPrazoDias) = Val(lblMetaPrazo) Then
                lblEntregueSN.Caption = "ENTREGUE - No Prazo"
            ElseIf Val(lblPrazoDias) > Val(lblMetaPrazo) Then
                If Val(lblAbonoDias) > 0 Then
                    lblEntregueSN.Caption = "ENTREGUE-Atraso: " & Trim$(Str(Val(lblPrazoDias) - Val(lblMetaPrazo))) & " dia(s) (Abono)"
                Else
                    lblEntregueSN.Caption = "ENTREGUE - Atraso: " & Trim$(Str(Val(lblPrazoDias) - Val(lblMetaPrazo))) & " dia(s)"
                End If
            ElseIf Val(lblPrazoDias) < Val(lblMetaPrazo) Then
                lblEntregueSN.Caption = "ENTREGUE - Antecipado: " & Trim$(Str(Val(lblMetaPrazo) - Val(lblPrazoDias))) & " dia(s)"
            End If
        Else
            lblEntregueSN.Caption = "OK. ENTREGUE"
        End If
    ElseIf xtemocorr = "2" Then
        lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
        lblEntregueSN.Caption = "OCORR/Pendente"
    ElseIf xtemocorr = "N" Then
        If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= datahora("data") Then
            lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
            lblEntregueSN.Caption = "EM TRÂNSITO"
            lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
        Else
            lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
            lblEntregueSN.Caption = "SEM POSIÇÃO há " & Trim$(Str(Val(datahora("data") - CDate(Mid$(CVar(de_informa.rsSel_Ctc_SAC.Fields("prev_entrega")), 1, 10))))) & " dia(s)"
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

    DoEvents

'responsável de frete
    
    If Len(Trim$(de_informa.rsSel_Ctc_SAC.Fields("respons_nome"))) < 5 Then
        If de_informa.rsSel_BuscaConsig.State = 1 Then de_informa.rsSel_BuscaConsig.Close
        de_informa.Sel_BuscaConsig de_informa.rsSel_Ctc_SAC.Fields("respons_cgc")
        If de_informa.rsSel_BuscaConsig.RecordCount > 0 Then
            lblRespons_Nome = de_informa.rsSel_BuscaConsig.Fields("nome")
        Else
            lblRespons_Nome = ""
        End If
    Else
        lblRespons_Nome = de_informa.rsSel_Ctc_SAC.Fields("respons_nome")
    End If
    
    
'Notas Fiscais do CTC
    
    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
    de_informa.Sel_NFsdoCTC xfilialctc
    fraNfs.Caption = "Notas Fiscas do CTC: " & de_informa.rsSel_NFsdoCTC.RecordCount & " NF(s)"
    
    gridNfsdoCTC.DataMember = "sel_nfsdoctc"
    gridNfsdoCTC.Refresh
    
'    gridNFsEspeciais.DataMember = "sel_nfsdoctc"
'    gridNFsEspeciais.Refresh
    
    
'    If de_informa.rsSel_NFsdoCTC.RecordCount > 0 Then
'        lblNfs = ""
'        For xflexnf = 1 To de_informa.rsSel_NFsdoCTC.RecordCount
'            lblNfs = lblNfs & Trim$(de_informa.rsSel_NFsdoCTC.Fields("numnf")) & " - "
'            de_informa.rsSel_NFsdoCTC.MoveNext
'        Next
'        If Mid$(lblNfs, Len(lblNfs) - 2, 3) = " - " Then lblNfs = Mid$(lblNfs, 1, Len(lblNfs) - 3)
'    End If
    
'ATUALIZA GRID DE ÚLTIMAS CONSULTAS DE USUÁRIOS
    If de_informa.rsSel_Ultconssac.State = 1 Then de_informa.rsSel_Ultconssac.Close
    de_informa.Sel_Ultconssac xfilialctc
    gridUltCons.DataMember = "sel_ultconssac"
    gridUltCons.Refresh
    
'ATUALIZA GRID DE MANIFESTO
    If de_informa.rsSel_ManifestoPorCTC.State = 1 Then de_informa.rsSel_ManifestoPorCTC.Close
    de_informa.Sel_ManifestoPorCTC xfilialctc
    gridManifesto.DataMember = "sel_manifestoporctc"
    gridManifesto.Refresh
    If de_informa.rsSel_ManifestoPorCTC.RecordCount > 0 Then
        gridManifesto.Enabled = True
    Else
        gridManifesto.Enabled = False
    End If
    
'ZERA GRID DOS CTCS DO MANIFESTO E LABEL DE ESTATÍSTICAS
    gridCTCManifesto.DataMember = ""
    gridCTCManifesto.Refresh
    
    lblManifEstatCTC = ""
    lblManifEstatNF = ""
    lblManifEstatVOL = ""
    lblManifEstatPESO = ""
    lblManifEstatVALMERC = ""
    lblManifEstatFRETE = ""
    lblManifEstatIndNfCtc = ""
    lblManifEstatIndFretePeso = ""
    lblManifEstatIndFreteNf = ""
    
    DoEvents
    
    TxtFilial.SetFocus
    'atualiza usuário e hora que utilisou

    de_informa.ins_ultconssac xfilialctc, xusuario, datahora("datahora")
        
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "CONSULTA", xusuario, "INFORMAÇÃO SAC - CONSULTA CTC: " & xfilialctc
        
    'verifica se existe imagem scanneada do CTC/CTR
    
    If chkAutoScan.Value = 1 Then
    
        If de_informa.rsSel_imagem.State = 1 Then de_informa.rsSel_imagem.Close
        de_informa.Sel_Imagem xfilialctc
        Dim Img() As Byte, I As Long
        
        If de_informa.rsSel_imagem.RecordCount < 1 Then
            If de_informa.rsSel_imagem.RecordCount = 0 And Mid(Trim(txtCtc.Text), 1, 1) = "1" Then
                Dim xTextoAuxiliar As String
                xTextoAuxiliar = txtCtc.Text
                xTextoAuxiliar = Mid(xTextoAuxiliar, 3, 2) & Mid(xTextoAuxiliar, 1, 2) & Mid(xTextoAuxiliar, 5)
                If de_informa.rsSel_imagem.State = 1 Then de_informa.rsSel_imagem.Close
                de_informa.Sel_Imagem transctc(TxtFilial, xTextoAuxiliar)
                If de_informa.rsSel_imagem.RecordCount < 1 Then
                    'NAO ENCONTRA
                    Image1.Picture = LoadPicture(App.Path & "\semscan.jpg")
                    Image1.Visible = False
                    fraPODScan.Visible = True
                Else
                    'ENCONTROU / formato rede local intec
                    Image1.Visible = True
                    fraPODScan.Visible = False
                    Img = de_informa.rsSel_imagem.Fields("imagem")
                    Open "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG" For Binary As #2
                    Put #2, , Img
                    Close #2
                    Image1.Picture = LoadPicture("C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG")
                    Kill "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG"
                End If
            Else
                'NAO ENCONTRA
                Image1.Picture = LoadPicture(App.Path & "\semscan.jpg")
                Image1.Visible = False
                fraPODScan.Visible = True
                End If
        Else
        'ENCONTROU
            Image1.Visible = True
            fraPODScan.Visible = False
            Img = de_informa.rsSel_imagem.Fields("imagem")
            Open "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG" For Binary As #2
            Put #2, , Img
            Close #2
            Image1.Picture = LoadPicture("C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG")
            Kill "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG"
        End If
    End If
    
'Procura Informação de AWB - Atualizado em 15/12/2003

    If De_Aereo.rsSelAWB_CTC.State = 1 Then De_Aereo.rsSelAWB_CTC.Close
    De_Aereo.SelAWB_CTC xfilialctc
    
        If De_Aereo.rsSelAWB_CTC.RecordCount > 0 Then
        CmdAWB.Enabled = True
        Else
        CmdAWB.Enabled = False
        End If
    
    cmbProcurar.Caption = "Procurar"
    cmbProcurar.Enabled = True
    DoEvents
    
    
End Sub
Private Sub cmbSair_Click()
    Unload Me
End Sub

Private Sub CmdAWB_Click()
frmAWB.Show 1
End Sub

Private Sub cmdFinalizaSolic_Click()
End Sub

Private Sub cmdGravaCancela_Click()
End Sub

Private Sub cmdImagem_Click()
    cmdImagem.Caption = "A G U A R D E  ...."
    cmdImagem.Enabled = False
    DoEvents
    
    If de_informa.rsSel_imagem.State = 1 Then de_informa.rsSel_imagem.Close
    de_informa.Sel_Imagem transctc(TxtFilial, txtCtc)

    If de_informa.rsSel_imagem.RecordCount < 1 Then
        If de_informa.rsSel_imagem.RecordCount = 0 And Mid(Trim(txtCtc.Text), 1, 1) = "1" Then
            Dim xTextoAuxiliar As String
            xTextoAuxiliar = txtCtc.Text
            xTextoAuxiliar = Mid(xTextoAuxiliar, 3, 2) & Mid(xTextoAuxiliar, 1, 2) & Mid(xTextoAuxiliar, 5)
            If de_informa.rsSel_imagem.State = 1 Then de_informa.rsSel_imagem.Close
            de_informa.Sel_Imagem transctc(TxtFilial, xTextoAuxiliar)
            If de_informa.rsSel_imagem.RecordCount < 1 Then
                MsgBox "Não Encontrado Imagem Para Este CTC !", vbInformation, "Scanner"
                cmdImagem.Caption = "POD SCANNER ..."
                cmdImagem.Enabled = True
                Exit Sub
            Else
                cmdImagem.Caption = "POD SCANNER ..."
                cmdImagem.Enabled = True
                'ENCONTROU / formato rede local intec
                Call Image1_DblClick
            End If
        Else
            cmdImagem.Caption = "POD SCANNER ..."
            cmdImagem.Enabled = True
            MsgBox "Não Encontrado Imagem Para Este CTC !"
            Exit Sub
        End If
    Else
        cmdImagem.Caption = "POD SCANNER ..."
        cmdImagem.Enabled = True
        'ENCONTROU / formato rede local intec
        Call Image1_DblClick
    End If
End Sub

Private Sub cmdImprTela_Click()
    Printer.KillDoc
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdMostraAbono_Click()
    frmAbonoDetalhe.lblEmissao = lblData
    frmAbonoDetalhe.lblDestino = Trim$(lblCidade_Dest) & " - " & lblUf_Dest
    frmAbonoDetalhe.lblMeta = lblMetaPrazo
    frmAbonoDetalhe.lblPrevisao = lblPrevEntr
    frmAbonoDetalhe.lblEntrega = lblDtBaixaPre
    frmAbonoDetalhe.lblReal = lblPrazoDias
    frmAbonoDetalhe.lblAbono = lblAbonoDias
    frmAbonoDetalhe.lblAbonador = de_informa.rsSel_ConsOcorr.Fields("usu_abono")
    frmAbonoDetalhe.lblDtAbono = de_informa.rsSel_ConsOcorr.Fields("data_abono")
    frmAbonoDetalhe.lblObsAbono = de_informa.rsSel_ConsOcorr.Fields("obs_abono")
    frmAbonoDetalhe.Show 1
End Sub

Private Sub cmdNovaOcorrInf_Click()
End Sub

Private Sub cmdNovasolic_Click()
End Sub

Private Sub cmdPesq_Click()
    optCTC.Value = True
    optCTC_Click
    DoEvents
    frmPesquisaCTC.Show 1
End Sub
Private Sub cmdRastPrevEntr_Click()
    frmRastrPrazo.lblFilialctc = transctc(frmSac.TxtFilial, frmSac.txtCtc)
    frmRastrPrazo.lblModal = frmSac.lblModal
    frmRastrPrazo.lblCidadeDest = frmSac.lblCidade_Dest
    frmRastrPrazo.lblUFdest = frmSac.lblUf_Dest
    frmRastrPrazo.lblEmissao = frmSac.lblData
    frmRastrPrazo.lblHsEmiss = frmSac.lblHora
    frmRastrPrazo.lblEntrega = lblPrevEntr.Caption
    frmRastrPrazo.lblPrazo = lblMetaPrazo.Caption
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
    'de_informa.Alt_AtualPrazoSCTC transctc(frmSac.txtfilial, frmSac.txtCTC)
    frmAtualPrazos.lblFilialctc = transctc(frmSac.TxtFilial, frmSac.txtCtc)
    frmAtualPrazos.Show 1
    If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
    de_informa.Sel_ConsOcorr transctc(frmSac.TxtFilial, frmSac.txtCtc), "01"
    frmRastrPrazo.lblFilialctc = transctc(frmSac.TxtFilial, frmSac.txtCtc)
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
    cmbProcurar_Click
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Activate()
'    If optCTC.Value = True Then
'        If xultimofilial <> "" Then
'           txtfilial.Text = xultimofilial
'           txtCTC.Text = xultimoctc
'        End If
'        txtfilial.SetFocus
'    ElseIf optNf.Value = True Then
'        txtNumNf.SetFocus
'    End If
'
End Sub
Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    GridConsOcorr.DataMember = ""
    gridUltCons.DataMember = ""
    gridCTCManifesto.DataMember = ""
    gridManifesto.DataMember = ""
    cmbProcurar.Enabled = True
    cmbSair.Enabled = True
    If xultimofilial <> "" Then
        TxtFilial.Text = xultimofilial
        txtCtc.Text = xultimoctc
    End If
    
    TabOcorrencias.TabEnabled(1) = False
    
'    optCTC_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Me.Caption <> "SAC - Informação de Transporte - Acompanhamento (chamada)" Then
        mdiInforma.Toolbar1.Visible = True
        mdiInforma.StatusBar1.Visible = True
    End If
    Set frmSac = Nothing
'    If de_informa.rsSel_NFBasecli.State = 1 Then de_informa.rsSel_NFBasecli.Close
'    gridClienteEspeciais.DataMember = "Sel_NFBasecli"
'    gridClienteEspeciais.Refresh
    
'    If de_informa.rsSel_NFVideolar.State = 1 Then de_informa.rsSel_NFVideolar.Close
'    gridVideolarEspeciais.DataMember = "Sel_NFVideolar"
'    gridVideolarEspeciais.Refresh
    
'    If de_informa.rsSel_CheckReceb.State = 1 Then de_informa.rsSel_CheckReceb.Close
'    gridCheckEspeciais.DataMember = "Sel_CheckReceb"
'    gridCheckEspeciais.Refresh
    
End Sub

Private Sub gridClienteEspeciais_Click()

End Sub

Private Sub GridConsOcorr_Click()
    If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
    'atualiza o campo de obs de ocorrência quando clicado no grid
        lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
    End If
End Sub

Private Sub gridManifesto_Click()

Dim xvolumes As Long, xpeso As Currency, xvalmerc As Currency, xfretetotal As Currency

Me.MousePointer = 11

'ATUALIZA RECORDSET QUE IRÁ ATUALIZAR O GRID DE CTCS DO MANIFESTO
    If de_informa.rsSel_CTCsdoManifesto.State = 1 Then de_informa.rsSel_CTCsdoManifesto.Close
    de_informa.Sel_CTCsdoManifesto gridManifesto.Columns(2)
    
'ATUALIZA DADOS ESTATÍSTICOS DE MANIFESTO
    
    
    If de_informa.rsSel_CTCsdoManifesto.RecordCount > 0 Then
        de_informa.rsSel_CTCsdoManifesto.MoveFirst
    End If
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
    
    If de_informa.rsSel_CTCsdoManifesto.RecordCount > 0 Then
        lblManifEstatCTC = de_informa.rsSel_CTCsdoManifesto.RecordCount
        lblManifEstatNF = de_informa.rsSel_NfsdoManifesto.Fields("qtd")
        lblManifEstatVOL = xvolumes
        lblManifEstatPESO = Format(xpeso, "##,##0.0")
        lblManifEstatVALMERC = Format(xvalmerc, "##,###,##0.00")
        lblManifEstatFRETE = Format(xfretetotal, "##,###,##0.00")
        lblManifEstatIndNfCtc = Format(Val(lblManifEstatNF) / Val(lblManifEstatCTC), "#,##0.0")
        lblManifEstatIndFretePeso = Format(xfretetotal / xpeso, "#,##0.00")
        lblManifEstatIndFreteNf = Format(xfretetotal / xvalmerc, "##0.00%")
    End If
    
    
    'ATUALIZA O GRID DE MANIFESTO
    
    gridCTCManifesto.DataMember = "Sel_CTCsdoManifesto"
    gridCTCManifesto.Refresh

    
Me.MousePointer = 0
End Sub

Private Sub gridNFsEspeciais_Click()

End Sub

Private Sub Image1_DblClick()
    Dim Img() As Byte, I As Long
    frmImagem.lblTotImg = Trim$(Val(de_informa.rsSel_imagem.RecordCount))
    frmImagem.lblNumImg = 1
    
    If de_informa.rsSel_imagem.RecordCount > 1 Then
        frmImagem.cmdAnterior.Enabled = False
        frmImagem.cmdProximo.Enabled = True
    Else
        frmImagem.cmdAnterior.Enabled = False
        frmImagem.cmdProximo.Enabled = False
    End If
    
    Img = de_informa.rsSel_imagem.Fields("imagem")
    frmImagem.lblDataHora = de_informa.rsSel_imagem.Fields("data")
    Open "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG" For Binary As #2
    Put #2, , Img
    Close #2
    On Error Resume Next
    frmImagem.Image1.Picture = LoadPicture("C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG")
    frmImagem.lblCtc = transctc(TxtFilial, txtCtc)
    Kill "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG"
    frmImagem.Show 1
End Sub

Private Sub Label69_Click()

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
Private Sub TxtFilial_Change()
    On Error Resume Next
    If Len(TxtFilial.Text) >= 2 Then txtCtc.SetFocus
End Sub
Private Sub TxtFilial_GotFocus()
    TxtFilial.SelStart = 0
    TxtFilial.SelLength = 2
End Sub
Private Sub optCTC_Click()
    On Error Resume Next
    TxtFilial.Visible = True
    txtCtc.Visible = True
    txtNumNf.Visible = False
    TxtFilial.SetFocus
End Sub
Private Sub optNf_Click()
    On Error Resume Next
    TxtFilial.Visible = False
    txtCtc.Visible = False
    txtNumNf.Visible = True
    txtNumNf.SetFocus
End Sub
Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        If TxtFilial.Text = "" Then
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
    If TxtFilial.Text <> "" Then
        If Not IsNumeric(TxtFilial.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            TxtFilial.SetFocus
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


