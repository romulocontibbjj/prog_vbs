VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmGerencial 
   Caption         =   "Informações Gerenciais"
   ClientHeight    =   8070
   ClientLeft      =   1875
   ClientTop       =   1590
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   13005
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "CTCs Fiscais Por Cliente"
      TabPicture(0)   =   "frmGerencial.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraMensagem0"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "CTCs Fiscais"
      TabPicture(1)   =   "frmGerencial.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraMensagem"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Notas Fiscais de Serviço"
      TabPicture(2)   =   "frmGerencial.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame8"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraMensagem1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Movimento BOMI"
      TabPicture(3)   =   "frmGerencial.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Resumo"
      TabPicture(4)   =   "frmGerencial.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Frame Frame9 
         Caption         =   "Totais"
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
         TabIndex        =   109
         Top             =   5760
         Width           =   11415
         Begin VB.Label Label19 
            Caption         =   $"frmGerencial.frx":008C
            Height          =   1335
            Left            =   1320
            TabIndex        =   110
            Top             =   120
            Width           =   5775
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Emissões de CTRs - Negociação BOMI FARMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   107
         Top             =   480
         Width           =   11415
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmGerencial.frx":01C5
            Height          =   4815
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   8493
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
            DataMember      =   "Sel_GerNfsEmitidos"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "Filial Origem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde de NFS"
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
            BeginProperty Column02 
               DataField       =   "valornf"
               Caption         =   "Valor da NFS"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2174,74
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1454,74
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   2294,929
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fraMensagem1 
         Height          =   1455
         Left            =   -71400
         TabIndex        =   49
         Top             =   2040
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "AGUARDE ...."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Left            =   1200
            TabIndex        =   50
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame fraMensagem0 
         Height          =   1455
         Left            =   3720
         TabIndex        =   75
         Top             =   2040
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "AGUARDE ...."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Left            =   1200
            TabIndex        =   76
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Total Geral Emitido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   97
         Top             =   480
         Width           =   11415
         Begin VB.CommandButton Command12 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   99
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdImprTela1Cliente 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   98
            Top             =   1440
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid gridEmitidoCliente 
            Bindings        =   "frmGerencial.frx":01DE
            Height          =   1575
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsEmitidosCliente"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "respons_nome"
               Caption         =   "Cliente"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde CTCs"
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
            BeginProperty Column02 
               DataField       =   "fretefinal"
               Caption         =   "Valor Frete"
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
            BeginProperty Column03 
               DataField       =   "valmerc"
               Caption         =   "Valor Mercadoria"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2789,858
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1365,165
               EndProperty
            EndProperty
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   106
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Total Frete:"
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
            Left            =   7440
            TabIndex        =   105
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   104
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblTotCtcEmitCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   103
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblTotValMercEmitCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   102
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotFreteEmitCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   8880
            TabIndex        =   101
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "CTCs Faturados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   87
         Top             =   2520
         Width           =   11415
         Begin VB.CommandButton Command10 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   89
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdImprTela2Cliente 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   88
            Top             =   1440
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid gridFaturadoCliente 
            Bindings        =   "frmGerencial.frx":01F7
            Height          =   1575
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsFaturadosCliente"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "respons_nome"
               Caption         =   "Cliente"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde CTCs"
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
            BeginProperty Column02 
               DataField       =   "fretefinal"
               Caption         =   "Valor Frete"
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
            BeginProperty Column03 
               DataField       =   "valmerc"
               Caption         =   "Valor Mercadoria"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2805,166
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1065,26
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1379,906
               EndProperty
            EndProperty
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   96
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Total Frete:"
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
            Left            =   7440
            TabIndex        =   95
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   94
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblTotCtcFatCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   93
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblTotValMercFatCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   92
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotFreteFatCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   8880
            TabIndex        =   91
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "CTCs Não Faturados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   77
         Top             =   4680
         Width           =   11415
         Begin VB.CommandButton Command6 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   80
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdImprTela3Cliente 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   79
            Top             =   1440
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid gridNaoFatCliente 
            Bindings        =   "frmGerencial.frx":0210
            Height          =   1575
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsNaoFatCliente"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "respons_nome"
               Caption         =   "Cliente"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "qtde"
               Caption         =   "Qtde CTCs"
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
            BeginProperty Column02 
               DataField       =   "fretefinal"
               Caption         =   "Valor Frete"
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
            BeginProperty Column03 
               DataField       =   "valmerc"
               Caption         =   "Valor Mercadoria"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2849,953
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1379,906
               EndProperty
            EndProperty
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   86
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Total Frete:"
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
            Left            =   7440
            TabIndex        =   85
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   84
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblTotCtcNaoFatCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   83
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblTotValMercNaoFatCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   82
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotFreteNaoFatCliente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   8880
            TabIndex        =   81
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Total de Notas de Serviço Emitidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   67
         Top             =   480
         Width           =   11415
         Begin VB.CommandButton Command9 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   69
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdImprTela4 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   68
            Top             =   1440
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid gridEmitidoNF 
            Bindings        =   "frmGerencial.frx":0229
            Height          =   1575
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerNfsEmitidos"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "Filial Origem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde de NFS"
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
            BeginProperty Column02 
               DataField       =   "valornf"
               Caption         =   "Valor da NFS"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2174,74
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1454,74
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   2294,929
               EndProperty
            EndProperty
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Total de NFs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   74
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total NFs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   73
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label lblTotNfsEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   72
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblTotValorNFSEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   71
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Notas de Serviço Faturadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   59
         Top             =   2520
         Width           =   11415
         Begin VB.CommandButton Command7 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   61
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdImprTela5 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   60
            Top             =   1440
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid gridFaturadoNF 
            Bindings        =   "frmGerencial.frx":0242
            Height          =   1575
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerNfsFaturados"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "Filial Origem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde de NFS"
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
            BeginProperty Column02 
               DataField       =   "valornf"
               Caption         =   "Valor da NFS"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2174,74
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1425,26
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   2340,284
               EndProperty
            EndProperty
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Total de NFs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   66
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total NFs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   65
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label lblTotNfsFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   64
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblTotValorNFSFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   63
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Notas de Serviço Não Faturadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   51
         Top             =   4680
         Width           =   11415
         Begin VB.CommandButton Command5 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   54
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdImprTela6 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   53
            Top             =   1440
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid gridNaoFatNF 
            Bindings        =   "frmGerencial.frx":025B
            Height          =   1575
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerNfsNaoFat"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "Filial Origem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "qtde"
               Caption         =   "Qtde de NFS"
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
            BeginProperty Column02 
               DataField       =   "valornf"
               Caption         =   "Valor da NFS"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2174,74
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1454,74
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   2324,977
               EndProperty
            EndProperty
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Total de NFs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   58
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total NFs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   57
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label lblTotNfsNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   56
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblTotValorNFSNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   55
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame fraMensagem 
         Height          =   1455
         Left            =   -71400
         TabIndex        =   47
         Top             =   2040
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "AGUARDE ...."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Left            =   1200
            TabIndex        =   48
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "CTCs Não Faturados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   20
         Top             =   4680
         Width           =   11415
         Begin MSDataGridLib.DataGrid gridNaoFat 
            Bindings        =   "frmGerencial.frx":0274
            Height          =   1575
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsNaoFat"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "Filial Origem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde CTCs"
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
            BeginProperty Column02 
               DataField       =   "fretefinal"
               Caption         =   "Valor Frete"
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
            BeginProperty Column03 
               DataField       =   "valmerc"
               Caption         =   "Valor Mercadoria"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1349,858
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1995,024
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdImprTela3 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   13
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   14
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label lblTotFreteNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   40
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblTotValMercNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   39
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotCtcNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   38
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   37
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Total Frete:"
            Height          =   195
            Left            =   7440
            TabIndex        =   36
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   35
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CTCs Faturados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   19
         Top             =   2520
         Width           =   11415
         Begin VB.CommandButton cmdImprTela2 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   12
            Top             =   1440
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid gridFaturado 
            Bindings        =   "frmGerencial.frx":028D
            Height          =   1575
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsFaturados"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "Filial Origem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde CTCs"
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
            BeginProperty Column02 
               DataField       =   "fretefinal"
               Caption         =   "Valor Frete"
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
            BeginProperty Column03 
               DataField       =   "valmerc"
               Caption         =   "Valor Mercadoria"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1349,858
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1980,284
               EndProperty
            EndProperty
         End
         Begin VB.Label lblTotFreteFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   34
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblTotValMercFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   33
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotCtcFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   32
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Total Frete:"
            Height          =   195
            Left            =   7440
            TabIndex        =   30
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   29
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total de CTCs Emitidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   11415
         Begin VB.CommandButton cmdImprTela1 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   9
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Imprimir/Arquivo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   10
            Top             =   1440
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid gridEmitido 
            Bindings        =   "frmGerencial.frx":02A6
            Height          =   1575
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsEmitidos"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "Filial Origem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "qtde"
               Caption         =   "Qtde CTCs"
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
            BeginProperty Column02 
               DataField       =   "fretefinal"
               Caption         =   "Valor Frete"
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
            BeginProperty Column03 
               DataField       =   "valmerc"
               Caption         =   "Valor Mercadoria"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1365,165
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1950,236
               EndProperty
            EndProperty
         End
         Begin VB.Label lblTotFreteEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   28
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblTotValMercEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   27
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotCtcEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   26
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   25
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Total Frete:"
            Height          =   195
            Left            =   7440
            TabIndex        =   24
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   23
            Top             =   240
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
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
      TabIndex        =   15
      Top             =   0
      Width           =   11655
      Begin VB.OptionButton Option3 
         Caption         =   "Aereo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   680
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Rodo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   680
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rodo+Aereo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   680
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   615
         Left            =   10800
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   615
         Left            =   9720
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   240
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
         Left            =   840
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   45
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Responsável:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6600
         TabIndex        =   44
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   43
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Responsável:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3480
         TabIndex        =   42
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmGerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()

End Sub

Private Sub cmdImprTela1_Click()
    Printer.KillDoc
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdImprTela1Cliente_Click()
Printer.KillDoc

If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
Me.PrintForm
End Sub

Private Sub cmdImprTela2_Click()
    Printer.KillDoc
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdImprTela2Cliente_Click()
Printer.KillDoc

If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
Me.PrintForm
End Sub

Private Sub cmdImprTela3_Click()
    Printer.KillDoc
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdImprTela3Cliente_Click()
Printer.KillDoc

If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
Me.PrintForm
End Sub

Private Sub cmdImprTela4_Click()
Printer.KillDoc

If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
Me.PrintForm


End Sub

Private Sub cmdImprTela5_Click()
Printer.KillDoc

If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
Me.PrintForm
End Sub

Private Sub cmdImprTela6_Click()
Printer.KillDoc

If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
Me.PrintForm
End Sub

Private Sub cmdProcessar_Click()
    Dim xTotCtcEmit As Currency, xTotValMercEmit As Currency, xTotFreteEmit As Currency
    Dim xTotCtcFat As Currency, xTotValMercFat As Currency, xTotFreteFat As Currency
    Dim xTotCtcNaoFat As Currency, xTotValMercNaoFat As Currency, xTotFreteNaoFat As Currency
    
    Dim xTotNFSEmit As Currency, xTotValorNFSEmit As Currency
    Dim xTotNFSFat As Currency, xTotValorNFSFat As Currency
    Dim xTotNFSNaoFat As Currency, xTotValorNFSNaoFat As Currency

    fraMensagem.Visible = True
    fraMensagem0.Visible = True
    fraMensagem1.Visible = True
    cmdProcessar.Enabled = False
    cmdSair.Enabled = False
    SSTab1.Enabled = False
    DoEvents
    
'****************** CTCs FISCAIS ********************
    
'**** POR CLIENTE ****
    
    If de_informa.rsSel_GerCtcsEmitidosCliente.State = 1 Then de_informa.rsSel_GerCtcsEmitidosCliente.Close
    de_informa.Sel_GerCtcsEmitidosCliente mskPer1, mskPer2, "%", "%", "%"
    
    gridEmitidoCliente.DataMember = "Sel_GerCtcsEmitidosCliente"
    gridEmitidoCliente.Refresh
    
    If de_informa.rsSel_GerCtcsEmitidosClienteTotal.State = 1 Then de_informa.rsSel_GerCtcsEmitidosClienteTotal.Close
    de_informa.Sel_GerCtcsEmitidosClienteTotal mskPer1, mskPer2, "%", "%", "%"
    
    lblTotCtcEmitCliente = Format(de_informa.rsSel_GerCtcsEmitidosClienteTotal.Fields("qtde"), "##,###,##0")
    lblTotValMercEmitCliente = Format(de_informa.rsSel_GerCtcsEmitidosClienteTotal.Fields("valmerc"), "##,###,##0.00")
    lblTotFreteEmitCliente = Format(de_informa.rsSel_GerCtcsEmitidosClienteTotal.Fields("fretefinal"), "##,###,##0.00")
    
    lblTotCtcEmit = Format(de_informa.rsSel_GerCtcsEmitidosClienteTotal.Fields("qtde"), "##,###,##0")
    lblTotValMercEmit = Format(de_informa.rsSel_GerCtcsEmitidosClienteTotal.Fields("valmerc"), "##,###,##0.00")
    lblTotFreteEmit = Format(de_informa.rsSel_GerCtcsEmitidosClienteTotal.Fields("fretefinal"), "##,###,##0.00")
    
    DoEvents
    
    If de_informa.rsSel_GerCtcsFaturadosCliente.State = 1 Then de_informa.rsSel_GerCtcsFaturadosCliente.Close
    de_informa.Sel_GerCtcsFaturadosCliente mskPer1, mskPer2, "%", "%", "%"
    
    gridFaturadoCliente.DataMember = "Sel_GerCtcsFaturadosCliente"
    gridFaturadoCliente.Refresh
    
    If de_informa.rsSel_GerCtcsFaturadosClienteTotal.State = 1 Then de_informa.rsSel_GerCtcsFaturadosClienteTotal.Close
    de_informa.Sel_GerCtcsFaturadosClienteTotal mskPer1, mskPer2, "%", "%", "%"
    
    lblTotCtcFatCliente = Format(de_informa.rsSel_GerCtcsFaturadosClienteTotal.Fields("qtde"), "##,###,##0")
    lblTotValMercFatCliente = Format(de_informa.rsSel_GerCtcsFaturadosClienteTotal.Fields("valmerc"), "##,###,##0.00")
    lblTotFreteFatCliente = Format(de_informa.rsSel_GerCtcsFaturadosClienteTotal.Fields("fretefinal"), "##,###,##0.00")
    
    lblTotCtcFat = Format(de_informa.rsSel_GerCtcsFaturadosClienteTotal.Fields("qtde"), "##,###,##0")
    lblTotValMercFat = Format(de_informa.rsSel_GerCtcsFaturadosClienteTotal.Fields("valmerc"), "##,###,##0.00")
    lblTotFreteFat = Format(de_informa.rsSel_GerCtcsFaturadosClienteTotal.Fields("fretefinal"), "##,###,##0.00")
    
    DoEvents
    
    If de_informa.rsSel_GerCtcsNaoFatCliente.State = 1 Then de_informa.rsSel_GerCtcsNaoFatCliente.Close
    de_informa.Sel_GerCtcsNaoFatCliente mskPer1, mskPer2, "%", "%", "%"
    
    gridNaoFatCliente.DataMember = "Sel_GerCtcsNaoFatCliente"
    gridNaoFatCliente.Refresh
    
    If de_informa.rsSel_GerCtcsNaoFatClienteTotal.State = 1 Then de_informa.rsSel_GerCtcsNaoFatClienteTotal.Close
    de_informa.Sel_GerCtcsNaoFatClienteTotal mskPer1, mskPer2, "%", "%", "%"
    
    lblTotCtcNaoFatCliente = Format(de_informa.rsSel_GerCtcsNaoFatClienteTotal.Fields("qtde"), "##,###,##0")
    lblTotValMercNaoFatCliente = Format(de_informa.rsSel_GerCtcsNaoFatClienteTotal.Fields("valmerc"), "##,###,##0.00")
    lblTotFreteNaoFatCliente = Format(de_informa.rsSel_GerCtcsNaoFatClienteTotal.Fields("fretefinal"), "##,###,##0.00")
    
    lblTotCtcNaoFat = Format(de_informa.rsSel_GerCtcsNaoFatClienteTotal.Fields("qtde"), "##,###,##0")
    lblTotValMercNaoFat = Format(de_informa.rsSel_GerCtcsNaoFatClienteTotal.Fields("valmerc"), "##,###,##0.00")
    lblTotFreteNaoFat = Format(de_informa.rsSel_GerCtcsNaoFatClienteTotal.Fields("fretefinal"), "##,###,##0.00")
    
    DoEvents
    
'**** POR ORIGEM ****
    
    If de_informa.rsSel_GerCtcsEmitidos.State = 1 Then de_informa.rsSel_GerCtcsEmitidos.Close
    de_informa.Sel_GerCtcsEmitidos mskPer1, mskPer2, "%", "%", "%"
    
    gridEmitido.DataMember = "Sel_GerCtcsEmitidos"
    gridEmitido.Refresh
    DoEvents
    
    If de_informa.rsSel_GerCtcsFaturados.State = 1 Then de_informa.rsSel_GerCtcsFaturados.Close
    de_informa.Sel_GerCtcsFaturados mskPer1, mskPer2, "%", "%", "%"
    
    gridFaturado.DataMember = "Sel_GerCtcsFaturados"
    gridFaturado.Refresh
    DoEvents
    
    If de_informa.rsSel_GerCtcsNaoFat.State = 1 Then de_informa.rsSel_GerCtcsNaoFat.Close
    de_informa.Sel_GerCtcsNaoFat mskPer1, mskPer2, "%", "%", "%"
    
    gridNaoFat.DataMember = "Sel_GerCtcsNaoFat"
    gridNaoFat.Refresh
    DoEvents
    
'****************** NOTAS FISCAIS DE SERVIÇO ********************
    
    If de_informa.rsSel_GerNfsEmitidos.State = 1 Then de_informa.rsSel_GerNfsEmitidos.Close
    de_informa.Sel_GerNfsEmitidos mskPer1, mskPer2, "%"
    
    gridEmitidoNF.DataMember = "Sel_GerNfsEmitidos"
    gridEmitidoNF.Refresh
    DoEvents
    
    If de_informa.rsSel_GerNfsFaturados.State = 1 Then de_informa.rsSel_GerNfsFaturados.Close
    de_informa.Sel_GerNfsFaturados mskPer1, mskPer2, "%"
    
    gridFaturadoNF.DataMember = "Sel_GerNfsFaturados"
    gridFaturadoNF.Refresh
    DoEvents
    
    If de_informa.rsSel_GerNfsNaoFat.State = 1 Then de_informa.rsSel_GerNfsNaoFat.Close
    de_informa.Sel_GerNfsNaoFat mskPer1, mskPer2, "%"
    
    gridNaoFatNF.DataMember = "Sel_GerNfsNaoFat"
    gridNaoFatNF.Refresh
    DoEvents
    
    xTotNFSEmit = 0
    xTotValorNFSEmit = 0
    xTotNFSFat = 0
    xTotValorNFSFat = 0
    xTotNFSNaoFat = 0
    xTotValorNFSNaoFat = 0
    
    de_informa.rsSel_GerNfsEmitidos.MoveFirst
    
    Do Until de_informa.rsSel_GerNfsEmitidos.EOF
        xTotNFSEmit = xTotNFSEmit + de_informa.rsSel_GerNfsEmitidos.Fields("qtde")
        xTotValorNFSEmit = xTotValorNFSEmit + de_informa.rsSel_GerNfsEmitidos.Fields("valornf")
        de_informa.rsSel_GerNfsEmitidos.MoveNext
    Loop
    
    lblTotNfsEmit = Format(xTotNFSEmit, "##,###,##0")
    lblTotValorNFSEmit = Format(xTotValorNFSEmit, "##,###,##0.00")
    
    de_informa.rsSel_GerNfsFaturados.MoveFirst
    
    Do Until de_informa.rsSel_GerNfsFaturados.EOF
        xTotNFSFat = xTotNFSFat + de_informa.rsSel_GerNfsFaturados.Fields("qtde")
        xTotValorNFSFat = xTotValorNFSFat + de_informa.rsSel_GerNfsFaturados.Fields("valornf")
        de_informa.rsSel_GerNfsFaturados.MoveNext
    Loop
    
    lblTotNfsFat = Format(xTotNFSFat, "##,###,##0")
    lblTotValorNFSFat = Format(xTotValorNFSFat, "##,###,##0.00")
    
    de_informa.rsSel_GerNfsNaoFat.MoveFirst
    
    Do Until de_informa.rsSel_GerNfsNaoFat.EOF
        xTotNFSNaoFat = xTotNFSNaoFat + de_informa.rsSel_GerNfsNaoFat.Fields("qtde")
        xTotValorNFSNaoFat = xTotValorNFSNaoFat + de_informa.rsSel_GerNfsNaoFat.Fields("valornf")
        de_informa.rsSel_GerNfsNaoFat.MoveNext
    Loop
    
    lblTotNfsNaoFat = Format(xTotNFSNaoFat, "##,###,##0")
    lblTotValorNFSNaoFat = Format(xTotValorNFSNaoFat, "##,###,##0.00")
    
    fraMensagem.Visible = False
    fraMensagem0.Visible = False
    fraMensagem1.Visible = False
    cmdProcessar.Enabled = True
    cmdSair.Enabled = True
    SSTab1.Enabled = True
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mdiFatura.ToolFaturamento.Visible = False
    If de_informa.rsSel_GerCtcsEmitidos.State = 1 Then de_informa.rsSel_GerCtcsEmitidos.Close
    gridEmitido.DataMember = "Sel_GerCtcsEmitidos"
    gridEmitido.Refresh
    DoEvents
    
    If de_informa.rsSel_GerCtcsFaturados.State = 1 Then de_informa.rsSel_GerCtcsFaturados.Close
    gridFaturado.DataMember = "Sel_GerCtcsFaturados"
    gridFaturado.Refresh
    DoEvents
    
    If de_informa.rsSel_GerCtcsNaoFat.State = 1 Then de_informa.rsSel_GerCtcsNaoFat.Close
    gridNaoFat.DataMember = "Sel_GerCtcsNaoFat"
    gridNaoFat.Refresh
    DoEvents


    If de_informa.rsSel_GerNfsEmitidos.State = 1 Then de_informa.rsSel_GerNfsEmitidos.Close
    gridEmitidoNF.DataMember = "Sel_GerNfsEmitidos"
    gridEmitidoNF.Refresh
    DoEvents
    
    If de_informa.rsSel_GerNfsFaturados.State = 1 Then de_informa.rsSel_GerNfsFaturados.Close
    gridFaturadoNF.DataMember = "Sel_GerNfsFaturados"
    gridFaturadoNF.Refresh
    DoEvents
    
    If de_informa.rsSel_GerNfsNaoFat.State = 1 Then de_informa.rsSel_GerNfsNaoFat.Close
    gridNaoFatNF.DataMember = "Sel_GerNfsNaoFat"
    gridNaoFatNF.Refresh
    DoEvents




End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiFatura.ToolFaturamento.Visible = True
End Sub

Private Sub Label25_Click()

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
    End If
End Sub

Private Sub SSTab1_DblClick()

End Sub
