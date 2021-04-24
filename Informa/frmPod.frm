VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPod 
   Caption         =   "Informação de Entregas e Ocorrências"
   ClientHeight    =   8010
   ClientLeft      =   300
   ClientTop       =   1620
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Baixas (Pré e Física) - Ocorrências"
      TabPicture(0)   =   "frmPod.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "xt"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraProcura"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Pré-Baixa por Manifesto (Somente a Data de Entrega)"
      TabPicture(1)   =   "frmPod.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame13"
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(2)=   "Frame11"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   -74880
         TabIndex        =   89
         Top             =   480
         Width           =   2295
         Begin VB.Label lblFilialManifesto 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame12 
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
         Height          =   6135
         Left            =   -74880
         TabIndex        =   78
         Top             =   1320
         Width           =   11535
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
            Height          =   4455
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   7858
            _Version        =   393216
            BackColor       =   -2147483644
            BackColorSel    =   12582912
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.CommandButton cmdGravarPreBxMnf 
            Caption         =   "Gravar Pré-Baixa"
            Height          =   495
            Left            =   9360
            TabIndex        =   55
            Top             =   5520
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   91
            Text            =   "01"
            Top             =   5640
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskDataMnf 
            Height          =   285
            Left            =   1800
            TabIndex        =   54
            Top             =   5640
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   -2147483634
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   5640
            TabIndex        =   95
            Top             =   5640
            Width           =   765
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ENTREGA REALIZADA"
            Height          =   285
            Left            =   6600
            TabIndex        =   94
            Top             =   5640
            Width           =   2415
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cód. de Ocorrência:"
            Height          =   195
            Left            =   3360
            TabIndex        =   93
            Top             =   5640
            Width           =   1425
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Data da Pré-Baixa:"
            Height          =   195
            Left            =   240
            TabIndex        =   92
            Top             =   5640
            Width           =   1335
         End
         Begin VB.Label lblPlacaMnf 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5520
            TabIndex        =   88
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblDataMnf 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   87
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblMotoristaMnf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7200
            TabIndex        =   86
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label lblPropMnf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   85
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblQtdeCtcMnf 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   10920
            TabIndex        =   84
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Qtde.Entregas:"
            Height          =   195
            Left            =   9720
            TabIndex        =   83
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Placa:"
            Height          =   195
            Left            =   5040
            TabIndex        =   82
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário:"
            Height          =   195
            Left            =   2040
            TabIndex        =   81
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Motorista:"
            Height          =   195
            Left            =   6480
            TabIndex        =   80
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Emissao:"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " Filial  e  Manifesto "
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
         Left            =   -72480
         TabIndex        =   77
         Top             =   480
         Width           =   9135
         Begin VB.CommandButton cmdImprTela2 
            Height          =   495
            Left            =   8520
            Picture         =   "frmPod.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   160
            Width           =   495
         End
         Begin VB.CommandButton cmdSair2 
            Caption         =   "Sair"
            Height          =   375
            Left            =   5040
            TabIndex        =   56
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton cmdBuscaCTCs 
            Caption         =   "Buscar CTCs Deste Manifesto"
            Height          =   375
            Left            =   2160
            TabIndex        =   52
            Top             =   240
            Width           =   2595
         End
         Begin VB.TextBox txtManifesto 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   840
            MaxLength       =   6
            TabIndex        =   51
            Top             =   300
            Width           =   1050
         End
         Begin VB.TextBox txtFilialmnf 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   200
            MaxLength       =   2
            TabIndex        =   50
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame fraProcura 
         Caption         =   "Núm. da  Filial e CTC"
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
         Left            =   4335
         TabIndex        =   76
         Top             =   480
         Width           =   7320
         Begin VB.CommandButton cmdImprTela 
            Height          =   495
            Left            =   6720
            Picture         =   "frmPod.frx":07AA
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   160
            Width           =   495
         End
         Begin VB.CommandButton cmbGravar 
            Caption         =   "Gravar a Ocorr."
            Enabled         =   0   'False
            Height          =   375
            Left            =   3480
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmbSair 
            Caption         =   "Canc/Sair"
            Height          =   375
            Left            =   5160
            TabIndex        =   17
            Top             =   240
            Width           =   945
         End
         Begin VB.CommandButton cmdProcurar 
            Caption         =   "Procurar..."
            Height          =   375
            Left            =   2400
            TabIndex        =   3
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox txtFilial 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   0
            Top             =   360
            Width           =   435
         End
         Begin VB.TextBox txtCtc 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   840
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtNumNf 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   1965
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do CTC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   58
         Top             =   3840
         Width           =   6375
         Begin VB.Frame Frame3 
            Caption         =   "Origem"
            Height          =   615
            Left            =   120
            TabIndex        =   65
            Top             =   840
            Width           =   6135
            Begin VB.Label lblRemet 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   3495
            End
            Begin VB.Label lblRemetCid 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3720
               TabIndex        =   67
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblRemetUf 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5640
               TabIndex        =   66
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Destino"
            Height          =   615
            Left            =   120
            TabIndex        =   61
            Top             =   1510
            Width           =   6135
            Begin VB.Label lblDest 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Width           =   3495
            End
            Begin VB.Label lblDestCid 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3720
               TabIndex        =   63
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblDestUf 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5640
               TabIndex        =   62
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Notas Fiscais"
            Height          =   1335
            Left            =   120
            TabIndex        =   59
            Top             =   2160
            Width           =   6135
            Begin VB.Label lblNfs 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   975
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   5895
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Emissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   2520
            TabIndex        =   74
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblDtEmiss 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   73
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblHsEmiss 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   72
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   3960
            TabIndex        =   71
            Top             =   480
            Width           =   480
         End
         Begin VB.Label lblModal 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4680
            TabIndex        =   70
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label lblPrioridade 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NORMAL"
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
            Left            =   5160
            TabIndex        =   69
            Top             =   160
            Width           =   960
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Atuais Ocorrências deste CTC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   6375
         Begin VB.CommandButton cmdExclOcorr 
            Caption         =   "Excluir a Ocorrência Selecionada"
            Enabled         =   0   'False
            Height          =   360
            Left            =   120
            TabIndex        =   15
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CheckBox chkObsOcorr 
            Caption         =   "Comentários de Ocorrência ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2040
            Width           =   2895
         End
         Begin MSDataGridLib.DataGrid gridOcorr 
            Bindings        =   "frmPod.frx":0F1C
            Height          =   1575
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   2778
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
            DataMember      =   "Sel_ConsOcorr2"
            ColumnCount     =   10
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
               DataField       =   "cod_ocorr"
               Caption         =   "Cd."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "Ocorrência"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "usu_ocorr"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               Caption         =   "usu_dataocorr"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            BeginProperty Column07 
               DataField       =   "rel_arq_data"
               Caption         =   "rel_arq_data"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "rel_arq_num"
               Caption         =   "rel_arq_num"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "codigo"
               Caption         =   "codigo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   524,976
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   315,213
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   4589,858
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column05 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column06 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column08 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column09 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1289,764
               EndProperty
            EndProperty
         End
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
         Left            =   2535
         TabIndex        =   46
         Top             =   480
         Width           =   1800
         Begin VB.OptionButton optNf 
            Caption         =   "Por Núm de NF"
            Height          =   270
            Left            =   105
            TabIndex        =   48
            Top             =   420
            Width           =   1455
         End
         Begin VB.OptionButton optCTC 
            Caption         =   "Por Núm. de CTC"
            Height          =   255
            Left            =   105
            TabIndex        =   47
            Top             =   210
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "S T A T U S"
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
         TabIndex        =   44
         Top             =   480
         Width           =   2415
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
            TabIndex        =   45
            Top             =   360
            Width           =   2145
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ocorrência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   6600
         TabIndex        =   22
         Top             =   1320
         Width           =   5055
         Begin VB.Frame Frame6 
            Caption         =   "Para Ocorrência 01 - ENTREGA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4935
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   4815
            Begin VB.OptionButton optBaixaFinal 
               Caption         =   "Baixa Física ou Ambas"
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
               Left            =   2040
               TabIndex        =   8
               ToolTipText     =   "Considerado Data de Entrega Independente da data de Pré-Baixa"
               Top             =   360
               Width           =   2295
            End
            Begin VB.OptionButton optPreBaixa 
               Caption         =   "Pré Baixa"
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
               Left            =   480
               TabIndex        =   7
               ToolTipText     =   "Considerado com Data de Entrega na ausência de Baixa Física"
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox chkObsEntr 
               Caption         =   "Comentários/ Observações de Entrega ..."
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   4440
               Width           =   4575
            End
            Begin VB.Frame Frame9 
               Caption         =   "PRÉ-BAIXA"
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
               Left            =   120
               TabIndex        =   32
               Top             =   720
               Width           =   4575
               Begin VB.Frame fraPreBaixa 
                  Caption         =   "Dados da Pré Baixa (Emails, Telefone, etc)"
                  Height          =   1095
                  Left            =   120
                  TabIndex        =   33
                  Top             =   240
                  Width           =   4335
                  Begin VB.TextBox txtRecPreBx 
                     BackColor       =   &H8000000E&
                     Height          =   285
                     Left            =   1320
                     MaxLength       =   25
                     TabIndex        =   9
                     Top             =   720
                     Width           =   2895
                  End
                  Begin VB.Label lblHsPreBx 
                     BackColor       =   &H8000000E&
                     BorderStyle     =   1  'Fixed Single
                     Height          =   285
                     Left            =   3480
                     TabIndex        =   38
                     Top             =   360
                     Width           =   735
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     Caption         =   "Recebedor...:"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   37
                     Top             =   720
                     Width           =   975
                  End
                  Begin VB.Label Label16 
                     AutoSize        =   -1  'True
                     Caption         =   "Data Entrega:"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   36
                     Top             =   360
                     Width           =   990
                  End
                  Begin VB.Label Label17 
                     AutoSize        =   -1  'True
                     Caption         =   "Hora:"
                     Height          =   195
                     Left            =   3000
                     TabIndex        =   35
                     Top             =   360
                     Width           =   390
                  End
                  Begin VB.Label lblDtPreBx 
                     BackColor       =   &H8000000E&
                     BorderStyle     =   1  'Fixed Single
                     ForeColor       =   &H8000000D&
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   34
                     Top             =   360
                     Width           =   1455
                  End
               End
               Begin VB.CommandButton cmdExclPreBx 
                  Caption         =   "EXCLUIR esta Pré-Baixa"
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   12
                  Top             =   1320
                  Width           =   2895
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "BAIXA FÍSICA"
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
               TabIndex        =   24
               Top             =   2450
               Width           =   4575
               Begin VB.Frame fraBaixaFinal 
                  Caption         =   "Dados da Baixa Física (Com o CTC Físico)"
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   25
                  Top             =   240
                  Width           =   4335
                  Begin VB.TextBox txtRecBx 
                     BackColor       =   &H8000000E&
                     Height          =   285
                     Left            =   1320
                     MaxLength       =   25
                     TabIndex        =   10
                     Top             =   720
                     Width           =   2895
                  End
                  Begin VB.CheckBox chkCanhoto 
                     Caption         =   "Possui o Canhoto da Nota Fiscal do Cliente ?"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   26
                     Top             =   1080
                     Width           =   3495
                  End
                  Begin VB.Label Label19 
                     AutoSize        =   -1  'True
                     Caption         =   "Recebedor...:"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   31
                     Top             =   600
                     Width           =   975
                  End
                  Begin VB.Label lblDtBx 
                     BackColor       =   &H8000000E&
                     BorderStyle     =   1  'Fixed Single
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   30
                     Top             =   360
                     Width           =   1455
                  End
                  Begin VB.Label Label21 
                     AutoSize        =   -1  'True
                     Caption         =   "Data Entrega:"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   29
                     Top             =   360
                     Width           =   990
                  End
                  Begin VB.Label Label22 
                     AutoSize        =   -1  'True
                     Caption         =   "Hora:"
                     Height          =   195
                     Left            =   3000
                     TabIndex        =   28
                     Top             =   360
                     Width           =   390
                  End
                  Begin VB.Label lblHsBx 
                     BackColor       =   &H8000000E&
                     BorderStyle     =   1  'Fixed Single
                     Height          =   285
                     Left            =   3480
                     TabIndex        =   27
                     Top             =   360
                     Width           =   735
                  End
               End
               Begin VB.CommandButton cmdExclBx 
                  Caption         =   "EXCLUIR esta Baixa Física"
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   13
                  Top             =   1560
                  Width           =   3135
               End
            End
         End
         Begin VB.TextBox txtCodOcorr 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskHora 
            Height          =   285
            Left            =   2400
            TabIndex        =   5
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Hs:"
            Height          =   195
            Left            =   2160
            TabIndex        =   42
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cód. de Ocorrência:"
            Height          =   195
            Left            =   3120
            TabIndex        =   41
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label lblDescOcorr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   40
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   765
         End
      End
   End
   Begin VB.Label lblcontroletela 
      AutoSize        =   -1  'True
      Caption         =   "normal"
      Height          =   195
      Left            =   6360
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblbxfinalSim 
      Height          =   255
      Left            =   9960
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmPod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Coluna As Integer
Private Linha As Integer


Private Sub chkCanhoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub chkObsEntr_Click()
    If chkObsEntr.Value = 1 Then
        frmObsOcorr.Caption = "Observação de Entrega"
        frmObsOcorr.Show 1
        chkObsEntr.Value = 0
        chkObsOcorr.Value = 0
        cmdProcurar_Click
    End If
End Sub

Private Sub chkObsOcorr_Click()
    If chkObsOcorr.Value = 1 Then
        frmObsOcorr.Show 1
        chkObsEntr.Value = 0
        chkObsOcorr.Value = 0
        cmdProcurar_Click
    End If
End Sub

Private Sub cmbGravar_Click()

frmPod.MousePointer = 11
'tratamento de acerto aws (data de emissão)------------------------------------------------

'If mskEmissaoNova.Text <> "__/__/____" Then
'    If IsDate(mskEmissaoNova.Text) Then
'        If CDate(mskEmissaoNova) <> CDate(lblDtEmiss) Then
'            'alterar data de emissão deste CTC no tb_ctc_esp e tb_ocorr
'            de_informa.Acerto_AltDataCTC CDate(mskEmissaoNova.Text), transctc(txtFilial, txtCtc)
'            de_informa.Acerto_AltDataOcorr CDate(mskEmissaoNova.Text), transctc(txtFilial, txtCtc)
'            lblDtEmiss.Caption = mskEmissaoNova.Text
'            lblEmissao2.Caption = mskEmissaoNova.Text
'            MsgBox "OK ! Data de Emissão Alterada !"
'        End If
'    End If
'End If

'------------------------------------------------------------------------------------------

    Dim xcanhoto As String
    Dim xcontnfscanhoto As Integer
    Dim xabonodias As Long
    xcanhoto = ""
'STORE PROCEDURES ocorr1 = Dados de Pré baixa  (ALT E INS)
                 'ocorr2 = Dados de Baixa Final  (ALT E INS)
                 'ocorr3 = Dados de Pré e Baixa Final (ambas) (ALT E INS)
                 'ocorr4 = Dados de Ocorrência  (INS)
                 
'INDICAÇÕES DO tem_ocorr   (N, 0, 1 ou 2)
                 'N = indica que não há ocorrência nem baixa OU em Trânsito
                 '0 = indica processo com ocorrência, mas fechado
                 '1 = indica ctc já entregue
                 '2 = indica ctc com ocorrência, mas NÃO fechado (pendente)
                 
'TRATAMENTO DE  B A I X A S

    If Not IsDate(mskData.Text) Then
        frmPod.MousePointer = 0
        MsgBox "Data Inválida !", vbCritical, "Erro"
        mskData.SetFocus
        Exit Sub
    End If
    
    
    'tratamento acerto aws ----------------------------------------------------------------
   ' If mskEmissaoNova.Text <> "__/__/____" Then
   '     If CDate(mskData.Text) < CDate(mskEmissaoNova) Then
   '         MsgBox "Erro ! Data da Ocorrência/Entrega anterior à emissão.", vbCritical, "Erro"
   '         mskData.SetFocus
   '         Exit Sub
   '     End If
   ' End If
    '--------------------------------------------------------------------------------------

    If mskHora.Text = "__:__" Then
        mskHora.Mask = ""
        mskHora.Text = "00:00"
        mskHora.Mask = "##:##"
    End If

    If txtCodOcorr.Text = "01" Then  'se for "01" (entrega realizada/baixa)
        'verifica se campos estão digitados
        If optBaixaFinal = False And optPreBaixa = False Then
            frmPod.MousePointer = 0
            MsgBox "Escolha a forma de Baixa: Pré-Baixa ou Baixa Física (Final) !"
            Exit Sub
        End If
        If mskData.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Data", vbOKOnly + vbCritical, "ERRO"
            mskData.SetFocus
            Exit Sub
        ElseIf mskHora.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Hora", vbOKOnly + vbCritical, "ERRO"
            mskHora.SetFocus
            Exit Sub
        ElseIf txtCodOcorr.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Cod. Ocorrência", vbOKOnly + vbCritical, "ERRO"
            txtCodOcorr.SetFocus
            Exit Sub
        ElseIf optBaixaFinal = True Then
            If txtRecBx.Text = "" Then
                frmPod.MousePointer = 0
                MsgBox "Dados Inválidos ! Campo: Recebedor (Baixa Final)", vbOKOnly + vbCritical, "ERRO"
                txtRecBx.SetFocus
                Exit Sub
            End If
        ElseIf optPreBaixa = True Then
            If txtRecPreBx.Text = "" Then
                frmPod.MousePointer = 0
                MsgBox "Dados Inválidos ! Campo: Recebedor (Pré Baixa)", vbOKOnly + vbCritical, "ERRO"
                txtRecPreBx.SetFocus
                Exit Sub
            End If
        End If
        'verifica se o CTC já tem ocorrência fechada cadastrada. Caso tenha não possibilita a baixa
        If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
            frmPod.MousePointer = 0
            MsgBox "Este CTC está Baixado, indicando que esta entrega não ocorreria. Caso deseje baixar como ENTREGA, você deve primeiro excluir a Ocorrência  C T C   B A I X A D O", vbOKCancel
            txtCodOcorr.SetFocus
            Exit Sub
        'verifica se o ctc já possui ocorrência e se a data que se quer baixar não é menor que a data de alguma ocorrência
        ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
            de_informa.rsSel_ConsOcorr2.MoveFirst
            Do Until de_informa.rsSel_ConsOcorr2.EOF
                If CDate(mskData.Text) < de_informa.rsSel_ConsOcorr2.Fields("data") Then
                    frmPod.MousePointer = 0
                    If MsgBox("Você está tentando baixar um CTC com data Menor que uma Ocorrência Cadastrada. Você tem certeza que quer baixar este CTC como entrega nesta data ?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
                        mskData.SetFocus
                        Exit Sub
                    Else
                        Exit Do
                    End If
                End If
                de_informa.rsSel_ConsOcorr2.MoveNext
            Loop
        End If
            
'início do processo de baixa.

        If optPreBaixa.Value = True Then         'se for uma pré baixa
            'procura se o CTC já está baixado pré
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr transctc(TxtFilial.Text, txtCtc.Text), "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
               'Este CTC já Contém Pré Baixa.
                frmPod.MousePointer = 0
                MsgBox "Este CTC já está baixado com Pré-Baixa. Caso esteja tentando alterar a data que já está cadastrada, você deve antes excluir esta Pré-Baixa e lançar novamente com a data correta. Exclusão de Entrega/Ocorrências só pode ser realizada por usuário que possui este direito de acesso. Esta informação não será gravada no sistema.", vbOKOnly + vbExclamation
                mskData.SetFocus
                Exit Sub
                
                    'frmPod.MousePointer = 11
                    'de_informa.cn_informa.BeginTrans
                    'If de_informa.rsSel_ConsOcorr.Fields("baixadofinal") = "S" Then
                    '   de_informa.alt_ocorr1ow transctc(txtFilial.Text, txtCtc.Text), mskData.Text, mskHora.Text, RTrim(txtRecPreBx.Text), xusuario, CVar(Date) & " " & CVar(Time())
                    '   de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                    'Else
                    '   'atual_sitla=S  =>  atualizar o sistema SITLA
                    '   de_informa.alt_ocorr1 transctc(txtFilial.Text, txtCtc.Text), mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecPreBx.Text), xusuario, CVar(Date) & " " & CVar(Time()), "S", Date
                    '   de_informa.Alt_AtClienteNFBranco transctc(txtFilial.Text, txtCtc.Text)
                    '   de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                    'End If
                    
                    'LOG DE USUÁRIO
                    'de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA"
                    'de_informa.cn_informa.CommitTrans

            Else 'SE NÃO HOUVER NENHUMA BAIXA, INCLUI ...  (INSERT "01")
                de_informa.cn_informa.BeginTrans
                de_informa.ins_ocorr1 transctc(TxtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecPreBx.Text), xusuario, datahora("datahora"), "S", datahora("data")
                de_informa.alt_temocorr_sn "1", transctc(TxtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                de_informa.Alt_AtClienteNFBranco transctc(TxtFilial.Text, txtCtc.Text)
                    
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA"
                de_informa.cn_informa.CommitTrans
                
                'atualiza os prazos
                frmAtualPrazos.lblFilialctc = transctc(TxtFilial.Text, txtCtc.Text)
                frmAtualPrazos.Show 1
            End If
        ElseIf optBaixaFinal.Value = True Then   'se for uma baixa final ou ambas
            If txtRecBx.Text = "" Then
                frmPod.MousePointer = 0
               MsgBox "Dados Inválidos ! Campo: Recebedor", vbOKOnly + vbCritical, "ERRO"
               txtRecBx.SetFocus
               Exit Sub
            End If
            If chkCanhoto.Value = 1 Then
                frmContrCanhotos.lstPresentes.Clear
                frmContrCanhotos.lstFaltantes.Clear
                xcanhoto = "S"
                If chkCanhoto.Enabled = True Then
                    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
                    de_informa.Sel_NFsdoCTC transctc(TxtFilial.Text, txtCtc.Text)
                    If de_informa.rsSel_NFsdoCTC.RecordCount > 0 Then
                        Do Until de_informa.rsSel_NFsdoCTC.EOF
                            frmContrCanhotos.lstPresentes.AddItem de_informa.rsSel_NFsdoCTC.Fields("numnf")
                            de_informa.rsSel_NFsdoCTC.MoveNext
                        Loop
                        frmContrCanhotos.lblFilialctc = transctc(TxtFilial.Text, txtCtc.Text)
                        frmContrCanhotos.fraPresentes.Caption = frmContrCanhotos.lstPresentes.ListCount & " Canhotos"
                        frmContrCanhotos.Show 1
                        If lblcontroletela.Caption = "cancelar" Then
                            lblcontroletela.Caption = "normal"
                            Unload frmContrCanhotos
                            Me.MousePointer = 0
                            cmdProcurar_Click
                            Exit Sub
                        End If
                    Else
                        xcanhoto = "N"  'pois não há NFS
                    End If
                End If
            Else
                xcanhoto = "N"
            End If
            'procura se o CTC já está baixado final
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr transctc(TxtFilial.Text, txtCtc.Text), "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
                If de_informa.rsSel_ConsOcorr.Fields("baixadofinal") = "S" Then
                    'Este CTC já está baixado (ambos ou final)
                    frmPod.MousePointer = 0
                    MsgBox "Este CTC já está baixado com Baixa-Física. Caso esteja tentando alterar a data que já está cadastrada, você deve antes excluir esta Baixa-Física e lançar novamente com a data correta. Exclusão de Entrega/Ocorrências só pode ser realizada por usuário que possui este direito de acesso. Esta informação não será gravada no sistema.", vbOKOnly + vbExclamation
                    Exit Sub
                        
                        'frmPod.MousePointer = 11
                        'inicia a transação
                        'de_informa.cn_informa.BeginTrans
                        
                        'atualiza com os dados de baixa física
                        'de_informa.alt_ocorr2 transctc(txtfilial.Text, txtCTC.Text), mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecBx.Text), xusuario, CVar(Date) & " " & CVar(Time()), "S", Date, xcanhoto
                        'de_informa.alt_temocorr_sn "1", transctc(txtfilial.Text, txtCTC.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                        
                        'atualiza as NFs que contém o canhoto
                        
                        'For xcontnfscanhoto = 1 To frmContrCanhotos.lstPresentes.ListCount
                        '    de_informa.Alt_CanhotoNFSN "S", transctc(txtfilial.Text, txtCTC.Text), frmContrCanhotos.lstPresentes.List(xcontnfscanhoto - 1)
                        'Next
                        
                        'atualiza as NFs que NÃO contém o canhoto
                        
                        'For xcontnfscanhoto = 1 To frmContrCanhotos.lstFaltantes.ListCount
                        '    de_informa.Alt_CanhotoNFSN "N", transctc(txtfilial.Text, txtCTC.Text), frmContrCanhotos.lstFaltantes.List(xcontnfscanhoto - 1)
                        'Next
                        
                        'atualiza status de envio de informação para o cliente
                        'de_informa.Alt_AtClienteNFBranco transctc(txtfilial.Text, txtCTC.Text)
                        
                        'lblbxfinalSim = "SIM" 'identifica label invisível como SIM para controle se executa pergunta de relatório do Protocolo para Setor de Arquivo
                        
                        'LOG DE USUÁRIO
                        'de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtfilial.Text, txtCTC.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FINAL (FÍSICA)"
                        
                        'finaliza transação
                        'de_informa.cn_informa.CommitTrans
                        
                Else  'SE NÃO HOUVER BAIXA FINAL, INCLUI NO REGISTRO DE BAIXA (UPDATE "01" BAIXA FINAL)
                
                    'Se Não há baixa Física é porque há baixa final, pois o mesmo já está baixado !
                    
                    If CDate(mskData) <> CDate(lblDtPreBx) Then
                    
                        Me.MousePointer = 0
                        
                        If MsgBox("ATENÇÃO ! A Data desta Baixa Física que você está querendo cadastrar é diferente da data da Pré-Baixa que já está cadastrada para este CTC/NF." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Você tem certeza que esta informação está correta e deseja realmente gravar esta Baixa-Física com a data diferente da Pré-Baixa ?", vbYesNo + vbQuestion + vbCritical, "ATENÇÃO") = vbNo Then
                            MsgBox "OK ! Operação Cancelada. Esta informação não será gravada no sistema."
                            txtRecBx.Text = ""
                            txtCodOcorr.Text = ""
                            mskData.SetFocus
                            Exit Sub
                        End If
                        
                    End If
                        
                    Me.MousePointer = 11
                    'inicia a transação
                    de_informa.cn_informa.BeginTrans
                                                
                    'atualiza os dados de entrega
                    de_informa.alt_ocorr2 transctc(TxtFilial.Text, txtCtc.Text), CDate(lblDtPreBx), lblHsPreBx, mskData.Text, mskHora.Text, RTrim(txtRecBx.Text), xusuario, datahora("datahora"), "N", datahora("data"), xcanhoto
                    de_informa.alt_temocorr_sn "1", transctc(TxtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                    
                    'atualiza as NFs que contém o canhoto
                        
                    For xcontnfscanhoto = 1 To frmContrCanhotos.lstPresentes.ListCount
                        de_informa.Alt_CanhotoNFSN "S", transctc(TxtFilial.Text, txtCtc.Text), frmContrCanhotos.lstPresentes.List(xcontnfscanhoto - 1)
                    Next
                        
                    'atualiza as NFs que NÃO contém o canhoto
                        
                    For xcontnfscanhoto = 1 To frmContrCanhotos.lstFaltantes.ListCount
                        de_informa.Alt_CanhotoNFSN "N", transctc(TxtFilial.Text, txtCtc.Text), frmContrCanhotos.lstFaltantes.List(xcontnfscanhoto - 1)
                    Next
                    
                    Unload frmContrCanhotos
                    
                    'atualiza o campo de informação para o cliente
                    de_informa.Alt_AtClienteNFBranco transctc(TxtFilial.Text, txtCtc.Text)
                    
                    lblbxfinalSim = "SIM" 'identifica label invisível como SIM para controle se executa pergunta de relatório do Protocolo para Setor de Arquivo
                        
                    'LOG DE USUÁRIO
                    de_informa.ins_LogUsuario "INCLUSAO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FINAL/FÍSICA (JÁ HAVIA PRÉ-BAIXA)"
                    
                    'finaliza Transação
                    de_informa.cn_informa.CommitTrans
                    
                    'atualiza os prazos
                    frmAtualPrazos.lblFilialctc = transctc(TxtFilial.Text, txtCtc.Text)
                    frmAtualPrazos.Show 1
                            
                End If
                
            Else  'SE NÃO HOUVER NENHUMA BAIXA, INCLUI AMBOS (PRE E FINAL)
            
                'inicia a transação
                de_informa.cn_informa.BeginTrans
                
                'atualiza os dados de entrega
                de_informa.ins_ocorr3 transctc(TxtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecPreBx), mskData.Text, mskHora.Text, RTrim(txtRecBx.Text), xusuario, datahora("datahora"), "S", datahora("data"), xcanhoto 'insere aS baixas ambas
                
                'atualiza o status do CTC
                de_informa.alt_temocorr_sn "1", transctc(TxtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                
                'atualiza as NFs que contém o canhoto
                        
                For xcontnfscanhoto = 1 To frmContrCanhotos.lstPresentes.ListCount
                    de_informa.Alt_CanhotoNFSN "S", transctc(TxtFilial.Text, txtCtc.Text), frmContrCanhotos.lstPresentes.List(xcontnfscanhoto - 1)
                Next
                        
                'atualiza as NFs que NÃO contém o canhoto
                        
                For xcontnfscanhoto = 1 To frmContrCanhotos.lstFaltantes.ListCount
                    de_informa.Alt_CanhotoNFSN "N", transctc(TxtFilial.Text, txtCtc.Text), frmContrCanhotos.lstFaltantes.List(xcontnfscanhoto - 1)
                Next
                
                Unload frmContrCanhotos
                
                lblbxfinalSim = "SIM" 'identifica label invisível como SIM para controle se executa pergunta de relatório do Protocolo para Setor de Arquivo
                
                'atualiza campo para informação para o cliente
                de_informa.Alt_AtClienteNFBranco transctc(TxtFilial.Text, txtCtc.Text)
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "INCLUSAO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA + BAIXA FINAL/FÍSICA"
                
                'finaliza transação
                de_informa.cn_informa.CommitTrans
                
                'atualiza os prazos
                frmAtualPrazos.lblFilialctc = transctc(TxtFilial.Text, txtCtc.Text)
                frmAtualPrazos.Show 1
                
            End If
        End If
        
'TRATAMENTO DE  O C O R R Ê N C I A S
        
    Else   'se nao for baixa (ocorr # 01) então é somente ocorrência
        'verifica se campos estão digitados
        If mskData.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Data", vbOKOnly + vbCritical, "ERRO"
            mskData.SetFocus
            Exit Sub
        ElseIf mskHora.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Hora", vbOKOnly + vbCritical, "ERRO"
            mskHora.SetFocus
            Exit Sub
        ElseIf txtCodOcorr.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Cod. Ocorrência", vbOKOnly + vbCritical, "ERRO"
            txtCodOcorr.SetFocus
            Exit Sub
        End If
        
        If txtCodOcorr.Text = "00" Then   'se for ocorr 00
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then
                frmPod.MousePointer = 0
                MsgBox "CTC já Baixado como Entregue. Não é Possível informar Ocorrência  C T C   B A I X A D O"
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
                frmPod.MousePointer = 0
                MsgBox "CTC já possui Ocorrência  C T C   B A I X A D O"
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "N" Then
                frmPod.MousePointer = 0
                MsgBox "Você só pode informar Ocorrência  C T C   B A I X A D O, se o CTC já tiver alguma ocorrência que a explique o motivo."
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
                de_informa.rsSel_ConsOcorr2.MoveFirst
                Do Until de_informa.rsSel_ConsOcorr2.EOF
                   If CDate(mskData.Text) < de_informa.rsSel_ConsOcorr2.Fields("data") Then
                       MsgBox "A Data da Baixa Deve ser maior ou igual a última ocorrência cadastrada.", vbOKOnly, "Erro"
                       mskData.SetFocus
                       Exit Sub
                    End If
                    de_informa.rsSel_ConsOcorr2.MoveNext
                Loop
            End If
        Else   'se não for é ocorrência normal
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then
                If IsDate(lblDtBx.Caption) Then
                    If CDate(mskData.Text) > CDate(lblDtBx.Caption) Then
                        frmPod.MousePointer = 0
                        If MsgBox("Você está tentando incluir uma Ocorrência com data Posterior à sua Data de Entrega. Você tem certeza que deseja informar esta ocorrência com esta data ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        End If
                    End If
                ElseIf IsDate(lblDtPreBx.Caption) Then
                    If CDate(mskData.Text) > CDate(lblDtPreBx.Caption) Then
                        frmPod.MousePointer = 0
                        If MsgBox("Você está tentando incluir uma Ocorrência com data Posterior à Data de Entrega. Você tem certeza que deseja informar esta ocorrência com esta data ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
                de_informa.rsSel_ConsOcorr2.MoveFirst
                Do Until de_informa.rsSel_ConsOcorr2.EOF
                    If de_informa.rsSel_ConsOcorr2.Fields("cod_ocorr") = "00" Then
                        If CDate(mskData.Text) > de_informa.rsSel_ConsOcorr2.Fields("data") Then
                            frmPod.MousePointer = 0
                            If MsgBox("Você está tentando lançar uma ocorrência com data posterior a uma ocorrência  C T C   B A I X A D O", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                                mskData.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    de_informa.rsSel_ConsOcorr2.MoveNext
                Loop
            End If
        End If
        
        'ATUALIZA BD COM OS DADOS DA OCORRÊNCIA
        
        de_informa.cn_informa.BeginTrans
        
        If txtCodOcorr.Text = "00" Then   'se for ocorr 00
            de_informa.ins_ocorr4cod00 transctc(TxtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
            txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xusuario, datahora("datahora")
            '0 = IDENTIFICA COMO CTC COM OCORRÊNCIA FECHADA
            'If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
            de_informa.alt_temocorr_sn "0", transctc(TxtFilial.Text, txtCtc.Text)   'atualiza arquivo de CTC com tem_ocorr = 0
            'End If
        Else  'se for outro tipo de ocorrência
            de_informa.ins_ocorr4 transctc(TxtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
            txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xusuario, datahora("datahora")
            '2 = IDENTIFICA COMO CTC COM OCORRÊNCIAS PENDENTE
            de_informa.Alt_AtClienteNFBranco transctc(TxtFilial.Text, txtCtc.Text)
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") <> "1" And de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") <> "0" Then
                de_informa.alt_temocorr_sn "2", transctc(TxtFilial.Text, txtCtc.Text)   'atualiza arquivo de CTC com tem_ocorr = 2
                If txtCodOcorr = "39" Or txtCodOcorr = "84" Then  'pre-baixa automática por ser CTC/NF Retido para COnferência
                    If MsgBox("Você está lançando uma ocorrência de retenção de Doctos. para conferência, onde provavelmente a entrega foi realizada nesta data." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "O Informa pode incluir automaticamente uma Pré-Baixa para este CTC nesta data. Você Confirma ?", vbYesNo, "Pré-Baixa Automática") = vbYes Then
                        de_informa.ins_ocorr1 transctc(TxtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), "01", "ENTREGA REALIZADA", _
                        mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, ".", "AUTO-PREBX", datahora("datahora"), "S", datahora("data")
                        de_informa.alt_temocorr_sn "1", transctc(TxtFilial.Text, txtCtc.Text)
                        MsgBox "Pré-Baixa Automática Gravada !"
                    End If
                End If
            End If
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then
                If txtCodOcorr = "26" Or txtCodOcorr = "85" Then  'abono automático para atraso
                    If de_informa.rsSel_CTCEntrega.State = 1 Then de_informa.rsSel_CTCEntrega.Close
                    de_informa.Sel_CTCEntrega transctc(TxtFilial.Text, txtCtc.Text)
                    If de_informa.rsSel_CTCEntrega.RecordCount > 0 Then
                        If de_informa.rsSel_CTCEntrega.Fields("diasuteis") - _
                           de_informa.rsSel_CTCEntrega.Fields("abonodias") > _
                           de_informa.rsSel_CTCEntrega.Fields("prazoentr") Then
                           'está em atraso, lançar abono automático
                           xabonodias = de_informa.rsSel_CTCEntrega.Fields("diasuteis") - de_informa.rsSel_CTCEntrega.Fields("prazoentr")
                           de_informa.Alt_AbonoAtraso xabonodias, "AUTOMATIC", datahora("DATAHORA"), "Abono Automático Devido Ocorrência", transctc(TxtFilial.Text, txtCtc.Text)
                        End If
                    End If
                End If
            End If
        End If
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSAO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr
        
        de_informa.cn_informa.CommitTrans
        
    End If
    
    mskData.Mask = ""
    mskData.Text = ""
    mskData.Mask = "##/##/####"
    'mskData.Enabled = False
    'mskData.BackColor = &H8000000E   'branco
    mskHora.Mask = ""
    mskHora.Text = ""
    mskHora.Mask = "##:##"
    'mskHora.Enabled = False
    'mskHora.BackColor = &H8000000E   'branco
    

    txtCodOcorr = ""
    'txtCodOcorr.Enabled = False
    'txtCodOcorr.BackColor = &H8000000E   'branco
    lblDescOcorr.Caption = ""
    txtRecPreBx.BackColor = &H8000000E   'branco
    txtRecBx.BackColor = &H8000000E   'branco
    cmbGravar.Enabled = False
    frmPod.MousePointer = 0
    cmdProcurar_Click
    MsgBox "OK ! OCORRÊNCIA REGISTRADA.", vbOKOnly + vbExclamation
    TxtFilial.SetFocus
End Sub
Private Sub cmbSair_Click()
    If mskData.Text <> "__/__/____" Then   'quer dizer que é CANCELAR
        mskData.Mask = ""
        mskData.Text = ""
        mskData.Mask = "##/##/####"
        mskHora.Mask = ""
        mskHora.Text = ""
        mskHora.Mask = "##:##"
        txtCodOcorr = ""
        lblDescOcorr.Caption = ""
        txtRecPreBx.Text = ""
        txtRecBx.Text = ""
        txtRecPreBx.BackColor = &H8000000E   'branco
        txtRecBx.BackColor = &H8000000E   'branco
        cmdProcurar_Click
        TxtFilial.SetFocus
    Else                            'caso contrário é SAIR
        'frmAtualPrazos.Show 1
        'If lblbxfinalSim = "SIM" Then
            'If MsgBox("Deseja Imprimir o Relatório de CTCs Físicos Baixados, para Envio dos Documentos para o Arquivo ? (PROTOCOLO)", vbQuestion + vbYesNo, "Confirmação de Relatório") = vbYes Then
            '    mdiInforma.StatusBar1.Panels.Item(2).Text = "AGUARDE IMPRESSÃO DO RELATÓRIO ..."
            '    DoEvents
            '    Call rel_arquivo
            '    mdiInforma.StatusBar1.Panels.Item(2).Text = ""
            '    DoEvents
            'End If
        'End If
        Set frmPod = Nothing
        Unload Me
    End If
End Sub

Private Sub cmdComentario_Click()
    frmObsOcorr.Show 1
End Sub

Private Sub cmdCalendario_Click()

End Sub
Private Sub cmdBuscaCTCs_Click()
    Dim X As Integer
    
    MSFlexGrid1.Rows = 1
    DoEvents
    MSFlexGrid1.Rows = 2
    DoEvents
    MSFlexGrid1.FixedRows = 1
    DoEvents
    
    If de_informa.rsSel_ManifestoPorNum.State = 1 Then de_informa.rsSel_ManifestoPorNum.Close
    de_informa.Sel_ManifestoPorNum transmanif(txtFilialmnf, txtManifesto)

    If de_informa.rsSel_ManifestoPorNum.RecordCount < 1 Then
        
        lblFilialManifesto = ""
        lblDataMnf = ""
        lblPropMnf = ""
        lblPlacaMnf = ""
        lblMotoristaMnf = ""
        lblQtdeCtcMnf = ""
        MsgBox "Número de Filial-Manifesto Não Encontrado !!", vbCritical
        
    Else
        
        lblFilialManifesto = transmanif(txtFilialmnf, txtManifesto)
        MSFlexGrid1.Rows = de_informa.rsSel_ManifestoPorNum.RecordCount + 1
        
        lblDataMnf = de_informa.rsSel_ManifestoPorNum.Fields("dtemissao")
        lblPropMnf = de_informa.rsSel_ManifestoPorNum.Fields("proprietario")
        lblPlacaMnf = de_informa.rsSel_ManifestoPorNum.Fields("placaveic")
        lblMotoristaMnf = de_informa.rsSel_ManifestoPorNum.Fields("motorista")
        lblQtdeCtcMnf = de_informa.rsSel_ManifestoPorNum.RecordCount
        
        For X = 1 To de_informa.rsSel_ManifestoPorNum.RecordCount
            MSFlexGrid1.TextMatrix(X, 1) = de_informa.rsSel_ManifestoPorNum.Fields("filialctc")
            MSFlexGrid1.TextMatrix(X, 2) = de_informa.rsSel_ManifestoPorNum.Fields("data")
            MSFlexGrid1.TextMatrix(X, 3) = de_informa.rsSel_ManifestoPorNum.Fields("remet_nome")
            MSFlexGrid1.TextMatrix(X, 4) = de_informa.rsSel_ManifestoPorNum.Fields("dest_nome")
            MSFlexGrid1.TextMatrix(X, 5) = de_informa.rsSel_ManifestoPorNum.Fields("cidade_dest")
            MSFlexGrid1.TextMatrix(X, 6) = de_informa.rsSel_ManifestoPorNum.Fields("uf_dest")
            MSFlexGrid1.TextMatrix(X, 7) = Format(de_informa.rsSel_ManifestoPorNum.Fields("valmerc"), "##,###,##0.00")
            MSFlexGrid1.TextMatrix(X, 8) = Format(de_informa.rsSel_ManifestoPorNum.Fields("peso"), "##,##0.0")
            MSFlexGrid1.TextMatrix(X, 9) = Trim$(de_informa.rsSel_ManifestoPorNum.Fields("nfs"))
            de_informa.rsSel_ManifestoPorNum.MoveNext
            Call ColocaCheck(X, 0, frmPod)  'Chama Função Para Colocar o CheckBox
            DoEvents
            Call VerificaCheck(X, 0, frmPod) 'Chama Função Para Marcar ou Desmarcar
            DoEvents
        Next
        
        mskDataMnf.BackColor = xamarelo1
        mskDataMnf.Enabled = True
        
        de_informa.rsSel_ManifestoPorNum.MoveFirst
        
        MSFlexGrid1.TextMatrix(0, 0) = "S/N"
        
        MSFlexGrid1.SetFocus
        
        'substituir por control + home
        
        For X = 1 To de_informa.rsSel_ManifestoPorNum.RecordCount
            SendKeys "{UP}"
        Next
        SendKeys "{UP}"
        
    End If
    

End Sub

Private Sub cmdExclBx_Click()
    If Mid$(xdireitos, 23, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        'Alteração no Registro de Baixa, (UPDATE) para BAIXAFINAL = N
        'Atualização do Campo DATA do TB_OCORR para a Data da Pré-Baixa
        If MsgBox("Confirma Exclusão dos Dados de BAIXA FÍSICA ? ", vbYesNo, "Atenção") = vbYes Then
        
            de_informa.cn_informa.BeginTrans
            
            de_informa.Alt_ExclBaixaFisica transctc(TxtFilial, txtCtc)
            
            de_informa.alt_ExclCanhotoNF transctc(TxtFilial, txtCtc)
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " BAIXA FÍSICA"
            
            de_informa.cn_informa.CommitTrans
            
            cmdProcurar_Click
        End If
    End If
End Sub

Private Sub cmdExclOcorr_Click()
    If Mid$(xdireitos, 22, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        Dim xcodocorr As String
        If MsgBox("Confirma a Exclusão da Ocorrência Selecionada ?", vbQuestion + vbYesNo, "Exclusão") = vbYes Then
            xcodocorr = GridOcorr.Columns(2)
            
            de_informa.cn_informa.BeginTrans
            
            de_informa.excl_ocorr GridOcorr.Columns(9)
            If xcodocorr = "00" Then 'se for "00" altera o temocorr para 02
                de_informa.alt_temocorr_sn "2", transctc(TxtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr
            End If
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " OCORR:" & GridOcorr.Columns(2) & "-" & GridOcorr.Columns(3)
            
            'SE HOUVER HOUVER ENTREGA, ABONODIAS = 0 DEVIDO EXCLUSÃO DE OCORRÊNCIA
            
            If de_informa.rsSel_CTCEntrega.State = 1 Then de_informa.rsSel_CTCEntrega.Close
            de_informa.Sel_CTCEntrega transctc(TxtFilial.Text, txtCtc.Text)
            
            If de_informa.rsSel_CTCEntrega.RecordCount > 0 Then
                de_informa.Alt_ExclAbono transctc(TxtFilial.Text, txtCtc.Text)
                MsgBox "Caso Este CTC tenha algum Abono de Atraso, este abono foi excluido."
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "ABONO ATRASO:" & transctc(TxtFilial.Text, txtCtc.Text) & " DEVIDO EXCLUSÃO DE OCORRENCIA."
            End If
            
            'busca as ocorrências e atualiza o grid de ocorrências
            
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 transctc(TxtFilial, txtCtc), "01"
            Set GridOcorr.DataSource = de_informa
            GridOcorr.DataMember = "Sel_ConsOcorr2"
            GridOcorr.Refresh
            
         'verifica se é a última ocorrência baixada e se é ocorr "00"
        'se for exclui ela também (pois o processo não está finalizado) e atualiza o grid novamente
 
            If de_informa.rsSel_ConsOcorr2.RecordCount = 1 Then
                If de_informa.rsSel_ConsOcorr2.Fields("cod_ocorr") = "00" Then
                    de_informa.excl_ocorr GridOcorr.Columns(9)
                    If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                    de_informa.Sel_ConsOcorr2 transctc(TxtFilial, txtCtc), "01"
                    Set GridOcorr.DataSource = de_informa
                    GridOcorr.DataMember = "Sel_ConsOcorr2"
                    GridOcorr.Refresh
                End If
            End If
            
        'se o grid estiver vazio e se não estiver baixa o CTC atualiza o temocorr para "N" (sem posição)
            If de_informa.rsSel_ConsOcorr2.RecordCount = 0 And de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") <> "1" Then  'verifica se não há mais ocorrência e se não está baixado
                de_informa.alt_temocorr_sn "N", transctc(TxtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr
            End If
                
            de_informa.cn_informa.CommitTrans
                
            cmdProcurar_Click
        End If
    End If
End Sub

Private Sub CmdObsEntr_Click()
    frmObsOcorr.Show 1
End Sub
Private Sub cmdGravarPreBxMnf_Click()
    Dim X As Integer, X2 As Integer, xDataHora As Date
    
    If Not IsDate(mskDataMnf) Then
        MsgBox "Atenção !!! A Data que você Digitou Para Pré-Baixa é Inválida.", vbCritical, "ERRO"
        mskDataMnf.SetFocus
        Exit Sub
    End If

    If MsgBox("Confira Atentamente a Data que Você Está Informando como Pré-Baixa !!!" _
              + Chr(10) + Chr(13) + Chr(10) + Chr(13) + _
              "Você Confirma Lançar Pré-Baixas com a Data " & mskDataMnf & " Para os CTCs do Manifesto " & txtFilialmnf & "-" & txtManifesto & " ? ", vbYesNo + vbQuestion, "ATENÇÃO") = vbYes Then
        
        SSTab1.Enabled = False
        frmPod.MousePointer = 11
        DoEvents
        
        xDataHora = datahora("DATAHORA")
    
        For X = 1 To MSFlexGrid1.Rows - 1
        
            If MSFlexGrid1.TextMatrix(X, 0) = "þ" Then   'verifica se está checado !!
            
                If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
                de_informa.Sel_Ctc_SAC MSFlexGrid1.TextMatrix(X, 1)
                
                If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "C" Then
                    MsgBox "O CTC " & MSFlexGrid1.TextMatrix(X, 1) & " encontra-se CANCELADO !! Não Será Possível Informar Posição de Entrega/Pré-Baixa.", vbCritical, "ERRO"
                    MSFlexGrid1.Row = X
                    For X2 = 0 To 9
                        MSFlexGrid1.Col = X2
                        MSFlexGrid1.CellBackColor = &HC0C0FF   'VERMELHO
                    Next
                ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
                    MsgBox "O CTC " & MSFlexGrid1.TextMatrix(X, 1) & " encontra-se BAIXADO SEM ENTREGA !! Não Será Possível Informar Posição de Entrega/Pré-Baixa.", vbCritical, "ERRO"
                    MSFlexGrid1.Row = X
                    For X2 = 0 To 9
                        MSFlexGrid1.Col = X2
                        MSFlexGrid1.CellBackColor = &HC0C0FF   'VERMELHO
                    Next
                ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then
                    MsgBox "O CTC " & MSFlexGrid1.TextMatrix(X, 1) & " já encontra-se Baixado com Data de Entrega !! Não Será Possível Informar Nova Posição de Entrega/Pré-Baixa.", vbCritical, "ERRO"
                    MSFlexGrid1.Row = X
                    For X2 = 0 To 9
                        MSFlexGrid1.Col = X2
                        MSFlexGrid1.CellBackColor = &HC0C0FF   'VERMELHO
                    Next
                ElseIf de_informa.rsSel_Ctc_SAC.Fields("data") > CDate(mskDataMnf) Then
                    MsgBox "O CTC " & MSFlexGrid1.TextMatrix(X, 1) & " possue Data de Emissão Posterior a Data Que Está Tentando Lançar a Pré-Baixa !! Este CTC não será baixado.", vbCritical, "ERRO"
                    MSFlexGrid1.Row = X
                    For X2 = 0 To 9
                        MSFlexGrid1.Col = X2
                        MSFlexGrid1.CellBackColor = &HC0C0FF   'VERMELHO
                    Next
                Else
                    de_informa.cn_informa.BeginTrans
                        de_informa.ins_ocorr1 MSFlexGrid1.TextMatrix(X, 1), de_informa.rsSel_Ctc_SAC.Fields("data"), _
                                              de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), "01", "ENTREGA REALIZADA", _
                                              mskDataMnf.Text, "00:00", mskDataMnf.Text, "00:00", "", xusuario, xDataHora, "S", xDataHora
                        
                        de_informa.alt_temocorr_sn "1", MSFlexGrid1.TextMatrix(X, 1)  'atualiza arquivo de CTC com tem_ocorr = 1
                        de_informa.Alt_AtClienteNFBranco MSFlexGrid1.TextMatrix(X, 1)
                            
                        'LOG DE USUÁRIO
                        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC:" & MSFlexGrid1.TextMatrix(X, 1) & " OCORR:01-ENTREGA REALIZADA PRÉ-BAIXA(PELO MANIFESTO)"
                    de_informa.cn_informa.CommitTrans
                    
                    'atualiza os prazos
                    frmAtualPrazos.lblFilialctc = MSFlexGrid1.TextMatrix(X, 1)
                    frmAtualPrazos.Show 1
                    
                    MSFlexGrid1.Row = X
                    For X2 = 0 To 9
                        MSFlexGrid1.Col = X2
                        MSFlexGrid1.CellBackColor = &HFFC0C0       'AZUL
                    Next
                
                End If
            
            End If
            
        Next X
        
        MsgBox "Final do Processamento !!", vbInformation
        
        frmPod.MousePointer = 0
        SSTab1.Enabled = True
    
    End If
End Sub

Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
Me.PrintForm
    
End Sub

Private Sub cmdImprTela2_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdProcurar_Click()
    If optNf.Value = True Then  'Se a procura for por NF
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
                frmDuplNF.Caption = "POD - Número de NFs Duplicadas"
                DoEvents
                frmDuplNF.Show 1  'direciona para o form que trata casos de NF duplicadas
                Exit Sub
            Else  'Caso seja encontrada somente uma NF
                optCTC_Click
                TxtFilial.Text = Mid(de_informa.rsSel_NF_SAC.Fields("filialctc"), 1, 2)
                txtCtc.Text = Mid(de_informa.rsSel_NF_SAC.Fields("filialctc"), 3, 8) 'Busca a Filial e o CTC com base na NF
            End If
    End If
    optCTC.Value = True
        Dim xtemocorr As String
        If TxtFilial.Text = "" Or txtCtc.Text = "" Then
            MsgBox "Filial / CTC Inválidos !", vbCritical, "Erro"
            Exit Sub
        End If
        If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
        de_informa.Sel_Ctc_SAC transctc(TxtFilial, txtCtc)  'Procura na Tabela a Filial/CTC
        If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then
            MsgBox "Número de Filial/CTC Não Encontrados !", vbCritical + vbOKOnly, "Erro"
            TxtFilial.SetFocus
            Exit Sub
        End If
'REGISTRA VARIÁVEIS GLOBAIS DE FILIAL E CTC PARA UTILIZAÇÃO EM OUTROS FORMS
        xultimofilial = TxtFilial.Text
        xultimoctc = txtCtc.Text
'ATUALIZA DADOS DO CTC NO FORM
        lblDtEmiss.Caption = de_informa.rsSel_Ctc_SAC.Fields("data")
        
        
        'tratamento de data de emissão (acerto aws)--------------------------------------
        'mskEmissaoNova.Text = de_informa.rsSel_Ctc_SAC.Fields("data")
        'lblEmissao2 = de_informa.rsSel_Ctc_SAC.Fields("data")
        '--------------------------------------------------------------------------------
        
        
        lblHsEmiss.Caption = de_informa.rsSel_Ctc_SAC.Fields("hora")
        lblRemet.Caption = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
        lblRemetCid.Caption = de_informa.rsSel_Ctc_SAC.Fields("cidade_orig")
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli de_informa.rsSel_Ctc_SAC.Fields("remet_cgc")
        lblRemetUf = de_informa.rsSel_ConsCadCli.Fields("uf")
        lblDest.Caption = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
        lblDestCid.Caption = de_informa.rsSel_Ctc_SAC.Fields("cidade_dest")
        lblDestUf.Caption = de_informa.rsSel_Ctc_SAC.Fields("uf_dest")
        lblNfs.Caption = de_informa.rsSel_Ctc_SAC.Fields("nfs")
        lblModal.Caption = de_informa.rsSel_Ctc_SAC.Fields("modal")
        
        If de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "URGÊNCIA" Or _
            de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "PRIORIDADE" Then
            LblPrioridade.ForeColor = &HC0&
        Else
            LblPrioridade.ForeColor = &H80000012
        End If
        LblPrioridade = de_informa.rsSel_Ctc_SAC.Fields("prioridade")
        
        'LIMPA OS DADOS DO FORM
        mskData.Mask = ""
        mskData.Text = ""
        mskData.Mask = "##/##/####"
        mskHora.Mask = ""
        mskHora.Text = ""
        mskHora.Mask = "##:##"
        txtCodOcorr = ""
        lblDescOcorr.Caption = ""
        xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") 'verifica se tem Ocorrência
        cmbGravar.Enabled = True
        mskData.Enabled = True
        mskHora.Enabled = True
        txtCodOcorr.Enabled = True
        
        lblEntregueSN.ToolTipText = ""
        If xtemocorr = "0" Then
           lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
           lblEntregueSN.Caption = "OCORR/Baixado"
        ElseIf xtemocorr = "1" Then
           lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
           lblEntregueSN.Caption = "OK. ENTREGUE"
        ElseIf xtemocorr = "2" Then
           lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
           lblEntregueSN.Caption = "OCORR/Pendente"
        ElseIf xtemocorr = "N" Then
            If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= datahora("data") Then
                lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
                lblEntregueSN.Caption = "EM TRÂNSITO"
                lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
            Else
                lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
                lblEntregueSN.Caption = "SEM POSIÇÃO"
                lblEntregueSN.ToolTipText = "SEM POSIÇÃO = Após a Previsão de Entrega"
            End If
        ElseIf xtemocorr = "C" Then
            cmbGravar.Enabled = False
            mskData.Enabled = False
            mskHora.Enabled = False
            txtCodOcorr.Enabled = False
            lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
            lblEntregueSN.Caption = "CTC CANCELADO"
            lblEntregueSN.ToolTipText = "Cancelado em:" & de_informa.rsSel_Ctc_SAC.Fields("canc_data") & _
                                        "  Usuário:" & de_informa.rsSel_Ctc_SAC.Fields("canc_usu") & _
                                        "  Motivo:" & de_informa.rsSel_Ctc_SAC.Fields("canc_obs")
        End If

        'se tiver busca as ocorrências e atualiza o grid de ocorrências
        
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 transctc(TxtFilial, txtCtc), "01"
            Set GridOcorr.DataSource = de_informa
            GridOcorr.DataMember = "Sel_ConsOcorr2"
            GridOcorr.Refresh
            If de_informa.rsSel_ConsOcorr2.RecordCount = 0 Then
                cmdExclOcorr.Enabled = False
                chkObsOcorr.Enabled = False
            Else
                cmdExclOcorr.Enabled = True
                chkObsOcorr.Enabled = True
            End If

        'se houver baixa atualiza campos de baixa. ocorrência = 01
        
        If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
        de_informa.Sel_ConsOcorr transctc(TxtFilial, txtCtc), "01"
        If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr.Fields("baixadopre") = "S" Then
              'SE HOUVER PRÉ-BAIXA ATUALIZA CAMPOS DE PRÉ-BAIXA NO FORM
                lblDtPreBx.Caption = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
                lblHsPreBx.Caption = de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")
                txtRecPreBx.Text = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                cmdExclPreBx.Enabled = True
            Else
              'SE NÃO HOUVER PRÉ-BAIXA ATUALIZA CAMPOS COM BRANCOS ("")
                lblDtPreBx.Caption = ""
                lblHsPreBx.Caption = ""
                txtRecPreBx.Text = ""
                cmdExclPreBx.Enabled = False
            End If
            If de_informa.rsSel_ConsOcorr.Fields("baixadofinal") = "S" Then
               'SE HOUVER BAIXA FINAL ATUALIZA CAMPOS DE BAIXA FINAL NO FORM
                chkCanhoto.Enabled = False
                If de_informa.rsSel_ConsOcorr.Fields("canhotonf") = "S" Then
                    chkCanhoto.Value = 1
                Else
                    chkCanhoto.Value = 0
                End If
                lblDtBx.Caption = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                lblHsBx.Caption = de_informa.rsSel_ConsOcorr.Fields("hsbaixa")
                txtRecBx.Text = de_informa.rsSel_ConsOcorr.Fields("receb")
                cmdExclBx.Enabled = True
                'If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("canhotonf")) Then
                '    If de_informa.rsSel_ConsOcorr.Fields("canhotonf") = "S" Then
                '        chkCanhoto.Value = 1
                '    Else
                '        chkCanhoto.Value = 0
                '    End If
                'Else
                '    chkCanhoto.Value = 0
                'End If
            Else
                chkCanhoto.Enabled = True
               'SE NÃO HOUVER BAIXA FINAL ATUALIZA CAMPOS COM BRANCOS ("")
                lblDtBx.Caption = ""
                lblHsBx.Caption = ""
                txtRecBx.Text = ""
                chkCanhoto.Value = 0
                cmdExclBx.Enabled = False
            End If
            chkObsEntr.Enabled = True
        Else
                chkCanhoto.Enabled = True
                lblDtBx.Caption = ""
                lblHsBx.Caption = ""
                txtRecBx.Text = ""
                lblDtPreBx.Caption = ""
                lblHsPreBx.Caption = ""
                txtRecPreBx.Text = ""
                chkCanhoto.Value = 0
                chkObsEntr.Enabled = False
                cmdExclBx.Enabled = False
                cmdExclPreBx.Enabled = False
        End If
        mskData.BackColor = &HC0FFFF      'AMARELO
        mskHora.BackColor = &HC0FFFF      'AMARELO
        txtCodOcorr.BackColor = &HC0FFFF      'AMARELO
        mskData.Enabled = True
        mskHora.Enabled = True
        txtCodOcorr.Enabled = True
        mskData.SetFocus
        cmbGravar.Enabled = True
        If xtemocorr = "C" Then
            txtCtc.SetFocus
            cmbGravar.Enabled = False
            mskData.Enabled = False
            mskHora.Enabled = False
            txtCodOcorr.Enabled = False
        End If
End Sub

Private Sub cmdExclPreBx_Click()
    If Mid$(xdireitos, 23, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        'Exclusão do Registro de Baixa, pois se o CTC tiver baixa física também,
        'a mesma é excluida pois não é possivel CTC baixado Físico sem Pré-Baixa.
        'Não é Possível Excluir Pré-Baixa e deixar a baixa física !
        If Len(lblDtBx) > 0 Then  'tem baixa física
            If MsgBox("Confirma Exclusão dos Dados de BAIXA (Pré-Baixa e Baixa Física) ? ", vbYesNo, "Atenção") = vbYes Then
            
                de_informa.cn_informa.BeginTrans
            
                de_informa.excl_BaixaPOD transctc(TxtFilial, txtCtc)
                
                'exclui informação sobre canhoto
                de_informa.alt_ExclCanhotoNF transctc(TxtFilial, txtCtc)
                
                'se houver ocorrência, tem_ocorr = '2', caso contrário tem_ocorr = 'N'
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    de_informa.alt_temocorr_sn "2", transctc(TxtFilial, txtCtc)
                Else
                    de_informa.alt_temocorr_sn "N", transctc(TxtFilial, txtCtc)
                End If
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " PRÉ-BAIXA/FÍSICA"
                
                de_informa.cn_informa.CommitTrans
                
                cmdProcurar_Click
            End If
        Else  'é só Pré-Baixa
            If MsgBox("Confirma Exclusão dos Dados de PRÉ-BAIXA ? ", vbYesNo, "Atenção") = vbYes Then
            
                de_informa.cn_informa.BeginTrans
            
                de_informa.excl_BaixaPOD transctc(TxtFilial, txtCtc)
                'se houver ocorrência, tem_ocorr = '2', caso contrário tem_ocorr = 'N'
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    de_informa.alt_temocorr_sn "2", transctc(TxtFilial, txtCtc)
                Else
                    de_informa.alt_temocorr_sn "N", transctc(TxtFilial, txtCtc)
                End If
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(TxtFilial.Text, txtCtc.Text) & " PRÉ-BAIXA"
                
                de_informa.cn_informa.CommitTrans
                
                cmdProcurar_Click
            End If
        End If
    End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdSair2_Click()
    Set frmPod = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    If xultimofilial <> "" Then
        TxtFilial.Text = xultimofilial
        txtCtc.Text = xultimoctc
    End If
        TxtFilial.SetFocus















End Sub

Private Sub Form_Load()
    Dim X As Integer
        
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    
    'CONFIGURA OS OPTIONS, FRAMES E CHECKS
      optBaixaFinal.Enabled = False
      optPreBaixa.Enabled = False
      fraPreBaixa.Enabled = False
      fraBaixaFinal.Enabled = False
      GridOcorr.DataMember = ""
      GridOcorr.Refresh

      MSFlexGrid1.Cols = 10 ' Determinal Numero de Colunas do Flexgrid
      MSFlexGrid1.FixedCols = 0 'Determinal Quantas Colunas Fixas Vai Ter o Flexgrid
      MSFlexGrid1.FixedRows = 1 'Determinal Quantas Linhas Fixas Vai Ter o Flexgrid
      
      MSFlexGrid1.ColWidth(0) = 400
      MSFlexGrid1.ColWidth(1) = 1000
      MSFlexGrid1.ColWidth(2) = 1000
      MSFlexGrid1.ColWidth(3) = 2200
      MSFlexGrid1.ColWidth(4) = 2200
      MSFlexGrid1.ColWidth(5) = 1800
      MSFlexGrid1.ColWidth(6) = 400
      MSFlexGrid1.ColWidth(7) = 1100
      MSFlexGrid1.ColWidth(8) = 800
      MSFlexGrid1.ColWidth(9) = 10000
      
      MSFlexGrid1.TextMatrix(0, 0) = ""
      MSFlexGrid1.TextMatrix(0, 1) = "Filial-CTC"
      MSFlexGrid1.TextMatrix(0, 2) = "Data"
      MSFlexGrid1.TextMatrix(0, 3) = "Remetente"
      MSFlexGrid1.TextMatrix(0, 4) = "Destinatário"
      MSFlexGrid1.TextMatrix(0, 5) = "Cidade"
      MSFlexGrid1.TextMatrix(0, 6) = "UF"
      MSFlexGrid1.TextMatrix(0, 7) = "Valor"
      MSFlexGrid1.TextMatrix(0, 8) = "Peso"
      MSFlexGrid1.TextMatrix(0, 9) = "Notas Fiscais"
      
      MSFlexGrid1.ColAlignment(1) = 4
      MSFlexGrid1.ColAlignment(2) = 4
      MSFlexGrid1.ColAlignment(9) = 1

      

      
      
        
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.StatusBar1.Visible = True
    Set frmPod = Nothing
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskData_LostFocus()
    If mskData.Text <> "__/__/____" Then
        mskData.Text = century(mskData.Text)
        If Not IsDate(mskData.Text) Or Mid(mskData.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskData.SetFocus
            Exit Sub
        End If
        'tratamento acerto aws ---------------------------------------------
        If CDate(mskData.Text) < CDate(lblDtEmiss) Then
            MsgBox "Erro ! Data anterior à emissão.", vbCritical, "Erro"
            mskData.SetFocus
            Exit Sub
        End If
        '------------------------------------------------------------------
        If CDate(mskData.Text) > datahora("data") Then
            MsgBox "Erro ! Data posterior à hoje.", vbCritical, "Erro"
            mskData.SetFocus
            Exit Sub
        End If
    End If
End Sub


Private Sub mskDataMnf_GotFocus()
    mskDataMnf.SelStart = 0
    mskDataMnf.SelLength = 10

End Sub

Private Sub mskDataMnf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskDataMnf_LostFocus()
    If mskDataMnf.Text <> "__/__/____" Then
        mskDataMnf.Text = century(mskDataMnf.Text)
        If Not IsDate(mskDataMnf.Text) Or Mid(mskDataMnf.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskDataMnf.SetFocus
            Exit Sub
        End If
        'tratamento acerto aws ---------------------------------------------
        If CDate(mskDataMnf.Text) < CDate(lblDataMnf) Then
            MsgBox "Erro ! Data anterior à emissão do Manifesto.", vbCritical, "Erro"
            mskDataMnf.SetFocus
            Exit Sub
        End If
        '------------------------------------------------------------------
        If CDate(mskDataMnf.Text) > datahora("data") Then
            MsgBox "Erro ! Data posterior à hoje.", vbCritical, "Erro"
            mskDataMnf.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub mskHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskHora_LostFocus()
    If mskHora.Text <> "__:__" Then
        If Mid(mskHora.Text, 1, 2) > 23 Or Mid(mskHora.Text, 4, 2) > 59 Then
            MsgBox "Hora Inválida !", vbCritical, "Erro"
            mskHora.SetFocus
            Exit Sub
        End If
    Else
        mskHora.Text = "00:00"
    End If
End Sub

Private Sub optBaixaFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optCTC_Click()
    On Error Resume Next
    fraProcura.Caption = "Núm. da Filial e CTC"
    TxtFilial.Visible = True
    txtCtc.Visible = True
    txtNumNf.Visible = False
    TxtFilial.SetFocus
End Sub

Private Sub optNf_Click()
    On Error Resume Next
    fraProcura.Caption = "Núm. da NF"
    TxtFilial.Visible = False
    txtCtc.Visible = False
    txtNumNf.Visible = True
    txtNumNf.SetFocus
End Sub

Private Sub optPreBaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error Resume Next
    If SSTab1.Tab = 0 Then
        TxtFilial.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        txtFilialmnf.SetFocus
    End If
End Sub
Private Sub TxtCodOcorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCodOcorr)) = 2 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    ElseIf KeyAscii = 13 And Len(Trim$(txtCodOcorr)) = 0 Then   'TECLA ENTER
        KeyAscii = 0
        frmBuscaOcorrencias.Show 1
        If Len(Trim$(txtCodOcorr)) = 2 Then
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub
Private Sub txtCtc_Change()
    On Error Resume Next
    If Len(txtCtc.Text) >= 8 Then cmdProcurar.SetFocus
End Sub

Private Sub txtCTC_GotFocus()
   'RECEBER FOCO SELECIONADO
    txtCtc.SelStart = 0
    txtCtc.SelLength = 8
End Sub

Private Sub mskData_GotFocus()
    mskData.SelStart = 0
    mskData.SelLength = 10
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
End Sub

Private Sub TxtFilial_Change()
    On Error Resume Next
    If Len(TxtFilial.Text) >= 2 Then txtCtc.SetFocus
End Sub

Private Sub TxtFilial_GotFocus()
   'RECEBER FOCO SELECIONADO
    TxtFilial.SelStart = 0
    TxtFilial.SelLength = 2
End Sub

Private Sub mskHora_GotFocus()
    mskHora.SelStart = 0
    mskHora.SelLength = 5
End Sub

Private Sub optBaixaFinal_Click()
       fraPreBaixa.Enabled = True
       fraBaixaFinal.Enabled = False
       txtRecPreBx.BackColor = &HC0FFFF      'AMARELO
       txtRecBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = False
       'txtRecPreBx.SetFocus
       chkCanhoto.Value = 0


'    If lblDtPreBx.Caption = "" Then
'       fraPreBaixa.Enabled = True
'    Else
       fraPreBaixa.Enabled = False
'    End If
       fraBaixaFinal.Enabled = True
       txtRecBx.BackColor = &HC0FFFF      'amarelo
       txtRecPreBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = True
       'txtRecBx.SetFocus
       chkCanhoto.Value = 1
End Sub
Private Sub optPreBaixa_Click()
       fraPreBaixa.Enabled = True
       fraBaixaFinal.Enabled = False
       txtRecPreBx.BackColor = &HC0FFFF      'AMARELO
       txtRecBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = False
       'txtRecPreBx.SetFocus
       chkCanhoto.Value = 0
End Sub
Private Sub txtCodOcorr_Change()
    If txtCodOcorr.Text = "01" Then
        optBaixaFinal.Enabled = True
        optPreBaixa.Enabled = True
        If optPreBaixa.Value = True Then
            optPreBaixa_Click
        ElseIf optBaixaFinal.Value = True Then
            optBaixaFinal_Click
        End If
    Else
        optBaixaFinal.Enabled = False
        optPreBaixa.Enabled = False
        fraPreBaixa.Enabled = False
        fraBaixaFinal.Enabled = False
        txtRecPreBx.BackColor = &H8000000E       'BRANCO
        txtRecBx.BackColor = &H8000000E       'BRANCO
    End If
End Sub
Private Sub TxtCodOcorr_GotFocus()
    'RECEBER FOCO SELECIONADO
    txtCodOcorr.SelStart = 0
    txtCodOcorr.SelLength = 65000
End Sub
Private Sub txtCodOcorr_LostFocus()
    If txtCodOcorr.Text = "" Then
        Exit Sub
    Else
    'VERIFICA O CÓDIGO DE OCORRÊNCIA QUANDO DIGITADO E ATUALIZA A LABEL DE DESCRICAO DE OCORR
        If de_informa.rsSel_ConsCadOcor.State = 1 Then de_informa.rsSel_ConsCadOcor.Close
        de_informa.Sel_ConsCadOcor txtCodOcorr
        If de_informa.rsSel_ConsCadOcor.RecordCount > 0 Then
            lblDescOcorr.Caption = de_informa.rsSel_ConsCadOcor.Fields("descricao")
        Else
            MsgBox "Código de Ocorrência Inválido !", vbOKOnly + vbCritical, "Erro"
            txtCodOcorr.SetFocus
        End If
    End If
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

Private Sub txtFilialmnf_Change()
   On Error Resume Next
    If Len(txtFilialmnf.Text) >= 2 Then txtManifesto.SetFocus
End Sub

Private Sub txtFilialmnf_GotFocus()
   'RECEBER FOCO SELECIONADO
    txtFilialmnf.SelStart = 0
    txtFilialmnf.SelLength = 2
End Sub
Private Sub txtFilialmnf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtManifesto_GotFocus()
   'RECEBER FOCO SELECIONADO
    txtManifesto.SelStart = 0
    txtManifesto.SelLength = 6

End Sub

Private Sub txtManifesto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
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

Private Sub txtRecBx_Change()
    If optBaixaFinal.Value = True And lblDtPreBx = "" Then
        txtRecPreBx.Text = txtRecBx.Text
    End If
    If txtRecPreBx = "" And lblDtPreBx <> "" Then
        txtRecPreBx = "."
    End If
End Sub

Private Sub txtRecBx_GotFocus()
    txtRecBx.SelStart = 0
    txtRecBx.SelLength = 25
    chkCanhoto.Value = 1
End Sub

Private Sub txtRecBx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRecBx_LostFocus()
    txtRecBx.Text = UCase(txtRecBx.Text)
End Sub

Private Sub txtRecPreBx_GotFocus()
    txtRecPreBx.SelStart = 0
    txtRecPreBx.SelLength = 25
End Sub

Private Sub txtRecPreBx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtRecPreBx_LostFocus()
    txtRecPreBx.Text = UCase(txtRecPreBx.Text)
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 32 Then 'Caso Seja Precionado a Tecla Spaço ou Enter
      With MSFlexGrid1 'Seta o Flexgrid
          Call VerificaCheck(.Row, .Col, frmPod) 'Chama Função Para Marcar ou Desmarcar
      End With
  End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then 'Caso Seja dado um Click com o Botao do Mouse Sobre o Flexgrid
        With MSFlexGrid1 'Seta o Flexgrid
            If .Text <> "þ" And .Text <> "q" Then 'Se não for Check Box
              'If Text1.Text <> "" Then
              '  MSFlexGrid1.TextMatrix(Linha, Coluna) = Trim(Text1.Text) 'Diz que texto da Celula e igual ao text1
              'End If
              
              '.Col = .MouseCol 'Pega Coluna do Click do Mouse
              '.Row = .MouseRow 'Pega Linha do Click do Mouse
              'Linha = .MouseRow 'Pega Linha do Click do Mouse
              'Coluna = .MouseCol 'Pega Coluna do Click do Mouse
              'Text1.Width = .CellWidth 'Coloca o Tamanho do Text1 igual a celula celecionada
              'Text1.Height = .CellHeight 'Coloca o Tamanho do Text1 igual a celula celecionada
              'Text1.Move .CellLeft + .Left, .CellTop + .Top 'Posiciona o  text1 em cima da celula
              'Text1.Text = .Text 'Coloca o texto da Celula no text1
              'Text1.Visible = True 'deixa o text1 visivel
              
              'Coloca o Text1 na frente e posiciona o foco nele
              ' If Text1.Visible Then
              '  Text1.ZOrder
              '  Text1.SetFocus
              'End If
              
            End If
            If .MouseRow <> 0 Then 'Se nao For a Primeira Linha
                Call VerificaCheck(.MouseRow, .MouseCol, frmPod) 'Chama Função Para Marcar ou Desmarcar
            End If
        End With
    End If
End Sub

