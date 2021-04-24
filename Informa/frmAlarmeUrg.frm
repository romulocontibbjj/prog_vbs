VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAlarmeUrg 
   Caption         =   "Alarme Informa - PENDENTES (Sem Posição)"
   ClientHeight    =   7800
   ClientLeft      =   270
   ClientTop       =   840
   ClientWidth     =   11415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11415
   Begin VB.Timer tm_atualiza 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "URGÊNCIAS"
      TabPicture(0)   =   "frmAlarmeUrg.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraUrgencias"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdImprListUrg"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSairUrg"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdImprTelaUrg"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCopiarUrg"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "PRIORIDADES"
      TabPicture(1)   =   "frmAlarmeUrg.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "cmdImprListPri"
      Tab(1).Control(2)=   "cmdSairPri"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "fraPrioridades"
      Tab(1).Control(5)=   "cmdImprTelaPri"
      Tab(1).Control(6)=   "cmdCopiarPri"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "GERENCIAL"
      TabPicture(2)   =   "frmAlarmeUrg.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLegendaTotal"
      Tab(2).Control(1)=   "lblLegendaRodo"
      Tab(2).Control(2)=   "lblLegendaAereo"
      Tab(2).Control(3)=   "flexAereoGerTot"
      Tab(2).Control(4)=   "flexAereoGer"
      Tab(2).Control(5)=   "flexRodoGer"
      Tab(2).Control(6)=   "flexGeralGer"
      Tab(2).Control(7)=   "flexRodoGerTot"
      Tab(2).Control(8)=   "flexGeralGerTot"
      Tab(2).Control(9)=   "cmdImprGer1"
      Tab(2).Control(10)=   "cmdSairGer"
      Tab(2).Control(11)=   "comboMesAnoGeral"
      Tab(2).Control(12)=   "cmdProcessarGerGeral"
      Tab(2).ControlCount=   13
      Begin VB.CommandButton cmdProcessarGerGeral 
         Caption         =   "Processar"
         Height          =   375
         Left            =   -69120
         TabIndex        =   66
         Top             =   7080
         Width           =   1575
      End
      Begin VB.ComboBox comboMesAnoGeral 
         Height          =   315
         ItemData        =   "frmAlarmeUrg.frx":0054
         Left            =   -74640
         List            =   "frmAlarmeUrg.frx":0056
         TabIndex        =   65
         Text            =   "Mes/Ano"
         Top             =   7080
         Width           =   2295
      End
      Begin VB.CommandButton cmdSairGer 
         Caption         =   "S A I R"
         Height          =   375
         Left            =   -65520
         TabIndex        =   64
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprGer1 
         Caption         =   "Impr.Tela"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67320
         TabIndex        =   59
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton cmdCopiarPri 
         Caption         =   "Copiar..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70200
         TabIndex        =   58
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopiarUrg 
         Caption         =   "Copiar..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   57
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprTelaPri 
         Caption         =   "Imprimir Tela..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68880
         TabIndex        =   54
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprTelaUrg 
         Caption         =   "Imprimir Tela..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   6120
         TabIndex        =   53
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Frame fraPrioridades 
         Caption         =   "CTCs com Prioridade (Tratando Transit-Time)"
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
         Left            =   -74880
         TabIndex        =   49
         Top             =   4200
         Width           =   10935
         Begin MSDataGridLib.DataGrid gridPrioridades 
            Bindings        =   "frmAlarmeUrg.frx":0058
            Height          =   2295
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4048
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
            DataMember      =   "Sel_Urgencias"
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
            BeginProperty Column03 
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
            BeginProperty Column04 
               DataField       =   "prioridade"
               Caption         =   "prioridade"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            BeginProperty Column06 
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
            BeginProperty Column07 
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
            BeginProperty Column08 
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
            BeginProperty Column09 
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
            BeginProperty Column10 
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
            BeginProperty Column11 
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
            BeginProperty Column12 
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
               BeginProperty Column00 
                  ColumnWidth     =   1110,047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1154,835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   3135,118
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   3254,74
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   2594,835
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   420,095
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   780,095
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
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
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame11 
            Caption         =   "Observação de Emissão do CTC"
            Height          =   855
            Left            =   120
            TabIndex        =   40
            Top             =   2640
            Width           =   10695
            Begin VB.Label lblObsEmissPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Origem"
            Height          =   975
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   5295
            Begin VB.Label lblRemetPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   5055
            End
            Begin VB.Label lblCidadeOrigPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   38
               Top             =   600
               Width           =   3255
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Destino"
            Height          =   975
            Left            =   5520
            TabIndex        =   33
            Top             =   840
            Width           =   5295
            Begin VB.Label lblDestPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   5055
            End
            Begin VB.Label lblCidadeDestPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   35
               Top             =   600
               Width           =   3255
            End
            Begin VB.Label lblUfDestPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3600
               TabIndex        =   34
               Top             =   600
               Width           =   375
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Notas Fiscais"
            Height          =   615
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Width           =   10695
            Begin VB.Label lblNfsPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "STATUS"
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
            Left            =   8880
            TabIndex        =   29
            Top             =   120
            Width           =   1935
            Begin VB.Label lblStatusPri 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
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
               Left            =   945
               TabIndex        =   30
               Top             =   240
               Width           =   90
            End
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   4560
            TabIndex        =   56
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblDataEmiPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblHoraEmiPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5040
            TabIndex        =   47
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblFilialCtcPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   46
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "CTC: "
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   2280
            TabIndex        =   44
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   6240
            TabIndex        =   43
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblModalPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6840
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdSairPri 
         Caption         =   "SAIR"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -65280
         TabIndex        =   25
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprListPri 
         Caption         =   "Imprimir Listagem..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67080
         TabIndex        =   24
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CommandButton cmdSairUrg 
         Caption         =   "SAIR"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   23
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprListUrg 
         Caption         =   "Imprimir Listagem..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         TabIndex        =   22
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Frame fraUrgencias 
         Caption         =   "CTCs com Urgência (Não tratando Transit-Time)"
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
         Left            =   120
         TabIndex        =   20
         Top             =   4200
         Width           =   10935
         Begin MSDataGridLib.DataGrid gridUrgentes 
            Bindings        =   "frmAlarmeUrg.frx":0071
            Height          =   2295
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
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
            DataMember      =   "Sel_Urgencias"
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
            BeginProperty Column03 
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
            BeginProperty Column04 
               DataField       =   "prioridade"
               Caption         =   "prioridade"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            BeginProperty Column06 
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
            BeginProperty Column07 
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
            BeginProperty Column08 
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
            BeginProperty Column09 
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
            BeginProperty Column10 
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
            BeginProperty Column11 
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
            BeginProperty Column12 
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
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  ColumnWidth     =   1110,047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1154,835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   3135,118
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   3254,74
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   2594,835
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   420,095
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   780,095
               EndProperty
            EndProperty
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
         TabIndex        =   1
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame13 
            Caption         =   "STATUS"
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
            Left            =   8880
            TabIndex        =   26
            Top             =   120
            Width           =   1935
            Begin VB.Label lblStatusUrg 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
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
               Left            =   945
               TabIndex        =   27
               Top             =   240
               Width           =   90
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Notas Fiscais"
            Height          =   615
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   10695
            Begin VB.Label lblNfsUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Destino"
            Height          =   975
            Left            =   5520
            TabIndex        =   7
            Top             =   840
            Width           =   5295
            Begin VB.Label lblUfDestUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3600
               TabIndex        =   10
               Top             =   600
               Width           =   375
            End
            Begin VB.Label lblCidadeDestUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   9
               Top             =   600
               Width           =   3255
            End
            Begin VB.Label lblDestUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   5055
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Origem"
            Height          =   975
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   5295
            Begin VB.Label lblCidadeOrigUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   6
               Top             =   600
               Width           =   3255
            End
            Begin VB.Label lblRemetUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   5055
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Observação de Emissão do CTC"
            Height          =   855
            Left            =   120
            TabIndex        =   2
            Top             =   2640
            Width           =   10695
            Begin VB.Label lblObsEmissUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   120
               TabIndex        =   3
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   4560
            TabIndex        =   55
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblModalUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6840
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   6240
            TabIndex        =   18
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   2280
            TabIndex        =   17
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "CTC: "
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblFilialctcUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblHoraEmiUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5040
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblDataEmiUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGeralGerTot 
         Height          =   375
         Left            =   -74640
         TabIndex        =   60
         Top             =   6480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexRodoGerTot 
         Height          =   375
         Left            =   -74640
         TabIndex        =   61
         Top             =   4320
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGeralGer 
         Height          =   1575
         Left            =   -74640
         TabIndex        =   62
         Top             =   4800
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2778
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexRodoGer 
         Height          =   1575
         Left            =   -74640
         TabIndex        =   63
         Top             =   2640
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2778
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexAereoGer 
         Height          =   1575
         Left            =   -74640
         TabIndex        =   67
         Top             =   480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2778
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexAereoGerTot 
         Height          =   375
         Left            =   -74640
         TabIndex        =   71
         Top             =   2160
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblLegendaAereo 
         Caption         =   "A É R E O"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   70
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLegendaRodo 
         Caption         =   "R O D O"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74880
         TabIndex        =   69
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblLegendaTotal 
         Caption         =   "T O T A L"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   68
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "PRIORIDADES PENDENTES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   -74640
         TabIndex        =   52
         Top             =   6960
         Width           =   4020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "URGÊNCIAS PENDENTES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   480
         TabIndex        =   51
         Top             =   6960
         Width           =   3750
      End
   End
End
Attribute VB_Name = "frmAlarmeUrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub cmdCopiarPri_Click()
    Dim atcopiar As CClipboard, xlinha As String
    
    xlinha = "CTC: " & gridPrioridades.Columns(0) & "  Data: " & gridPrioridades.Columns(1) & "  Modal: " & gridPrioridades.Columns(3) & "  " & gridPrioridades.Columns(4) & Chr(13) & Chr(10) & _
             "Remetente: " & gridPrioridades.Columns(5) & Chr(13) & Chr(10) & _
             "Cidade - UF: " & Trim$(gridPrioridades.Columns(8)) & " - " & gridPrioridades.Columns(9) & Chr(13) & Chr(10) & _
             "Obs: " & Trim$(gridPrioridades.Columns(11)) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Set atcopiar = New CClipboard
    
    atcopiar.Clear
    atcopiar.SetText xlinha
    
    'MsgBox xlinha

End Sub

Private Sub cmdCopiarUrg_Click()
    Dim atcopiar As CClipboard, xlinha As String
    
    xlinha = "CTC: " & gridUrgentes.Columns(0) & "  Data: " & gridUrgentes.Columns(1) & "  Modal: " & gridUrgentes.Columns(3) & "  " & gridUrgentes.Columns(4) & Chr(13) & Chr(10) & _
             "Remetente: " & gridUrgentes.Columns(5) & Chr(13) & Chr(10) & _
             "Cidade - UF: " & Trim$(gridUrgentes.Columns(8)) & " - " & gridUrgentes.Columns(9) & Chr(13) & Chr(10) & _
             "Obs: " & Trim$(gridUrgentes.Columns(11)) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Set atcopiar = New CClipboard
    
    atcopiar.Clear
    atcopiar.SetText xlinha
    
    'MsgBox xlinha
End Sub

Private Sub cmdImprGer1_Click()
    
    Me.MousePointer = 11
    comboMesAnoGeral.Enabled = False
    cmdProcessarGerGeral.Enabled = False
    cmdDetalhe.Enabled = False
    cmdSairGer.Enabled = False
    SSTab1.Enabled = False
    flexAereoGer.Enabled = False
    flexRodoGer.Enabled = False
    flexGeralGer.Enabled = False
    cmdImprGer1.Enabled = False
    DoEvents
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    
    Me.MousePointer = 0
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdSairGer.Enabled = True
    SSTab1.Enabled = True
    flexAereoGer.Enabled = True
    flexRodoGer.Enabled = True
    flexGeralGer.Enabled = True
    DoEvents
    lblMensagem.Caption = comboMesAnoGeral.Text
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdImprGer1.Enabled = True
    If comboMesAnoGeral.ListIndex = 0 Then
        cmdDetalhe.Enabled = True
    Else
        cmdDetalhe.Enabled = False
    End If

End Sub
Private Sub cmdImprListPri_Click()

Dim xColuna As Single, xlinha As Single, xremetente As String, xdestinatario As String, xcidadeuf As String
    
cmdImprTelaUrg.Enabled = False
cmdImprListUrg.Enabled = False
cmdSairUrg.Enabled = False
cmdImprTelaPri.Enabled = False
cmdImprListPri.Enabled = False
cmdSairPri.Enabled = False
    
If Printer.Orientation = vbPRORLandscape Then Printer.Orientation = vbPRORPortrait
    
If de_informa.rsSel_Prioridades.RecordCount > 0 Then
    
    de_informa.rsSel_Prioridades.MoveFirst
    xColuna = 1
    xlinha = 0
    Do Until de_informa.rsSel_Prioridades.EOF
        If xlinha = 0 And xColuna = 1 Then   'identifica inicio da página/cabeçário
            Printer.FontName = "Courier New"
            Printer.Print
            Printer.Print
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontUnderline = True
            Printer.Print Spc(5); "INTEC TRANSPORTES"
            Printer.FontUnderline = False
            Printer.Print
            Printer.Print Spc(5); "RELATÓRIO DE PRIORIDADES PENDENTES (Sem Posição)"
            Printer.Print Spc(5); "USUÁRIO: " & xusuario
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.Print Spc(9); "Filial-CTC"; Spc(5); "Data"; Spc(7); "Modal"; Spc(5); "Remetente"; Spc(14); "Destinatário"; Spc(11); "Cidade/UF"; Spc(10); "Status"
            Printer.FontSize = 12
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.FontBold = False
            Printer.FontUnderline = False
        End If
        
        xremetente = Trim$(Mid$(de_informa.rsSel_Prioridades.Fields("remet_nome"), 1, 20))
        If Len(xremetente) < 20 Then
            xremetente = xremetente & String(20 - Len(xremetente), " ")
        End If
        xdestinatario = Trim$(Mid$(de_informa.rsSel_Prioridades.Fields("dest_nome"), 1, 20))
        If Len(xdestinatario) < 20 Then
            xdestinatario = xdestinatario & String(20 - Len(xdestinatario), " ")
        End If
        xcidadeuf = Trim$(Mid$(de_informa.rsSel_Prioridades.Fields("Cidade_dest"), 1, 15)) & "-" & de_informa.rsSel_Prioridades.Fields("uf_dest")
        If Len(xcidadeuf) < 18 Then
            xcidadeuf = xcidadeuf & String(18 - Len(xcidadeuf), " ")
        End If

        Printer.Print Spc(9); de_informa.rsSel_Prioridades.Fields("filialctc"); Spc(2); _
                                zeros(Day(de_informa.rsSel_Prioridades.Fields("data")), 2) & "/" & _
                                zeros(Month(de_informa.rsSel_Prioridades.Fields("data")), 2) & "/" & _
                                zeros(Year(de_informa.rsSel_Prioridades.Fields("data")), 4); Spc(2); _
                                Trim$(de_informa.rsSel_Prioridades.Fields("modal")); Spc(2 + (10 - Len(Trim$(de_informa.rsSel_Prioridades.Fields("modal"))))); _
                                xremetente; Spc(3); xdestinatario; Spc(3); xcidadeuf; Spc(3); _
                                de_informa.rsSel_Prioridades.Fields("tem_ocorr")
        xlinha = xlinha + 1
        If xlinha = 70 Then
            xlinha = 0
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontBold = False
            Printer.FontStrikethru = False
            Printer.Print
            Printer.NewPage
        End If
        de_informa.rsSel_Prioridades.MoveNext
    Loop
    
    de_informa.rsSel_Prioridades.MoveFirst
            
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontStrikethru = True
    Printer.Print Spc(5); String(132, " ")
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.FontStrikethru = False
    Printer.Print
    Printer.NewPage
    Printer.EndDoc   'finaliza spool da impressão
    DoEvents
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "IMPRESSÃO", xusuario, "RELATÓRIO DE PRIORIDADES"
        
    MsgBox "RELATÓRIO ENVIADO PARA IMPRESSÃO !"
    
Else
    MsgBox "Não Há Dados a Serem Impressos !"
End If

cmdImprTelaUrg.Enabled = True
cmdImprListUrg.Enabled = True
cmdSairUrg.Enabled = True
cmdImprTelaPri.Enabled = True
cmdImprListPri.Enabled = True
cmdSairPri.Enabled = True

End Sub

Private Sub cmdImprListUrg_Click()

Dim xColuna As Single, xlinha As Single, xremetente As String, xdestinatario As String, xcidadeuf As String

cmdImprTelaUrg.Enabled = False
cmdImprListUrg.Enabled = False
cmdSairUrg.Enabled = False
cmdImprTelaPri.Enabled = False
cmdImprListPri.Enabled = False
cmdSairPri.Enabled = False
    
If Printer.Orientation = vbPRORLandscape Then Printer.Orientation = vbPRORPortrait

If de_informa.rsSel_Urgencias.RecordCount > 0 Then
    de_informa.rsSel_Urgencias.MoveFirst
    xColuna = 1
    xlinha = 0
    Do Until de_informa.rsSel_Urgencias.EOF
        If xlinha = 0 And xColuna = 1 Then   'identifica inicio da página/cabeçário
            Printer.FontName = "Courier New"
            Printer.Print
            Printer.Print
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontUnderline = True
            Printer.Print Spc(5); "INTEC TRANSPORTES"
            Printer.FontUnderline = False
            Printer.Print
            Printer.Print Spc(5); "RELATÓRIO DE URGÊNCIAS PENDENTES (Sem Posição)"
            Printer.Print Spc(5); "USUÁRIO: " & xusuario
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.Print Spc(9); "Filial-CTC"; Spc(5); "Data"; Spc(7); "Modal"; Spc(5); "Remetente"; Spc(14); "Destinatário"; Spc(11); "Cidade/UF"; Spc(10); "Status"
            Printer.FontSize = 12
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.FontBold = False
            Printer.FontUnderline = False
        End If
        
        xremetente = Trim$(Mid$(de_informa.rsSel_Urgencias.Fields("remet_nome"), 1, 20))
        If Len(xremetente) < 20 Then
            xremetente = xremetente & String(20 - Len(xremetente), " ")
        End If
        xdestinatario = Trim$(Mid$(de_informa.rsSel_Urgencias.Fields("dest_nome"), 1, 20))
        If Len(xdestinatario) < 20 Then
            xdestinatario = xdestinatario & String(20 - Len(xdestinatario), " ")
        End If
        xcidadeuf = Trim$(Mid$(de_informa.rsSel_Urgencias.Fields("Cidade_dest"), 1, 15)) & "-" & de_informa.rsSel_Urgencias.Fields("uf_dest")
        If Len(xcidadeuf) < 18 Then
            xcidadeuf = xcidadeuf & String(18 - Len(xcidadeuf), " ")
        End If

        Printer.Print Spc(9); de_informa.rsSel_Urgencias.Fields("filialctc"); Spc(2); _
                                zeros(Day(de_informa.rsSel_Urgencias.Fields("data")), 2) & "/" & _
                                zeros(Month(de_informa.rsSel_Urgencias.Fields("data")), 2) & "/" & _
                                zeros(Year(de_informa.rsSel_Urgencias.Fields("data")), 4); Spc(2); _
                                Trim$(de_informa.rsSel_Urgencias.Fields("modal")); Spc(2 + (10 - Len(Trim$(de_informa.rsSel_Urgencias.Fields("modal"))))); _
                                xremetente; Spc(3); xdestinatario; Spc(3); xcidadeuf; Spc(3); _
                                de_informa.rsSel_Urgencias.Fields("tem_ocorr")
        xlinha = xlinha + 1
        If xlinha = 70 Then
            xlinha = 0
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontBold = False
            Printer.FontStrikethru = False
            Printer.Print
            Printer.NewPage
        End If
        de_informa.rsSel_Urgencias.MoveNext
    Loop
    
    de_informa.rsSel_Urgencias.MoveFirst
            
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontStrikethru = True
    Printer.Print Spc(5); String(132, " ")
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.FontStrikethru = False
    Printer.Print
    Printer.NewPage
    Printer.EndDoc   'finaliza spool da impressão
    DoEvents
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "IMPRESSÃO", xusuario, "RELATÓRIO DE URGÊNCIA"
        
    MsgBox "RELATÓRIO ENVIADO PARA IMPRESSÃO !"
Else
    MsgBox "Não Há Dados a Serem Impressos !"
End If
    
cmdImprTelaUrg.Enabled = True
cmdImprListUrg.Enabled = True
cmdSairUrg.Enabled = True
cmdImprTelaPri.Enabled = True
cmdImprListPri.Enabled = True
cmdSairPri.Enabled = True

End Sub

Private Sub cmdImprTelaPri_Click()
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
End Sub

Private Sub cmdImprTelaUrg_Click()
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
    
End Sub
Private Sub cmdProcessarGerGeral_Click()
    Dim xTotValmerc As Currency, xTotFrete As Currency, xTotPeso As Currency
    Dim xTotVol As Long, xTotCtc As Long, xTotNf As Long, xdataper1 As Date, xdataper2 As Date
    Dim xTotValMercMesAnt As Currency, xTotFreteMesAnt As Currency, xTotPesoMesAnt As Currency
    Dim xTotVolMesAnt As Long, xTotCtcMesAnt As Long, xTotNfMesAnt As Long
    
    Me.MousePointer = 11
    comboMesAnoGeral.Enabled = False
    cmdProcessarGerGeral.Enabled = False
    cmdSairGer.Enabled = False
    SSTab1.Enabled = False
    flexAereoGer.Enabled = False
    flexRodoGer.Enabled = False
    flexGeralGer.Enabled = False
    cmdImprGer1.Enabled = False
    
    flexAereoGer.Rows = 1
    flexAereoGer.Rows = 2
    flexAereoGer.FixedRows = 1
    flexRodoGer.Rows = 1
    flexRodoGer.Rows = 2
    flexRodoGer.FixedRows = 1
    flexGeralGer.Rows = 1
    flexGeralGer.Rows = 2
    flexGeralGer.FixedRows = 1
    flexAereoGerTot.Rows = 0
    flexAereoGerTot.Rows = 1
    flexRodoGerTot.Rows = 0
    flexRodoGerTot.Rows = 1
    flexGeralGerTot.Rows = 0
    flexGeralGerTot.Rows = 1
    
    DoEvents
    
'MONTANDO O AÉREO
    
    xdataper1 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4) & "/" & _
                Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2) & "/" & "01")

    If comboMesAnoGeral.ListIndex = 0 Then  'se for o mês atual, pega somente até a data de hoje, pelo dia
        xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4) & "/" & _
                    Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2) & "/" & Day(datahora("DATA")))
    Else 'senão, pega até o último dia do mês
        xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4) & "/" & _
                    Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2) & "/" & UltDiaMes(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2), Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4)))
    End If
    
    If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
    de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "AEREO"
        
    If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
        flexAereoGer.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
        xTotValmerc = 0
        xTotFrete = 0
        xTotPeso = 0
        xTotVol = 0
        xTotCtc = 0
        xTotNf = 0
        For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
            flexAereoGer.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexAereoGer.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
            flexAereoGer.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
            flexAereoGer.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
            flexAereoGer.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
            flexAereoGer.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
            flexAereoGer.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
            flexAereoGer.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
            If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
            'busca qtde de nfs de cada filial
            de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "AEREO", de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexAereoGer.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
            
            xTotValmerc = xTotValmerc + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
            xTotFrete = xTotFrete + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
            xTotPeso = xTotPeso + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
            xTotVol = xTotVol + de_informa.rsSel_AlarmMovGer.Fields("tvol")
            xTotCtc = xTotCtc + de_informa.rsSel_AlarmMovGer.Fields("qtd")
            xTotNf = xTotNf + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
            
            de_informa.rsSel_AlarmMovGer.MoveNext
            DoEvents
        Next
        
        flexAereoGerTot.TextMatrix(0, 2) = "SUB-TOTAL ............"
        flexAereoGerTot.TextMatrix(0, 3) = Format(xTotValmerc, "###,###,###,##0.00")
        flexAereoGerTot.TextMatrix(0, 4) = Format(xTotFrete, "###,###,##0.00")
        flexAereoGerTot.TextMatrix(0, 5) = Format(xTotFrete / xTotValmerc, "##0.000%")
        flexAereoGerTot.TextMatrix(0, 6) = Format(xTotPeso, "###,###,##0.0")
        flexAereoGerTot.TextMatrix(0, 7) = Format(xTotVol, "###,###,##0")
        flexAereoGerTot.TextMatrix(0, 8) = Format(xTotCtc, "###,###,##0")
        flexAereoGerTot.TextMatrix(0, 9) = Format(xTotNf, "###,###,##0")
        
        DoEvents
    End If

'MONTANDO O RODO
    
    If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
    de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "RODOVIARIO"
    
    
    
    If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
        flexRodoGer.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
        xTotValmerc = 0
        xTotFrete = 0
        xTotPeso = 0
        xTotVol = 0
        xTotCtc = 0
        xTotNf = 0
        For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
            flexRodoGer.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexRodoGer.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
            flexRodoGer.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
            flexRodoGer.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
            flexRodoGer.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
            flexRodoGer.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
            flexRodoGer.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
            flexRodoGer.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
            If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
            'busca qtde de nfs de cada filial
            de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "RODOVIARIO", de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexRodoGer.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
            xTotValmerc = xTotValmerc + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
            xTotFrete = xTotFrete + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
            xTotPeso = xTotPeso + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
            xTotVol = xTotVol + de_informa.rsSel_AlarmMovGer.Fields("tvol")
            xTotCtc = xTotCtc + de_informa.rsSel_AlarmMovGer.Fields("qtd")
            xTotNf = xTotNf + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
            
            de_informa.rsSel_AlarmMovGer.MoveNext
            DoEvents
        Next
        
        flexRodoGerTot.TextMatrix(0, 2) = "SUB-TOTAL ............"
        flexRodoGerTot.TextMatrix(0, 3) = Format(xTotValmerc, "###,###,###,##0.00")
        flexRodoGerTot.TextMatrix(0, 4) = Format(xTotFrete, "###,###,##0.00")
        flexRodoGerTot.TextMatrix(0, 5) = Format(xTotFrete / xTotValmerc, "##0.000%")
        flexRodoGerTot.TextMatrix(0, 6) = Format(xTotPeso, "###,###,##0.0")
        flexRodoGerTot.TextMatrix(0, 7) = Format(xTotVol, "###,###,##0")
        flexRodoGerTot.TextMatrix(0, 8) = Format(xTotCtc, "###,###,##0")
        flexRodoGerTot.TextMatrix(0, 9) = Format(xTotNf, "###,###,##0")
        DoEvents
    End If

'MONTANDO O GERAL (RODO + AEREO)
    
    xTotValmerc = 0
    xTotFrete = 0
    xTotPeso = 0
    xTotVol = 0
    xTotCtc = 0
    xTotNf = 0
    
    If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
    de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "%"
    
    If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
        flexGeralGer.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
        For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
            flexGeralGer.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexGeralGer.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
            flexGeralGer.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
            flexGeralGer.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
            flexGeralGer.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
            flexGeralGer.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
            flexGeralGer.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
            flexGeralGer.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
            If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
            'busca qtde de nfs de cada filial
            de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "%", de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexGeralGer.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
            
            xTotValmerc = xTotValmerc + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
            xTotFrete = xTotFrete + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
            xTotPeso = xTotPeso + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
            xTotVol = xTotVol + de_informa.rsSel_AlarmMovGer.Fields("tvol")
            xTotCtc = xTotCtc + de_informa.rsSel_AlarmMovGer.Fields("qtd")
            xTotNf = xTotNf + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
            
            de_informa.rsSel_AlarmMovGer.MoveNext
            DoEvents
        Next

        flexGeralGerTot.TextMatrix(0, 2) = "TOTAL " & comboMesAnoGeral.Text & ".............."
        flexGeralGerTot.TextMatrix(0, 3) = Format(xTotValmerc, "###,###,###,##0.00")
        flexGeralGerTot.TextMatrix(0, 4) = Format(xTotFrete, "###,###,##0.00")
        flexGeralGerTot.TextMatrix(0, 5) = Format(xTotFrete / xTotValmerc, "##0.000%")
        flexGeralGerTot.TextMatrix(0, 6) = Format(xTotPeso, "###,###,##0.0")
        flexGeralGerTot.TextMatrix(0, 7) = Format(xTotVol, "###,###,##0")
        flexGeralGerTot.TextMatrix(0, 8) = Format(xTotCtc, "###,###,##0")
        flexGeralGerTot.TextMatrix(0, 9) = Format(xTotNf, "###,###,##0")
        DoEvents
    End If

    If comboMesAnoGeral.ListIndex = 0 Then

        'MONTANDO O GERAL (RODO + AEREO) / DO MÊS ANTERIOR
        
        xdataper1 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                    Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & "01")
                    
        If IsDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                  Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & _
                  Day(datahora("DATA"))) Then
            xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                        Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & Day(datahora("DATA")))
                  
        Else
            xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                        Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & _
                        UltDiaMes(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2), _
                                  Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4)))
        End If
        
        'tratar se este mes já estiver no dia 31 e no mês anterior não tiver esta data
        
        If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
        de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "%"
        
        If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
'            flexGeralGerMesAnt.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
            xTotValMercMesAnt = 0
            xTotFreteMesAnt = 0
            xTotPesoMesAnt = 0
            xTotVolMesAnt = 0
            xTotCtcMesAnt = 0
            xTotNfMesAnt = 0
            For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
            '    flexGeralGerMesAnt.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
            '    flexGeralGerMesAnt.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
            '    flexGeralGerMesAnt.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
            '    flexGeralGerMesAnt.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
            '    flexGeralGerMesAnt.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
            '    flexGeralGerMesAnt.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
            '    flexGeralGerMesAnt.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
            '    flexGeralGerMesAnt.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
            '    If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
                'busca qtde de nfs de cada filial
             '   de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "%", de_informa.rsSel_AlarmMovGer.Fields("filial")
             '   flexGeralGerMesAnt.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
                
                xTotValMercMesAnt = xTotValMercMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
                xTotFreteMesAnt = xTotFreteMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
                xTotPesoMesAnt = xTotPesoMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
                xTotVolMesAnt = xTotVolMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tvol")
                xTotCtcMesAnt = xTotCtcMesAnt + de_informa.rsSel_AlarmMovGer.Fields("qtd")
                xTotNfMesAnt = xTotNfMesAnt + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
                
                de_informa.rsSel_AlarmMovGer.MoveNext
                DoEvents
                DoEvents
            Next
            
         '   flexGeralGerTotMesAnt.TextMatrix(0, 2) = "TOTAL " & comboMesAnoGeral.List(1) & ".............."
         '   flexGeralGerTotMesAnt.TextMatrix(0, 3) = Format(xTotValMercMesAnt, "###,###,###,##0.00")
         '   flexGeralGerTotMesAnt.TextMatrix(0, 4) = Format(xTotFreteMesAnt, "###,###,##0.00")
         '   flexGeralGerTotMesAnt.TextMatrix(0, 5) = Format(xTotFreteMesAnt / xTotValMercMesAnt, "##0.000%")
         '   flexGeralGerTotMesAnt.TextMatrix(0, 6) = Format(xTotPesoMesAnt, "###,###,##0.0")
         '   flexGeralGerTotMesAnt.TextMatrix(0, 7) = Format(xTotVolMesAnt, "###,###,##0")
         '   flexGeralGerTotMesAnt.TextMatrix(0, 8) = Format(xTotCtcMesAnt, "###,###,##0")
         '   flexGeralGerTotMesAnt.TextMatrix(0, 9) = Format(xTotNfMesAnt, "###,###,##0")
            DoEvents
        End If
        
     '   flexVarPercent.TextMatrix(0, 2) = "VARIAÇÃO (%).........."
     '   flexVarPercent.TextMatrix(0, 3) = Format((xTotValmerc - xTotValMercMesAnt) / xTotValMercMesAnt, "##0.00%")
     '   flexVarPercent.TextMatrix(0, 4) = Format((xTotFrete - xTotFreteMesAnt) / xTotFreteMesAnt, "##0.00%")
     '   flexVarPercent.TextMatrix(0, 6) = Format((xTotPeso - xTotPesoMesAnt) / xTotPesoMesAnt, "##0.00%")
     '   flexVarPercent.TextMatrix(0, 7) = Format((xTotVol - xTotVolMesAnt) / xTotVolMesAnt, "##0.00%")
     '   flexVarPercent.TextMatrix(0, 8) = Format((xTotCtc - xTotCtcMesAnt) / xTotCtcMesAnt, "##0.00%")
     '   flexVarPercent.TextMatrix(0, 9) = Format((xTotNf - xTotNfMesAnt) / xTotNfMesAnt, "##0.00%")
        
   '     lblTexto.Caption = "Comparação do Mês Atual (" & comboMesAnoGeral.List(0) & ") Com o Mês Anterior (" & comboMesAnoGeral.List(1) & ")  no Mesmo Período."
   '     cmdDetalhe.Enabled = True
        
    Else
    End If
    
    Me.MousePointer = 0
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdSairGer.Enabled = True
    SSTab1.Enabled = True
    flexAereoGer.Enabled = True
    flexRodoGer.Enabled = True
    flexGeralGer.Enabled = True
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdImprGer1.Enabled = True

End Sub

Private Sub cmdSairGer_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
    
End Sub

Private Sub cmdSairPri_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
End Sub

Private Sub cmdSairUrg_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command9_Click()

End Sub

Private Sub DataGrid3_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub flexUrgRegiao_Click()

End Sub

Private Sub comboMesAnoCliente_Change()

End Sub

Private Sub FlexPerfAir_Click()

End Sub

Private Sub FlexPerfRodo_Click()

End Sub

Private Sub FlexPorModal_Click()

End Sub

Private Sub Form_Load()
    Dim PosWin As Long
    PosWin = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    'Configura as Flex
    
    'gerencial - movto geral - aereo
    flexAereoGer.Cols = 10
    flexAereoGer.ColWidth(0) = 150
    flexAereoGer.ColWidth(1) = 400
    flexAereoGer.ColWidth(2) = 1600
    flexAereoGer.ColWidth(3) = 1700
    flexAereoGer.ColWidth(4) = 1400
    flexAereoGer.ColWidth(5) = 970
    flexAereoGer.ColWidth(6) = 1200
    flexAereoGer.ColWidth(7) = 900
    flexAereoGer.ColWidth(8) = 900
    flexAereoGer.ColWidth(9) = 900
    
    flexAereoGer.TextMatrix(0, 1) = "Filial"
    flexAereoGer.TextMatrix(0, 2) = "Nome Filial"
    flexAereoGer.TextMatrix(0, 3) = "Vlr. Mercadoria"
    flexAereoGer.TextMatrix(0, 4) = "Frete Líquido"
    flexAereoGer.TextMatrix(0, 5) = "Frete/Valor"
    flexAereoGer.TextMatrix(0, 6) = "Peso"
    flexAereoGer.TextMatrix(0, 7) = "Volumes"
    flexAereoGer.TextMatrix(0, 8) = "CTCs"
    flexAereoGer.TextMatrix(0, 9) = "NFs"
    
    flexAereoGerTot.Cols = 10
    flexAereoGerTot.ColWidth(0) = 150
    flexAereoGerTot.ColWidth(1) = 400
    flexAereoGerTot.ColWidth(2) = 1600
    flexAereoGerTot.ColWidth(3) = 1700
    flexAereoGerTot.ColWidth(4) = 1400
    flexAereoGerTot.ColWidth(5) = 970
    flexAereoGerTot.ColWidth(6) = 1200
    flexAereoGerTot.ColWidth(7) = 900
    flexAereoGerTot.ColWidth(8) = 900
    flexAereoGerTot.ColWidth(9) = 900

    'gerencial - movto geral - rodo
    flexRodoGer.Cols = 10
    flexRodoGer.ColWidth(0) = 150
    flexRodoGer.ColWidth(1) = 400
    flexRodoGer.ColWidth(2) = 1600
    flexRodoGer.ColWidth(3) = 1700
    flexRodoGer.ColWidth(4) = 1400
    flexRodoGer.ColWidth(5) = 970
    flexRodoGer.ColWidth(6) = 1200
    flexRodoGer.ColWidth(7) = 900
    flexRodoGer.ColWidth(8) = 900
    flexRodoGer.ColWidth(9) = 900
    
    flexRodoGer.TextMatrix(0, 1) = "Filial"
    flexRodoGer.TextMatrix(0, 2) = "Nome Filial"
    flexRodoGer.TextMatrix(0, 3) = "Vlr. Mercadoria"
    flexRodoGer.TextMatrix(0, 4) = "Frete Líquido"
    flexRodoGer.TextMatrix(0, 5) = "Frete/Valor"
    flexRodoGer.TextMatrix(0, 6) = "Peso"
    flexRodoGer.TextMatrix(0, 7) = "Volumes"
    flexRodoGer.TextMatrix(0, 8) = "CTCs"
    flexRodoGer.TextMatrix(0, 9) = "NFs"

    flexRodoGerTot.Cols = 10
    flexRodoGerTot.ColWidth(0) = 150
    flexRodoGerTot.ColWidth(1) = 400
    flexRodoGerTot.ColWidth(2) = 1600
    flexRodoGerTot.ColWidth(3) = 1700
    flexRodoGerTot.ColWidth(4) = 1400
    flexRodoGerTot.ColWidth(5) = 970
    flexRodoGerTot.ColWidth(6) = 1200
    flexRodoGerTot.ColWidth(7) = 900
    flexRodoGerTot.ColWidth(8) = 900
    flexRodoGerTot.ColWidth(9) = 900

    'gerencial - movto geral - rodo
    flexGeralGer.Cols = 10
    flexGeralGer.ColWidth(0) = 150
    flexGeralGer.ColWidth(1) = 400
    flexGeralGer.ColWidth(2) = 1600
    flexGeralGer.ColWidth(3) = 1700
    flexGeralGer.ColWidth(4) = 1400
    flexGeralGer.ColWidth(5) = 970
    flexGeralGer.ColWidth(6) = 1200
    flexGeralGer.ColWidth(7) = 900
    flexGeralGer.ColWidth(8) = 900
    flexGeralGer.ColWidth(9) = 900
    
    flexGeralGer.TextMatrix(0, 1) = "Filial"
    flexGeralGer.TextMatrix(0, 2) = "Nome Filial"
    flexGeralGer.TextMatrix(0, 3) = "Vlr. Mercadoria"
    flexGeralGer.TextMatrix(0, 4) = "Frete Líquido"
    flexGeralGer.TextMatrix(0, 5) = "Frete/Valor"
    flexGeralGer.TextMatrix(0, 6) = "Peso"
    flexGeralGer.TextMatrix(0, 7) = "Volumes"
    flexGeralGer.TextMatrix(0, 8) = "CTCs"
    flexGeralGer.TextMatrix(0, 9) = "NFs"
    
    flexGeralGerTot.Cols = 10
    flexGeralGerTot.ColWidth(0) = 150
    flexGeralGerTot.ColWidth(1) = 400
    flexGeralGerTot.ColWidth(2) = 1600
    flexGeralGerTot.ColWidth(3) = 1700
    flexGeralGerTot.ColWidth(4) = 1400
    flexGeralGerTot.ColWidth(5) = 970
    flexGeralGerTot.ColWidth(6) = 1200
    flexGeralGerTot.ColWidth(7) = 900
    flexGeralGerTot.ColWidth(8) = 900
    flexGeralGerTot.ColWidth(9) = 900
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAlarmeUrg = Nothing
End Sub

Private Sub gridClientes_Click()

End Sub

Private Sub gridPrioridades_Click()
    cmdCopiarPri.Enabled = True
End Sub

Private Sub gridPrioridades_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
        If cmdImprListPri.Enabled = True Then
            lblFilialCtcPri = gridPrioridades.Columns(0)
            lblDataEmiPri = gridPrioridades.Columns(1)
            lblHoraEmiPri = gridPrioridades.Columns(2)
            lblModalPri = gridPrioridades.Columns(3)
            lblRemetPri = gridPrioridades.Columns(5)
            lblCidadeOrigPri = gridPrioridades.Columns(6)
            lblDestPri = gridPrioridades.Columns(7)
            lblCidadeDestPri = gridPrioridades.Columns(8)
            lblUfDestPri = gridPrioridades.Columns(9)
            lblNfsPri = gridPrioridades.Columns(10)
            lblObsEmissPri = gridPrioridades.Columns(11)
            If gridPrioridades.Columns(12) = "N" Then
                lblStatusPri = "SEM POSIÇÃO"
            ElseIf gridPrioridades.Columns(12) = "2" Then
                lblStatusPri = "EM OCORRÊNCIA"
            End If
            cmdCopiarPri.Enabled = True
        End If

End Sub

Private Sub gridUrgentes_Click()
        cmdCopiarUrg.Enabled = True
End Sub

Private Sub gridUrgentes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If cmdImprListUrg.Enabled = True Then
        lblFilialctcUrg = gridUrgentes.Columns(0)
        lblDataEmiUrg = gridUrgentes.Columns(1)
        lblHoraEmiUrg = gridUrgentes.Columns(2)
        lblModalUrg = gridUrgentes.Columns(3)
        lblRemetUrg = gridUrgentes.Columns(5)
        lblCidadeOrigUrg = gridUrgentes.Columns(6)
        lblDestUrg = gridUrgentes.Columns(7)
        lblCidadeDestUrg = gridUrgentes.Columns(8)
        lblUfDestUrg = gridUrgentes.Columns(9)
        lblNfsUrg = gridUrgentes.Columns(10)
        lblObsEmissUrg = gridUrgentes.Columns(11)
        If gridUrgentes.Columns(12) = "N" Then
            lblStatusUrg = "SEM POSIÇÃO"
        ElseIf gridUrgentes.Columns(12) = "2" Then
            lblStatusUrg = "EM OCORRÊNCIA"
        End If
        cmdCopiarUrg.Enabled = True
    End If
End Sub

Private Sub tm_atualiza_Timer()
    
    tm_atualiza.Interval = 0
    
    If de_informa.rsSel_Urgencias.State = 1 Then de_informa.rsSel_Urgencias.Close
    de_informa.Sel_Urgencias
    
    If de_informa.rsSel_Prioridades.State = 1 Then de_informa.rsSel_Prioridades.Close
    de_informa.Sel_Prioridades
    
    If Mid$(xdireitos, 30, 1) = "0" Then
        SSTab1.TabEnabled(2) = False
    Else
        SSTab1.TabEnabled(2) = True
    End If
    
    If (de_informa.rsSel_Urgencias.RecordCount + de_informa.rsSel_Prioridades.RecordCount) > 0 Then
        
        de_informa.ins_LogUsuario "CONSULTA", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    
        If de_informa.rsSel_Urgencias.RecordCount > 0 Then   'URGENCIAS
            gridUrgentes.DataMember = "sel_urgencias"
            gridUrgentes.Refresh
            lblFilialctcUrg = gridUrgentes.Columns(0)
            lblDataEmiUrg = gridUrgentes.Columns(1)
            lblHoraEmiUrg = gridUrgentes.Columns(2)
            lblModalUrg = gridUrgentes.Columns(3)
            lblRemetUrg = gridUrgentes.Columns(5)
            lblCidadeOrigUrg = gridUrgentes.Columns(6)
            lblDestUrg = gridUrgentes.Columns(7)
            lblCidadeDestUrg = gridUrgentes.Columns(8)
            lblUfDestUrg = gridUrgentes.Columns(9)
            lblNfsUrg = gridUrgentes.Columns(10)
            lblObsEmissUrg = gridUrgentes.Columns(11)
            If gridUrgentes.Columns(12) = "N" Then
                lblStatusUrg = "SEM POSIÇÃO"
            ElseIf gridUrgentes.Columns(12) = "2" Then
                lblStatusUrg = "EM OCORRÊNCIA"
            End If
            fraUrgencias = "CTCs com Urgência (Não tratando Transit-Time): " & Trim$(Str(de_informa.rsSel_Urgencias.RecordCount))
            gridUrgentes.SetFocus
            DoEvents
            Beep
        Else
            MsgBox "Não Há URGÊNCIAS Pendentes mas há PRIORIDADES Pendentes. Clique na Aba PRIORIDADES PENDENTES."
        End If
        
        If de_informa.rsSel_Prioridades.RecordCount > 0 Then   'PRIORIDADES
            gridPrioridades.DataMember = "sel_prioridades"
            gridPrioridades.Refresh
            
            lblFilialCtcPri = gridPrioridades.Columns(0)
            lblDataEmiPri = gridPrioridades.Columns(1)
            lblHoraEmiPri = gridPrioridades.Columns(2)
            lblModalPri = gridPrioridades.Columns(3)
            lblRemetPri = gridPrioridades.Columns(5)
            lblCidadeOrigPri = gridPrioridades.Columns(6)
            lblDestPri = gridPrioridades.Columns(7)
            lblCidadeDestPri = gridPrioridades.Columns(8)
            lblUfDestPri = gridPrioridades.Columns(9)
            lblNfsPri = gridPrioridades.Columns(10)
            lblObsEmissPri = gridPrioridades.Columns(11)
            If gridPrioridades.Columns(12) = "N" Then
                lblStatusPri = "SEM POSIÇÃO"
            ElseIf gridPrioridades.Columns(12) = "2" Then
                lblStatusPri = "EM OCORRÊNCIA"
            End If
            fraPrioridades = "CTCs com Prioridade (Tratando Transit-Time): " & Trim$(Str(de_informa.rsSel_Prioridades.RecordCount))
            DoEvents
            Beep
        End If
    Else
        MsgBox "Não há Urgências / Prioridades Pendentes !"
    End If
    
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    

    'preencher combos mes/ano - gerencial
    
    Call combomesano(comboMesAnoGeral)
    
    comboMesAnoGeral.ListIndex = 0
    
    'preenche clientes que podem ser analisados
    
    If de_informa.rsSel_ClientesAlarmMov.State = 1 Then de_informa.rsSel_ClientesAlarmMov.Close
    de_informa.Sel_ClientesAlarmMov
    

End Sub
