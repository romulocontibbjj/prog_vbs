VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAnEntregas 
   Caption         =   "Análise de Entregas"
   ClientHeight    =   7590
   ClientLeft      =   1470
   ClientTop       =   2070
   ClientWidth     =   12030
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGerarArq 
      Caption         =   "Gerar Arquivo"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7440
      TabIndex        =   160
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPOD 
      Caption         =   "Ocorr / POD ..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   8760
      TabIndex        =   135
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdNova 
      Caption         =   "Nova ..."
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdVaiParaSac 
      Caption         =   "Consulta SAC..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   8760
      TabIndex        =   159
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAbona 
      Caption         =   "Justificativa Atraso"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10200
      TabIndex        =   136
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   495
      Left            =   10200
      Picture         =   "frmAnEntregas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   10920
      TabIndex        =   124
      Top             =   120
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Norte / Nordeste"
      TabPicture(0)   =   "frmAnEntregas.frx":0772
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Sul / Sudeste / C.Oeste - Totais"
      TabPicture(1)   =   "frmAnEntregas.frx":078E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(2)=   "Frame10"
      Tab(1).Control(3)=   "Frame11"
      Tab(1).Control(4)=   "Frame12"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Resumo Geral"
      TabPicture(2)   =   "frmAnEntregas.frx":07AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).Control(2)=   "Frame14"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Análise dos Atrasos"
      TabPicture(3)   =   "frmAnEntregas.frx":07C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame17"
      Tab(3).Control(1)=   "Frame16"
      Tab(3).Control(2)=   "Frame24"
      Tab(3).Control(3)=   "fraAtraso"
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame17 
         Caption         =   "Detalhe Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -65400
         TabIndex        =   134
         Top             =   2520
         Width           =   2055
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Em:"
            Height          =   195
            Left            =   120
            TabIndex        =   158
            Top             =   2400
            Width           =   270
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Em:"
            Height          =   195
            Left            =   120
            TabIndex        =   157
            Top             =   1200
            Width           =   270
         End
         Begin VB.Label lblDtBxPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   142
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblRecebPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   141
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblUsuBxPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   145
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblUsuDtBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   156
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblDtBx 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   144
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblReceb 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   143
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblUsuBx 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   146
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblUsuDtBaixa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   155
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Usu:"
            Height          =   195
            Left            =   120
            TabIndex        =   154
            Top             =   2160
            Width           =   330
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Rec:"
            Height          =   195
            Left            =   120
            TabIndex        =   153
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   152
            Top             =   1680
            Width           =   390
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Usu:"
            Height          =   195
            Left            =   120
            TabIndex        =   151
            Top             =   960
            Width           =   330
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Rec:"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   149
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Física"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   148
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Pré-Baixa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   147
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "UFs"
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
         Left            =   -74880
         TabIndex        =   132
         Top             =   360
         Width           =   1095
         Begin VB.ListBox lstUFs 
            Height          =   4545
            ItemData        =   "frmAnEntregas.frx":07E2
            Left            =   120
            List            =   "frmAnEntregas.frx":07E4
            TabIndex        =   133
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Ocorrências"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -73800
         TabIndex        =   129
         Top             =   2520
         Width           =   8295
         Begin MSDataGridLib.DataGrid GridConsOcorr 
            Bindings        =   "frmAnEntregas.frx":07E6
            Height          =   1335
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   2355
            _Version        =   393216
            AllowUpdate     =   0   'False
            ForeColor       =   -2147483641
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
         Begin VB.Label lblObs_Ocorr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   960
            Left            =   120
            TabIndex        =   131
            Top             =   1680
            Width           =   8055
         End
      End
      Begin VB.Frame fraAtraso 
         Caption         =   "CTCs em Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -73800
         TabIndex        =   127
         Top             =   360
         Width           =   10455
         Begin MSDataGridLib.DataGrid gridAtrasos 
            Bindings        =   "frmAnEntregas.frx":07FF
            Height          =   1815
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3201
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
            DataMember      =   "Sel_CtcsAtrasosComAbono"
            ColumnCount     =   11
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
               DataField       =   "emissao"
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
               Caption         =   "Previsão"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "entrega"
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
               DataField       =   "prz_meta"
               Caption         =   "Prz.Meta"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "prz_real"
               Caption         =   "Prz.Real"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "prz_abono"
               Caption         =   "Abono"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            BeginProperty Column10 
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
               MarqueeStyle    =   3
               BeginProperty Column00 
                  ColumnWidth     =   1019,906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   734,74
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   599,811
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1995,024
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   345,26
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   2894,74
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   3254,74
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Gráfico por Região"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   116
         Top             =   3120
         Width           =   11535
         Begin MSChart20Lib.MSChart GrafNO 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmAnEntregas.frx":0818
            TabIndex        =   117
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSChart20Lib.MSChart GrafND 
            Height          =   1815
            Left            =   2160
            OleObjectBlob   =   "frmAnEntregas.frx":2388
            TabIndex        =   118
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSChart20Lib.MSChart GrafSD 
            Height          =   1815
            Left            =   4320
            OleObjectBlob   =   "frmAnEntregas.frx":3F01
            TabIndex        =   119
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSChart20Lib.MSChart GrafSU 
            Height          =   1815
            Left            =   6240
            OleObjectBlob   =   "frmAnEntregas.frx":5A77
            TabIndex        =   120
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSChart20Lib.MSChart GrafCO 
            Height          =   1815
            Left            =   8400
            OleObjectBlob   =   "frmAnEntregas.frx":75E1
            TabIndex        =   121
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Gráfico Total Brasil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -67080
         TabIndex        =   114
         Top             =   360
         Width           =   3735
         Begin MSChart20Lib.MSChart GrafBR 
            Height          =   2415
            Left            =   120
            OleObjectBlob   =   "frmAnEntregas.frx":917A
            TabIndex        =   115
            Top             =   240
            Visible         =   0   'False
            Width           =   3495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Por Região"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   7695
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Região NORTE"
            Height          =   195
            Left            =   240
            TabIndex        =   113
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Região NORDESTE"
            Height          =   195
            Left            =   240
            TabIndex        =   112
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Região SUDESTE"
            Height          =   195
            Left            =   240
            TabIndex        =   111
            Top             =   1200
            Width           =   1320
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Região SUL"
            Height          =   195
            Left            =   240
            TabIndex        =   110
            Top             =   1560
            Width           =   870
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Região CENTRO-OESTE"
            Height          =   195
            Left            =   240
            TabIndex        =   109
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "CTCs"
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
            Left            =   2400
            TabIndex        =   108
            Top             =   240
            Width           =   465
         End
         Begin VB.Label lblCTCsSD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   107
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblCTCsSU 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   106
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblCTCsCO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   105
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Sem Pos."
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
            Left            =   3120
            TabIndex        =   104
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   4200
            TabIndex        =   103
            Top             =   240
            Width           =   150
         End
         Begin VB.Label lblCTCsND 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   102
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblCTCsNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   101
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "No Prazo"
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
            Left            =   4800
            TabIndex        =   100
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Atraso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   5760
            TabIndex        =   99
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   6720
            TabIndex        =   98
            Top             =   240
            Width           =   150
         End
         Begin VB.Label lblNCTCsSD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3120
            TabIndex        =   97
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblNCTCsSU 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3120
            TabIndex        =   96
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblNCTCsCO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3120
            TabIndex        =   95
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblNCTCsND 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3120
            TabIndex        =   94
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblNCTCsNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3120
            TabIndex        =   93
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblPerc1SD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3960
            TabIndex        =   92
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblPerc1SU 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3960
            TabIndex        =   91
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblPerc1CO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3960
            TabIndex        =   90
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblPerc1ND 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3960
            TabIndex        =   89
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblPerc1NO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3960
            TabIndex        =   88
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblOnTimeSD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            TabIndex        =   87
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblOnTimeSU 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            TabIndex        =   86
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblOnTimeCO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            TabIndex        =   85
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblOnTimeND 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            TabIndex        =   84
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblOnTimeNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            TabIndex        =   83
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblDelaySD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   5640
            TabIndex        =   82
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblDelaySU 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   5640
            TabIndex        =   81
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblDelayCO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   5640
            TabIndex        =   80
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblDelayND 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   5640
            TabIndex        =   79
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblDelayNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   5640
            TabIndex        =   78
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblPerc2SD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6480
            TabIndex        =   77
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblPerc2SU 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6480
            TabIndex        =   76
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblPerc2CO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6480
            TabIndex        =   75
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblPerc2ND 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6480
            TabIndex        =   74
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblPerc2NO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6480
            TabIndex        =   73
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label84 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL BRASIL"
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
            Left            =   240
            TabIndex        =   72
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblCTCsBR 
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
            Left            =   2280
            TabIndex        =   71
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblNCTCsBR 
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
            Left            =   3120
            TabIndex        =   70
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblPerc1BR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
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
            Left            =   3960
            TabIndex        =   69
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblOnTimeBR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            TabIndex        =   68
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblDelayBR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   5640
            TabIndex        =   67
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblPerc2BR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
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
            Left            =   6480
            TabIndex        =   66
            Top             =   2280
            Width           =   735
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Estados"
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
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   1575
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            Caption         =   "Total Todas as Regiões"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   51
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            X1              =   960
            X2              =   120
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label31 
            Caption         =   "REGIÃO SUL"
            Height          =   435
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   855
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000010&
            X1              =   1440
            X2              =   120
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            X1              =   960
            X2              =   120
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            X1              =   120
            X2              =   1440
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label61 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MT"
            Height          =   255
            Left            =   1080
            TabIndex        =   49
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label Label60 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MS"
            Height          =   255
            Left            =   1080
            TabIndex        =   48
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label59 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GO"
            Height          =   255
            Left            =   1080
            TabIndex        =   47
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label58 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DF"
            Height          =   255
            Left            =   1080
            TabIndex        =   46
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label Label29 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SC"
            Height          =   255
            Left            =   1080
            TabIndex        =   45
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RS"
            Height          =   255
            Left            =   1080
            TabIndex        =   44
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PR"
            Height          =   255
            Left            =   1080
            TabIndex        =   43
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label26 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SP"
            Height          =   255
            Left            =   1080
            TabIndex        =   42
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label25 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RJ"
            Height          =   255
            Left            =   1080
            TabIndex        =   41
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MG"
            Height          =   255
            Left            =   1080
            TabIndex        =   40
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label23 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ES"
            Height          =   255
            Left            =   1080
            TabIndex        =   39
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "REGIÃO C.OESTE"
            Height          =   435
            Left            =   120
            TabIndex        =   38
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "REGIÃO SUDESTE"
            Height          =   435
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "CTCs"
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
         Left            =   -73320
         TabIndex        =   35
         Top             =   480
         Width           =   2175
         Begin MSFlexGridLib.MSFlexGrid FlexCtcs2 
            Height          =   2895
            Left            =   60
            TabIndex        =   52
            Top             =   480
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   12
            Cols            =   3
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flexTotCTCs 
            Height          =   375
            Left            =   60
            TabIndex        =   59
            Top             =   3720
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Capitais"
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
         Left            =   -71160
         TabIndex        =   34
         Top             =   480
         Width           =   2775
         Begin MSFlexGridLib.MSFlexGrid FlexCapitais2 
            Height          =   2895
            Left            =   60
            TabIndex        =   56
            Top             =   480
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid FlexTotCapitais 
            Height          =   375
            Left            =   60
            TabIndex        =   60
            Top             =   3720
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Interior"
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
         Left            =   -68400
         TabIndex        =   33
         Top             =   480
         Width           =   2775
         Begin MSFlexGridLib.MSFlexGrid FlexInterior2 
            Height          =   2895
            Left            =   60
            TabIndex        =   57
            Top             =   480
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid FlexTotInterior 
            Height          =   375
            Left            =   60
            TabIndex        =   61
            Top             =   3720
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Total do Estado"
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
         Left            =   -65640
         TabIndex        =   32
         Top             =   480
         Width           =   2295
         Begin MSFlexGridLib.MSFlexGrid FlexTotUf2 
            Height          =   2895
            Left            =   60
            TabIndex        =   58
            Top             =   480
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   12
            Cols            =   3
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid FlexTotTotuf 
            Height          =   375
            Left            =   60
            TabIndex        =   62
            Top             =   3720
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Total do Estado"
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
         Left            =   9360
         TabIndex        =   31
         Top             =   480
         Width           =   2295
         Begin MSFlexGridLib.MSFlexGrid FlexTotUf1 
            Height          =   4095
            Left            =   60
            TabIndex        =   55
            Top             =   480
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            Cols            =   3
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Interior"
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
         Left            =   6600
         TabIndex        =   30
         Top             =   480
         Width           =   2775
         Begin MSFlexGridLib.MSFlexGrid FlexInterior1 
            Height          =   4095
            Left            =   60
            TabIndex        =   54
            Top             =   480
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            Cols            =   4
            FixedCols       =   0
            GridColorFixed  =   16777215
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Capitais"
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
         Left            =   3840
         TabIndex        =   9
         Top             =   480
         Width           =   2775
         Begin MSFlexGridLib.MSFlexGrid FlexCapitais1 
            Height          =   4095
            Left            =   60
            TabIndex        =   53
            Top             =   480
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            Cols            =   4
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "CTCs"
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
         Left            =   1680
         TabIndex        =   10
         Top             =   480
         Width           =   2175
         Begin MSFlexGridLib.MSFlexGrid flexCtcs1 
            Height          =   4095
            Left            =   60
            TabIndex        =   126
            Top             =   480
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            Cols            =   3
            FixedCols       =   0
            ScrollBars      =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Estados"
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
         TabIndex        =   11
         Top             =   480
         Width           =   1575
         Begin VB.Line Line3 
            BorderColor     =   &H80000010&
            X1              =   1440
            X2              =   120
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   960
            X2              =   120
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   120
            X2              =   1440
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label12 
            Caption         =   "REGIÃO NORTE"
            Height          =   435
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "REGIÃO NORDESTE"
            Height          =   435
            Left            =   120
            TabIndex        =   28
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AC"
            Height          =   255
            Left            =   1080
            TabIndex        =   27
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AM"
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AP"
            Height          =   255
            Left            =   1080
            TabIndex        =   25
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PA"
            Height          =   255
            Left            =   1080
            TabIndex        =   24
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label30 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RO"
            Height          =   255
            Left            =   1080
            TabIndex        =   23
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RR"
            Height          =   255
            Left            =   1080
            TabIndex        =   22
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TO"
            Height          =   255
            Left            =   1080
            TabIndex        =   21
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AL"
            Height          =   255
            Left            =   1080
            TabIndex        =   20
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BA"
            Height          =   255
            Left            =   1080
            TabIndex        =   19
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SE"
            Height          =   255
            Left            =   1080
            TabIndex        =   18
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PE"
            Height          =   255
            Left            =   1080
            TabIndex        =   17
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PB"
            Height          =   255
            Left            =   1080
            TabIndex        =   16
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RN"
            Height          =   255
            Left            =   1080
            TabIndex        =   15
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CE"
            Height          =   255
            Left            =   1080
            TabIndex        =   14
            Top             =   3840
            Width           =   375
         End
         Begin VB.Label Label20 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PI"
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   4080
            Width           =   375
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MA"
            Height          =   255
            Left            =   1080
            TabIndex        =   12
            Top             =   4320
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Selecionados"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin VB.CheckBox chkAnalise 
         Caption         =   "abono"
         Height          =   255
         Left            =   3600
         TabIndex        =   138
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCgcRemet 
         AutoSize        =   -1  'True
         Caption         =   "CGC"
         Height          =   195
         Left            =   1920
         TabIndex        =   137
         Top             =   120
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblModal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   123
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Modal:"
         Height          =   195
         Left            =   4920
         TabIndex        =   122
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblPrazo 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6360
         TabIndex        =   64
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Tab. Prazo:"
         Height          =   195
         Left            =   5520
         TabIndex        =   63
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cliente / Remetente:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Período..............: De"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1935
         TabIndex        =   5
         Top             =   360
         Width           =   3465
      End
      Begin VB.Label lblDataPer1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1935
         TabIndex        =   4
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   3240
         TabIndex        =   3
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblDataPer2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   720
         Width           =   1170
      End
   End
   Begin VB.Label lblAutoriza 
      AutoSize        =   -1  'True
      Caption         =   "NAO"
      Height          =   195
      Left            =   600
      TabIndex        =   140
      Top             =   6960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lbltrata 
      AutoSize        =   -1  'True
      Caption         =   "nao"
      Height          =   195
      Left            =   9600
      TabIndex        =   139
      Top             =   6960
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmAnEntregas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Graf1_OLEStartDrag(Data As MSChart20Lib.DataObject, AllowedEffects As Long)

End Sub

Private Sub cmdPOD_Click()
    xultimofilial = Mid(gridAtrasos.Columns(0), 1, 2)
    xultimoctc = Mid(gridAtrasos.Columns(0), 3, 8)
    frmPod.Caption = "Informação de Entregas e Ocorrências - An. Entregas (Atrasos)"
    frmPod.TxtFilial = Mid(gridAtrasos.Columns(0), 1, 2)
    frmPod.txtCtc = Mid(gridAtrasos.Columns(0), 3, 8)
    DoEvents
    frmPod.Show
    DoEvents
    frmPod.cmdProcurar.SetFocus
    DoEvents
    SendKeys "{ENTER}"
    DoEvents
End Sub

Private Sub cmdVaiParaSac_Click()
    xultimofilial = Mid(gridAtrasos.Columns(0), 1, 2)
    xultimoctc = Mid(gridAtrasos.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (An. Atrasos)"
    frmSac.TxtFilial = Mid(gridAtrasos.Columns(0), 1, 2)
    frmSac.txtCtc = Mid(gridAtrasos.Columns(0), 3, 8)
    DoEvents
    frmSac.Show
    DoEvents
    frmSac.cmbProcurar.SetFocus
    DoEvents
    SendKeys "{ENTER}"
    DoEvents
End Sub

Private Sub Command1_Click()

End Sub

Private Sub GridConsOcorr_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
End Sub

Private Sub cmdAbrir_Click()

End Sub
Private Sub cmdAbona_Click()
    If Mid$(xdireitos, 31, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmAbonaAtraso.lblPrazoContr = Val(gridAtrasos.Columns(4))
        frmAbonaAtraso.lblPrazoReal = Val(gridAtrasos.Columns(5))
        frmAbonaAtraso.lblDiasAtraso = Val(gridAtrasos.Columns(5)) - Val(gridAtrasos.Columns(4))
        frmAbonaAtraso.txtDiasJustif = Val(frmAbonaAtraso.lblDiasAtraso)
        frmAbonaAtraso.Show 1
    End If
End Sub

Private Sub cmdImprTela_Click()
    Printer.KillDoc
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Printer.Zoom = 50
    Me.PrintForm
End Sub

Private Sub cmdNova_Click()
    Unload frmAnEntregas
    frmEscCliPer.Caption = "Análise de Entregas"
    frmEscCliPer.Show 1
End Sub

Private Sub cmdSair_Click()
    Set frmAnEntregas = Nothing
    Unload frmEscCliPer
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    
'configura tamanho das colunas das flexgrids

'tabs1
    flexCtcs1.Cols = 3
    flexCtcs1.ColWidth(0) = 660  'CTCs tab1
    flexCtcs1.ColWidth(1) = 750
    flexCtcs1.ColWidth(2) = 620
    FlexCapitais1.ColWidth(0) = 540  'Capitais Tab1
    FlexCapitais1.ColWidth(1) = 750
    FlexCapitais1.ColWidth(2) = 720
    FlexCapitais1.ColWidth(3) = 620
    FlexInterior1.ColWidth(0) = 540  'Interior tab1
    FlexInterior1.ColWidth(1) = 750
    FlexInterior1.ColWidth(2) = 720
    FlexInterior1.ColWidth(3) = 620
    FlexTotUf1.ColWidth(0) = 750  'Total UF tab1
    FlexTotUf1.ColWidth(1) = 750
    FlexTotUf1.ColWidth(2) = 650
    
'tabs2
    FlexCtcs2.ColWidth(0) = 660  'CTCs tab2
    FlexCtcs2.ColWidth(1) = 750
    FlexCtcs2.ColWidth(2) = 620
    FlexCapitais2.ColWidth(0) = 540  'Capitais Tab2
    FlexCapitais2.ColWidth(1) = 750
    FlexCapitais2.ColWidth(2) = 720
    FlexCapitais2.ColWidth(3) = 620
    FlexInterior2.ColWidth(0) = 540  'Interior tab2
    FlexInterior2.ColWidth(1) = 750
    FlexInterior2.ColWidth(2) = 720
    FlexInterior2.ColWidth(3) = 620
    FlexTotUf2.ColWidth(0) = 750  'Total UF tab2
    FlexTotUf2.ColWidth(1) = 750
    FlexTotUf2.ColWidth(2) = 650
'Totais do tab2
    flexTotCTCs.ColWidth(0) = 660  'Totais CTCs tab2
    flexTotCTCs.ColWidth(1) = 750
    flexTotCTCs.ColWidth(2) = 620
    FlexTotCapitais.ColWidth(0) = 540  'Totais Capitais Tab2
    FlexTotCapitais.ColWidth(1) = 750
    FlexTotCapitais.ColWidth(2) = 720
    FlexTotCapitais.ColWidth(3) = 620
    FlexTotInterior.ColWidth(0) = 540  'Totais Interior tab2
    FlexTotInterior.ColWidth(1) = 750
    FlexTotInterior.ColWidth(2) = 720
    FlexTotInterior.ColWidth(3) = 620
    FlexTotTotuf.ColWidth(0) = 750  'Totais do Total UF tab2
    FlexTotTotuf.ColWidth(1) = 750
    FlexTotTotuf.ColWidth(2) = 650
    
'configura dos cabeçarios das flexgrids
    
'tab1
    flexCtcs1.Row = 0       'CTCs tab1
    flexCtcs1.Col = 0
    flexCtcs1.Text = "  Qtde."
    flexCtcs1.Col = 1
    flexCtcs1.Text = "Sem Pos."
    flexCtcs1.Col = 2
    flexCtcs1.Text = "   %"
    FlexCapitais1.Row = 0   'Capitais tab1
    FlexCapitais1.Col = 0
    FlexCapitais1.Text = " Meta"
    FlexCapitais1.Col = 1
    FlexCapitais1.Text = "No Prazo"
    FlexCapitais1.Col = 2
    FlexCapitais1.Text = "Atraso"
    FlexCapitais1.Col = 3
    FlexCapitais1.Text = "   %"
    FlexInterior1.Row = 0   'Interior tab1
    FlexInterior1.Col = 0
    FlexInterior1.Text = " Meta"
    FlexInterior1.Col = 1
    FlexInterior1.Text = "No Prazo"
    FlexInterior1.Col = 2
    FlexInterior1.Text = "Atraso"
    FlexInterior1.Col = 3
    FlexInterior1.Text = "   %"
    FlexTotUf1.Row = 0   'Total do Estado tab1
    FlexTotUf1.Col = 0
    FlexTotUf1.Text = "No Prazo"
    FlexTotUf1.Col = 1
    FlexTotUf1.Text = "Atraso"
    FlexTotUf1.Col = 2
    FlexTotUf1.Text = "   %"
    
'tab2
    FlexCtcs2.Row = 0       'CTCs tab2
    FlexCtcs2.Col = 0
    FlexCtcs2.Text = "  Qtde."
    FlexCtcs2.Col = 1
    FlexCtcs2.Text = "Sem Pos."
    FlexCtcs2.Col = 2
    FlexCtcs2.Text = "   %"
    FlexCapitais2.Row = 0   'Capitais tab2
    FlexCapitais2.Col = 0
    FlexCapitais2.Text = " Meta"
    FlexCapitais2.Col = 1
    FlexCapitais2.Text = "No Prazo"
    FlexCapitais2.Col = 2
    FlexCapitais2.Text = "Atraso"
    FlexCapitais2.Col = 3
    FlexCapitais2.Text = "   %"
    FlexInterior2.Row = 0   'Interior tab2
    FlexInterior2.Col = 0
    FlexInterior2.Text = " Meta"
    FlexInterior2.Col = 1
    FlexInterior2.Text = "No Prazo"
    FlexInterior2.Col = 2
    FlexInterior2.Text = "Atraso"
    FlexInterior2.Col = 3
    FlexInterior2.Text = "   %"
    FlexTotUf2.Row = 0   'Total do Estado tab2
    FlexTotUf2.Col = 0
    FlexTotUf2.Text = "No Prazo"
    FlexTotUf2.Col = 1
    FlexTotUf2.Text = "Atraso"
    FlexTotUf2.Col = 2
    FlexTotUf2.Text = "   %"
    
    lstUFs.Clear
    lstUFs.AddItem "Todos"
    lstUFs.AddItem "AC"
    lstUFs.AddItem "AM"
    lstUFs.AddItem "AP"
    lstUFs.AddItem "PA"
    lstUFs.AddItem "RO"
    lstUFs.AddItem "RR"
    lstUFs.AddItem "TO"
    lstUFs.AddItem "AL"
    lstUFs.AddItem "BA"
    lstUFs.AddItem "SE"
    lstUFs.AddItem "PE"
    lstUFs.AddItem "PB"
    lstUFs.AddItem "RN"
    lstUFs.AddItem "CE"
    lstUFs.AddItem "PI"
    lstUFs.AddItem "MA"
    lstUFs.AddItem "ES"
    lstUFs.AddItem "MG"
    lstUFs.AddItem "RJ"
    lstUFs.AddItem "SP"
    lstUFs.AddItem "PR"
    lstUFs.AddItem "RS"
    lstUFs.AddItem "SC"
    lstUFs.AddItem "DF"
    lstUFs.AddItem "GO"
    lstUFs.AddItem "MS"
    lstUFs.AddItem "MT"
    


End Sub

Private Sub MSChart1_OLEStartDrag(Data As MSChart20Lib.DataObject, AllowedEffects As Long)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmAnEntregas = Nothing
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub gridAtrasos_GotFocus()
    lbltrata = "SIM"

    cmdVaiParaSac.Enabled = True
    cmdPOD.Enabled = True

    If chkAnalise.Value = 1 Then
        If de_informa.rsSel_CtcsAtrasosComAbono.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
            frmAnEntregas.GridConsOcorr.Refresh
            
            If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                cmdAbona.Enabled = True

            Else
                lblObs_Ocorr.Caption = ""
                cmdAbona.Enabled = False
            End If
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"

            frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
            Else
                frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            End If
        End If
    Else
        If de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
            frmAnEntregas.GridConsOcorr.Refresh
            
            If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                cmdAbona.Enabled = True
            Else
                lblObs_Ocorr.Caption = ""
                cmdAbona.Enabled = False
            End If
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
            Else
                frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            End If
        End If
    End If
    
   
End Sub

Private Sub gridAtrasos_LostFocus()
lbltrata = "NAO"
End Sub

Private Sub gridAtrasos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If lbltrata = "SIM" Then
    If chkAnalise.Value = 1 Then
        If de_informa.rsSel_CtcsAtrasosComAbono.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
            frmAnEntregas.GridConsOcorr.Refresh
            
            If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                cmdAbona.Enabled = True
            Else
                lblObs_Ocorr.Caption = ""
                cmdAbona.Enabled = False
            End If

            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"

            frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
            Else
                frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            End If
        End If
    Else
        If de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
            frmAnEntregas.GridConsOcorr.Refresh
            
            If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                cmdAbona.Enabled = True
            Else
                lblObs_Ocorr.Caption = ""
                cmdAbona.Enabled = False
            End If
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
            Else
                frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            End If
        End If
    End If
    
End If
End Sub

Private Sub GridConsOcorr_Click()
    lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
End Sub

Private Sub lstUFs_Click()
    If Mid$(lblModal, 1, 1) = "R" Then
        xmodal = "RODOVIARIO"
    Else
        xmodal = "AEREO"
    End If
    
    If lstUFs.Text = "Todos" Then
        If chkAnalise.Value = 1 Then
            If de_informa.rsSel_CtcsAtrasosComAbono.State = 1 Then de_informa.rsSel_CtcsAtrasosComAbono.Close
            de_informa.Sel_CtcsAtrasosComAbono lblCgcRemet, CDate(lblDataPer1), CDate(lblDataPer2), xmodal, "%"
            frmAnEntregas.gridAtrasos.DataMember = "sel_ctcsatrasoscomabono"
            frmAnEntregas.gridAtrasos.Refresh
            fraAtraso.Caption = "CTCs em Atraso: " & de_informa.rsSel_CtcsAtrasosComAbono.RecordCount
            If de_informa.rsSel_CtcsAtrasosComAbono.RecordCount > 0 Then
                If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
                
                frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
                frmAnEntregas.GridConsOcorr.Refresh
                
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                Else
                    lblObs_Ocorr.Caption = ""
                End If
                
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
    
                frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
                frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
                frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
                
                If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                    frmAnEntregas.lblDtBx = ""
                    frmAnEntregas.lblReceb = ""
                    frmAnEntregas.lblUsuBx = ""
                    frmAnEntregas.lblUsuDtBaixa = ""
                Else
                    frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                    frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                    frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                    frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
                End If
                frmAnEntregas.cmdAbona.Enabled = True
            Else
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                GridConsOcorr.DataMember = "Sel_ConsOcorr"
                GridConsOcorr.Refresh
                frmAnEntregas.lblDtBxPre = ""
                frmAnEntregas.lblRecebPre = ""
                frmAnEntregas.lblUsuBxPre = ""
                frmAnEntregas.lblUsuDtBaixaPre = ""
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblObs_Ocorr = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
                frmAnEntregas.cmdAbona.Enabled = False
            End If
        Else
            If de_informa.rsSel_CtcsAtrasosSemAbono.State = 1 Then de_informa.rsSel_CtcsAtrasosSemAbono.Close
            de_informa.Sel_CtcsAtrasosSemAbono lblCgcRemet, CDate(lblDataPer1), CDate(lblDataPer2), xmodal, "%"
            frmAnEntregas.gridAtrasos.DataMember = "sel_ctcsatrasossemabono"
            frmAnEntregas.gridAtrasos.Refresh
            fraAtraso.Caption = "CTCs em Atraso: " & de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount
            If de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount > 0 Then
                If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
                
                frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
                frmAnEntregas.GridConsOcorr.Refresh
                
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                Else
                    lblObs_Ocorr.Caption = ""
                End If
                
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
                
                frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
                frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
                frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
                
                If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                    frmAnEntregas.lblDtBx = ""
                    frmAnEntregas.lblReceb = ""
                    frmAnEntregas.lblUsuBx = ""
                    frmAnEntregas.lblUsuDtBaixa = ""
                Else
                    frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                    frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                    frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                    frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
                End If
                frmAnEntregas.cmdAbona.Enabled = True
            Else
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                GridConsOcorr.DataMember = "Sel_ConsOcorr"
                GridConsOcorr.Refresh
                frmAnEntregas.lblDtBxPre = ""
                frmAnEntregas.lblRecebPre = ""
                frmAnEntregas.lblUsuBxPre = ""
                frmAnEntregas.lblUsuDtBaixaPre = ""
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblObs_Ocorr = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
                frmAnEntregas.cmdAbona.Enabled = False
            End If
        End If
    Else
        If chkAnalise.Value = 1 Then
            If de_informa.rsSel_CtcsAtrasosComAbono.State = 1 Then de_informa.rsSel_CtcsAtrasosComAbono.Close
            de_informa.Sel_CtcsAtrasosComAbono lblCgcRemet, CDate(lblDataPer1), CDate(lblDataPer2), xmodal, lstUFs.Text
            frmAnEntregas.gridAtrasos.DataMember = "sel_ctcsatrasoscomabono"
            frmAnEntregas.gridAtrasos.Refresh
            fraAtraso.Caption = "CTCs em Atraso: " & de_informa.rsSel_CtcsAtrasosComAbono.RecordCount
            If de_informa.rsSel_CtcsAtrasosComAbono.RecordCount > 0 Then
                If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
                
                frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
                frmAnEntregas.GridConsOcorr.Refresh
                
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                Else
                    lblObs_Ocorr.Caption = ""
                End If
                
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
    
                frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
                frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
                frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
                
                If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                    frmAnEntregas.lblDtBx = ""
                    frmAnEntregas.lblReceb = ""
                    frmAnEntregas.lblUsuBx = ""
                    frmAnEntregas.lblUsuDtBaixa = ""
                Else
                    frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                    frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                    frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                    frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
                End If
                frmAnEntregas.cmdAbona.Enabled = True
            Else
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                GridConsOcorr.DataMember = "Sel_ConsOcorr"
                GridConsOcorr.Refresh
                frmAnEntregas.lblDtBxPre = ""
                frmAnEntregas.lblRecebPre = ""
                frmAnEntregas.lblUsuBxPre = ""
                frmAnEntregas.lblUsuDtBaixaPre = ""
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblObs_Ocorr = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
                
                frmAnEntregas.cmdAbona.Enabled = False
            End If
        Else
            If de_informa.rsSel_CtcsAtrasosSemAbono.State = 1 Then de_informa.rsSel_CtcsAtrasosSemAbono.Close
            de_informa.Sel_CtcsAtrasosSemAbono lblCgcRemet, CDate(lblDataPer1), CDate(lblDataPer2), xmodal, lstUFs.Text
            frmAnEntregas.gridAtrasos.DataMember = "sel_ctcsatrasossemabono"
            frmAnEntregas.gridAtrasos.Refresh
            fraAtraso.Caption = "CTCs em Atraso: " & de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount
            If de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount > 0 Then
                If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
                
                frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
                frmAnEntregas.GridConsOcorr.Refresh
                
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
                Else
                    lblObs_Ocorr.Caption = ""
                End If
                
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
                
                frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
                frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
                frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
                
                If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                    frmAnEntregas.lblDtBx = ""
                    frmAnEntregas.lblReceb = ""
                    frmAnEntregas.lblUsuBx = ""
                    frmAnEntregas.lblUsuDtBaixa = ""
                Else
                    frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                    frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                    frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                    frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
                End If
                frmAnEntregas.cmdAbona.Enabled = True
            Else
                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                GridConsOcorr.DataMember = "Sel_ConsOcorr"
                GridConsOcorr.Refresh
                frmAnEntregas.lblDtBxPre = ""
                frmAnEntregas.lblRecebPre = ""
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblObs_Ocorr = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
                frmAnEntregas.cmdAbona.Enabled = False
            End If
        End If
    End If
            

End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Or SSTab1.Tab = 2 Then
        cmdAbona.Enabled = False
        cmdVaiParaSac.Enabled = False
        cmdPOD.Enabled = False
    End If
End Sub

