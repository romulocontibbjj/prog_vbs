VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAnEstat 
   Caption         =   "Análise Estatística"
   ClientHeight    =   8370
   ClientLeft      =   825
   ClientTop       =   960
   ClientWidth     =   12030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGeraXls 
      Caption         =   "Gerar no EXCEL ..."
      Height          =   375
      Left            =   8160
      TabIndex        =   228
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   495
      Left            =   10080
      Picture         =   "frmAnOper.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   10920
      TabIndex        =   153
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdNova 
      Caption         =   "Nova ..."
      Height          =   375
      Left            =   8160
      TabIndex        =   152
      Top             =   240
      Width           =   1695
   End
   Begin TabDlg.SSTab TabAnOper 
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Norte/Nordeste"
      TabPicture(0)   =   "frmAnOper.frx":0772
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame24"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Sudeste/Sul/C.Oeste"
      TabPicture(1)   =   "frmAnOper.frx":078E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame11"
      Tab(1).Control(2)=   "Frame25"
      Tab(1).Control(3)=   "Frame13"
      Tab(1).Control(4)=   "Frame10"
      Tab(1).Control(5)=   "Frame8"
      Tab(1).Control(6)=   "Frame12"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Estatística por Peso"
      TabPicture(2)   =   "frmAnOper.frx":07AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).Control(1)=   "Frame17"
      Tab(2).Control(2)=   "Frame16"
      Tab(2).Control(3)=   "Frame15"
      Tab(2).Control(4)=   "Frame14"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Estatística por Val.Merc."
      TabPicture(3)   =   "frmAnOper.frx":07C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame19"
      Tab(3).Control(1)=   "Frame20"
      Tab(3).Control(2)=   "Frame21"
      Tab(3).Control(3)=   "Frame22"
      Tab(3).Control(4)=   "Frame23"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Tráfego Mútuo"
      TabPicture(4)   =   "frmAnOper.frx":07E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "tabSub"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Índices"
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
         Left            =   8520
         TabIndex        =   19
         Top             =   840
         Width           =   3135
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPercentNONE 
            Height          =   4095
            Left            =   120
            TabIndex        =   160
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Expedições"
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
         Left            =   7320
         TabIndex        =   21
         Top             =   840
         Width           =   1215
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexExpedNONE 
            Height          =   4095
            Left            =   120
            TabIndex        =   159
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Volumes"
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
         Left            =   6120
         TabIndex        =   229
         Top             =   840
         Width           =   1215
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexVolNONE 
            Height          =   4095
            Left            =   120
            TabIndex        =   230
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Peso"
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
         Left            =   4920
         TabIndex        =   22
         Top             =   840
         Width           =   1215
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPesoNONE 
            Height          =   4095
            Left            =   120
            TabIndex        =   158
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frete"
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
         Left            =   3240
         TabIndex        =   20
         Top             =   840
         Width           =   1695
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexFreteNONE 
            Height          =   4095
            Left            =   120
            TabIndex        =   157
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Valor Mercadoria"
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
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   1695
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexValmerNONE 
            Height          =   4095
            Left            =   120
            TabIndex        =   156
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
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
         TabIndex        =   10
         Top             =   840
         Width           =   1455
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexUF1 
            Height          =   4095
            Left            =   840
            TabIndex        =   161
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   17
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Região Nordeste"
            Height          =   435
            Left            =   120
            TabIndex        =   12
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Região Norte"
            Height          =   435
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   615
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   120
            X2              =   1320
            Y1              =   740
            Y2              =   740
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   1320
            X2              =   120
            Y1              =   2420
            Y2              =   2420
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000010&
            X1              =   1320
            X2              =   120
            Y1              =   4560
            Y2              =   4560
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Índices"
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
         Left            =   -66480
         TabIndex        =   24
         Top             =   840
         Width           =   3135
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPercenttot 
            Height          =   375
            Left            =   120
            TabIndex        =   172
            Top             =   3600
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPercentSDSU 
            Height          =   3015
            Left            =   120
            TabIndex        =   167
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   12
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Expedições"
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
         Left            =   -67680
         TabIndex        =   26
         Top             =   840
         Width           =   1215
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexExpedTot 
            Height          =   375
            Left            =   120
            TabIndex        =   171
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexExpedSDSU 
            Height          =   3015
            Left            =   120
            TabIndex        =   166
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   12
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Volumes"
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
         Left            =   -68880
         TabIndex        =   231
         Top             =   840
         Width           =   1215
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexVolTot 
            Height          =   375
            Left            =   120
            TabIndex        =   232
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexVolSDSU 
            Height          =   3015
            Left            =   120
            TabIndex        =   233
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   12
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Peso"
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
         Left            =   -70080
         TabIndex        =   27
         Top             =   840
         Width           =   1215
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPesoTot 
            Height          =   375
            Left            =   120
            TabIndex        =   170
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPesoSDSU 
            Height          =   3015
            Left            =   120
            TabIndex        =   165
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   12
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frete"
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
         Left            =   -71760
         TabIndex        =   25
         Top             =   840
         Width           =   1695
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexFreteTot 
            Height          =   375
            Left            =   120
            TabIndex        =   169
            Top             =   3600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexFreteSDSU 
            Height          =   3015
            Left            =   120
            TabIndex        =   164
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   12
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Valor Mercadoria"
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
         Left            =   -73440
         TabIndex        =   23
         Top             =   840
         Width           =   1695
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexValMerTot 
            Height          =   375
            Left            =   120
            TabIndex        =   168
            Top             =   3600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexValMerSDSU 
            Height          =   3015
            Left            =   120
            TabIndex        =   163
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   12
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
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
         TabIndex        =   13
         Top             =   840
         Width           =   1455
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexUF2 
            Height          =   2895
            Left            =   840
            TabIndex        =   162
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   12
            FixedCols       =   0
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label21 
            Caption         =   "Região Sudeste"
            Height          =   435
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Região Centro Oeste"
            Height          =   555
            Left            =   120
            TabIndex        =   16
            Top             =   2520
            Width           =   615
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            X1              =   120
            X2              =   1320
            Y1              =   740
            Y2              =   740
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            X1              =   1320
            X2              =   120
            Y1              =   1690
            Y2              =   1690
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000010&
            X1              =   1320
            X2              =   120
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label Label31 
            Caption         =   "Região Sul"
            Height          =   435
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   615
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            X1              =   1320
            X2              =   120
            Y1              =   2410
            Y2              =   2410
         End
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
            TabIndex        =   14
            Top             =   3600
            Width           =   1095
         End
      End
      Begin TabDlg.SSTab tabSub 
         Height          =   6375
         Left            =   -74880
         TabIndex        =   155
         Top             =   480
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11245
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Por UF"
         TabPicture(0)   =   "frmAnOper.frx":07FE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FlexMutuoPorUF"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdGeraArqMutuoUF"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Resumo"
         TabPicture(1)   =   "frmAnOper.frx":081A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "flexMutuoResumo"
         Tab(1).Control(1)=   "tabMutuoTotais"
         Tab(1).Control(2)=   "grafValmerc"
         Tab(1).Control(3)=   "grafReceitaTran"
         Tab(1).Control(4)=   "grafReceitaTotal"
         Tab(1).Control(5)=   "grafReceita"
         Tab(1).ControlCount=   6
         Begin VB.CommandButton cmdGeraArqMutuoUF 
            Caption         =   "Gerar Arquivo TXT (Por UF) ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   173
            Top             =   480
            Width           =   2655
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexMutuoPorUF 
            Bindings        =   "frmAnOper.frx":0836
            Height          =   5295
            Left            =   120
            TabIndex        =   174
            Top             =   960
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   9340
            _Version        =   393216
            Cols            =   11
            FocusRect       =   0
            HighLight       =   0
            DataMember      =   "Sel_AnEstatSubPorUF"
            _NumberOfBands  =   1
            _Band(0).Cols   =   11
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
            _Band(0)._NumMapCols=   10
            _Band(0)._MapCol(0)._Name=   "transp_sub"
            _Band(0)._MapCol(0)._RSIndex=   0
            _Band(0)._MapCol(1)._Name=   "regiaogeo"
            _Band(0)._MapCol(1)._RSIndex=   1
            _Band(0)._MapCol(2)._Name=   "uf_dest"
            _Band(0)._MapCol(2)._RSIndex=   2
            _Band(0)._MapCol(3)._Name=   "qtd"
            _Band(0)._MapCol(3)._RSIndex=   3
            _Band(0)._MapCol(3)._Alignment=   7
            _Band(0)._MapCol(4)._Name=   "tvalmerc"
            _Band(0)._MapCol(4)._RSIndex=   4
            _Band(0)._MapCol(4)._Alignment=   7
            _Band(0)._MapCol(5)._Name=   "tfretetotal"
            _Band(0)._MapCol(5)._RSIndex=   5
            _Band(0)._MapCol(5)._Alignment=   7
            _Band(0)._MapCol(6)._Name=   "tfretepago"
            _Band(0)._MapCol(6)._RSIndex=   6
            _Band(0)._MapCol(6)._Alignment=   7
            _Band(0)._MapCol(7)._Name=   "receita"
            _Band(0)._MapCol(7)._RSIndex=   7
            _Band(0)._MapCol(7)._Alignment=   7
            _Band(0)._MapCol(8)._Name=   "tpeso"
            _Band(0)._MapCol(8)._RSIndex=   8
            _Band(0)._MapCol(9)._Name=   "tvolumes"
            _Band(0)._MapCol(9)._RSIndex=   9
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexMutuoResumo 
            Bindings        =   "frmAnOper.frx":084F
            Height          =   2895
            Left            =   -74880
            TabIndex        =   175
            Top             =   480
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   5106
            _Version        =   393216
            Cols            =   9
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            SelectionMode   =   1
            DataMember      =   "Sel_AnEstatSubGrid"
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
            _Band(0)._NumMapCols=   8
            _Band(0)._MapCol(0)._Name=   "transp_sub"
            _Band(0)._MapCol(0)._RSIndex=   0
            _Band(0)._MapCol(1)._Name=   "tfretetotal"
            _Band(0)._MapCol(1)._RSIndex=   5
            _Band(0)._MapCol(1)._Alignment=   7
            _Band(0)._MapCol(2)._Name=   "tfretepago"
            _Band(0)._MapCol(2)._RSIndex=   6
            _Band(0)._MapCol(2)._Alignment=   7
            _Band(0)._MapCol(3)._Name=   "receita"
            _Band(0)._MapCol(3)._RSIndex=   7
            _Band(0)._MapCol(3)._Alignment=   7
            _Band(0)._MapCol(4)._Name=   "qtd"
            _Band(0)._MapCol(4)._RSIndex=   1
            _Band(0)._MapCol(4)._Alignment=   7
            _Band(0)._MapCol(5)._Name=   "tpeso"
            _Band(0)._MapCol(5)._RSIndex=   2
            _Band(0)._MapCol(6)._Name=   "tvolumes"
            _Band(0)._MapCol(6)._RSIndex=   3
            _Band(0)._MapCol(7)._Name=   "tvalmerc"
            _Band(0)._MapCol(7)._RSIndex=   4
            _Band(0)._MapCol(7)._Alignment=   7
         End
         Begin TabDlg.SSTab tabMutuoTotais 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   176
            Top             =   3480
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   4895
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Totais"
            TabPicture(0)   =   "frmAnOper.frx":0868
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblTotReceitaporCTC"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label80"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblTotReceitaValor"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label78"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblTotFretePagoValor"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label76"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lblTotFreteCobrValor"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label74"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label73"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "lblTotVol"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label33"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "lblTotValmerc"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label45"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "lblTotFreteCobr"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "Label50"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "lblTotFretePago"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Label56"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "lblTotReceita"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "Label62"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "lblTotPercCobrPago"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "Label48"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "lblTotCtc"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "Label52"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "lblTotPeso"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).ControlCount=   24
            TabCaption(1)   =   "Transportador"
            TabPicture(1)   =   "frmAnOper.frx":0884
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label5"
            Tab(1).Control(1)=   "Label70"
            Tab(1).Control(2)=   "Label68"
            Tab(1).Control(3)=   "Label66"
            Tab(1).Control(4)=   "Label64"
            Tab(1).Control(5)=   "Label54"
            Tab(1).Control(6)=   "lblTotTranCobrPago"
            Tab(1).Control(7)=   "lblTotTranValMinimo"
            Tab(1).Control(8)=   "lblTotTranPercNegoc"
            Tab(1).Control(9)=   "Label2"
            Tab(1).Control(10)=   "Label1"
            Tab(1).Control(11)=   "lblTotTranPercFreteCobrValor"
            Tab(1).Control(12)=   "Label4"
            Tab(1).Control(13)=   "lblTotTranPercFretePagoValor"
            Tab(1).Control(14)=   "Label6"
            Tab(1).Control(15)=   "lblTotTranPercReceitaValor"
            Tab(1).Control(16)=   "Label8"
            Tab(1).Control(17)=   "lblTotTranReceitaporCTC"
            Tab(1).Control(18)=   "lblTransportador"
            Tab(1).Control(19)=   "lblTotTranVol"
            Tab(1).Control(20)=   "lblTotTranPeso"
            Tab(1).Control(21)=   "lblTotTranCtc"
            Tab(1).Control(22)=   "lblTotTranValMerc"
            Tab(1).ControlCount=   23
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Val. Mercad."
               Height          =   195
               Left            =   -74880
               TabIndex        =   227
               Top             =   480
               Width           =   900
            End
            Begin VB.Label lblTotPeso 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1560
               TabIndex        =   221
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "Peso Transportado:"
               Height          =   195
               Left            =   120
               TabIndex        =   220
               Top             =   1200
               Width           =   1395
            End
            Begin VB.Label lblTotCtc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1560
               TabIndex        =   219
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Quantidade CTCs:"
               Height          =   195
               Left            =   120
               TabIndex        =   218
               Top             =   840
               Width           =   1305
            End
            Begin VB.Label lblTotPercCobrPago 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3960
               TabIndex        =   217
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               Caption         =   "Cobrado/Pago %:"
               Height          =   195
               Left            =   2640
               TabIndex        =   216
               Top             =   1560
               Width           =   1260
            End
            Begin VB.Label lblTotReceita 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3600
               TabIndex        =   215
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Receita:"
               Height          =   195
               Left            =   2640
               TabIndex        =   214
               Top             =   1200
               Width           =   600
            End
            Begin VB.Label lblTotFretePago 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3600
               TabIndex        =   213
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               Caption         =   "Frete Pago:"
               Height          =   195
               Left            =   2640
               TabIndex        =   212
               Top             =   840
               Width           =   825
            End
            Begin VB.Label lblTotFreteCobr 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3600
               TabIndex        =   211
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Frete Cobr.:"
               Height          =   195
               Left            =   2640
               TabIndex        =   210
               Top             =   480
               Width           =   825
            End
            Begin VB.Label lblTotValmerc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1080
               TabIndex        =   209
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Val. Mercad."
               Height          =   195
               Left            =   120
               TabIndex        =   208
               Top             =   480
               Width           =   900
            End
            Begin VB.Label lblTotVol 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1560
               TabIndex        =   207
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label73 
               AutoSize        =   -1  'True
               Caption         =   "Volumes Transp:"
               Height          =   195
               Left            =   120
               TabIndex        =   206
               Top             =   1560
               Width           =   1185
            End
            Begin VB.Label Label74 
               AutoSize        =   -1  'True
               Caption         =   "Frete Cobr/Valor:"
               Height          =   195
               Left            =   120
               TabIndex        =   205
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label lblTotFreteCobrValor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1560
               TabIndex        =   204
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label76 
               AutoSize        =   -1  'True
               Caption         =   "Frete Pago/Valor:"
               Height          =   195
               Left            =   120
               TabIndex        =   203
               Top             =   2400
               Width           =   1260
            End
            Begin VB.Label lblTotFretePagoValor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1560
               TabIndex        =   202
               Top             =   2400
               Width           =   975
            End
            Begin VB.Label Label78 
               AutoSize        =   -1  'True
               Caption         =   "Receita/Valor:"
               Height          =   195
               Left            =   2640
               TabIndex        =   201
               Top             =   2040
               Width           =   1035
            End
            Begin VB.Label lblTotReceitaValor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3960
               TabIndex        =   200
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Label80 
               AutoSize        =   -1  'True
               Caption         =   "Receita Por CTC:"
               Height          =   195
               Left            =   2640
               TabIndex        =   199
               Top             =   2400
               Width           =   1245
            End
            Begin VB.Label lblTotReceitaporCTC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3960
               TabIndex        =   198
               Top             =   2400
               Width           =   855
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               Caption         =   "Com Valor Mínimo:"
               Height          =   195
               Left            =   -72360
               TabIndex        =   197
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               Caption         =   "Perc. Negociado:"
               Height          =   195
               Left            =   -72360
               TabIndex        =   196
               Top             =   840
               Width           =   1245
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               Caption         =   "Peso Transportado:"
               Height          =   195
               Left            =   -74880
               TabIndex        =   194
               Top             =   1200
               Width           =   1395
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "Quantidade CTCs:"
               Height          =   195
               Left            =   -74880
               TabIndex        =   192
               Top             =   840
               Width           =   1305
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "Cobrado/Pago %:"
               Height          =   195
               Left            =   -72360
               TabIndex        =   191
               Top             =   1560
               Width           =   1260
            End
            Begin VB.Label lblTotTranCobrPago 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -70920
               TabIndex        =   190
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label lblTotTranValMinimo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -70920
               TabIndex        =   189
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label lblTotTranPercNegoc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -70920
               TabIndex        =   188
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Volumes Transp:"
               Height          =   195
               Left            =   -74880
               TabIndex        =   186
               Top             =   1560
               Width           =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Frete Cobr/Valor:"
               Height          =   195
               Left            =   -74880
               TabIndex        =   185
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label lblTotTranPercFreteCobrValor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -73440
               TabIndex        =   184
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Frete Pago/Valor:"
               Height          =   195
               Left            =   -74880
               TabIndex        =   183
               Top             =   2400
               Width           =   1260
            End
            Begin VB.Label lblTotTranPercFretePagoValor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -73440
               TabIndex        =   182
               Top             =   2400
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Receita/Valor:"
               Height          =   195
               Left            =   -72360
               TabIndex        =   181
               Top             =   2040
               Width           =   1035
            End
            Begin VB.Label lblTotTranPercReceitaValor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -71040
               TabIndex        =   180
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Receita Por CTC:"
               Height          =   195
               Left            =   -72360
               TabIndex        =   179
               Top             =   2400
               Width           =   1245
            End
            Begin VB.Label lblTotTranReceitaporCTC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -71040
               TabIndex        =   178
               Top             =   2400
               Width           =   855
            End
            Begin VB.Label lblTransportador 
               AutoSize        =   -1  'True
               Caption         =   "Transportador"
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
               Left            =   -71880
               TabIndex        =   177
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label lblTotTranVol 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -73440
               TabIndex        =   187
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label lblTotTranPeso 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -73440
               TabIndex        =   195
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label lblTotTranCtc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -73440
               TabIndex        =   193
               Top             =   840
               Width           =   975
            End
            Begin VB.Label lblTotTranValMerc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   -73920
               TabIndex        =   226
               Top             =   480
               Width           =   1455
            End
         End
         Begin MSChart20Lib.MSChart grafValmerc 
            Height          =   2895
            Left            =   -69840
            OleObjectBlob   =   "frmAnOper.frx":08A0
            TabIndex        =   222
            Top             =   3360
            Visible         =   0   'False
            Width           =   3135
         End
         Begin MSChart20Lib.MSChart grafReceitaTran 
            Height          =   2895
            Left            =   -69840
            OleObjectBlob   =   "frmAnOper.frx":2E15
            TabIndex        =   223
            Top             =   480
            Visible         =   0   'False
            Width           =   3135
         End
         Begin MSChart20Lib.MSChart grafReceitaTotal 
            Height          =   2895
            Left            =   -66720
            OleObjectBlob   =   "frmAnOper.frx":49C7
            TabIndex        =   224
            Top             =   480
            Visible         =   0   'False
            Width           =   3135
         End
         Begin MSChart20Lib.MSChart grafReceita 
            Height          =   2895
            Left            =   -66720
            OleObjectBlob   =   "frmAnOper.frx":6585
            TabIndex        =   225
            Top             =   3360
            Visible         =   0   'False
            Width           =   3135
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Peso"
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
         Left            =   -65280
         TabIndex        =   142
         Top             =   840
         Width           =   1815
         Begin VB.Label lblNF4B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   151
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblNF4D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   150
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblNF4C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   149
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblNF4E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   148
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lblNF4F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   147
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label lblNF4G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   146
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label lblNF4H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   145
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label lblNF4Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   144
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label lblNF4A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Valor de Frete"
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
         Left            =   -69720
         TabIndex        =   132
         Top             =   840
         Width           =   2055
         Begin VB.Label lblNF3A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   141
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblNF3Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label lblNF3H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   139
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label lblNF3G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   138
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label lblNF3F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   137
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label lblNF3E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblNF3D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblNF3C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblNF3B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Valor de Mercadoria"
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
         Left            =   -67560
         TabIndex        =   122
         Top             =   840
         Width           =   2175
         Begin VB.Label lblNF2B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   131
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblNF2D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblNF2C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label lblNF2E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label lblNF2F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label lblNF2G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label lblNF2H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label lblNF2Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label lblNF2A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Quant. CTCs"
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
         Left            =   -72360
         TabIndex        =   103
         Top             =   840
         Width           =   2535
         Begin VB.Label lblNF1A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   121
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblNF1Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label lblNF1H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label lblNF1G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label lblNF1F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lblNF1E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblNF1C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblNF1D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblNF1B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblNFPerA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   112
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblNFPerTotal 
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
            Height          =   255
            Left            =   1560
            TabIndex        =   111
            Top             =   4320
            Width           =   735
         End
         Begin VB.Label lblNFPerH 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   110
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label lblNFPerG 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   109
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label lblNFPerF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   108
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label lblNFPerE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   107
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lblNFPerC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   106
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblNFPerD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   105
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblNFPerB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   104
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Faixa de Valor"
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
         Left            =   -74760
         TabIndex        =   93
         Top             =   840
         Width           =   2295
         Begin VB.Label Label107 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 100 a 500,00"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label106 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 1.001 a 2.500,00"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label105 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 501 a 1.000,00"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label104 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 2.501 a 5.000,00"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label103 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 5.000 a 10.000,00"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label102 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 10.000 a 20.000,00"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label101 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Acima de 20.000,00"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label99 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total"
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
            Left            =   240
            TabIndex        =   95
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label Label98 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abaixo de 100,00"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Peso"
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
         Left            =   -65280
         TabIndex        =   82
         Top             =   840
         Width           =   1815
         Begin VB.Label lblPeso4A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPeso4Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label lblPeso4I 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label lblPeso4H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label lblPeso4G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lblPeso4F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblPeso4E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblPeso4C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblPeso4D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblPeso4B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Valor de Frete"
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
         Left            =   -69720
         TabIndex        =   71
         Top             =   840
         Width           =   2055
         Begin VB.Label lblPeso3B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label lblPeso3D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblPeso3C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblPeso3E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblPeso3F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblPeso3G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblPeso3H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblPeso3I 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label lblPeso3Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   3960
            Width           =   1575
         End
         Begin VB.Label lblPeso3A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Valor de Mercadoria"
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
         Left            =   -67560
         TabIndex        =   60
         Top             =   840
         Width           =   2175
         Begin VB.Label lblPeso2A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblPeso2Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label lblPeso2I 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label lblPeso2H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label lblPeso2G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label lblPeso2F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lblPeso2E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblPeso2C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label lblPeso2D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblPeso2B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Quant. CTCs"
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
         Left            =   -72360
         TabIndex        =   39
         Top             =   840
         Width           =   2535
         Begin VB.Label lblper1B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   59
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblper1D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   58
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblper1C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   57
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblper1E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   56
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblper1F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   55
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblper1G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   54
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lblper1H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   53
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label lblper1I 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   52
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label lblPercTotal 
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
            Height          =   255
            Left            =   1560
            TabIndex        =   51
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label lblPer1A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   50
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblPeso1B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblPeso1D 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblPeso1C 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblPeso1E 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblPeso1F 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label lblPeso1G 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblPeso1H 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label lblPeso1I 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label lblPeso1Total 
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
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label lblPeso1A 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Faixa de Peso"
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
         Left            =   -74760
         TabIndex        =   28
         Top             =   840
         Width           =   2295
         Begin VB.Label Label46 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 0 a 20 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label44 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total"
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
            Left            =   240
            TabIndex        =   37
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label Label43 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Acima 5.000 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label42 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 2.501 a 5.000 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label Label41 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 1.001 a 2.500 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label Label40 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 501 a 1.000 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label39 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 201 a 500 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label38 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 51 a 100 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label37 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 101 a 200 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label36 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De 21 a 50 Kg"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   840
            Width           =   1815
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
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Label lblDataPer2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblDataPer1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1935
         TabIndex        =   6
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1935
         TabIndex        =   5
         Top             =   360
         Width           =   5625
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Período..............: De"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cliente / Remetente:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Modal:"
         Height          =   195
         Left            =   4920
         TabIndex        =   2
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblModal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAnEstat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Excel As Excel.Application
Dim ExcelWBk As Excel.Workbook
Dim ExcelWS1 As Excel.Worksheet
Dim ExcelWS2 As Excel.Worksheet
Dim ExcelWS3 As Excel.Worksheet
Private Sub cmdGeraXls_Click()
    
    Set Excel = CreateObject("Excel.Application")
    Set Excel = GetObject(, "Excel.Application")
    Excel.Visible = True
    Set ExcelWBk = Excel.Workbooks.Add
    Set ExcelWS1 = ExcelWBk.Worksheets(1)
    Set ExcelWS2 = ExcelWBk.Worksheets(2)
    Set ExcelWS3 = ExcelWBk.Worksheets(3)
    
    ExcelWS1.Cells.Font.Name = "Verdana"
    
    Excel.ActiveWindow.Zoom = 75
    
    ExcelWS1.Name = "Por UF"
    ExcelWS2.Name = "Por Peso"
    ExcelWS3.Name = "Por Val.Merc."
    
    ExcelWS1.Range("A:Z").Borders.ColorIndex = 2
    ExcelWS1.Range("A:Z").Borders(xlEdgeTop).Weight = xlThin
    ExcelWS1.Range("A:Z").Borders(xlEdgeRight).Weight = xlThin
    ExcelWS1.Range("A:Z").Borders(xlEdgeLeft).Weight = xlThin
    ExcelWS1.Range("A:Z").Borders(xlEdgeBottom).Weight = xlThin
    ExcelWS1.Range("A:Z").Borders(xlInsideHorizontal).Weight = xlThin
    ExcelWS1.Range("A:Z").Borders(xlInsideVertical).Weight = xlThin
    
    ExcelWS1.Cells(1, 1) = "ANÁLISE ESTATÍSTICA - POR UF"
    ExcelWS1.Cells(2, 1) = "Cliente: " & lblCliente.Caption
    ExcelWS1.Cells(3, 1) = "Modal: " & lblModal
    ExcelWS1.Cells(4, 1) = "Período: " & lblDataPer1 & " a " & lblDataPer2
    
    ExcelWS1.Range(ExcelWS1.Cells(1, 1), ExcelWS1.Cells(4, 1)).Font.Size = 12
    ExcelWS1.Range(ExcelWS1.Cells(1, 1), ExcelWS1.Cells(4, 1)).Font.Bold = True
    
    ExcelWS1.Cells(7, 1) = "UF"
    ExcelWS1.Cells(7, 2) = "Val. Merc."
    ExcelWS1.Cells(7, 3) = "Frete"
    ExcelWS1.Cells(7, 4) = "Peso"
    ExcelWS1.Cells(7, 5) = "Expedições"
    ExcelWS1.Cells(7, 6) = "Peso/Qtde"
    ExcelWS1.Cells(7, 7) = "Frete/Peso"
    ExcelWS1.Cells(7, 8) = "Frete/ValMerc"
        
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(7, 8)).Font.Bold = True
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(7, 8)).Interior.ColorIndex = 15

    
    For xlinha = 1 To 16
        ExcelWS1.Cells(xlinha + 7, 1) = FlexUF1.TextMatrix(xlinha, 0)
        ExcelWS1.Cells(xlinha + 7, 2) = SoNumeros(FlexValmerNONE.TextMatrix(xlinha, 0)) / 100
        ExcelWS1.Cells(xlinha + 7, 3) = SoNumeros(FlexFreteNONE.TextMatrix(xlinha, 0)) / 100
        ExcelWS1.Cells(xlinha + 7, 4) = SoNumeros(FlexPesoNONE.TextMatrix(xlinha, 0)) / 10
        ExcelWS1.Cells(xlinha + 7, 5) = SoNumeros(FlexExpedNONE.TextMatrix(xlinha, 0))
        ExcelWS1.Cells(xlinha + 7, 6) = SoNumeros(FlexPercentNONE.TextMatrix(xlinha, 0)) / 10
        ExcelWS1.Cells(xlinha + 7, 7) = SoNumeros(FlexPercentNONE.TextMatrix(xlinha, 1))
        ExcelWS1.Cells(xlinha + 7, 8) = (SoNumeros(FlexPercentNONE.TextMatrix(xlinha, 2)) / 1000) / 100
    Next
    
    For xlinha = 1 To 11
        ExcelWS1.Cells(xlinha + 23, 1) = FlexUF2.TextMatrix(xlinha, 0)
        ExcelWS1.Cells(xlinha + 23, 2) = SoNumeros(FlexValMerSDSU.TextMatrix(xlinha, 0)) / 100
        ExcelWS1.Cells(xlinha + 23, 3) = SoNumeros(FlexFreteSDSU.TextMatrix(xlinha, 0)) / 100
        ExcelWS1.Cells(xlinha + 23, 4) = SoNumeros(FlexPesoSDSU.TextMatrix(xlinha, 0)) / 10
        ExcelWS1.Cells(xlinha + 23, 5) = SoNumeros(FlexExpedSDSU.TextMatrix(xlinha, 0))
        ExcelWS1.Cells(xlinha + 23, 6) = SoNumeros(FlexPercentSDSU.TextMatrix(xlinha, 0)) / 10
        ExcelWS1.Cells(xlinha + 23, 7) = SoNumeros(FlexPercentSDSU.TextMatrix(xlinha, 1)) / 10
        ExcelWS1.Cells(xlinha + 23, 8) = (SoNumeros(FlexPercentSDSU.TextMatrix(xlinha, 2)) / 1000) / 100
    Next
    
        'totais
        
        ExcelWS1.Cells(35, 1) = "Total"
        ExcelWS1.Cells(35, 2) = SoNumeros(FlexValMerTot.TextMatrix(0, 0)) / 100
        ExcelWS1.Cells(35, 3) = SoNumeros(FlexFreteTot.TextMatrix(0, 0)) / 100
        ExcelWS1.Cells(35, 4) = SoNumeros(FlexPesoTot.TextMatrix(0, 0)) / 10
        ExcelWS1.Cells(35, 5) = SoNumeros(FlexExpedTot.TextMatrix(0, 0))
        ExcelWS1.Cells(35, 6) = SoNumeros(FlexPercenttot.TextMatrix(0, 0)) / 10
        ExcelWS1.Cells(35, 7) = SoNumeros(FlexPercenttot.TextMatrix(0, 1)) / 10
        ExcelWS1.Cells(35, 8) = (SoNumeros(FlexPercenttot.TextMatrix(0, 2)) / 1000) / 100
        
    ExcelWS1.Range(ExcelWS1.Cells(35, 1), ExcelWS1.Cells(35, 8)).Font.Bold = True
    
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(35, 8)).Borders.ColorIndex = 1
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(35, 8)).Borders(xlEdgeTop).Weight = xlThin
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(35, 8)).Borders(xlEdgeLeft).Weight = xlThin
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(35, 8)).Borders(xlEdgeRight).Weight = xlThin
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(35, 8)).Borders(xlEdgeBottom).Weight = xlThin
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(35, 8)).Borders(xlInsideVertical).Weight = xlThin
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(35, 8)).Borders(xlInsideHorizontal).Weight = xlThin
    
    
    ExcelWS1.Range(ExcelWS1.Cells(8, 2), ExcelWS1.Cells(35, 8)).Style = "Comma"
    
    'ValMerc / Frete
    ExcelWS1.Range(ExcelWS1.Cells(8, 2), ExcelWS1.Cells(35, 3)).NumberFormat = "###,###,###,##0.00"
    'peso 1 casa decimal
    ExcelWS1.Range(ExcelWS1.Cells(8, 4), ExcelWS1.Cells(35, 4)).NumberFormat = "###,###,##0.0"
    'expediçoes 0 casas decimais
    ExcelWS1.Range(ExcelWS1.Cells(8, 5), ExcelWS1.Cells(35, 5)).NumberFormat = "###,##0"
    'indices peso/qtde e frete/peso - 1 casa decimal
    ExcelWS1.Range(ExcelWS1.Cells(8, 6), ExcelWS1.Cells(35, 7)).NumberFormat = "###,###,##0.0"
    'indice frete/merc - 3 casas decimais percentual
    ExcelWS1.Range(ExcelWS1.Cells(8, 8), ExcelWS1.Cells(35, 8)).NumberFormat = "0.000%"
    
    ExcelWS1.Range("B:H").EntireColumn.AutoFit
    ExcelWS1.Range("A:A").EntireColumn.ColumnWidth = 10
    
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(7, 8)).HorizontalAlignment = xlCenter
    ExcelWS1.Range(ExcelWS1.Cells(7, 1), ExcelWS1.Cells(7, 8)).VerticalAlignment = xlCenter
    ExcelWS1.Range("7:7").EntireRow.RowHeight = 20
        
End Sub

Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdNova_Click()
    Unload frmAnEstat
    'Me.Hide
    frmEscCliPer.Caption = "Análise Estatística"
    frmEscCliPer.Show 1
    
End Sub

Private Sub cmdSair_Click()
    Unload frmEscCliPer
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub flexMutuoResumo_Click()
    Dim xTransportador As String, xFreteCobrado As Currency, xFretePagoSub As Currency, xreceita As Currency
    Dim xTotTranCtc As Long, xTotTranPeso As Currency, xTotTranVol As Long, xTotTranValMerc As Currency
            
        grafReceitaTran.Visible = True
        flexMutuoResumo.Col = 1
        xTransportador = flexMutuoResumo.Text
        de_informa.rsSel_AnEstatSubGrid.MoveFirst
        Do While de_informa.rsSel_AnEstatSubGrid.Fields("transp_sub") <> xTransportador
            de_informa.rsSel_AnEstatSubGrid.MoveNext
        Loop
        

        xFreteCobrado = de_informa.rsSel_AnEstatSubGrid.Fields("tfretetotal")
        xFretePagoSub = de_informa.rsSel_AnEstatSubGrid.Fields("tfretepago")
        xreceita = de_informa.rsSel_AnEstatSubGrid.Fields("receita")
        xTotTranCtc = de_informa.rsSel_AnEstatSubGrid.Fields("qtd")
        xTotTranPeso = de_informa.rsSel_AnEstatSubGrid.Fields("tpeso")
        xTotTranVol = de_informa.rsSel_AnEstatSubGrid.Fields("tvolumes")
        xTotTranValMerc = de_informa.rsSel_AnEstatSubGrid.Fields("tvalmerc")
        
        
        'flexMutuoResumo.Col = 2
        'xFreteCobrado = Val(flexMutuoResumo.Text)
        'flexMutuoResumo.Col = 3
        'xFretePagoSub = Val(flexMutuoResumo.Text)
        'flexMutuoResumo.Col = 4
        'xreceita = Val(flexMutuoResumo.Text)
        'flexMutuoResumo.Col = 5
        'xTotTranCtc = Val(flexMutuoResumo.Text)
        'flexMutuoResumo.Col = 6
        'xTotTranPeso = Val(flexMutuoResumo.Text)
        'flexMutuoResumo.Col = 7
        'xTotTranVol = Val(flexMutuoResumo.Text)
        'flexMutuoResumo.Col = 8
        'xTotTranValMerc = Val(flexMutuoResumo.Text)
       
        lblTransportador = xTransportador
        lblTotTranCtc = Format(xTotTranCtc, "###,##0")
        lblTotTranPeso = Format(xTotTranPeso, "###,##0.0")
        lblTotTranVol = Format(xTotTranVol, "###,##0")
        lblTotTranValMerc = Format(xTotTranValMerc, "#,###,###,##0.00")
        lblTotTranCobrPago = Format(xFretePagoSub / xFreteCobrado, "##0.00%")
        lblTotTranPercFreteCobrValor = Format(xFreteCobrado / xTotTranValMerc, "##0.000%")
        lblTotTranPercFretePagoValor = Format(xFretePagoSub / xTotTranValMerc, "##0.000%")
        lblTotTranPercReceitaValor = Format(xreceita / xTotTranValMerc, "##0.000%")
        lblTotTranReceitaporCTC = Format(xreceita / xTotTranCtc, "###,##0.00")

        frmAnEstat.grafReceitaTran.RowLabel = lblTransportador
        frmAnEstat.grafReceitaTran.Column = 1
        frmAnEstat.grafReceitaTran.Data = xreceita
        frmAnEstat.grafReceitaTran.ColumnLabel = "Receita"
        frmAnEstat.grafReceitaTran.Column = 2
        frmAnEstat.grafReceitaTran.Data = xFretePagoSub
        frmAnEstat.grafReceitaTran.ColumnLabel = "Fr.Pago"


        'busca o transportador
        
        If de_informa.rsSel_BuscaSubContratado.State = 1 Then de_informa.rsSel_BuscaSubContratado.Close
        de_informa.Sel_BuscaSubContratado lblTransportador
        
        lblTotTranPercNegoc = Format(de_informa.rsSel_BuscaSubContratado.Fields("percentual"), "##0.00%")
        lblTotTranValMinimo = Format(de_informa.rsSel_BuscaSubContratado.Fields("minimo"), "###,##0.00")
        
        
        
End Sub

Private Sub Form_Activate()
'    Unload frmEscCliPer
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    
    FlexUF1.Cols = 1
    FlexUF2.Cols = 1
    
    FlexValmerNONE.Cols = 1
    FlexFreteNONE.Cols = 1
    FlexPesoNONE.Cols = 1
    FlexVolNONE.Cols = 1
    FlexExpedNONE.Cols = 1
    FlexPercentNONE.Cols = 3
    
    FlexValMerSDSU.Cols = 1
    FlexFreteSDSU.Cols = 1
    FlexPesoSDSU.Cols = 1
    FlexVolSDSU.Cols = 1
    FlexExpedSDSU.Cols = 1
    FlexPercentSDSU.Cols = 3
    
    FlexValMerTot.Cols = 1
    FlexFreteTot.Cols = 1
    FlexPesoTot.Cols = 1
    FlexVolTot.Cols = 1
    FlexExpedTot.Cols = 1
    FlexPercenttot.Cols = 3
    
    FlexUF1.ColWidth(0) = 440
    FlexUF2.ColWidth(0) = 440
    
    FlexValmerNONE.ColWidth(0) = 1400
    FlexFreteNONE.ColWidth(0) = 1400
    FlexPesoNONE.ColWidth(0) = 910
    FlexExpedNONE.ColWidth(0) = 910
    FlexVolNONE.ColWidth(0) = 910
    FlexPercentNONE.ColWidth(0) = 948
    FlexPercentNONE.ColWidth(1) = 948
    FlexPercentNONE.ColWidth(2) = 948
    
    FlexValMerSDSU.ColWidth(0) = 1400
    FlexFreteSDSU.ColWidth(0) = 1400
    FlexPesoSDSU.ColWidth(0) = 910
    FlexExpedSDSU.ColWidth(0) = 910
    FlexVolSDSU.ColWidth(0) = 910
    FlexPercentSDSU.ColWidth(0) = 948
    FlexPercentSDSU.ColWidth(1) = 948
    FlexPercentSDSU.ColWidth(2) = 948
    
    FlexValMerTot.ColWidth(0) = 1400
    FlexFreteTot.ColWidth(0) = 1400
    FlexPesoTot.ColWidth(0) = 910
    FlexExpedTot.ColWidth(0) = 910
    FlexVolTot.ColWidth(0) = 910
    FlexPercenttot.ColWidth(0) = 948
    FlexPercenttot.ColWidth(1) = 948
    FlexPercenttot.ColWidth(2) = 948
    
    FlexUF1.Row = 0
    FlexUF1.Text = " UF"
    FlexUF2.Row = 0
    FlexUF2.Text = " UF"

    FlexValmerNONE.Row = 0
    FlexValmerNONE.Text = "            R$"
    FlexFreteNONE.Row = 0
    FlexFreteNONE.Text = "            R$"
    FlexPesoNONE.Row = 0
    FlexPesoNONE.Text = "      Kg"
    FlexVolNONE.Row = 0
    FlexVolNONE.Text = "   Qtde."
    FlexExpedNONE.Row = 0
    FlexExpedNONE.Text = "   Qtde."
    FlexPercentNONE.Row = 0
    FlexPercentNONE.Col = 0
    FlexPercentNONE.Text = "Peso/Qtde."
    FlexPercentNONE.Col = 1
    FlexPercentNONE.Text = "Frete/Peso"
    FlexPercentNONE.Col = 2
    FlexPercentNONE.Text = "Frete/Merc."
    
    FlexValMerSDSU.Row = 0
    FlexValMerSDSU.Text = "            R$"
    FlexFreteSDSU.Row = 0
    FlexFreteSDSU.Text = "            R$"
    FlexPesoSDSU.Row = 0
    FlexPesoSDSU.Text = "      Kg"
    FlexVolSDSU.Row = 0
    FlexVolSDSU.Text = "   Qtde."
    FlexExpedSDSU.Row = 0
    FlexExpedSDSU.Text = "   Qtde."
    FlexPercentSDSU.Row = 0
    FlexPercentSDSU.Col = 0
    FlexPercentSDSU.Text = "Peso/Qtde."
    FlexPercentSDSU.Col = 1
    FlexPercentSDSU.Text = "Frete/Peso"
    FlexPercentSDSU.Col = 2
    FlexPercentSDSU.Text = "Frete/Merc."
    
    'preenche os estados
    
    FlexUF1.Row = 1
    FlexUF1.Text = " AC"
    FlexUF1.Row = 2
    FlexUF1.Text = " AM"
    FlexUF1.Row = 3
    FlexUF1.Text = " AP"
    FlexUF1.Row = 4
    FlexUF1.Text = " PA"
    FlexUF1.Row = 5
    FlexUF1.Text = " RO"
    FlexUF1.Row = 6
    FlexUF1.Text = " RR"
    FlexUF1.Row = 7
    FlexUF1.Text = " TO"
    FlexUF1.Row = 8
    FlexUF1.Text = " AL"
    FlexUF1.Row = 9
    FlexUF1.Text = " BA"
    FlexUF1.Row = 10
    FlexUF1.Text = " SE"
    FlexUF1.Row = 11
    FlexUF1.Text = " PE"
    FlexUF1.Row = 12
    FlexUF1.Text = " PB"
    FlexUF1.Row = 13
    FlexUF1.Text = " RN"
    FlexUF1.Row = 14
    FlexUF1.Text = " CE"
    FlexUF1.Row = 15
    FlexUF1.Text = " PI"
    FlexUF1.Row = 16
    FlexUF1.Text = " MA"
    
    FlexUF2.Row = 1
    FlexUF2.Text = " ES"
    FlexUF2.Row = 2
    FlexUF2.Text = " MG"
    FlexUF2.Row = 3
    FlexUF2.Text = " RJ"
    FlexUF2.Row = 4
    FlexUF2.Text = " SP"
    FlexUF2.Row = 5
    FlexUF2.Text = " PR"
    FlexUF2.Row = 6
    FlexUF2.Text = " RS"
    FlexUF2.Row = 7
    FlexUF2.Text = " SC"
    FlexUF2.Row = 8
    FlexUF2.Text = " DF"
    FlexUF2.Row = 9
    FlexUF2.Text = " GO"
    FlexUF2.Row = 10
    FlexUF2.Text = " MS"
    FlexUF2.Row = 11
    FlexUF2.Text = " MT"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmAnEstat = Nothing
End Sub

Private Sub Label79_Click()

End Sub

