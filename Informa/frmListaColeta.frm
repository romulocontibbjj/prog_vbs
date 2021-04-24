VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmListaColeta 
   Caption         =   "Listagem de Status de Ordens de Coleta"
   ClientHeight    =   7815
   ClientLeft      =   -945
   ClientTop       =   495
   ClientWidth     =   12915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmListaColeta.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   12915
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ChkAtualiza 
      Caption         =   "Atualizar minha lista automaticamente"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9240
      TabIndex        =   34
      Top             =   1140
      Width           =   3075
   End
   Begin VB.Timer Tempo 
      Left            =   120
      Top             =   7320
   End
   Begin VB.CommandButton CmdGerarTXT 
      Caption         =   "Gerar Arq. TXT da Lista Abaixo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   29
      Top             =   240
      Width           =   3195
   End
   Begin VB.CommandButton CmdBaixar 
      Caption         =   "POD Coleta..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   28
      Top             =   660
      Width           =   3195
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar Coleta..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   27
      Top             =   1080
      Width           =   3195
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
      Height          =   1815
      Left            =   180
      TabIndex        =   9
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cbFiliais 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmListaColeta.frx":000C
         Left            =   480
         List            =   "frmListaColeta.frx":000E
         TabIndex        =   39
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Frame fraPorEmissao 
         Height          =   975
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton opt30d 
            Caption         =   "Últimos 30 dias"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optPer15d 
            Caption         =   "Últimos 15 dias"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton opt60d 
            Caption         =   "Últimos 60 dias"
            Height          =   435
            Left            =   1920
            TabIndex        =   11
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame fraPorPeriodo 
         Height          =   855
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   1680
            TabIndex        =   15
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1440
            TabIndex        =   18
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "( Intervalo Máximo de 60 dias )"
            Height          =   195
            Left            =   390
            TabIndex        =   17
            Top             =   120
            Width           =   2160
         End
      End
      Begin VB.Frame fraPorMesAno 
         Height          =   855
         Left            =   1680
         TabIndex        =   22
         Top             =   300
         Visible         =   0   'False
         Width           =   3015
         Begin VB.ComboBox comboMesAnoAcomp 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   480
            TabIndex        =   23
            Text            =   "Mes/Ano"
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.OptionButton optPorEmissao 
         Caption         =   "Emissão nos .."
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optPorPeriodo 
         Caption         =   "Por Período .."
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optPorMes 
         Caption         =   "Por Mês .."
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Filial"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5295
      Left            =   180
      TabIndex        =   8
      Top             =   2040
      Width           =   12555
      Begin TabDlg.SSTab SSTab1 
         Height          =   4875
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   8599
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Pendentes"
         TabPicture(0)   =   "frmListaColeta.frx":0010
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblRegistros(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGridMov(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Em Ocorrência"
         TabPicture(1)   =   "frmListaColeta.frx":002C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DataGridMov(1)"
         Tab(1).Control(1)=   "LblRegistros(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Efetuadas"
         TabPicture(2)   =   "frmListaColeta.frx":0048
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "DataGridMov(2)"
         Tab(2).Control(1)=   "LblRegistros(2)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Baixadas"
         TabPicture(3)   =   "frmListaColeta.frx":0064
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "DataGridMov(3)"
         Tab(3).Control(1)=   "LblRegistros(3)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Canceladas"
         TabPicture(4)   =   "frmListaColeta.frx":0080
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "DataGridMov(4)"
         Tab(4).Control(1)=   "LblRegistros(4)"
         Tab(4).ControlCount=   2
         Begin MSDataGridLib.DataGrid DataGridMov 
            Height          =   3975
            Index           =   0
            Left            =   180
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   7011
            _Version        =   393216
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridMov 
            Height          =   3975
            Index           =   1
            Left            =   -74820
            TabIndex        =   30
            Top             =   720
            Visible         =   0   'False
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   7011
            _Version        =   393216
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridMov 
            Height          =   3975
            Index           =   2
            Left            =   -74820
            TabIndex        =   31
            Top             =   720
            Visible         =   0   'False
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   7011
            _Version        =   393216
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridMov 
            Height          =   3975
            Index           =   3
            Left            =   -74820
            TabIndex        =   35
            Top             =   720
            Visible         =   0   'False
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   7011
            _Version        =   393216
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridMov 
            Height          =   3975
            Index           =   4
            Left            =   -74820
            TabIndex        =   36
            Top             =   720
            Visible         =   0   'False
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   7011
            _Version        =   393216
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label LblRegistros 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   -63780
            TabIndex        =   38
            Top             =   480
            Width           =   75
         End
         Begin VB.Label LblRegistros 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   -63780
            TabIndex        =   37
            Top             =   480
            Width           =   75
         End
         Begin VB.Label LblRegistros 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   -63795
            TabIndex        =   33
            Top             =   480
            Width           =   75
         End
         Begin VB.Label LblRegistros 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   -63840
            TabIndex        =   32
            Top             =   480
            Width           =   75
         End
         Begin VB.Label LblRegistros 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   11205
            TabIndex        =   26
            Top             =   480
            Width           =   75
         End
      End
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   660
      Width           =   3195
   End
   Begin VB.CommandButton CmdProcessar 
      Caption         =   "Processar"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   240
      Width           =   3195
   End
   Begin VB.Frame Frame2 
      Caption         =   "Coletas..."
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
      Left            =   7200
      TabIndex        =   0
      Top             =   5280
      Width           =   4575
      Begin VB.OptionButton Option11 
         Caption         =   "Todas"
         Height          =   195
         Left            =   2040
         TabIndex        =   7
         Top             =   660
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Confirmadas/Atendidas"
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Em Trânsito"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Canceladas"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Pendentes"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmListaColeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xIndex As Integer
Dim I As Integer
Dim XstrTodos As String
Dim xDataInicial As Date
Dim xDataFinal As Date
Dim zPND As Integer
Dim zEMO As Integer
Dim zFIN As Integer
Dim zBAI As Integer
Dim zCAN As Integer
Dim xPND As Integer
Dim xEMO As Integer
Dim xFIN As Integer
Dim xBAI As Integer
Dim xCAN As Integer
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub CK_todasFil_Click()
If troca <> "OK" Then
    cbFiliais.Text = ""
End If
troca = "Nao"
End Sub

Private Sub CmdBaixar_Click()
    If Mid$(xdireitos, 33, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
    DataGridMov(SSTab1.Tab).Col = 0
    frmPodColeta.TxtFilial.Text = Mid(DataGridMov(SSTab1.Tab).Text, 1, 2)
    frmPodColeta.TxtColeta.Text = Mid(DataGridMov(SSTab1.Tab).Text, 3)
    Tempo.Interval = 0
    frmPodColeta.Show 1
    CmdProcessar_Click
    End If
End Sub

Private Sub CmdConsultar_Click()
DataGridMov(SSTab1.Tab).Col = 0
frmConsultaColeta.TxtFilial.Text = Mid(DataGridMov(SSTab1.Tab).Text, 1, 2)
frmConsultaColeta.TxtColeta.Text = Mid(DataGridMov(SSTab1.Tab).Text, 3)
Tempo.Interval = 0
frmConsultaColeta.Show 1
CmdProcessar_Click
End Sub

Private Sub CmdProcessar_Click()
If cbFiliais.Text = "TODAS" Or cbFiliais.Text = "" Then

    Dim StrTodos As String
    For I = 0 To cbFiliais.ListCount
        If I <> 0 And I <> cbFiliais.ListCount Then
            StrTodos = StrTodos & Mid(cbFiliais.List(I), 1, 2) & ","
        End If
    Next
    
    XstrTodos = Mid(StrTodos, 1, Len(StrTodos) - 1)
    
    CmdGerarTXT.Enabled = False
    CmdBaixar.Enabled = False
    CmdConsultar.Enabled = False
    CmdProcessar.Enabled = False
    cmdSair.Enabled = False
    Me.MousePointer = 11
    DoEvents
    
    SSTab1.Enabled = True
        If optPorEmissao.Value = True Then  'por emissao
            If optPer15d.Value = True Then
                xDataInicial = datahora("data") - 15
                xDataFinal = datahora("data")
            ElseIf opt30d.Value = True Then
                xDataInicial = datahora("data") - 30
                xDataFinal = datahora("data")
            ElseIf opt60d.Value = True Then
                xDataInicial = datahora("data") - 60
                xDataFinal = datahora("data")
            Else
                MsgBox "Período Escolhido Inválido !"
                Exit Sub
            End If
        End If
        If optPorMes.Value = True Then   'por mes
            xDataInicial = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 "01")
            xDataFinal = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 UltDiaMes(Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2)), _
                           Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4))))
            If CDate(xDataFinal) > CDate(datahora("DATA")) Then xDataFinal = datahora("DATA")
        End If
    
        If optPorPeriodo.Value = True Then   'por periodo
            If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
                MsgBox "Período Escolhido Inválido !"
                mskPer1.SetFocus
                Exit Sub
            End If
        
            If CDate(mskPer1) > CDate(mskPer2) Then
                MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
                mskPer1.SetFocus
                Exit Sub
            End If
        
            xDataInicial = CDate(mskPer1)
            xDataFinal = CDate(mskPer2)
        
            If xDataFinal - xDataInicial > 62 Then
                MsgBox "Período Escolhido Maior que 60 Dias ! Escolha um Período Menor."
                mskPer1.SetFocus
                Exit Sub
            End If
        End If

        xIndex = 0
        If de_informa.rsColetaSelPendente.State = 1 Then de_informa.rsColetaSelPendente.Close
            de_informa.ColetaSelPendente CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelPendente"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelPendente.RecordCount
            SSTab1.TabCaption(xIndex) = "Pendentes (" & de_informa.rsColetaSelPendente.RecordCount & ")"
            DoEvents
        
        xIndex = 1
        If de_informa.rsColetaSelEmOcorrencia.State = 1 Then de_informa.rsColetaSelEmOcorrencia.Close
            de_informa.ColetaSelEmOcorrencia CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelEmOcorrencia"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelEmOcorrencia.RecordCount
            SSTab1.TabCaption(xIndex) = "Em Ocorrência (" & de_informa.rsColetaSelEmOcorrencia.RecordCount & ")"
            DoEvents

        xIndex = 2
        If de_informa.rsColetaSelFinalizada.State = 1 Then de_informa.rsColetaSelFinalizada.Close
            de_informa.ColetaSelFinalizada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelFinalizada"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelFinalizada.RecordCount
            SSTab1.TabCaption(xIndex) = "Efetuadas (" & de_informa.rsColetaSelFinalizada.RecordCount & ")"
            DoEvents

        xIndex = 3
        If de_informa.rsColetaSelBaixada.State = 1 Then de_informa.rsColetaSelBaixada.Close
            de_informa.ColetaSelBaixada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelbaixada"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelBaixada.RecordCount
            SSTab1.TabCaption(xIndex) = "Baixadas (" & de_informa.rsColetaSelBaixada.RecordCount & ")"
            DoEvents

        xIndex = 4
        If de_informa.rsColetaSelCancelada.State = 1 Then de_informa.rsColetaSelCancelada.Close
            de_informa.ColetaSelCancelada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelcancelada"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelCancelada.RecordCount
            SSTab1.TabCaption(xIndex) = "Canceladas (" & de_informa.rsColetaSelCancelada.RecordCount & ")"
            DoEvents

        If de_informa.rsColetaSelPendente.RecordCount > 0 Then
            DataGridMov(0).Enabled = True
        Else
            DataGridMov(0).Enabled = False
        End If
    
        If de_informa.rsColetaSelEmOcorrencia.RecordCount > 0 Then
            DataGridMov(1).Enabled = True
        Else
            DataGridMov(1).Enabled = False
        End If
    
        If de_informa.rsColetaSelFinalizada.RecordCount > 0 Then
            DataGridMov(2).Enabled = True
        Else
            DataGridMov(2).Enabled = False
        End If
    
        If de_informa.rsColetaSelCancelada.RecordCount > 0 Then
            DataGridMov(3).Enabled = True
        Else
            DataGridMov(3).Enabled = False
        End If
    
        If de_informa.rsColetaSelBaixada.RecordCount > 0 Then
            DataGridMov(4).Enabled = True
        Else
            DataGridMov(4).Enabled = False
        End If

        ChkAtualiza.Enabled = True
        ChkAtualiza.Value = 1
        Tempo.Interval = 15000
        

        DataGridMov(0).Visible = True
        DataGridMov(1).Visible = True
        DataGridMov(2).Visible = True
        DataGridMov(3).Visible = True
        DataGridMov(4).Visible = True

        CmdGerarTXT.Enabled = False
        CmdProcessar.Enabled = True
        cmdSair.Enabled = True
        Me.MousePointer = 0
        DoEvents
Else
    
    'MsgBox ("Filial") & cbFiliais.Text
    XstrTodos = Mid(cbFiliais.Text, 1, 2)
    ChkAtualiza.Enabled = True
    ChkAtualiza.Value = 1
    Tempo.Interval = 15000
        
    CmdGerarTXT.Enabled = False
    CmdBaixar.Enabled = False
    CmdConsultar.Enabled = False
    CmdProcessar.Enabled = False
    cmdSair.Enabled = False
    Me.MousePointer = 11
    DoEvents
    
    SSTab1.Enabled = True
        If optPorEmissao.Value = True Then  'por emissao
            If optPer15d.Value = True Then
                xDataInicial = datahora("data") - 15
                xDataFinal = datahora("data")
            ElseIf opt30d.Value = True Then
                xDataInicial = datahora("data") - 30
                xDataFinal = datahora("data")
            ElseIf opt60d.Value = True Then
                xDataInicial = datahora("data") - 60
                xDataFinal = datahora("data")
            Else
                MsgBox "Período Escolhido Inválido !"
                Exit Sub
            End If
        End If
        If optPorMes.Value = True Then   'por mes
            xDataInicial = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 "01")
            xDataFinal = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 UltDiaMes(Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2)), _
                           Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4))))
            If CDate(xDataFinal) > CDate(datahora("DATA")) Then xDataFinal = datahora("DATA")
        End If
    
        If optPorPeriodo.Value = True Then   'por periodo
            If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
                MsgBox "Período Escolhido Inválido !"
                mskPer1.SetFocus
                Exit Sub
            End If
        
            If CDate(mskPer1) > CDate(mskPer2) Then
                MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
                mskPer1.SetFocus
                Exit Sub
            End If
        
            xDataInicial = CDate(mskPer1)
            xDataFinal = CDate(mskPer2)
        
            If xDataFinal - xDataInicial > 62 Then
                MsgBox "Período Escolhido Maior que 60 Dias ! Escolha um Período Menor."
                mskPer1.SetFocus
                Exit Sub
            End If
        End If

        xIndex = 0
        If de_informa.rsColetaSelPendente.State = 1 Then de_informa.rsColetaSelPendente.Close
            de_informa.ColetaSelPendente CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelPendente"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelPendente.RecordCount
            SSTab1.TabCaption(xIndex) = "Pendentes (" & de_informa.rsColetaSelPendente.RecordCount & ")"
            DoEvents
        
        xIndex = 1
        If de_informa.rsColetaSelEmOcorrencia.State = 1 Then de_informa.rsColetaSelEmOcorrencia.Close
            de_informa.ColetaSelEmOcorrencia CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelEmOcorrencia"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelEmOcorrencia.RecordCount
            SSTab1.TabCaption(xIndex) = "Em Ocorrência (" & de_informa.rsColetaSelEmOcorrencia.RecordCount & ")"
            DoEvents

        xIndex = 2
        If de_informa.rsColetaSelFinalizada.State = 1 Then de_informa.rsColetaSelFinalizada.Close
            de_informa.ColetaSelFinalizada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelFinalizada"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelFinalizada.RecordCount
            SSTab1.TabCaption(xIndex) = "Efetuadas (" & de_informa.rsColetaSelFinalizada.RecordCount & ")"
            DoEvents

        xIndex = 3
        If de_informa.rsColetaSelBaixada.State = 1 Then de_informa.rsColetaSelBaixada.Close
            de_informa.ColetaSelBaixada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelbaixada"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelBaixada.RecordCount
            SSTab1.TabCaption(xIndex) = "Baixadas (" & de_informa.rsColetaSelBaixada.RecordCount & ")"
            DoEvents

        xIndex = 4
        If de_informa.rsColetaSelCancelada.State = 1 Then de_informa.rsColetaSelCancelada.Close
            de_informa.ColetaSelCancelada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Set DataGridMov(xIndex).DataSource = de_informa
            DataGridMov(xIndex).DataMember = "ColetaSelcancelada"
            DataGridMov(xIndex).Refresh
            LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelCancelada.RecordCount
            SSTab1.TabCaption(xIndex) = "Canceladas (" & de_informa.rsColetaSelCancelada.RecordCount & ")"
            DoEvents

        If de_informa.rsColetaSelPendente.RecordCount > 0 Then
            DataGridMov(0).Enabled = True
        Else
            DataGridMov(0).Enabled = False
        End If
    
        If de_informa.rsColetaSelEmOcorrencia.RecordCount > 0 Then
            DataGridMov(1).Enabled = True
        Else
            DataGridMov(1).Enabled = False
        End If
    
        If de_informa.rsColetaSelFinalizada.RecordCount > 0 Then
            DataGridMov(2).Enabled = True
        Else
            DataGridMov(2).Enabled = False
        End If
    
        If de_informa.rsColetaSelCancelada.RecordCount > 0 Then
            DataGridMov(3).Enabled = True
        Else
            DataGridMov(3).Enabled = False
        End If
    
        If de_informa.rsColetaSelBaixada.RecordCount > 0 Then
            DataGridMov(4).Enabled = True
        Else
            DataGridMov(4).Enabled = False
        End If

        ChkAtualiza.Enabled = True
        ChkAtualiza.Value = 1
        Tempo.Interval = 15000
               

        DataGridMov(0).Visible = True
        DataGridMov(1).Visible = True
        DataGridMov(2).Visible = True
        DataGridMov(3).Visible = True
        DataGridMov(4).Visible = True

        CmdGerarTXT.Enabled = False
        CmdProcessar.Enabled = True
        cmdSair.Enabled = True
        Me.MousePointer = 0
        DoEvents
    
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGridMov_Click(Index As Integer)
CmdGerarTXT.Enabled = False
CmdBaixar.Enabled = True
CmdConsultar.Enabled = True
DoEvents
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    Call combomesano(comboMesAnoAcomp)
    comboMesAnoAcomp.ListIndex = 0
If de_informa.rsSel_UsuarioFiliais.State = 1 Then de_informa.rsSel_UsuarioFiliais.Close
        de_informa.sel_Usuariofiliais
        If de_informa.rsSel_UsuarioFiliais.RecordCount <= 0 Then
            MsgBox ("Nenhuma Filial Cadastrada... Verificar!"), vbInformation, "Busca Filial"
        Else
        cbFiliais.AddItem ("TODAS"), 0
        Do Until de_informa.rsSel_UsuarioFiliais.EOF
            cbFiliais.AddItem de_informa.rsSel_UsuarioFiliais.Fields("filial") & " -  " & de_informa.rsSel_UsuarioFiliais.Fields("nomefilial")
            de_informa.rsSel_UsuarioFiliais.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
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

Private Sub optPorFilial_Click()
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

Private Sub Tempo_Timer()
If ChkAtualiza.Value = 1 Then
    If cbFiliais.Text = "TODAS" Or cbFiliais.Text = "" Then
        For I = 0 To cbFiliais.ListCount
            If I <> 0 And I <> cbFiliais.ListCount Then
                StrTodos = StrTodos & Mid(cbFiliais.List(I), 1, 2) & ","
            End If
        Next
       XstrTodos = Mid(StrTodos, 1, Len(StrTodos) - 1)
        If optPorEmissao.Value = True Then  'por emissao
            If optPer15d.Value = True Then
                xDataInicial = datahora("data") - 15
                xDataFinal = datahora("data")
            ElseIf opt30d.Value = True Then
                xDataInicial = datahora("data") - 30
                xDataFinal = datahora("data")
            ElseIf opt60d.Value = True Then
                xDataInicial = datahora("data") - 60
                xDataFinal = datahora("data")
            Else
                MsgBox "Período Escolhido Inválido !"
                Exit Sub
            End If
        End If
        If optPorMes.Value = True Then   'por mes
            xDataInicial = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                     Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                     "01")
            xDataFinal = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                     Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                     UltDiaMes(Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2)), _
                               Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4))))
                               
            If CDate(xDataFinal) > CDate(datahora("DATA")) Then xDataFinal = datahora("DATA")
        End If
        If optPorPeriodo.Value = True Then   'por periodo
            If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
                MsgBox "Período Escolhido Inválido !"
                mskPer1.SetFocus
                Exit Sub
            End If
            If CDate(mskPer1) > CDate(mskPer2) Then
                MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
                mskPer1.SetFocus
                Exit Sub
            End If
            xDataInicial = CDate(mskPer1)
            xDataFinal = CDate(mskPer2)
            If xDataFinal - xDataInicial > 62 Then
                MsgBox "Período Escolhido Maior que 60 Dias ! Escolha um Período Menor."
                mskPer1.SetFocus
                Exit Sub
            End If
        End If
        'Dim PosWin As Long
        'PosWin = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        Dim zPND As Integer
        Dim zEMO As Integer
        Dim zFIN As Integer
        Dim zBAI As Integer
        Dim zCAN As Integer
        Dim xPND As Integer
        Dim xEMO As Integer
        Dim xFIN As Integer
        Dim xBAI As Integer
        Dim xCAN As Integer

        If de_informa.rsColetaSelPendente.State <> 1 Then
            de_informa.ColetaSelPendente CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zPND = de_informa.rsColetaSelPendente.RecordCount
        Else
            zPND = de_informa.rsColetaSelPendente.RecordCount
        End If
        
        If de_informa.rsColetaSelEmOcorrencia.State <> 1 Then
            de_informa.ColetaSelEmOcorrencia CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zEMO = de_informa.rsColetaSelEmOcorrencia.RecordCount
        Else
            zEMO = de_informa.rsColetaSelEmOcorrencia.RecordCount
        End If
        
        If de_informa.rsColetaSelFinalizada.State <> 1 Then
            de_informa.ColetaSelFinalizada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zFIN = de_informa.rsColetaSelFinalizada.RecordCount
        Else
            zFIN = de_informa.rsColetaSelFinalizada.RecordCount
        End If
        
        If de_informa.rsColetaSelBaixada.EOF Then
        zBAI = 0
        Else
            If de_informa.rsColetaSelBaixada.Status <> 1 Then
                de_informa.ColetaSelBaixada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Else
                zBAI = de_informa.rsColetaSelBaixada.RecordCount
            End If
            zBAI = de_informa.rsColetaSelBaixada.RecordCount
        End If
        
        If de_informa.rsColetaSelCancelada.State <> 1 Then
            de_informa.ColetaSelCancelada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zCAN = de_informa.rsColetaSelCancelada.RecordCount
        Else
            zCAN = de_informa.rsColetaSelCancelada.RecordCount
        End If
        
        xIndex = 0
        If de_informa.rsColetaSelPendente.State = 1 Then de_informa.rsColetaSelPendente.Close
        de_informa.ColetaSelPendente CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xPND = de_informa.rsColetaSelPendente.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelPendente"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelPendente.RecordCount
        SSTab1.TabCaption(xIndex) = "Pendentes (" & de_informa.rsColetaSelPendente.RecordCount & ")"
        DoEvents
        xIndex = 1
        If de_informa.rsColetaSelEmOcorrencia.State = 1 Then de_informa.rsColetaSelEmOcorrencia.Close
        de_informa.ColetaSelEmOcorrencia CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xEMO = de_informa.rsColetaSelEmOcorrencia.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelEmOcorrencia"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelEmOcorrencia.RecordCount
        SSTab1.TabCaption(xIndex) = "Em Ocorrência (" & de_informa.rsColetaSelEmOcorrencia.RecordCount & ")"
        DoEvents
        xIndex = 2
        If de_informa.rsColetaSelFinalizada.State = 1 Then de_informa.rsColetaSelFinalizada.Close
        de_informa.ColetaSelFinalizada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xFIN = de_informa.rsColetaSelFinalizada.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelFinalizada"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelFinalizada.RecordCount
        SSTab1.TabCaption(xIndex) = "Efetuadas (" & de_informa.rsColetaSelFinalizada.RecordCount & ")"
        DoEvents
        xIndex = 3
        If de_informa.rsColetaSelBaixada.State = 1 Then de_informa.rsColetaSelBaixada.Close
        de_informa.ColetaSelBaixada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xBAI = de_informa.rsColetaSelBaixada.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelbaixada"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelBaixada.RecordCount
        SSTab1.TabCaption(xIndex) = "Baixadas (" & de_informa.rsColetaSelBaixada.RecordCount & ")"
        DoEvents
        xIndex = 4
        If de_informa.rsColetaSelCancelada.State = 1 Then de_informa.rsColetaSelCancelada.Close
        de_informa.ColetaSelCancelada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xCAN = de_informa.rsColetaSelCancelada.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelcancelada"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelCancelada.RecordCount
        SSTab1.TabCaption(xIndex) = "Canceladas (" & de_informa.rsColetaSelCancelada.RecordCount & ")"
        DoEvents
        'End If
    
        If (zPND <> xPND) Or (zEMO <> xEMO) Or (zFIN <> xFIN) Or (zBAI <> xBAI) Or (zCAN <> xCAN) Then
        mdiInforma.WindowState = vbMaximized
        mdiInforma.SetFocus
        End If
        If de_informa.rsColetaSelPendente.RecordCount > 0 Then
            DataGridMov(0).Enabled = True
        Else
            DataGridMov(0).Enabled = False
            de_informa.rsColetaSelPendente.Close
        End If
        If de_informa.rsColetaSelEmOcorrencia.RecordCount > 0 Then
            DataGridMov(1).Enabled = True
        Else
            DataGridMov(1).Enabled = False
        End If
        If de_informa.rsColetaSelFinalizada.RecordCount > 0 Then
            DataGridMov(2).Enabled = True
        Else
            DataGridMov(2).Enabled = False
        End If
        If de_informa.rsColetaSelCancelada.RecordCount > 0 Then
            DataGridMov(3).Enabled = True
        Else
            DataGridMov(3).Enabled = False
        End If
        If de_informa.rsColetaSelBaixada.RecordCount > 0 Then
            DataGridMov(4).Enabled = True
        Else
            DataGridMov(4).Enabled = False
        End If
        DoEvents
    
    Else
        
        XstrTodos = Mid(cbFiliais.Text, 1, 2)
        If optPorEmissao.Value = True Then  'por emissao
            If optPer15d.Value = True Then
                xDataInicial = datahora("data") - 15
                xDataFinal = datahora("data")
            ElseIf opt30d.Value = True Then
                xDataInicial = datahora("data") - 30
                xDataFinal = datahora("data")
            ElseIf opt60d.Value = True Then
                xDataInicial = datahora("data") - 60
                xDataFinal = datahora("data")
            Else
                MsgBox "Período Escolhido Inválido !"
                Exit Sub
            End If
        End If
        If optPorMes.Value = True Then   'por mes
            xDataInicial = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                     Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                     "01")
            xDataFinal = CDate(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                     Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                     UltDiaMes(Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 5, 2)), _
                               Val(Mid$(comboMesAnoAcomp.ItemData(comboMesAnoAcomp.ListIndex), 1, 4))))
                               
            If CDate(xDataFinal) > CDate(datahora("DATA")) Then xDataFinal = datahora("DATA")
        End If
        If optPorPeriodo.Value = True Then   'por periodo
            If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
                MsgBox "Período Escolhido Inválido !"
                mskPer1.SetFocus
                Exit Sub
            End If
            If CDate(mskPer1) > CDate(mskPer2) Then
                MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
                mskPer1.SetFocus
                Exit Sub
            End If
            xDataInicial = CDate(mskPer1)
            xDataFinal = CDate(mskPer2)
            If xDataFinal - xDataInicial > 62 Then
                MsgBox "Período Escolhido Maior que 60 Dias ! Escolha um Período Menor."
                mskPer1.SetFocus
                Exit Sub
            End If
        End If
        'Dim PosWin As Long
        'PosWin = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        'Dim zPND As Integer
        'Dim zEMO As Integer
        'Dim zFIN As Integer
        'Dim zBAI As Integer
        'Dim zCAN As Integer
        'Dim xPND As Integer
        'Dim xEMO As Integer
        'Dim xFIN As Integer
        'Dim xBAI As Integer
        'Dim xCAN As Integer
        If de_informa.rsColetaSelPendente.State <> 1 Then
            de_informa.ColetaSelPendente CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zPND = de_informa.rsColetaSelPendente.RecordCount
        Else
            zPND = de_informa.rsColetaSelPendente.RecordCount
        End If
        
        If de_informa.rsColetaSelEmOcorrencia.State <> 1 Then
            de_informa.ColetaSelEmOcorrencia CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zEMO = de_informa.rsColetaSelEmOcorrencia.RecordCount
        Else
            zEMO = de_informa.rsColetaSelEmOcorrencia.RecordCount
        End If
        
        If de_informa.rsColetaSelFinalizada.State <> 1 Then
            de_informa.ColetaSelFinalizada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zFIN = de_informa.rsColetaSelFinalizada.RecordCount
        Else
            zFIN = de_informa.rsColetaSelFinalizada.RecordCount
        End If
        
        If de_informa.rsColetaSelBaixada.EOF Then
        zBAI = 0
        Else
            If de_informa.rsColetaSelBaixada.Status <> 1 Then
                de_informa.ColetaSelBaixada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            Else
                zBAI = de_informa.rsColetaSelBaixada.RecordCount
            End If
            zBAI = de_informa.rsColetaSelBaixada.RecordCount
        End If
        
        If de_informa.rsColetaSelCancelada.State <> 1 Then
            de_informa.ColetaSelCancelada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
            zCAN = de_informa.rsColetaSelCancelada.RecordCount
        Else
            zCAN = de_informa.rsColetaSelCancelada.RecordCount
        End If
        
        'zPND = de_informa.rsColetaSelPendente.RecordCount
        'zEMO = de_informa.rsColetaSelEmOcorrencia.RecordCount
        'zFIN = de_informa.rsColetaSelFinalizada.RecordCount
        'zBAI = de_informa.rsColetaSelBaixada.RecordCount
        'zCAN = de_informa.rsColetaSelCancelada.RecordCount
        
        xIndex = 0
        If de_informa.rsColetaSelPendente.State = 1 Then de_informa.rsColetaSelPendente.Close
        de_informa.ColetaSelPendente CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xPND = de_informa.rsColetaSelPendente.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelPendente"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelPendente.RecordCount
        SSTab1.TabCaption(xIndex) = "Pendentes (" & de_informa.rsColetaSelPendente.RecordCount & ")"
        DoEvents
        xIndex = 1
        If de_informa.rsColetaSelEmOcorrencia.State = 1 Then de_informa.rsColetaSelEmOcorrencia.Close
        de_informa.ColetaSelEmOcorrencia CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xEMO = de_informa.rsColetaSelEmOcorrencia.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelEmOcorrencia"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelEmOcorrencia.RecordCount
        SSTab1.TabCaption(xIndex) = "Em Ocorrência (" & de_informa.rsColetaSelEmOcorrencia.RecordCount & ")"
        DoEvents
        xIndex = 2
        If de_informa.rsColetaSelFinalizada.State = 1 Then de_informa.rsColetaSelFinalizada.Close
        de_informa.ColetaSelFinalizada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xFIN = de_informa.rsColetaSelFinalizada.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelFinalizada"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelFinalizada.RecordCount
        SSTab1.TabCaption(xIndex) = "Efetuadas (" & de_informa.rsColetaSelFinalizada.RecordCount & ")"
        DoEvents
        xIndex = 3
        If de_informa.rsColetaSelBaixada.State = 1 Then de_informa.rsColetaSelBaixada.Close
        de_informa.ColetaSelBaixada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xBAI = de_informa.rsColetaSelBaixada.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelbaixada"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelBaixada.RecordCount
        SSTab1.TabCaption(xIndex) = "Baixadas (" & de_informa.rsColetaSelBaixada.RecordCount & ")"
        DoEvents
        xIndex = 4
        If de_informa.rsColetaSelCancelada.State = 1 Then de_informa.rsColetaSelCancelada.Close
        de_informa.ColetaSelCancelada CDate(xDataInicial), CDate(xDataFinal), XstrTodos
        xCAN = de_informa.rsColetaSelCancelada.RecordCount
        Set DataGridMov(xIndex).DataSource = de_informa
        DataGridMov(xIndex).DataMember = "ColetaSelcancelada"
        DataGridMov(xIndex).Refresh
        LblRegistros(xIndex).Caption = "Registros Retornados: " & de_informa.rsColetaSelCancelada.RecordCount
        SSTab1.TabCaption(xIndex) = "Canceladas (" & de_informa.rsColetaSelCancelada.RecordCount & ")"
        DoEvents
        'End If
        If (zPND <> xPND) Or (zEMO <> xEMO) Or (zFIN <> xFIN) Or (zBAI <> xBAI) Or (zCAN <> xCAN) Then
        mdiInforma.WindowState = vbMaximized
        mdiInforma.SetFocus
        End If
            If de_informa.rsColetaSelPendente.RecordCount > 0 Then
            DataGridMov(0).Enabled = True
        Else
            DataGridMov(0).Enabled = False
            de_informa.rsColetaSelPendente.Close
        End If
        If de_informa.rsColetaSelEmOcorrencia.RecordCount > 0 Then
            DataGridMov(1).Enabled = True
        Else
            DataGridMov(1).Enabled = False
        End If
        If de_informa.rsColetaSelFinalizada.RecordCount > 0 Then
            DataGridMov(2).Enabled = True
        Else
            DataGridMov(2).Enabled = False
        End If
        If de_informa.rsColetaSelCancelada.RecordCount > 0 Then
            DataGridMov(3).Enabled = True
        Else
            DataGridMov(3).Enabled = False
        End If
        If de_informa.rsColetaSelBaixada.RecordCount > 0 Then
            DataGridMov(4).Enabled = True
        Else
            DataGridMov(4).Enabled = False
        End If
        DoEvents
    End If
End If
End Sub
