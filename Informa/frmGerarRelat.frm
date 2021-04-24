VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGerarRelat 
   Caption         =   "Gerar Relatórios / Arquivos"
   ClientHeight    =   7920
   ClientLeft      =   1110
   ClientTop       =   1695
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   11940
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   46
      Top             =   3000
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8070
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Movimentação"
      TabPicture(0)   =   "frmGerarRelat.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Posição de Entrega/Status"
      TabPicture(1)   =   "frmGerarRelat.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Dados Analíticos"
      TabPicture(2)   =   "frmGerarRelat.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2655
         Left            =   840
         TabIndex        =   48
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4683
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Documentos e Filiais"
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
      Left            =   7800
      TabIndex        =   32
      Top             =   1320
      Width           =   4095
      Begin VB.CheckBox Check5 
         Caption         =   "REE"
         Height          =   195
         Left            =   1080
         TabIndex        =   52
         Top             =   1080
         Width           =   650
      End
      Begin VB.CheckBox Check4 
         Caption         =   "DEV"
         Height          =   195
         Left            =   1080
         TabIndex        =   51
         Top             =   1290
         Width           =   650
      End
      Begin VB.CheckBox Check3 
         Caption         =   "TRA"
         Height          =   195
         Left            =   140
         TabIndex        =   50
         Top             =   1290
         Width           =   650
      End
      Begin VB.CheckBox Check2 
         Caption         =   "ENT"
         Height          =   195
         Left            =   140
         TabIndex        =   49
         Top             =   1080
         Width           =   650
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   45
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Excluir Filial"
         Height          =   195
         Left            =   2280
         TabIndex        =   44
         Top             =   1080
         Width           =   1150
      End
      Begin VB.OptionButton Option8 
         Caption         =   "CTRs"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   760
         Width           =   1335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "CTCs e CTRs"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "CTCs e COBs"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   36
         Top             =   645
         Width           =   495
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   35
         Top             =   290
         Width           =   495
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   34
         Top             =   645
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   290
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   195
         Left            =   3120
         TabIndex        =   40
         Top             =   645
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   195
         Left            =   3120
         TabIndex        =   39
         Top             =   290
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   195
         Left            =   1920
         TabIndex        =   38
         Top             =   645
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   195
         Left            =   1920
         TabIndex        =   37
         Top             =   290
         Width           =   345
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Localidade"
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
      Left            =   3360
      TabIndex        =   24
      Top             =   1320
      Width           =   4335
      Begin VB.TextBox txtCidade 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   30
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   3240
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Interior"
         Height          =   195
         Left            =   1680
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Capital"
         Height          =   195
         Left            =   3240
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Todo Estado"
         Height          =   195
         Left            =   1680
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cmbUF 
         Height          =   315
         Left            =   480
         TabIndex        =   25
         Text            =   "Todos"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   960
         TabIndex        =   47
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   3135
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   300
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
         TabIndex        =   22
         Top             =   300
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   1480
         TabIndex        =   23
         Top             =   300
         Width           =   90
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Modal"
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
      TabIndex        =   16
      Top             =   2160
      Width           =   3135
      Begin VB.OptionButton optAir 
         Caption         =   "Aéreo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optRodo 
         Caption         =   "Rodoviário"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ambos"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1095
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   720
         MaxLength       =   8
         TabIndex        =   12
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destinatário"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   720
         MaxLength       =   8
         TabIndex        =   8
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remetente"
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
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtCGCRem 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   720
         MaxLength       =   8
         TabIndex        =   4
         Top             =   360
         Width           =   1590
      End
      Begin VB.CommandButton cmdBuscaREM 
         Caption         =   "?"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblNomeRem 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.OLE OLE1 
      Height          =   1215
      Left            =   4440
      TabIndex        =   53
      Top             =   3360
      Width           =   3015
   End
End
Attribute VB_Name = "frmGerarRelat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
