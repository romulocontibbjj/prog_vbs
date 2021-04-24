VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRecalcPrazo 
   Caption         =   "Recalcular Prazos de Entrega"
   ClientHeight    =   6840
   ClientLeft      =   1650
   ClientTop       =   1335
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9360
   Begin VB.Frame fraConsCli 
      Caption         =   "Consulta Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   7365
      Begin VB.TextBox txtBuscaNome 
         Height          =   285
         Left            =   1575
         MaxLength       =   25
         TabIndex        =   15
         Top             =   3045
         Width           =   1905
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Height          =   330
         Left            =   6195
         TabIndex        =   14
         Top             =   3045
         Width           =   1065
      End
      Begin VB.OptionButton optBuscaInic 
         Caption         =   "Busca no Início do Texto"
         Height          =   195
         Left            =   3675
         TabIndex        =   13
         Top             =   2940
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton optBuscaTodo 
         Caption         =   "Busca no Texto Todo"
         Height          =   195
         Left            =   3675
         TabIndex        =   12
         Top             =   3150
         Width           =   2220
      End
      Begin MSDataGridLib.DataGrid GridConsCli 
         Bindings        =   "frmRecalcPrazo.frx":0000
         Height          =   2535
         Left            =   105
         TabIndex        =   16
         Top             =   315
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   4471
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
         DataMember      =   "Sel_CadCli"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cgc"
            Caption         =   "cgc"
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
            DataField       =   "nome"
            Caption         =   "nome"
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
            DataField       =   "cidade"
            Caption         =   "cidade"
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
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3750,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Busca por Nome:"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   3045
         Width           =   1230
      End
   End
   Begin VB.Frame fraDados 
      Caption         =   "Seleção dos Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7365
      Begin VB.CheckBox chkSair 
         Caption         =   "Sair"
         Height          =   330
         Left            =   5775
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processa..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   4200
         TabIndex        =   3
         Top             =   1320
         Width           =   1485
      End
      Begin VB.TextBox txtCGCCli 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2205
         MaxLength       =   8
         TabIndex        =   2
         Top             =   840
         Width           =   1590
      End
      Begin VB.CheckBox chkTodosCli 
         Caption         =   "Todos os Clientes Remetentes"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2625
         TabIndex        =   5
         Top             =   1260
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
         Left            =   1155
         TabIndex        =   6
         Top             =   1260
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
      Begin VB.Label lblNomeCli 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3885
         TabIndex        =   10
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2415
         TabIndex        =   9
         Top             =   1260
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Período:  De"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CGC do Cliente/Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmRecalcPrazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
