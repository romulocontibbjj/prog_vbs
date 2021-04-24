VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEscCliAnEstat2 
   Caption         =   "Análise Estatística 2 - Seleção dos Dados"
   ClientHeight    =   7350
   ClientLeft      =   1890
   ClientTop       =   1485
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   9330
   Begin VB.Frame fraConsCli 
      Caption         =   "Busca Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   120
      TabIndex        =   43
      Top             =   4200
      Width           =   7485
      Begin VB.TextBox txtBuscaNome 
         Height          =   285
         Left            =   1575
         MaxLength       =   25
         TabIndex        =   47
         Top             =   2520
         Width           =   1905
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Height          =   330
         Left            =   6240
         TabIndex        =   46
         Top             =   2520
         Width           =   1065
      End
      Begin VB.OptionButton optBuscaInic 
         Caption         =   "Busca no Início do Texto"
         Height          =   195
         Left            =   3795
         TabIndex        =   45
         Top             =   2460
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton optBuscaTodo 
         Caption         =   "Busca no Texto Todo"
         Height          =   195
         Left            =   3795
         TabIndex        =   44
         Top             =   2670
         Width           =   2220
      End
      Begin MSDataGridLib.DataGrid GridConsCli 
         Bindings        =   "frmEscCliAnEstat2.frx":0000
         Height          =   2055
         Left            =   120
         TabIndex        =   48
         Top             =   315
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   3625
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
               ColumnWidth     =   1590,236
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Busca por Nome:"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   2520
         Width           =   1230
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Sair"
      Height          =   375
      Left            =   7680
      TabIndex        =   42
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Processar ..."
      Height          =   375
      Left            =   7680
      TabIndex        =   41
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cliente Remetente"
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
      TabIndex        =   35
      Top             =   240
      Width           =   3735
      Begin VB.CheckBox chkTodosEstab 
         Caption         =   "Todos os Estabelecimentos"
         Height          =   225
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.TextBox txtCGCCli 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   37
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkTodosCli 
         Caption         =   "Todos os Clientes Remetentes"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblNomeCli 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CGC Remetente:"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Período 1"
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
      Left            =   3960
      TabIndex        =   25
      Top             =   2880
      Width           =   3615
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   720
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2640
         TabIndex        =   28
         Text            =   "2002"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   360
         Width           =   255
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   840
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Top             =   840
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
         Caption         =   "OU"
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
         Left            =   360
         TabIndex        =   34
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2040
         TabIndex        =   33
         Top             =   840
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período 1"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   1560
      Width           =   3615
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Text            =   "2002"
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   720
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   840
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Top             =   840
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2040
         TabIndex        =   24
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "OU"
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
         Left            =   360
         TabIndex        =   23
         Top             =   720
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções Pré-Padronizadas para Período"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   3735
      Begin VB.OptionButton Option4 
         Caption         =   "Mês Anterior Ano Atual / Mes Ano Passado"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Mes Ano Atual / Mes Atual Ano Passado"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mês Anterior / Anterior 1 / Anterior 2"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mês Atual / Anterior / Anterior 1"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame frmPer1 
      Caption         =   "Período 1"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Text            =   "2002"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   840
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
         Left            =   720
         TabIndex        =   7
         Top             =   840
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OU"
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
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmEscCliAnEstat2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Set frmEscCliAnEstat2 = Nothing
End Sub
