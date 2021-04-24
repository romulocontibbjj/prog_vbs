VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmImprGeral 
   Caption         =   "Seleção dos Dados para Impressão"
   ClientHeight    =   6585
   ClientLeft      =   1275
   ClientTop       =   1230
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdProcessa 
      Caption         =   "Processa..."
      Enabled         =   0   'False
      Height          =   330
      Left            =   3120
      TabIndex        =   29
      Top             =   6120
      Width           =   1515
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   330
      Left            =   5160
      TabIndex        =   28
      Top             =   6120
      Width           =   1515
   End
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
      Height          =   3000
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   9075
      Begin VB.TextBox txtBuscaNome 
         Height          =   285
         Left            =   7200
         MaxLength       =   25
         TabIndex        =   18
         Top             =   1920
         Width           =   1785
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Height          =   330
         Left            =   7560
         TabIndex        =   17
         Top             =   2400
         Width           =   1065
      End
      Begin VB.OptionButton optBuscaInic 
         Caption         =   "Busca no Início do Texto"
         Height          =   435
         Left            =   7320
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton optBuscaTodo 
         Caption         =   "Busca no Texto Todo"
         Height          =   435
         Left            =   7320
         TabIndex        =   15
         Top             =   1080
         Width           =   1500
      End
      Begin MSDataGridLib.DataGrid GridConsCli 
         Bindings        =   "frmImprGeral.frx":0000
         Height          =   2535
         Left            =   105
         TabIndex        =   19
         Top             =   315
         Width           =   6960
         _ExtentX        =   12277
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
               ColumnWidth     =   1379,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3539,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1635,024
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Busca por Nome:"
         Height          =   195
         Left            =   7200
         TabIndex        =   20
         Top             =   1680
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
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9045
      Begin VB.Frame Frame2 
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
         Height          =   1335
         Left            =   5400
         TabIndex        =   22
         Top             =   1320
         Width           =   3495
         Begin VB.CheckBox Check1 
            Caption         =   "Brasil"
            Height          =   195
            Left            =   720
            TabIndex        =   26
            Top             =   360
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   720
            TabIndex        =   25
            Top             =   840
            Width           =   2655
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2520
            TabIndex        =   24
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   2160
            TabIndex        =   23
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Posição de Entrega"
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
         TabIndex        =   21
         Top             =   1680
         Width           =   5175
         Begin VB.CheckBox Check7 
            Caption         =   "Com Ocorr. - Pendente"
            Height          =   255
            Left            =   960
            TabIndex        =   34
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox Check6 
            Caption         =   "OK. Entregue"
            Height          =   255
            Left            =   3120
            TabIndex        =   33
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Com Ocorr. - Fechado"
            Height          =   195
            Left            =   3120
            TabIndex        =   32
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Sem Posição - Pendente"
            Height          =   255
            Left            =   960
            TabIndex        =   31
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Todos"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.CheckBox chkTodosEstab 
         Caption         =   "Todos os Estabelecimentos"
         Height          =   465
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.TextBox txtCGCCli 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2205
         MaxLength       =   8
         TabIndex        =   5
         Top             =   840
         Width           =   1590
      End
      Begin VB.CheckBox chkTodosCli 
         Caption         =   "Todos os Clientes Remetentes"
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkModal 
         Caption         =   "Todos Modais"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5280
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.OptionButton optRodo 
         Caption         =   "Rodoviário"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7560
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optAir 
         Caption         =   "Aéreo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7560
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2880
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   13
         Top             =   840
         Width           =   4980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   1260
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Período:  De"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CGC do Cliente/Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Modal:"
         Height          =   195
         Left            =   6960
         TabIndex        =   9
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmImprGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmImprGeral = Nothing
End Sub

Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtBuscaNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCGCCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
