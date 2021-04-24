VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBaixaExpress 
   Caption         =   "Baixa POD's Express"
   ClientHeight    =   6240
   ClientLeft      =   180
   ClientTop       =   1440
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   11760
   Begin VB.Frame fraComandos 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   88
      Top             =   5280
      Width           =   4215
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "Limpar Tela"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   2760
         TabIndex        =   69
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdBaixar 
         Caption         =   "Baixar Todos..."
         Height          =   375
         Left            =   1440
         TabIndex        =   68
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status dos Processos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   86
      Top             =   5280
      Width           =   6855
      Begin VB.Label lblStatusGer 
         Caption         =   "Descrever o Status do Processo..."
         Height          =   495
         Left            =   120
         TabIndex        =   87
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados dos CTCs"
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
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   10
         Left            =   8760
         TabIndex        =   66
         Top             =   4320
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   10
         Left            =   7800
         TabIndex        =   65
         Top             =   4320
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   9
         Left            =   8760
         TabIndex        =   60
         Top             =   3960
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   9
         Left            =   7800
         TabIndex        =   59
         Top             =   3960
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   8
         Left            =   8760
         TabIndex        =   54
         Top             =   3600
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   8
         Left            =   7800
         TabIndex        =   53
         Top             =   3600
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   7
         Left            =   8760
         TabIndex        =   48
         Top             =   3240
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   7
         Left            =   7800
         TabIndex        =   47
         Top             =   3240
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   6
         Left            =   8760
         TabIndex        =   42
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   6
         Left            =   7800
         TabIndex        =   41
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   5
         Left            =   8760
         TabIndex        =   36
         Top             =   2520
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   5
         Left            =   7800
         TabIndex        =   35
         Top             =   2520
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   4
         Left            =   8760
         TabIndex        =   30
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   4
         Left            =   7800
         TabIndex        =   29
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   3
         Left            =   8760
         TabIndex        =   24
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   3
         Left            =   7800
         TabIndex        =   23
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   2
         Left            =   8760
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   2
         Left            =   7800
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   1
         Left            =   8760
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   1
         Left            =   7800
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox ChkBx 
         Caption         =   "Bx.Final"
         Height          =   195
         Index           =   0
         Left            =   8760
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkPre 
         Caption         =   "Pré-Bx."
         Height          =   195
         Index           =   0
         Left            =   7800
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   10
         Left            =   3960
         TabIndex        =   64
         Top             =   4320
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   10
         Left            =   720
         TabIndex        =   62
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   61
         Top             =   4320
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   9
         Left            =   3960
         TabIndex        =   58
         Top             =   3960
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   9
         Left            =   720
         TabIndex        =   56
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   55
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   8
         Left            =   3960
         TabIndex        =   52
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   8
         Left            =   720
         TabIndex        =   50
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   49
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   7
         Left            =   3960
         TabIndex        =   46
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   7
         Left            =   720
         TabIndex        =   44
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   6
         Left            =   3960
         TabIndex        =   40
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   6
         Left            =   720
         TabIndex        =   38
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   5
         Left            =   3960
         TabIndex        =   34
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   5
         Left            =   720
         TabIndex        =   32
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   28
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   4
         Left            =   720
         TabIndex        =   26
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   22
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   16
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   10
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtReceb 
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   4
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtCtc 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFilial 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   495
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   33
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   39
         Top             =   2880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   45
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   51
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   57
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEntrega 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   63
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   11
         Left            =   3240
         TabIndex        =   89
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   12
         Left            =   3240
         TabIndex        =   90
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   13
         Left            =   3240
         TabIndex        =   91
         Top             =   1440
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   14
         Left            =   3240
         TabIndex        =   92
         Top             =   1800
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   15
         Left            =   3240
         TabIndex        =   93
         Top             =   2160
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   16
         Left            =   3240
         TabIndex        =   94
         Top             =   2520
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   17
         Left            =   3240
         TabIndex        =   95
         Top             =   2880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   18
         Left            =   3240
         TabIndex        =   96
         Top             =   3240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   19
         Left            =   3240
         TabIndex        =   97
         Top             =   3600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   20
         Left            =   3240
         TabIndex        =   98
         Top             =   3960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Index           =   21
         Left            =   3240
         TabIndex        =   99
         Top             =   4320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Left            =   3360
         TabIndex        =   101
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   2280
         TabIndex        =   100
         Top             =   480
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   9720
         TabIndex        =   85
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   9720
         TabIndex        =   84
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   9720
         TabIndex        =   83
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   9720
         TabIndex        =   82
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   9720
         TabIndex        =   81
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   9720
         TabIndex        =   80
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   9720
         TabIndex        =   79
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   9720
         TabIndex        =   78
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   9720
         TabIndex        =   77
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   9720
         TabIndex        =   76
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   75
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   10080
         TabIndex        =   74
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Baixa"
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
         Left            =   7800
         TabIndex        =   73
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Recebedor"
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
         Left            =   4440
         TabIndex        =   72
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "          Entrega          "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   71
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial e CTC"
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
         Left            =   480
         TabIndex        =   70
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmBaixaExpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBaixar_Click()
    Dim xind As Integer
    For xind = 0 To 10
        If txtFilial(xind) <> "" And txtCtc(xind) <> "" Then
            'BAIXAR POD
            
            
            
            
            
            
            
        Else
        End If
    Next
End Sub

Private Sub cmdLimpar_Click()
    If MsgBox("Confirma LIMPAR TELA ?", vbYesNo + vbQuestion, "Confirmação") = vbYes Then
        For xind = 0 To 10
            txtFilial(xind) = ""
            txtCtc(xind) = ""
            txtReceb(xind) = ""
            lblStatus(xind) = ""
            ChkBx(xind) = 0
            chkPre(xind) = 0
            mskEntrega(xind).Mask = ""
            mskEntrega(xind).Text = ""
            mskEntrega(xind).Mask = "##/##/####"
            mskHora(xind).Mask = ""
            mskHora(xind).Text = ""
            mskHora(xind).Mask = "##:##"
        Next
        txtFilial(1).SetFocus
    End If
End Sub

Private Sub mskEntrega_LostFocus(Index As Integer)
    Dim xind As Integer
    For xind = 0 To 10
        mskEntrega(xind) = century(mskEntrega(xind))
    Next
End Sub
