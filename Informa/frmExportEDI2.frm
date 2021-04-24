VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmExportEDI 
   Caption         =   "Exporta EDI"
   ClientHeight    =   9690
   ClientLeft      =   120
   ClientTop       =   705
   ClientWidth     =   14715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   14715
   Begin VB.Frame Frame10 
      Caption         =   "Arquivo de Fatura - DOCCOB (Cliente)"
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
      TabIndex        =   47
      Top             =   2830
      Width           =   4215
      Begin VB.OptionButton optPorCliente 
         Caption         =   "Por CNPJ Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optPorFatura 
         Caption         =   "Por Filial Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCgcDocCob 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1680
         MaxLength       =   14
         TabIndex        =   53
         Top             =   680
         Width           =   1335
      End
      Begin VB.TextBox txtFatura 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDocCod 
         Caption         =   "Gerar Arquivo"
         Height          =   615
         Left            =   3120
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "ou"
         Height          =   195
         Left            =   2760
         TabIndex        =   54
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "CONEMB"
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
      TabIndex        =   29
      Top             =   120
      Width           =   4215
      Begin VB.Frame FraDataCONEMB 
         Caption         =   "Intervalo de Data"
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   3975
         Begin MSMask.MaskEdBox MskDataFinalCONEMB 
            Height          =   285
            Left            =   2280
            TabIndex        =   37
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDataInicialCONEMB 
            Height          =   285
            Left            =   480
            TabIndex        =   35
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1920
            TabIndex        =   36
            Top             =   360
            Width           =   105
         End
      End
      Begin VB.CommandButton CmdCONEMB 
         Caption         =   "Gerar CONEMB"
         Height          =   435
         Left            =   960
         TabIndex        =   40
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Frame fraCliente 
         Height          =   1095
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   3975
         Begin VB.TextBox txtCGCRem 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   8
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdBuscaREM 
            Caption         =   "?"
            Height          =   255
            Left            =   2280
            TabIndex        =   39
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox chkTodosEstab 
            Caption         =   "Todos Estabel."
            Height          =   345
            Left            =   2760
            TabIndex        =   31
            Top             =   240
            Value           =   1  'Checked
            Width           =   1125
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblNomeRem 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   3795
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Arquivo Bonagura - Faturamento e CTC Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8880
      TabIndex        =   28
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame9 
         Height          =   2055
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   4095
         Begin VB.CommandButton cmdBonaFAT 
            Caption         =   "Gerar Arquivo Faturas Bonagura"
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   1560
            Width           =   3615
         End
         Begin VB.CommandButton cmdBonaCTC 
            Caption         =   "Gerar Arquivo CTCs Bonagura"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   600
            Width           =   3615
         End
         Begin MSMask.MaskEdBox mskDataPer1 
            Height          =   285
            Left            =   600
            TabIndex        =   43
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataPer2 
            Height          =   285
            Left            =   2160
            TabIndex        =   44
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   285
            Left            =   600
            TabIndex        =   50
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12632256
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   285
            Left            =   2160
            TabIndex        =   51
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12632256
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1920
            TabIndex        =   52
            Top             =   1200
            Width           =   90
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3960
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1920
            TabIndex        =   45
            Top             =   240
            Width           =   90
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Arquivo para CORREIOS"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   2640
      Width           =   4335
      Begin VB.CommandButton cmdCorreio 
         Caption         =   "Gera Arquivo ..."
         Height          =   495
         Left            =   2520
         TabIndex        =   26
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         Caption         =   "No Período de ...  (máximo de 60 dias)"
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3975
         Begin MSMask.MaskEdBox mskPer2Correio 
            Height          =   285
            Left            =   2280
            TabIndex        =   25
            Top             =   360
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
         Begin MSMask.MaskEdBox mskPer1Correio 
            Height          =   285
            Left            =   600
            TabIndex        =   23
            Top             =   360
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1920
            TabIndex        =   24
            Top             =   360
            Width           =   90
         End
      End
      Begin VB.Label Label21 
         Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\CORREIOS\"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Arquivo para Bomi (BOSS)"
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
      Left            =   8880
      TabIndex        =   14
      Top             =   2640
      Width           =   4335
      Begin VB.Frame Frame6 
         Caption         =   "No Período de ...  (máximo de 60 dias)"
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3975
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   2280
            TabIndex        =   20
            Top             =   360
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
            Left            =   600
            TabIndex        =   18
            Top             =   360
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1920
            TabIndex        =   19
            Top             =   360
            Width           =   90
         End
      End
      Begin VB.CommandButton cmdGeraBomi 
         Caption         =   "Gerar Arquivo - INTEC.TXT"
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\BOMIBRASIL\"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arquivo de Ocorrências - OCOREN"
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
      TabIndex        =   9
      Top             =   4680
      Width           =   13095
      Begin VB.ComboBox cmb_clientes 
         Height          =   315
         ItemData        =   "frmExportEDI.frx":0000
         Left            =   600
         List            =   "frmExportEDI.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmd_alcon 
         Caption         =   "ALCON"
         Height          =   255
         Left            =   600
         TabIndex        =   99
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame fra_cgcs 
         Caption         =   "CGCS"
         Height          =   3915
         Left            =   4320
         TabIndex        =   75
         Top             =   240
         Visible         =   0   'False
         Width           =   5955
         Begin VB.TextBox txt_gillette 
            Height          =   285
            Left            =   960
            TabIndex        =   83
            Text            =   "04490850"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txt_medley 
            Height          =   285
            Left            =   1020
            TabIndex        =   82
            Text            =   "50929710000179"
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox txt_bayer 
            Height          =   285
            Left            =   1020
            TabIndex        =   81
            Text            =   "14372981000102"
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txt_givaudan 
            Height          =   285
            Left            =   1020
            TabIndex        =   80
            Text            =   "61188488000117"
            Top             =   2700
            Width           =   1695
         End
         Begin VB.TextBox txt_videolar 
            Height          =   285
            Left            =   1020
            TabIndex        =   79
            Text            =   "04229761"
            Top             =   3060
            Width           =   1695
         End
         Begin VB.TextBox txt_alcon 
            Height          =   285
            Left            =   4260
            TabIndex        =   78
            Text            =   "60412327"
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox txt_glaxo 
            Height          =   285
            Left            =   4260
            TabIndex        =   77
            Text            =   "33247743"
            Top             =   1260
            Width           =   1575
         End
         Begin VB.TextBox txt_boehringer 
            Height          =   285
            Left            =   4260
            TabIndex        =   76
            Text            =   "60831658"
            Top             =   2220
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\MEDLEY"
            Height          =   495
            Left            =   60
            TabIndex        =   98
            Top             =   1380
            Width           =   2655
         End
         Begin VB.Label Label7 
            Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\GILLETTE"
            Height          =   495
            Left            =   60
            TabIndex        =   97
            Top             =   660
            Width           =   2655
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Gillette: "
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
            Left            =   60
            TabIndex        =   96
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "MedLey:"
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
            Left            =   60
            TabIndex        =   95
            Top             =   1140
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bayer: "
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
            Left            =   60
            TabIndex        =   94
            Top             =   1860
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\BAYER"
            Height          =   435
            Left            =   180
            TabIndex        =   93
            Top             =   2220
            Width           =   2730
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Givaudan:"
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
            Left            =   60
            TabIndex        =   92
            Top             =   2700
            Width           =   885
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Videolar:"
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
            Left            =   60
            TabIndex        =   91
            Top             =   3060
            Width           =   765
         End
         Begin VB.Label Label15 
            Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\VIDEOLAR"
            Height          =   435
            Left            =   60
            TabIndex        =   90
            Top             =   3300
            Width           =   2730
         End
         Begin VB.Label Label18 
            Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\ALCON"
            Height          =   435
            Left            =   3180
            TabIndex        =   89
            Top             =   720
            Width           =   2730
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Alcon:"
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
            Left            =   3180
            TabIndex        =   88
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Glaxo:"
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
            Left            =   3180
            TabIndex        =   87
            Top             =   1260
            Width           =   555
         End
         Begin VB.Label Label26 
            Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\GLAXO"
            Height          =   435
            Left            =   3180
            TabIndex        =   86
            Top             =   1620
            Width           =   2730
         End
         Begin VB.Label Label27 
            Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\BOEHRINGER"
            Height          =   435
            Left            =   3120
            TabIndex        =   85
            Top             =   2640
            Width           =   2730
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Boheringer:"
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
            Left            =   3180
            TabIndex        =   84
            Top             =   2220
            Width           =   990
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "SELECIONE CLIENTES"
         Height          =   4155
         Left            =   10320
         TabIndex        =   63
         Top             =   240
         Width           =   2655
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Left            =   720
            Top             =   3600
         End
         Begin VB.CheckBox chk_correios 
            Caption         =   "CORREIOS"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk_bomi 
            Caption         =   "BOMI"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chk_bona 
            Caption         =   "BONAGURA"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk_alcon 
            Caption         =   "ALCON"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CommandButton cmd_todos 
            Caption         =   "GERAR"
            Height          =   255
            Left            =   360
            TabIndex        =   69
            Top             =   3360
            Width           =   1335
         End
         Begin VB.CheckBox chk_medley 
            Caption         =   "MEDLEY"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   540
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chk_videolar 
            Caption         =   "VIDEOLAR"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   840
            Value           =   1  'Checked
            Width           =   1155
         End
         Begin VB.CheckBox chk_boe 
            Caption         =   "BOEHRINGER"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1395
         End
         Begin VB.CheckBox chk_bayer 
            Caption         =   "BAYER"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1140
            Width           =   1215
         End
         Begin VB.CheckBox chk_gillette 
            Caption         =   "GILLETTE"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label lab_correios 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   108
            Top             =   2880
            Width           =   800
         End
         Begin VB.Label lab_bomi 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   105
            Top             =   2520
            Width           =   800
         End
         Begin VB.Label lab_bona 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   103
            Top             =   2160
            Width           =   800
         End
         Begin VB.Label lab_alcon 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   101
            Top             =   1800
            Width           =   800
         End
         Begin VB.Label lab_bayer 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   74
            Top             =   1140
            Width           =   800
         End
         Begin VB.Label lab_boeh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   73
            Top             =   1440
            Width           =   800
         End
         Begin VB.Label lab_medley 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   72
            Top             =   540
            Width           =   800
         End
         Begin VB.Label lab_videolar 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   71
            Top             =   840
            Width           =   800
         End
         Begin VB.Label lab_gillette 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   70
            Top             =   240
            Width           =   800
         End
      End
      Begin VB.CommandButton cmd_medley 
         Caption         =   "MEDLEY"
         Height          =   255
         Left            =   600
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmd_videolar 
         Caption         =   "VIDEOLAR"
         Height          =   255
         Left            =   600
         TabIndex        =   61
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmd_bayer 
         Caption         =   "BAYER"
         Height          =   255
         Left            =   600
         TabIndex        =   60
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmd_boeh 
         Caption         =   "BOEHRINGER"
         Height          =   255
         Left            =   600
         TabIndex        =   59
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmd_cgcs 
         Caption         =   "MOSTRAR CGCS"
         Height          =   255
         Left            =   600
         TabIndex        =   58
         Top             =   1200
         Width           =   1635
      End
      Begin VB.CommandButton cmd_gillette 
         Caption         =   "GILLETTE"
         Height          =   255
         Left            =   600
         TabIndex        =   57
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtCGC 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENVIAR ARQUIVOS DE FATURA DA GILLETTE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1815
         Left            =   600
         TabIndex        =   106
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "CGC:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Padrão PROCEDA"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arquivo de Embarques MEDLEY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdLerArquivoMedley 
         Caption         =   "Ler Arquivo Bomi  -  MEDLEY.TXT"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdGeraArqMedley 
         Caption         =   "Gerar Arquivo  -  BOMDDMMAAAA.TXT"
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Padrão MEDLEY"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label8 
         Caption         =   "CGC:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "O Arquivo da BOMI deve estar no diretório C:\INFORMA\MEDLEY"
         Height          =   435
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   2580
      End
      Begin VB.Label Label10 
         Caption         =   "Arquivo EDI Gerado no Diretório C:\INFORMA\EXP_EDI\MEDLEY"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblCGC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50929710000179"
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "SAIR"
      Height          =   375
      Left            =   13320
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmExportEDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xqtdreg As Integer
Public xmsg As Integer





Private Sub cmb_clientes_Click()

Select Case cmb_clientes.Text

Case "ALCON":
txtCgc.Text = txt_alcon.Text
Case "GILLETTE":
txtCgc.Text = txt_gillette.Text
txtCgcDocCob.Text = txt_gillette.Text
Case "MEDLEY":
txtCgc.Text = txt_medley.Text
Case "GIVAUDAN":
txtCgc.Text = txt_givaudan.Text
txtCgcDocCob.Text = txt_givaudan.Text
Case "BOEHRINGER"
txtCgc.Text = txt_boehringer.Text
txtCgcDocCob.Text = txt_boehringer.Text
Case "VIDEOLAR":
txtCgc.Text = txt_videolar.Text
txtCgcDocCob.Text = txt_videolar.Text
Case "GLAXO":
txtCgc.Text = txt_glaxo.Text
txtCgcDocCob.Text = txt_glaxo.Text
Case "BAYER":
txtCgc.Text = txt_bayer.Text
txtCgcDocCob.Text = txt_bayer.Text
End Select








End Sub

Private Sub cmd_alcon_Click()
txtCgc.Text = "60412327"
CmdProcessar_Click
End Sub

Private Sub cmd_bayer_Click()
txtCgc.Text = "14372981000102"
CmdProcessar_Click
End Sub

Private Sub cmd_boeh_Click()
Dim inicio As Date
Dim final As Date

final = Mid(Now, 1, 10)
inicio = final - 10

txtCgc.Text = "60831658"
CmdProcessar_Click



'If MsgBox("Deseja gerar Arquivos de Conhecimentos referente a 10 dias?", vbExclamation + vbYesNo, "CONHECIMETO") = vbYes Then

TxtCGCRem.Text = "60831658"
MskDataInicialCONEMB.Mask = ""
MskDataFinalCONEMB.Mask = ""
MskDataInicialCONEMB = inicio
MskDataFinalCONEMB = final
CmdCONEMB_Click

'End If



End Sub

Private Sub cmd_cgcs_Click()

If fra_cgcs.Visible = False Then
    fra_cgcs.Visible = True
    cmd_cgcs.Caption = "OCULTAR CGCS"
ElseIf fra_cgcs.Visible = True Then
    fra_cgcs.Visible = False
    cmd_cgcs.Caption = "MOSTRAR CGCS"
End If


End Sub

Private Sub cmd_gillette_Click()
Dim inicio As Date
Dim final As Date

final = Mid(Now, 1, 10)
inicio = final - 10



txtCgc.Text = "04490850"
CmdProcessar_Click

'If MsgBox("Deseja gerar Arquivos de Conhecimentos referente a 10 dias?", vbExclamation + vbYesNo, "CONHECIMETO") = vbYes Then

TxtCGCRem.Text = "04490850"
MskDataInicialCONEMB.Mask = ""
MskDataFinalCONEMB.Mask = ""
MskDataInicialCONEMB = inicio
MskDataFinalCONEMB = final
CmdCONEMB_Click

'End If






End Sub

Private Sub cmd_medley_Click()
txtCgc.Text = "50929710000179"
CmdProcessar_Click

'If MsgBox("Deseja Gerar Arquivos de Conhecimentos?", vbInformation + vbYesNo, "CONHECIMENTOS") = vbYes Then
    cmdGeraArqMedley_Click
'End If


End Sub

Private Sub cmd_todos_Click()
xmsg = 1


If chk_bona.Value = 1 Then
    cmdBonaFAT_Click
    lab_bona.Caption = xqtdreg
    Frame11.Refresh
End If


If chk_gillette.Value = 1 Then
    cmd_gillette_Click
    lab_gillette.Caption = xqtdreg
    Frame11.Refresh
    
    If Weekday(Date) = 2 Then
        txtCgcDocCob.Text = "04490850"
        cmdDocCod_Click
    End If
        
    
End If

If chk_medley.Value = 1 Then
    cmd_medley_Click
    lab_medley.Caption = xqtdreg
    Frame11.Refresh
End If

If chk_videolar.Value = 1 Then
    cmd_videolar_Click
    lab_videolar.Caption = xqtdreg
    Frame11.Refresh
End If

If chk_boe.Value = 1 Then
    cmd_boeh_Click
    lab_boeh.Caption = xqtdreg
    Frame11.Refresh
End If


If chk_bayer.Value = 1 Then
    cmd_bayer_Click
    lab_bayer.Caption = xqtdreg
    Frame11.Refresh
End If

If chk_alcon.Value = 1 Then
    cmd_alcon_Click
    lab_alcon.Caption = xqtdreg
    Frame11.Refresh
End If

If chk_bomi.Value = 1 Then
    mskPer1.Mask = ""
    mskPer1.Text = InputBox("INTEC.TXT", "Data Inicial:", "  /  /  ")
    mskPer2.Mask = ""
    mskPer2.Text = InputBox("INTEC.TXT", "Data Final:", "  /  /  ")
    cmdGeraBomi_Click
    lab_bomi.Caption = xqtdreg
    Frame11.Refresh
End If

If chk_correios.Value = 1 Then
    mskPer1Correio.Mask = ""
    mskPer1Correio.Text = InputBox("CORREIOS", "Data Inicial:", "  /  /  ")
    mskPer2Correio.Mask = ""
    mskPer2Correio.Text = InputBox("CORREIOS", "Data Final:", "  /  /  ")
    cmdCorreio_Click
    lab_correios.Caption = xqtdreg
    Frame11.Refresh
End If

xmsg = 0




End Sub

Private Sub cmd_videolar_Click()
txtCgc.Text = "04229761"
CmdProcessar_Click
End Sub

Private Sub cmdBonaCTC_Click()
    Dim xEmissao As String, xDocto As String, xserie As String, xEspecie As String, xCfop As String, xuf As String
    Dim xCnpj As String, xSPC1 As String, xContabil As String, xSPC2 As String, xBase_icms As String
    Dim xAliq_icms As String, xVlr_icms As String, xIse_icms As String, xOut_icms As String, xSPC3 As String
    Dim xDest_cnpj As String, xDest_nome As String, xDest_ie As String, xDest_uf As String
    Dim xRemet_cnpj As String, xRemet_nome As String, xRemet_ie As String, xRemet_uf As String
    Dim xClient_cnpj As String, xClient_nome As String, xClient_ie As String, xClient_uf As String
    Dim xQuem_paga As String, xTipo_Transp As String, xlinha As String
    Dim xDocto2 As String, xSerie2 As String, xSubserie2 As String, xEmissao2 As String, xValor2 As String

    If Not IsDate(mskDataPer1) Then
        MsgBox "Data Inválida !", vbCritical
        mskDataPer1.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskDataPer2) Then
        MsgBox "Data Inválida !", vbCritical
        mskDataPer2.SetFocus
        Exit Sub
    End If
    
    'busca Faturas não atualizadas
    If de_informa.rsSel_CTCEDIContabil.State = 1 Then de_informa.rsSel_CTCEDIContabil.Close
    de_informa.Sel_CTCEDIContabil CDate(mskDataPer1), CDate(mskDataPer2)
    
    If de_informa.rsSel_CTCEDIContabil.RecordCount > 0 Then
    
        'abre arquivo
        Open "C:\INFORMA\CONTABIL\M5CTC.TXT" For Output As #1
        
        Do Until de_informa.rsSel_CTCEDIContabil.EOF
        
            If de_informa.rsSel_CTCEDIContabil.Fields("tem_ocorr") = "C" And IsNull(de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif")) Then
                de_informa.rsSel_CTCEDIContabil.MoveNext
            ElseIf de_informa.rsSel_CTCEDIContabil.Fields("tem_ocorr") = "C" And de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif") = "" Then
                de_informa.rsSel_CTCEDIContabil.MoveNext
            Else
                If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
                de_informa.Sel_CadFilial de_informa.rsSel_CTCEDIContabil.Fields("filial")
                        
                xEmissao = zeros(Day(de_informa.rsSel_CTCEDIContabil.Fields("data")), 2) & "/"
                xEmissao = xEmissao & zeros(Month(de_informa.rsSel_CTCEDIContabil.Fields("data")), 2) & "/"
                xEmissao = xEmissao & Trim$(Str(Year(de_informa.rsSel_CTCEDIContabil.Fields("data"))))
                xDocto = zeros2(Str(CDbl(Mid$(de_informa.rsSel_CTCEDIContabil.Fields("filialctc"), 3, 8))), 6)
                xserie = "001"
                xEspecie = "CTC"
                
                If IsNull(de_informa.rsSel_CTCEDIContabil.Fields("cfop")) Or Trim$(de_informa.rsSel_CTCEDIContabil.Fields("cfop")) = "" Then
                    If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
                    de_informa.Sel_CadCliCGC de_informa.rsSel_CTCEDIContabil.Fields("respons_cgc")
                    If Trim$(de_informa.rsSel_CadCliCGC.Fields("cfop")) = "" Then
                        If de_informa.rsSel_CadCliCGC.Fields("cfop") = "SP" Then
                            xCfop = "5.353"
                        Else
                            xCfop = "6.353"
                        End If
                    Else
                        xCfop = Format(de_informa.rsSel_CadCliCGC.Fields("cfop"), "#,###")
                    End If
                Else
                    xCfop = Format(de_informa.rsSel_CTCEDIContabil.Fields("cfop"), "#,###")
                End If
                
                xuf = de_informa.rsSel_CadFilial.Fields("uf")
                xCnpj = Format(de_informa.rsSel_CadFilial.Fields("cgc"), "@@.@@@.@@@/@@@@-@@")
                
                'xpro_codigo,pro_descricao,pro_cst,xsequencia
                xSPC1 = Space(76)
                
                If de_informa.rsSel_CTCEDIContabil.Fields("subtrib") = "S" Then
                    xContabil = SoNumeros(Format(de_informa.rsSel_CTCEDIContabil.Fields("fretetotal"), "#########0.00"))
                Else
                    xContabil = SoNumeros(Format(de_informa.rsSel_CTCEDIContabil.Fields("fretetotalbruto"), "#########0.00"))
                End If
                
                xContabil = String(14 - Len(xContabil), " ") & Mid$(xContabil, 1, Len(xContabil) - 2) & "." & Mid$(xContabil, Len(xContabil) - 1)
                
                xSPC2 = " 0.00" & "           0.00" & "           0.00" & "           0.00" & "           0.00"
                
                If de_informa.rsSel_CTCEDIContabil.Fields("tribut") = "S" And de_informa.rsSel_CTCEDIContabil.Fields("subtrib") = "N" Then
                    xBase_icms = SoNumeros(Format(de_informa.rsSel_CTCEDIContabil.Fields("fretetotalbruto"), "#########0.00"))
                    xBase_icms = String(14 - Len(xBase_icms), " ") & Mid$(xBase_icms, 1, Len(xBase_icms) - 2) & "." & Mid$(xBase_icms, Len(xBase_icms) - 1)
                    xAliq_icms = SoNumeros(Format(de_informa.rsSel_CTCEDIContabil.Fields("ALIQUOTA") * 100, "#0.00"))
                    xAliq_icms = String(4 - Len(xAliq_icms), " ") & Mid$(xAliq_icms, 1, Len(xAliq_icms) - 2) & "." & Mid$(xAliq_icms, Len(xAliq_icms) - 1)
                    xVlr_icms = SoNumeros(Format(de_informa.rsSel_CTCEDIContabil.Fields("fretetotalbruto") - de_informa.rsSel_CTCEDIContabil.Fields("fretetotal"), "#########0.00"))
                    xVlr_icms = String(14 - Len(xVlr_icms), " ") & Mid$(xVlr_icms, 1, Len(xVlr_icms) - 2) & "." & Mid$(xVlr_icms, Len(xVlr_icms) - 1)
                    xIse_icms = "           0.00"
                    xOut_icms = "           0.00"
                Else
                    xBase_icms = "           0.00"
                    xAliq_icms = " 0.00"
                    xVlr_icms = "           0.00"
                    xIse_icms = "           0.00"
                    xOut_icms = "           0.00"
                End If
                
                xSPC3 = "           0.00" & " 0.00" & "           0.00" & Space(100) & "00.00" & "           0.00" & _
                        "               " & Space(20) & Space(5) & "         0.0000" & "           0.00" & "           0.00" & _
                        "           0.00" & Space(20) & Space(50) & Space(20) & Space(2)
                xDest_cnpj = de_informa.rsSel_CTCEDIContabil.Fields("dest_cgc") & Space(6)
                xDest_nome = de_informa.rsSel_CTCEDIContabil.Fields("dest_nome") & _
                             String(50 - Len(de_informa.rsSel_CTCEDIContabil.Fields("dest_nome")), " ")
                xDest_ie = de_informa.rsSel_CTCEDIContabil.Fields("dest_ie") & _
                             String(20 - Len(de_informa.rsSel_CTCEDIContabil.Fields("dest_ie")), " ")
                xDest_uf = de_informa.rsSel_CTCEDIContabil.Fields("dest_uf")
                xRemet_cnpj = de_informa.rsSel_CTCEDIContabil.Fields("remet_cgc") & Space(6)
                xRemet_nome = de_informa.rsSel_CTCEDIContabil.Fields("remet_nome") & _
                             String(50 - Len(de_informa.rsSel_CTCEDIContabil.Fields("remet_nome")), " ")
                xRemet_ie = de_informa.rsSel_CTCEDIContabil.Fields("remet_ie") & _
                             String(20 - Len(de_informa.rsSel_CTCEDIContabil.Fields("remet_ie")), " ")
                xRemet_uf = de_informa.rsSel_CTCEDIContabil.Fields("remet_uf")
                xClient_cnpj = de_informa.rsSel_CTCEDIContabil.Fields("respons_cgc") & Space(6)
                xClient_nome = de_informa.rsSel_CTCEDIContabil.Fields("respons_nome") & _
                             String(50 - Len(de_informa.rsSel_CTCEDIContabil.Fields("respons_nome")), " ")
                xClient_ie = de_informa.rsSel_CTCEDIContabil.Fields("respons_ie") & _
                             String(20 - Len(de_informa.rsSel_CTCEDIContabil.Fields("respons_ie")), " ")
                xClient_uf = de_informa.rsSel_CTCEDIContabil.Fields("respons_uf")
                xQuem_paga = "C"
                xTipo_Transp = Mid$(de_informa.rsSel_CTCEDIContabil.Fields("modal"), 1, 1)
                
                If IsNull(de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif")) Then
                    xlinha = "FIS-CTRCS" & xEmissao & xDocto & xserie & xEspecie & xCfop & xuf & xCnpj & xSPC1 & xContabil & xSPC2 & _
                             xBase_icms & xAliq_icms & xVlr_icms & xIse_icms & xOut_icms & xSPC3 & xDest_cnpj & _
                             xDest_nome & xDest_ie & xDest_uf & xRemet_cnpj & xRemet_nome & xRemet_ie & xRemet_uf & _
                             xClient_cnpj & xClient_nome & xClient_ie & xClient_uf & xQuem_paga & xTipo_Transp & "I" & Space(10) & _
                             zeros2(Str(CDbl(Mid$(de_informa.rsSel_CTCEDIContabil.Fields("filialctc"), 1, 2))), 2)
                    Print #1, xlinha
                ElseIf de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif") = "" Then
                    xlinha = "FIS-CTRCS" & xEmissao & xDocto & xserie & xEspecie & xCfop & xuf & xCnpj & xSPC1 & xContabil & xSPC2 & _
                             xBase_icms & xAliq_icms & xVlr_icms & xIse_icms & xOut_icms & xSPC3 & xDest_cnpj & _
                             xDest_nome & xDest_ie & xDest_uf & xRemet_cnpj & xRemet_nome & xRemet_ie & xRemet_uf & _
                             xClient_cnpj & xClient_nome & xClient_ie & xClient_uf & xQuem_paga & xTipo_Transp & "I" & Space(10) & _
                             zeros2(Str(CDbl(Mid$(de_informa.rsSel_CTCEDIContabil.Fields("filialctc"), 1, 2))), 2)
                    Print #1, xlinha
                ElseIf de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif") = "A" Then
                    If de_informa.rsSel_CTCEDIContabil.Fields("tem_ocorr") = "C" Then
                        xlinha = "FIS-CTRCS" & xEmissao & xDocto & xserie & xEspecie & xCfop & xuf & xCnpj & xSPC1 & xContabil & xSPC2 & _
                                 xBase_icms & xAliq_icms & xVlr_icms & xIse_icms & xOut_icms & xSPC3 & xDest_cnpj & _
                                 xDest_nome & xDest_ie & xDest_uf & xRemet_cnpj & xRemet_nome & xRemet_ie & xRemet_uf & _
                                 xClient_cnpj & xClient_nome & xClient_ie & xClient_uf & xQuem_paga & xTipo_Transp & "E" & Space(10) & _
                                 zeros2(Str(CDbl(Mid$(de_informa.rsSel_CTCEDIContabil.Fields("filialctc"), 1, 2))), 2)
                        Print #1, xlinha
                    Else
                        xlinha = "FIS-CTRCS" & xEmissao & xDocto & xserie & xEspecie & xCfop & xuf & xCnpj & xSPC1 & xContabil & xSPC2 & _
                                 xBase_icms & xAliq_icms & xVlr_icms & xIse_icms & xOut_icms & xSPC3 & xDest_cnpj & _
                                 xDest_nome & xDest_ie & xDest_uf & xRemet_cnpj & xRemet_nome & xRemet_ie & xRemet_uf & _
                                 xClient_cnpj & xClient_nome & xClient_ie & xClient_uf & xQuem_paga & xTipo_Transp & "E" & Space(10) & _
                                 zeros2(Str(CDbl(Mid$(de_informa.rsSel_CTCEDIContabil.Fields("filialctc"), 1, 2))), 2)
                        Print #1, xlinha
                        xlinha = "FIS-CTRCS" & xEmissao & xDocto & xserie & xEspecie & xCfop & xuf & xCnpj & xSPC1 & xContabil & xSPC2 & _
                                 xBase_icms & xAliq_icms & xVlr_icms & xIse_icms & xOut_icms & xSPC3 & xDest_cnpj & _
                                 xDest_nome & xDest_ie & xDest_uf & xRemet_cnpj & xRemet_nome & xRemet_ie & xRemet_uf & _
                                 xClient_cnpj & xClient_nome & xClient_ie & xClient_uf & xQuem_paga & xTipo_Transp & "I" & Space(10) & _
                                 zeros2(Str(CDbl(Mid$(de_informa.rsSel_CTCEDIContabil.Fields("filialctc"), 1, 2))), 2)
                        Print #1, xlinha
                    End If
                End If
                
                'registro nf - vinc
                
                If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
                de_informa.Sel_NFsdoCTC de_informa.rsSel_CTCEDIContabil.Fields("filialctc")
                
                Do Until de_informa.rsSel_NFsdoCTC.EOF
                    
                    If Len(Trim$(de_informa.rsSel_NFsdoCTC.Fields("numnf"))) > 6 Then
                        xDocto2 = Mid$(de_informa.rsSel_NFsdoCTC.Fields("numnf"), 1, 6)
                    Else
                        xDocto2 = zeros2(de_informa.rsSel_NFsdoCTC.Fields("numnf"), 6)
                    End If
                    
                    xSerie2 = zeros2(de_informa.rsSel_NFsdoCTC.Fields("serie"), 3)
                    xSubserie2 = "   "
                    If IsDate(de_informa.rsSel_NFsdoCTC.Fields("emissao_nf")) Then
                        xEmissao2 = zeros(Day(de_informa.rsSel_NFsdoCTC.Fields("emissao_nf")), 2) & "/"
                        xEmissao2 = xEmissao & zeros(Month(de_informa.rsSel_NFsdoCTC.Fields("emissao_nf")), 2) & "/"
                        xEmissao2 = xEmissao & Trim$(Str(Year(de_informa.rsSel_NFsdoCTC.Fields("emissao_nf"))))
                    Else
                        xEmissao2 = xEmissao
                    End If
                    
                    xValor2 = SoNumeros(Format(de_informa.rsSel_NFsdoCTC.Fields("valornf"), "#########0.00"))
                    xValor2 = String(14 - Len(xValor2), " ") & Mid$(xValor2, 1, Len(xValor2) - 2) & "." & Mid$(xValor2, Len(xValor2) - 1)
                    
                    If IsNull(de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif")) Then
                        xlinha = "NF-VINC" & xDocto2 & xSerie2 & xSubserie2 & xEmissao2 & xValor2 & "I"
                        Print #1, xlinha
                    ElseIf de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif") = "" Then
                        xlinha = "NF-VINC" & xDocto2 & xSerie2 & xSubserie2 & xEmissao2 & xValor2 & "I"
                        Print #1, xlinha
                    ElseIf de_informa.rsSel_CTCEDIContabil.Fields("at_ctc_cif") = "A" Then
                        If de_informa.rsSel_CTCEDIContabil.Fields("tem_ocorr") = "C" Then
                        Else
                            xlinha = "NF-VINC" & xDocto2 & xSerie2 & xSubserie2 & xEmissao2 & xValor2 & "E"
                            Print #1, xlinha
                            xlinha = "NF-VINC" & xDocto2 & xSerie2 & xSubserie2 & xEmissao2 & xValor2 & "I"
                            Print #1, xlinha
                        End If
                    End If
                    
                    de_informa.rsSel_NFsdoCTC.MoveNext
                
                Loop
                
                de_informa.rsSel_CTCEDIContabil.MoveNext
                
            End If
            
        Loop
        
        de_informa.rsSel_CTCEDIContabil.MoveFirst
        
'        Do Until de_informa.rsSel_CTCEDIContabil.EOF
        
            'ATUALIZA EDI GERADO = S
'            de_informa.Alt_AtEdiCTC de_informa.rsSel_CTCEDIContabil.Fields("filialctc")
        
'            de_informa.rsSel_CTCEDIContabil.MoveNext
            
'        Loop
    
        Close #1
        MsgBox "OK ! Arquivo Gerado."
        
    
    Else
        MsgBox "Não Há Novas Faturas a Serem Atualizadas !"
        Exit Sub
    End If






End Sub
Private Sub cmdBonaFAT_Click()
    Dim xfilial As String, xFatura As String, xEmissao As String, xVencto As String, xvalor As String
    Dim xDesconto As String, xnome As String, xEndereco As String, xTelefone As String, xBairro As String
    Dim xCidade As String, xCep As String, xCnpj As String, xBanco As String, xAgencia As String, xNomeBanco As String
    Dim xSituacao As String, xlinha As String
    Dim xrs As Recordset
    
    'busca Faturas não atualizadas
    If de_informa.rsSel_FaturaEDIContabil.State = 1 Then de_informa.rsSel_FaturaEDIContabil.Close
    de_informa.Sel_FaturaEDIContabil
    
    If de_informa.rsSel_FaturaEDIContabil.RecordCount > 0 Then
        xqtdreg = de_informa.rsSel_FaturaEDIContabil.RecordCount
    
        'abre arquivo
        Open "C:\INFORMA\CONTABIL\M5FAT" & zeros(Day(datahora("DATA")), 2) & _
                                            zeros(Month(datahora("DATA")), 2) & _
                                            zeros(Hour(datahora("HORA")), 2) & _
                                            zeros(Minute(datahora("HORA")), 2) & ".TXT" For Output As #1
                                                
        Do Until de_informa.rsSel_FaturaEDIContabil.EOF
        
            If de_informa.rsSel_FaturaEDIContabil.Fields("status") = "C" And de_informa.rsSel_FaturaEDIContabil.Fields("at_edi_contabil") = "" Then
                de_informa.rsSel_FaturaEDIContabil.MoveNext
            Else
                xfilial = Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura"), 1, 2)
                xFatura = zeros(CDbl(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura"), 3, 6)), 10)
                xEmissao = zeros(Day(de_informa.rsSel_FaturaEDIContabil.Fields("emissao")), 2) & "/"
                xEmissao = xEmissao & zeros(Month(de_informa.rsSel_FaturaEDIContabil.Fields("emissao")), 2) & "/"
                xEmissao = xEmissao & Mid$(Trim$(Str(Year(de_informa.rsSel_FaturaEDIContabil.Fields("emissao")))), 3, 2)
                xVencto = zeros(Day(de_informa.rsSel_FaturaEDIContabil.Fields("vencimento")), 2) & "/"
                xVencto = xVencto & zeros(Month(de_informa.rsSel_FaturaEDIContabil.Fields("vencimento")), 2) & "/"
                xVencto = xVencto & Mid$(Trim$(Str(Year(de_informa.rsSel_FaturaEDIContabil.Fields("vencimento")))), 3, 2)
                xvalor = zeros(Int(de_informa.rsSel_FaturaEDIContabil.Fields("valorfatura") * 100), 14)
                xvalor = Mid$(xvalor, 1, 12) & "." & Mid$(xvalor, 13, 2)
                xDesconto = zeros(Int(de_informa.rsSel_FaturaEDIContabil.Fields("abatimento") * 100), 14)
                xDesconto = Mid$(xDesconto, 1, 12) & "." & Mid$(xDesconto, 13, 2)
                xnome = de_informa.rsSel_FaturaEDIContabil.Fields("cliente_nome") & _
                        String(40 - Len(de_informa.rsSel_FaturaEDIContabil.Fields("cliente_nome")), " ")
                xEndereco = de_informa.rsSel_FaturaEDIContabil.Fields("endcob") & _
                            String(48 - Len(de_informa.rsSel_FaturaEDIContabil.Fields("endcob")), " ")
                xTelefone = de_informa.rsSel_FaturaEDIContabil.Fields("telefonecob") & _
                            String(20 - Len(de_informa.rsSel_FaturaEDIContabil.Fields("telefonecob")), " ")
                xBairro = Space(15)
                xCidade = Mid(de_informa.rsSel_FaturaEDIContabil.Fields("cidadecob"), 1, 25) & _
                          String(25 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cidadecob"), 1, 25)), " ")
                xCep = Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 1, 5) & _
                       String(5 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 1, 5)), " ") & "-"
                xCep = xCep & Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 6, 3) & _
                       String(3 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 6, 3)), " ")
                xCnpj = Format(de_informa.rsSel_FaturaEDIContabil.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
                xBanco = zeros(CDbl(de_informa.rsSel_FaturaEDIContabil.Fields("banco")), 4)
                xAgencia = zeros(CDbl(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("conta"), 1, 4)), 4)
                xNomeBanco = Mid(de_informa.rsSel_FaturaEDIContabil.Fields("banconome"), 1, 10) & _
                             String(10 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("banconome"), 1, 10)), " ")
        
                'linha registro
                
                If de_informa.rsSel_FaturaEDIContabil.Fields("at_edi_contabil") = "" Then
                
                    xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                             xBairro & xCidade & xCep & xCnpj & xBanco & xAgencia & xNomeBanco & "I"
                
                    Print #1, xlinha
                
                ElseIf de_informa.rsSel_FaturaEDIContabil.Fields("at_edi_contabil") = "A" Then
                
                    If de_informa.rsSel_FaturaEDIContabil.Fields("status") = "C" Then
                
                        xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                                 xBairro & xCidade & xCep & xCnpj & xBanco & xAgencia & xNomeBanco & "E"
                                 
                        Print #1, xlinha
                        
                    Else
                    
                        xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                                 xBairro & xCidade & xCep & xCnpj & xBanco & xAgencia & xNomeBanco & "E"
                                 
                        Print #1, xlinha
                        
                        xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                                 xBairro & xCidade & xCep & xCnpj & xBanco & xAgencia & xNomeBanco & "I"
                        
                        Print #1, xlinha
                        
                    End If
                    
                End If
                    
                de_informa.rsSel_FaturaEDIContabil.MoveNext
                
            End If
        Loop
        
        de_informa.rsSel_FaturaEDIContabil.MoveFirst
        
        Do Until de_informa.rsSel_FaturaEDIContabil.EOF
        
            'ATUALIZA EDI GERADO = S
            de_informa.Alt_AtEdiFatura de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura")
            
            de_informa.rsSel_FaturaEDIContabil.MoveNext
            
        Loop
                    
        Close #1
        If xmsg = 0 Then
        MsgBox "OK ! Arquivo Gerado."
        End If
        
    Else
        If xmsg = 0 Then
        MsgBox "Não Há Novas Faturas a Serem Atualizadas !"
        End If
        
        Exit Sub
    End If

End Sub

Private Sub cmdCorreio_Click()
    Dim xdestinatario As String, xCep As String, xpeso As Currency, xanotacoes As String, xlinha As String, xfile As String
    
    Me.MousePointer = 11
    
    If de_informa.rsSel_BuscaCTCCorreios.State = 1 Then de_informa.rsSel_BuscaCTCCorreios.Close
    de_informa.Sel_BuscaCTCCorreios CDate(mskPer1Correio), CDate(mskPer2Correio)
    
    If de_informa.rsSel_BuscaCTCCorreios.RecordCount < 1 Then
    xqtdreg = de_informa.rsSel_BuscaCTCCorreios.RecordCount
        MsgBox "Não Há Dados para o Período Selecionado !"
        Exit Sub
    Else
    
        xfile = "INT" & zeros(Day(datahora("DATA")), 2) & _
                zeros(Month(datahora("DATA")), 2) & _
                Mid$(datahora("HORA"), 1, 2) & Mid$(datahora("HORA"), 4, 2) & ".TXT"
    
        Open "C:\INFORMA\CORREIOS\" & xfile For Output As #1
    
        Do Until de_informa.rsSel_BuscaCTCCorreios.EOF
        
            xlinha = ""
            
            xdestinatario = Trim$(de_informa.rsSel_BuscaCTCCorreios.Fields("dest_nome")) & _
                            String(50 - Len(Trim$(de_informa.rsSel_BuscaCTCCorreios.Fields("dest_nome"))), " ")
            xCep = de_informa.rsSel_BuscaCTCCorreios.Fields("dest_cep")
            If Len(Trim$(xCep)) < 8 Then
                If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
                de_informa.Sel_CadCliCGC de_informa.rsSel_BuscaCTCCorreios.Fields("dest_cgc")
                xCep = de_informa.rsSel_CadCliCGC.Fields("cep")
                If Len(Trim$(xCep)) < 8 Then
                    xCep = "        "
                End If
            End If
            
            xpeso = de_informa.rsSel_BuscaCTCCorreios.Fields("peso") * 1000
            xpesoCHAR = Format(xpeso, "#########0")
            xpesoCHAR = String(10 - Len(xpesoCHAR), "0") & xpesoCHAR
            
            xanotacoes = de_informa.rsSel_BuscaCTCCorreios.Fields("numnf") & "-" & _
                         de_informa.rsSel_BuscaCTCCorreios.Fields("serie")
                         
            xanotacoes = xanotacoes & String(50 - Len(Trim$(xanotacoes)), " ")
                         
            
            xlinha = xdestinatario & xCep & xpesoCHAR & xanotacoes
            
            Print #1, xlinha
            
            de_informa.rsSel_BuscaCTCCorreios.MoveNext

        Loop
        
        Close #1
        
        Me.MousePointer = 0
        
        MsgBox "Geração Finalizada !"
        
        
    End If
    
End Sub

Private Sub cmdDocCod_Click()
    Dim xarquivo As String, xrs As Recordset, xDataAgora As Date, xHoraAgora As Variant
    Dim xValorTot As Currency, xlinha As String, xnome As String
    
    If optPorCliente = True And Len(Trim$(txtCgcDocCob)) < 8 Then
        MsgBox "Número de CNPJ Inválido. Deve ter no mínimo os 8 primeiro dígitos do CNPJ !"
        Exit Sub
    ElseIf optPorFatura = True And Len(Trim$(txtFatura)) < 8 Then
        MsgBox "Número de Fatura Inválido. Deve ter 8 dígitos !"
        Exit Sub
    End If
    
    If optPorCliente.Value = True Then
        If de_informa.rsSel_EDICobCnpj.State = 1 Then de_informa.rsSel_EDICobCnpj.Close
        de_informa.Sel_EDICobCnpj txtCgcDocCob & "%"
        Set xrs = de_informa.rsSel_EDICobCnpj
    ElseIf optPorFatura.Value = True Then
        If de_informa.rsSel_EDICobFatura.State = 1 Then de_informa.rsSel_EDICobFatura.Close
        de_informa.Sel_EDICobFatura txtFatura
        Set xrs = de_informa.rsSel_EDICobFatura
    End If
    
    If xrs.RecordCount < 1 Then
        MsgBox "Não Há Dados Para Esta Seleção !", vbInformation
        Exit Sub
    End If
    
    cmdDocCod.Enabled = False
    cmdDocCod.Caption = "Aguarde..."
    
    xDataAgora = datahora("data")
    xHoraAgora = datahora("hora")
    
    'abre arquivo
    If txtCgcDocCob.Text = "04490850" Then
    xnome = "C:\INFORMA\GILLETTE\INTEFAT" & zeros(Day(xDataAgora), 2) & _
                                        zeros(Month(xDataAgora), 2)
    Else
    
    
    Open "C:\INFORMA\EDI_EXP\COBRANCA\INTECCOB_" & zeros(Day(xDataAgora), 2) & _
                                        zeros(Month(xDataAgora), 2) & _
                                        zeros(Hour(xHoraAgora), 2) & _
                                        zeros(Minute(xHoraAgora), 2) & ".TXT" For Output As #1
                                        
    xlinha = "000INTEC CARGO                        " & Mid$(Trim$(xrs.Fields("cliente_nome")), 1, 35) & _
             String(35 - Len(Mid$(Trim$(xrs.Fields("cliente_nome")), 1, 35)), " ") & zeros(Day(xDataAgora), 2) & _
             zeros(Month(xDataAgora), 2) & Mid$(Trim$(Year(xDataAgora)), 3, 2) & zeros(Hour(xHoraAgora), 2) & _
             zeros(Minute(xHoraAgora), 2) & "COB" & zeros(Day(xDataAgora), 2) & zeros(Month(xDataAgora), 2) & _
             zeros(Hour(xHoraAgora), 2) & zeros(Minute(xHoraAgora), 2) & "0" & Space(75)
            
    Print #1, xlinha
    
    xlinha = "350COBRA" & zeros(Day(xDataAgora), 2) & zeros(Month(xDataAgora), 2) & _
             zeros(Hour(xHoraAgora), 2) & zeros(Minute(xHoraAgora), 2) & "0" & Space(153)
    
    Print #1, xlinha
    
    If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
    de_informa.Sel_CadFilial Mid$(xrs.Fields("filialfatura"), 1, 2)
    
    xlinha = "351" & de_informa.rsSel_CadFilial.Fields("cgc") & "INTEC INTEGRACAO NACIONAL DE TRANSPORTES" & Space(113)
    
    Print #1, xlinha
    
    Do Until xrs.EOF
    
        xlinha = "352" & Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10) & _
                 String(10 - Len(Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10)), " ") & "0" & "U  " & _
                 zeros2(Mid$(xrs.Fields("filialfatura"), 3, 6), 10) & _
                 zeros(Day(xrs.Fields("emissao")), 2) & zeros(Month(xrs.Fields("emissao")), 2) & Trim$(Year(xrs.Fields("emissao"))) & _
                 zeros(Day(xrs.Fields("vencimento")), 2) & zeros(Month(xrs.Fields("vencimento")), 2) & Trim$(Year(xrs.Fields("vencimento"))) & _
                 zeros2(SoNumeros(Format(xrs.Fields("valorfatura"), "#########0.00")), 15) & "   " & _
                 zeros2(SoNumeros(Format(xrs.Fields("descicms"), "#########0.00")), 15) & _
                 "000000000000000" & "00000000" & zeros2(SoNumeros(Format(xrs.Fields("abatimento"), "#########0.00")), 15) & _
                 Trim$(xrs.Fields("banconome")) & String(35 - Len(Trim$(xrs.Fields("banconome"))), " ") & _
                 zeros2(Mid$(xrs.Fields("conta"), 1, InStr(1, xrs.Fields("conta"), ".") - 1), 4) & " " & _
                 zeros2(Mid$(xrs.Fields("conta"), InStr(1, xrs.Fields("conta"), ".") + 1, Abs(InStr(1, xrs.Fields("conta"), "-") - InStr(1, xrs.Fields("conta"), "."))), 10) & _
                 "  " & "I" & Space(3)
                 
        Print #1, xlinha
        
        If de_informa.rsSel_EDICobCTCs.State = 1 Then de_informa.rsSel_EDICobCTCs.Close
        de_informa.Sel_EDICobCTCs xrs.Fields("filialfatura")
        
        Do Until de_informa.rsSel_EDICobCTCs.EOF
        
            If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
            de_informa.Sel_CadFilial Mid$(de_informa.rsSel_EDICobCTCs.Fields("filialctc"), 1, 2)
            
            xlinha = "353" & Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10) & _
                     String(10 - Len(Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10)), " ") & _
                     "     " & zeros2(Mid$(de_informa.rsSel_EDICobCTCs.Fields("filialctc"), 3), 12) & Space(140)

            Print #1, xlinha
            
            If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
            de_informa.Sel_NFsdoCTC de_informa.rsSel_EDICobCTCs.Fields("filialctc")
            
            Do Until de_informa.rsSel_NFsdoCTC.EOF
            
                xlinha = "354" & de_informa.rsSel_NFsdoCTC.Fields("serie") & String(3 - Len(de_informa.rsSel_NFsdoCTC.Fields("serie")), " ") & _
                         zeros2(de_informa.rsSel_NFsdoCTC.Fields("numnf"), 8) & String(30, "0") & _
                         zeros2(de_informa.rsSel_NFsdoCTC.Fields("cliente_cgc"), 14) & Space(112)
            
                Print #1, xlinha
                
                de_informa.rsSel_NFsdoCTC.MoveNext
                
            Loop
            
            de_informa.rsSel_EDICobCTCs.MoveNext
            
        Loop
        
        xValorTot = xValorTot + xrs.Fields("valorfatura")
        
        xrs.MoveNext
        
    Loop
    
    xlinha = "355" & zeros(xrs.RecordCount, 4) & zeros2(SoNumeros(Format(xValorTot, "#########0.00")), 15) & Space(148)
    
    Print #1, xlinha
    
    Close #1
    
    xrs.MoveFirst
    
    Do Until xrs.EOF
        de_informa.Alt_EDICobATSim xrs.Fields("filialfatura")
        xrs.MoveNext
    Loop
    
    cmdDocCod.Enabled = True
    cmdDocCod.Caption = "Gerar Arquivo"
    
    MsgBox "Arquivo Gerado ! " + Chr(10) + Chr(13) + Chr(10) + Chr(13) + xarquivo, vbInformation
    End If
    
End Sub

Private Sub cmdGeraArqMedley_Click()
    
     
    Me.MousePointer = 11
    
    If de_informa.rsSel_EDICtcsMedley.State = 1 Then de_informa.rsSel_EDICtcsMedley.Close
    de_informa.Sel_EDICtcsMedley "50929710000179"

    If de_informa.rsSel_EDICtcsMedley.RecordCount > 0 Then
    
        'abre o arquivo
        
        Open "C:\INFORMA\EDI_EXP\MEDLEY\EDICONINT" & zeros(Day(datahora("data")), 2) & _
              zeros(Month(datahora("data")), 2) & Mid$(Trim$(Str(Year(datahora("data")))), 3, 2) & _
              ".TXT" _
        For Output As #1
        
        Do Until de_informa.rsSel_EDICtcsMedley.EOF
        
            'verifica o consignatário para confirmar que é da operação Bomi/Intec
            If Not (Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "50929710" Or _
                    Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "04019475" Or _
                    Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "52134798" Or _
                    Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "02426290") Then
                de_informa.rsSel_EDICtcsMedley.MoveNext
            Else
        
                'tratamento da série da NF
            
                If de_informa.rsSel_EDICtcsMedley.Fields("numnfnum") > 200000 Then
                    xserie = "1  "
                Else
                    xserie = "2  "
                End If
                
                xnumnf = zeros(de_informa.rsSel_EDICtcsMedley.Fields("numnfnum"), 6)
                
                If Not IsNull(de_informa.rsSel_EDICtcsMedley.Fields("pesonf")) Then
                    xpeso = zeros(de_informa.rsSel_EDICtcsMedley.Fields("pesonf") * 1000, 13)
                Else
                    xpeso = "0000000000000"
                End If
                
                xctc = "BOM" & Mid$(zeros(Val(de_informa.rsSel_EDICtcsMedley.Fields("ctc")), 8), 3, 6)
            
                If Not IsNull(de_informa.rsSel_EDICtcsMedley.Fields("volumesnf")) Then
                    xvolumes = zeros(de_informa.rsSel_EDICtcsMedley.Fields("volumesnf"), 6)
                Else
                    xvolumes = "000000"
                End If
                

            
                'data de embarque
                
                xdata = ""
                xdata = Trim$(Str(Year(de_informa.rsSel_EDICtcsMedley.Fields("data"))))
                xdata = xdata & zeros(Month(de_informa.rsSel_EDICtcsMedley.Fields("data")), 2)
                xdata = xdata & zeros(Day(de_informa.rsSel_EDICtcsMedley.Fields("data")), 2)
                
                'linha registro
                
                xlinha = "50" & "0001" & xserie & xnumnf & xpeso & xctc & xvolumes & xdata
                
                Print #1, xlinha
                
                'ATUALIZA EDI GERADO = S
                
                de_informa.alt_EDICtcMedleySim de_informa.rsSel_EDICtcsMedley.Fields("filialctc")
                
                de_informa.rsSel_EDICtcsMedley.MoveNext
                    
            End If
            DoEvents
        Loop
        Close #1
        If xmsg = 0 Then
        MsgBox "OK ! Geração de Arquivo EDI DE CONHECIMENTOS MEDLEY FINALIZADO."
        End If
    Else
        If xmsg = 0 Then
        MsgBox "Não Há Dados Para Geração de Arquivo EDI de Conhecimentos"
        End If
        
    End If
    
    Me.MousePointer = 0
    
End Sub
Private Sub cmdGeraBomi_Click()
    Dim xfrete As Currency, xfretechar As String, xdtctc As String, xhsctc As String, xdtmnf As String, xhsmnf As String
    Dim xmnf As String, xdtentrega As String, xhsentrega As String, xreceb As String, xnf As String, xlinha As String
    
    
    Me.MousePointer = 11
    
    If de_informa.rsSel_BossCTC.State = 1 Then de_informa.rsSel_BossCTC.Close
    de_informa.Sel_BossCTC CDate(mskPer1), CDate(mskPer2)
    
    If de_informa.rsSel_BossCTC.RecordCount < 1 Then
        MsgBox "Não Há Dados para o Período Selecionado !"
        Exit Sub
    Else
    
        Open "C:\INFORMA\BOMIBRASIL\INTEC.TXT" For Output As #1
    
        Do Until de_informa.rsSel_BossCTC.EOF
            xlinha = ""
        
            'dados do CTC/CTR
            
            If de_informa.rsSel_BossCiaCod.State = 1 Then de_informa.rsSel_BossCiaCod.Close
            de_informa.Sel_BossCiaCod Mid$(de_informa.rsSel_BossCTC.Fields("remet_cgc"), 1, 8)
            
            xcodcia = de_informa.rsSel_BossCiaCod.Fields("codigo")
            
            xnf = zeros(de_informa.rsSel_BossCTC.Fields("numnf"), 6)
            xserie = de_informa.rsSel_BossCTC.Fields("serie")
            
            If Len(xserie) = 0 Or Len(xserie) > 1 Then
                xserie = " "
            End If
            
            xdtctc = zeros(Day(de_informa.rsSel_BossCTC.Fields("data")), 2) & _
                     zeros(Month(de_informa.rsSel_BossCTC.Fields("data")), 2) & _
                     Mid$(Trim$(Str(Year(de_informa.rsSel_BossCTC.Fields("data")))), 3, 2)
                     
            If Mid$(de_informa.rsSel_BossCTC.Fields("hora"), 3, 1) = ":" Then
                xhsctc = Mid$(de_informa.rsSel_BossCTC.Fields("hora"), 1, 2) & _
                         Mid$(de_informa.rsSel_BossCTC.Fields("hora"), 4, 2) & "00"
            ElseIf Mid$(de_informa.rsSel_BossCTC.Fields("hora"), 2, 1) = ":" Then
                xhsctc = "0" & Mid$(de_informa.rsSel_BossCTC.Fields("hora"), 1, 1) & _
                         Mid$(de_informa.rsSel_BossCTC.Fields("hora"), 3, 2) & "00"
            Else
                xhsctc = "      "
            End If
            
                     
            xfrete = de_informa.rsSel_BossCTC.Fields("fretetotalbruto") * 100
            xfretechar = Format(xfrete, "##############0")
            xfretechar = String(15 - Len(xfretechar), "0") & xfretechar
        
            'dados de Manifesto
                    
            If de_informa.rsSel_BossMnf.State = 1 Then de_informa.rsSel_BossMnf.Close
            de_informa.Sel_BossMnf de_informa.rsSel_BossCTC.Fields("filialctc")
            
            If de_informa.rsSel_BossMnf.RecordCount < 1 Then
                xmnf = Space(8)
                xdtmnf = Space(6)
                xhsmnf = Space(6)
            Else
                xmnf = de_informa.rsSel_BossMnf.Fields("filialmanifesto")
                xdtmnf = zeros(Day(de_informa.rsSel_BossMnf.Fields("dtemissao")), 2) & _
                         zeros(Month(de_informa.rsSel_BossMnf.Fields("dtemissao")), 2) & _
                         Mid$(Trim$(Str(Year(de_informa.rsSel_BossMnf.Fields("dtemissao")))), 3, 2)
                         
                If de_informa.rsSel_BossMnf.Fields("hsemissao") = "" Then
                    xhsmnf = Space(6)
                Else
                    xhsmnf = Mid$(de_informa.rsSel_BossMnf.Fields("hsemissao"), 1, 2) & _
                             Mid$(de_informa.rsSel_BossMnf.Fields("hsemissao"), 4, 2) & "00"
                End If
            End If
            
            If de_informa.rsSel_BossCTC.Fields("tem_ocorr") = "1" Then
            
                If de_informa.rsSel_BossEntrega.State = 1 Then de_informa.rsSel_BossEntrega.Close
                de_informa.Sel_BossEntrega de_informa.rsSel_BossCTC.Fields("filialctc")
                
                xdtentrega = zeros(Day(de_informa.rsSel_BossEntrega.Fields("data")), 2) & _
                         zeros(Month(de_informa.rsSel_BossEntrega.Fields("data")), 2) & _
                         Mid$(Trim$(Str(Year(de_informa.rsSel_BossEntrega.Fields("data")))), 3, 2)
                
                If de_informa.rsSel_BossEntrega.Fields("hora") = "" Then
                    xhsentrega = Space(6)
                Else
                    xhsentrega = Mid$(de_informa.rsSel_BossEntrega.Fields("hora"), 1, 2) & _
                                 Mid$(de_informa.rsSel_BossEntrega.Fields("hora"), 4, 2) & "00"
                End If
                
                If IsNull(de_informa.rsSel_BossEntrega.Fields("receb")) Then
                    xreceb = Trim$(de_informa.rsSel_BossEntrega.Fields("recebpre")) & _
                             String(30 - Len(Trim$(de_informa.rsSel_BossEntrega.Fields("recebpre"))), " ")
                Else
                    If Len(Trim$(de_informa.rsSel_BossEntrega.Fields("recebpre"))) > _
                       Len(Trim$(de_informa.rsSel_BossEntrega.Fields("receb"))) Then
                        xreceb = Trim$(de_informa.rsSel_BossEntrega.Fields("recebpre")) & _
                                 String(30 - Len(Trim$(de_informa.rsSel_BossEntrega.Fields("recebpre"))), " ")
                    Else
                        xreceb = Trim$(de_informa.rsSel_BossEntrega.Fields("receb")) & _
                                 String(30 - Len(Trim$(de_informa.rsSel_BossEntrega.Fields("receb"))), " ")
                    End If
                End If
                
            Else
            
                xdtentrega = Space(6)
                xhsentrega = Space(6)
                xreceb = Space(30)
                
            End If
            
            'linha registro
            
            xlinha = xcodcia & xserie & xnf & de_informa.rsSel_BossCTC.Fields("filialctc") & xdtctc & xhsctc & _
                     xfretechar & xmnf & xdtmnf & xhsmnf & xdtentrega & xhsentrega & _
                     Mid$(de_informa.rsSel_BossCTC.Fields("modal"), 1, 1) & xreceb & Space(179)
            
            Print #1, xlinha
            
            de_informa.rsSel_BossCTC.MoveNext
            
        Loop
        
        Close #1
        
        MsgBox "Geração Finalizada !"
                
    End If
        
    Me.MousePointer = 0

End Sub

Private Sub cmdLerArquivoMedley_Click()
    Dim xlinha As String, xnf As Long, xvolumes As Integer, xpeso As Currency, xvalornf As Currency

     Open "C:\INFORMA\MEDLEY\NOTAS.TXT" For Input As #1
     lblOk3 = 0
     lblNaoOk3 = 0
     lblLida3 = 0
     Do Until EOF(1)
        Line Input #1, xlinha
        If Val(Mid$(xlinha, 1, 8)) > 0 Then
            lblLida3 = Val(lblLida3) + 1
            xnf = Val(Mid$(xlinha, 1, 8))
            xvolumes = Val(Mid$(xlinha, 9, 4))
            xpeso = Val(Mid$(xlinha, 13, 8)) / 10
            xvalornf = Val(Mid$(xlinha, 21, 16)) / 100
            If de_informa.rsSel_CgcNFEmissao.State = 1 Then de_informa.rsSel_CgcNFEmissao.Close
            de_informa.Sel_CgcNFEmissao "50929710%", xnf
            If de_informa.rsSel_CgcNFEmissao.RecordCount > 0 Then
                de_informa.alt_DadosMedleyNF xvolumes, xpeso, xvalornf, "50929710%", xnf
                lblOk3 = Val(lblOk3) + 1
            Else
                lblNaoOk3 = Val(lblNaoOk3) + 1
            End If
            DoEvents
        End If
     Loop
     MsgBox "Processo Finalizado !"
     Close #1

End Sub
Private Sub CmdProcessar_Click()
    Dim xlinha As String, xRemet_nome As String, xdata As String, xhora As String, xid_intercam As String
    Dim xnumnf As String, xobs_ocorr As String, xdataoco As String, xhoraoco As String, xserie As String
    Dim xarquivo As String
    Dim xnomearquivo As String
    Dim xnumarq As Integer
    
    
    Me.MousePointer = 11
    If de_informa.rsSel_EDI_Ocorr.State = 1 Then de_informa.rsSel_EDI_Ocorr.Close
    de_informa.Sel_EDI_Ocorr txtCgc & "%"
    
    If de_informa.rsSel_EDI_Ocorr.RecordCount > 0 Then
    
    xqtdreg = de_informa.rsSel_EDI_Ocorr.RecordCount
    
        'definição do diretório de gravação
        
        If Mid$(txtCgc, 1, 8) = "50929710" Then 'medley
            xarquivo = "C:\INFORMA\EDI_EXP\MEDLEY\EDIOCOINT" & zeros(Day(datahora("data")), 2) & zeros(Month(datahora("data")), 2) & Mid$(Trim$(Str(Year(datahora("data")))), 3, 2) & ".TXT"
        ElseIf Mid$(txtCgc, 1, 8) = "04490850" Then 'gillette
            xarquivo = "C:\INFORMA\GILLETTE\INTEOCO" & zeros(Day(datahora("data")), 2) & zeros(Month(datahora("data")), 2) & ".TXT"
        ElseIf Mid$(txtCgc, 1, 8) = "14372981" Then 'bayer
            xarquivo = "C:\INFORMA\EDI_EXP\BAYER\INTEOCO" & zeros(Day(datahora("data")), 2) & zeros(Month(datahora("data")), 2) & ".TXT"
        ElseIf Mid$(txtCgc, 1, 8) = "61188488" Then  'GIVAUDAN
            xarquivo = "C:\INFORMA\EDI_EXP\GIVAUDAN\EDIOCOINT" & zeros(Day(datahora("data")), 2) & zeros(Month(datahora("data")), 2) & Mid$(Trim$(Str(Year(datahora("data")))), 3, 2) & ".TXT"
        
        
        'SE FOR VIDEOLAR
        
        ElseIf Mid$(txtCgc, 1, 8) = "04229761" Then  'VIDEOLAR
            
            'PEGA O NUMERO MAX + 1 DE REGISTRO
            de_informa.rsSel_maxarq.Open
                        
            If IsNull(de_informa.rsSel_maxarq.Fields("MAX")) = True Then
                
                xnumarq = 1
                
            Else
            
                xnumarq = Trim(de_informa.rsSel_maxarq.Fields("MAX"))
            
            End If
            de_informa.rsSel_maxarq.Close
            
            'CRIA O NOME DO ARQUIVO
            xnomearquivo = "LFTOCO" & String(4 - Len(xnumarq), "0") & xnumarq & ".TXT"
            
            
            'CRIO O LOCAL NO QUAL SERÁ SALVO O ARQ.
            xarquivo = "C:\INFORMA\EDI_EXP\VIDEOLAR\" & xnomearquivo
            
            'ALTERA TB_MEM
             de_informa.up_tbmem xnomearquivo, xnumarq, de_informa.rsSel_EDI_Ocorr.RecordCount
             
             
        ElseIf Mid$(txtCgc, 1, 8) = "60412327" Then  'ALCON
            xarquivo = "C:\INFORMA\EDI_EXP\ALCON\INTEOCO" & zeros(Day(datahora("data")), 2) & zeros(Month(datahora("data")), 2) & ".TXT"
        ElseIf Mid$(txtCgc, 1, 8) = "33247743" Then  'GLAXO
            xarquivo = "C:\INFORMA\EDI_EXP\GLAXO\INTOCOREN" & zeros(Day(datahora("data")), 2) & zeros(Month(datahora("data")), 2) & ".TXT"
        ElseIf Mid$(txtCgc, 1, 8) = "60831658" Then  'BOEHRINGER
            xarquivo = "C:\INFORMA\EDI_EXP\BOEHRINGER\OC" & zeros(Day(datahora("data")), 2) & zeros(Month(datahora("data")), 2) & Mid$(Str(Year(datahora("data"))), 3, 2) & ".TXT"
        Else
            Me.MousePointer = 0
            MsgBox "CGC não Configurado para Geração de EDI de Ocorrências"
            Exit Sub
        End If
        
        Open xarquivo For Output As #1
        
        'tratamentos de dados para o arquivo (cabecários)
        
        'nome remetente / embarcador
        xRemet_nome = Trim$(de_informa.rsSel_EDI_Ocorr.Fields("remet_nome"))
        
        If Len(xRemet_nome) > 35 Then
            xRemet_nome = Mid$(xRemet_nome, 1, 35)
        ElseIf Len(xRemet_nome) < 35 Then
            xRemet_nome = xRemet_nome + Space(35 - Len(xRemet_nome))
        End If
        
        'data
        xdata = ""
        If Len(Trim$(Str(Day(datahora("data"))))) = 1 Then
            xdata = xdata & "0" & Trim$(Str(Day(datahora("data"))))
        Else
            xdata = xdata & Trim$(Str(Day(datahora("data"))))
        End If
        If Len(Trim$(Str(Month(datahora("data"))))) = 1 Then
            xdata = xdata & "0" & Trim$(Str(Month(datahora("data"))))
        Else
            xdata = xdata & Trim$(Str(Month(datahora("data"))))
        End If
        xdata = xdata & Trim$(Str(Year(datahora("data"))))
        
        'hora
        xhora = Mid(Trim$(Str(Time())), 1, 2) & Mid(Trim$(Str(Time())), 4, 2)
        
        'identif. de intercambio
        xid_intercam = "OCO" & Mid(xdata, 1, 4) & Mid(xhora, 1, 4) & "0"
        
        'REGISTRO 000
        
        xlinha = "000INTEC TRANSPORTES                  " & xRemet_nome & Mid(xdata, 1, 4) & Mid$(xdata, 7, 2) & xhora & xid_intercam & Space(25)
        Print #1, xlinha
        
        'REGISTRO 340
        
        xlinha = "340OCORR" & Mid$(xid_intercam, 4, 9) & Space(103)
        Print #1, xlinha
        
        'REGISTRO 341
        
        xlinha = "34152134798000320INTEC INTEGRACAO NAC TRANSP ENC. CARGAS" & Space(64)
        Print #1, xlinha
    
        'inicio do laço na recordset
        
        Do Until de_informa.rsSel_EDI_Ocorr.EOF
            
            If Mid$(txtCgc, 1, 8) = "50929710" And Not (Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "50929710" Or _
                                    Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "04019475" Or _
                                    Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "52134798" Or _
                                    Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "02426290") Then
                de_informa.rsSel_EDI_Ocorr.MoveNext
            Else

                'tratamento dos dados do detalhe (ocorrência)
            
                'número da NF
                xnumnf = String(8 - Len(Trim$(de_informa.rsSel_EDI_Ocorr.Fields("numnf"))), "0") & Trim$(de_informa.rsSel_EDI_Ocorr.Fields("numnf"))
        
                'observação de ocorrência
                If Not IsNull(de_informa.rsSel_EDI_Ocorr.Fields("obs_ocorr")) Then
                    xobs_ocorr = Trim$(de_informa.rsSel_EDI_Ocorr.Fields("obs_ocorr"))
                Else
                    xobs_ocorr = Space(70)
                End If
                If Len(xobs_ocorr) > 70 Then
                    xobs_ocorr = Mid$(xobs_ocorr, 1, 70)
                ElseIf Len(xobs_ocorr) < 70 Then
                    xobs_ocorr = xobs_ocorr + Space(70 - Len(xobs_ocorr))
                End If
            
                'data ocorrência
                xdataoco = ""
                If Len(Trim$(Str(Day(de_informa.rsSel_EDI_Ocorr.Fields("data"))))) = 1 Then
                    xdataoco = xdataoco & "0" & Trim$(Str(Day(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                Else
                    xdataoco = xdataoco & Trim$(Str(Day(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                End If
                If Len(Trim$(Str(Month(de_informa.rsSel_EDI_Ocorr.Fields("data"))))) = 1 Then
                    xdataoco = xdataoco & "0" & Trim$(Str(Month(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                Else
                    xdataoco = xdataoco & Trim$(Str(Month(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                End If
                xdataoco = xdataoco & Trim$(Str(Year(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
            
                'hora ocorrência
                xhoraoco = Mid$(de_informa.rsSel_EDI_Ocorr.Fields("hora"), 1, 2) & Mid$(de_informa.rsSel_EDI_Ocorr.Fields("hora"), 4, 2)
                If xhoraoco = "" Then
                    xhoraoco = "0000"
                End If
            
                'SERIE DA NF
                If Mid$(txtCgc, 1, 8) = "50929710" Then     'medley
                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
                        xserie = "   "
                    Else
                        xserie = Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))) & _
                                 String(3 - (Len(Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))))), " ")
                    End If
                ElseIf Mid$(txtCgc, 1, 8) = "04490850" Then 'gillette
                    xserie = "001"
                ElseIf Mid$(txtCgc, 1, 8) = "14372981" Then 'bayer
                    xserie = "001"
                ElseIf Mid$(txtCgc, 1, 8) = "61188488" Then 'bayer
                    xserie = "004"
                ElseIf Mid$(txtCgc, 1, 8) = "60412327" Then 'ALCON
                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
                        xserie = "01 "
                    Else
                        xserie = zeros(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")), 2) & " "
                    End If
                ElseIf Mid$(txtCgc, 1, 8) = "33247743" Then 'GLAXO
'                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
'                       xserie = "01 "
'                    Else
'                        xserie = zeros(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")), 2) & " "
'                    End If
                        xserie = "01 "
                Else
                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
                        xserie = "   "
                    Else
                        xserie = Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))) & _
                                 String(3 - (Len(Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))))), " ")
                    End If
                End If
                
                'REGISTRO 342

                xlinha = "342" & de_informa.rsSel_EDI_Ocorr.Fields("remet_cgc") & xserie & xnumnf & de_informa.rsSel_EDI_Ocorr.Fields("cod_ocorr") & _
                        xdataoco & xhoraoco & "00" & xobs_ocorr & Space(6)
                Print #1, xlinha
            
                'ATUALIZA EDI GERADO = S
            
                de_informa.Alt_EDI_Ocorr de_informa.rsSel_EDI_Ocorr.Fields("codigo")
            
                de_informa.rsSel_EDI_Ocorr.MoveNext
                
            End If
            
        Loop
        Close #1
        
    If xmsg = 0 Then
        MsgBox "OK ! Geração de Arquivo EDI FINALIZADO"
    End If
    
    
    Else
        
    If xmsg = 0 Then
        MsgBox "Não Há Dados Para Geração de Arquivo EDI de Ocorrências"
    End If
    
    
    End If
    Me.MousePointer = 0
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub cmdBuscaREM_Click()
    frmBuscaCLI.Caption = "Busca Cliente Exporta EDI"
    frmBuscaCLI.Show 1
End Sub


Private Sub chkTodosEstab_Click()
    If chkTodosEstab.Value = 1 Then
        TxtCGCRem.MaxLength = 8
    Else
        TxtCGCRem.MaxLength = 14
    End If
    TxtCGCRem.SetFocus
End Sub


P




Private Sub Form_Load()
If Weekday(Date) = 3 Then
    Label31.Visible = True
End If

End Sub

Private Sub optPorCliente_Click()
    If optPorCliente.Value = True Then
        txtCgcDocCob.Enabled = True
        txtCgcDocCob.BackColor = xamarelo1
        txtFatura.Enabled = False
        txtFatura.BackColor = xbranco
        txtCgcDocCob.SetFocus
    Else
        txtCgcDocCob.Enabled = False
        txtCgcDocCob.BackColor = xbranco
        txtFatura.Enabled = True
        txtFatura.BackColor = xamarelo1
        txtFatura.SetFocus
    End If
End Sub

Private Sub optPorFatura_Click()
     If optPorCliente.Value = True Then
        txtCgcDocCob.Enabled = True
        txtCgcDocCob.BackColor = xamarelo1
        txtFatura.Enabled = False
        txtFatura.BackColor = xbranco
        txtCgcDocCob.SetFocus
    Else
        txtCgcDocCob.Enabled = False
        txtCgcDocCob.BackColor = xbranco
        txtFatura.Enabled = True
        txtFatura.BackColor = xamarelo1
        txtFatura.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()

If Time = "08:00" Then
    cmd_todos_Click
    Exit Sub
End If


End Sub

Private Sub txtCGCRem_GotFocus()
    TxtCGCRem.SelStart = 0
    TxtCGCRem.SelLength = 14
End Sub

Private Sub txtCGCRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtCGCRem_LostFocus()
    If TxtCGCRem.Text = "%" Then
        lblNomeRem.Caption = "TODOS CLIENTES"
        Exit Sub
    End If
    If Len(Trim$(TxtCGCRem.Text)) <> TxtCGCRem.MaxLength And Len(Trim$(TxtCGCRem.Text)) > 0 Then
        MsgBox "Quantidade de Caracteres Inválida para este Número de CGC !"
        TxtCGCRem.SetFocus
        SendKeys "{END}"
    End If
    If TxtCGCRem.Text <> "" Then
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli Trim(TxtCGCRem) & "%"
        If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
            lblNomeRem.Caption = de_informa.rsSel_ConsCadCli.Fields("nome")
            CmdCONEMB.SetFocus
        Else
            TxtCGCRem.SetFocus
        End If
    Else
        lblNomeRem.Caption = ""
    End If
End Sub

Private Sub CmdCONEMB_Click()
Dim xRegID As String
Dim xAgora As Date
Dim xNomeFile As String
Dim Filial As String
Dim xserie As String
Dim CTC As String
Dim xdata As String
Dim xCondicaoFrete As String
Dim xpeso As String
Dim FreteTotal As String
Dim BaseCalc As String
Dim Aliquota As String
Dim ICMS As String
Dim FreteVolume As String
Dim FreteValor As String
Dim SecCat As String
Dim ITR As String
Dim Despacho As String
Dim Pedagio As String
Dim AdEme As String
Dim Subst As String
Dim CFOP As String
Dim CgcEmissor As String
Dim CgcEmbarc As String
Dim xAcao As String
Dim xTipoCon As String
Dim xFiller As String

Dim xRemID As String
Dim xDesID As String

Dim xDia As String
Dim xmes As String
Dim xano As String

Dim xH As String
Dim xM As String
Dim xhora As String
Dim xIntID As String
Dim xlinha As String
Dim xDocID As String
Dim xcgc As String
Dim xRazaoSocial As String

Dim xRecSize As Integer

Dim TotFrete As Long
Dim TotCTC As Long

Dim xTotFrete As String
Dim xTotCtc As String

Dim xarquivo As String

Dim dia As String
Dim mes As String
Dim ano As String

dia = Mid(Now, 1, 2)
mes = Mid(Now, 4, 2)
ano = Mid(Now, 9, 2)

If de_informa.rsCONEMBSel.State = 1 Then de_informa.rsCONEMBSel.Close
de_informa.CONEMBSel CDate(MskDataInicialCONEMB.Text), CDate(MskDataFinalCONEMB.Text), Trim(UCase(TxtCGCRem.Text)) & "%"
If de_informa.rsCONEMBSel.RecordCount > 0 Then
XNARQ = Int((Hour(Time) + Day(Date) + Year(Date) + Minute(Time)) / Second(Time))

xAgora = datahora("DATAHORA")
xNomeFile = Mid$(Trim(UCase(TxtCGCRem.Text)), 1, 8) & "_" & zeros(Day(xAgora), 2) & zeros(Month(xAgora), 2) & "_" & zeros(Hour(xAgora), 2) & zeros(Minute(xAgora), 2) & ".txt"

If TxtCGCRem = "04490850" Then
    xarquivo = "C:\INFORMA\GILLETTE\"
    xNomeFile = "INTECON" & dia & mes & ".TXT"
ElseIf TxtCGCRem = "60831658" Then
    xarquivo = "C:\INFORMA\EDI_EXP\BOEHRINGER\"
    xNomeFile = "CO" & dia & mes & ano & ".TXT"
Else
    xarquivo = "C:\INFORMA\EDI_EXP\PROCEDA\"
    
End If




    Open xarquivo & xNomeFile For Output As #1

    
xRecSize = 680

xRegID = "000"
xRemID = "INTEC TRANSPORTES"
xDesID = UCase(lblNomeRem.Caption)


xDia = String(2 - Len(Trim(Str(Day(Date)))), "0") & Trim(Str(Day(Date)))
xmes = String(2 - Len(Trim(Str(Month(Date)))), "0") & Trim(Str(Month(Date)))
xano = String(2 - Len(Mid(Trim(Str(Year(Date))), 3, 2)), "0") & Mid(Trim(Str(Year(Date))), 3, 2)
xdata = xDia & xmes & xano

xH = String(2 - Len(Trim(Str(Hour(Time)))), "0") & Trim(Str(Hour(Time)))
xM = String(2 - Len(Trim(Str(Minute(Time)))), "0") & Trim(Str(Minute(Time)))
xhora = xH & xM

xIntID = "CON" & xDia & xmes & xhora & "0"

xRegID = String(3 - Len(Trim(xRegID)), " ") & Trim(xRegID)
xRemID = Trim(xRemID) & String(35 - Len(Trim(xRemID)), " ")
xDesID = Trim(xDesID) & String(35 - Len(Trim(xDesID)), " ")
xdata = Trim(xdata) & String(6 - Len(Trim(xdata)), " ")
xhora = Trim(xhora) & String(4 - Len(Trim(xhora)), " ")
xIntID = Trim(xIntID) & String(12 - Len(Trim(xIntID)), " ")
xFiller = String(585, " ")
xlinha = xRegID & xRemID & xDesID & xdata & xhora & xIntID & xFiller
    If Len(xlinha) <> xRecSize Then
    MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
    Close #1
    Else
    Print #1, xlinha
    End If
    

    
xRegID = "320"
xDocID = "CONHE" & xDia & xmes & xhora & "1"
xFiller = String(663, " ")
xlinha = xRegID & xDocID & xFiller
    If Len(xlinha) <> xRecSize Then
    MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
    Close #1
    Else
    Print #1, xlinha
    End If
    
xRegID = "321"
xcgc = "52134798000320"
xRazaoSocial = "INTEC TRANSPORTES LTDA"

xRegID = Trim(xRegID) & String(3 - Len(Trim(xRegID)), " ")
xcgc = Trim(xcgc) & String(14 - Len(Trim(xcgc)), " ")
xRazaoSocial = Trim(xRazaoSocial) & String(40 - Len(Trim(xRazaoSocial)), " ")
xFiller = String(623, " ")
xlinha = xRegID & xcgc & xRazaoSocial & xFiller
    
    If Len(xlinha) <> xRecSize Then
    MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
    Close #1
    Else
    Print #1, xlinha
    End If
    
TotFrete = 0
TotCTC = 0

    Do Until de_informa.rsCONEMBSel.EOF
    
    TotFrete = TotFrete + de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
    TotCTC = TotCTC + 1
    
    xRegID = "322"
    Filial = de_informa.rsCONEMBSel.Fields("FILIAL")
    xserie = "U"
    CTC = de_informa.rsCONEMBSel.Fields("CTC")
    xDia = String(2 - Len(Trim(Str(Day(de_informa.rsCONEMBSel.Fields("DATA"))))), "0") & Trim(Str(Day(de_informa.rsCONEMBSel.Fields("DATA"))))
    xmes = String(2 - Len(Trim(Str(Month(de_informa.rsCONEMBSel.Fields("DATA"))))), "0") & Trim(Str(Month(de_informa.rsCONEMBSel.Fields("DATA"))))
    xano = String(4 - Len(Trim(Str(Year(de_informa.rsCONEMBSel.Fields("DATA"))))), "0") & Trim(Str(Year(de_informa.rsCONEMBSel.Fields("DATA"))))
    xdata = xDia & xmes & xano
    xCondicaoFrete = de_informa.rsCONEMBSel.Fields("FPAG")
        If xCondicaoFrete = "1-CIF" Then xCondicaoFrete = "C"
        If xCondicaoFrete = "2-FOB" Then xCondicaoFrete = "F"
        If xCondicaoFrete = "A PAGAR" Then xCondicaoFrete = "F"
        If xCondicaoFrete = "AGO   G" Then xCondicaoFrete = "C"
        If xCondicaoFrete = "PAGO" Then xCondicaoFrete = "C"
        
    xpeso = de_informa.rsCONEMBSel.Fields("PESO")
    
    If Mid$(de_informa.rsCONEMBSel.Fields("respons_cgc"), 1, 8) = "04490850" Then
        If de_informa.rsCONEMBSel.Fields("subtrib") = "S" Then
            FreteTotal = de_informa.rsCONEMBSel.Fields("FRETETOTAL")
        Else
            FreteTotal = de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
        End If
    Else
        FreteTotal = de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
    End If
    
    BaseCalc = de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
    Aliquota = de_informa.rsCONEMBSel.Fields("ALIQUOTA") * 100
    ICMS = FreteTotal * (Aliquota / 100)
    FreteVolume = 0
    FreteValor = de_informa.rsCONEMBSel.Fields("FRETEVALORBR")
    SecCat = de_informa.rsCONEMBSel.Fields("TXCOLETABR") + de_informa.rsCONEMBSel.Fields("TXENTREGAredbr")
    ITR = "0"
    Despacho = "0"
    Pedagio = de_informa.rsCONEMBSel.Fields("PEDAGIOBR")
    AdEme = "0"
    Subst = de_informa.rsCONEMBSel.Fields("SUBTRIB")
        If Subst = "S" Then Subst = "1"
        If Subst = "N" Then Subst = "2"
    CFOP = Mid(de_informa.rsCONEMBSel.Fields("CFOP"), 1, 3)
    CgcEmissor = "52134798000320"
    CgcEmbarc = de_informa.rsCONEMBSel.Fields("REMET_CGC")
        If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
        de_informa.Sel_NFsdoCTC transctc(Filial, CTC)
        
    Dim NFs(1 To 40, 1 To 2) As String
        For k = 1 To 40
            If de_informa.rsSel_NFsdoCTC.EOF = False Then
            NFs(k, 1) = de_informa.rsSel_NFsdoCTC.Fields("SERIE")
            NFs(k, 2) = de_informa.rsSel_NFsdoCTC.Fields("NUMNF")
            de_informa.rsSel_NFsdoCTC.MoveNext
            Else
            NFs(k, 1) = ""
            NFs(k, 2) = ""
            End If
                
            If Val(NFs(k, 1)) > 0 Then
            NFs(k, 1) = String(3 - Len(Trim(Str(Val(NFs(k, 1))))), "0") & Trim(Str(Val(NFs(k, 1))))
            Else
            NFs(k, 1) = String(3 - Len(Trim(NFs(k, 1))), " ") & Trim(NFs(k, 1))
            End If
            
        NFs(k, 2) = String(8 - Len(Trim(Str(Val(NFs(k, 2))))), "0") & Trim(Str(Val(NFs(k, 2))))
        Next
    Dim StringNFs As String
    StringNFs = ""
        For k = 1 To 40
        StringNFs = StringNFs & NFs(k, 1) & NFs(k, 2)
        Next
    
    xAcao = "I"
    xTipoCon = "N"
        
    xpeso = Trim(Str(Int(Val(xpeso * 100))))
    FreteTotal = Trim(Str(Int(Val(FreteTotal * 100))))
    BaseCalc = Trim(Str(Int(Val(BaseCalc * 100))))
    Aliquota = Trim(Str(Int(Val(Aliquota * 100))))
    ICMS = Trim(Str(Int(Val(xicms * 100))))
    FreteVolume = Trim(Str(Int(Val(FreteVolume * 100))))
    FreteValor = Trim(Str(Int(Val(FreteValor * 100))))
    SecCat = Trim(Str(Int(Val(SecCat * 100))))
    ITR = Trim(Str(Int(Val(ITR * 100))))
    Despacho = Trim(Str(Int(Val(Despacho * 100))))
    Pedagio = Trim(Str(Int(Val(pedadio * 100))))
    AdEme = Trim(Str(Int(Val(AdEme * 100))))
    
    xRegID = Trim(xRegID)
    Filial = Trim(Filial)
    xserie = Trim(xserie)
    CTC = Trim(CTC)
    xdata = Trim(xdata)
    xCondicaoFrete = Trim(xCondicaoFrete)
    xpeso = Trim(xpeso)
    FreteTotal = Trim(FreteTotal)
    BaseCalc = Trim(BaseCalc)
    Aliquota = Trim(Aliquota)
    ICMS = Trim(ICMS)
    FreteVolume = Trim(FreteVolume)
    FreteValor = Trim(FreteValor)
    SecCat = Trim(SecCat)
    ITR = Trim(ITR)
    Despacho = Trim(Despacho)
    Pedagio = Trim(Pedagio)
    AdEme = Trim(AdEme)
    Subst = Trim(Subst)
    CFOP = Trim(CFOP)
    CgcEmissor = Trim(CgcEmissor)
    CgcEmbarc = Trim(CgcEmbarc)
    xAcao = Trim(xAcao)
    xTipoCon = Trim(xTipoCon)
    
    xRegID = xRegID & String(3 - Len(xRegID), " ")
    Filial = Filial & String(10 - Len(Filial), " ")
    xserie = xserie & String(5 - Len(xserie), " ")
    CTC = CTC & String(12 - Len(CTC), " ")
    xdata = xdata & String(8 - Len(xdata), " ")
    xCondicaoFrete = xCondicaoFrete & String(1 - Len(xCondicaoFrete), " ")
    xpeso = String(7 - Len(xpeso), "0") & xpeso
    FreteTotal = String(15 - Len(FreteTotal), "0") & FreteTotal
    BaseCalc = String(15 - Len(BaseCalc), "0") & BaseCalc
    Aliquota = String(4 - Len(Aliquota), "0") & Aliquota
    ICMS = String(15 - Len(ICMS), "0") & ICMS
    FreteVolume = String(15 - Len(FreteVolume), "0") & FreteVolume
    FreteValor = String(15 - Len(FreteValor), "0") & FreteValor
    SecCat = String(15 - Len(SecCat), "0") & SecCat
    ITR = String(15 - Len(ITR), "0") & ITR
    Despacho = String(15 - Len(Despacho), "0") & Despacho
    Pedagio = String(15 - Len(Pedagio), "0") & Pedagio
    AdEme = String(15 - Len(AdEme), "0") & AdEme
    Subst = Subst & String(1 - Len(Subst), " ")
    CFOP = CFOP & String(3 - Len(CFOP), " ")
    CgcEmissor = CgcEmissor & String(14 - Len(CgcEmissor), " ")
    CgcEmbarc = CgcEmbarc & String(14 - Len(CgcEmbarc), " ")
    xAcao = xAcao & String(1 - Len(xAcao), " ")
    xTipoCon = xTipoCon & String(1 - Len(xTipoCon), " ")
    xFiller = String(6, " ")

    xlinha = xRegID & Filial & xserie & CTC & xdata & xCondicaoFrete & xpeso & FreteTotal & BaseCalc & Aliquota & ICMS & FreteVolume & FreteValor & SecCat & ITR & Despacho & Pedagio & AdEme & Subst & CFOP & CgcEmissor & CgcEmbarc & StringNFs & xAcao & xTipoCon & xFiller
        
        If Len(xlinha) <> xRecSize Then
        MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
        Close #1
        Else
        Print #1, xlinha
        End If
    
    de_informa.rsCONEMBSel.MoveNext
    Loop

xRegID = "323"
xTotFrete = Trim(Str(Int(TotFrete * 100)))
xTotCtc = Trim(Str(TotCTC))

xTotCtc = String(4 - Len(xTotCtc), "0") & xTotCtc
xTotFrete = String(15 - Len(xTotFrete), "0") & xTotFrete
xFiller = String(658, " ")

xlinha = xRegID & xTotCtc & xTotFrete & xFiller

Print #1, xlinha
    
Close #1
MsgBox "Arquivo C:\INFORMA\EDI_EXP\PROCEDA\" & xNomeFile & " gerado com sucesso!", vbInformation
Else
MsgBox "Sua Pesquisa não retornou dado algum. Por favor, tente novamente.", vbInformation
End If
    


    
End Sub
