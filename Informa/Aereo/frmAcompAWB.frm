VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAcompAWB 
   Caption         =   "Acompanhamento de AWB"
   ClientHeight    =   6015
   ClientLeft      =   1935
   ClientTop       =   3165
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Icon            =   "frmAcompAWB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   435
      Left            =   5280
      TabIndex        =   122
      Top             =   7740
      Width           =   3255
   End
   Begin VB.Frame FraOBS 
      Caption         =   "Observações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   5220
      TabIndex        =   35
      Top             =   3660
      Width           =   6615
      Begin VB.TextBox TxtOBS 
         BackColor       =   &H00C0FFFF&
         Height          =   3615
         Left            =   120
         MaxLength       =   239
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   435
      Left            =   8580
      TabIndex        =   21
      Top             =   7740
      Width           =   3255
   End
   Begin VB.Frame FraVoo 
      Caption         =   "Vôo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   107
      Top             =   3660
      Width           =   5055
      Begin VB.Frame FraCON 
         Caption         =   "Conexão"
         Enabled         =   0   'False
         Height          =   1935
         Left            =   120
         TabIndex        =   116
         Top             =   1260
         Width           =   4815
         Begin VB.TextBox TxtConVoo 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1215
            TabIndex        =   10
            Top             =   900
            Width           =   1170
         End
         Begin VB.TextBox TxtBuscaCON 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   540
            Width           =   675
         End
         Begin VB.TextBox TxtSiglaCON 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   124
            Top             =   540
            Width           =   495
         End
         Begin VB.TextBox TxtAeroportoCON 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1380
            TabIndex        =   123
            Top             =   540
            Width           =   3315
         End
         Begin VB.TextBox TxtConHoraCheg 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1215
            MaxLength       =   5
            TabIndex        =   12
            Top             =   1500
            Width           =   1170
         End
         Begin VB.TextBox TxtConHoraPart 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3525
            MaxLength       =   5
            TabIndex        =   14
            Top             =   1500
            Width           =   1170
         End
         Begin MSMask.MaskEdBox MskConPart 
            Height          =   285
            Left            =   3525
            TabIndex        =   13
            Top             =   1200
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskConCheg 
            Height          =   285
            Left            =   1215
            TabIndex        =   11
            Top             =   1200
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nº Vôo"
            Height          =   195
            Left            =   600
            TabIndex        =   125
            Top             =   960
            Width           =   510
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Hora Cheg."
            Height          =   195
            Left            =   315
            TabIndex        =   121
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Dt. Chegada"
            Height          =   195
            Left            =   225
            TabIndex        =   120
            Top             =   1260
            Width           =   900
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Hora Partida"
            Height          =   195
            Left            =   2580
            TabIndex        =   119
            Top             =   1560
            Width           =   885
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data Partida"
            Height          =   195
            Left            =   2580
            TabIndex        =   118
            Top             =   1260
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Aeroporto da Conexão"
            Height          =   195
            Left            =   120
            TabIndex        =   117
            Top             =   300
            Width           =   1590
         End
      End
      Begin VB.CheckBox ChkCon 
         Caption         =   "Vôo Com Conexão"
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
         Left            =   180
         TabIndex        =   8
         Top             =   1020
         Width           =   3135
      End
      Begin MSMask.MaskEdBox MskDataPart 
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Top             =   600
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin VB.Frame FraRetira 
         Caption         =   "Cliente Retirou Mercadoria?"
         Height          =   735
         Left            =   120
         TabIndex        =   113
         Top             =   3720
         Width           =   4815
         Begin VB.TextBox TxtVolumesRetira 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3540
            MaxLength       =   12
            TabIndex        =   19
            Top             =   300
            Width           =   1155
         End
         Begin VB.OptionButton OptRetiraNao 
            Caption         =   "Não"
            Height          =   195
            Left            =   1020
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptRetiraSim 
            Caption         =   "Sim"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Volumes Retirados"
            Height          =   195
            Left            =   2100
            TabIndex        =   114
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.TextBox TxtVoo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1260
         TabIndex        =   5
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox TxtHoraPart 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   7
         Top             =   600
         Width           =   1170
      End
      Begin VB.TextBox TxtHoraCheg 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   16
         Top             =   3300
         Width           =   1170
      End
      Begin MSMask.MaskEdBox MskDataCheg 
         Height          =   285
         Left            =   1260
         TabIndex        =   15
         Top             =   3300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Partida"
         Height          =   195
         Left            =   300
         TabIndex        =   112
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nº Vôo"
         Height          =   195
         Left            =   675
         TabIndex        =   111
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora Partida"
         Height          =   195
         Left            =   2640
         TabIndex        =   110
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Chegada"
         Height          =   195
         Left            =   180
         TabIndex        =   109
         Top             =   3360
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora Chegada"
         Height          =   195
         Left            =   2520
         TabIndex        =   108
         Top             =   3360
         Width           =   1035
      End
   End
   Begin VB.Frame FraEmissao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   99
      Top             =   2940
      Width           =   11715
      Begin VB.TextBox TxtHora 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6420
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtEmissor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   240
         Width           =   2715
      End
      Begin VB.TextBox TxtEmissao 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Left            =   6000
         TabIndex        =   106
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Emissor"
         Height          =   195
         Left            =   120
         TabIndex        =   104
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   3840
         TabIndex        =   103
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "Dados do Expedidor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   3300
      TabIndex        =   74
      Top             =   60
      Width           =   8535
      Begin VB.TextBox TxtNome 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   87
         Top             =   240
         Width           =   7335
      End
      Begin VB.TextBox TxtCGC 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   86
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtInscrEst 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   85
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtEnd 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   84
         Top             =   1380
         Width           =   4395
      End
      Begin VB.TextBox TxtCEP 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   83
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtTel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3060
         TabIndex        =   82
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtFAX 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   81
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtApolice 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   80
         Top             =   1980
         Width           =   1755
      End
      Begin VB.TextBox TxtSeguradora 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   79
         Top             =   1980
         Width           =   3615
      End
      Begin VB.TextBox TxtBairro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   78
         Top             =   1380
         Width           =   1875
      End
      Begin VB.CommandButton CmdDados 
         Caption         =   ">"
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
         Index           =   1
         Left            =   8160
         TabIndex        =   77
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox TxtCidade 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   76
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox TxtUF 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   6540
         TabIndex        =   75
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   98
         Top             =   285
         Width           =   420
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   97
         Top             =   2325
         Width           =   405
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   96
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "I.E."
         Height          =   195
         Index           =   1
         Left            =   3660
         TabIndex        =   95
         Top             =   2325
         Width           =   240
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "End."
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   94
         Top             =   1425
         Width           =   330
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   7020
         X2              =   120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FAX:"
         Height          =   195
         Index           =   1
         Left            =   4905
         TabIndex        =   93
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   92
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Index           =   1
         Left            =   2355
         TabIndex        =   91
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seguradora:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   90
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Apólice:"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   89
         Top             =   2025
         Width           =   570
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Dados Complementares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2310
         TabIndex        =   88
         Top             =   780
         Width           =   2520
      End
   End
   Begin VB.Frame FraAWB 
      Caption         =   "AWB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   70
      Top             =   1500
      Width           =   2895
      Begin VB.TextBox TxtDig 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   115
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox TxtFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         TabIndex        =   72
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox TxtSiglaCiaAerea 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   660
         TabIndex        =   71
         Top             =   300
         Width           =   435
      End
      Begin VB.Label TxtAWB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   73
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "Dados do Destinatário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   3300
      TabIndex        =   22
      Top             =   780
      Width           =   8535
      Begin VB.TextBox TxtUF 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6540
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TxtCidade 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   25
         Top             =   1080
         Width           =   5775
      End
      Begin VB.CommandButton CmdDados 
         Caption         =   ">"
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
         Index           =   0
         Left            =   8160
         TabIndex        =   51
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox TxtBairro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   45
         Top             =   1380
         Width           =   1875
      End
      Begin VB.TextBox TxtSeguradora 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1020
         TabIndex        =   44
         Top             =   1980
         Width           =   3615
      End
      Begin VB.TextBox TxtApolice 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   43
         Top             =   1980
         Width           =   1755
      End
      Begin VB.TextBox TxtFAX 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   42
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtTel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3060
         TabIndex        =   41
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtCEP 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   40
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtEnd 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   38
         Top             =   1380
         Width           =   4395
      End
      Begin VB.TextBox TxtInscrEst 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   36
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtCGC 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   26
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtNome 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Dados Complementares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   2310
         TabIndex        =   58
         Top             =   780
         Width           =   2520
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Apólice:"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   57
         Top             =   2025
         Width           =   570
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seguradora:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Index           =   0
         Left            =   2355
         TabIndex        =   55
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   54
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FAX:"
         Height          =   195
         Index           =   0
         Left            =   4905
         TabIndex        =   53
         Top             =   1740
         Width           =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   7020
         X2              =   120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "End."
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   39
         Top             =   1425
         Width           =   330
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "I.E."
         Height          =   195
         Index           =   0
         Left            =   3660
         TabIndex        =   37
         Top             =   2325
         Width           =   240
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   29
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   28
         Top             =   2325
         Width           =   405
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   27
         Top             =   285
         Width           =   420
      End
   End
   Begin VB.Frame FraBuscaMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   120
      TabIndex        =   59
      Top             =   60
      Width           =   3135
      Begin VB.Frame FraBuscaAWB 
         Caption         =   "Número e Dígito do AWB"
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
         Left            =   120
         TabIndex        =   60
         Top             =   180
         Width           =   2895
         Begin VB.TextBox TxtBuscaDig 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2340
            MaxLength       =   1
            TabIndex        =   3
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox TxtBuscaAWB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   960
            MaxLength       =   10
            TabIndex        =   2
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox TxtBuscaSiglaAWB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   540
            MaxLength       =   2
            TabIndex        =   1
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox TxtBuscaAWBFilial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   0
            Top             =   240
            Width           =   435
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   2895
      End
   End
   Begin VB.Frame FraAeoportos 
      Caption         =   "Aeroportos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3060
      TabIndex        =   30
      Top             =   1500
      Width           =   8775
      Begin VB.TextBox TxtAeroportoVIA 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   31
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox TxtSiglaVIA 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   32
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox TxtAeroportoDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4140
         TabIndex        =   50
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox TxtSiglaDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3660
         TabIndex        =   49
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox TxtAeroportoExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   46
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox TxtSiglaExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   660
         TabIndex        =   47
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Left            =   3060
         TabIndex        =   52
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Origem"
         Height          =   195
         Left            =   105
         TabIndex        =   48
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "VIA"
         Height          =   195
         Left            =   6045
         TabIndex        =   33
         Top             =   345
         Width           =   255
      End
   End
   Begin VB.Frame FraEspecie 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   34
      Top             =   2220
      Width           =   11715
      Begin VB.TextBox TxtModal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   10260
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtClienteRetira 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8820
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtDescrIATA 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   240
         Width           =   3795
      End
      Begin VB.TextBox TxtPerecivel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtEspecie 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Modal Frete"
         Height          =   195
         Left            =   9360
         TabIndex        =   69
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Retira?"
         Height          =   195
         Left            =   7740
         TabIndex        =   67
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Perecível?"
         Height          =   195
         Left            =   6180
         TabIndex        =   64
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Espécie"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   300
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmAcompAWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public COMBODIGESP As Integer

Private Sub ChkCon_Click()
    If ChkCon.Value = 1 Then
    FraCON.Enabled = True
    TxtBuscaCON.Text = ""
    TxtSiglaCON.Text = ""
    TxtAeroportoCON.Text = ""
    TxtConVoo.Text = ""
    MskConCheg.Text = ""
    MskConPart.Text = ""
    TxtConHoraCheg.Text = ""
    TxtConHoraPart.Text = ""
    TxtBuscaCON.BackColor = xAmarelo
    MskConCheg.BackColor = xAmarelo
    MskConPart.BackColor = xAmarelo
    TxtConHoraCheg.BackColor = xAmarelo
    TxtConHoraPart.BackColor = xAmarelo
    TxtConVoo.BackColor = xAmarelo
    TxtBuscaCON.Enabled = True
    MskConCheg.Enabled = True
    MskConPart.Enabled = True
    TxtConHoraCheg.Enabled = True
    TxtConHoraPart.Enabled = True
    TxtConVoo.Enabled = True
    TxtBuscaCON.SetFocus
    Else
    FraCON.Enabled = False
    TxtBuscaCON.Text = ""
    TxtSiglaCON.Text = ""
    TxtAeroportoCON.Text = ""
    TxtConVoo.Text = ""
    MskConCheg.Text = ""
    MskConPart.Text = ""
    TxtConHoraCheg.Text = ""
    TxtConHoraPart.Text = ""
    TxtBuscaCON.BackColor = xBranco
    MskConCheg.BackColor = xBranco
    MskConPart.BackColor = xBranco
    TxtConHoraCheg.BackColor = xBranco
    TxtConHoraPart.BackColor = xBranco
    TxtConVoo.BackColor = xBranco
    TxtBuscaCON.Enabled = False
    MskConCheg.Enabled = False
    MskConPart.Enabled = False
    TxtConHoraCheg.Enabled = False
    TxtConHoraPart.Enabled = False
    TxtConVoo.Enabled = False
    End If
    
End Sub

Private Sub CmdBuscar_Click()
If de_informa.rsSelAWB.State = 1 Then de_informa.rsSelAWB.Close
If de_informa.rsSelAWBVoo.State = 1 Then de_informa.rsSelAWBVoo.Close

        If Len(Trim(TxtBuscaAWBFilial.Text)) = 0 Or Len(Trim(TxtBuscaSiglaAWB.Text)) = 0 Or Len(Trim(TxtBuscaAWB.Text)) = 0 Or Len(Trim(TxtBuscaDig.Text)) = 0 Then
        Exit Sub
        End If
        
    xAWB = Trim(TxtBuscaAWB.Text)
    xDig = Trim(TxtBuscaDig.Text)
    xCodAwb = String(2 - Len(Trim(Str(Val(TxtBuscaAWBFilial.Text)))), "0") & Trim(Str(Val(TxtBuscaAWBFilial.Text))) & UCase(Trim(TxtBuscaSiglaAWB.Text)) & String(10 - Len(Trim(Str(Val(xAWB)))), "0") & Trim(Str(Val(xAWB))) & Trim(Str(Val(xDig)))
    
        
        de_informa.SelAWB xCodAwb
        de_informa.SelAWBVoo xCodAwb
        
        If de_informa.rsSelAWB.RecordCount > 0 Then
        Call LimpaTela(Me)
        TxtFilial.Text = de_informa.rsSelAWB.Fields("filial")
        TxtSiglaCiaAerea.Text = de_informa.rsSelAWB.Fields("cia")
        TxtAWB.Caption = de_informa.rsSelAWB.Fields("awb")
        TxtDig.Text = de_informa.rsSelAWB.Fields("dig")
        TxtSiglaExpedidor.Text = de_informa.rsSelAWB.Fields("siglaorigem")
        TxtSiglaVIA.Text = de_informa.rsSelAWB.Fields("siglavia")
        TxtSiglaDestinatario.Text = de_informa.rsSelAWB.Fields("siglades")
        
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadeorigem")) Then TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadeorigem")) & " - " & de_informa.rsSelAWB.Fields("uforigem") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportoorigem")) & ")"
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadevia")) Then TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadevia")) & " - " & de_informa.rsSelAWB.Fields("ufvia") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportovia")) & ")"
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadedestino")) Then TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadedestino")) & " - " & de_informa.rsSelAWB.Fields("ufdestino") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportodestino")) & ")"
        
        TxtEspecie.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("especie"))
        TxtDescrIATA.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("descrprodsis"))
        
            If de_informa.rsSelAWB.Fields("perecivel") = "S" Then
            TxtPerecivel.Text = "S"
            Else
            TxtPerecivel.Text = "N"
            End If
            
            If de_informa.rsSelAWB.Fields("retira") = "S" Then
            TxtClienteRetira.Text = "S"
            Else
            TxtClienteRetira.Text = "N"
            End If
            
        TxtModal.Text = de_informa.rsSelAWB.Fields("modal")
        TxtEmissor.Text = de_informa.rsSelAWB.Fields("emissor")
        TxtEmissao.Text = de_informa.rsSelAWB.Fields("data")
        TxtHora.Text = de_informa.rsSelAWB.Fields("hora")
        
            If de_informa.rsSelAWB.Fields("cancelado") = "X" Then
            TxtStatus.Text = "AWB Cancelado"
            Else
            TxtStatus.Text = ""
            End If
        
        
        X = 1
        
        If Not IsNull(de_informa.rsSelAWB.Fields("nomeexp")) Then TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("nomeexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("endexp")) Then TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("endexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("bairroexp")) Then TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("bairroexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadexp")) Then TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("ufexp")) Then TxtUF(X).Text = de_informa.rsSelAWB.Fields("ufexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("telexp")) Then TxtTel(X).Text = de_informa.rsSelAWB.Fields("telexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("faxexp")) Then TxtFAX(X).Text = de_informa.rsSelAWB.Fields("faxexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("cnpjexp")) Then TxtCGC(X).Text = de_informa.rsSelAWB.Fields("cnpjexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("inscrestexp")) Then TxtInscrEst(X).Text = de_informa.rsSelAWB.Fields("inscrestexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("cepexp")) Then TxtCEP(X).Text = de_informa.rsSelAWB.Fields("cepexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("segexp")) Then TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("segexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("apoliceexp")) Then TxtApolice(X).Text = de_informa.rsSelAWB.Fields("apoliceexp")
        
        X = 0
        
        If Not IsNull(de_informa.rsSelAWB.Fields("nomedes")) Then TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("nomedes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("enddes")) Then TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("enddes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("bairrodes")) Then TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("bairrodes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadedes")) Then TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadedes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("ufdes")) Then TxtUF(X).Text = de_informa.rsSelAWB.Fields("ufdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("teldes")) Then TxtTel(X).Text = de_informa.rsSelAWB.Fields("teldes")
        If Not IsNull(de_informa.rsSelAWB.Fields("faxdes")) Then TxtFAX(X).Text = de_informa.rsSelAWB.Fields("faxdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("cnpjdes")) Then TxtCGC(X).Text = de_informa.rsSelAWB.Fields("cnpjdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("inscrestdes")) Then TxtInscrEst(X).Text = de_informa.rsSelAWB.Fields("inscrestdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("cepdes")) Then TxtCEP(X).Text = de_informa.rsSelAWB.Fields("cepdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("segdes")) Then TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("segdes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("apolicedes")) Then TxtApolice(X).Text = de_informa.rsSelAWB.Fields("apolicedes")
        
        CmdGravar.Enabled = True
        TxtVoo.SetFocus
        End If
        
        If de_informa.rsSelAWBVoo.RecordCount > 0 Then
        TxtVoo.Text = de_informa.rsSelAWBVoo.Fields("voo")
        MskDataCheg.Text = de_informa.rsSelAWBVoo.Fields("data_chegada")
        TxtHoraCheg.Text = de_informa.rsSelAWBVoo.Fields("hora_chegada")
        MskDataPart.Text = de_informa.rsSelAWBVoo.Fields("data_partida")
        TxtHoraPart.Text = de_informa.rsSelAWBVoo.Fields("data_partida")
        TxtVolumesRetira.Text = de_informa.rsSelAWBVoo.Fields("volumesretirados")
        TxtOBS.Text = de_informa.rsSelAWBVoo.Fields("obs")
            If de_informa.rsSelAWBVoo.Fields("clienteretirou") = "S" Then
            OptRetiraSim.Value = True
            Else
            OptRetiraNao.Value = True
            End If
        TxtConVoo.Text = de_informa.rsSelAWBVoo.Fields("convoo")
        TxtSiglaCON.Text = de_informa.rsSelAWBVoo.Fields("consigla")
        TxtAeroportoCON.Text = PriMaiuscula(de_informa.rsSelAWBVoo.Fields("concidade")) & " - " & de_informa.rsSelAWBVoo.Fields("conuf") & " (" & PriMaiuscula(de_informa.rsSelAWBVoo.Fields("conaeroporto")) & ")"
        MskConCheg.Text = de_informa.rsSelAWBVoo.Fields("condtcheg")
        TxtConHoraCheg.Text = de_informa.rsSelAWBVoo.Fields("conhoracheg")
        MskConPart.Text = de_informa.rsSelAWBVoo.Fields("condtpart")
        TxtConHoraPart.Text = de_informa.rsSelAWBVoo.Fields("conhorapart")
        End If
End Sub

Private Sub CmdDados_Click(Index As Integer)

Dim xFrame As Frame
Dim Botao As CommandButton
Dim HMax As Integer
Dim HMin As Integer
Set Botao = CmdDados(Index)

Set xFrame = Fra(Index)
HMax = 2715
HMin = 675

    If xFrame.Height = HMin Then
    xFrame.ZOrder (0)
    DoEvents
    Call TravaFrame(frmConsultaAWB, xFrame, 0)
    xFrame.Height = HMax
    Botao.Caption = "<"
    DoEvents
    ElseIf xFrame.Height = HMax Then
    Call TravaFrame(frmConsultaAWB, xFrame, 1)
    xFrame.Height = HMin
    Botao.Caption = ">"
    DoEvents
    End If

End Sub

Private Sub cmdGravar_Click()

    If Len(Trim(TxtVoo.Text)) = 0 Then
    MsgBox "É necessário que você informe o número do Vôo.", vbCritical, ""
    Exit Sub
    End If
    
    If MsgBox("Você confirma a inclusão destes dados?", vbYesNo + vbQuestion, "") = vbNo Then
    Exit Sub
    End If

CmdGravar.Enabled = False

xAWB = Trim(TxtAWB.Caption)
xDig = Trim(TxtDig.Text)
xCodAwb = String(2 - Len(Trim(Str(Val(TxtFilial.Text)))), "0") & Trim(Str(Val(TxtFilial.Text))) & UCase(Trim(TxtSiglaCiaAerea.Text)) & String(10 - Len(Trim(Str(Val(xAWB)))), "0") & Trim(Str(Val(xAWB))) & Trim(Str(Val(xDig)))
'xCodAwb = TransCodAWB(TxtFilial.Text, xtSiglaCiaAerea.Text, xawb, xDig)

xconsigla = UCase(TxtSiglaCON.Text)
    If Len(Trim(TxtAeroportoCON.Text)) > 0 Then
    XCONCIDADE = UCase(Trim(Mid(TxtAeroportoCON.Text, 1, InStr(1, TxtAeroportoCON.Text, "-", vbTextCompare) - 1)))
    xconuf = UCase(Trim(Mid(TxtAeroportoCON.Text, InStr(1, TxtAeroportoCON.Text, "-", vbTextCompare) + 1, 3)))
    xconAeroporto = UCase(Trim(Mid(TxtAeroportoCON.Text, InStr(1, TxtAeroportoCON.Text, "(", vbTextCompare) + 1, Len(TxtAeroportoCON.Text) - (InStr(1, TxtAeroportoCON.Text, "(", vbTextCompare) + 1))))
    Else
    XCONCIDADE = ""
    xconuf = ""
    xconAeroporto = ""
    End If


xRetira = ""
    If OptRetiraSim.Value = True Then xRetira = "S"

If de_informa.rsSelAWBVoo.State = 1 Then de_informa.rsSelAWBVoo.Close
de_informa.SelAWBVoo xCodAwb

    If de_informa.rsSelAWBVoo.RecordCount > 0 Then
    de_informa.UpdateAWBVoo xCodAwb, Trim(UCase(TxtVoo.Text)), _
    MskDataPart.Text, TxtHoraPart.Text, _
    Trim(UCase(TxtConVoo.Text)), xconsigla, XCONCIDADE, xconuf, xconaereporto, MskConCheg.Text, TxtConHoraCheg, MskConPart.Text, TxtConHoraPart.Text, _
    MskDataCheg.Text, TxtHoraCheg.Text, xRetira, Trim(TxtVolumesRetira.Text), _
    Trim(UCase(TxtOBS.Text)), CDate(DataHora("DATA")), xUsuario
    Else
    de_informa.InsertAWBVoo xCodAwb, Trim(UCase(TxtVoo.Text)), _
    MskDataPart.Text, TxtHoraPart.Text, _
    Trim(UCase(TxtConVoo.Text)), xconsigla, XCONCIDADE, xconuf, xconaereporto, MskConCheg.Text, TxtConHoraCheg, MskConPart.Text, TxtConHoraPart.Text, _
    MskDataCheg.Text, TxtHoraCheg.Text, xRetira, Trim(TxtVolumesRetira.Text), _
    Trim(UCase(TxtOBS.Text)), CDate(DataHora("DATA")), xUsuario
    End If

Call LimpaTela(Me)

TxtBuscaAWBFilial.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub


Private Sub OptAWB_Click()
    If OptAWB.Value = True Then
    FraAWB.Visible = True
    FraNF.Visible = False
    FraCTC.Visible = False
    ElseIf OptCTC.Value = True Then
    FraAWB.Visible = False
    FraNF.Visible = False
    FraCTC.Visible = True
    ElseIf OptNF.Value = True Then
    FraAWB.Visible = False
    FraNF.Visible = True
    FraCTC.Visible = False
    End If
DoEvents
End Sub

Private Sub optCTC_Click()
    If OptAWB.Value = True Then
    FraAWB.Visible = True
    FraNF.Visible = False
    FraCTC.Visible = False
    ElseIf OptCTC.Value = True Then
    FraAWB.Visible = False
    FraNF.Visible = False
    FraCTC.Visible = True
    ElseIf OptNF.Value = True Then
    FraAWB.Visible = False
    FraNF.Visible = True
    FraCTC.Visible = False
    End If
DoEvents
End Sub

Private Sub optNf_Click()
    If OptAWB.Value = True Then
    FraAWB.Visible = True
    FraNF.Visible = False
    FraCTC.Visible = False
    ElseIf OptCTC.Value = True Then
    FraAWB.Visible = False
    FraNF.Visible = False
    FraCTC.Visible = True
    ElseIf OptNF.Value = True Then
    FraAWB.Visible = False
    FraNF.Visible = True
    FraCTC.Visible = False
    End If
DoEvents
End Sub

Private Sub TxtADValorem_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtADValorem.Text = SoNumero(TxtADValorem.Text)
            TxtADValorem.Text = Mid(TxtADValorem.Text, 1, Len(TxtADValorem.Text) - 1)
            TxtADValorem.Text = Val(TxtADValorem.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtADValorem.Text = SoNumero(TxtADValorem.Text) & Chr(KeyAscii)
    TxtADValorem.Text = Val(TxtADValorem.Text) / 100
    KeyAscii = 0
    End If

TxtADValorem.Text = Format(TxtADValorem.Text, "###,###,###,##0.00")
TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")
End Sub

Private Sub Form_Activate()
If xACOMP = True Then
TxtBuscaAWBFilial.Text = frmAIRAcompanha.FlexAWB.TextMatrix(frmAIRAcompanha.FlexAWB.Row, 1)
TxtBuscaSiglaAWB.Text = frmAIRAcompanha.FlexAWB.TextMatrix(frmAIRAcompanha.FlexAWB.Row, 2)
TxtBuscaAWB.Text = frmAIRAcompanha.FlexAWB.TextMatrix(frmAIRAcompanha.FlexAWB.Row, 3)
TxtBuscaDig.Text = frmAIRAcompanha.FlexAWB.TextMatrix(frmAIRAcompanha.FlexAWB.Row, 4)
Call CmdBuscar_Click
xACOMP = False
End If
End Sub

Private Sub MskDataCheg_GotFocus()
Call Date_MskEdBox_GotFocus(MskDataCheg)
End Sub

Private Sub MskDataCheg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub MskDataCheg_LostFocus()
Call Date_MskEdBox_LostFocus(MskDataCheg)
End Sub

Private Sub MskConCheg_GotFocus()
Call Date_MskEdBox_GotFocus(MskConCheg)
End Sub

Private Sub MskConCheg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub MskConCheg_LostFocus()
Call Date_MskEdBox_LostFocus(MskConCheg)
End Sub

Private Sub MskCONPART_GotFocus()
Call Date_MskEdBox_GotFocus(MskConPart)
End Sub

Private Sub MskCONPART_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub MskCONPART_LostFocus()
Call Date_MskEdBox_LostFocus(MskConPart)
End Sub


Private Sub MskDataPart_GotFocus()
Call Date_MskEdBox_GotFocus(MskDataPart)
End Sub

Private Sub MskDataPart_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub MskDataPart_LostFocus()
Call Date_MskEdBox_LostFocus(MskDataPart)
End Sub


Private Sub OptRetiraNao_Click()
    If OptRetiraSim.Value = True Then
    TxtVolumesRetira.Enabled = True
    TxtVolumesRetira.BackColor = xAmarelo
    TxtVolumesRetira.SetFocus
    Else
    TxtVolumesRetira.Text = ""
    TxtVolumesRetira.Enabled = False
    TxtVolumesRetira.BackColor = xBranco
    End If
End Sub

Private Sub OptRetiraNao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub OptRetiraSim_Click()
    If OptRetiraSim.Value = True Then
    TxtVolumesRetira.Enabled = True
    TxtVolumesRetira.BackColor = xAmarelo
    TxtVolumesRetira.SetFocus
    Else
    TxtVolumesRetira.Text = ""
    TxtVolumesRetira.Enabled = False
    TxtVolumesRetira.BackColor = xBranco
    End If
End Sub

Private Sub OptRetiraSim_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtBuscaAWB_GotFocus()
TxtBuscaAWB.SelStart = 0
TxtBuscaAWB.SelLength = 100
End Sub

Private Sub TxtBuscaAWB_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtBuscaAWBFilial_GotFocus()
TxtBuscaAWBFilial.SelStart = 0
TxtBuscaAWBFilial.SelLength = 100
End Sub

Private Sub TxtBuscaAWBFilial_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtBuscaCTC_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtBuscaCON_Change()
Dim X As Integer
X = TxtBuscaCON.SelStart
TxtBuscaCON.Text = UCase(TxtBuscaCON.Text)
TxtBuscaCON.SelStart = X
End Sub

Private Sub TxtBuscaCON_LostFocus()
With TxtBuscaCON
    If Len(Trim(.Text)) > 0 Then
    If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
    If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
    
    de_informa.SelAeroportoSigla .Text & "%"
        If de_informa.rsSelAeroportoSigla.RecordCount > 0 Then
        TxtSiglaCON.Text = de_informa.rsSelAeroportoSigla.Fields("sigla")
        TxtAeroportoCON.Text = PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("localidade")) & " - " & de_informa.rsSelAeroportoSigla.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("aeroporto")) & ")"
        Else
        If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
        de_informa.SelAeroportoCidade .Text & "%"
            If de_informa.rsSelAeroportoCidade.RecordCount > 0 Then
            TxtSiglaCON.Text = de_informa.rsSelAeroportoCidade.Fields("sigla")
            TxtAeroportoCON.Text = PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & de_informa.rsSelAeroportoCidade.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("aeroporto")) & ")"
            Else
            MsgBox "Termo não encontrado!", vbCritical, ""
            End If
        End If
    End If
End With
End Sub

Private Sub TxtBuscacon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub


Private Sub TxtBuscaDig_GotFocus()
TxtBuscaDig.SelStart = 0
TxtBuscaDig.SelLength = 100
End Sub

Private Sub TxtBuscaDig_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtBuscaNF_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub


Private Sub TxtBuscaSiglaAWB_Change()
Dim X As Integer
X = TxtBuscaSiglaAWB.SelStart
TxtBuscaSiglaAWB.Text = UCase(TxtBuscaSiglaAWB.Text)
TxtBuscaSiglaAWB.SelStart = X
End Sub

Private Sub TxtBuscaSiglaAWB_GotFocus()
TxtBuscaSiglaAWB.SelStart = 0
TxtBuscaSiglaAWB.SelLength = 100
End Sub

Private Sub TxtBuscaSiglaAWB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub


Private Sub TxtDescrIATA_GotFocus()
TxtDescrIATA.SelStart = 0
TxtDescrIATA.SelLength = 200
End Sub

Private Sub TxtDescrIATA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmEmissaoCODSIATA.Show 1
ChkPerecivel.SetFocus
LblAtualizarFrete.Caption = "Sim"
End If
End Sub

Private Sub TxtDescrIATA_LostFocus()
'If Len(Trim(TxtDescrIATA.Text)) = 0 Then

'End If
End Sub

Private Sub TxtFreteNacional_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtFreteNacional.Text = SoNumero(TxtFreteNacional.Text)
            TxtFreteNacional.Text = Mid(TxtFreteNacional.Text, 1, Len(TxtFreteNacional.Text) - 1)
            TxtFreteNacional.Text = Val(TxtFreteNacional.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtFreteNacional.Text = SoNumero(TxtFreteNacional.Text) & Chr(KeyAscii)
    TxtFreteNacional.Text = Val(TxtFreteNacional.Text) / 100
    KeyAscii = 0
    End If

TxtFreteNacional.Text = Format(TxtFreteNacional.Text, "###,###,###,##0.00")
TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")
End Sub


Private Sub TxtOBSEmissao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub TxtTXDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtTXDestino.Text = SoNumero(TxtTXDestino.Text)
            TxtTXDestino.Text = Mid(TxtTXDestino.Text, 1, Len(TxtTXDestino.Text) - 1)
            TxtTXDestino.Text = Val(TxtTXDestino.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtTXDestino.Text = SoNumero(TxtTXDestino.Text) & Chr(KeyAscii)
    TxtTXDestino.Text = Val(TxtTXDestino.Text) / 100
    KeyAscii = 0
    End If

TxtTXDestino.Text = Format(TxtTXDestino.Text, "###,###,###,##0.00")
TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")

End Sub


Private Sub TxtVolumes_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii = 8 Then
        Else
        KeyAscii = 0
        End If
    End If
End Sub



Private Sub TxtBuscaFilial_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub


Private Sub TxtDescrOutros1_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtDescrOutros2_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtBuscaFilial_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtDescrOutros1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtDescrOutros2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtOBSEmissao_Change()
If Len(Trim(TxtOBSEmissao.Text)) > 0 Then
TxtOBSEmissao.Text = UCase(TxtOBSEmissao.Text)
TxtOBSEmissao.SelStart = Len(TxtOBSEmissao.Text)
End If
End Sub

Private Sub TxtOutros1_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtOutros1.Text = SoNumero(TxtOutros1.Text)
            TxtOutros1.Text = Mid(TxtOutros1.Text, 1, Len(TxtOutros1.Text) - 1)
            TxtOutros1.Text = Val(TxtOutros1.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtOutros1.Text = SoNumero(TxtOutros1.Text) & Chr(KeyAscii)
    TxtOutros1.Text = Val(TxtOutros1.Text) / 100
    KeyAscii = 0
    End If

TxtOutros1.Text = Format(TxtOutros1.Text, "###,###,###,##0.00")
TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")
End Sub

Private Sub Txtoutros2_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtOutros2.Text = SoNumero(TxtOutros2.Text)
            TxtOutros2.Text = Mid(TxtOutros2.Text, 1, Len(TxtOutros2.Text) - 1)
            TxtOutros2.Text = Val(TxtOutros2.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtOutros2.Text = SoNumero(TxtOutros2.Text) & Chr(KeyAscii)
    TxtOutros2.Text = Val(TxtOutros2.Text) / 100
    KeyAscii = 0
    End If

TxtOutros2.Text = Format(TxtOutros2.Text, "###,###,###,##0.00")
TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")
End Sub

Private Sub TxtPesoReal_KeyPress(KeyAscii As Integer)
LblAtualizarFrete.Caption = "Sim"
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtPesoReal.Text = SoNumero(TxtPesoReal.Text)
            TxtPesoReal.Text = Mid(TxtPesoReal.Text, 1, Len(TxtPesoReal.Text) - 1)
            TxtPesoReal.Text = Val(TxtPesoReal.Text) / 10
            KeyAscii = 0
            End If
        End If
    Else
    TxtPesoReal.Text = SoNumero(TxtPesoReal.Text) & Chr(KeyAscii)
    TxtPesoReal.Text = Val(TxtPesoReal.Text) / 10
    KeyAscii = 0
    End If

TxtPesoReal.Text = Format(TxtPesoReal.Text, "###,###,###,##0.0")
End Sub



Private Sub TxtTXRedesp_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtTXRedesp.Text = SoNumero(TxtTXRedesp.Text)
            TxtTXRedesp.Text = Mid(TxtTXRedesp.Text, 1, Len(TxtTXRedesp.Text) - 1)
            TxtTXRedesp.Text = Val(TxtTXRedesp.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtTXRedesp.Text = SoNumero(TxtTXRedesp.Text) & Chr(KeyAscii)
    TxtTXRedesp.Text = Val(TxtTXRedesp.Text) / 100
    KeyAscii = 0
    End If

TxtTXRedesp.Text = Format(TxtTXRedesp.Text, "###,###,###,##0.00")
TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")
End Sub

Private Sub TxtCONHoraCheg_GotFocus()
    If Len(Trim(TxtConHoraCheg.Text)) = 0 Then
    TxtConHoraCheg.Text = "__:__"
    End If
End Sub

Private Sub TxtCONHoraCheg_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    ElseIf KeyAscii = 8 Then
    TxtConHoraCheg.Text = "__:__"
    Else
        If TxtConHoraCheg.Text = "__:__" Then
            If Chr(KeyAscii) > 2 Then
            KeyAscii = 0
            Else
            TxtConHoraCheg.Text = Chr(KeyAscii) & "_:__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtConHoraCheg.Text, 2) = "_:__" Then
            If Mid(TxtConHoraCheg.Text, 1, 1) = "2" Then
                If Chr(KeyAscii) > 3 Then
                KeyAscii = 0
                Else
                TxtConHoraCheg.Text = Mid(TxtConHoraCheg.Text, 1, 1) & Chr(KeyAscii) & ":__"
                KeyAscii = 0
                End If
            Else
            TxtConHoraCheg.Text = Mid(TxtConHoraCheg.Text, 1, 1) & Chr(KeyAscii) & ":__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtConHoraCheg.Text, 3) = ":__" Then
            If Chr(KeyAscii) > 5 Then
            KeyAscii = 0
            Else
            TxtConHoraCheg.Text = Mid(TxtConHoraCheg.Text, 1, 3) & Chr(KeyAscii) & "_"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtConHoraCheg.Text, 5) = "_" Then
            TxtConHoraCheg.Text = Mid(TxtConHoraCheg.Text, 1, 4) & Chr(KeyAscii) & "_"
            KeyAscii = 0
        End If
    KeyAscii = 0
    End If
End Sub

Private Sub TxtCONHoraCheg_LostFocus()
    If InStr(1, TxtConHoraCheg.Text, "_", vbTextCompare) > 0 Then
    TxtConHoraCheg.Text = ""
    End If
End Sub


Private Sub TxtCONHoraPART_GotFocus()
    If Len(Trim(TxtConHoraPart.Text)) = 0 Then
    TxtConHoraPart.Text = "__:__"
    End If
End Sub

Private Sub TxtCONHoraPART_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    ElseIf KeyAscii = 8 Then
    TxtConHoraPart.Text = "__:__"
    Else
        If TxtConHoraPart.Text = "__:__" Then
            If Chr(KeyAscii) > 2 Then
            KeyAscii = 0
            Else
            TxtConHoraPart.Text = Chr(KeyAscii) & "_:__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtConHoraPart.Text, 2) = "_:__" Then
            If Mid(TxtConHoraPart.Text, 1, 1) = "2" Then
                If Chr(KeyAscii) > 3 Then
                KeyAscii = 0
                Else
                TxtConHoraPart.Text = Mid(TxtConHoraPart.Text, 1, 1) & Chr(KeyAscii) & ":__"
                KeyAscii = 0
                End If
            Else
            TxtConHoraPart.Text = Mid(TxtConHoraPart.Text, 1, 1) & Chr(KeyAscii) & ":__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtConHoraPart.Text, 3) = ":__" Then
            If Chr(KeyAscii) > 5 Then
            KeyAscii = 0
            Else
            TxtConHoraPart.Text = Mid(TxtConHoraPart.Text, 1, 3) & Chr(KeyAscii) & "_"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtConHoraPart.Text, 5) = "_" Then
            TxtConHoraPart.Text = Mid(TxtConHoraPart.Text, 1, 4) & Chr(KeyAscii) & "_"
            KeyAscii = 0
        End If
    KeyAscii = 0
    End If
End Sub

Private Sub TxtCONHoraPART_LostFocus()
    If InStr(1, TxtConHoraPart.Text, "_", vbTextCompare) > 0 Then
    TxtConHoraPart.Text = ""
    End If
End Sub

Private Sub TxtHoraCheg_GotFocus()
    If Len(Trim(TxtHoraCheg.Text)) = 0 Then
    TxtHoraCheg.Text = "__:__"
    End If
End Sub

Private Sub TxtHoraCheg_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    ElseIf KeyAscii = 8 Then
    TxtHoraCheg.Text = "__:__"
    Else
        If TxtHoraCheg.Text = "__:__" Then
            If Chr(KeyAscii) > 2 Then
            KeyAscii = 0
            Else
            TxtHoraCheg.Text = Chr(KeyAscii) & "_:__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHoraCheg.Text, 2) = "_:__" Then
            If Mid(TxtHoraCheg.Text, 1, 1) = "2" Then
                If Chr(KeyAscii) > 3 Then
                KeyAscii = 0
                Else
                TxtHoraCheg.Text = Mid(TxtHoraCheg.Text, 1, 1) & Chr(KeyAscii) & ":__"
                KeyAscii = 0
                End If
            Else
            TxtHoraCheg.Text = Mid(TxtHoraCheg.Text, 1, 1) & Chr(KeyAscii) & ":__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHoraCheg.Text, 3) = ":__" Then
            If Chr(KeyAscii) > 5 Then
            KeyAscii = 0
            Else
            TxtHoraCheg.Text = Mid(TxtHoraCheg.Text, 1, 3) & Chr(KeyAscii) & "_"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHoraCheg.Text, 5) = "_" Then
            TxtHoraCheg.Text = Mid(TxtHoraCheg.Text, 1, 4) & Chr(KeyAscii) & "_"
            KeyAscii = 0
        End If
    KeyAscii = 0
    End If
End Sub

Private Sub TxtHoraCheg_LostFocus()
    If InStr(1, TxtHoraCheg.Text, "_", vbTextCompare) > 0 Then
    TxtHoraCheg.Text = ""
    End If
End Sub

Private Sub TxtHoraPart_GotFocus()
    If Len(Trim(TxtHoraPart.Text)) = 0 Then
    TxtHoraPart.Text = "__:__"
    End If
End Sub

Private Sub TxtHoraPart_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    ElseIf KeyAscii = 8 Then
    TxtHoraPart.Text = "__:__"
    Else
        If TxtHoraPart.Text = "__:__" Then
            If Chr(KeyAscii) > 2 Then
            KeyAscii = 0
            Else
            TxtHoraPart.Text = Chr(KeyAscii) & "_:__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHoraPart.Text, 2) = "_:__" Then
            If Mid(TxtHoraPart.Text, 1, 1) = "2" Then
                If Chr(KeyAscii) > 3 Then
                KeyAscii = 0
                Else
                TxtHoraPart.Text = Mid(TxtHoraPart.Text, 1, 1) & Chr(KeyAscii) & ":__"
                KeyAscii = 0
                End If
            Else
            TxtHoraPart.Text = Mid(TxtHoraPart.Text, 1, 1) & Chr(KeyAscii) & ":__"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHoraPart.Text, 3) = ":__" Then
            If Chr(KeyAscii) > 5 Then
            KeyAscii = 0
            Else
            TxtHoraPart.Text = Mid(TxtHoraPart.Text, 1, 3) & Chr(KeyAscii) & "_"
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHoraPart.Text, 5) = "_" Then
            TxtHoraPart.Text = Mid(TxtHoraPart.Text, 1, 4) & Chr(KeyAscii) & "_"
            KeyAscii = 0
        End If
    KeyAscii = 0
    End If
End Sub

Private Sub TxtHoraPart_LostFocus()
    If InStr(1, TxtHoraPart.Text, "_", vbTextCompare) > 0 Then
    TxtHoraPart.Text = ""
    End If
End Sub



Private Sub TxtOBS_Change()
Dim X As Integer
X = TxtOBS.SelStart
TxtOBS.Text = UCase(TxtOBS.Text)
TxtOBS.SelStart = X
End Sub

Private Sub TxtOBS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtVolumesRetira_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtVoo_Change()
Dim X As Integer
X = TxtVoo.SelStart
TxtVoo.Text = UCase(TxtVoo.Text)
TxtVoo.SelStart = X
End Sub

Private Sub TxtVoo_GotFocus()
TxtVoo.SelStart = 0
TxtVoo.SelLength = 12
End Sub

Private Sub TxtVoo_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtconVoo_Change()
Dim X As Integer
X = TxtConVoo.SelStart
TxtConVoo.Text = UCase(TxtConVoo.Text)
TxtConVoo.SelStart = X
End Sub

Private Sub TxtconVoo_GotFocus()
TxtConVoo.SelStart = 0
TxtConVoo.SelLength = 12
End Sub

Private Sub TxtconVoo_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

