VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissao 
   Caption         =   "Emissão de AWB"
   ClientHeight    =   8535
   ClientLeft      =   300
   ClientTop       =   2010
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoAir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.Frame FraExpedidor 
      Caption         =   "Expedidor/Origem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   3480
      TabIndex        =   36
      Top             =   120
      Width           =   4035
      Begin VB.TextBox TxtUFExpedidor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   45
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton CmdDadosExpedidor 
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
         Left            =   3660
         TabIndex        =   142
         Top             =   180
         Width           =   255
      End
      Begin VB.TextBox TxtBairroEXP 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2940
         TabIndex        =   140
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox TxtCidadeExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   43
         Top             =   1380
         Width           =   2775
      End
      Begin VB.TextBox TxtSeguradoraExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   130
         Top             =   3060
         Width           =   2715
      End
      Begin VB.TextBox TxtApoliceExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   129
         Top             =   3360
         Width           =   2715
      End
      Begin VB.TextBox TxtFAXExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   121
         Top             =   2760
         Width           =   2715
      End
      Begin VB.TextBox TxtTelExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   120
         Top             =   2460
         Width           =   2715
      End
      Begin VB.TextBox TxtCEPExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   119
         Top             =   2160
         Width           =   2715
      End
      Begin VB.TextBox TxtEndExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   117
         Top             =   1080
         Width           =   2235
      End
      Begin VB.TextBox TxtInscrEstExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   115
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox TxtBuscaExpedidor 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox TxtNomeExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   46
         Top             =   480
         Width           =   3195
      End
      Begin VB.TextBox TxtCGCExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   38
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label Label36 
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
         Left            =   757
         TabIndex        =   165
         Top             =   1860
         Width           =   2520
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Apólice:"
         Height          =   195
         Left            =   465
         TabIndex        =   164
         Top             =   3405
         Width           =   570
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seguradora:"
         Height          =   195
         Left            =   165
         TabIndex        =   163
         Top             =   3105
         Width           =   870
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   360
         TabIndex        =   162
         Top             =   2505
         Width           =   675
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   675
         TabIndex        =   161
         Top             =   2205
         Width           =   360
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FAX:"
         Height          =   195
         Left            =   690
         TabIndex        =   160
         Top             =   2805
         Width           =   345
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3900
         X2              =   120
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "End."
         Height          =   195
         Left            =   345
         TabIndex        =   118
         Top             =   1125
         Width           =   330
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "I.E."
         Height          =   195
         Left            =   2160
         TabIndex        =   116
         Top             =   825
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   270
         TabIndex        =   112
         Top             =   825
         Width           =   405
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   255
         TabIndex        =   48
         Top             =   525
         Width           =   420
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ, Apelido ou Nome"
         Height          =   195
         Left            =   360
         TabIndex        =   47
         Top             =   225
         Width           =   1710
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   1425
         Width           =   495
      End
   End
   Begin VB.Frame FraCiaAerea 
      Caption         =   "Cia. Aérea/Tab."
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
      Left            =   60
      TabIndex        =   57
      Top             =   780
      Width           =   3375
      Begin VB.CommandButton CmdDadosCia 
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
         Left            =   3000
         TabIndex        =   151
         Top             =   315
         Width           =   255
      End
      Begin VB.TextBox TxtBuscaSiglaCia 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   1
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox TxtSiglaCiaAerea 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   58
         Top             =   300
         Width           =   435
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3540
         X2              =   120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label TxtNomeCiaAerea 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   108
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label TxtInscrEstCiaAerea 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   107
         Top             =   1140
         Width           =   2655
      End
      Begin VB.Label TxtCGCCiaAerea 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   106
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   420
         TabIndex        =   68
         Top             =   885
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Inscr. Est."
         Height          =   195
         Left            =   180
         TabIndex        =   67
         Top             =   1185
         Width           =   705
      End
   End
   Begin VB.Frame FraBotoes 
      Height          =   8295
      Left            =   10800
      TabIndex        =   156
      Top             =   60
      Width           =   1095
      Begin VB.CommandButton CmdCancAWB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar AWB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton CmdLimpaTela 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Limpar Tela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5520
         Width           =   855
      End
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H008080FF&
         Caption         =   "Sair"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   6840
         Width           =   855
      End
      Begin VB.CommandButton CmdEmitir 
         BackColor       =   &H0080FF80&
         Caption         =   "Emitir AWB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdImportarDados 
         BackColor       =   &H0080C0FF&
         Caption         =   "Usar Dados de AWB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton CmdVia2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Buscar AWB para Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.Frame FraDestinatario 
      Caption         =   "Destinatário/Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   3480
      TabIndex        =   49
      Top             =   3180
      Width           =   4035
      Begin VB.TextBox TxtUFDestinatario 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   51
         Top             =   1380
         Width           =   435
      End
      Begin VB.TextBox TxtCidadeDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   52
         Top             =   1380
         Width           =   2775
      End
      Begin VB.CommandButton CmdDadosDestinatario 
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
         Left            =   3660
         TabIndex        =   148
         Top             =   180
         Width           =   255
      End
      Begin VB.TextBox TxtBairroDEST 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   141
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox TxtSeguradoraDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   132
         Top             =   3060
         Width           =   2715
      End
      Begin VB.TextBox TxtApoliceDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   131
         Top             =   3360
         Width           =   2715
      End
      Begin VB.TextBox TxtFAXDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   128
         Top             =   2760
         Width           =   2715
      End
      Begin VB.TextBox TxtTelDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   127
         Top             =   2460
         Width           =   2715
      End
      Begin VB.TextBox TxtCEPDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   126
         Top             =   2160
         Width           =   2715
      End
      Begin VB.TextBox TxtEndDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   124
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TxtInscrEstDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   122
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox TxtBuscaDestinatario 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox TxtCGCDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   53
         Top             =   780
         Width           =   1395
      End
      Begin VB.TextBox TxtNomeDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   50
         Top             =   480
         Width           =   3195
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
         Left            =   757
         TabIndex        =   171
         Top             =   1860
         Width           =   2520
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Apólice:"
         Height          =   195
         Left            =   540
         TabIndex        =   170
         Top             =   3420
         Width           =   570
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seguradora:"
         Height          =   195
         Left            =   240
         TabIndex        =   169
         Top             =   3120
         Width           =   870
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   435
         TabIndex        =   168
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   750
         TabIndex        =   167
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FAX:"
         Height          =   195
         Left            =   765
         TabIndex        =   166
         Top             =   2820
         Width           =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3900
         X2              =   120
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ, Apelido ou Nome"
         Height          =   195
         Left            =   420
         TabIndex        =   133
         Top             =   225
         Width           =   1710
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "End."
         Height          =   195
         Left            =   345
         TabIndex        =   125
         Top             =   1125
         Width           =   330
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "I.E."
         Height          =   195
         Left            =   2160
         TabIndex        =   123
         Top             =   825
         Width           =   240
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   180
         TabIndex        =   56
         Top             =   1425
         Width           =   495
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   270
         TabIndex        =   55
         Top             =   825
         Width           =   405
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   255
         TabIndex        =   54
         Top             =   525
         Width           =   420
      End
   End
   Begin VB.Frame FraAWB 
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
      Left            =   60
      TabIndex        =   157
      Top             =   1500
      Width           =   3375
      Begin VB.TextBox TxtAWB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TxtDig 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2820
         MaxLength       =   1
         TabIndex        =   3
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame FraSpot 
      Caption         =   "Tarifa Spot"
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   136
      Top             =   7080
      Width           =   3195
      Begin VB.CommandButton CmdCancSpot 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   1560
         TabIndex        =   35
         Top             =   960
         Width           =   1515
      End
      Begin VB.CommandButton CmdContinuar 
         Caption         =   "Continuar"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtAutorizador 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   255
         TabIndex        =   32
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox TxtKilo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2100
         TabIndex        =   33
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Autorizador"
         Height          =   195
         Left            =   120
         TabIndex        =   138
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Valor por Kilo"
         Height          =   195
         Left            =   1080
         TabIndex        =   137
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Frame FraTaxas 
      Caption         =   "Taxas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   7560
      TabIndex        =   78
      Top             =   1980
      Width           =   3195
      Begin VB.TextBox Txt_transp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         TabIndex        =   172
         Text            =   "0.00"
         Top             =   2944
         Width           =   1335
      End
      Begin VB.TextBox TxtKiloCob 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   153
         Text            =   "0.00"
         Top             =   1484
         Width           =   1335
      End
      Begin VB.TextBox TxtTipoADVAL 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   139
         Top             =   2100
         Width           =   315
      End
      Begin VB.TextBox TxtTipoTaxa 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   113
         Top             =   900
         Width           =   1335
      End
      Begin VB.CommandButton CmdTarifaSpot 
         Caption         =   "Tarifa Spot"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   540
         Width           =   2955
      End
      Begin VB.CommandButton CmdCalcularTaxas 
         Caption         =   "Calcular Tarifas"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2955
      End
      Begin VB.TextBox TxtOutros2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox TxtOutros1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   3540
         Width           =   1335
      End
      Begin VB.TextBox TxtICMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   101
         Text            =   "0.00"
         Top             =   4740
         Width           =   1335
      End
      Begin VB.TextBox TxtDescrOutros2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   30
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox TxtDescrOutros1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   28
         Top             =   3540
         Width           =   1575
      End
      Begin VB.TextBox TxtAliquota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   102
         Top             =   4470
         Width           =   1335
      End
      Begin VB.TextBox TxtTXRedesp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox TxtTXDestino 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   82
         Text            =   "0.00"
         Top             =   2652
         Width           =   1335
      End
      Begin VB.TextBox TxtTXOrigem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   99
         Text            =   "0.00"
         Top             =   2360
         Width           =   1335
      End
      Begin VB.TextBox TxtADValorem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   80
         Text            =   "0.00"
         Top             =   2100
         Width           =   1035
      End
      Begin VB.TextBox TxtFreteRegional 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   81
         Text            =   "0.00"
         Top             =   1776
         Width           =   1335
      End
      Begin VB.TextBox TxtFreteNacional 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   79
         Text            =   "0.00"
         Top             =   1192
         Width           =   1335
      End
      Begin VB.TextBox TxtFreteTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   89
         Text            =   "0.00"
         Top             =   4140
         Width           =   1335
      End
      Begin VB.Label Label43 
         Caption         =   "Tx. Transp....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   173
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Kilo..........:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   154
         Top             =   1545
         Width           =   1575
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Taxa.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   114
         Top             =   945
         Width           =   1575
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS..........:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   103
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tx. Redesp....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   87
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tx. Ter. Dest.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   86
         Top             =   2745
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tx. Ter. Orig.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   100
         Top             =   2445
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ad. Val.......:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   85
         Top             =   2145
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Frete Reg.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   84
         Top             =   1845
         Width           =   1575
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Frete Nac.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   83
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Alíquota......:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   104
         Top             =   4485
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Frete + Taxas"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   88
         Top             =   4185
         Width           =   1365
      End
   End
   Begin VB.Frame FraRetira 
      Caption         =   "Cliente Retira?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   7560
      TabIndex        =   75
      Top             =   840
      Width           =   3195
      Begin VB.TextBox TxtLocalRetirada 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   24
         Top             =   720
         Width           =   2955
      End
      Begin VB.OptionButton OptRetiraSim 
         Caption         =   "Sim"
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
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton OptRetiraNao 
         Caption         =   "Não"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Local de Retirada"
         Height          =   195
         Left            =   120
         TabIndex        =   155
         Top             =   480
         Width           =   1260
      End
   End
   Begin VB.Frame FraModalFrete 
      Caption         =   "Modalidade do Frete"
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
      Left            =   7560
      TabIndex        =   74
      Top             =   120
      Width           =   3195
      Begin VB.OptionButton OptAPagar 
         Caption         =   "A Pagar"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1860
         TabIndex        =   21
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton OptPago 
         Caption         =   "Pago"
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
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame FraVolumes 
      Caption         =   "Volumes e Peso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   60
      TabIndex        =   69
      Top             =   5640
      Width           =   3375
      Begin VB.TextBox TxtVolumes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   2400
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TxtPesoReal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   1260
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton CmdIncluirVolume 
         Caption         =   "Incluir Cubagens"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox TxtPesoCubado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   70
         Top             =   2160
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid FlexGridVolumes 
         Height          =   1275
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2249
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Volumes"
         Height          =   195
         Left            =   2400
         TabIndex        =   158
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Peso Real"
         Height          =   195
         Left            =   1260
         TabIndex        =   73
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Peso Cub."
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   1920
         Width           =   735
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
      Height          =   1275
      Left            =   3480
      TabIndex        =   59
      Top             =   1860
      Width           =   4035
      Begin VB.TextBox TxtAeroportoVIA 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1860
         TabIndex        =   60
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TxtSiglaVIA 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1380
         TabIndex        =   61
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox TxtAeroportoDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1860
         TabIndex        =   147
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox TxtSiglaDestinatario 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1380
         TabIndex        =   146
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox TxtAeroportoExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1860
         TabIndex        =   143
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscaSiglaDEST 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox TxtSiglaExpedidor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1380
         TabIndex        =   144
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtBuscaSiglaExp 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox TxtBuscaSiglaVIA 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Left            =   120
         TabIndex        =   152
         Top             =   585
         Width           =   540
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Origem"
         Height          =   195
         Left            =   165
         TabIndex        =   145
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "VIA"
         Height          =   195
         Left            =   405
         TabIndex        =   62
         Top             =   885
         Width           =   255
      End
   End
   Begin VB.Frame FraEspecie 
      Caption         =   "Espécie"
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
      Left            =   60
      TabIndex        =   76
      Top             =   2160
      Width           =   3375
      Begin VB.ComboBox ComboEspecie 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   3135
      End
   End
   Begin VB.Frame FraProduto 
      Caption         =   "Descrição do Produto (Tecle ENTER para Escolher)"
      Height          =   915
      Left            =   3480
      TabIndex        =   98
      Top             =   4980
      Width           =   4035
      Begin VB.TextBox TxtDescrIATA 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   3795
      End
      Begin VB.TextBox TxtDescrOutros 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         MaxLength       =   25
         TabIndex        =   17
         Top             =   1380
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox ChkPerecivel 
         Caption         =   "Produto Perecível"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3795
      End
      Begin VB.ComboBox ComboProduto 
         Height          =   315
         Left            =   3720
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   3795
      End
   End
   Begin VB.Frame FraOBS 
      Caption         =   "Observações de Emissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   3480
      TabIndex        =   105
      Top             =   5940
      Width           =   4035
      Begin VB.TextBox TxtOBSEmissao 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         MaxLength       =   239
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label LblAtualizarFrete 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sim"
         Height          =   195
         Left            =   600
         TabIndex        =   149
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame FraNFs 
      Caption         =   "Notas Fiscais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   63
      Top             =   3000
      Width           =   3375
      Begin VB.TextBox TxtTotalVM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2340
         TabIndex        =   65
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton CmdIncluirNF 
         Caption         =   "Incluir Notas Fiscais"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   1635
      End
      Begin MSFlexGridLib.MSFlexGrid FlexGridNFs 
         Height          =   1815
         Left            =   120
         TabIndex        =   64
         Top             =   660
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   1920
         TabIndex        =   66
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame FraFiliais 
      Caption         =   "Informa a Filial de Emissão"
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
      Left            =   60
      TabIndex        =   90
      Top             =   60
      Width           =   3375
      Begin VB.CommandButton CmdDadosFilial 
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
         Left            =   3000
         TabIndex        =   150
         Top             =   315
         Width           =   255
      End
      Begin VB.TextBox TxtBuscaFilial 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Top             =   300
         Width           =   435
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3540
         X2              =   120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label TxtLicensaFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   111
         Top             =   1740
         Width           =   1995
      End
      Begin VB.Label TxtCidadeFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   95
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Label TxtFilial 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   134
         Top             =   300
         Width           =   435
      End
      Begin VB.Label TxtNomeFilial 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   135
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label TxtSiglaFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2940
         TabIndex        =   109
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Lic. IATA"
         Height          =   195
         Left            =   180
         TabIndex        =   110
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label TxtUFFilial 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3060
         TabIndex        =   97
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   330
         TabIndex        =   96
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   420
         TabIndex        =   94
         Top             =   885
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Inscr. Est."
         Height          =   195
         Left            =   120
         TabIndex        =   93
         Top             =   1185
         Width           =   705
      End
      Begin VB.Label TxtCGCFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   92
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label TxtInscrEstFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   91
         Top             =   1140
         Width           =   2655
      End
   End
   Begin VB.Frame FraIATA 
      Caption         =   "Cod. Prod. IATA"
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
      Left            =   -2160
      TabIndex        =   77
      Top             =   5400
      Visible         =   0   'False
      Width           =   3675
      Begin VB.TextBox TxtCodIATA 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   4
         Top             =   540
         Visible         =   0   'False
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmEmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public COMBODIGESP As Integer

Public xIMPNomeEXP As String, xIMPCGCEXP As String, xIMPInscEstEXP As String, _
xIMPEndEXP As String, xIMPBairroEXP As String, xIMPCidadeEXP As String, _
xIMPCepEXP As String, xIMPUFEXP As String, xIMPTelEXP As String, _
xIMPFAXEXP As String, xIMPNomeDEST As String, xIMPCGCDEST As String, _
xIMPInscEstDEST As String, xIMPEndDEST As String, xIMPBairroDEST As String, _
xIMPCidadeDEST As String, xIMPCepDEST As String, xIMPUFDEST As String, _
xIMPTelDEST As String, xIMPFAXDEST As String, xIMPOrigem As String, _
xIMPVia As String, xIMPCidadeDESTINO As String, xIMPSIGLA As String, _
xIMPReqTranspMinuta As String, xIMPNumControle As String, xIMPInscrEstCiaAerea As String, _
xIMPCNPJCiaAerea As String, xIMPVlDecTRANSP As String, xIMPVlDecSUFRAMA As String, _
xIMPDescrEmbalagem As String, xIMPQteVol As String, xIMPPesoReal As String, _
xIMPPesoTax As String, xIMPTrecho As String, xIMPCl As String, xIMPCodigo As String, _
xIMPKilo As String, xIMPFreteNacEscopo As String, xIMPFreteRegEscopo As String, _
xIMPNatureza As String, xIMPTxDescrDevAg As String, xIMPTxDescrDevTransp As String, _
xIMPFreteNacional As String, xIMPFreteRegional As String, xIMPAdValorem As String, _
xIMPTipoADVAL As String, xIMPTxTerrOrig As String, xIMPTxTerrDest As String, _
xIMPTxRedesp As String, xIMPTxAgente As String, xIMPTxDevTransp As String, _
xIMPDescrTxOutros1 As String, xIMPTxOutros1 As String, xIMPDescrTxOutros2 As String, _
xIMPTxOutros2 As String, xIMPFreteTotal As String, xIMPStrObservacao As String, _
xIMPStrRetiraSIM As String, xIMPStrRetiraNAO As String, xIMPStrLocalRetira As String, _
xIMPHorarioAt As String, xIMPStrTelefone As String, xIMPStrTotalServ As String, _
xIMPStrBaseCalculo As String, xIMPStrAliquota As String, xIMPStrICMS As String, _
xIMPAgenteEmissor As String, xIMPCodIATA As String, xIMPDtEmissao As String
Public xIMPHoraEmissao As String, xIMPNaturezaOp As String, xIMPCFOP As String, _
xIMPEmissor As String, xIMPLocalidade As String, xIMPMatricula As String, _
XIMPAUX As String, StringNF As String, xIMPObsICMS As String, xIMPObsPerecivel As String, _
xIMPObsSeguro As String, xIMPStrNF01 As String, xIMPStrNF02 As String, xIMPStrNF03 As String, _
xIMPStrNF04 As String, xIMPStrNF05 As String, xIMPStrNF06 As String, _
xIMPStrNF07 As String, xIMPStrNF08 As String, xIMPStrNF09 As String, _
xIMPStrNF10 As String, xIMPStrNF11 As String, xIMPStrNF12 As String, _
xIMPStrObservacao01 As String, xIMPStrObservacao02 As String, xIMPStrObservacao03 As String, _
xIMPStrObservacao04 As String, xDim As String

Private Sub ChkPerecivel_Click()
    If ChkPerecivel.Value = 1 Then
    ChkPerecivel.FontBold = True
    Else
    ChkPerecivel.FontBold = False
    End If
End Sub

Private Sub CmdCancAWB_Click()
    frmEmissaoCANCAWB.Show 1
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdCancSpot_Click()
    FraAeoportos.Enabled = True
    FraCiaAerea.Enabled = True
    FraDestinatario.Enabled = True
    FraEspecie.Enabled = True
    FraExpedidor.Enabled = True
    FraFiliais.Enabled = True
    FraIATA.Enabled = True
    FraModalFrete.Enabled = True
    FraNFs.Enabled = True
    FraOBS.Enabled = True
    FraProduto.Enabled = True
    FraRetira.Enabled = True
    FraTaxas.Enabled = True
    FraVolumes.Enabled = True
    FraSpot.Enabled = False
    
    TxtAutorizador.Enabled = False
    TxtKilo.Enabled = False
    
    TxtAutorizador.BackColor = xBranco
    TxtKilo.BackColor = xBranco
    TxtAutorizador.Text = ""
    TxtKilo.Text = ""
    
    CmdCalcularTaxas.SetFocus
End Sub

Private Sub CmdContinuar_Click()
Dim xFrete, xPesoTx, xVlKilo As Currency

If Len(Trim(TxtKilo.Text)) = 0 Then
    MsgBox "Você não informou o Valor do Kilo para a Tarifa Spot!", vbExclamation, ""
    Exit Sub
ElseIf Len(Trim(TxtAutorizador.Text)) = 0 Then
    MsgBox "Você não informou o nome do Autorizador para esta Terifa Spot!", vbExclamation, ""
    Exit Sub
ElseIf CDbl(TxtKilo.Text) = 0 Then
    MsgBox "O Valor por Kilo não pode ser nulo!", vbExclamation, ""
    Exit Sub
End If

    If TxtPesoReal.Text > TxtPesoCubado.Text Then
    xPesoTx = TxtPesoReal.Text
    Else
    xPesoTx = TxtPesoCubado.Text
End If

xVlKilo = TxtKilo.Text
xFrete = xPesoTx * xVlKilo

TxtKiloCob.Text = Format(xVlKilo, "###,###,###,##0.00")

TxtTipoTaxa.Text = "Spot"

TxtFreteNacional.Text = xFrete


TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")
TxtFreteNacional.Text = Format(TxtFreteNacional.Text, "###,###,###,##0.00")

FraAeoportos.Enabled = True
FraCiaAerea.Enabled = True
FraDestinatario.Enabled = True
FraEspecie.Enabled = True
FraExpedidor.Enabled = True
FraFiliais.Enabled = True
FraIATA.Enabled = True
FraModalFrete.Enabled = True
FraNFs.Enabled = True
FraOBS.Enabled = True
FraProduto.Enabled = True
FraRetira.Enabled = True
FraTaxas.Enabled = True
FraVolumes.Enabled = True
FraSpot.Enabled = False
TxtAutorizador.Enabled = False
TxtKilo.Enabled = False
TxtAutorizador.BackColor = xBranco
TxtKilo.BackColor = xBranco
CmdCalcularTaxas.SetFocus
CmdTarifaSpot.Enabled = False
End Sub

Private Sub MascaraAWB_Click()
Open "\\REGIS\EPSON" For Output As #1
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXX    XXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXX    XXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXX    XXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(72) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXX    XXXXXXX" & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXX    XXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXX  XXXXXX   XXXXXXXXXXXXXXXXXXX   XXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "XXXX  XXXXXXX XXXXXXX  XXXXXX  X  XXXX XXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX  XXXXXX  XXXX  XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX  XXXXXXXXXXXX  XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX  XXXXXXXXXXXX  XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "                        XXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                        XXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                        XXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX   XXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXX"
Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXX   XXXXXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXXX"
DoEvents


DoEvents
Close #1
End Sub

Private Sub CmdDadosCia_Click()
Dim xFrame As Frame
Dim Botao As CommandButton
Dim HMax As Integer
Dim HMin As Integer
Set Botao = ActiveControl


Set xFrame = FraCiaAerea
HMax = 1515
HMin = 675

    If xFrame.Height = HMin Then
    xFrame.ZOrder (0)
    DoEvents
    Call TravaFrame(frmEmissao, xFrame, 0)
    xFrame.Height = HMax
    Botao.Caption = "<"
    DoEvents
    ElseIf xFrame.Height = HMax Then
    Call TravaFrame(frmEmissao, xFrame, 1)
    xFrame.Height = HMin
    Botao.Caption = ">"
    DoEvents
    End If
FraSpot.Enabled = False
End Sub

Private Sub CmdDadosDestinatario_Click()

Dim xFrame As Frame
Dim Botao As CommandButton
Dim HMax As Integer
Dim HMin As Integer
Set Botao = ActiveControl


Set xFrame = FraDestinatario
HMax = 3855
HMin = 1755

    If xFrame.Height = HMin Then
    xFrame.ZOrder (0)
    DoEvents
    Call TravaFrame(frmEmissao, xFrame, 0)
    xFrame.Height = HMax
    Botao.Caption = "<"
    DoEvents
    ElseIf xFrame.Height = HMax Then
    Call TravaFrame(frmEmissao, xFrame, 1)
    xFrame.Height = HMin
    Botao.Caption = ">"
    DoEvents
    End If
FraSpot.Enabled = False
End Sub

Private Sub CmdDadosFilial_Click()
Dim xFrame As Frame
Dim Botao As CommandButton
Dim HMax As Integer
Dim HMin As Integer
Set Botao = ActiveControl
Set xFrame = FraFiliais
HMax = 2115
HMin = 675

    If xFrame.Height = HMin Then
    xFrame.ZOrder (0)
    DoEvents
    Call TravaFrame(frmEmissao, xFrame, 0)
    xFrame.Height = HMax
    Botao.Caption = "<"
    DoEvents
    ElseIf xFrame.Height = HMax Then
    Call TravaFrame(frmEmissao, xFrame, 1)
    xFrame.Height = HMin
    Botao.Caption = ">"
    DoEvents
    End If
FraSpot.Enabled = False
End Sub

Private Sub CmdEmitir_Click()
 ' TOMADA DE DADOS
 
 
If Len(Trim(TxtLicensaFilial.Caption)) = 0 Then
MsgBox "ERRO! A Filial informada não está licensiada pela IATA a emitir AWBs. Entre em contato com o responsável pelo Aéreo ou com o Administrador do Sistema.", vbCritical, ""
Exit Sub
ElseIf Len(Trim(TxtAWB.Text)) = 0 Then
MsgBox "Você não informou o número do AWB a ser Emitido!", vbCritical, ""
Exit Sub
ElseIf Len(Trim(TxtDig.Text)) = 0 Then
MsgBox "Você não informou o dígito do AWB a ser Emitido!", vbCritical, ""
Exit Sub
ElseIf Len(Trim(ComboEspecie.Text)) = 0 Then
MsgBox "Você não informou a Espécie da Embalagem!", vbCritical, ""
Exit Sub
ElseIf Val(SemPonto(TxtFreteTotal.Text)) = 0 Then
MsgBox "ERRO! O Frete Total é nulo. Revise seus dados antes de continuar.", vbCritical, ""
Exit Sub
ElseIf Len(Trim(TxtDescrOutros1.Text)) = 0 And Val(SemPonto(TxtOutros1.Text)) > 0 Then
MsgBox "Você informou o Valor de uma taxa Adicional porém não descreveu que Taxa seria esta...", vbCritical, ""
Exit Sub
ElseIf Len(Trim(TxtDescrOutros2.Text)) = 0 And Val(SemPonto(TxtOutros2.Text)) > 0 Then
MsgBox "Você informou o Valor de uma taxa Adicional porém não descreveu que Taxa seria esta...", vbCritical, ""
Exit Sub
ElseIf Len(Trim(TxtDescrOutros1.Text)) > 0 And Val(SemPonto(TxtOutros1.Text)) = 0 Then
MsgBox "Você informou uma taxa Adicional porém não inseriu um valor para a mesma...", vbCritical, ""
Exit Sub
ElseIf Len(Trim(TxtDescrOutros2.Text)) > 0 And Val(SemPonto(TxtOutros2.Text)) = 0 Then
MsgBox "Você informou uma taxa Adicional porém não inseriu um valor para a mesma...", vbCritical, ""
Exit Sub
'ElseIf Len(Trim(ComboProduto.Text)) = 0 Then
'MsgBox "Você não especificou o tipo do seu produto...", vbCritical, ""
'Exit Sub
'ElseIf ComboProduto.Text = "Outros" And Len(Trim(TxtDescrOutros.Text)) = 0 Then
'MsgBox "Você selecionou a Categoria 'Outros' de Produtos porém não informou a descrição do mesmo...", vbCritical, ""
'Exit Sub
ElseIf Val(TxtVolumes.Text) = 0 Then
MsgBox "Você não informou os volumes que viajarão...", vbCritical, ""
Exit Sub
ElseIf ComboProduto.Text = "Outros" And Len(Trim(TxtDescrOutros.Text)) = 0 Then
MsgBox "Você selecionou a Categoria 'Outros' de Produtos porém não informou a descrição do mesmo...", vbCritical, ""
Exit Sub
ElseIf LblAtualizarFrete.Caption = "Sim" Then
MsgBox "Você pode ter alterado alguns dados desde a última vez que você calculou as tarifas. Antes de continuar, calcule novamente e revise as Taxas componentes do Frete.", vbCritical, ""
Exit Sub
End If

Me.MousePointer = 11
DoEvents

If Acao = "IMPRIMIR" Then
'CONFIGURACAO DE IMPRESSORAS - Inicio
Dim SETIMPLinha As String
Dim SETIMPLinhaPC As String
Dim SETIMPImpressoraAtual As Printer
Dim SETIMPAchouIMP As Boolean
Dim NomeMicro As String

    If Dir("c:\printer.cfg") = "" Then
    MsgBox "Você não possui o arquivo de configuração de impressoras. Antes de continuar, é imprescindível que você configure as configure.", vbExclamation, "IMPRESSORAS"
    frmControleImpressoras.Show 1
    End If

AchouIMP = False
    
    If Dir("c:\printer.cfg") <> "" Then
        Open "c:\printer.cfg" For Input As #1
        Do Until EOF(1)
            Line Input #1, SETIMPxLinha
            If Mid(SETIMPxLinha, 1, 3) = "AWB" Then
            SETIMPImpressoraPadrao = Mid(SETIMPxLinha, 5)
            SETIMPAchouIMP = True
            Exit Do
            End If
        Loop
        Close #1
    End If
    
    
    If SETIMPAchouIMP = False Then
    MsgBox "Não existe impressora configurada para esta operação. Corrija este problema indo ao menu Configurações e depois em Impressoras e configure em qual impressora os AWBs deverão ser impressos.", vbCritical, "ERRO!"
    Exit Sub
    End If
    
    For Each SETIMPImpressoraAtual In Printers
        If SETIMPImpressoraAtual.DeviceName = SETIMPImpressoraPadrao Then
            Set Printer = SETIMPImpressoraAtual
            DoEvents
            Exit For
        End If
    Next

    If Mid(SETIMPImpressoraPadrao, 1, 1) <> "\" Then
    SETIMPImpressoraPadrao = "LPT1"
    End If

'CONFIGURACAO DE IMPRESSORAS - Fim
End If

'If de_informa.rsVerificacaodeForm.State = 1 Then de_informa.rsVerificacaodeForm.Close
'de_informa.VerificacaodeForm TxtSiglaCiaAerea.Text, TxtFilial.Caption

'    If de_informa.rsVerificacaodeForm.RecordCount = 0 Then
'    MsgBox "ERRO! Não existe formulário disponível para a Emissão deste AWB.", vbCritical, ""
'    Exit Sub
'    End If
    

Dim xNotas As Integer
Dim xVolumes As Integer
Dim Y0 As Integer
Dim xTipoADVal As Integer

Dim xItemNF, xItemVol As Long
Dim xCodAwb As String

Dim xDescrTxOutros1, xDescrTxOutro2, xDescrProdSis, xDescrProdOutros, xPerecivel, _
xModal, xOBSEmissor, xOBSSis, xRetira, xLocalRetirada, xFilialCTC As String
Dim xValMerc, xPesoReal, xPesoCubado, xKilo, xSpotKilo, xFreteNacional, xFreteRegional, xADValorem, _
xPercADVal, xTxOrigem, xTxDestino, xTxRedesp, xTxAgente, xTxTransp, xTxOutros1, xTxOutros2, _
xFreteTotal, xAliquota, xICMS, xFreteTotalLiq As Currency
    If TxtKiloCob.Text = "" Then
    xKilo = 0
    Else
    xKilo = CDbl(TxtKiloCob.Text)
    End If
    
    If TxtKilo.Text = "" Then
    xSpotKilo = 0
    Else
    xSpotKilo = CDbl(TxtKilo.Text)
    End If
    
xFreteNacional = CDbl(TxtFreteNacional.Text)
xFreteRegional = CDbl(TxtFreteRegional.Text)
xADValorem = CDbl(TxtADValorem.Text)
xTipoADVal = Val(TxtTipoADVAL.Text)
    If xTipoADVal = 1 Then
    xPercADVal = 0.33
    ElseIf xTipoADVal = 2 Then
    xPercADVal = 0.66
    End If
xTxOrigem = CDbl(TxtTXOrigem.Text)
xTxDestino = CDbl(TxtTXDestino.Text)
xTxRedesp = CDbl(TxtTXRedesp.Text)
xTxAgente = 0


'***********************Alteração - Lincoln  ****************************
If Txt_transp.Text = "" Then
    xTxTransp = 0
Else
    xTxTransp = CDbl(Txt_transp.Text)
End If
'***********************Alteração - Lincoln  ****************************


    If CDbl(TxtOutros1.Text) > 0 Then
    xTxOutros1 = CDbl(TxtOutros1.Text)
    xDescrTxOutros1 = UCase(Trim(TxtDescrOutros1.Text))
    Else
    xTxOutros1 = 0
    xDescrTxOutros1 = ""
    End If
    
    If CDbl(TxtOutros2.Text) > 0 Then
    xTxOutros2 = CDbl(TxtOutros2.Text)
    xdescrtxoutros2 = UCase(Trim(TxtDescrOutros2.Text))
    Else
    xTxOutros2 = 0
    xdescrtxoutros2 = ""
    End If

    
xFreteTotal = CDbl(TxtFreteTotal.Text)
xICMS = CDbl(TxtICMS.Text)
xFreteTotalLiq = xFreteTotal - xICMS

    If TxtAliquota.Text <> "ISENTO" Then
    xAliquota = CDbl(TxtAliquota.Text)
    xICMS = CDbl(TxtICMS.Text)
    Else
    xAliquota = 0
    xICMS = 0
    End If
    
    'If Len(Trim(TxtAutorizador.Text)) > 0 Then
    'xKilo = CDbl(TxtKilo.Text)
    'Else
    'xKilo = 0
    'End If
    
xNotas = FlexGridNFs.Rows - 1
xValMerc = CDbl(TxtTotalVM.Text)
xPesoReal = CDbl(TxtPesoReal.Text)
If IsNumeric(TxtPesoCubado.Text) = True Then
xPesoCubado = CDbl(TxtPesoCubado.Text)
Else
xPesoCubado = 0
End If
'xVolumes = 0
'    If FlexGridVolumes.Rows > 1 Then
'        For Y0 = 1 To FlexGridVolumes.Rows - 1
'        xVolumes = xVolumes + Val(FlexGridVolumes.TextMatrix(Y0, 0))
'        Next
'    End If
xVolumes = Val(TxtVolumes.Text)

xDescrProdSis = UCase(Trim(TxtDescrIATA.Text))
xDescrProdOutros = ""

xPerecivel = ""
If ChkPerecivel.Value = 1 Then xPerecivel = "S"

    If OptPago.Value = True Then
    xModal = "PAGO"
    ElseIf OptAPagar.Value = True Then
    xModal = "A PAGAR"
    End If
    
    If OptRetiraSim.Value = True Then
    xRetira = "S"
    ElseIf OptRetiraNao.Value = True Then
    xRetira = "N"
    End If
xLocalRetirada = UCase(Trim(TxtLocalRetirada.Text))
xOBSEmissor = UCase(Trim(TxtOBSEmissao.Text))

xOBSSis = ""

If xModal = "PAGO" Then
    If xAliquota = 4 Then
    xOBSSis = xOBSSis & "ICMS - ALIQUOTA DE 4% - RESOLUCAO 95/96 SENADO FEDERAL "
    End If
    
    If Len(Trim(TxtApoliceExpedidor.Text)) > 0 Then
    xOBSSis = xOBSSis & "Seguradora: " & UCase(Trim(TxtSeguradoraExpedidor.Text)) & "/" & Trim(TxtApoliceExpedidor.Text) & " "
    End If
    
    If xPerecivel = "S" Then
    xOBSSis = xOBSSis & "P E R E C I V E L - Prazo de Duracao: 48 hs"
    End If
ElseIf xModal = "A PAGAR" Then
    If xAliquota > 0 Then
    xOBSSis = xOBSSis & "ICMS - ALIQUOTA DE 4% - RESOLUCAO 95/96 SENADO FEDERAL   "
    End If
    
    If Len(Trim(TxtApoliceDestinatario.Text)) > 0 Then
    xOBSSis = xOBSSis & "Seguradora: " & UCase(Trim(TxtSeguradoraDestinatario.Text)) & "/" & Trim(TxtApoliceDestinatario.Text) & "   "
    End If
    
    If xPerecivel = "S" Then
    xOBSSis = xOBSSis & "P E R E C I V E L - Prazo de Duracao: 48 hs"
    End If
End If

    If UCase(Trim(TxtSiglaCiaAerea.Text)) = "VP" Then
    xOBSSis = xOBSSis & "****  A N C  50097  ****"
    End If
    

If de_informa.rsSelAWBCodItemNOTA.State = 1 Then de_informa.rsSelAWBCodItemNOTA.Close
de_informa.SelAWBCodItemNOTA
    If de_informa.rsSelAWBCodItemNOTA.RecordCount = 0 Or IsNull(de_informa.rsSelAWBCodItemNOTA.Fields("coditem")) = True Then
    xItemNF = 1
    Else
    xItemNF = de_informa.rsSelAWBCodItemNOTA.Fields("coditem") + 1
    End If

If de_informa.rsSelAWBCodItemVOL.State = 1 Then de_informa.rsSelAWBCodItemVOL.Close
de_informa.SelAWBCodItemVOL
    If de_informa.rsSelAWBCodItemVOL.RecordCount = 0 Or IsNull(de_informa.rsSelAWBCodItemVOL.Fields("coditem")) = True Then
    xItemVol = 1
    Else
    xItemVol = de_informa.rsSelAWBCodItemVOL.Fields("coditem") + 1
    End If

'Aqui Comeca o Transction
'de_informa.cn_informa.BeginTrans

If de_informa.rsConfereNumeroAWB.State = 1 Then de_informa.rsConfereNumeroAWB.Close
de_informa.ConfereNumeroAWB TxtSiglaCiaAerea.Text, TxtFilial.Caption, TxtAWB.Text

    If de_informa.rsConfereNumeroAWB.RecordCount = 0 Then
    MsgBox "Este formulário não está cadastrado.", vbCritical, ""
    Me.MousePointer = 0
    Exit Sub
    ElseIf de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") = "C" Then
    MsgBox "O formulário para este AWB está cancelado. Para utilizá-lo, vá até o cadastro de formulários e descancele-o.", vbCritical, ""
    Me.MousePointer = 0
    Exit Sub
    ElseIf de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") = "E" And Acao = "GRAVAR" Then
    MsgBox "Este AWB já foi emitido. Seus dados podem ser alterados se você escolher Alterar AWB.", vbCritical, ""
    Me.MousePointer = 0
    Exit Sub
    ElseIf Acao = "IMPRIMIR" And de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") <> "E" And de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") <> "I" Then
    MsgBox "Este AWB não está emitido. Não é possível imprimí-lo.", vbCritical, ""
    Me.MousePointer = 0
    Exit Sub
    ElseIf de_informa.rsConfereNumeroAWB.Fields("dig") <> TxtDig.Text Then
    MsgBox "O dígito para este AWB não confere.", vbCritical, ""
    Me.MousePointer = 0
    Exit Sub
    End If
    
'xAWB = de_informa.rsCapturaNumeroAWB.Fields("numero")
'xDig = de_informa.rsCapturaNumeroAWB.Fields("dig")
'xCodAwb = String(2 - Len(Trim(Str(Val(TxtFilial.Caption)))), "0") & Trim(Str(Val(TxtFilial.Caption))) & UCase(Trim(TxtSiglaCiaAerea.Text)) & String(10 - Len(Trim(Str(Val(xAWB)))), "0") & Trim(Str(Val(xAWB))) & Trim(Str(Val(xDig)))

xAWB = Trim(TxtAWB.Text)
xDig = Trim(TxtDig.Text)
xCodAwb = String(2 - Len(Trim(Str(Val(TxtFilial.Caption)))), "0") & Trim(Str(Val(TxtFilial.Caption))) & UCase(Trim(TxtSiglaCiaAerea.Text)) & String(10 - Len(Trim(Str(Val(xAWB)))), "0") & Trim(Str(Val(xAWB))) & Trim(Str(Val(xDig)))
If de_informa.rsConsultaAWB.State = 1 Then de_informa.rsConsultaAWB.Close
de_informa.ConsultaAWB xCodAwb

    If de_informa.rsConsultaAWB.RecordCount > 0 And Acao = "GRAVAR" Then
    MsgBox "ERRO! Este número de AWB para esta Cia. Aérea consta como já emitido. Por favor, tente novamente.", vbCritical, ""
    Me.MousePointer = 0
    Exit Sub
    End If

'Primeira Parte do Insert do AWB
Dim xSiglaOrigem As String
Dim xCidadeOrigem As String
Dim xUFOrigem As String
Dim xAeroportoOrigem As String
Dim xSiglaVIA As String
Dim xCidadeVIA As String
Dim xUFVIA As String
Dim xAeroportoVIA As String
Dim xSiglaDestino As String
Dim xCidadeDestino As String
Dim xUFDestino As String
Dim xAeroportoDestino As String

xSiglaOrigem = UCase(TxtSiglaExpedidor.Text)
xCidadeOrigem = UCase(Trim(Mid(TxtAeroportoExpedidor.Text, 1, InStr(1, TxtAeroportoExpedidor.Text, "-", vbTextCompare) - 1)))
xUFOrigem = UCase(Trim(Mid(TxtAeroportoExpedidor.Text, InStr(1, TxtAeroportoExpedidor.Text, "-", vbTextCompare) + 1, 3)))
xAeroportoOrigem = UCase(Trim(Mid(TxtAeroportoExpedidor.Text, InStr(1, TxtAeroportoExpedidor.Text, "(", vbTextCompare) + 1, Len(TxtAeroportoExpedidor.Text) - (InStr(1, TxtAeroportoExpedidor.Text, "(", vbTextCompare) + 1))))

xSiglaVIA = UCase(TxtSiglaVIA.Text)
xCidadeVIA = UCase(Trim(Mid(TxtAeroportoVIA.Text, 1, InStr(1, TxtAeroportoVIA.Text, "-", vbTextCompare) - 1)))
xUFVIA = UCase(Trim(Mid(TxtAeroportoVIA.Text, InStr(1, TxtAeroportoVIA.Text, "-", vbTextCompare) + 1, 3)))
xAeroportoVIA = UCase(Trim(Mid(TxtAeroportoVIA.Text, InStr(1, TxtAeroportoVIA.Text, "(", vbTextCompare) + 1, Len(TxtAeroportoVIA.Text) - (InStr(1, TxtAeroportoVIA.Text, "(", vbTextCompare) + 1))))

xSiglaDestino = UCase(TxtSiglaDestinatario.Text)
xCidadeDestino = UCase(Trim(Mid(TxtAeroportoDestinatario.Text, 1, InStr(1, TxtAeroportoDestinatario.Text, "-", vbTextCompare) - 1)))
xUFDestino = UCase(Trim(Mid(TxtAeroportoDestinatario.Text, InStr(1, TxtAeroportoDestinatario.Text, "-", vbTextCompare) + 1, 3)))
xAeroportoDestino = UCase(Trim(Mid(TxtAeroportoDestinatario.Text, InStr(1, TxtAeroportoDestinatario.Text, "(", vbTextCompare) + 1, Len(TxtAeroportoDestinatario.Text) - (InStr(1, TxtAeroportoDestinatario.Text, "(", vbTextCompare) + 1))))

If Acao = "GRAVAR" Then
de_informa.cn_informa.BeginTrans
de_informa.InsAWB xCodAwb, Trim(Str(xAWB)), Trim(Str(xDig)), _
UCase(Trim(TxtSiglaCiaAerea.Text)), UCase(Trim(TxtNomeCiaAerea.Caption)), Trim(TxtCGCCiaAerea.Caption), Trim(TxtInscrEstCiaAerea.Caption), _
Trim(TxtFilial.Caption), Trim(TxtLicensaFilial.Caption), Mid(Trim(TxtDescrIATA.Text), 1, 3), _
UCase(Trim(ComboEspecie.Text)), xNotas, xValMerc, xVolumes, "0", "0", "0", xPesoReal, xPesoCubado, _
Trim(TxtCGCExpedidor.Text), Trim(TxtInscrEstExpedidor.Text), UCase(Trim(TxtNomeExpedidor.Text)), UCase(Trim(TxtEndExpedidor.Text)), UCase(Trim(TxtCEPExpedidor.Text)), UCase(Trim(TxtBairroEXP.Text)), UCase(Trim(TxtCidadeExpedidor.Text)), UCase(Trim(TxtUFExpedidor.Text)), UCase(Trim(TxtTelExpedidor.Text)), UCase(Trim(TxtFAXExpedidor.Text)), UCase(Trim(TxtSeguradoraExpedidor.Text)), Trim(TxtApoliceExpedidor.Text), _
Trim(TxtCGCDestinatario.Text), Trim(TxtInscrEstDestinatario.Text), UCase(Trim(TxtNomeDestinatario.Text)), UCase(Trim(TxtEndDestinatario.Text)), UCase(Trim(TxtCEPDestinatario.Text)), UCase(Trim(TxtBairroEXP.Text)), UCase(Trim(TxtCidadeDestinatario.Text)), UCase(Trim(TxtUFDestinatario.Text)), UCase(Trim(TxtTelDestinatario.Text)), UCase(Trim(TxtFAXDestinatario.Text)), UCase(Trim(TxtSeguradoraDestinatario.Text)), Trim(TxtApoliceDestinatario.Text), _
xSiglaOrigem, xCidadeOrigem, xUFOrigem, xAeroportoOrigem, _
xSiglaVIA, xCidadeVIA, xUFVIA, xAeroportoVIA, _
xSiglaDestino, xCidadeDestino, xUFDestino, xAeroportoDestino
'Segunda Parte do Insert do AWB
de_informa.InsAWB2 xCodAwb, UCase(Trim(TxtTipoTaxa.Text)), xKilo, xFreteNacional, xFreteRegional, xADValorem, _
xTipoADVal, xPercADVal, xTxOrigem, xTxDestino, xTxRedesp, xTxAgente, xTxTransp, xTxOutros1, _
xDescrTxOutros1, xTxOutros2, xdescrtxoutros2, xFreteTotal, xAliquota, xICMS, xFreteTotalLiq, _
UCase(Trim(TxtAutorizador.Text)), xSpotKilo, xDescrProdSis, xDescrProdOutros, xPerecivel, xModal, _
xLocalRetirada, xRetira, xOBSEmissor, xOBSSis, CDate(DataHora("Data")), CVar(DataHora("Hora")), xUsuario, UCase(Trim(TxtSiglaFilial.Caption))


    For Y0 = 1 To FlexGridNFs.Rows - 1
        If Len(Trim(FlexGridNFs.TextMatrix(Y0, 1))) = 1 Then
        xFilialCTC = "0" & Trim(FlexGridNFs.TextMatrix(Y0, 1))
        Else
        xFilialCTC = Trim(FlexGridNFs.TextMatrix(Y0, 1))
        End If
        
        If Len(Trim(FlexGridNFs.TextMatrix(Y0, 2))) > 0 Then
        xFilialCTC = xFilialCTC & String(8 - Len(Trim(FlexGridNFs.TextMatrix(Y0, 2))), "0") & Trim(FlexGridNFs.TextMatrix(Y0, 2))
        Else
        xFilialCTC = "0"
        End If
        
    If Trim(FlexGridNFs.TextMatrix(Y0, 5)) = "" Then FlexGridNFs.TextMatrix(Y0, 5) = "0"
        
        If Len(Trim(FlexGridNFs.TextMatrix(Y0, 0))) > 0 Then
        de_informa.InsAWBNota xItemNF, _
                                xCodAwb, _
                                xFilialCTC, _
                                Trim(FlexGridNFs.TextMatrix(Y0, 0)), _
                                Trim(FlexGridNFs.TextMatrix(Y0, 3)), _
                                Trim(FlexGridNFs.TextMatrix(Y0, 4)), _
                                CDbl(Trim(FlexGridNFs.TextMatrix(Y0, 5)))
        End If
    xItemNF = xItemNF + 1
    Next
    
    
    For Y0 = 1 To FlexGridVolumes.Rows - 1
    If Len(Trim(FlexGridVolumes.TextMatrix(Y0, 4))) = "" Then FlexGridVolumes.TextMatrix(Y0, 4) = "0"
    
     de_informa.InsAWBVol xItemVol, xCodAwb, _
                            Val(FlexGridVolumes.TextMatrix(Y0, 0)), _
                            Val(FlexGridVolumes.TextMatrix(Y0, 1)), _
                            Val(FlexGridVolumes.TextMatrix(Y0, 2)), _
                            Val(FlexGridVolumes.TextMatrix(Y0, 3)), _
                            Val(FlexGridVolumes.TextMatrix(Y0, 4)), _
                            (((Val(FlexGridVolumes.TextMatrix(Y0, 1)) * Val(FlexGridVolumes.TextMatrix(Y0, 2)) * Val(FlexGridVolumes.TextMatrix(Y0, 3))) / 6000) * Val(FlexGridVolumes.TextMatrix(Y0, 0)))
    xItemVol = xItemVol + 1
    Next
    
de_informa.AltStatusFormularioItem "E", CDate(DataHora("DATA")), Trim(TxtAWB.Text), Trim(TxtDig.Text), UCase(Trim(TxtSiglaCiaAerea.Text)), Trim(TxtFilial.Caption)

de_informa.cn_informa.CommitTrans
ElseIf Acao = "ALTERAR" Then
de_informa.cn_informa.BeginTrans
de_informa.ALT_AIRAWB xCodAwb, Trim(Str(xAWB)), Trim(Str(xDig)), _
UCase(Trim(TxtSiglaCiaAerea.Text)), UCase(Trim(TxtNomeCiaAerea.Caption)), Trim(TxtCGCCiaAerea.Caption), Trim(TxtInscrEstCiaAerea.Caption), _
Trim(TxtFilial.Caption), Trim(TxtLicensaFilial.Caption), Mid(Trim(TxtDescrIATA.Text), 1, 3), _
UCase(Trim(ComboEspecie.Text)), xNotas, xValMerc, xVolumes, "0", "0", "0", xPesoReal, xPesoCubado, _
Trim(TxtCGCExpedidor.Text), Trim(TxtInscrEstExpedidor.Text), UCase(Trim(TxtNomeExpedidor.Text)), UCase(Trim(TxtEndExpedidor.Text)), UCase(Trim(TxtCEPExpedidor.Text)), UCase(Trim(TxtBairroEXP.Text)), UCase(Trim(TxtCidadeExpedidor.Text)), UCase(Trim(TxtUFExpedidor.Text)), UCase(Trim(TxtTelExpedidor.Text)), UCase(Trim(TxtFAXExpedidor.Text)), UCase(Trim(TxtSeguradoraExpedidor.Text)), Trim(TxtApoliceExpedidor.Text), _
Trim(TxtCGCDestinatario.Text), Trim(TxtInscrEstDestinatario.Text), UCase(Trim(TxtNomeDestinatario.Text)), UCase(Trim(TxtEndDestinatario.Text)), UCase(Trim(TxtCEPDestinatario.Text)), UCase(Trim(TxtBairroEXP.Text)), UCase(Trim(TxtCidadeDestinatario.Text)), UCase(Trim(TxtUFDestinatario.Text)), UCase(Trim(TxtTelDestinatario.Text)), UCase(Trim(TxtFAXDestinatario.Text)), UCase(Trim(TxtSeguradoraDestinatario.Text)), Trim(TxtApoliceDestinatario.Text), _
xSiglaOrigem, xCidadeOrigem, xUFOrigem, xAeroportoOrigem, _
xSiglaVIA, xCidadeVIA, xUFVIA, xAeroportoVIA, _
xSiglaDestino, xCidadeDestino, xUFDestino, xAeroportoDestino
'Segunda Parte do Insert do AWB
de_informa.ALT_AIRAWB2 xCodAwb, UCase(Trim(TxtTipoTaxa.Text)), xKilo, xFreteNacional, xFreteRegional, xADValorem, _
xTipoADVal, xPercADVal, xTxOrigem, xTxDestino, xTxRedesp, xTxAgente, xTxTransp, xTxOutros1, _
xDescrTxOutros1, xTxOutros2, xdescrtxoutros2, xFreteTotal, xAliquota, xICMS, xFreteTotalLiq, _
UCase(Trim(TxtAutorizador.Text)), xKilo, xDescrProdSis, xDescrProdOutros, xPerecivel, xModal, _
xLocalRetirada, xRetira, xOBSEmissor, xOBSSis, CDate(DataHora("Data")), CVar(DataHora("Hora")), xUsuario, UCase(Trim(TxtSiglaFilial.Caption))

de_informa.DELETEAWBNota xCodAwb
de_informa.DELETEAWBVOL xCodAwb


    For Y0 = 1 To FlexGridNFs.Rows - 1
        If Len(Trim(FlexGridNFs.TextMatrix(Y0, 1))) = 1 Then
        xFilialCTC = "0" & Trim(FlexGridNFs.TextMatrix(Y0, 1))
        Else
        xFilialCTC = Trim(FlexGridNFs.TextMatrix(Y0, 1))
        End If
        
        If Len(Trim(FlexGridNFs.TextMatrix(Y0, 2))) > 0 Then
        xFilialCTC = xFilialCTC & String(8 - Len(Trim(FlexGridNFs.TextMatrix(Y0, 2))), "0") & Trim(FlexGridNFs.TextMatrix(Y0, 2))
        Else
        xFilialCTC = "0"
        End If
        
    If Trim(FlexGridNFs.TextMatrix(Y0, 5)) = "" Then FlexGridNFs.TextMatrix(Y0, 5) = "0"
        
        If Len(Trim(FlexGridNFs.TextMatrix(Y0, 0))) > 0 Then
        de_informa.InsAWBNota xItemNF, _
                                xCodAwb, _
                                xFilialCTC, _
                                Trim(FlexGridNFs.TextMatrix(Y0, 0)), _
                                Trim(FlexGridNFs.TextMatrix(Y0, 3)), _
                                Trim(FlexGridNFs.TextMatrix(Y0, 4)), _
                                CDbl(Trim(FlexGridNFs.TextMatrix(Y0, 5)))
        End If
    xItemNF = xItemNF + 1
    Next
    
    
    For Y0 = 1 To FlexGridVolumes.Rows - 1
    If Len(Trim(FlexGridVolumes.TextMatrix(Y0, 4))) = "" Then FlexGridVolumes.TextMatrix(Y0, 4) = "0"
    
    de_informa.InsAWBVol xItemVol, xCodAwb, _
                            Val(FlexGridVolumes.TextMatrix(Y0, 0)), _
                            Val(FlexGridVolumes.TextMatrix(Y0, 1)), _
                            Val(FlexGridVolumes.TextMatrix(Y0, 2)), _
                            Val(FlexGridVolumes.TextMatrix(Y0, 3)), _
                            CDbl(FlexGridVolumes.TextMatrix(Y0, 4)), _
                            (((Val(FlexGridVolumes.TextMatrix(Y0, 1)) * Val(FlexGridVolumes.TextMatrix(Y0, 2)) * Val(FlexGridVolumes.TextMatrix(Y0, 3))) / 6000) * Val(FlexGridVolumes.TextMatrix(Y0, 0)))
    xItemVol = xItemVol + 1
    Next
    
''de_informa.AltStatusFormularioItem xAWB, _
                                    UCase(Trim(TxtSiglaCiaAerea.Text)), _
                                    Trim(TxtFilial.Caption), _
                                    "E", _
                                    CDate(DataHora("data"))
    

de_informa.cn_informa.CommitTrans
End If

'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO
'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO
'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO

'Declaracao Extritamente de Variaveis para Tomada de Dados

If Acao = "IMPRIMIR" Then

de_informa.cn_informa.BeginTrans

StringNF = ""

    For Y0 = 1 To FlexGridNFs.Rows - 1
        If Mid(Trim(FlexGridNFs.TextMatrix(Y0, 0)), 1, 1) = "N" Then
            'If Len(Trim(FlexGridNFs.TextMatrix(Y0, 4))) > 0 Then
            'StringNF = StringNF & FlexGridNFs.TextMatrix(Y0, 3) & "-" & FlexGridNFs.TextMatrix(Y0, 4)
            'Else
            StringNF = StringNF & FlexGridNFs.TextMatrix(Y0, 3)
            'End If
        StringNF = StringNF & "/"
        Else
        StringNF = StringNF & "DECLARACAO/"
        End If
    Next
StringNF = StringNF & String(720 - Len(Trim(StringNF)), " ")
xIMPStrNF01 = ""
xIMPStrNF02 = Trim(Mid(StringNF, 1, 60))
xIMPStrNF03 = Trim(Mid(StringNF, 61, 60))
xIMPStrNF04 = Trim(Mid(StringNF, 121, 60))
xIMPStrNF05 = Trim(Mid(StringNF, 181, 60))
xIMPStrNF06 = Trim(Mid(StringNF, 241, 60))
xIMPStrNF07 = Trim(Mid(StringNF, 301, 60))
xIMPStrNF08 = Trim(Mid(StringNF, 361, 60))
xIMPStrNF09 = Trim(Mid(StringNF, 421, 60))
xIMPStrNF10 = Trim(Mid(StringNF, 481, 60))
xIMPStrNF11 = Trim(Mid(StringNF, 541, 60))
xIMPStrNF12 = Trim(Mid(StringNF, 601, 60))
'xIMPStrNF12 = Trim(Mid(StringNF, 661, 60))

xIMPNomeEXP = UCase(Trim(TxtNomeExpedidor.Text))
xIMPCGCEXP = Trim(TxtCGCExpedidor.Text)
xIMPInscEstEXP = Trim(TxtInscrEstExpedidor.Text)
xIMPEndEXP = UCase(Trim(TxtEndExpedidor.Text))
xIMPBairroEXP = UCase(Trim(TxtBairroEXP.Text))
xIMPCidadeEXP = UCase(Trim(TxtCidadeExpedidor.Text))
xIMPCepEXP = Trim(TxtCEPExpedidor.Text)
xIMPUFEXP = UCase(Trim(TxtUFExpedidor.Text))
xIMPTelEXP = Trim(TxtTelExpedidor.Text)
xIMPFAXEXP = Trim(TxtFAXExpedidor.Text)
xIMPNomeDEST = UCase(Trim(TxtNomeDestinatario.Text))
xIMPCGCDEST = Trim(TxtCGCDestinatario.Text)
xIMPInscEstDEST = Trim(TxtInscrEstDestinatario.Text)
xIMPEndDEST = UCase(Trim(TxtEndDestinatario.Text))
xIMPBairroDEST = UCase(Trim(TxtBairroDEST.Text))
xIMPCidadeDEST = UCase(Trim(TxtCidadeDestinatario.Text))
xIMPCepDEST = Trim(TxtCEPDestinatario.Text)
xIMPUFDEST = UCase(Trim(TxtUFDestinatario.Text))
xIMPTelDEST = Trim(TxtTelDestinatario.Text)
xIMPFAXDEST = Trim(TxtFAXDestinatario.Text)
xIMPOrigem = UCase(Trim(TxtSiglaExpedidor.Text))
xIMPVia = IIf((UCase(Trim(TxtSiglaDestinatario.Text)) = UCase(Trim(TxtSiglaVIA.Text))), "", UCase(Trim(TxtSiglaVIA.Text)))

If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
de_informa.SelAeroportoSigla UCase(Trim(TxtSiglaVIA.Text))

xIMPCidadeDESTINO = UCase(Trim(de_informa.rsSelAeroportoSigla.Fields("localidade")))
'xIMPCidadeDESTINO = UCase(Trim(TxtCidadeDestinatario.Text))
xIMPSIGLA = UCase(Trim(TxtSiglaVIA.Text))

xIMPReqTranspMinuta = xAWB & xDig
xIMPNumControle = xAWB & xDig
xIMPInscrEstCiaAerea = Trim(TxtInscrEstCiaAerea.Caption)
xIMPCNPJCiaAerea = Trim(TxtCGCCiaAerea.Caption)
    If OptAPagar.Value = True And Len(Trim(TxtApoliceDestinatario.Text)) = 0 Then
    xIMPStrNF12 = "F R E T E    A   P A G A R"
    xIMPVlDecTRANSP = Trim(TxtTotalVM.Text)
    xIMPVlDecSUFRAMA = Trim(TxtTotalVM.Text)
    ElseIf OptAPagar.Value = True And Len(Trim(TxtApoliceDestinatario.Text)) > 0 Then
    xIMPStrNF12 = "F R E T E    A   P A G A R"
    xIMPVlDecTRANSP = UCase(Trim(TxtSeguradoraDestinatario.Text)) & "/" & Trim(TxtApoliceDestinatario.Text)
    xIMPVlDecSUFRAMA = UCase(Trim(TxtSeguradoraDestinatario.Text)) & "/" & Trim(TxtApoliceDestinatario.Text)
    ElseIf OptPago.Value = True And Len(Trim(TxtApoliceExpedidor.Text)) = 0 Then
    xIMPStrNF12 = "F R E T E    P A G O"
    xIMPVlDecTRANSP = Trim(TxtTotalVM.Text)
    xIMPVlDecSUFRAMA = Trim(TxtTotalVM.Text)
    ElseIf OptPago.Value = True And Len(Trim(TxtApoliceExpedidor.Text)) > 0 Then
    xIMPStrNF12 = "F R E T E    P A G O"
    xIMPVlDecTRANSP = UCase(Trim(TxtSeguradoraExpedidor.Text)) & "/" & Trim(TxtApoliceExpedidor.Text)
    xIMPVlDecSUFRAMA = UCase(Trim(TxtSeguradoraExpedidor.Text)) & "/" & Trim(TxtApoliceExpedidor.Text)
    End If
    
xIMPDescrEmbalagem = UCase(Trim(ComboEspecie.Text))
xDim = " - "
    For Y0 = 1 To FlexGridVolumes.Rows - 1
        If Val(FlexGridVolumes.TextMatrix(Y0, 0)) > 0 Then
        xDim = xDim & Val(FlexGridVolumes.TextMatrix(Y0, 0)) & "x" & "(" & Val(FlexGridVolumes.TextMatrix(Y0, 1)) & "x" & Val(FlexGridVolumes.TextMatrix(Y0, 2)) & "x" & Val(FlexGridVolumes.TextMatrix(Y0, 3)) & ")/"
        'xDim = xDim & "(" & Val(FlexGridVolumes.TextMatrix(Y0, 1)) & "x" & Val(FlexGridVolumes.TextMatrix(Y0, 2)) & "x" & Val(FlexGridVolumes.TextMatrix(Y0, 3)) & ")/"
        End If
    Next
xIMPDescrEmbalagem = xIMPDescrEmbalagem & xDim
xIMPQteVol = Trim(Str(xVolumes))
xIMPPesoReal = Trim(TxtPesoReal.Text)
xIMPPesoTax = IIf(CDbl(TxtPesoReal.Text) > CDbl(TxtPesoCubado.Text), TxtPesoReal.Text, TxtPesoCubado.Text)
    If Len(Trim(TxtSiglaVIA.Text)) = 0 Then
    xIMPTrecho = UCase(Trim(TxtSiglaExpedidor.Text)) & "/" & UCase(Trim(TxtSiglaDestinatario.Text))
    Else
    xIMPTrecho = UCase(Trim(TxtSiglaExpedidor.Text)) & "/" & UCase(Trim(TxtSiglaVIA.Text))
    End If
xIMPCl = UCase(Mid(Trim(TxtTipoTaxa.Text), 1, 1))
xIMPCodigo = IIf(xIMPCl = "E", Mid(Trim(TxtDescrIATA.Text), 1, 3), "")
xIMPKilo = IIf(xIMPCl = "M", "T/M", xKilo)
xIMPFreteNacEscopo = TxtFreteNacional.Text
xIMPNatureza = IIf(ChkPerecivel.Value = 1, UCase(Mid(Trim(TxtDescrIATA.Text), 7)) & " PERECIVEL", UCase(Mid(Trim(TxtDescrIATA.Text), 7)))
xIMPTxDescrDevAg = ""
xIMPTxDescrDevTransp = ""
xIMPFreteNacional = TxtFreteNacional.Text
xIMPFreteRegional = TxtFreteRegional.Text
xIMPAdValorem = TxtADValorem.Text
xIMPTipoADVAL = TxtTipoADVAL.Text
xIMPTxTerrOrig = TxtTXOrigem.Text
xIMPTxTerrDest = TxtTXDestino.Text
xIMPTxRedesp = TxtTXRedesp.Text
xIMPTxAgente = "0,00"


'xIMPTxDevTransp = "0,00"
xIMPTxDevTransp = Txt_transp.Text


xIMPDescrTxOutros1 = UCase(Trim(TxtDescrOutros1.Text))
xIMPTxOutros1 = TxtOutros1.Text
xIMPDescrTxOutros2 = UCase(Trim(TxtDescrOutros2.Text))
xIMPTxOutros2 = TxtOutros2.Text
xIMPFreteTotal = TxtFreteTotal.Text



xIMPFreteNacional = String(10 - Len(Trim(xIMPFreteNacional)), " ") & Trim(xIMPFreteNacional)
xIMPFreteRegional = String(10 - Len(Trim(xIMPFreteRegional)), " ") & Trim(xIMPFreteRegional)
xIMPAdValorem = String(10 - Len(Trim(xIMPAdValorem)), " ") & Trim(xIMPAdValorem)
xIMPTxTerrOrig = String(10 - Len(Trim(xIMPTxTerrOrig)), " ") & Trim(xIMPTxTerrOrig)
xIMPTxTerrDest = String(10 - Len(Trim(xIMPTxTerrDest)), " ") & Trim(xIMPTxTerrDest)
xIMPTxRedesp = String(10 - Len(Trim(xIMPTxRedesp)), " ") & Trim(xIMPTxRedesp)
xIMPTxAgente = String(10 - Len(Trim(xIMPTxAgente)), " ") & Trim(xIMPTxAgente)
xIMPTxDevTransp = String(10 - Len(Trim(xIMPTxDevTransp)), " ") & Trim(xIMPTxDevTransp)
xIMPTxOutros1 = String(10 - Len(Trim(xIMPTxOutros1)), " ") & Trim(xIMPTxOutros1)
xIMPTxOutros2 = String(10 - Len(Trim(xIMPTxOutros2)), " ") & Trim(xIMPTxOutros2)
xIMPFreteTotal = String(10 - Len(Trim(xIMPFreteTotal)), " ") & Trim(xIMPFreteTotal)



xIMPStrObservacao = UCase(Trim(TxtOBSEmissao.Text))

xIMPStrObservacao = xIMPStrObservacao & String(240 - Len(Trim(xIMPStrObservacao)), " ")
xIMPStrObservacao01 = Trim(Mid(xIMPStrObservacao, 1, 60))
xIMPStrObservacao02 = Trim(Mid(xIMPStrObservacao, 61, 60))
xIMPStrObservacao03 = Trim(Mid(xIMPStrObservacao, 121, 60))

    If UCase(Trim(TxtSiglaCiaAerea.Text)) = "VP" Then
    xIMPStrObservacao04 = Trim("****  A N C  50097  ****")
    Else
    xIMPStrObservacao04 = Trim(Mid(xIMPStrObservacao, 181, 60))
    End If


xIMPObsICMS = ""
xIMPObsPerecivel = ""
xIMPObsSeguro = ""

If xModal = "PAGO" Then
    If xAliquota = 4 Then
    xIMPObsICMS = "ICMS - ALIQUOTA DE 4% - RESOLUCAO 95/96 SENADO FEDERAL "
    End If
    
    If Len(Trim(TxtApoliceExpedidor.Text)) > 0 Then
    xIMPObsSeguro = "Seguradora: " & UCase(Trim(TxtSeguradoraExpedidor.Text)) & "/" & Trim(TxtApoliceExpedidor.Text) & " "
    End If
    
    If xPerecivel = "S" Then
    xIMPObsPerecivel = "P E R E C I V E L - Prazo de Duracao: 48 hs"
    End If
ElseIf xModal = "A PAGAR" Then
    If xAliquota = 4 Then
    xIMPObsICMS = "ICMS - ALIQUOTA DE 4% - RESOLUCAO 95/96 SENADO FEDERAL "
    End If
    
    If Len(Trim(TxtApoliceDestinatario.Text)) > 0 Then
    xIMPObsSeguro = "Seguradora: " & UCase(Trim(TxtSeguradoraExpedidor.Text)) & "/" & Trim(TxtApoliceExpedidor.Text) & " "
    End If
    
    If xPerecivel = "S" Then
    xIMPObsPerecivel = "P E R E C I V E L - Prazo de Duracao: 48 hs"
    End If
End If

xIMPStrRetiraSIM = IIf(OptRetiraSim.Value = True, "XXX", "")
xIMPStrRetiraNAO = IIf(OptRetiraNao.Value = True, "XXX", "")
xIMPStrLocalRetira = xLocalRetirada
xIMPHorarioAt = ""
xIMPStrTelefone = ""
xIMPStrTotalServ = xIMPFreteTotal
xIMPStrBaseCalculo = xIMPFreteTotal
xIMPStrAliquota = TxtAliquota.Text
xIMPStrICMS = TxtICMS.Text


xIMPStrTotalServ = String(10 - Len(Trim(xIMPStrTotalServ)), " ") & Trim(xIMPStrTotalServ)
xIMPStrBaseCalculo = String(10 - Len(Trim(xIMPStrBaseCalculo)), " ") & Trim(xIMPStrBaseCalculo)
xIMPStrAliquota = String(10 - Len(Trim(xIMPStrAliquota)), " ") & Trim(xIMPStrAliquota)
xIMPStrICMS = String(10 - Len(Trim(xIMPStrICMS)), " ") & Trim(xIMPStrICMS)


'xIMPAgenteEmissor = Trim(UCase(TxtNomeFilial.Caption))
xIMPAgenteEmissor = "INTEC CARGO LTDA"
xIMPCodIATA = TxtLicensaFilial.Caption
'xIMPDtEmissao = DataHora("DATA")
'xIMPHoraEmissao = DataHora("HORA")
xIMPDtEmissao = xDataIMP
xIMPHoraEmissao = xHoraIMP
xIMPNaturezaOp = "SERV. TRANSP. AEREO"
xIMPCFOP = "6.63"
'xIMPEmissor = xUsuario
xIMPEmissor = xUsuarioIMP & "/" & xUsuario
xIMPLocalidade = TxtSiglaFilial.Caption
xIMPMatricula = ""

'Declaracao Extritamente de Variaveis para Referencias de Impressao

If UCase(Trim(TxtSiglaCiaAerea.Text)) = "RG" Or UCase(Trim(TxtSiglaCiaAerea.Text)) = "OC" Then
Call MascaraAWBVARIG
ElseIf UCase(Trim(TxtSiglaCiaAerea.Text)) = "P8" Then
Call MascaraAWBP8
ElseIf UCase(Trim(TxtSiglaCiaAerea.Text)) = "VP" Then
Call MascaraAWBVASP
ElseIf UCase(Trim(TxtSiglaCiaAerea.Text)) = "KK" Then
Call MascaraAWBTAM
End If

de_informa.AltStatusFormularioItem "I", DataHora("DATA"), TxtAWB.Text, TxtDig.Text, TxtSiglaCiaAerea.Text, TxtFilial.Caption
de_informa.cn_informa.CommitTrans
End If
'LIMPEZA DA TELA
Call LimpaTela(Me)
FlexGridNFs.Clear
FlexGridNFs.Rows = 0
FlexGridNFs.Cols = 0
FlexGridVolumes.Clear
FlexGridVolumes.Rows = 0
FlexGridVolumes.Cols = 0
ComboProduto.Clear
ComboEspecie.Clear

If de_informa.rsSelEspecie.State = 1 Then de_informa.rsSelEspecie.Close
de_informa.SelEspecie

ComboEspecie.Clear
    Do Until de_informa.rsSelEspecie.EOF
    ComboEspecie.AddItem PriMaiuscula(de_informa.rsSelEspecie.Fields("especie"))
    de_informa.rsSelEspecie.MoveNext
    Loop
    
If de_informa.rsSelProdINT.State = 1 Then de_informa.rsSelProdINT.Close
de_informa.SelProdINT

ComboProduto.Clear
    Do Until de_informa.rsSelProdINT.EOF
    ComboProduto.AddItem PriMaiuscula(de_informa.rsSelProdINT.Fields("descricao"))
    de_informa.rsSelProdINT.MoveNext
    Loop
TxtFreteNacional.Text = "0.00"
TxtFreteRegional.Text = "0.00"
TxtADValorem.Text = "0.00"
TxtTXOrigem.Text = "0.00"
TxtTXDestino.Text = "0.00"
TxtTXRedesp.Text = "0.00"
TxtOutros1.Text = "0.00"
TxtOutros2.Text = "0.00"
TxtFreteTotal.Text = "0.00"

Call TravaFrame(frmEmissao, FraFiliais, 1)
FraSpot.Enabled = False
FraFiliais.Enabled = True

CmdEmitir.Caption = "Gravar AWB"

Acao = "GRAVAR"

TxtBuscaFilial.SetFocus
Me.MousePointer = 0
DoEvents
End Sub

Private Sub CmdDadosExpedidor_Click()
Dim xFrame As Frame
Dim Botao As CommandButton
Dim HMax As Integer
Dim HMin As Integer
Set Botao = ActiveControl


Set xFrame = FraExpedidor
HMax = 3855
HMin = 1755

    If xFrame.Height = HMin Then
    xFrame.ZOrder (0)
    DoEvents
    Call TravaFrame(frmEmissao, xFrame, 0)
    xFrame.Height = HMax
    Botao.Caption = "<"
    DoEvents
    ElseIf xFrame.Height = HMax Then
    Call TravaFrame(frmEmissao, xFrame, 1)
    xFrame.Height = HMin
    Botao.Caption = ">"
    DoEvents
    End If
FraSpot.Enabled = False
End Sub

Private Sub CmdImportarDados_Click()
xUsuarioIMP = ""
xDataIMP = ""
xHoraIMP = ""
AUXCanc = "BUSCAR"
frmEmissaoVia2.Show 1
AUXCanc = ""
xUsuarioIMP = ""
xDataIMP = ""
xHoraIMP = ""
Call TravaFrame(frmEmissao, FraAWB, 1)
FraSpot.Enabled = False
Acao = "GRAVAR"
FraAWB.Enabled = True
TxtAWB.Text = ""
TxtDig.Text = ""
CmdEmitir.Caption = "Gravar AWB"
LblAtualizarFrete.Caption = "Sim"
End Sub

Private Sub CmdLimpaTela_Click()
Call LimpaTela(Me)
Call TravaFrame(frmEmissao, FraAWB, 1)
FraSpot.Enabled = False
FraAWB.Enabled = True
FlexGridNFs.Clear
FlexGridNFs.Rows = 0
FlexGridNFs.Cols = 0
FlexGridVolumes.Clear
FlexGridVolumes.Rows = 0
FlexGridVolumes.Cols = 0
ComboProduto.Clear
ComboEspecie.Clear

If de_informa.rsSelEspecie.State = 1 Then de_informa.rsSelEspecie.Close
de_informa.SelEspecie

ComboEspecie.Clear
    Do Until de_informa.rsSelEspecie.EOF
    ComboEspecie.AddItem PriMaiuscula(de_informa.rsSelEspecie.Fields("especie"))
    de_informa.rsSelEspecie.MoveNext
    Loop
    
If de_informa.rsSelProdINT.State = 1 Then de_informa.rsSelProdINT.Close
de_informa.SelProdINT

ComboProduto.Clear
    Do Until de_informa.rsSelProdINT.EOF
    ComboProduto.AddItem PriMaiuscula(de_informa.rsSelProdINT.Fields("descricao"))
    de_informa.rsSelProdINT.MoveNext
    Loop
TxtFreteNacional.Text = "0.00"
TxtFreteRegional.Text = "0.00"
TxtKiloCob.Text = "0.00"
TxtADValorem.Text = "0.00"
TxtTXOrigem.Text = "0.00"
TxtTXDestino.Text = "0.00"
TxtTXRedesp.Text = "0.00"
TxtOutros1.Text = "0.00"
TxtOutros2.Text = "0.00"
TxtFreteTotal.Text = "0.00"
TxtICMS.Text = "0.00"
CmdEmitir.Caption = "Gravar AWB"
Acao = "GRAVAR"
LblAtualizarFrete.Caption = "Sim"
TxtBuscaFilial.SetFocus
End Sub

Private Sub CmdVia2_Click()
Do While True
Call LimpaTela(Me)
xUsuarioIMP = ""
xDataIMP = ""
xHoraIMP = ""
FlexGridNFs.Clear
FlexGridNFs.Rows = 0
FlexGridNFs.Cols = 0
FlexGridVolumes.Clear
FlexGridVolumes.Rows = 0
FlexGridVolumes.Cols = 0
ComboProduto.Clear
ComboEspecie.Clear
Call TravaFrame(frmEmissao, FraBotoes, 1)
FraSpot.Enabled = False
AUXCanc = "IMPRIMIR"
frmEmissaoIMPRIMIRAWB.Show 1
AUXCanc = ""
    If Acao = "IMPRIMIR" Then
    CmdEmitir_Click
    Else
    Exit Do
    End If
Loop
End Sub

Private Sub ComboEspecie_KeyPress(KeyAscii As Integer)
Dim xTextoVelho As String, xTextoNovo As String, Y As Integer

'Asc (UCase(Chr(KeyAscii)))

    If KeyAscii <> 13 And KeyAscii <> 8 Then
    xTextoVelho = Left(ComboEspecie.Text, ComboEspecie.SelStart) & Chr(KeyAscii)
    xTextoNovo = ""
        For Y = 0 To ComboEspecie.ListCount - 1
            If Len(xTextoVelho) <= Len(ComboEspecie.List(Y)) Then
                If UCase(Mid(ComboEspecie.List(Y), 1, Len(xTextoVelho))) = UCase(xTextoVelho) Then
                xTextoNovo = Mid(ComboEspecie.List(Y), Len(xTextoVelho) + 1)
                Y = ComboEspecie.ListCount
                End If
            End If
        Next
    ComboEspecie.Text = UCase(xTextoVelho) & xTextoNovo
    ComboEspecie.SelStart = Len(xTextoVelho)
    ComboEspecie.SelLength = 1000
    ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    ElseIf KeyAscii = 8 Then
        If Len(ComboEspecie.Text) > 0 Then
            If ComboEspecie.SelStart > 0 Then
            xTextoVelho = Mid(ComboEspecie.Text, 1, ComboEspecie.SelStart - 1)
            Else
            xTextoVelho = Mid(ComboEspecie.Text, 1, ComboEspecie.SelStart)
            End If
        ComboEspecie.Text = UCase(xTextoVelho)
        ComboEspecie.SelStart = Len(xTextoVelho)
        ComboEspecie.SelLength = 1000
        End If
    End If
KeyAscii = 0
End Sub


Private Sub ComboEspecie_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub ComboEspecie_LostFocus()
Dim Y As Integer, xTexto As String

xTexto = ""

        For Y = 0 To ComboEspecie.ListCount - 1
            If UCase(Trim((ComboEspecie.Text))) = UCase(Trim(ComboEspecie.List(Y))) Then
            xTexto = ComboEspecie.List(Y)
            Y = ComboEspecie.ListCount
            End If
        Next
ComboEspecie.Text = xTexto
End Sub

Private Sub ComboProduto_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub


Private Sub Command2_Click()

End Sub


Private Sub OptAPagar_Click()
If OptPago.Value = True Then
OptPago.FontBold = True
OptAPagar.FontBold = False
LblAtualizarFrete.Caption = "Sim"
Else
OptPago.FontBold = False
OptAPagar.FontBold = True
LblAtualizarFrete.Caption = "Sim"
End If
End Sub

Private Sub OptPago_Click()
If OptPago.Value = True Then
OptPago.FontBold = True
OptAPagar.FontBold = False
LblAtualizarFrete.Caption = "Sim"
Else
OptPago.FontBold = False
OptAPagar.FontBold = True
LblAtualizarFrete.Caption = "Sim"
End If
End Sub

Private Sub OptRetiraNao_Click()
If OptRetiraSim.Value = True Then
OptRetiraSim.FontBold = True
OptRetiraNao.FontBold = False
TxtLocalRetirada.Enabled = False
TxtLocalRetirada.BackColor = xBranco
TxtLocalRetirada.Text = ""
LblAtualizarFrete.Caption = "Sim"
Else
OptRetiraSim.FontBold = False
OptRetiraNao.FontBold = True
TxtLocalRetirada.Enabled = True
TxtLocalRetirada.BackColor = xAmarelo
'TxtLocalRetirada.SetFocus
LblAtualizarFrete.Caption = "Sim"
End If
End Sub

Private Sub OptRetiraSim_Click()
If OptRetiraSim.Value = True Then
OptRetiraSim.FontBold = True
OptRetiraNao.FontBold = False
TxtLocalRetirada.Enabled = False
TxtLocalRetirada.BackColor = xBranco
TxtLocalRetirada.Text = ""
LblAtualizarFrete.Caption = "Sim"
Else
OptRetiraSim.FontBold = False
OptRetiraNao.FontBold = True
TxtLocalRetirada.Enabled = True
TxtLocalRetirada.BackColor = xAmarelo
TxtLocalRetirada.SetFocus
LblAtualizarFrete.Caption = "Sim"
End If
End Sub

Private Sub Text1_Change()

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

Private Sub TxtAWB_LostFocus()

If Len(Trim(TxtFilial.Caption)) > 0 And Len(Trim(TxtSiglaCiaAerea.Text)) > 0 And Len(Trim(TxtAWB.Text)) > 0 Then

If de_informa.rsConfereNumeroAWB.State = 1 Then de_informa.rsConfereNumeroAWB.Close
de_informa.ConfereNumeroAWB TxtSiglaCiaAerea.Text, TxtFilial.Caption, TxtAWB.Text

    If de_informa.rsConfereNumeroAWB.RecordCount = 0 Then
    MsgBox "Este formulário não está cadastrado!.", vbCritical, ""
    TxtAWB.Text = ""
    TxtDig.Text = ""
    TxtAWB.SetFocus
    Exit Sub
    ElseIf de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") = "C" Then
    MsgBox "O formulário para este AWB está cancelado. Para utilizá-lo, vá até o cadastro de formulários e descancele-o.", vbCritical, ""
    TxtAWB.Text = ""
    TxtDig.Text = ""
    TxtAWB.SetFocus
    Exit Sub


    ElseIf de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") = "E" Or de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") = "I" Then
    MsgBox "Este AWB já foi emitido. Seus dados podem ser alterados se você escolher Alterar AWB.", vbCritical, ""
    TxtAWB.Text = ""
    TxtDig.Text = ""
    TxtAWB.SetFocus
    Exit Sub
    
    
    
    Else
    TxtDig.Text = de_informa.rsConfereNumeroAWB.Fields("dig")
    End If
Else
TxtAWB.Text = ""
TxtDig.Text = ""
End If
End Sub

Private Sub TxtBuscaFilial_LostFocus()
If Len(Trim(TxtBuscaFilial.Text)) > 0 Then
LblAtualizarFrete.Caption = "Sim"
    TxtBuscaFilial.Text = Trim(String(2 - Len(Trim(Str(Val(TxtBuscaFilial.Text)))), "0")) & Trim(Str(Val(TxtBuscaFilial.Text)))
    If de_informa.rsSelFiliais.State = 1 Then de_informa.rsSelFiliais.Close
    de_informa.SelFiliais TxtBuscaFilial.Text
    
    If de_informa.rsSelFiliais.RecordCount > 0 Then
        If IsNull(de_informa.rsSelFiliais.Fields("filial")) = False Then TxtFilial.Caption = de_informa.rsSelFiliais.Fields("filial")
        If IsNull(de_informa.rsSelFiliais.Fields("nomefilial")) = False Then TxtNomeFilial.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("nomefilial"))
        If IsNull(de_informa.rsSelFiliais.Fields("cgc")) = False Then TxtCGCFilial.Caption = de_informa.rsSelFiliais.Fields("cgc")
        If IsNull(de_informa.rsSelFiliais.Fields("inscrest")) = False Then TxtInscrEstFilial.Caption = de_informa.rsSelFiliais.Fields("inscrest")
        If IsNull(de_informa.rsSelFiliais.Fields("cidade")) = False Then TxtCidadeFilial.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("cidade"))
        If IsNull(de_informa.rsSelFiliais.Fields("uf")) = False Then TxtUFFilial.Caption = de_informa.rsSelFiliais.Fields("uf")
        If IsNull(de_informa.rsSelFiliais.Fields("licensaIATA")) = False Then TxtLicensaFilial.Caption = de_informa.rsSelFiliais.Fields("licensaIATA")
        If IsNull(de_informa.rsSelFiliais.Fields("siglaIATA")) = False Then TxtSiglaFilial.Caption = de_informa.rsSelFiliais.Fields("siglaIATA")
    DoEvents
    End If
End If
End Sub

Private Sub TxtBuscaSiglaCia_LostFocus()
If Len(Trim(TxtBuscaSiglaCia.Text)) > 0 Then
LblAtualizarFrete.Caption = "Sim"
    If de_informa.rsSelCiaAerea.State = 1 Then de_informa.rsSelCiaAerea.Close
    de_informa.SelCiaAerea Trim(UCase(TxtBuscaSiglaCia.Text))
    
        If de_informa.rsSelCiaAerea.RecordCount > 0 Then
        TxtSiglaCiaAerea.Text = UCase(de_informa.rsSelCiaAerea.Fields("codcia"))
        TxtNomeCiaAerea.Caption = PriMaiuscula(de_informa.rsSelCiaAerea.Fields("fantasia"))
        TxtCGCCiaAerea.Caption = de_informa.rsSelCiaAerea.Fields("cgc")
        TxtInscrEstCiaAerea.Caption = de_informa.rsSelCiaAerea.Fields("inscrest")
        Else
        TxtSiglaCiaAerea.Text = ""
        TxtNomeCiaAerea.Caption = ""
        TxtCGCCiaAerea.Caption = ""
        TxtInscrEstCiaAerea.Caption = ""
        MsgBox "Sigla de Cia. Aérea não encontrada! Por favor, tente novamente...", vbCritical, ""
        TxtBuscaSiglaCia.SetFocus
        End If
Else
TxtNomeCiaAerea.Caption = ""
TxtCGCCiaAerea.Caption = ""
TxtInscrEstCiaAerea.Caption = ""
End If
End Sub

Private Sub TxtCodIATA_LostFocus()
With TxtCodIATA
    If Len(Trim(.Text)) > 0 Then
        LblAtualizarFrete.Caption = "Sim"
        If Trim(.Text) = "?" Then
        frmEmissaoCODSIATA.Show 1
        Else
        .Text = String(3 - Len(Trim(Str(Val(.Text)))), "0") & Trim(Str(Val(.Text)))
        If de_informa.rsSelIATA.State = 1 Then de_informa.rsSelIATA.Close
        de_informa.SelIATA .Text
            If de_informa.rsSelIATA.RecordCount > 0 Then
            TxtDescrIATA.Text = .Text & " - " & PriMaiuscula(Trim(de_informa.rsSelIATA.Fields("descricao")))
            Else
            MsgBox "Este Código não foi encontrado!", vbCritical, ""
            End If
        End If
    End If

End With

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

Private Sub TxtFreteRegional_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtFreteRegional.Text = SoNumero(TxtFreteRegional.Text)
            TxtFreteRegional.Text = Mid(TxtFreteRegional.Text, 1, Len(TxtFreteRegional.Text) - 1)
            TxtFreteRegional.Text = Val(TxtFreteRegional.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtFreteRegional.Text = SoNumero(TxtFreteRegional.Text) & Chr(KeyAscii)
    TxtFreteRegional.Text = Val(TxtFreteRegional.Text) / 100
    KeyAscii = 0
    End If

TxtFreteRegional.Text = Format(TxtFreteRegional.Text, "###,###,###,##0.00")
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

Private Sub TxtTXOrigem_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtTXOrigem.Text = SoNumero(TxtTXOrigem.Text)
            TxtTXOrigem.Text = Mid(TxtTXOrigem.Text, 1, Len(TxtTXOrigem.Text) - 1)
            TxtTXOrigem.Text = Val(TxtTXOrigem.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtTXOrigem.Text = SoNumero(TxtTXOrigem.Text) & Chr(KeyAscii)
    TxtTXOrigem.Text = Val(TxtTXOrigem.Text) / 100
    KeyAscii = 0
    End If

TxtTXOrigem.Text = Format(TxtTXOrigem.Text, "###,###,###,##0.00")
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

Private Sub TxtAutorizador_Change()
If Len(Trim(TxtAutorizador.Text)) > 0 Then
TxtAutorizador.Text = UCase(TxtAutorizador.Text)
TxtAutorizador.SelStart = Len(TxtAutorizador.Text)
End If
End Sub

Private Sub TxtAutorizador_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtAutorizador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtAWB_GotFocus()
TxtAWB.SelStart = 0
TxtAWB.SelLength = 100
End Sub

Private Sub TxtAWB_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtBuscaDestinatario_Change()
If Len(Trim(TxtBuscaDestinatario.Text)) > 0 Then
TxtBuscaDestinatario.Text = UCase(TxtBuscaDestinatario.Text)
TxtBuscaDestinatario.SelStart = Len(TxtBuscaDestinatario.Text)
End If
End Sub

Private Sub TxtBuscaDestinatario_GotFocus()
TxtBuscaDestinatario.SelStart = 0
TxtBuscaDestinatario.SelLength = 100
End Sub

Private Sub TxtBuscaExpedidor_Change()
If Len(Trim(TxtBuscaExpedidor.Text)) > 0 Then
TxtBuscaExpedidor.Text = UCase(TxtBuscaExpedidor.Text)
TxtBuscaExpedidor.SelStart = Len(TxtBuscaExpedidor.Text)
End If
End Sub

Private Sub TxtBuscaExpedidor_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtBuscaFilial_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtBuscaSiglaCia_Change()
If Len(Trim(TxtBuscaSiglaCia.Text)) > 0 Then
TxtBuscaSiglaCia.Text = UCase(TxtBuscaSiglaCia.Text)
TxtBuscaSiglaCia.SelStart = Len(TxtBuscaSiglaCia.Text)
End If
End Sub

Private Sub TxtBuscaSiglaCia_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtBuscaSiglaDEST_Change()
If Len(Trim(TxtBuscaSiglaDEST.Text)) > 0 Then
TxtBuscaSiglaDEST.Text = UCase(TxtBuscaSiglaDEST.Text)
TxtBuscaSiglaDEST.SelStart = Len(TxtBuscaSiglaDEST.Text)
End If
End Sub

Private Sub TxtBuscaSiglaDEST_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub


Private Sub TxtBuscaSiglaExp_Change()
If Len(Trim(TxtBuscaSiglaExp.Text)) > 0 Then
TxtBuscaSiglaExp.Text = UCase(TxtBuscaSiglaExp.Text)
TxtBuscaSiglaExp.SelStart = Len(TxtBuscaSiglaExp.Text)
End If
End Sub

Private Sub TxtBuscaSiglaExp_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtBuscaSiglaVIA_Change()
If Len(Trim(TxtBuscaSiglaVIA.Text)) > 0 Then
TxtBuscaSiglaVIA.Text = UCase(TxtBuscaSiglaVIA.Text)
TxtBuscaSiglaVIA.SelStart = Len(TxtBuscaSiglaVIA.Text)
End If
End Sub

Private Sub TxtBuscaSiglaVIA_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtCodIATA_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub TxtDescrOutros_Change()
If Len(TxtDescrOutros.Text) > 0 Then
TxtDescrOutros.Text = UCase(TxtDescrOutros.Text)
TxtDescrOutros.SelStart = Len(TxtDescrOutros.Text)
End If
End Sub

Private Sub TxtDescrOutros_GotFocus()
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

Private Sub TxtDig_GotFocus()
TxtDig.SelStart = 0
TxtDig.SelLength = 3
End Sub

Private Sub TxtDig_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtKilo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
            If KeyAscii = 8 Then
            TxtKilo.Text = SoNumero(TxtKilo.Text)
            TxtKilo.Text = Mid(TxtKilo.Text, 1, Len(TxtKilo.Text) - 1)
            TxtKilo.Text = Val(TxtKilo.Text) / 100
            KeyAscii = 0
            End If
        End If
    Else
    TxtKilo.Text = SoNumero(TxtKilo.Text) & Chr(KeyAscii)
    TxtKilo.Text = Val(TxtKilo.Text) / 100
    KeyAscii = 0
    End If

TxtKilo.Text = Format(TxtKilo.Text, "###,###,###,##0.00")
End Sub

Private Sub CmdBuscaCiaAerea_Click()
    frmEmissaoBuscaCiaAerea.Show 1
End Sub

Private Sub ChkPerecivel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub CmdCalcularTaxas_Click()

If Len(Trim(TxtSiglaCiaAerea)) = 0 Then
MsgBox "É preciso que você Especifique uma Cia. Aérea antes.", vbInformation, ""
TxtBuscaSiglaCia.SetFocus
Exit Sub
ElseIf Len(Trim(TxtCGCExpedidor.Text)) = 0 Then
MsgBox "É preciso que você Especifique um Expedidor antes.", vbInformation, ""
TxtBuscaExpedidor.SetFocus
Exit Sub
ElseIf Len(Trim(TxtSiglaExpedidor.Text)) = 0 Then
MsgBox "É preciso que você Especifique uma Origem antes.", vbInformation, ""
TxtBuscaSiglaExp.SetFocus
Exit Sub
ElseIf Len(Trim(TxtCGCDestinatario.Text)) = 0 Then
MsgBox "É preciso que você Especifique um Destinatario antes.", vbInformation, ""
TxtBuscaDestinatario.SetFocus
Exit Sub
ElseIf Len(Trim(TxtSiglaDestinatario.Text)) = 0 Then
MsgBox "É preciso que você Especifique um Destino antes.", vbInformation, ""
TxtBuscaSiglaDEST.SetFocus
Exit Sub
ElseIf Len(Trim(TxtSiglaVIA.Text)) = 0 Then
MsgBox "É preciso que você Especifique uma VIA antes.", vbInformation, ""
TxtBuscaSiglaVIA.SetFocus
Exit Sub
'ElseIf Val(TxtPesoCubado.Text) < 1 Then
'MsgBox "O Peso Cubado não pode ser menor que 1.", vbInformation, ""
'TxtPesoCubado.SetFocus
'Exit Sub
ElseIf Val(TxtPesoReal.Text) < 1 Then
MsgBox "O Peso Real não pode ser menor que 1.", vbInformation, ""
TxtPesoReal.SetFocus
Exit Sub
ElseIf Len(Trim(TxtDescrIATA.Text)) = 0 Then
MsgBox "Você não Especificou um Produto.", vbExclamation, ""
Exit Sub
End If

Dim xCodCia, xOrigem, xDestino, xCGC, xCodTETC As String
Dim xCodTab As Integer
Dim xTxKgGERAL, xTxKgTETC, xDescTETC, xDescGERAL, xPesoTx, xTxMin, xCharter, xCorteCharter, _
xTxTerrestre, xCorteTerrestre, xExcedTerrestre, xFreteNacional, xMenorTx, xVlMerc, xADVAL, _
xTipoADVal, xICMS As Currency

If IsNumeric(TxtPesoCubado.Text) = False Then TxtPesoCubado.Text = 0
If IsNumeric(TxtPesoReal.Text) = False Then TxtPesoReal.Text = 0

xVlMerc = TxtTotalVM.Text
xCodCia = TxtSiglaCiaAerea.Text
xTxKgGERAL = 0
xDescGERAL = 0
xTxKgTETC = 0
xDescTETC = 0
xCGC = Mid(TxtCGCExpedidor.Text, 1, 8)


'***********Alteração - Lincoln - Pesos em intervalos de 0.5 **********************Start
'***********Calcula o decimal ***********

If CDbl(TxtPesoCubado.Text) > CDbl(TxtPesoReal.Text) Then
    xPesoTx = (CDbl(TxtPesoCubado.Text))
Else
    xPesoTx = (CDbl(TxtPesoReal.Text))
End If

'***********Alteração - Lincoln - Pesos em intervalos de 0.5 **********************End

If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
de_informa.SelAeroportoSigla TxtSiglaExpedidor.Text
xOrigem = de_informa.rsSelAeroportoSigla.Fields("localidade")

If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
de_informa.SelAeroportoSigla TxtSiglaVIA.Text
xDestino = de_informa.rsSelAeroportoSigla.Fields("localidade")


If de_informa.rsSelCODTAB.State = 1 Then de_informa.rsSelCODTAB.Close
de_informa.SelCODTAB xCodCia, xOrigem, xCGC & "%"

    If de_informa.rsSelCODTAB.RecordCount > 0 Then
    xCodTab = de_informa.rsSelCODTAB.Fields("codtab")
    Else
    If de_informa.rsSelCODTAB.State = 1 Then de_informa.rsSelCODTAB.Close
    de_informa.SelCODTAB xCodCia, xOrigem, "%"
        If de_informa.rsSelCODTAB.RecordCount > 0 Then
        xCodTab = de_informa.rsSelCODTAB.Fields("codtab")
        xTxTerrestre = de_informa.rsSelCODTAB.Fields("taxadestino")
        xCorteTerrestre = de_informa.rsSelCODTAB.Fields("cortedestino")
        xExcedTerrestre = de_informa.rsSelCODTAB.Fields("exceddestino")
        Else
        MsgBox "Não existe Tabela cadastrada para estes parâmetros. Verifique os dados inseridos ou contate o Administador do Sistema.", vbCritical, ""
        Exit Sub
        End If
    End If
    
        

If de_informa.rsSelPrecosGERAL.State = 1 Then de_informa.rsSelPrecosGERAL.Close
de_informa.SelPrecosGERAL xCodTab, xDestino

    If de_informa.rsSelPrecosGERAL.RecordCount = 0 Then
    MsgBox "Não Existe Tabela Cadastrada para este Destino. Verifique os dados inseridos ou contate o Administador do Sistema.", vbCritical, ""
    Exit Sub
    Else
    xTxMin = de_informa.rsSelPrecosGERAL.Fields("taxaminima")
        If xPesoTx <= 25.5 Then
        xTxKgGERAL = de_informa.rsSelPrecosGERAL.Fields("ate25")
        ElseIf xPesoTx <= 50.5 Then
        xTxKgGERAL = de_informa.rsSelPrecosGERAL.Fields("ate50")
        ElseIf xPesoTx <= 300.5 Then
        xTxKgGERAL = de_informa.rsSelPrecosGERAL.Fields("ate300")
        ElseIf xPesoTx <= 500.5 Then
        xTxKgGERAL = de_informa.rsSelPrecosGERAL.Fields("ate500")
        ElseIf xPesoTx <= 1000.5 Then
        xTxKgGERAL = de_informa.rsSelPrecosGERAL.Fields("ate1000")
        Else
        xTxKgGERAL = de_informa.rsSelPrecosGERAL.Fields("acima1000")
        End If
    xDescGERAL = de_informa.rsSelPrecosGERAL.Fields("descontogeral")
    xCharter = de_informa.rsSelPrecosGERAL.Fields("charter")
    xCorteCharter = de_informa.rsSelPrecosGERAL.Fields("cortecharter")
    xTxDestino = Trim(de_informa.rsSelPrecosGERAL.Fields("txterrestre"))
    End If

    If Len(Trim(TxtDescrIATA.Text)) > 0 Then
    xCodTETC = Mid(Trim(TxtDescrIATA.Text), 1, 3)
    If de_informa.rsSelPrecosTETC.State = 1 Then de_informa.rsSelPrecosTETC.Close
    de_informa.SelPrecosTETC xCodTab, xDestino, xCodTETC
        If de_informa.rsSelPrecosTETC.RecordCount > 0 Then
            If de_informa.rsSelPrecosTETC.Fields("usargeral") <> "S" Then
            xTxKgTETC = de_informa.rsSelPrecosTETC.Fields("porkilo")
            xDescTETC = de_informa.rsSelPrecosTETC.Fields("desconto")
            End If
        ElseIf xCodTETC <> "000" Then
            If MsgBox("Não existe Tabela Específica para estes parâmetros... Você deseja utilizar a Tabela Geral para este caso?", vbYesNo + vbExclamation, "") = vbNo Then
            MsgBox "Contate o Administrador do Sistema.", vbExclamation, ""
            Exit Sub
            End If
        End If
    End If
    
    TxtTipoTaxa.Text = ""
    
    If xTxKgTETC > 0 Then
        TxtKiloCob.Text = Format(CDbl(xTxKgTETC - ((xTxKgTETC / 100) * xDescTETC)), "###,###,###,##0.00")
        TxtTipoTaxa.Text = "Específica"
    Else
        TxtKiloCob.Text = Format(CDbl(xTxKgGERAL - ((xTxKgGERAL / 100) * xDescGERAL)), "###,###,###,##0.00")
        TxtTipoTaxa.Text = "Geral"
    End If

    xFreteNacional = TxtKiloCob.Text * xPesoTx


    '*************Alteração - Arredondar cálculo p/  duas casas  - Lincoln ***************** Start
    If xCodCia = "RG" Then
        TxtKiloCob.Text = Format(CDbl(TxtKiloCob.Text), "###,###,###,##0.00")
    End If
    '*************Alteração - Arredondar cálculo p/  duas casas  - Lincoln ***************** End
    
    
    If xFreteNacional < xTxMin Then
        xFreteNacional = xTxMin
        TxtTipoTaxa.Text = "Mínima"
        TxtKiloCob.Text = ""
    End If
    
    If xCharter > 0 Then
        If xPesoTx >= xCorteCharter Then
        xFreteNacional = xPesoTx * xCharter
        TxtTipoTaxa.Text = "Charter"
        TxtKiloCob.Text = xCharter
        End If
    End If
    
    TxtFreteNacional.Text = xFreteNacional
    TxtFreteRegional.Text = "0,00"
    TxtTXOrigem.Text = "0,00"
    
        If xCodCia = "RG" Then
            If (xPesoTx - 10) <= 300 Then
            xExcedTerrestre = 0.25
            ElseIf (xPesoTx - 10) <= 1000 Then
            xExcedTerrestre = 0.18
            Else
            xExcedTerrestre = 0.12
            End If
        ElseIf xCodCia = "P8" And xDestino = "ARARAQUARA" Then
        xTxTerrestre = 13.58
        xCorteTerrestre = 10
        xExcedTerrestre = 0.38
        End If
    
        If OptRetiraSim.Value = True And xTxDestino = "S" And xDestino <> "MANAUS" Then
            If xTxTerrestre > 0 Then
                If xPesoTx >= xCorteTerrestre Then
                TxtTXDestino.Text = xTxTerrestre + ((xPesoTx - xCorteTerrestre) * xExcedTerrestre)
                Else
                TxtTXDestino.Text = xTxTerrestre
                End If
            Else
            TxtTXDestino.Text = "0,00"
            End If
        TxtLocalRetirada.Text = "RETIRA AGENCIA"
        ElseIf OptRetiraNao.Value = True Then
            If xTxTerrestre > 0 Then
                If xPesoTx >= xCorteTerrestre Then
                TxtTXDestino.Text = xTxTerrestre + ((xPesoTx - xCorteTerrestre) * xExcedTerrestre)
                Else
                TxtTXDestino.Text = xTxTerrestre
                End If
            Else
            TxtTXDestino.Text = "0,00"
            End If
        ElseIf OptRetiraSim.Value = True And xTxDestino = "" Then
        TxtTXDestino.Text = "0,00"
        TxtLocalRetirada.Text = "RETIRA AEROPORTO"
        End If

    
If de_informa.rsSelMenorTxGERAL.State = 1 Then de_informa.rsSelMenorTxGERAL.Close
de_informa.SelMenorTxGERAL xCodTab, xDestino

If de_informa.rsSelMenorTxTETC.State = 1 Then de_informa.rsSelMenorTxTETC.Close
de_informa.SelMenorTxtetc xCodTab, xDestino

    If de_informa.rsSelMenorTxTETC.Fields("porkilo") < de_informa.rsSelMenorTxGERAL.Fields("acima1000") Then
    xMenorTx = de_informa.rsSelMenorTxTETC.Fields("porkilo")
    Else
    xMenorTx = de_informa.rsSelMenorTxGERAL.Fields("acima1000")
    End If
    
    
    
    If (xVlMerc / xPesoTx) >= (xMenorTx * 100) Then
    xTipoADVal = 2
    xADVAL = 0.66
    Else
    xTipoADVal = 1
    xADVAL = 0.33
    End If
    
    TxtTipoADVAL.Text = xTipoADVal
    
    If Len(Trim(TxtApoliceExpedidor.Text)) > 0 And OptPago.Value = True Then
    TxtADValorem.Text = 0
    Else
        If Len(Trim(TxtApoliceDestinatario.Text)) > 0 And OptAPagar.Value = True Then
        TxtADValorem.Text = 0
        Else
        TxtADValorem.Text = (xVlMerc * (xADVAL / 100))
        TxtTipoADVAL.Text = xTipoADVal
        End If
    End If
    
    If Mid(Trim(TxtCodIATA), 1, 3) = "001" Or Mid(Trim(TxtCodIATA), 1, 3) = "140" Then
    TxtAliquota.Text = "ISENTO"
    Else
        If Len(Trim(TxtInscrEstDestinatario.Text)) > 0 And UCase(Trim(TxtInscrEstDestinatario.Text)) <> "ISENTO" Then
        TxtAliquota.Text = "4"
        Else
        TxtAliquota.Text = "12"
        End If
    End If
    
'***********Alteração - Lincoln - Inclusão de tx transportador **********************Start 24/09/2004
If xCodCia = "RG" Then
    If de_informa.rsSel_destino.State = 1 Then
        de_informa.rsSel_destino.Close
    End If
    
    de_informa.Sel_destino TxtSiglaDestinatario
    
    If de_informa.rsSel_destino.EOF = False Then
        If de_informa.rsSel_destino.Fields("regiaogeo") = "NORTE" Then
            Txt_transp.Text = 0.3
        ElseIf de_informa.rsSel_destino.Fields("regiaogeo") = "NORDESTE" Then
            Txt_transp.Text = 0.3
        ElseIf de_informa.rsSel_destino.Fields("regiaogeo") = "CENTRO-OESTE" Then
            Txt_transp.Text = 0.22
        ElseIf de_informa.rsSel_destino.Fields("regiaogeo") = "CENTRO -OESTE" Then
            Txt_transp.Text = 0.22
        ElseIf de_informa.rsSel_destino.Fields("regiaogeo") = "SUL" Then
            Txt_transp.Text = 0.15
        ElseIf de_informa.rsSel_destino.Fields("regiaogeo") = "SUDESTE" Then
            Txt_transp.Text = 0.15
        Else
            Txt_transp.Text = 0.15
        End If
    End If
        
    Txt_transp.Text = CDbl(Txt_transp.Text) * CDbl(xPesoTx)
End If


'VARIG é cobrado taxa de transporte
If xCodCia = "RG" Then
    TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text) + CDbl(Txt_transp.Text)), "###,###,###,##0.00")
Else
    TxtFreteTotal.Text = Format((CDbl(TxtFreteNacional.Text) + CDbl(TxtFreteRegional.Text) + CDbl(TxtADValorem.Text) + CDbl(TxtTXOrigem.Text) + CDbl(TxtTXDestino.Text) + CDbl(TxtTXRedesp.Text) + CDbl(TxtOutros1.Text) + CDbl(TxtOutros2.Text)), "###,###,###,##0.00")
End If
'***********Alteração - Lincoln - Inclusão de tx transportador **********************Start 24/09/2004


Txt_transp.Text = Format(Txt_transp.Text, "###,###,###,##0.00")
TxtKiloCob.Text = Format(TxtKiloCob.Text, "###,###,###,##0.00")
TxtFreteNacional.Text = Format(TxtFreteNacional.Text, "###,###,###,##0.00")
TxtADValorem.Text = Format(TxtADValorem.Text, "###,###,###,##0.00")
TxtTXOrigem.Text = Format(TxtTXOrigem.Text, "###,###,###,##0.00")
TxtTXDestino.Text = Format(TxtTXDestino.Text, "###,###,###,##0.00")
TxtTXRedesp.Text = Format(TxtTXRedesp.Text, "###,###,###,##0.00")
TxtOutros1.Text = Format(TxtOutros1.Text, "###,###,###,##0.00")
TxtOutros2.Text = Format(TxtOutros2.Text, "###,###,###,##0.00")

If TxtAliquota.Text <> "ISENTO" Then
xICMS = ((Val(SemPonto(TxtFreteTotal.Text)) / 100) * Val(TxtAliquota.Text)) / 100
Else
xICMS = 0
End If

TxtICMS.Text = xICMS
TxtICMS.Text = Format(TxtICMS.Text, "###,###,###,##0.00")

LblAtualizarFrete.Caption = "Nao"
CmdTarifaSpot.Enabled = True


End Sub

Private Sub CmdIncluirNF_Click()
LblAtualizarFrete.Caption = "Sim"
frmEmissaoIncluiNF.Show 1
End Sub

Private Sub CmdIncluirVolume_Click()
LblAtualizarFrete.Caption = "Sim"
frmEmissaoIncluiVolume.Show 1
End Sub


Private Sub CmdTarifaSpot_Click()
FraAeoportos.Enabled = False
FraCiaAerea.Enabled = False
FraDestinatario.Enabled = False
FraEspecie.Enabled = False
FraExpedidor.Enabled = False
FraFiliais.Enabled = False
FraIATA.Enabled = False
FraModalFrete.Enabled = False
FraNFs.Enabled = False
FraOBS.Enabled = False
FraProduto.Enabled = False
FraRetira.Enabled = False
FraTaxas.Enabled = False
FraVolumes.Enabled = False
FraSpot.Enabled = True

TxtAutorizador.Enabled = True
TxtKilo.Enabled = True

TxtAutorizador.BackColor = xAmarelo
TxtKilo.BackColor = xAmarelo

TxtAutorizador.SetFocus



End Sub


Private Sub ComboProduto_Click()
    If ComboProduto.Text = "Outros" Then
    TxtDescrOutros.Text = ""
    TxtDescrOutros.BackColor = xAmarelo
    TxtDescrOutros.Enabled = True
    TxtDescrOutros.SetFocus
    Else
    TxtDescrOutros.Text = ""
    TxtDescrOutros.BackColor = xBranco
    TxtDescrOutros.Enabled = False
    End If
End Sub

Private Sub ComboProduto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
Else
KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
Acao = "GRAVAR"

CmdEmitir.Caption = "Gravar AWB"
If de_informa.rsSelEspecie.State = 1 Then de_informa.rsSelEspecie.Close
de_informa.SelEspecie

ComboEspecie.Clear
    Do Until de_informa.rsSelEspecie.EOF
    ComboEspecie.AddItem UCase(de_informa.rsSelEspecie.Fields("especie"))
    de_informa.rsSelEspecie.MoveNext
    Loop
    
If de_informa.rsSelProdINT.State = 1 Then de_informa.rsSelProdINT.Close
de_informa.SelProdINT

ComboProduto.Clear
    Do Until de_informa.rsSelProdINT.EOF
    ComboProduto.AddItem PriMaiuscula(de_informa.rsSelProdINT.Fields("descricao"))
    de_informa.rsSelProdINT.MoveNext
    Loop


'StringDireitos = Mid(StringDireitos, 1, 35) & "1" & Mid(StringDireitos, 37)
End Sub

Private Sub OptAPagar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub OptPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub OptRetiraNao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub OptRetiraSim_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtBuscaDestinatario_KeyPress(KeyAscii As Integer)
Dim xBusca As String

If KeyAscii = 13 Then
    With TxtBuscaDestinatario
        If Len(Trim(.Text)) > 0 Then
        LblAtualizarFrete.Caption = "Sim"
        xBusca = .Text
        Call LimpaFrame(frmEmissao, FraDestinatario.Caption)
        TxtNomeDestinatario.Text = ""
        TxtCGCDestinatario.Text = ""
        TxtCidadeDestinatario.Text = ""
        TxtUFDestinatario.Text = ""
        TxtInscrEstDestinatario.Text = ""
        TxtEndDestinatario.Text = ""
        TxtBairroDEST.Text = ""
        TxtCEPDestinatario.Text = ""
        TxtTelDestinatario.Text = ""
        TxtFAXDestinatario.Text = ""
       
        If de_informa.rsSelClienteAPELIDO.State = 1 Then de_informa.rsSelClienteAPELIDO.Close
        If de_informa.rsSelClienteCNPJ.State = 1 Then de_informa.rsSelClienteCNPJ.Close
        If de_informa.rsSelClienteFANTASIA.State = 1 Then de_informa.rsSelClienteFANTASIA.Close
        If de_informa.rsSelClienteNOME.State = 1 Then de_informa.rsSelClienteNOME.Close
        de_informa.SelClienteAPELIDO xBusca & "%"
            If de_informa.rsSelClienteAPELIDO.RecordCount = 1 Then
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("nome")) = False Then TxtNomeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("nome"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("cgc")) = False Then TxtCGCDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("cgc"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("cidade")) = False Then TxtCidadeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("cidade"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("uf")) = False Then TxtUFDestinatario.Text = de_informa.rsSelClienteAPELIDO.Fields("uf")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("ie")) = False Then TxtInscrEstDestinatario.Text = de_informa.rsSelClienteAPELIDO.Fields("ie")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("endereco")) = False Then TxtEndDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("endereco"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("bairro")) = False Then TxtBairroDEST.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("bairro"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("cep")) = False Then TxtCEPDestinatario.Text = de_informa.rsSelClienteAPELIDO.Fields("cep")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("pabx")) = False Then TxtTelDestinatario.Text = de_informa.rsSelClienteAPELIDO.Fields("pabx")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("fax")) = False Then TxtFAXDestinatario.Text = de_informa.rsSelClienteAPELIDO.Fields("fax")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("seguradora")) = False Then TxtSeguradoraDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("seguradora"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("APOLICE")) = False Then TxtApoliceDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("APOLICE"))
            SendKeys "{TAB}"
            KeyAscii = 0
            ElseIf de_informa.rsSelClienteAPELIDO.RecordCount > 1 Then
            frmEmissaoFiltraDest.Show 1
            SendKeys "{TAB}"
            KeyAscii = 0
            Else
            de_informa.rsSelClienteAPELIDO.Close
            de_informa.SelClienteCNPJ xBusca & "%"
                If de_informa.rsSelClienteCNPJ.RecordCount = 1 Then
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("nome")) = False Then TxtNomeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("nome"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("cgc")) = False Then TxtCGCDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("cgc"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("cidade")) = False Then TxtCidadeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("cidade"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("uf")) = False Then TxtUFDestinatario.Text = de_informa.rsSelClienteCNPJ.Fields("uf")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("ie")) = False Then TxtInscrEstDestinatario.Text = de_informa.rsSelClienteCNPJ.Fields("ie")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("endereco")) = False Then TxtEndDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("endereco"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("bairro")) = False Then TxtBairroDEST.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("bairro"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("cep")) = False Then TxtCEPDestinatario.Text = de_informa.rsSelClienteCNPJ.Fields("cep")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("pabx")) = False Then TxtTelDestinatario.Text = de_informa.rsSelClienteCNPJ.Fields("pabx")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("fax")) = False Then TxtFAXDestinatario.Text = de_informa.rsSelClienteCNPJ.Fields("fax")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("seguradora")) = False Then TxtSeguradoraDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("seguradora"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("APOLICE")) = False Then TxtApoliceDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("APOLICE"))
                SendKeys "{TAB}"
                KeyAscii = 0
                ElseIf de_informa.rsSelClienteCNPJ.RecordCount > 1 Then
                frmEmissaoFiltraDest.Show 1
                SendKeys "{TAB}"
                KeyAscii = 0
                Else
                de_informa.rsSelClienteCNPJ.Close
                de_informa.SelClientefantasia xBusca & "%"
                    If de_informa.rsSelClienteFANTASIA.RecordCount = 1 Then
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("nome")) = False Then TxtNomeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("nome"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("cgc")) = False Then TxtCGCDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("cgc"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("cidade")) = False Then TxtCidadeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("cidade"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("uf")) = False Then TxtUFDestinatario.Text = de_informa.rsSelClienteFANTASIA.Fields("uf")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("ie")) = False Then TxtInscrEstDestinatario.Text = de_informa.rsSelClienteFANTASIA.Fields("ie")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("endereco")) = False Then TxtEndDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("endereco"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("bairro")) = False Then TxtBairroDEST.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("bairro"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("cep")) = False Then TxtCEPDestinatario.Text = de_informa.rsSelClienteFANTASIA.Fields("cep")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("pabx")) = False Then TxtTelDestinatario.Text = de_informa.rsSelClienteFANTASIA.Fields("pabx")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("fax")) = False Then TxtFAXDestinatario.Text = de_informa.rsSelClienteFANTASIA.Fields("fax")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("seguradora")) = False Then TxtSeguradoraDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("seguradora"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("APOLICE")) = False Then TxtApoliceDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("APOLICE"))
                    SendKeys "{TAB}"
                    KeyAscii = 0
                    ElseIf de_informa.rsSelClienteFANTASIA.RecordCount > 1 Then
                    frmEmissaoFiltraDest.Show 1
                    SendKeys "{TAB}"
                    KeyAscii = 0
                    Else
                    de_informa.rsSelClienteFANTASIA.Close
                    de_informa.SelClientenome "%" & xBusca & "%"
                        If de_informa.rsSelClienteNOME.RecordCount = 1 Then
                        If IsNull(de_informa.rsSelClienteNOME.Fields("nome")) = False Then TxtNomeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("nome"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("cgc")) = False Then TxtCGCDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("cgc"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("cidade")) = False Then TxtCidadeDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("cidade"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("uf")) = False Then TxtUFDestinatario.Text = de_informa.rsSelClienteNOME.Fields("uf")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("ie")) = False Then TxtInscrEstDestinatario.Text = de_informa.rsSelClienteNOME.Fields("ie")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("endereco")) = False Then TxtEndDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("endereco"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("bairro")) = False Then TxtBairroDEST.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("bairro"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("cep")) = False Then TxtCEPDestinatario.Text = de_informa.rsSelClienteNOME.Fields("cep")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("pabx")) = False Then TxtTelDestinatario.Text = de_informa.rsSelClienteNOME.Fields("pabx")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("fax")) = False Then TxtFAXDestinatario.Text = de_informa.rsSelClienteNOME.Fields("fax")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("seguradora")) = False Then TxtSeguradoraDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("seguradora"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("APOLICE")) = False Then TxtApoliceDestinatario.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("APOLICE"))
                        SendKeys "{TAB}"
                        KeyAscii = 0
                        ElseIf de_informa.rsSelClienteNOME.RecordCount > 1 Then
                        frmEmissaoFiltraDest.Show 1
                        SendKeys "{TAB}"
                        KeyAscii = 0
                        Else
                        MsgBox "Termo não encontrado!", vbCritical, ""
                        TxtBuscaDestinatario.SetFocus
                        End If
                    End If
                End If
            End If
        Else
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    End With
End If
End Sub

Private Sub TxtBuscaExpedidor_KeyPress(KeyAscii As Integer)
Dim xBusca As String

If KeyAscii = 13 Then
    With TxtBuscaExpedidor
        If Len(Trim(.Text)) > 0 Then
        LblAtualizarFrete.Caption = "Sim"
        xBusca = .Text
        
        Call LimpaFrame(frmEmissao, FraExpedidor.Caption)
        
        If de_informa.rsSelClienteAPELIDO.State = 1 Then de_informa.rsSelClienteAPELIDO.Close
        If de_informa.rsSelClienteCNPJ.State = 1 Then de_informa.rsSelClienteCNPJ.Close
        If de_informa.rsSelClienteFANTASIA.State = 1 Then de_informa.rsSelClienteFANTASIA.Close
        If de_informa.rsSelClienteNOME.State = 1 Then de_informa.rsSelClienteNOME.Close
        de_informa.SelClienteAPELIDO xBusca & "%"
            If de_informa.rsSelClienteAPELIDO.RecordCount = 1 Then
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("nome")) = False Then TxtNomeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("nome"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("cgc")) = False Then TxtCGCExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("cgc"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("cidade")) = False Then TxtCidadeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("cidade"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("uf")) = False Then TxtUFExpedidor.Text = de_informa.rsSelClienteAPELIDO.Fields("uf")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("ie")) = False Then TxtInscrEstExpedidor.Text = de_informa.rsSelClienteAPELIDO.Fields("ie")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("endereco")) = False Then TxtEndExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("endereco"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("bairro")) = False Then TxtBairroEXP.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("bairro"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("cep")) = False Then TxtCEPExpedidor.Text = de_informa.rsSelClienteAPELIDO.Fields("cep")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("pabx")) = False Then TxtTelExpedidor.Text = de_informa.rsSelClienteAPELIDO.Fields("pabx")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("fax")) = False Then TxtFAXExpedidor.Text = de_informa.rsSelClienteAPELIDO.Fields("fax")
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("seguradora")) = False Then TxtSeguradoraExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("seguradora"))
            If IsNull(de_informa.rsSelClienteAPELIDO.Fields("APOLICE")) = False Then TxtApoliceExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("APOLICE"))
            SendKeys "{TAB}"
            KeyAscii = 0
            ElseIf de_informa.rsSelClienteAPELIDO.RecordCount > 1 Then
            frmEmissaoFiltraEXP.Show 1
            SendKeys "{TAB}"
            KeyAscii = 0
            Else
            de_informa.rsSelClienteAPELIDO.Close
            de_informa.SelClienteCNPJ xBusca & "%"
                If de_informa.rsSelClienteCNPJ.RecordCount = 1 Then
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("nome")) = False Then TxtNomeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("nome"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("cgc")) = False Then TxtCGCExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("cgc"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("cidade")) = False Then TxtCidadeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("cidade"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("uf")) = False Then TxtUFExpedidor.Text = de_informa.rsSelClienteCNPJ.Fields("uf")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("ie")) = False Then TxtInscrEstExpedidor.Text = de_informa.rsSelClienteCNPJ.Fields("ie")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("endereco")) = False Then TxtEndExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("endereco"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("bairro")) = False Then TxtBairroEXP.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("bairro"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("cep")) = False Then TxtCEPExpedidor.Text = de_informa.rsSelClienteCNPJ.Fields("cep")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("pabx")) = False Then TxtTelExpedidor.Text = de_informa.rsSelClienteCNPJ.Fields("pabx")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("fax")) = False Then TxtFAXExpedidor.Text = de_informa.rsSelClienteCNPJ.Fields("fax")
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("seguradora")) = False Then TxtSeguradoraExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("seguradora"))
                If IsNull(de_informa.rsSelClienteCNPJ.Fields("APOLICE")) = False Then TxtApoliceExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("APOLICE"))
                SendKeys "{TAB}"
                KeyAscii = 0
                ElseIf de_informa.rsSelClienteCNPJ.RecordCount > 1 Then
                frmEmissaoFiltraEXP.Show 1
                SendKeys "{TAB}"
                KeyAscii = 0
                Else
                de_informa.rsSelClienteCNPJ.Close
                de_informa.SelClientefantasia xBusca & "%"
                    If de_informa.rsSelClienteFANTASIA.RecordCount = 1 Then
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("nome")) = False Then TxtNomeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("nome"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("cgc")) = False Then TxtCGCExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("cgc"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("cidade")) = False Then TxtCidadeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("cidade"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("uf")) = False Then TxtUFExpedidor.Text = de_informa.rsSelClienteFANTASIA.Fields("uf")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("ie")) = False Then TxtInscrEstExpedidor.Text = de_informa.rsSelClienteFANTASIA.Fields("ie")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("endereco")) = False Then TxtEndExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("endereco"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("bairro")) = False Then TxtBairroEXP.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("bairro"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("cep")) = False Then TxtCEPExpedidor.Text = de_informa.rsSelClienteFANTASIA.Fields("cep")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("pabx")) = False Then TxtTelExpedidor.Text = de_informa.rsSelClienteFANTASIA.Fields("pabx")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("fax")) = False Then TxtFAXExpedidor.Text = de_informa.rsSelClienteFANTASIA.Fields("fax")
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("seguradora")) = False Then TxtSeguradoraExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("seguradora"))
                    If IsNull(de_informa.rsSelClienteFANTASIA.Fields("APOLICE")) = False Then TxtApoliceExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("APOLICE"))
                    SendKeys "{TAB}"
                    KeyAscii = 0
                    ElseIf de_informa.rsSelClienteFANTASIA.RecordCount > 1 Then
                    frmEmissaoFiltraEXP.Show 1
                    SendKeys "{TAB}"
                    KeyAscii = 0
                    Else
                    de_informa.rsSelClienteFANTASIA.Close
                    de_informa.SelClientenome "%" & xBusca & "%"
                        If de_informa.rsSelClienteNOME.RecordCount = 1 Then
                        If IsNull(de_informa.rsSelClienteNOME.Fields("nome")) = False Then TxtNomeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("nome"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("cgc")) = False Then TxtCGCExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("cgc"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("cidade")) = False Then TxtCidadeExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("cidade"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("uf")) = False Then TxtUFExpedidor.Text = de_informa.rsSelClienteNOME.Fields("uf")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("ie")) = False Then TxtInscrEstExpedidor.Text = de_informa.rsSelClienteNOME.Fields("ie")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("endereco")) = False Then TxtEndExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("endereco"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("bairro")) = False Then TxtBairroEXP.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("bairro"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("cep")) = False Then TxtCEPExpedidor.Text = de_informa.rsSelClienteNOME.Fields("cep")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("pabx")) = False Then TxtTelExpedidor.Text = de_informa.rsSelClienteNOME.Fields("pabx")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("fax")) = False Then TxtFAXExpedidor.Text = de_informa.rsSelClienteNOME.Fields("fax")
                        If IsNull(de_informa.rsSelClienteNOME.Fields("seguradora")) = False Then TxtSeguradoraExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("seguradora"))
                        If IsNull(de_informa.rsSelClienteNOME.Fields("APOLICE")) = False Then TxtApoliceExpedidor.Text = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("APOLICE"))
                        SendKeys "{TAB}"
                        KeyAscii = 0
                        ElseIf de_informa.rsSelClienteNOME.RecordCount > 1 Then
                        frmEmissaoFiltraEXP.Show 1
                        SendKeys "{TAB}"
                        KeyAscii = 0
                        Else
                        MsgBox "Termo não encontrado!", vbCritical, ""
                        TxtBuscaExpedidor.SetFocus
                        End If
                    End If
                End If
            End If
        End If
    End With
End If
End Sub


Private Sub TxtBuscaFilial_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtBuscaSiglaCia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub TxtBuscaSiglaDEST_KeyPress(KeyAscii As Integer)
With TxtBuscaSiglaDEST
If KeyAscii = 13 Then
    If Len(Trim(.Text)) > 0 Then
    LblAtualizarFrete.Caption = "Sim"
    If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
    If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
    de_informa.SelAeroportoSigla .Text & "%"
        If de_informa.rsSelAeroportoSigla.RecordCount > 1 Then
        frmEmissaoBuscaAeroportoDEST.Show 1
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf de_informa.rsSelAeroportoSigla.RecordCount = 1 Then
        TxtSiglaDestinatario.Text = de_informa.rsSelAeroportoSigla.Fields("sigla")
        TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("localidade")) & " - " & de_informa.rsSelAeroportoSigla.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("aeroporto")) & ")"
        TxtSiglaVIA.Text = TxtSiglaDestinatario.Text
        TxtAeroportoVIA.Text = TxtAeroportoDestinatario.Text
        
        If de_informa.rsSelRepres.State = 1 Then de_informa.rsSelRepres.Close
        de_informa.SelRepres de_informa.rsSelAeroportoSigla.Fields("localidade"), de_informa.rsSelAeroportoSigla.Fields("uf")
            If de_informa.rsSelRepres.RecordCount > 0 Then
                If MsgBox("Deseja utilizar o Representante INTEC como Destinatário?", vbYesNo + vbQuestion, "") = vbYes Then
                    Call LimpaFrame(frmEmissao, FraDestinatario.Caption)
                    If de_informa.rsSelRepres.RecordCount = 1 Then
                    TxtNomeDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("nome"))
                    TxtCGCDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("cgc"))
                    TxtCidadeDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("LOCALIDADE"))
                    TxtUFDestinatario.Text = de_informa.rsSelRepres.Fields("uf")
                    TxtInscrEstDestinatario.Text = de_informa.rsSelRepres.Fields("inscr_est")
                    TxtEndDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("endereco"))
                    TxtBairroDEST.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("bairro"))
                    TxtCEPDestinatario.Text = de_informa.rsSelRepres.Fields("cep")
                    TxtTelDestinatario.Text = de_informa.rsSelRepres.Fields("telcom")
                    TxtFAXDestinatario.Text = de_informa.rsSelRepres.Fields("fax")
                    
                    If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
                    de_informa.SelAeroportoCidade de_informa.rsSelRepres.Fields("cidaderetira")
                    TxtSiglaVIA.Text = de_informa.rsSelAeroportoCidade.Fields("sigla")
                    TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & de_informa.rsSelAeroportoCidade.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("aeroporto")) & ")"
                    
                    ElseIf de_informa.rsSelRepres.RecordCount > 1 Then
                    frmEmissaoFiltraRESPRES.Show 1
                    End If
                'TxtDescrIATA.SetFocus
                
                End If
            End If
            SendKeys "{TAB}"
            KeyAscii = 0
        Else
        If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
        de_informa.SelAeroportoCidade .Text & "%"
            If de_informa.rsSelAeroportoCidade.RecordCount > 1 Then
            frmEmissaoBuscaAeroportoDEST.Show 1
            SendKeys "{TAB}"
            KeyAscii = 0
            ElseIf de_informa.rsSelAeroportoCidade.RecordCount = 1 Then
            TxtSiglaDestinatario.Text = de_informa.rsSelAeroportoCidade.Fields("sigla")
            TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & de_informa.rsSelAeroportoCidade.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("aeroporto")) & ")"
            TxtSiglaVIA.Text = TxtSiglaDestinatario.Text
            TxtAeroportoVIA.Text = TxtAeroportoDestinatario.Text
            If de_informa.rsSelRepres.State = 1 Then de_informa.rsSelRepres.Close
            de_informa.SelRepres de_informa.rsSelAeroportoCidade.Fields("localidade"), de_informa.rsSelAeroportoCidade.Fields("uf")
                If de_informa.rsSelRepres.RecordCount > 0 Then
                    If MsgBox("Deseja utilizar o Representante INTEC como Destinatário?", vbYesNo + vbQuestion, "") = vbYes Then
                    Call LimpaFrame(frmEmissao, FraDestinatario.Caption)
                        If de_informa.rsSelRepres.RecordCount = 1 Then
                        TxtNomeDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("nome"))
                        TxtCGCDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("cgc"))
                        TxtCidadeDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("LOCALIDADE"))
                        TxtUFDestinatario.Text = de_informa.rsSelRepres.Fields("uf")
                        TxtInscrEstDestinatario.Text = de_informa.rsSelRepres.Fields("inscr_est")
                        TxtEndDestinatario.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("endereco"))
                        TxtBairroDEST.Text = PriMaiuscula(de_informa.rsSelRepres.Fields("bairro"))
                        TxtCEPDestinatario.Text = de_informa.rsSelRepres.Fields("cep")
                        TxtTelDestinatario.Text = de_informa.rsSelRepres.Fields("telcom")
                        TxtFAXDestinatario.Text = de_informa.rsSelRepres.Fields("fax")
                        
                        If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
                        de_informa.SelAeroportoCidade de_informa.rsSelRepres.Fields("cidaderetira")
                        TxtSiglaVIA.Text = de_informa.rsSelAeroportoCidade.Fields("sigla")
                        TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & de_informa.rsSelAeroportoCidade.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("aeroporto")) & ")"
                        
                        ElseIf de_informa.rsSelRepres.RecordCount > 1 Then
                        frmEmissaoFiltraRESPRES.Show 1
                        End If
                    'TxtDescrIATA.SetFocus
                    End If
                End If
            SendKeys "{TAB}"
            KeyAscii = 0
            Else
            MsgBox "Termo não encontrado!", vbCritical, ""
            End If
        End If
    Else
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End If
End With

End Sub

Private Sub TxtBuscaSiglaExp_KeyPress(KeyAscii As Integer)

With TxtBuscaSiglaExp
If KeyAscii = 13 Then
    If Len(Trim(.Text)) > 0 Then
    LblAtualizarFrete.Caption = "Sim"
    If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
    If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
    
    de_informa.SelAeroportoSigla .Text & "%"
        If de_informa.rsSelAeroportoSigla.RecordCount > 1 Then
        frmEmissaoBuscaAeroportoEXP.Show 1
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf de_informa.rsSelAeroportoSigla.RecordCount = 1 Then
        TxtSiglaExpedidor.Text = de_informa.rsSelAeroportoSigla.Fields("sigla")
        TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("localidade")) & " - " & de_informa.rsSelAeroportoSigla.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("aeroporto")) & ")"
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
        de_informa.SelAeroportoCidade .Text & "%"
            If de_informa.rsSelAeroportoCidade.RecordCount > 1 Then
            frmEmissaoBuscaAeroportoEXP.Show 1
            SendKeys "{TAB}"
            KeyAscii = 0
            ElseIf de_informa.rsSelAeroportoCidade.RecordCount = 1 Then
            TxtSiglaExpedidor.Text = de_informa.rsSelAeroportoCidade.Fields("sigla")
            TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & de_informa.rsSelAeroportoCidade.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("aeroporto")) & ")"
            SendKeys "{TAB}"
            KeyAscii = 0
            Else
            MsgBox "Termo não encontrado!", vbCritical, ""
            End If
        End If
    Else
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End If
End With
End Sub

Private Sub TxtBuscaSiglaVIA_KeyPress(KeyAscii As Integer)
With TxtBuscaSiglaVIA
If KeyAscii = 13 Then
    If Len(Trim(.Text)) > 0 Then
    LblAtualizarFrete.Caption = "Sim"
    If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
    If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
    de_informa.SelAeroportoSigla .Text & "%"
        If de_informa.rsSelAeroportoSigla.RecordCount > 1 Then
        frmEmissaoBuscaAeroportoDEST.Show 1
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf de_informa.rsSelAeroportoSigla.RecordCount = 1 Then
        TxtSiglaVIA.Text = de_informa.rsSelAeroportoSigla.Fields("sigla")
        TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("localidade")) & " - " & de_informa.rsSelAeroportoSigla.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("aeroporto")) & ")"
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
        de_informa.SelAeroportoCidade .Text & "%"
            If de_informa.rsSelAeroportoCidade.RecordCount > 1 Then
            frmEmissaoBuscaAeroportoDEST.Show 1
            SendKeys "{TAB}"
            KeyAscii = 0
            ElseIf de_informa.rsSelAeroportoCidade.RecordCount = 1 Then
            TxtSiglaVIA.Text = de_informa.rsSelAeroportoCidade.Fields("sigla")
            TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & de_informa.rsSelAeroportoCidade.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("aeroporto")) & ")"
            SendKeys "{TAB}"
            KeyAscii = 0
            Else
            MsgBox "Termo não encontrado!", vbCritical, ""
            End If
        End If
    Else
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End If
End With

End Sub


Private Sub TxtCodIATA_KeyPress(KeyAscii As Integer)
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

Private Sub TxtLocalRetirada_Change()
If Len(TxtLocalRetirada.Text) > 0 Then
TxtLocalRetirada.Text = UCase(TxtLocalRetirada.Text)
TxtLocalRetirada.SelStart = Len(TxtLocalRetirada.Text)
End If
End Sub

Private Sub TxtLocalRetirada_KeyPress(KeyAscii As Integer)
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


Private Sub ImpressaoPorPrintVARIG()

xAUX = 55
xIMPNomeEXP = xIMPNomeEXP & String(xAUX - Len(xIMPNomeEXP), " ")
xAUX = 55
xIMPCGCEXP = xIMPCGCEXP & String(xAUX - Len(xIMPCGCEXP), " ")
xAUX = 55
xIMPInscEstEXP = xIMPInscEstEXP & String(xAUX - Len(xIMPInscEstEXP), " ")
xAUX = 55
xIMPEndEXP = xIMPEndEXP & String(xAUX - Len(xIMPEndEXP), " ")
xAUX = 25
xIMPBairroEXP = xIMPBairroEXP & String(xAUX - Len(xIMPBairroEXP), " ")
xAUX = 25
xIMPCidadeEXP = xIMPCidadeEXP & String(xAUX - Len(xIMPCidadeEXP), " ")
xAUX = 20
xIMPCepEXP = xIMPCepEXP & String(xAUX - Len(xIMPCepEXP), " ")
xAUX = 5
xIMPUFEXP = xIMPUFEXP & String(xAUX - Len(xIMPUFEXP), " ")
xAUX = 20
xIMPTelEXP = xIMPTelEXP & String(xAUX - Len(xIMPTelEXP), " ")
xAUX = 55
xIMPFAXEXP = xIMPFAXEXP & String(xAUX - Len(xIMPFAXEXP), " ")
xAUX = 55
xIMPNomeDEST = xIMPNomeDEST & String(xAUX - Len(xIMPNomeDEST), " ")
xAUX = 55
xIMPCGCDEST = xIMPCGCDEST & String(xAUX - Len(xIMPCGCDEST), " ")
xAUX = 55
xIMPInscEstDEST = xIMPInscEstDEST & String(xAUX - Len(xIMPInscEstDEST), " ")
xAUX = 55
xIMPEndDEST = xIMPEndDEST & String(xAUX - Len(xIMPEndDEST), " ")
xAUX = 25
xIMPBairroDEST = xIMPBairroDEST & String(xAUX - Len(xIMPBairroDEST), " ")
xAUX = 25
xIMPCidadeDEST = xIMPCidadeDEST & String(xAUX - Len(xIMPCidadeDEST), " ")
xAUX = 20
xIMPCepDEST = xIMPCepDEST & String(xAUX - Len(xIMPCepDEST), " ")
xAUX = 5
xIMPUFDEST = xIMPUFDEST & String(xAUX - Len(xIMPUFDEST), " ")
xAUX = 20
xIMPTelDEST = xIMPTelDEST & String(xAUX - Len(xIMPTelDEST), " ")
xAUX = 55
xIMPFAXDEST = xIMPFAXDEST & String(xAUX - Len(xIMPFAXDEST), " ")
xAUX = 10
xIMPOrigem = String(Int((xAUX - Len(xIMPOrigem)) / 2), " ") & xIMPOrigem & String((xAUX - Len(xIMPOrigem)) - Int((xAUX - Len(xIMPOrigem)) / 2), " ")
xAUX = 11
xIMPVia = String(Int((xAUX - Len(xIMPVia)) / 2), " ") & xIMPVia & String((xAUX - Len(xIMPVia)) - Int((xAUX - Len(xIMPVia)) / 2), " ")
xAUX = 25
xIMPCidadeDESTINO = String(Int((xAUX - Len(xIMPCidadeDESTINO)) / 2), " ") & xIMPCidadeDESTINO & String((xAUX - Len(xIMPCidadeDESTINO)) - Int((xAUX - Len(xIMPCidadeDESTINO)) / 2), " ")
xAUX = 13
xIMPSIGLA = String(Int((xAUX - Len(xIMPSIGLA)) / 2), " ") & xIMPSIGLA & String((xAUX - Len(xIMPSIGLA)) - Int((xAUX - Len(xIMPSIGLA)) / 2), " ")
xAUX = 20
xIMPReqTranspMinuta = String(Int((xAUX - Len(xIMPReqTranspMinuta)) / 2), " ") & xIMPReqTranspMinuta & String((xAUX - Len(xIMPReqTranspMinuta)) - Int((xAUX - Len(xIMPReqTranspMinuta)) / 2), " ")
xAUX = 37
xIMPNumControle = String(Int((xAUX - Len(xIMPNumControle)) / 2), " ") & xIMPNumControle & String((xAUX - Len(xIMPNumControle)) - Int((xAUX - Len(xIMPNumControle)) / 2), " ")
xAUX = 29
xIMPInscrEstCiaAerea = String(Int((xAUX - Len(xIMPInscrEstCiaAerea)) / 2), " ") & xIMPInscrEstCiaAerea & String((xAUX - Len(xIMPInscrEstCiaAerea)) - Int((xAUX - Len(xIMPInscrEstCiaAerea)) / 2), " ")
xAUX = 29
xIMPCNPJCiaAerea = String(Int((xAUX - Len(xIMPCNPJCiaAerea)) / 2), " ") & xxIMPCNPJCiaAerea & String((xAUX - Len(xIMPCNPJCiaAerea)) - Int((xAUX - Len(xIMPCNPJCiaAerea)) / 2), " ")
xAUX = 29
xIMPVlDecTRANSP = String(xAUX - Len(xIMPVlDecTRANSP), " ") & xIMPVlDecTRANSP
xAUX = 29
xIMPVlDecSUFRAMA = String(xAUX - Len(xIMPVlDecSUFRAMA), " ") & xIMPVlDecSUFRAMA
xAUX = 62
xIMPDescrEmbalagem = String(Int((xAUX - Len(xIMPDescrEmbalagem)) / 2), " ") & xIMPDescrEmbalagem & String((xAUX - Len(xIMPDescrEmbalagem)) - Int((xAUX - Len(xIMPDescrEmbalagem)) / 2), " ")
xAUX = 7
xIMPQteVol = String(Int((xAUX - Len(xIMPQteVol)) / 2), " ") & xIMPQteVol & String((xAUX - Len(xIMPQteVol)) - Int((xAUX - Len(xIMPQteVol)) / 2), " ")
xAUX = 11
xIMPPesoReal = String(Int((xAUX - Len(xIMPPesoReal)) / 2), " ") & xIMPPesoReal & String((xAUX - Len(xIMPPesoReal)) - Int((xAUX - Len(xIMPPesoReal)) / 2), " ")
xAUX = 11
xIMPPesoTax = String(Int((xAUX - Len(xIMPPesoTax)) / 2), " ") & xIMPPesoTax & String((xAUX - Len(xIMPPesoTax)) - Int((xAUX - Len(xIMPPesoTax)) / 2), " ")
xAUX = 13
xIMPTrecho = String(Int((xAUX - Len(xIMPTrecho)) / 2), " ") & xIMPTrecho & String((xAUX - Len(xIMPTrecho)) - Int((xAUX - Len(xIMPTrecho)) / 2), " ")
xAUX = 4
xIMPCl = String(Int((xAUX - Len(xIMPCl)) / 2), " ") & xIMPCl & String((xAUX - Len(xIMPCl)) - Int((xAUX - Len(xIMPCl)) / 2), " ")
xAUX = 7
xIMPCodigo = String(Int((xAUX - Len(xIMPCodigo)) / 2), " ") & xIMPCodigo & String((xAUX - Len(xIMPCodigo)) - Int((xAUX - Len(xIMPCodigo)) / 2), " ")
xAUX = 12
xIMPKilo = String(Int((xAUX - Len(xIMPKilo)) / 2), " ") & xIMPKilo & String((xAUX - Len(xIMPKilo)) - Int((xAUX - Len(xIMPKilo)) / 2), " ")
xAUX = 19
xIMPFreteNacEscopo = String(Int((xAUX - Len(xIMPFreteNacEscopo)) / 2), " ") & xIMPFreteNacEscopo & String((xAUX - Len(xIMPFreteNacEscopo)) - Int((xAUX - Len(xIMPFreteNacEscopo)) / 2), " ")
xAUX = 33
xIMPNatureza = String(Int((xAUX - Len(xIMPNatureza)) / 2), " ") & xIMPNatureza & String((xAUX - Len(xIMPNatureza)) - Int((xAUX - Len(xIMPNatureza)) / 2), " ")
xAUX = 29
xIMPTxDescrDevAg = String(Int((xAUX - Len(xIMPTxDescrDevAg)) / 2), " ") & xIMPTxDescrDevAg & String((xAUX - Len(xIMPTxDescrDevAg)) - Int((xAUX - Len(xIMPTxDescrDevAg)) / 2), " ")
xAUX = 29
xIMPTxDescrDevTransp = String(Int((xAUX - Len(xIMPTxDescrDevTransp)) / 2), " ") & xIMPTxDescrDevTransp & String((xAUX - Len(xIMPTxDescrDevTransp)) - Int((xAUX - Len(xIMPTxDescrDevTransp)) / 2), " ")
xAUX = 21
xIMPFreteNacional = String(xAUX - Len(xIMPFreteNacional), " ") & xIMPFreteNacional
xAUX = 21
xIMPFreteRegional = String(xAUX - Len(xIMPFreteRegional), " ") & xIMPFreteRegional
xAUX = 21
xIMPAdValorem = String(xAUX - Len(xIMPAdValorem), " ") & xIMPAdValorem
xAUX = 6
xIMPTipoADVAL = String(Int((xAUX - Len(xIMPTipoADVAL)) / 2), " ") & xIMPTipoADVAL & String((xAUX - Len(xIMPTipoADVAL)) - Int((xAUX - Len(xIMPTipoADVAL)) / 2), " ")
xAUX = 21
xIMPTxTerrOrig = String(xAUX - Len(xIMPTxTerrOrig), " ") & xIMPTxTerrOrig
xAUX = 21
xIMPTxTerrDest = String(xAUX - Len(xIMPTxTerrDest), " ") & xIMPTxTerrDest
xAUX = 21
xIMPTxRedesp = String(xAUX - Len(xIMPTxRedesp), " ") & xIMPTxRedesp
xAUX = 21
xIMPTxAgente = String(xAUX - Len(xIMPTxAgente), " ") & xIMPTxAgente
xAUX = 21
xIMPTxDevTransp = String(xAUX - Len(xIMPTxDevTransp), " ") & xIMPTxDevTransp
xAUX = 14
xIMPDescrTxOutros1 = String(Int((xAUX - Len(xIMPDescrTxOutros1)) / 2), " ") & xIMPDescrTxOutros1 & String((xAUX - Len(xIMPDescrTxOutros1)) - Int((xAUX - Len(xIMPDescrTxOutros1)) / 2), " ")
xAUX = 21
xIMPTxOutros1 = String(xAUX - Len(xIMPTxOutros1), " ") & xIMPTxOutros1
xAUX = 14
xIMPDescrTxOutros2 = String(Int((xAUX - Len(xIMPDescrTxOutros2)) / 2), " ") & xIMPDescrTxOutros2 & String((xAUX - Len(xIMPDescrTxOutros2)) - Int((xAUX - Len(xIMPDescrTxOutros2)) / 2), " ")
xAUX = 21
xIMPTxOutros2 = String(xAUX - Len(xIMPTxOutros2), " ") & xIMPTxOutros2
xAUX = 21
xIMPFreteTotal = String(xAUX - Len(xIMPFreteTotal), " ") & xIMPFreteTotal
'xaux =
'xIMPStrRetira = xIMPStrRetira & String(xaux - Len(xIMPStrRetira), " ")
'xaux =
'xIMPStrLocalRetira = xIMPStrLocalRetira & String(xaux - Len(xIMPStrLocalRetira), " ")
xAUX = 60
xIMPHorarioAt = xIMPHorarioAt & String(xAUX - Len(xIMPHorarioAt), " ")
xAUX = 60
xIMPStrTelefone = xIMPStrTelefone & String(xAUX - Len(xIMPStrTelefone), " ")
xAUX = 23
xIMPStrTotalServ = String(xAUX - Len(xIMPStrTotalServ), " ") & xIMPStrTotalServ
xAUX = 23
xIMPStrBaseCalculo = String(xAUX - Len(xIMPStrBaseCalculo), " ") & xIMPStrBaseCalculo
xAUX = 6
xIMPStrAliquota = String(xAUX - Len(Trim(xIMPStrAliquota)), " ") & xIMPStrAliquota
xAUX = 23
xIMPStrICMS = String(xAUX - Len(xIMPStrICMS), " ") & xIMPStrICMS
xAUX = 27
xIMPAgenteEmissor = String(Int((xAUX - Len(xIMPAgenteEmissor)) / 2), " ") & xIMPAgenteEmissor & String((xAUX - Len(xIMPAgenteEmissor)) - Int((xAUX - Len(xIMPAgenteEmissor)) / 2), " ")
xAUX = 27
xIMPCodIATA = String(Int((xAUX - Len(xIMPCodIATA)) / 2), " ") & xIMPCodIATA & String((xAUX - Len(xIMPCodIATA)) - Int((xAUX - Len(xIMPCodIATA)) / 2), " ")
xAUX = 27
xIMPDtEmissao = String(Int((xAUX - Len(xIMPDtEmissao)) / 2), " ") & xIMPDtEmissao & String((xAUX - Len(xIMPDtEmissao)) - Int((xAUX - Len(xIMPDtEmissao)) / 2), " ")
xAUX = 27
xIMPHoraEmissao = String(Int((xAUX - Len(xIMPHoraEmissao)) / 2), " ") & xIMPHoraEmissao & String((xAUX - Len(xIMPHoraEmissao)) - Int((xAUX - Len(xIMPHoraEmissao)) / 2), " ")
xAUX = 25
xIMPNaturezaOp = xIMPNaturezaOp & String(xAUX - Len(xIMPNaturezaOp), " ")
xAUX = 8
xIMPCFOP = xIMPCFOP & String(xAUX - Len(xIMPCFOP), " ")
xAUX = 59
xIMPEmissor = xIMPEmissor & String(xAUX - Len(xIMPEmissor), " ")
xAUX = 8
xIMPLocalidade = xIMPLocalidade & String(xAUX - Len(xIMPLocalidade), " ")
xAUX = 17
xIMPMatricula = xIMPMatricula & String(xAUX - Len(xIMPMatricula), " ")


xAUX = 8
If OptAPagar.Value = True Then xAUX = 50

Open SETIMPImpressoraPadrao For Output As #1
DoEvents
Print #1, Chr(15)
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, "                                                                         " & Mid(StringNF, 1, 60)
Print #1, "              " & xIMPNomeEXP & "  " & Mid(StringNF, 61, 60)
Print #1, "              " & xIMPCGCEXP & "  " & Mid(StringNF, 121, 60)
Print #1, "                " & xIMPInscEstEXP & "  " & Mid(StringNF, 181, 60)
Print #1, "            " & xIMPEndEXP & "  " & Mid(StringNF, 241, 60)
Print #1, "              " & xIMPBairroEXP & "     " & xIMPCidadeEXP & "  " & Mid(StringNF, 301, 60)
Print #1, "              " & xIMPCepEXP & "    " & xIMPUFEXP & "      " & xIMPTelEXP & "  " & Mid(StringNF, 361, 60)
Print #1, "              " & xfaxexp & "  " & Mid(StringNF, 361, 60)
Print #1, "                                                                         " & Mid(StringNF, 421, 60)
Print #1, "                                                                         " & Mid(StringNF, 481, 60)
Print #1, "                                                                         " & Mid(StringNF, 541, 60)

If OptPago.Value = True Then
Print #1, "              " & xIMPNomeDEST & "    F R E T E   P A G O"
Else
Print #1, "              " & xIMPNomeDEST & "    F R E T E   A   P A G A R"
End If

Print #1, "              " & xIMPCGCDEST & "                                                                     "
Print #1, "                " & xIMPInscEstDEST & "  " & xIMPReqTranspMinuta & "     " & xIMPNumControle
Print #1, "            " & xIMPEndDEST
Print #1, "              " & xIMPBairroDEST & "     " & xIMPCidadeDEST & "  " & xIMPInscrEstCiaAerea & "    " & xIMPCNPJCiaAerea
Print #1, "              " & xIMPCepDEST & "    " & xIMPUFDEST & "      " & xIMPTelDEST
Print #1, "              " & xIMPFAXDEST & "  " & xIMPVlDecTRANSP & "    " & xIMPVlDecSUFRAMA
Print #1, ""
Print #1, "       " & xIMPOrigem & "  " & xIMPVia & " " & xIMPCidadeDESTINO & "  " & xIMPSIGLA & "  " & xIMPDescrEmbalagem
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, "       " & xIMPQteVol & " " & xIMPPesoReal & "  " & xIMPPesoTax & " " & xIMPTrecho & " " & xIMPCl & " " & xIMPCodigo & "  " & xIMPKilo & "  " & xIMPFreteNacEscopo & " " & xIMPNatureza
Print #1, ""
Print #1, ""
Print #1, "                                                                          " & xIMPTxDescrDevAg & "   " & xIMPTxDescrDevTransp

    If OptPago.Value = True Then
    Print #1, ""
    Print #1, "        " & xIMPFreteNacional & "                                             " & Mid(xIMPStrObservacao, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 61, 60)
    Print #1, "        " & xIMPFreteRegional & "                                             " & Mid(xIMPStrObservacao, 121, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 181, 60)
    Print #1, "        " & xIMPAdValorem & "             " & xIMPTipoADVAL & "                          " & Mid(xIMPStrObservacao, 241, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 301, 60)
    Print #1, "        " & xIMPTxTerrOrig & "                                             " & Mid(xIMPStrObservacao, 361, 60)
    Print #1, "                                                                                                                                            "
        If OptRetiraSim.Value = True Then
        Print #1, "        " & xIMPTxTerrDest & "                                                                     XXX"
        Else
        Print #1, "        " & xIMPTxTerrDest & "                                                                                XXX"
        End If
    Print #1, ""
    Print #1, "        " & xIMPTxRedesp & "                                             " & Mid(xIMPStrLocalRetira, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrLocalRetira, 61, 60)
    Print #1, "        " & xIMPTxAgente
    Print #1, ""
    Print #1, "        " & xIMPTxDevTransp
    Print #1, ""
    Print #1, "        " & xIMPTxOutros1 & "   " & xIMPDescrTxOutros1
    Print #1, "                                                                                                              " & xIMPStrTotalServ
    Print #1, "        " & xIMPTxOutros2 & "   " & xIMPDescrTxOutros2 & "                                                                " & xIMPStrBaseCalculo
    Print #1, ""
    Print #1, "        " & xIMPFreteTotal & "                                              " & xIMPStrAliquota & "                    " & xIMPStrICMS
    
    Else
    
    Print #1, ""
    Print #1, "                                                  " & xIMPFreteNacional & "   " & Mid(xIMPStrObservacao, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 61, 60)
    Print #1, "                                                  " & xIMPFreteRegional & "   " & Mid(xIMPStrObservacao, 121, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 181, 60)
    Print #1, "                                          " & xIMPTipoADVAL & "  " & xIMPAdValorem & "   " & Mid(xIMPStrObservacao, 241, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 301, 60)
    Print #1, "                                                  " & xIMPTxTerrOrig & "   " & Mid(xIMPStrObservacao, 361, 60)
    Print #1, "                                                                                                                                            "
        If OptRetiraSim.Value = True Then
        Print #1, "                                                  " & xIMPTxTerrDest & "                           XXX"
        Else
        Print #1, "                                                  " & xIMPTxTerrDest & "                                      XXX"
        End If
    Print #1, ""
    Print #1, "                                                  " & xIMPTxRedesp & "    " & Mid(xIMPStrLocalRetira, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrLocalRetira, 61, 60)
    Print #1, "                                                  " & xIMPTxAgente
    Print #1, ""
    Print #1, "                                                  " & xIMPTxDevTransp
    Print #1, ""
    Print #1, "                                " & xIMPDescrTxOutros1 & "    " & xIMPDescrTxOutros2
    Print #1, "                                                                                                              " & xIMPStrTotalServ
    Print #1, "                                " & xIMPDescrTxOutros2 & "   " & xIMPTxOutros2 & "                                        " & xIMPStrBaseCalculo
    Print #1, ""
    Print #1, "                                                  " & xIMPFreteTotal & "               " & xIMPStrAliquota & "                    " & xIMPStrICMS
    End If

Print #1, ""
Print #1, "                                                                          " & xIMPAgenteEmissor & "     " & xIMPCodIATA
Print #1, ""
Print #1, "                                                                          " & xIMPDtEmissao & "     " & xIMPHoraEmissao
Print #1, ""
Print #1, "        " & xIMPNaturezaOp & "  " & xIMPCFOP & "   " & xIMPEmissor & "   " & xIMPLocalidade & "  " & xIMPMatricula
DoEvents
Close #1

End Sub


Private Sub ImpressaoPorObjectPrinter()

'de_informa.cn_informa.CommitTrans
End Sub


Sub EnviaIMP(xNomeFont As String, xSizeFont As Double, xCurrentX As Double, xCurrentY As Double, xIMP As String)
Printer.ScaleMode = vbMillimeters
Printer.Font.Name = xNomeFont
Printer.Font.Size = xSizeFont
Printer.CurrentX = xCurrentX
Printer.CurrentY = xCurrentY
Printer.Print xIMP
End Sub


Sub ImprimirporPRINTER()
Dim Fonte As String
Dim xREP As String
Dim FontSize As Double
Dim MargemTOP As Double
Dim Margem2TOP As Double
Dim MargemLeft As Double
Dim AuxTOP As Double
Dim AuxLEFT As Double
Dim Linha As Double
Dim AuxMODAL As Long



Open SETIMPImpressoraPadrao For Output As #1
Print #1, Chr(15)
DoEvents
Close #1


AuxTOP = 23.8
AuxLEFT = 25

xfatortipoletra = 54

Linha = 4.25
Fonte = "Courier New"

MargemTOP = 15
MargemTOP = MargemTOP + AuxTOP
FontSize = 10

'DADOS DE EXPEDIDOR
MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPNomeEXP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP & " ALT " & Trim(Str(AuxTOP)))

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCGCEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPInscEstEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPEndEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPBairroEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 93
xREP = xIMPCidadeEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCepEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 80
xREP = xIMPUFEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 110
xREP = xIMPTelEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPFAXEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (3 * Linha)


'DADOS DE DESTINATARIO
MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPNomeDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCGCDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPInscEstDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPEndDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPBairroDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 93
xREP = xIMPCidadeDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCepDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 80
xREP = xIMPUFDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 110
xREP = xIMPTelDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPFAXDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'Dados abaixo do Dest

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 5
xREP = xIMPOrigem
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 20
xREP = xIMPVia
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 65
xREP = xIMPCidadeDESTINO
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 135
xREP = xIMPSIGLA
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'Dados da linha Escopo
MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 3
xREP = xIMPQteVol
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 25
xREP = xIMPPesoReal
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 60
xREP = xIMPPesoTax
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 90
xREP = xIMPTrecho
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 123
xREP = xIMPCl
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'LINHA ESCOPO
MargemTOP = MargemTOP
MargemLeft = 138
xREP = xIMPCodigo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 168
xREP = xIMPKilo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 220
xREP = xIMPFreteNacEscopo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 248
xREP = xIMPNatureza
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

If OptPago.Value = True Then
AuxMODAL = 0
Else
AuxMODAL = 109
End If

MargemTOP = MargemTOP + (4 * Linha)

'Taxas
MargemTOP = MargemTOP + Linha
MargemLeft = 15
xREP = xIMPFreteNacional
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'MargemTOP = MargemTOP + (2 * Linha)
'MargemLeft = 15
'xREP = xIMPFreteRegional
'MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
'Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPAdValorem
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 90
xREP = xIMPTipoADVAL
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxTerrOrig
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxTerrDest
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (Linha)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxRedesp
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxAgente
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxDevTransp
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxOutros1
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 65
xREP = xIMPDescrTxOutros1
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxOutros2
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 65
xREP = xIMPDescrTxOutros2
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPFreteTotal
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = 10.8
MargemTOP = MargemTOP + AuxTOP

'Dados de NFs
MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF01
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF02
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF03
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF04
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF05
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF06
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF07
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF08
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF09
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF10
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF11
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF12
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPReqTranspMinuta
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 240
xREP = xIMPNumControle
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPInscrEstCiaAerea
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 245
xREP = xIMPCNPJCiaAerea
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPVlDecTRANSP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, 8, MargemLeft, MargemTOP, xREP)

MargemLeft = 250
xREP = xIMPVlDecSUFRAMA
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, 8, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPDescrEmbalagem
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 165
xREP = xIMPTxDescrDevAg
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 245
xREP = xIMPTxDescrDevTransp
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (5 * Linha)
MargemLeft = 170
xREP = xIMPStrObservacao01
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao02
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao03
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao04
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsSeguro)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsICMS)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsPerecivel)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

Margem2TOP = (MargemTOP + (3 * Linha)) - 2
MargemLeft = 227
xREP = xIMPStrRetiraSIM
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, Margem2TOP, xREP)

Margem2TOP = (MargemTOP + (3 * Linha)) - 2
MargemLeft = 255
xREP = xIMPStrRetiraNAO
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, Margem2TOP, xREP)

MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 170
xREP = xIMPStrLocalRetira
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (8 * Linha)
MargemLeft = 280
xREP = xIMPStrTotalServ
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (1 * Linha)
MargemLeft = 280
xREP = xIMPStrBaseCalculo
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 176
xREP = xIMPStrAliquota
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 280
xREP = xIMPStrICMS
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 170
xREP = xIMPAgenteEmissor
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 258
xREP = xIMPCodIATA
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 170
xREP = xIMPDtEmissao
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 258
xREP = xIMPHoraEmissao
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 6
xREP = xIMPNaturezaOp
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 75
xREP = xIMPCFOP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 100
xREP = xIMPEmissor
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 260
xREP = xIMPLocalidade
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)
Printer.EndDoc

End Sub
Sub ImprimirporPRINTERVarig()
Dim Fonte As String
Dim xREP As String
Dim FontSize As Double
Dim MargemTOP As Double
Dim Margem2TOP As Double
Dim MargemLeft As Double
Dim AuxTOP As Double
Dim AuxLEFT As Double
Dim Linha As Double
Dim AuxMODAL As Long



Open SETIMPImpressoraPadrao For Output As #1
Print #1, Chr(15)
DoEvents
Close #1


AuxTOP = 23.8
AuxLEFT = 25

xfatortipoletra = 54

Linha = 4.25
Fonte = "Courier New"

MargemTOP = 15
MargemTOP = MargemTOP + AuxTOP
FontSize = 10

'DADOS DE EXPEDIDOR
MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPNomeEXP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP & " ALT " & Trim(Str(AuxTOP)))

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCGCEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPInscEstEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPEndEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPBairroEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 93
xREP = xIMPCidadeEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCepEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 80
xREP = xIMPUFEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 110
xREP = xIMPTelEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPFAXEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (3 * Linha)


'DADOS DE DESTINATARIO
MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPNomeDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCGCDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPInscEstDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPEndDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPBairroDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 93
xREP = xIMPCidadeDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCepDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 80
xREP = xIMPUFDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 110
xREP = xIMPTelDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPFAXDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'Dados abaixo do Dest

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 5
xREP = xIMPOrigem
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 20
xREP = xIMPVia
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 65
xREP = xIMPCidadeDESTINO
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 135
xREP = xIMPSIGLA
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'Dados da linha Escopo
MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 3
xREP = xIMPQteVol
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 25
xREP = xIMPPesoReal
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 60
xREP = xIMPPesoTax
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 90
xREP = xIMPTrecho
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 123
xREP = xIMPCl
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'LINHA ESCOPO
MargemTOP = MargemTOP
MargemLeft = 138
xREP = xIMPCodigo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 168
xREP = xIMPKilo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 220
xREP = xIMPFreteNacEscopo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 248
xREP = xIMPNatureza
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

If OptPago.Value = True Then
AuxMODAL = 0
Else
AuxMODAL = 109
End If

MargemTOP = MargemTOP + (4 * Linha)

'Taxas
MargemTOP = MargemTOP + Linha
MargemLeft = 15
xREP = xIMPFreteNacional
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPFreteRegional
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPAdValorem
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 90
xREP = xIMPTipoADVAL
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxTerrOrig
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxTerrDest
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (Linha)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxRedesp
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxAgente
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxDevTransp
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxOutros1
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 65
xREP = xIMPDescrTxOutros1
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxOutros2
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 65
xREP = xIMPDescrTxOutros2
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPFreteTotal
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = 10.8
MargemTOP = MargemTOP + AuxTOP

'Dados de NFs
MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF01
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF02
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF03
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF04
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF05
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF06
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF07
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF08
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF09
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF10
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF11
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF12
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPReqTranspMinuta
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 240
xREP = xIMPNumControle
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPInscrEstCiaAerea
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 245
xREP = xIMPCNPJCiaAerea
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPVlDecTRANSP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, 8, MargemLeft, MargemTOP, xREP)

MargemLeft = 250
xREP = xIMPVlDecSUFRAMA
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, 8, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPDescrEmbalagem
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 165
xREP = xIMPTxDescrDevAg
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 245
xREP = xIMPTxDescrDevTransp
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (5 * Linha)
MargemLeft = 170
xREP = xIMPStrObservacao01
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao02
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao03
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao04
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsSeguro)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsICMS)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsPerecivel)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

Margem2TOP = (MargemTOP + (3 * Linha)) - 2
MargemLeft = 227
xREP = xIMPStrRetiraSIM
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, Margem2TOP, xREP)

Margem2TOP = (MargemTOP + (3 * Linha)) - 2
MargemLeft = 255
xREP = xIMPStrRetiraNAO
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, Margem2TOP, xREP)

MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 170
xREP = xIMPStrLocalRetira
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (8 * Linha)
MargemLeft = 280
xREP = xIMPStrTotalServ
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (1 * Linha)
MargemLeft = 280
xREP = xIMPStrBaseCalculo
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 176
xREP = xIMPStrAliquota
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 280
xREP = xIMPStrICMS
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 170
xREP = xIMPAgenteEmissor
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 258
xREP = xIMPCodIATA
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 170
xREP = xIMPDtEmissao
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 258
xREP = xIMPHoraEmissao
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 6
xREP = xIMPNaturezaOp
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 75
xREP = xIMPCFOP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 100
xREP = xIMPEmissor
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 260
xREP = xIMPLocalidade
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)
Printer.EndDoc

End Sub

Sub ImprimirporPRINTERP8()
Dim Fonte As String
Dim xREP As String
Dim FontSize As Double
Dim MargemTOP As Double
Dim Margem2TOP As Double
Dim MargemLeft As Double
Dim AuxTOP As Double
Dim AuxLEFT As Double
Dim Linha As Double
Dim AuxMODAL As Long



Open SETIMPImpressoraPadrao For Output As #1
Print #1, Chr(15)
DoEvents
Close #1


AuxTOP = 23.8
AuxLEFT = 25

xfatortipoletra = 54

Linha = 4.25
Fonte = "Courier New"

MargemTOP = 15
MargemTOP = MargemTOP + AuxTOP
FontSize = 10

'DADOS DE EXPEDIDOR
MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPNomeEXP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP & " ALT " & Trim(Str(AuxTOP)))

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCGCEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPInscEstEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPEndEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPBairroEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 93
xREP = xIMPCidadeEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCepEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 80
xREP = xIMPUFEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 110
xREP = xIMPTelEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPFAXEXP
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (3 * Linha)


'DADOS DE DESTINATARIO
MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPNomeDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCGCDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPInscEstDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPEndDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPBairroDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 93
xREP = xIMPCidadeDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPCepDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 80
xREP = xIMPUFDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 110
xREP = xIMPTelDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 18
xREP = xIMPFAXDEST
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'Dados abaixo do Dest

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 5
xREP = xIMPOrigem
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 20
xREP = xIMPVia
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 65
xREP = xIMPCidadeDESTINO
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 135
xREP = xIMPSIGLA
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'Dados da linha Escopo
MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 3
xREP = xIMPQteVol
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 25
xREP = xIMPPesoReal
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 60
xREP = xIMPPesoTax
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 90
xREP = xIMPTrecho
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 123
xREP = xIMPCl
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'LINHA ESCOPO
MargemTOP = MargemTOP
MargemLeft = 138
xREP = xIMPCodigo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 168
xREP = xIMPKilo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 220
xREP = xIMPFreteNacEscopo
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP
MargemLeft = 248
xREP = xIMPNatureza
MargemLeft = MargemLeft + AuxLEFT

Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

If OptPago.Value = True Then
AuxMODAL = 0
Else
AuxMODAL = 109
End If

MargemTOP = MargemTOP + (4 * Linha)

'Taxas
MargemTOP = MargemTOP + Linha
MargemLeft = 15
xREP = xIMPFreteNacional
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPFreteRegional
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPAdValorem
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 90
xREP = xIMPTipoADVAL
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxTerrOrig
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxTerrDest
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (Linha)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxRedesp
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxAgente
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxDevTransp
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPTxOutros1
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 65
xREP = xIMPDescrTxOutros1
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

'MargemTOP = MargemTOP + (2 * Linha)
'MargemLeft = 15
'xREP = xIMPTxOutros2
'MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
'Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)
'
'MargemLeft = 65
'xREP = xIMPDescrTxOutros2
'MargemLeft = MargemLeft + AuxLEFT
'Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 15
xREP = xIMPFreteTotal
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = 10.8
MargemTOP = MargemTOP + AuxTOP

'Dados de NFs
MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF01
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF02
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF03
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF04
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF05
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF06
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF07
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF08
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF09
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF10
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF11
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)


MargemTOP = MargemTOP + Linha
MargemLeft = 165
xREP = xIMPStrNF12
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPReqTranspMinuta
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 240
xREP = xIMPNumControle
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPInscrEstCiaAerea
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 245
xREP = xIMPCNPJCiaAerea
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPVlDecTRANSP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, 8, MargemLeft, MargemTOP, xREP)

MargemLeft = 250
xREP = xIMPVlDecSUFRAMA
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, 8, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 165
xREP = xIMPDescrEmbalagem
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 165
xREP = xIMPTxDescrDevAg
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 245
xREP = xIMPTxDescrDevTransp
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (5 * Linha)
MargemLeft = 170
xREP = xIMPStrObservacao01
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao02
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao03
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = xIMPStrObservacao04
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsSeguro)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsICMS)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + Linha
MargemLeft = 170
xREP = Trim(xIMPObsPerecivel)
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

Margem2TOP = (MargemTOP + (3 * Linha)) - 2
MargemLeft = 227
xREP = xIMPStrRetiraSIM
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, Margem2TOP, xREP)

Margem2TOP = (MargemTOP + (3 * Linha)) - 2
MargemLeft = 255
xREP = xIMPStrRetiraNAO
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, Margem2TOP, xREP)

MargemTOP = MargemTOP + (4 * Linha)
MargemLeft = 170
xREP = xIMPStrLocalRetira
MargemLeft = MargemLeft + AuxLEFT + AuxMODAL
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (8 * Linha)
MargemLeft = 280
xREP = xIMPStrTotalServ
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (1 * Linha)
MargemLeft = 280
xREP = xIMPStrBaseCalculo
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 176
xREP = xIMPStrAliquota
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 280
xREP = xIMPStrICMS
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 170
xREP = xIMPAgenteEmissor
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 258
xREP = xIMPCodIATA
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 170
xREP = xIMPDtEmissao
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 258
xREP = xIMPHoraEmissao
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemTOP = MargemTOP + (2 * Linha)
MargemLeft = 6
xREP = xIMPNaturezaOp
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 75
xREP = xIMPCFOP
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 100
xREP = xIMPEmissor
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)

MargemLeft = 260
xREP = xIMPLocalidade
MargemLeft = MargemLeft + AuxLEFT
Call EnviaIMP(Fonte, FontSize, MargemLeft, MargemTOP, xREP)
Printer.EndDoc

End Sub


Private Sub ImpressaoPorPrintVASP()

xAUX = 55
xIMPNomeEXP = xIMPNomeEXP & String(xAUX - Len(xIMPNomeEXP), " ")
xAUX = 55
xIMPCGCEXP = xIMPCGCEXP & String(xAUX - Len(xIMPCGCEXP), " ")
xAUX = 55
xIMPInscEstEXP = xIMPInscEstEXP & String(xAUX - Len(xIMPInscEstEXP), " ")
xAUX = 55
xIMPEndEXP = xIMPEndEXP & String(xAUX - Len(xIMPEndEXP), " ")
xAUX = 25
xIMPBairroEXP = xIMPBairroEXP & String(xAUX - Len(xIMPBairroEXP), " ")
xAUX = 25
xIMPCidadeEXP = xIMPCidadeEXP & String(xAUX - Len(xIMPCidadeEXP), " ")
xAUX = 20
xIMPCepEXP = xIMPCepEXP & String(xAUX - Len(xIMPCepEXP), " ")
xAUX = 5
xIMPUFEXP = xIMPUFEXP & String(xAUX - Len(xIMPUFEXP), " ")
xAUX = 20
xIMPTelEXP = xIMPTelEXP & String(xAUX - Len(xIMPTelEXP), " ")
xAUX = 55
xIMPFAXEXP = xIMPFAXEXP & String(xAUX - Len(xIMPFAXEXP), " ")
xAUX = 55
xIMPNomeDEST = xIMPNomeDEST & String(xAUX - Len(xIMPNomeDEST), " ")
xAUX = 55
xIMPCGCDEST = xIMPCGCDEST & String(xAUX - Len(xIMPCGCDEST), " ")
xAUX = 55
xIMPInscEstDEST = xIMPInscEstDEST & String(xAUX - Len(xIMPInscEstDEST), " ")
xAUX = 55
xIMPEndDEST = xIMPEndDEST & String(xAUX - Len(xIMPEndDEST), " ")
xAUX = 25
xIMPBairroDEST = xIMPBairroDEST & String(xAUX - Len(xIMPBairroDEST), " ")
xAUX = 25
xIMPCidadeDEST = xIMPCidadeDEST & String(xAUX - Len(xIMPCidadeDEST), " ")
xAUX = 20
xIMPCepDEST = xIMPCepDEST & String(xAUX - Len(xIMPCepDEST), " ")
xAUX = 5
xIMPUFDEST = xIMPUFDEST & String(xAUX - Len(xIMPUFDEST), " ")
xAUX = 20
xIMPTelDEST = xIMPTelDEST & String(xAUX - Len(xIMPTelDEST), " ")
xAUX = 55
xIMPFAXDEST = xIMPFAXDEST & String(xAUX - Len(xIMPFAXDEST), " ")
xAUX = 10
xIMPOrigem = String(Int((xAUX - Len(xIMPOrigem)) / 2), " ") & xIMPOrigem & String((xAUX - Len(xIMPOrigem)) - Int((xAUX - Len(xIMPOrigem)) / 2), " ")
xAUX = 11
xIMPVia = String(Int((xAUX - Len(xIMPVia)) / 2), " ") & xIMPVia & String((xAUX - Len(xIMPVia)) - Int((xAUX - Len(xIMPVia)) / 2), " ")
xAUX = 25
xIMPCidadeDESTINO = String(Int((xAUX - Len(xIMPCidadeDESTINO)) / 2), " ") & xIMPCidadeDESTINO & String((xAUX - Len(xIMPCidadeDESTINO)) - Int((xAUX - Len(xIMPCidadeDESTINO)) / 2), " ")
xAUX = 13
xIMPSIGLA = String(Int((xAUX - Len(xIMPSIGLA)) / 2), " ") & xIMPSIGLA & String((xAUX - Len(xIMPSIGLA)) - Int((xAUX - Len(xIMPSIGLA)) / 2), " ")
xAUX = 20
xIMPReqTranspMinuta = String(Int((xAUX - Len(xIMPReqTranspMinuta)) / 2), " ") & xIMPReqTranspMinuta & String((xAUX - Len(xIMPReqTranspMinuta)) - Int((xAUX - Len(xIMPReqTranspMinuta)) / 2), " ")
xAUX = 37
xIMPNumControle = String(Int((xAUX - Len(xIMPNumControle)) / 2), " ") & xIMPNumControle & String((xAUX - Len(xIMPNumControle)) - Int((xAUX - Len(xIMPNumControle)) / 2), " ")
xAUX = 29
xIMPInscrEstCiaAerea = String(Int((xAUX - Len(xIMPInscrEstCiaAerea)) / 2), " ") & xIMPInscrEstCiaAerea & String((xAUX - Len(xIMPInscrEstCiaAerea)) - Int((xAUX - Len(xIMPInscrEstCiaAerea)) / 2), " ")
xAUX = 29
xIMPCNPJCiaAerea = String(Int((xAUX - Len(xIMPCNPJCiaAerea)) / 2), " ") & xxIMPCNPJCiaAerea & String((xAUX - Len(xIMPCNPJCiaAerea)) - Int((xAUX - Len(xIMPCNPJCiaAerea)) / 2), " ")
xAUX = 29
xIMPVlDecTRANSP = String(xAUX - Len(xIMPVlDecTRANSP), " ") & xIMPVlDecTRANSP
xAUX = 29
xIMPVlDecSUFRAMA = String(xAUX - Len(xIMPVlDecSUFRAMA), " ") & xIMPVlDecSUFRAMA
xAUX = 62
xIMPDescrEmbalagem = String(Int((xAUX - Len(xIMPDescrEmbalagem)) / 2), " ") & xIMPDescrEmbalagem & String((xAUX - Len(xIMPDescrEmbalagem)) - Int((xAUX - Len(xIMPDescrEmbalagem)) / 2), " ")
xAUX = 7
xIMPQteVol = String(Int((xAUX - Len(xIMPQteVol)) / 2), " ") & xIMPQteVol & String((xAUX - Len(xIMPQteVol)) - Int((xAUX - Len(xIMPQteVol)) / 2), " ")
xAUX = 11
xIMPPesoReal = String(Int((xAUX - Len(xIMPPesoReal)) / 2), " ") & xIMPPesoReal & String((xAUX - Len(xIMPPesoReal)) - Int((xAUX - Len(xIMPPesoReal)) / 2), " ")
xAUX = 11
xIMPPesoTax = String(Int((xAUX - Len(xIMPPesoTax)) / 2), " ") & xIMPPesoTax & String((xAUX - Len(xIMPPesoTax)) - Int((xAUX - Len(xIMPPesoTax)) / 2), " ")
xAUX = 13
xIMPTrecho = String(Int((xAUX - Len(xIMPTrecho)) / 2), " ") & xIMPTrecho & String((xAUX - Len(xIMPTrecho)) - Int((xAUX - Len(xIMPTrecho)) / 2), " ")
xAUX = 4
xIMPCl = String(Int((xAUX - Len(xIMPCl)) / 2), " ") & xIMPCl & String((xAUX - Len(xIMPCl)) - Int((xAUX - Len(xIMPCl)) / 2), " ")
xAUX = 7
xIMPCodigo = String(Int((xAUX - Len(xIMPCodigo)) / 2), " ") & xIMPCodigo & String((xAUX - Len(xIMPCodigo)) - Int((xAUX - Len(xIMPCodigo)) / 2), " ")
xAUX = 12
xIMPKilo = String(Int((xAUX - Len(xIMPKilo)) / 2), " ") & xIMPKilo & String((xAUX - Len(xIMPKilo)) - Int((xAUX - Len(xIMPKilo)) / 2), " ")
xAUX = 19
xIMPFreteNacEscopo = String(Int((xAUX - Len(xIMPFreteNacEscopo)) / 2), " ") & xIMPFreteNacEscopo & String((xAUX - Len(xIMPFreteNacEscopo)) - Int((xAUX - Len(xIMPFreteNacEscopo)) / 2), " ")
xAUX = 33
xIMPNatureza = String(Int((xAUX - Len(xIMPNatureza)) / 2), " ") & xIMPNatureza & String((xAUX - Len(xIMPNatureza)) - Int((xAUX - Len(xIMPNatureza)) / 2), " ")
xAUX = 29
xIMPTxDescrDevAg = String(Int((xAUX - Len(xIMPTxDescrDevAg)) / 2), " ") & xIMPTxDescrDevAg & String((xAUX - Len(xIMPTxDescrDevAg)) - Int((xAUX - Len(xIMPTxDescrDevAg)) / 2), " ")
xAUX = 29
xIMPTxDescrDevTransp = String(Int((xAUX - Len(xIMPTxDescrDevTransp)) / 2), " ") & xIMPTxDescrDevTransp & String((xAUX - Len(xIMPTxDescrDevTransp)) - Int((xAUX - Len(xIMPTxDescrDevTransp)) / 2), " ")
xAUX = 21
xIMPFreteNacional = String(xAUX - Len(xIMPFreteNacional), " ") & xIMPFreteNacional
xAUX = 21
xIMPFreteRegional = String(xAUX - Len(xIMPFreteRegional), " ") & xIMPFreteRegional
xAUX = 21
xIMPAdValorem = String(xAUX - Len(xIMPAdValorem), " ") & xIMPAdValorem
xAUX = 6
xIMPTipoADVAL = String(Int((xAUX - Len(xIMPTipoADVAL)) / 2), " ") & xIMPTipoADVAL & String((xAUX - Len(xIMPTipoADVAL)) - Int((xAUX - Len(xIMPTipoADVAL)) / 2), " ")
xAUX = 21
xIMPTxTerrOrig = String(xAUX - Len(xIMPTxTerrOrig), " ") & xIMPTxTerrOrig
xAUX = 21
xIMPTxTerrDest = String(xAUX - Len(xIMPTxTerrDest), " ") & xIMPTxTerrDest
xAUX = 21
xIMPTxRedesp = String(xAUX - Len(xIMPTxRedesp), " ") & xIMPTxRedesp
xAUX = 21
xIMPTxAgente = String(xAUX - Len(xIMPTxAgente), " ") & xIMPTxAgente
xAUX = 21
xIMPTxDevTransp = String(xAUX - Len(xIMPTxDevTransp), " ") & xIMPTxDevTransp
xAUX = 14
xIMPDescrTxOutros1 = String(Int((xAUX - Len(xIMPDescrTxOutros1)) / 2), " ") & xIMPDescrTxOutros1 & String((xAUX - Len(xIMPDescrTxOutros1)) - Int((xAUX - Len(xIMPDescrTxOutros1)) / 2), " ")
xAUX = 21
xIMPTxOutros1 = String(xAUX - Len(xIMPTxOutros1), " ") & xIMPTxOutros1
xAUX = 14
xIMPDescrTxOutros2 = String(Int((xAUX - Len(xIMPDescrTxOutros2)) / 2), " ") & xIMPDescrTxOutros2 & String((xAUX - Len(xIMPDescrTxOutros2)) - Int((xAUX - Len(xIMPDescrTxOutros2)) / 2), " ")
xAUX = 21
xIMPTxOutros2 = String(xAUX - Len(xIMPTxOutros2), " ") & xIMPTxOutros2
xAUX = 21
xIMPFreteTotal = String(xAUX - Len(xIMPFreteTotal), " ") & xIMPFreteTotal
'xaux =
'xIMPStrRetira = xIMPStrRetira & String(xaux - Len(xIMPStrRetira), " ")
'xaux =
'xIMPStrLocalRetira = xIMPStrLocalRetira & String(xaux - Len(xIMPStrLocalRetira), " ")
xAUX = 60
xIMPHorarioAt = xIMPHorarioAt & String(xAUX - Len(xIMPHorarioAt), " ")
xAUX = 60
xIMPStrTelefone = xIMPStrTelefone & String(xAUX - Len(xIMPStrTelefone), " ")
xAUX = 23
xIMPStrTotalServ = String(xAUX - Len(xIMPStrTotalServ), " ") & xIMPStrTotalServ
xAUX = 23
xIMPStrBaseCalculo = String(xAUX - Len(xIMPStrBaseCalculo), " ") & xIMPStrBaseCalculo
xAUX = 6
xIMPStrAliquota = String(xAUX - Len(Trim(xIMPStrAliquota)), " ") & xIMPStrAliquota
xAUX = 23
xIMPStrICMS = String(xAUX - Len(xIMPStrICMS), " ") & xIMPStrICMS
xAUX = 27
xIMPAgenteEmissor = String(Int((xAUX - Len(xIMPAgenteEmissor)) / 2), " ") & xIMPAgenteEmissor & String((xAUX - Len(xIMPAgenteEmissor)) - Int((xAUX - Len(xIMPAgenteEmissor)) / 2), " ")
xAUX = 27
xIMPCodIATA = String(Int((xAUX - Len(xIMPCodIATA)) / 2), " ") & xIMPCodIATA & String((xAUX - Len(xIMPCodIATA)) - Int((xAUX - Len(xIMPCodIATA)) / 2), " ")
xAUX = 27
xIMPDtEmissao = String(Int((xAUX - Len(xIMPDtEmissao)) / 2), " ") & xIMPDtEmissao & String((xAUX - Len(xIMPDtEmissao)) - Int((xAUX - Len(xIMPDtEmissao)) / 2), " ")
xAUX = 27
xIMPHoraEmissao = String(Int((xAUX - Len(xIMPHoraEmissao)) / 2), " ") & xIMPHoraEmissao & String((xAUX - Len(xIMPHoraEmissao)) - Int((xAUX - Len(xIMPHoraEmissao)) / 2), " ")
xAUX = 25
xIMPNaturezaOp = xIMPNaturezaOp & String(xAUX - Len(xIMPNaturezaOp), " ")
xAUX = 8
xIMPCFOP = xIMPCFOP & String(xAUX - Len(xIMPCFOP), " ")
xAUX = 59
xIMPEmissor = xIMPEmissor & String(xAUX - Len(xIMPEmissor), " ")
xAUX = 8
xIMPLocalidade = xIMPLocalidade & String(xAUX - Len(xIMPLocalidade), " ")
xAUX = 17
xIMPMatricula = xIMPMatricula & String(xAUX - Len(xIMPMatricula), " ")


xAUX = 8
If OptAPagar.Value = True Then xAUX = 50

Open SETIMPImpressoraPadrao For Output As #1
DoEvents
Print #1, Chr(15)
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, "                                                                         " & Mid(StringNF, 1, 60)
Print #1, "              " & xIMPNomeEXP & "  " & Mid(StringNF, 61, 60)
Print #1, "              " & xIMPCGCEXP & "  " & Mid(StringNF, 121, 60)
Print #1, "                " & xIMPInscEstEXP & "  " & Mid(StringNF, 181, 60)
Print #1, "            " & xIMPEndEXP & "  " & Mid(StringNF, 241, 60)
Print #1, "              " & xIMPBairroEXP & "     " & xIMPCidadeEXP & "  " & Mid(StringNF, 301, 60)
Print #1, "              " & xIMPCepEXP & "    " & xIMPUFEXP & "      " & xIMPTelEXP & "  " & Mid(StringNF, 361, 60)
Print #1, "              " & xfaxexp & "  " & Mid(StringNF, 361, 60)
Print #1, "                                                                         " & Mid(StringNF, 421, 60)
Print #1, "                                                                         " & Mid(StringNF, 481, 60)
Print #1, "                                                                         " & Mid(StringNF, 541, 60)

If OptPago.Value = True Then
Print #1, "              " & xIMPNomeDEST & "    F R E T E   P A G O"
Else
Print #1, "              " & xIMPNomeDEST & "    F R E T E   A   P A G A R"
End If

Print #1, "              " & xIMPCGCDEST & "                                                                     "
Print #1, "                " & xIMPInscEstDEST & "  " & xIMPReqTranspMinuta & "     " & xIMPNumControle
Print #1, "            " & xIMPEndDEST
Print #1, "              " & xIMPBairroDEST & "     " & xIMPCidadeDEST & "  " & xIMPInscrEstCiaAerea & "    " & xIMPCNPJCiaAerea
Print #1, "              " & xIMPCepDEST & "    " & xIMPUFDEST & "      " & xIMPTelDEST
Print #1, "              " & xIMPFAXDEST & "  " & xIMPVlDecTRANSP & "    " & xIMPVlDecSUFRAMA
Print #1, ""
Print #1, "       " & xIMPOrigem & "  " & xIMPVia & " " & xIMPCidadeDESTINO & "  " & xIMPSIGLA & "  " & xIMPDescrEmbalagem
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, "       " & xIMPQteVol & " " & xIMPPesoReal & "  " & xIMPPesoTax & " " & xIMPTrecho & " " & xIMPCl & " " & xIMPCodigo & "  " & xIMPKilo & "  " & xIMPFreteNacEscopo & " " & xIMPNatureza
Print #1, ""
Print #1, ""
Print #1, "                                                                          " & xIMPTxDescrDevAg & "   " & xIMPTxDescrDevTransp

    If OptPago.Value = True Then
    Print #1, ""
    Print #1, "        " & xIMPFreteNacional & "                                             " & Mid(xIMPStrObservacao, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 61, 60)
    Print #1, "        " & xIMPFreteRegional & "                                             " & Mid(xIMPStrObservacao, 121, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 181, 60)
    Print #1, "        " & xIMPAdValorem & "             " & xIMPTipoADVAL & "                          " & Mid(xIMPStrObservacao, 241, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 301, 60)
    Print #1, "        " & xIMPTxTerrOrig & "                                             " & Mid(xIMPStrObservacao, 361, 60)
    Print #1, "                                                                                                                                            "
        If OptRetiraSim.Value = True Then
        Print #1, "        " & xIMPTxTerrDest & "                                                                     XXX"
        Else
        Print #1, "        " & xIMPTxTerrDest & "                                                                                XXX"
        End If
    Print #1, ""
    Print #1, "        " & xIMPTxRedesp & "                                             " & Mid(xIMPStrLocalRetira, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrLocalRetira, 61, 60)
    Print #1, "        " & xIMPTxAgente
    Print #1, ""
    Print #1, "        " & xIMPTxDevTransp
    Print #1, ""
    Print #1, "        " & xIMPTxOutros1 & "   " & xIMPDescrTxOutros1
    Print #1, "                                                                                                              " & xIMPStrTotalServ
    Print #1, "        " & xIMPTxOutros2 & "   " & xIMPDescrTxOutros2 & "                                                                " & xIMPStrBaseCalculo
    Print #1, ""
    Print #1, "        " & xIMPFreteTotal & "                                              " & xIMPStrAliquota & "                    " & xIMPStrICMS
    
    Else
    
    Print #1, ""
    Print #1, "                                                  " & xIMPFreteNacional & "   " & Mid(xIMPStrObservacao, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 61, 60)
    Print #1, "                                                  " & xIMPFreteRegional & "   " & Mid(xIMPStrObservacao, 121, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 181, 60)
    Print #1, "                                          " & xIMPTipoADVAL & "  " & xIMPAdValorem & "   " & Mid(xIMPStrObservacao, 241, 60)
    Print #1, "                                                                          " & Mid(xIMPStrObservacao, 301, 60)
    Print #1, "                                                  " & xIMPTxTerrOrig & "   " & Mid(xIMPStrObservacao, 361, 60)
    Print #1, "                                                                                                                                            "
        If OptRetiraSim.Value = True Then
        Print #1, "                                                  " & xIMPTxTerrDest & "                           XXX"
        Else
        Print #1, "                                                  " & xIMPTxTerrDest & "                                      XXX"
        End If
    Print #1, ""
    Print #1, "                                                  " & xIMPTxRedesp & "    " & Mid(xIMPStrLocalRetira, 1, 60)
    Print #1, "                                                                          " & Mid(xIMPStrLocalRetira, 61, 60)
    Print #1, "                                                  " & xIMPTxAgente
    Print #1, ""
    Print #1, "                                                  " & xIMPTxDevTransp
    Print #1, ""
    Print #1, "                                " & xIMPDescrTxOutros1 & "    " & xIMPDescrTxOutros2
    Print #1, "                                                                                                              " & xIMPStrTotalServ
    Print #1, "                                " & xIMPDescrTxOutros2 & "   " & xIMPTxOutros2 & "                                        " & xIMPStrBaseCalculo
    Print #1, ""
    Print #1, "                                                  " & xIMPFreteTotal & "               " & xIMPStrAliquota & "                    " & xIMPStrICMS
    End If

Print #1, ""
Print #1, "                                                                          " & xIMPAgenteEmissor & "     " & xIMPCodIATA
Print #1, ""
Print #1, "                                                                          " & xIMPDtEmissao & "     " & xIMPHoraEmissao
Print #1, ""
Print #1, "        " & xIMPNaturezaOp & "  " & xIMPCFOP & "   " & xIMPEmissor & "   " & xIMPLocalidade & "  " & xIMPMatricula
DoEvents
Close #1

End Sub

Private Sub MascaraAWBTAM()

Dim Linha(1 To 67) As String
Dim EspacosESQ As String
Dim EspacosMeio As String
Dim EspacosAUX As String

Dim Z As String

Dim xPAGOFreteNacional As String
Dim xPAGOAdValorem As String
Dim xPAGOTxTerrOrig As String
Dim xPAGOTxTerrDest As String
Dim xPAGOTxRedesp As String
Dim xPAGOTxAgente As String
Dim xPAGOTxDevTransp As String
Dim xPAGOTxOutros1 As String
Dim xPAGOTxOutros2 As String
Dim xPAGOFreteTotal As String
Dim xAPAGARFreteNacional As String
Dim xAPAGARAdValorem As String
Dim xAPAGARTxTerrOrig As String
Dim xAPAGARTxTerrDest As String
Dim xAPAGARTxRedesp As String
Dim xAPAGARTxAgente As String
Dim xAPAGARTxDevTransp As String
Dim xAPAGARTxOutros1 As String
Dim xAPAGARTxOutros2 As String
Dim xAPAGARFreteTotal As String


StringNF = Trim(StringNF)
xIMPStrNF01 = Trim(xIMPStrNF01)
xIMPStrNF02 = Trim(xIMPStrNF02)
xIMPStrNF03 = Trim(xIMPStrNF03)
xIMPStrNF04 = Trim(xIMPStrNF04)
xIMPStrNF05 = Trim(xIMPStrNF05)
xIMPStrNF06 = Trim(xIMPStrNF06)
xIMPStrNF07 = Trim(xIMPStrNF07)
xIMPStrNF08 = Trim(xIMPStrNF08)
xIMPStrNF09 = Trim(xIMPStrNF09)
xIMPStrNF10 = Trim(xIMPStrNF10)
xIMPStrNF11 = Trim(xIMPStrNF11)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPNomeEXP = Trim(xIMPNomeEXP)
xIMPCGCEXP = Trim(xIMPCGCEXP)
xIMPInscEstEXP = Trim(xIMPInscEstEXP)
xIMPEndEXP = Trim(xIMPEndEXP)
xIMPBairroEXP = Trim(xIMPBairroEXP)
xIMPCidadeEXP = Trim(xIMPCidadeEXP)
xIMPCepEXP = Trim(xIMPCepEXP)
xIMPUFEXP = Trim(xIMPUFEXP)
xIMPTelEXP = Trim(xIMPTelEXP)
xIMPFAXEXP = Trim(xIMPFAXEXP)
xIMPNomeDEST = Trim(xIMPNomeDEST)
xIMPCGCDEST = Trim(xIMPCGCDEST)
xIMPInscEstDEST = Trim(xIMPInscEstDEST)
xIMPEndDEST = Trim(xIMPEndDEST)
xIMPBairroDEST = Trim(xIMPBairroDEST)
xIMPCidadeDEST = Trim(xIMPCidadeDEST)
xIMPCepDEST = Trim(xIMPCepDEST)
xIMPUFDEST = Trim(xIMPUFDEST)
xIMPTelDEST = Trim(xIMPTelDEST)
xIMPFAXDEST = Trim(xIMPFAXDEST)
xIMPOrigem = Trim(xIMPOrigem)
xIMPVia = Trim(xIMPVia)
xIMPCidadeDESTINO = Trim(xIMPCidadeDESTINO)
xIMPSIGLA = Trim(xIMPSIGLA)
xIMPReqTranspMinuta = Trim(xIMPReqTranspMinuta)
xIMPNumControle = Trim(xIMPNumControle)
xIMPInscrEstCiaAerea = Trim(xIMPInscrEstCiaAerea)
xIMPCNPJCiaAerea = Trim(xIMPCNPJCiaAerea)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPVlDecTRANSP = Trim(xIMPVlDecTRANSP)
xIMPVlDecSUFRAMA = Trim(xIMPVlDecSUFRAMA)
xIMPDescrEmbalagem = Trim(xIMPDescrEmbalagem)
xIMPQteVol = Trim(xIMPQteVol)
xIMPPesoReal = Trim(xIMPPesoReal)
xIMPPesoTax = Trim(xIMPPesoTax)
xIMPTrecho = Trim(xIMPTrecho)
xIMPCl = Trim(xIMPCl)
xIMPCodigo = Trim(xIMPCodigo)
xIMPKilo = Trim(xIMPKilo)
xIMPFreteNacEscopo = Trim(xIMPFreteNacEscopo)
xIMPNatureza = Trim(xIMPNatureza)
xIMPTxDescrDevAg = Trim(xIMPTxDescrDevAg)
xIMPTxDescrDevTransp = Trim(xIMPTxDescrDevTransp)
xIMPFreteNacional = Trim(xIMPFreteNacional)
xIMPFreteRegional = Trim(xIMPFreteRegional)
xIMPAdValorem = Trim(xIMPAdValorem)
xIMPTipoADVAL = Trim(xIMPTipoADVAL)
xIMPTxTerrOrig = Trim(xIMPTxTerrOrig)
xIMPTxTerrDest = Trim(xIMPTxTerrDest)
xIMPTxRedesp = Trim(xIMPTxRedesp)
xIMPTxAgente = Trim(xIMPTxAgente)
xIMPTxDevTransp = Trim(xIMPTxDevTransp)
xIMPDescrTxOutros1 = Trim(xIMPDescrTxOutros1)
xIMPTxOutros1 = Trim(xIMPTxOutros1)
xIMPDescrTxOutros2 = Trim(xIMPDescrTxOutros2)
xIMPTxOutros2 = Trim(xIMPTxOutros2)
xIMPFreteTotal = Trim(xIMPFreteTotal)
xIMPStrObservacao = Trim(xIMPStrObservacao)
xIMPStrObservacao01 = Trim(xIMPStrObservacao01)
xIMPStrObservacao02 = Trim(xIMPStrObservacao02)
xIMPStrObservacao03 = Trim(xIMPStrObservacao03)
xIMPStrObservacao04 = Trim(xIMPStrObservacao04)
xIMPObsICMS = Trim(xIMPObsICMS)
xIMPObsPerecivel = Trim(xIMPObsPerecivel)
xIMPObsSeguro = Trim(xIMPObsSeguro)
xIMPStrRetiraSIM = Trim(xIMPStrRetiraSIM)
xIMPStrRetiraNAO = Trim(xIMPStrRetiraNAO)
xIMPStrLocalRetira = Trim(xIMPStrLocalRetira)
xIMPHorarioAt = Trim(xIMPHorarioAt)
xIMPStrTelefone = Trim(xIMPStrTelefone)
xIMPStrTotalServ = Trim(xIMPStrTotalServ)
xIMPStrBaseCalculo = Trim(xIMPStrBaseCalculo)
xIMPStrAliquota = Trim(xIMPStrAliquota)
xIMPStrICMS = Trim(xIMPStrICMS)
xIMPAgenteEmissor = Trim(xIMPAgenteEmissor)
xIMPCodIATA = Trim(xIMPCodIATA)
'xIMPDtEmissao = DataHora("DATA")
'xIMPHoraEmissao = DataHora("HORA")
xIMPDtEmissao = Trim(xDataIMP)
xIMPHoraEmissao = Trim(xIMPHoraEmissao)
xIMPNaturezaOp = Trim(xIMPNaturezaOp)
xIMPCFOP = Trim(xIMPCFOP)
'xIMPEmissor = xUsuario
xIMPEmissor = Trim(xIMPEmissor)
xIMPLocalidade = Trim(xIMPLocalidade)
xIMPMatricula = Trim(xIMPMatricula)

xIMPStrNF01 = Mid(xIMPStrNF01, 1, 60)
xIMPStrNF02 = Mid(xIMPStrNF02, 1, 60)
xIMPStrNF03 = Mid(xIMPStrNF03, 1, 60)
xIMPStrNF04 = Mid(xIMPStrNF04, 1, 60)
xIMPStrNF05 = Mid(xIMPStrNF05, 1, 60)
xIMPStrNF06 = Mid(xIMPStrNF06, 1, 60)
xIMPStrNF07 = Mid(xIMPStrNF07, 1, 60)
xIMPStrNF08 = Mid(xIMPStrNF08, 1, 60)
xIMPStrNF09 = Mid(xIMPStrNF09, 1, 60)
xIMPStrNF10 = Mid(xIMPStrNF10, 1, 60)
xIMPStrNF11 = Mid(xIMPStrNF11, 1, 60)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 60)
xIMPNomeEXP = Mid(xIMPNomeEXP, 1, 40)
xIMPCGCEXP = Mid(xIMPCGCEXP, 1, 40)
xIMPInscEstEXP = Mid(xIMPInscEstEXP, 1, 40)
xIMPEndEXP = Mid(xIMPEndEXP, 1, 40)
xIMPBairroEXP = Mid(xIMPBairroEXP, 1, 23)
xIMPCidadeEXP = Mid(xIMPCidadeEXP, 1, 29)
xIMPCepEXP = Mid(xIMPCepEXP, 1, 13)
xIMPUFEXP = Mid(xIMPUFEXP, 1, 7)
xIMPTelEXP = Mid(xIMPTelEXP, 1, 18)
xIMPFAXEXP = Mid(xIMPFAXEXP, 1, 18)
xIMPNomeDEST = Mid(xIMPNomeDEST, 1, 40)
xIMPCGCDEST = Mid(xIMPCGCDEST, 1, 40)
xIMPInscEstDEST = Mid(xIMPInscEstDEST, 1, 40)
xIMPEndDEST = Mid(xIMPEndDEST, 1, 40)
xIMPBairroDEST = Mid(xIMPBairroDEST, 1, 23)
xIMPCidadeDEST = Mid(xIMPCidadeDEST, 1, 29)
xIMPCepDEST = Mid(xIMPCepDEST, 1, 13)
xIMPUFDEST = Mid(xIMPUFDEST, 1, 7)
xIMPTelDEST = Mid(xIMPTelDEST, 1, 14)
xIMPFAXDEST = Mid(xIMPFAXDEST, 1, 40)
xIMPOrigem = Mid(xIMPOrigem, 1, 7)
xIMPVia = Mid(xIMPVia, 1, 6)
xIMPCidadeDESTINO = Mid(xIMPCidadeDESTINO, 1, 18)
xIMPSIGLA = Mid(xIMPSIGLA, 1, 8)
xIMPReqTranspMinuta = Mid(xIMPReqTranspMinuta, 1, 15)
xIMPNumControle = Mid(xIMPNumControle, 1, 25)
xIMPInscrEstCiaAerea = Mid(xIMPInscrEstCiaAerea, 1, 20)
xIMPCNPJCiaAerea = Mid(xIMPCNPJCiaAerea, 1, 20)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 35)
xIMPVlDecTRANSP = Mid(xIMPVlDecTRANSP, 1, 28)
xIMPVlDecSUFRAMA = Mid(xIMPVlDecSUFRAMA, 1, 28)
xIMPDescrEmbalagem = Mid(xIMPDescrEmbalagem, 1, 60)
xIMPQteVol = Mid(xIMPQteVol, 1, 5)
xIMPPesoReal = Mid(xIMPPesoReal, 1, 7)
xIMPPesoTax = Mid(xIMPPesoTax, 1, 8)
xIMPTrecho = Mid(xIMPTrecho, 1, 8)
xIMPCl = Mid(xIMPCl, 1, 2)
xIMPCodigo = Mid(xIMPCodigo, 1, 3)
xIMPKilo = Mid(xIMPKilo, 1, 9)
xIMPFreteNacEscopo = Mid(xIMPFreteNacEscopo, 1, 13)
xIMPTxDescrDevAg = Mid(xIMPTxDescrDevAg, 1, 1)
xIMPTxDescrDevTransp = Mid(xIMPTxDescrDevTransp, 1, 1)
xIMPDescrTxOutros1 = Mid(xIMPDescrTxOutros1, 1, 12)
xIMPDescrTxOutros2 = Mid(xIMPDescrTxOutros2, 1, 12)
xIMPStrObservacao01 = Mid(xIMPStrObservacao01, 1, 60)
xIMPStrObservacao02 = Mid(xIMPStrObservacao02, 1, 60)
xIMPStrObservacao03 = Mid(xIMPStrObservacao03, 1, 60)
xIMPStrObservacao04 = Mid(xIMPStrObservacao04, 1, 60)
xIMPObsICMS = Mid(xIMPObsICMS, 1, 60)
xIMPObsSeguro = Mid(xIMPObsSeguro, 1, 60)
xIMPStrLocalRetira = Mid(xIMPStrLocalRetira, 1, 42)
xIMPAgenteEmissor = Mid(xIMPAgenteEmissor, 1, 32)
xIMPCodIATA = Mid(xIMPCodIATA, 1, 28)
xIMPDtEmissao = Mid(xIMPDtEmissao, 1, 22)
xIMPHoraEmissao = Mid(xIMPHoraEmissao, 1, 18)
xIMPNaturezaOp = Mid(xIMPNaturezaOp, 1, 19)
xIMPCFOP = Mid(xIMPCFOP, 1, 6)
xIMPEmissor = Mid(xIMPEmissor, 1, 42)
xIMPLocalidade = Mid(xIMPLocalidade, 1, 6)





'65 LINHAS - 9 LINHAS EM BRANCO NO COMECO

EspacosESQ = String(3, " ")
EspacosMeio = String(2, " ")
EspacosAUX = String(5, " ")
Z = " "

If OptPago.Value = True Then
    xPAGOFreteNacional = xIMPFreteNacional
    xPAGOAdValorem = xIMPAdValorem
    xPAGOTxTerrOrig = xIMPTxTerrOrig
    xPAGOTxTerrDest = xIMPTxTerrDest
    xPAGOTxRedesp = xIMPTxRedesp
    xPAGOTxAgente = xIMPTxAgente
    xPAGOTxDevTransp = xIMPTxDevTransp
    xPAGOTxOutros1 = xIMPTxOutros1
    xPAGOTxOutros2 = xIMPTxOutros2
    xPAGOFreteTotal = xIMPFreteTotal
    xAPAGARFreteNacional = ""
    xAPAGARAdValorem = ""
    xAPAGARTxTerrOrig = ""
    xAPAGARTxTerrDest = ""
    xAPAGARTxRedesp = ""
    xAPAGARTxAgente = ""
    xAPAGARTxDevTransp = ""
    xAPAGARTxOutros1 = ""
    xAPAGARTxOutros2 = ""
    xAPAGARFreteTotal = ""
Else
    xPAGOFreteNacional = ""
    xPAGOAdValorem = ""
    xPAGOTxTerrOrig = ""
    xPAGOTxTerrDest = ""
    xPAGOTxRedesp = ""
    xPAGOTxAgente = ""
    xPAGOTxDevTransp = ""
    xPAGOTxOutros1 = ""
    xPAGOTxOutros2 = ""
    xPAGOFreteTotal = ""
    xAPAGARFreteNacional = xIMPFreteNacional
    xAPAGARAdValorem = xIMPAdValorem
    xAPAGARTxTerrOrig = xIMPTxTerrOrig
    xAPAGARTxTerrDest = xIMPTxTerrDest
    xAPAGARTxRedesp = xIMPTxRedesp
    xAPAGARTxAgente = xIMPTxAgente
    xAPAGARTxDevTransp = xIMPTxDevTransp
    xAPAGARTxOutros1 = xIMPTxOutros1
    xAPAGARTxOutros2 = xIMPTxOutros2
    xAPAGARFreteTotal = xIMPFreteTotal
End If


Linha(1) = ""
Linha(2) = ""
Linha(3) = ""
Linha(4) = ""
Linha(5) = ""
Linha(6) = ""
Linha(7) = ""
Linha(8) = ""
Linha(9) = ""
Linha(10) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, " ") & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF01 & String(60 - Len(Trim(xIMPStrNF01)), Z)
Linha(11) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, " ") & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF02 & String(60 - Len(Trim(xIMPStrNF02)), Z)
Linha(12) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeEXP & String(40 - Len(Trim(xIMPNomeEXP)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF03 & String(60 - Len(Trim(xIMPStrNF03)), Z)
Linha(13) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCEXP & String(40 - Len(Trim(xIMPCGCEXP)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF04 & String(60 - Len(Trim(xIMPStrNF04)), Z)
Linha(14) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstEXP & String(40 - Len(Trim(xIMPInscEstEXP)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF05 & String(60 - Len(Trim(xIMPStrNF05)), Z)
Linha(15) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndEXP & String(40 - Len(Trim(xIMPEndEXP)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF06 & String(60 - Len(Trim(xIMPStrNF06)), Z)
Linha(16) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroEXP & String(23 - Len(Trim(xIMPBairroEXP)), Z) & String(4, Z) & xIMPCepEXP & String(13 - Len(Trim(xIMPCepEXP)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF07 & String(60 - Len(Trim(xIMPStrNF07)), Z)
Linha(17) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCidadeEXP & String(29 - Len(Trim(xIMPCidadeEXP)), Z) & String(4, Z) & xIMPUFEXP & String(7 - Len(Trim(xIMPUFEXP)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF08 & String(60 - Len(Trim(xIMPStrNF08)), Z)
Linha(18) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPTelEXP & String(18 - Len(Trim(xIMPTelEXP)), Z) & String(4, Z) & xIMPFAXEXP & String(18 - Len(Trim(xIMPFAXEXP)), Z) & EspacosMeio & Chr(27) & "!" & Chr(72) & xIMPStrNF12 & String(34 - Len(Trim(xIMPStrNF12)), Z)
Linha(19) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(42 - Len(Trim("")), Z)
Linha(20) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(21) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(22) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeDEST & String(40 - Len(Trim(xIMPNomeDEST)), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPReqTranspMinuta & String(19 - Len(Trim(xIMPReqTranspMinuta)), Z) & String(3, Z) & xIMPNumControle & String(20 - Len(Trim(xIMPNumControle)), Z)
Linha(23) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCDEST & String(40 - Len(Trim(xIMPCGCDEST)), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(19 - Len(Trim("")), Z) & String(3, Z) & "" & String(20 - Len(Trim("")), Z)
Linha(24) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstDEST & String(40 - Len(Trim(xIMPInscEstDEST)), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPInscrEstCiaAerea & String(19 - Len(Trim(xIMPInscrEstCiaAerea)), Z) & String(3, Z) & xIMPCNPJCiaAerea & String(20 - Len(Trim(xIMPCNPJCiaAerea)), Z)
Linha(25) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndDEST & String(40 - Len(Trim(xIMPEndDEST)), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(19 - Len(Trim("")), Z) & String(3, Z) & "" & String(20 - Len(Trim("")), Z)
Linha(26) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroDEST & String(23 - Len(Trim(xIMPBairroDEST)), Z) & String(4, Z) & xIMPCepDEST & String(13 - Len(Trim(xIMPCepDEST)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPVlDecTRANSP & String(25 - Len(xIMPVlDecTRANSP), Z) & String(10, Z) & xIMPVlDecSUFRAMA & String(25 - Len(xIMPVlDecSUFRAMA), Z)
Linha(27) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCidadeDEST & String(29 - Len(Trim(xIMPCidadeDEST)), Z) & String(4, Z) & xIMPUFDEST & String(7 - Len(Trim(xIMPUFDEST)), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(19 - Len(Trim("")), Z) & String(3, Z) & "" & String(20 - Len(Trim("")), Z)
Linha(28) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPTelDEST & String(18 - Len(Trim(xIMPTelDEST)), Z) & String(4, Z) & xIMPFAXDEST & String(18 - Len(Trim(xIMPFAXDEST)), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPDescrEmbalagem & String(60 - Len(xIMPDescrEmbalagem), Z)
Linha(29) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(19 - Len(Trim("")), Z) & String(3, Z) & "" & String(20 - Len(Trim("")), Z)
Linha(30) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPOrigem & String(7 - Len(Trim(xIMPOrigem)), Z) & String(2, Z) & xIMPVia & String(6 - Len(Trim(xIMPVia)), Z) & String(3, Z) & xIMPCidadeDESTINO & String(18 - Len(Trim(xIMPCidadeDESTINO)), Z) & String(3, Z) & xIMPSIGLA & String(6 - Len(Trim(xIMPSIGLA)), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(19 - Len(Trim(xIMPFreteNacEscopo)), Z) & xIMPFreteNacEscopo & String(3, Z) & Chr(27) & "!" & Chr(20) & xIMPNatureza
Linha(31) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(19 - Len(Trim("")), Z) & String(3, Z) & "" & String(20 - Len(Trim("")), Z)
Linha(32) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPTxDescrDevAg & String(19 - Len(Trim(xIMPTxDescrDevAg)), Z) & String(3, Z) & xIMPTxDescrDevTransp & String(20 - Len(Trim(xIMPTxDescrDevTransp)), Z)
Linha(33) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(19 - Len(Trim("")), Z) & String(3, Z) & "" & String(20 - Len(Trim("")), Z)
Linha(34) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(4 - Len(Trim(xIMPQteVol)), Z) & xIMPQteVol & String(2, Z) & String(7 - Len(Trim(xIMPPesoReal)), Z) & xIMPPesoReal & String(1, Z) & String(7 - Len(Trim(xIMPPesoTax)), Z) & xIMPPesoTax & String(1, Z) & xIMPTrecho & String(7 - Len(Trim(xIMPTrecho)), Z) & String(2, Z) & Trim(xIMPCl) & String(2, Z) & xIMPCodigo & String(4 - Len(Trim(xIMPCodigo)), Z) & String(1, Z) & String(6 - Len(Trim(xIMPKilo)), Z) & xIMPKilo & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao01 & String(60 - Len(Trim(xIMPStrObservacao01)), Z)
Linha(35) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao02 & String(60 - Len(Trim(xIMPStrObservacao02)), Z)
Linha(36) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao03 & String(60 - Len(Trim(xIMPStrObservacao03)), Z)
Linha(37) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao04 & String(60 - Len(Trim(xIMPStrObservacao04)), Z)
Linha(38) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & "" & String(40 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao05 & String(60 - Len(Trim(xIMPStrObservacao05)), Z)
Linha(39) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Trim(xPAGOFreteNacional)), Z) & xPAGOFreteNacional & String(16, Z) & String(14 - Len(Trim(xAPAGARFreteNacional)), Z) & xAPAGARFreteNacional & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(42 - Len(Trim("")), Z)
Linha(40) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(42 - Len(Trim("")), Z)
Linha(41) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Trim(xPAGOAdValorem)), Z) & xPAGOAdValorem & String(10, Z) & xIMPTipoADVAL & String(4 - Len(Trim(xIMPTipoADVAL)), Z) & String(2, Z) & String(14 - Len(Trim(xAPAGARAdValorem)), Z) & xAPAGARAdValorem & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(42 - Len(Trim("")), Z)
Linha(42) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(42 - Len(Trim("")), Z)
Linha(43) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Trim(xPAGOTxTerrOrig)), Z) & xPAGOTxTerrOrig & String(16, Z) & String(14 - Len(Trim(xAPAGARTxTerrOrig)), Z) & xAPAGARTxTerrOrig & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(44) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(42 - Len(Trim("")), Z)
Linha(45) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Trim(xPAGOTxTerrDest)), Z) & xPAGOTxTerrDest & String(16, Z) & String(14 - Len(Trim(xAPAGARTxTerrDest)), Z) & xAPAGARTxTerrDest & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(46) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(47) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Trim(xPAGOTxRedesp)), Z) & xPAGOTxRedesp & String(16, Z) & String(14 - Len(Trim(xAPAGARTxRedesp)), Z) & xAPAGARTxRedesp & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsSeguro & String(60 - Len(Trim(xIMPObsSeguro)), Z)
Linha(48) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsICMS & String(60 - Len(Trim(xIMPObsICMS)), Z)
Linha(49) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Trim(xPAGOTxAgente)), Z) & xPAGOTxAgente & String(16, Z) & String(14 - Len(Trim(xAPAGARTxAgente)), Z) & xAPAGARTxAgente & EspacosMeio & Chr(27) & "!" & Chr(25) & Mid(xIMPObsPerecivel, 1, 42) & String(42 - Len(Trim(Mid(xIMPObsPerecivel, 1, 42))), Z)
Linha(50) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(51) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Trim(xPAGOTxDevTransp)), Z) & xPAGOTxDevTransp & String(16, Z) & String(14 - Len(Trim(xAPAGARTxDevTransp)), Z) & xAPAGARTxDevTransp & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPStrLocalRetira & String(42 - Len(Trim(xIMPStrLocalRetira)), Z)
Linha(52) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(53) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxOutros1), Z) & xPAGOTxOutros1 & String(2, Z) & Mid(xIMPDescrTxOutros1, 1, 12) & String(12 - Len(Mid(xIMPDescrTxOutros1, 1, 12)), Z) & String(2, Z) & String(14 - Len(xAPAGARTxOutros1), Z) & xAPAGARTxOutros1 & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(54) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(2, Z) & "" & String(12 - Len(Trim("")), Z) & String(2, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(55) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxOutros2), Z) & xPAGOTxOutros2 & String(2, Z) & Mid(xIMPDescrTxOutros2, 1, 12) & String(12 - Len(Mid(xIMPDescrTxOutros2, 1, 12)), Z) & String(2, Z) & String(14 - Len(Trim(xAPAGARTxOutros2)), Z) & xAPAGARTxOutros2 & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(56) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(2, Z) & "" & String(12 - Len(Trim("")), Z) & String(2, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Trim("")), Z)
Linha(57) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteTotal), Z) & xPAGOFreteTotal & String(16, Z) & String(14 - Len(xAPAGARFreteTotal), Z) & xAPAGARFreteTotal & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Z) & String(18 - Len(Trim(xIMPStrTotalServ)), Z) & xIMPStrTotalServ
Linha(58) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Z) & String(18 - Len(Trim(xIMPStrBaseCalculo)), Z) & xIMPStrBaseCalculo
Linha(59) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Z) & xIMPStrAliquota & String(5 - Len(xIMPStrAliquota), Z) & String(7, Z) & String(18 - Len(Trim(xIMPStrICMS)), Z) & xIMPStrICMS
Linha(60) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Z) & "" & String(18 - Len(Trim("")), Z)
Linha(61) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPAgenteEmissor & String(32 - Len(Trim(xIMPAgenteEmissor)), Z) & String(3, Z) & Chr(27) & "!" & Chr(20) & xIMPCodIATA
Linha(62) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(32 - Len(Trim("")), Z) & String(3, Z) & "" & String(7 - Len(Trim("")), Z)
Linha(63) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPDtEmissao & String(32 - Len(Trim(xIMPDtEmissao)), Z) & String(2, Z) & xIMPHoraEmissao & String(8 - Len(Trim(xIMPHoraEmissao)), Z)
Linha(64) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(32 - Len(Trim("")), Z) & String(3, Z) & "" & String(7 - Len(Trim("")), Z)
Linha(65) = Chr(27) & "!" & Chr(25) & EspacosESQ & "" & String(15 - Len(Trim("")), Z) & String(16, Z) & "" & String(14 - Len(Trim("")), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & "" & String(21 - Len(Trim("")), Z) & String(2, Z) & "" & String(19 - Len(Trim("")), Z)
Linha(66) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPNaturezaOp & String(26 - Len(xIMPNaturezaOp), Z) & String(2, Z) & xIMPCFOP & String(7 - Len(xIMPCFOP), Z) & String(3, Z) & xIMPEmissor & String(29 - Len(xIMPEmissor), Z) & String(2, Z) & xIMPLocalidade & String(20 - Len(xIMPLocalidade), Z)








'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXX    XXXXXXXXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXX    XXXXXXX" & Chr(27) & "!" & Chr(20) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXX    XXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(72) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXXXXXXXXXXXXX    XXXXXXX" & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "     XXXXXXXXXXXXXXXXXX    XXXXXXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXX  XXXXXX   XXXXXXXXXXXXXXXXXXX   XXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "XXXX  XXXXXXX XXXXXXX  XXXXXX  X  XXXX XXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX  XXXXXX  XXXX  XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX  XXXXXXXXXXXX  XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX  XXXXXXXXXXXX  XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXX                XXXXXXXXXXXXXX" & Chr(27) & "!" & Chr(25) & "                        XXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                        XXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                        XXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX   XXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXX"
'Print #1, Chr(27) & "!" & Chr(25) & "                                             " & Chr(27) & "!" & Chr(25) & "                                          "
'Print #1, Chr(27) & "!" & Chr(25) & "XXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXX   XXXXXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXXX"
DoEvents
Open SETIMPImpressoraPadrao For Output As #1

    For e = 1 To 66
    Print #1, Linha(e)
    Next

Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""

Close #1
End Sub


Private Sub MascaraAWBVARIG()

Dim Linha(1 To 67) As String
Dim EspacosESQ As String
Dim EspacosMeio As String
Dim EspacosAUX As String

Dim Z As String

Dim xPAGOFreteNacional As String
Dim xPAGOFreteRegional As String
Dim xPAGOAdValorem As String
Dim xPAGOTxTerrOrig As String
Dim xPAGOTxTerrDest As String
Dim xPAGOTxRedesp As String
Dim xPAGOTxAgente As String
Dim xPAGOTxDevTransp As String
Dim xPAGOTxOutros1 As String
Dim xPAGOTxOutros2 As String
Dim xPAGOFreteTotal As String
Dim xAPAGARFreteNacional As String
Dim xAPAGARFreteRegional As String
Dim xAPAGARAdValorem As String
Dim xAPAGARTxTerrOrig As String
Dim xAPAGARTxTerrDest As String
Dim xAPAGARTxRedesp As String
Dim xAPAGARTxAgente As String
Dim xAPAGARTxDevTransp As String
Dim xAPAGARTxOutros1 As String
Dim xAPAGARTxOutros2 As String
Dim xAPAGARFreteTotal As String


StringNF = Trim(StringNF)
xIMPStrNF01 = Trim(xIMPStrNF01)
xIMPStrNF02 = Trim(xIMPStrNF02)
xIMPStrNF03 = Trim(xIMPStrNF03)
xIMPStrNF04 = Trim(xIMPStrNF04)
xIMPStrNF05 = Trim(xIMPStrNF05)
xIMPStrNF06 = Trim(xIMPStrNF06)
xIMPStrNF07 = Trim(xIMPStrNF07)
xIMPStrNF08 = Trim(xIMPStrNF08)
xIMPStrNF09 = Trim(xIMPStrNF09)
xIMPStrNF10 = Trim(xIMPStrNF10)
xIMPStrNF11 = Trim(xIMPStrNF11)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPNomeEXP = Trim(xIMPNomeEXP)
xIMPCGCEXP = Trim(xIMPCGCEXP)
xIMPInscEstEXP = Trim(xIMPInscEstEXP)
xIMPEndEXP = Trim(xIMPEndEXP)
xIMPBairroEXP = Trim(xIMPBairroEXP)
xIMPCidadeEXP = Trim(xIMPCidadeEXP)
xIMPCepEXP = Trim(xIMPCepEXP)
xIMPUFEXP = Trim(xIMPUFEXP)
xIMPTelEXP = Trim(xIMPTelEXP)
xIMPFAXEXP = Trim(xIMPFAXEXP)
xIMPNomeDEST = Trim(xIMPNomeDEST)
xIMPCGCDEST = Trim(xIMPCGCDEST)
xIMPInscEstDEST = Trim(xIMPInscEstDEST)
xIMPEndDEST = Trim(xIMPEndDEST)
xIMPBairroDEST = Trim(xIMPBairroDEST)
xIMPCidadeDEST = Trim(xIMPCidadeDEST)
xIMPCepDEST = Trim(xIMPCepDEST)
xIMPUFDEST = Trim(xIMPUFDEST)
xIMPTelDEST = Trim(xIMPTelDEST)
xIMPFAXDEST = Trim(xIMPFAXDEST)
xIMPOrigem = Trim(xIMPOrigem)
xIMPVia = Trim(xIMPVia)
xIMPCidadeDESTINO = Trim(xIMPCidadeDESTINO)
xIMPSIGLA = Trim(xIMPSIGLA)
xIMPReqTranspMinuta = Trim(xIMPReqTranspMinuta)
xIMPNumControle = Trim(xIMPNumControle)
xIMPInscrEstCiaAerea = Trim(xIMPInscrEstCiaAerea)
xIMPCNPJCiaAerea = Trim(xIMPCNPJCiaAerea)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPVlDecTRANSP = Trim(xIMPVlDecTRANSP)
xIMPVlDecSUFRAMA = Trim(xIMPVlDecSUFRAMA)
xIMPDescrEmbalagem = Trim(xIMPDescrEmbalagem)
xIMPQteVol = Trim(xIMPQteVol)
xIMPPesoReal = Trim(xIMPPesoReal)
xIMPPesoTax = Trim(xIMPPesoTax)
xIMPTrecho = Trim(xIMPTrecho)
xIMPCl = Trim(xIMPCl)
xIMPCodigo = Trim(xIMPCodigo)
xIMPKilo = Trim(xIMPKilo)
xIMPFreteNacEscopo = Trim(xIMPFreteNacEscopo)
xIMPNatureza = Trim(xIMPNatureza)
xIMPTxDescrDevAg = Trim(xIMPTxDescrDevAg)
xIMPTxDescrDevTransp = Trim(xIMPTxDescrDevTransp)
xIMPFreteNacional = Trim(xIMPFreteNacional)
xIMPFreteRegional = Trim(xIMPFreteRegional)
xIMPAdValorem = Trim(xIMPAdValorem)
xIMPTipoADVAL = Trim(xIMPTipoADVAL)
xIMPTxTerrOrig = Trim(xIMPTxTerrOrig)
xIMPTxTerrDest = Trim(xIMPTxTerrDest)
xIMPTxRedesp = Trim(xIMPTxRedesp)
xIMPTxAgente = Trim(xIMPTxAgente)
xIMPTxDevTransp = Trim(xIMPTxDevTransp)
xIMPDescrTxOutros1 = Trim(xIMPDescrTxOutros1)
xIMPTxOutros1 = Trim(xIMPTxOutros1)
xIMPDescrTxOutros2 = Trim(xIMPDescrTxOutros2)
xIMPTxOutros2 = Trim(xIMPTxOutros2)
xIMPFreteTotal = Trim(xIMPFreteTotal)
xIMPStrObservacao = Trim(xIMPStrObservacao)
xIMPStrObservacao01 = Trim(xIMPStrObservacao01)
xIMPStrObservacao02 = Trim(xIMPStrObservacao02)
xIMPStrObservacao03 = Trim(xIMPStrObservacao03)
xIMPStrObservacao04 = Trim(xIMPStrObservacao04)
xIMPObsICMS = Trim(xIMPObsICMS)
xIMPObsPerecivel = Trim(xIMPObsPerecivel)
xIMPObsSeguro = Trim(xIMPObsSeguro)
xIMPStrRetiraSIM = Trim(xIMPStrRetiraSIM)
xIMPStrRetiraNAO = Trim(xIMPStrRetiraNAO)
xIMPStrLocalRetira = Trim(xIMPStrLocalRetira)
xIMPHorarioAt = Trim(xIMPHorarioAt)
xIMPStrTelefone = Trim(xIMPStrTelefone)
xIMPStrTotalServ = Trim(xIMPStrTotalServ)
xIMPStrBaseCalculo = Trim(xIMPStrBaseCalculo)
xIMPStrAliquota = Trim(xIMPStrAliquota)
xIMPStrICMS = Trim(xIMPStrICMS)
xIMPAgenteEmissor = Trim(xIMPAgenteEmissor)
xIMPCodIATA = Trim(xIMPCodIATA)
'xIMPDtEmissao = DataHora("DATA")
'xIMPHoraEmissao = DataHora("HORA")
xIMPDtEmissao = Trim(xDataIMP)
xIMPHoraEmissao = Trim(xIMPHoraEmissao)
xIMPNaturezaOp = Trim(xIMPNaturezaOp)
xIMPCFOP = Trim(xIMPCFOP)
'xIMPEmissor = xUsuario
xIMPEmissor = Trim(xIMPEmissor)
xIMPLocalidade = Trim(xIMPLocalidade)
xIMPMatricula = Trim(xIMPMatricula)

xIMPStrNF01 = Mid(xIMPStrNF01, 1, 60)
xIMPStrNF02 = Mid(xIMPStrNF02, 1, 60)
xIMPStrNF03 = Mid(xIMPStrNF03, 1, 60)
xIMPStrNF04 = Mid(xIMPStrNF04, 1, 60)
xIMPStrNF05 = Mid(xIMPStrNF05, 1, 60)
xIMPStrNF06 = Mid(xIMPStrNF06, 1, 60)
xIMPStrNF07 = Mid(xIMPStrNF07, 1, 60)
xIMPStrNF08 = Mid(xIMPStrNF08, 1, 60)
xIMPStrNF09 = Mid(xIMPStrNF09, 1, 60)
xIMPStrNF10 = Mid(xIMPStrNF10, 1, 60)
xIMPStrNF11 = Mid(xIMPStrNF11, 1, 60)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 60)
xIMPNomeEXP = Mid(xIMPNomeEXP, 1, 40)
xIMPCGCEXP = Mid(xIMPCGCEXP, 1, 40)
xIMPInscEstEXP = Mid(xIMPInscEstEXP, 1, 40)
xIMPEndEXP = Mid(xIMPEndEXP, 1, 40)
xIMPBairroEXP = Mid(xIMPBairroEXP, 1, 19)
xIMPCidadeEXP = Mid(xIMPCidadeEXP, 1, 18)
xIMPCepEXP = Mid(xIMPCepEXP, 1, 15)
xIMPUFEXP = Mid(xIMPUFEXP, 1, 4)
xIMPTelEXP = Mid(xIMPTelEXP, 1, 14)
xIMPFAXEXP = Mid(xIMPFAXEXP, 1, 40)
xIMPNomeDEST = Mid(xIMPNomeDEST, 1, 40)
xIMPCGCDEST = Mid(xIMPCGCDEST, 1, 40)
xIMPInscEstDEST = Mid(xIMPInscEstDEST, 1, 40)
xIMPEndDEST = Mid(xIMPEndDEST, 1, 40)
xIMPBairroDEST = Mid(xIMPBairroDEST, 1, 19)
xIMPCidadeDEST = Mid(xIMPCidadeDEST, 1, 18)
xIMPCepDEST = Mid(xIMPCepDEST, 1, 15)
xIMPUFDEST = Mid(xIMPUFDEST, 1, 4)
xIMPTelDEST = Mid(xIMPTelDEST, 1, 14)
xIMPFAXDEST = Mid(xIMPFAXDEST, 1, 40)
xIMPOrigem = Mid(xIMPOrigem, 1, 8)
xIMPVia = Mid(xIMPVia, 1, 8)
xIMPCidadeDESTINO = Mid(xIMPCidadeDESTINO, 1, 18)
xIMPSIGLA = Mid(xIMPSIGLA, 1, 8)
xIMPReqTranspMinuta = Mid(xIMPReqTranspMinuta, 1, 15)
xIMPNumControle = Mid(xIMPNumControle, 1, 25)
xIMPInscrEstCiaAerea = Mid(xIMPInscrEstCiaAerea, 1, 20)
xIMPCNPJCiaAerea = Mid(xIMPCNPJCiaAerea, 1, 20)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 35)
xIMPVlDecTRANSP = Mid(xIMPVlDecTRANSP, 1, 28)
xIMPVlDecSUFRAMA = Mid(xIMPVlDecSUFRAMA, 1, 28)
xIMPDescrEmbalagem = Mid(xIMPDescrEmbalagem, 1, 60)
xIMPQteVol = Mid(xIMPQteVol, 1, 5)
xIMPPesoReal = Mid(xIMPPesoReal, 1, 7)
xIMPPesoTax = Mid(xIMPPesoTax, 1, 8)
xIMPTrecho = Mid(xIMPTrecho, 1, 8)
xIMPCl = Mid(xIMPCl, 1, 2)
xIMPCodigo = Mid(xIMPCodigo, 1, 3)
xIMPKilo = Mid(xIMPKilo, 1, 9)
xIMPFreteNacEscopo = Mid(xIMPFreteNacEscopo, 1, 13)
xIMPTxDescrDevAg = Mid(xIMPTxDescrDevAg, 1, 1)
xIMPTxDescrDevTransp = Mid(xIMPTxDescrDevTransp, 1, 1)
xIMPDescrTxOutros1 = Mid(xIMPDescrTxOutros1, 1, 12)
xIMPDescrTxOutros2 = Mid(xIMPDescrTxOutros2, 1, 12)
xIMPStrObservacao01 = Mid(xIMPStrObservacao01, 1, 60)
xIMPStrObservacao02 = Mid(xIMPStrObservacao02, 1, 60)
xIMPStrObservacao03 = Mid(xIMPStrObservacao03, 1, 60)
xIMPStrObservacao04 = Mid(xIMPStrObservacao04, 1, 60)
xIMPObsICMS = Mid(xIMPObsICMS, 1, 60)
xIMPObsSeguro = Mid(xIMPObsSeguro, 1, 60)
xIMPStrLocalRetira = Mid(xIMPStrLocalRetira, 1, 42)
xIMPAgenteEmissor = Mid(xIMPAgenteEmissor, 1, 28)
xIMPCodIATA = Mid(xIMPCodIATA, 1, 28)
xIMPDtEmissao = Mid(xIMPDtEmissao, 1, 22)
xIMPHoraEmissao = Mid(xIMPHoraEmissao, 1, 18)
xIMPNaturezaOp = Mid(xIMPNaturezaOp, 1, 19)
xIMPCFOP = Mid(xIMPCFOP, 1, 6)
xIMPEmissor = Mid(xIMPEmissor, 1, 42)
xIMPLocalidade = Mid(xIMPLocalidade, 1, 6)




'65 LINHAS - 9 LINHAS EM BRANCO NO COMECO

EspacosESQ = String(3, " ")
EspacosMeio = String(2, " ")
EspacosAUX = String(5, " ")
Z = " "
Y = " "

If OptPago.Value = True Then
    xPAGOFreteNacional = xIMPFreteNacional
    xPAGOFreteRegional = xIMPFreteRegional
    xPAGOAdValorem = xIMPAdValorem
    xPAGOTxTerrOrig = xIMPTxTerrOrig
    xPAGOTxTerrDest = xIMPTxTerrDest
    xPAGOTxRedesp = xIMPTxRedesp
    xPAGOTxAgente = xIMPTxAgente
    xPAGOTxDevTransp = xIMPTxDevTransp
    xPAGOTxOutros1 = xIMPTxOutros1
    xPAGOTxOutros2 = xIMPTxOutros2
    xPAGOFreteTotal = xIMPFreteTotal
    xAPAGARFreteNacional = ""
    xAPAGARFreteRegional = ""
    xAPAGARAdValorem = ""
    xAPAGARTxTerrOrig = ""
    xAPAGARTxTerrDest = ""
    xAPAGARTxRedesp = ""
    xAPAGARTxAgente = ""
    'xAPAGARTxDevTransp = ""
    'xAPAGARTxDevTransp = xIMPTxDevTransp
    xAPAGARTxOutros1 = ""
    xAPAGARTxOutros2 = ""
    xAPAGARFreteTotal = ""
Else
    xPAGOFreteNacional = ""
    xPAGOFreteRegional = ""
    xPAGOAdValorem = ""
    xPAGOTxTerrOrig = ""
    xPAGOTxTerrDest = ""
    xPAGOTxRedesp = ""
    xPAGOTxAgente = ""
    xPAGOTxDevTransp = ""
    xPAGOTxOutros1 = ""
    xPAGOTxOutros2 = ""
    xPAGOFreteTotal = ""
    xAPAGARFreteNacional = xIMPFreteNacional
    xAPAGARFreteRegional = xIMPFreteRegional
    xAPAGARAdValorem = xIMPAdValorem
    xAPAGARTxTerrOrig = xIMPTxTerrOrig
    xAPAGARTxTerrDest = xIMPTxTerrDest
    xAPAGARTxRedesp = xIMPTxRedesp
    xAPAGARTxAgente = xIMPTxAgente
    xAPAGARTxDevTransp = xIMPTxDevTransp
    xAPAGARTxOutros1 = xIMPTxOutros1
    xAPAGARTxOutros2 = xIMPTxOutros2
    xAPAGARFreteTotal = xIMPFreteTotal
End If


Linha(1) = ""
Linha(2) = ""
Linha(3) = ""
Linha(4) = ""
Linha(5) = ""
Linha(6) = ""
Linha(7) = ""
Linha(8) = ""
Linha(9) = ""
Linha(10) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF01 & String(60 - Len(xIMPStrNF01), Z)
Linha(11) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF02 & String(60 - Len(xIMPStrNF02), Z)
Linha(12) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeEXP & String(40 - Len(xIMPNomeEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF03 & String(60 - Len(xIMPStrNF03), Z)
Linha(13) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCEXP & String(40 - Len(xIMPCGCEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF04 & String(60 - Len(xIMPStrNF04), Z)
Linha(14) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstEXP & String(40 - Len(xIMPInscEstEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF05 & String(60 - Len(xIMPStrNF05), Z)
Linha(15) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndEXP & String(40 - Len(xIMPEndEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF06 & String(60 - Len(xIMPStrNF06), Z)
Linha(16) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroEXP & String(19 - Len(xIMPBairroEXP), Z) & String(3, Y) & xIMPCidadeEXP & String(18 - Len(xIMPCidadeEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF07 & String(60 - Len(xIMPStrNF07), Z)
Linha(17) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCepEXP & String(15 - Len(xIMPCepEXP), Z) & String(3, Y) & xIMPUFEXP & String(4 - Len(xIMPUFEXP), Z) & String(4, Y) & xIMPTelEXP & String(14 - Len(xIMPTelEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF08 & String(60 - Len(xIMPStrNF08), Z)
Linha(18) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPFAXEXP & String(40 - Len(xIMPFAXEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF09 & String(60 - Len(xIMPStrNF09), Z)
Linha(19) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF10 & String(60 - Len(xIMPStrNF10), Z)
Linha(20) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF11 & String(60 - Len(xIMPStrNF11), Z)
Linha(21) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Z & String(60 - Len(Z), Z)
Linha(22) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeDEST & String(40 - Len(xIMPNomeDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(72) & xIMPStrNF12 & String(35 - Len(xIMPStrNF12), Z)
Linha(23) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCDEST & String(40 - Len(xIMPCGCDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(24) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstDEST & String(40 - Len(xIMPInscEstDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPReqTranspMinuta & String(15 - Len(xIMPReqTranspMinuta), Z) & String(2, Y) & xIMPNumControle & String(25 - Len(xIMPNumControle), Z)
Linha(25) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndDEST & String(40 - Len(xIMPEndDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(26) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroDEST & String(19 - Len(xIMPBairroDEST), Z) & String(3, Y) & Mid(xIMPCidadeDEST, 1, 18) & String(18 - Len(Mid(xIMPCidadeDEST, 1, 18)), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPInscrEstCiaAerea & String(20 - Len(xIMPInscrEstCiaAerea), Z) & String(2, Y) & xIMPCNPJCiaAerea & String(20 - Len(xIMPCNPJCiaAerea), Z)
Linha(27) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCepDEST & String(15 - Len(xIMPCepDEST), Z) & String(3, Y) & xIMPUFDEST & String(4 - Len(xIMPUFDEST), Z) & String(4, Y) & xIMPTelDEST & String(14 - Len(xIMPTelDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(28) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPFAXDEST & String(40 - Len(xIMPFAXDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPVlDecTRANSP & String(28 - Len(xIMPVlDecTRANSP), Z) & String(2, Y) & xIMPVlDecSUFRAMA & String(28 - Len(xIMPVlDecSUFRAMA), Z)
Linha(29) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(30) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPOrigem & String(8 - Len(xIMPOrigem), Z) & String(1, Y) & xIMPVia & String(8 - Len(xIMPVia), Z) & String(1, Y) & Mid(xIMPCidadeDESTINO, 1, 18) & String(18 - Len(Mid(xIMPCidadeDESTINO, 1, 18)), Z) & String(1, Y) & xIMPSIGLA & String(8 - Len(xIMPSIGLA), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPDescrEmbalagem & String(60 - Len(xIMPDescrEmbalagem), Z)
Linha(31) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(32) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(33) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(34) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(5 - Len(xIMPQteVol), Z) & xIMPQteVol & String(2, Y) & String(7 - Len(xIMPPesoReal), Z) & xIMPPesoReal & String(1, Y) & String(8 - Len(xIMPPesoTax), Z) & xIMPPesoTax & String(2, Y) & xIMPTrecho & String(8 - Len(xIMPTrecho), Z) & String(1, Y) & xIMPCl & String(2 - Len(xIMPCl), Z) & String(2, Y) & xIMPCodigo & String(3 - Len(xIMPCodigo), Z) & String(2, Y) & String(9 - Len(xIMPKilo), Z) & xIMPKilo & String(1, Y) & String(13 - Len(xIMPFreteNacEscopo), Z) & xIMPFreteNacEscopo & String(2, Y) & Chr(27) & "!" & Chr(20) & xIMPNatureza
Linha(35) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(36) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(37) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(38) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(39) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteNacional), Z) & xPAGOFreteNacional & String(16, Y) & String(14 - Len(xAPAGARFreteNacional), Z) & xAPAGARFreteNacional & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao01 & String(60 - Len(xIMPStrObservacao01), Z)
Linha(40) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao02 & String(60 - Len(xIMPStrObservacao02), Z)
Linha(41) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteRegional), Z) & xPAGOFreteRegional & String(16, Y) & String(14 - Len(xAPAGARFreteRegional), Z) & xAPAGARFreteRegional & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao03 & String(60 - Len(xIMPStrObservacao03), Z)
Linha(42) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao04 & String(60 - Len(xIMPStrObservacao04), Z)
Linha(43) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOAdValorem), Z) & xPAGOAdValorem & String(10, Y) & String(4 - Len(xIMPTipoADVAL), Z) & xIMPTipoADVAL & String(2, Y) & xAPAGARAdValorem & String(14 - Len(xAPAGARAdValorem), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsSeguro & String(60 - Len(xIMPObsSeguro), Z)
Linha(44) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsICMS & String(60 - Len(xIMPObsICMS), Z)
Linha(45) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxTerrOrig), Z) & xPAGOTxTerrOrig & String(16, Y) & String(14 - Len(xAPAGARTxTerrOrig), Z) & xAPAGARTxTerrOrig & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPObsPerecivel
Linha(46) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(17, Y) & Y & String(5, Y) & String(3, Y)
Linha(47) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxTerrDest), Z) & xPAGOTxTerrDest & String(16, Y) & String(14 - Len(xAPAGARTxTerrDest), Z) & xAPAGARTxTerrDest & EspacosMeio & Chr(27) & "!" & Chr(25) & String(17, Y) & xIMPStrRetiraSIM & String(3 - Len(xIMPStrRetiraSIM), Z) & String(5, Y) & xIMPStrRetiraNAO & String(3 - Len(xIMPStrRetiraNAO), Z)
Linha(48) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(42 - Len(Y), Y)
Linha(49) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxRedesp), Z) & xPAGOTxRedesp & String(16, Y) & String(14 - Len(xAPAGARTxRedesp), Z) & xAPAGARTxRedesp & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPStrLocalRetira & String(42 - Len(xIMPStrLocalRetira), Z)
Linha(50) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(51) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxAgente), Z) & xPAGOTxAgente & String(16, Y) & String(14 - Len(xAPAGARTxAgente), Z) & xAPAGARTxAgente & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(52) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(53) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxDevTransp), Z) & xPAGOTxDevTransp & String(16, Y) & String(14 - Len(xAPAGARTxDevTransp), Z) & xAPAGARTxDevTransp & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(54) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(55) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxOutros1), Z) & xPAGOTxOutros1 & String(2, Y) & xIMPDescrTxOutros1 & String(12 - Len(xIMPDescrTxOutros1), Z) & String(2, Y) & String(14 - Len(xAPAGARTxOutros1), Z) & xAPAGARTxOutros1 & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Y) & String(18 - Len(xIMPStrTotalServ), Z) & xIMPStrTotalServ
Linha(56) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(42 - Len(Y), Y)
Linha(57) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxOutros2), Z) & xPAGOTxOutros2 & String(2, Y) & xIMPDescrTxOutros2 & String(12 - Len(xIMPDescrTxOutros2), Z) & String(2, Y) & String(14 - Len(xAPAGARTxOutros2), Z) & xAPAGARTxOutros2 & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Y) & String(18 - Len(xIMPStrBaseCalculo), Z) & xIMPStrBaseCalculo
Linha(58) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(42 - Len(Y), Y)
Linha(59) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteTotal), Z) & xPAGOFreteTotal & String(16, Y) & String(14 - Len(xAPAGARFreteTotal), Z) & xAPAGARFreteTotal & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Z) & xIMPStrAliquota & String(5 - Len(xIMPStrAliquota), Z) & String(7, Y) & String(18 - Len(Trim(xIMPStrICMS)), Z) & xIMPStrICMS
Linha(60) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(61) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPAgenteEmissor & String(28 - Len(xIMPAgenteEmissor), Z) & String(3, Y) & xIMPCodIATA & String(28 - Len(xIMPCodIATA), Z)
Linha(62) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(63) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPDtEmissao & String(22 - Len(xIMPDtEmissao), Z) & String(3, Y) & xIMPHoraEmissao & String(18 - Len(xIMPHoraEmissao), Z)
Linha(64) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(65) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(65) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPNaturezaOp & String(19 - Len(xIMPNaturezaOp), Z) & String(1, Y) & xIMPCFOP & String(6 - Len(xIMPCFOP), Z) & String(2, Y) & xIMPEmissor & String(42 - Len(xIMPEmissor), Z) & String(1, Y) & xIMPLocalidade & String(6 - Len(xIMPLocalidade), Z) & String(1, Y) & Z & String(11 - Len(Z), Z)

DoEvents
Open SETIMPImpressoraPadrao For Output As #1
    For e = 1 To 66
    Print #1, Linha(e)
    Next
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""

Close #1
End Sub
Private Sub MascaraAWBVASP()

Dim Linha(1 To 67) As String
Dim EspacosESQ As String
Dim EspacosMeio As String
Dim EspacosAUX As String

Dim Z As String

Dim xPAGOFreteNacional As String
Dim xPAGOFreteRegional As String
Dim xPAGOAdValorem As String
Dim xPAGOTxTerrOrig As String
Dim xPAGOTxTerrDest As String
Dim xPAGOTxRedesp As String
Dim xPAGOTxAgente As String
Dim xPAGOTxDevTransp As String
Dim xPAGOTxOutros1 As String
Dim xPAGOTxOutros2 As String
Dim xPAGOFreteTotal As String
Dim xAPAGARFreteNacional As String
Dim xAPAGARFreteRegional As String
Dim xAPAGARAdValorem As String
Dim xAPAGARTxTerrOrig As String
Dim xAPAGARTxTerrDest As String
Dim xAPAGARTxRedesp As String
Dim xAPAGARTxAgente As String
Dim xAPAGARTxDevTransp As String
Dim xAPAGARTxOutros1 As String
Dim xAPAGARTxOutros2 As String
Dim xAPAGARFreteTotal As String


StringNF = Trim(StringNF)
xIMPStrNF01 = Trim(xIMPStrNF01)
xIMPStrNF02 = Trim(xIMPStrNF02)
xIMPStrNF03 = Trim(xIMPStrNF03)
xIMPStrNF04 = Trim(xIMPStrNF04)
xIMPStrNF05 = Trim(xIMPStrNF05)
xIMPStrNF06 = Trim(xIMPStrNF06)
xIMPStrNF07 = Trim(xIMPStrNF07)
xIMPStrNF08 = Trim(xIMPStrNF08)
xIMPStrNF09 = Trim(xIMPStrNF09)
xIMPStrNF10 = Trim(xIMPStrNF10)
xIMPStrNF11 = Trim(xIMPStrNF11)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPNomeEXP = Trim(xIMPNomeEXP)
xIMPCGCEXP = Trim(xIMPCGCEXP)
xIMPInscEstEXP = Trim(xIMPInscEstEXP)
xIMPEndEXP = Trim(xIMPEndEXP)
xIMPBairroEXP = Trim(xIMPBairroEXP)
xIMPCidadeEXP = Trim(xIMPCidadeEXP)
xIMPCepEXP = Trim(xIMPCepEXP)
xIMPUFEXP = Trim(xIMPUFEXP)
xIMPTelEXP = Trim(xIMPTelEXP)
xIMPFAXEXP = Trim(xIMPFAXEXP)
xIMPNomeDEST = Trim(xIMPNomeDEST)
xIMPCGCDEST = Trim(xIMPCGCDEST)
xIMPInscEstDEST = Trim(xIMPInscEstDEST)
xIMPEndDEST = Trim(xIMPEndDEST)
xIMPBairroDEST = Trim(xIMPBairroDEST)
xIMPCidadeDEST = Trim(xIMPCidadeDEST)
xIMPCepDEST = Trim(xIMPCepDEST)
xIMPUFDEST = Trim(xIMPUFDEST)
xIMPTelDEST = Trim(xIMPTelDEST)
xIMPFAXDEST = Trim(xIMPFAXDEST)
xIMPOrigem = Trim(xIMPOrigem)
xIMPVia = Trim(xIMPVia)
xIMPCidadeDESTINO = Trim(xIMPCidadeDESTINO)
xIMPSIGLA = Trim(xIMPSIGLA)
xIMPReqTranspMinuta = Trim(xIMPReqTranspMinuta)
xIMPNumControle = Trim(xIMPNumControle)
xIMPInscrEstCiaAerea = Trim(xIMPInscrEstCiaAerea)
xIMPCNPJCiaAerea = Trim(xIMPCNPJCiaAerea)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPVlDecTRANSP = Trim(xIMPVlDecTRANSP)
xIMPVlDecSUFRAMA = Trim(xIMPVlDecSUFRAMA)
xIMPDescrEmbalagem = Trim(xIMPDescrEmbalagem)
xIMPQteVol = Trim(xIMPQteVol)
xIMPPesoReal = Trim(xIMPPesoReal)
xIMPPesoTax = Trim(xIMPPesoTax)
xIMPTrecho = Trim(xIMPTrecho)
xIMPCl = Trim(xIMPCl)
xIMPCodigo = Trim(xIMPCodigo)
xIMPKilo = Trim(xIMPKilo)
xIMPFreteNacEscopo = Trim(xIMPFreteNacEscopo)
xIMPNatureza = Trim(xIMPNatureza)
xIMPTxDescrDevAg = Trim(xIMPTxDescrDevAg)
xIMPTxDescrDevTransp = Trim(xIMPTxDescrDevTransp)
xIMPFreteNacional = Trim(xIMPFreteNacional)
xIMPFreteRegional = Trim(xIMPFreteRegional)
xIMPAdValorem = Trim(xIMPAdValorem)
xIMPTipoADVAL = Trim(xIMPTipoADVAL)
xIMPTxTerrOrig = Trim(xIMPTxTerrOrig)
xIMPTxTerrDest = Trim(xIMPTxTerrDest)
xIMPTxRedesp = Trim(xIMPTxRedesp)
xIMPTxAgente = Trim(xIMPTxAgente)
xIMPTxDevTransp = Trim(xIMPTxDevTransp)
xIMPDescrTxOutros1 = Trim(xIMPDescrTxOutros1)
xIMPTxOutros1 = Trim(xIMPTxOutros1)
xIMPDescrTxOutros2 = Trim(xIMPDescrTxOutros2)
xIMPTxOutros2 = Trim(xIMPTxOutros2)
xIMPFreteTotal = Trim(xIMPFreteTotal)
xIMPStrObservacao = Trim(xIMPStrObservacao)
xIMPStrObservacao01 = Trim(xIMPStrObservacao01)
xIMPStrObservacao02 = Trim(xIMPStrObservacao02)
xIMPStrObservacao03 = Trim(xIMPStrObservacao03)
xIMPStrObservacao04 = Trim(xIMPStrObservacao04)
xIMPObsICMS = Trim(xIMPObsICMS)
xIMPObsPerecivel = Trim(xIMPObsPerecivel)
xIMPObsSeguro = Trim(xIMPObsSeguro)
xIMPStrRetiraSIM = Trim(xIMPStrRetiraSIM)
xIMPStrRetiraNAO = Trim(xIMPStrRetiraNAO)
xIMPStrLocalRetira = Trim(xIMPStrLocalRetira)
xIMPHorarioAt = Trim(xIMPHorarioAt)
xIMPStrTelefone = Trim(xIMPStrTelefone)
xIMPStrTotalServ = Trim(xIMPStrTotalServ)
xIMPStrBaseCalculo = Trim(xIMPStrBaseCalculo)
xIMPStrAliquota = Trim(xIMPStrAliquota)
xIMPStrICMS = Trim(xIMPStrICMS)
xIMPAgenteEmissor = Trim(xIMPAgenteEmissor)
xIMPCodIATA = Trim(xIMPCodIATA)
'xIMPDtEmissao = DataHora("DATA")
'xIMPHoraEmissao = DataHora("HORA")
xIMPDtEmissao = Trim(xDataIMP)
xIMPHoraEmissao = Trim(xIMPHoraEmissao)
xIMPNaturezaOp = Trim(xIMPNaturezaOp)
xIMPCFOP = Trim(xIMPCFOP)
'xIMPEmissor = xUsuario
xIMPEmissor = Trim(xIMPEmissor)
xIMPLocalidade = Trim(xIMPLocalidade)
xIMPMatricula = Trim(xIMPMatricula)

xIMPStrNF01 = Mid(xIMPStrNF01, 1, 60)
xIMPStrNF02 = Mid(xIMPStrNF02, 1, 60)
xIMPStrNF03 = Mid(xIMPStrNF03, 1, 60)
xIMPStrNF04 = Mid(xIMPStrNF04, 1, 60)
xIMPStrNF05 = Mid(xIMPStrNF05, 1, 60)
xIMPStrNF06 = Mid(xIMPStrNF06, 1, 60)
xIMPStrNF07 = Mid(xIMPStrNF07, 1, 60)
xIMPStrNF08 = Mid(xIMPStrNF08, 1, 60)
xIMPStrNF09 = Mid(xIMPStrNF09, 1, 60)
xIMPStrNF10 = Mid(xIMPStrNF10, 1, 60)
xIMPStrNF11 = Mid(xIMPStrNF11, 1, 60)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 60)
xIMPNomeEXP = Mid(xIMPNomeEXP, 1, 40)
xIMPCGCEXP = Mid(xIMPCGCEXP, 1, 40)
xIMPInscEstEXP = Mid(xIMPInscEstEXP, 1, 40)
xIMPEndEXP = Mid(xIMPEndEXP, 1, 40)
xIMPBairroEXP = Mid(xIMPBairroEXP, 1, 23)
xIMPCidadeEXP = Mid(xIMPCidadeEXP, 1, 29)
xIMPCepEXP = Mid(xIMPCepEXP, 1, 15)
xIMPUFEXP = Mid(xIMPUFEXP, 1, 8)
xIMPTelEXP = Mid(xIMPTelEXP, 1, 18)
xIMPFAXEXP = Mid(xIMPFAXEXP, 1, 20)
xIMPNomeDEST = Mid(xIMPNomeDEST, 1, 40)
xIMPCGCDEST = Mid(xIMPCGCDEST, 1, 40)
xIMPInscEstDEST = Mid(xIMPInscEstDEST, 1, 40)
xIMPEndDEST = Mid(xIMPEndDEST, 1, 40)
xIMPBairroDEST = Mid(xIMPBairroDEST, 1, 23)
xIMPCidadeDEST = Mid(xIMPCidadeDEST, 1, 29)
xIMPCepDEST = Mid(xIMPCepDEST, 1, 15)
xIMPUFDEST = Mid(xIMPUFDEST, 1, 8)
xIMPTelDEST = Mid(xIMPTelDEST, 1, 18)
xIMPFAXDEST = Mid(xIMPFAXDEST, 1, 20)
xIMPOrigem = Mid(xIMPOrigem, 1, 8)
xIMPVia = Mid(xIMPVia, 1, 8)
xIMPCidadeDESTINO = Mid(xIMPCidadeDESTINO, 1, 18)
xIMPSIGLA = Mid(xIMPSIGLA, 1, 8)
xIMPReqTranspMinuta = Mid(xIMPReqTranspMinuta, 1, 15)
xIMPNumControle = Mid(xIMPNumControle, 1, 25)
xIMPInscrEstCiaAerea = Mid(xIMPInscrEstCiaAerea, 1, 20)
xIMPCNPJCiaAerea = Mid(xIMPCNPJCiaAerea, 1, 20)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 35)
xIMPVlDecTRANSP = Mid(xIMPVlDecTRANSP, 1, 28)
xIMPVlDecSUFRAMA = Mid(xIMPVlDecSUFRAMA, 1, 28)
xIMPDescrEmbalagem = Mid(xIMPDescrEmbalagem, 1, 60)
xIMPQteVol = Mid(xIMPQteVol, 1, 5)
xIMPPesoReal = Mid(xIMPPesoReal, 1, 7)
xIMPPesoTax = Mid(xIMPPesoTax, 1, 8)
xIMPTrecho = Mid(xIMPTrecho, 1, 8)
xIMPCl = Mid(xIMPCl, 1, 2)
xIMPCodigo = Mid(xIMPCodigo, 1, 3)
xIMPKilo = Mid(xIMPKilo, 1, 9)
xIMPFreteNacEscopo = Mid(xIMPFreteNacEscopo, 1, 13)
xIMPTxDescrDevAg = Mid(xIMPTxDescrDevAg, 1, 1)
xIMPTxDescrDevTransp = Mid(xIMPTxDescrDevTransp, 1, 1)
xIMPDescrTxOutros1 = Mid(xIMPDescrTxOutros1, 1, 12)
xIMPDescrTxOutros2 = Mid(xIMPDescrTxOutros2, 1, 12)
xIMPStrObservacao01 = Mid(xIMPStrObservacao01, 1, 60)
xIMPStrObservacao02 = Mid(xIMPStrObservacao02, 1, 60)
xIMPStrObservacao03 = Mid(xIMPStrObservacao03, 1, 60)
xIMPStrObservacao04 = Mid(xIMPStrObservacao04, 1, 60)
xIMPObsICMS = Mid(xIMPObsICMS, 1, 60)
xIMPObsSeguro = Mid(xIMPObsSeguro, 1, 60)
xIMPStrLocalRetira = Mid(xIMPStrLocalRetira, 1, 42)
xIMPAgenteEmissor = Mid(xIMPAgenteEmissor, 1, 28)
xIMPCodIATA = Mid(xIMPCodIATA, 1, 28)
xIMPDtEmissao = Mid(xIMPDtEmissao, 1, 22)
xIMPHoraEmissao = Mid(xIMPHoraEmissao, 1, 18)
xIMPNaturezaOp = Mid(xIMPNaturezaOp, 1, 19)
xIMPCFOP = Mid(xIMPCFOP, 1, 6)
xIMPEmissor = Mid(xIMPEmissor, 1, 42)
xIMPLocalidade = Mid(xIMPLocalidade, 1, 6)





'65 LINHAS - 9 LINHAS EM BRANCO NO COMECO

EspacosESQ = String(3, " ")
EspacosMeio = String(2, " ")
EspacosAUX = String(5, " ")
Z = " "
Y = " "

If OptPago.Value = True Then
xPAGOFreteNacional = xIMPFreteNacional
xPAGOFreteRegional = xIMPFreteRegional
xPAGOAdValorem = xIMPAdValorem
xPAGOTxTerrOrig = xIMPTxTerrOrig
xPAGOTxTerrDest = xIMPTxTerrDest
xPAGOTxRedesp = xIMPTxRedesp
xPAGOTxAgente = xIMPTxAgente
xPAGOTxDevTransp = xIMPTxDevTransp
xPAGOTxOutros1 = xIMPTxOutros1
xPAGOTxOutros2 = xIMPTxOutros2
xPAGOFreteTotal = xIMPFreteTotal
xAPAGARFreteNacional = ""
xAPAGARFreteRegional = ""
xAPAGARAdValorem = ""
xAPAGARTxTerrOrig = ""
xAPAGARTxTerrDest = ""
xAPAGARTxRedesp = ""
xAPAGARTxAgente = ""
xAPAGARTxDevTransp = ""
xAPAGARTxOutros1 = ""
xAPAGARTxOutros2 = ""
xAPAGARFreteTotal = ""
Else
xPAGOFreteNacional = ""
xPAGOFreteRegional = ""
xPAGOAdValorem = ""
xPAGOTxTerrOrig = ""
xPAGOTxTerrDest = ""
xPAGOTxRedesp = ""
xPAGOTxAgente = ""
xPAGOTxDevTransp = ""
xPAGOTxOutros1 = ""
xPAGOTxOutros2 = ""
xPAGOFreteTotal = ""
xAPAGARFreteNacional = xIMPFreteNacional
xAPAGARFreteRegional = xIMPFreteRegional
xAPAGARAdValorem = xIMPAdValorem
xAPAGARTxTerrOrig = xIMPTxTerrOrig
xAPAGARTxTerrDest = xIMPTxTerrDest
xAPAGARTxRedesp = xIMPTxRedesp
xAPAGARTxAgente = xIMPTxAgente
xAPAGARTxDevTransp = xIMPTxDevTransp
xAPAGARTxOutros1 = xIMPTxOutros1
xAPAGARTxOutros2 = xIMPTxOutros2
xAPAGARFreteTotal = xIMPFreteTotal
End If


Linha(1) = ""
Linha(2) = ""
Linha(3) = ""
Linha(4) = ""
Linha(5) = ""
Linha(6) = ""
Linha(7) = ""
Linha(8) = ""
Linha(9) = ""
Linha(10) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF01 & String(60 - Len(xIMPStrNF01), Z)
Linha(11) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF02 & String(60 - Len(xIMPStrNF02), Z)
Linha(12) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeEXP & String(40 - Len(xIMPNomeEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF03 & String(60 - Len(xIMPStrNF03), Z)
Linha(13) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCEXP & String(40 - Len(xIMPCGCEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF04 & String(60 - Len(xIMPStrNF04), Z)
Linha(14) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstEXP & String(40 - Len(xIMPInscEstEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF05 & String(60 - Len(xIMPStrNF05), Z)
Linha(15) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndEXP & String(40 - Len(xIMPEndEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF06 & String(60 - Len(xIMPStrNF06), Z)
Linha(16) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroEXP & String(23 - Len(xIMPBairroEXP), Z) & String(2, Y) & xIMPCepEXP & String(15 - Len(xIMPCepEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF07 & String(60 - Len(xIMPStrNF07), Z)
Linha(17) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCidadeEXP & String(29 - Len(xIMPCidadeEXP), Z) & String(3, Y) & xIMPUFEXP & String(8 - Len(xIMPUFEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF08 & String(60 - Len(xIMPStrNF08), Z)
Linha(18) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPTelEXP & String(18 - Len(xIMPTelEXP), Z) & String(2, Y) & xIMPFAXEXP & String(20 - Len(xIMPFAXEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF09 & String(60 - Len(xIMPStrNF09), Z)
Linha(19) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF10 & String(60 - Len(xIMPStrNF10), Z)
Linha(20) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF11 & String(60 - Len(xIMPStrNF11), Z)
Linha(21) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Z & String(60 - Len(Z), Z)
Linha(22) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeDEST & String(40 - Len(xIMPNomeDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(72) & xIMPStrNF12 & String(35 - Len(xIMPStrNF12), Z)
Linha(23) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCDEST & String(40 - Len(xIMPCGCDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(24) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstDEST & String(40 - Len(xIMPInscEstDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPReqTranspMinuta & String(20 - Len(xIMPReqTranspMinuta), Z) & String(2, Y) & xIMPNumControle & String(20 - Len(xIMPNumControle), Z)
Linha(25) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndDEST & String(40 - Len(xIMPEndDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(26) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroDEST & String(23 - Len(xIMPBairroDEST), Z) & String(2, Y) & xIMPCepDEST & String(15 - Len(xIMPCepDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPInscrEstCiaAerea & String(20 - Len(xIMPInscrEstCiaAerea), Z) & String(2, Y) & xIMPCNPJCiaAerea & String(20 - Len(xIMPCNPJCiaAerea), Z)
Linha(27) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCidadeDEST & String(29 - Len(xIMPCidadeDEST), Z) & String(3, Y) & xIMPUFDEST & String(8 - Len(xIMPUFDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(28) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPTelDEST & String(18 - Len(xIMPTelDEST), Z) & String(2, Y) & xIMPFAXDEST & String(20 - Len(xIMPFAXDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPVlDecTRANSP & String(28 - Len(xIMPVlDecTRANSP), Z) & String(2, Y) & xIMPVlDecSUFRAMA & String(28 - Len(xIMPVlDecSUFRAMA), Z)
Linha(29) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(30) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPOrigem & String(8 - Len(xIMPOrigem), Z) & String(1, Y) & xIMPVia & String(8 - Len(xIMPVia), Z) & String(1, Y) & xIMPCidadeDESTINO & String(18 - Len(xIMPCidadeDESTINO), Z) & String(1, Y) & xIMPSIGLA & String(8 - Len(xIMPSIGLA), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPDescrEmbalagem & String(60 - Len(xIMPDescrEmbalagem), Z)
Linha(31) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(32) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(33) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(34) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(5 - Len(xIMPQteVol), Z) & xIMPQteVol & String(2, Y) & String(7 - Len(xIMPPesoReal), Z) & xIMPPesoReal & String(1, Y) & String(8 - Len(xIMPPesoTax), Z) & xIMPPesoTax & String(2, Y) & xIMPTrecho & String(8 - Len(xIMPTrecho), Z) & String(1, Y) & xIMPCl & String(2 - Len(xIMPCl), Z) & String(2, Y) & xIMPCodigo & String(3 - Len(xIMPCodigo), Z) & String(2, Y) & String(9 - Len(xIMPKilo), Z) & xIMPKilo & String(1, Y) & String(13 - Len(xIMPFreteNacEscopo), Z) & xIMPFreteNacEscopo & String(2, Y) & Chr(27) & "!" & Chr(20) & xIMPNatureza
Linha(35) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(36) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(37) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(38) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(39) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteNacional), Z) & xPAGOFreteNacional & String(16, Y) & String(14 - Len(xAPAGARFreteNacional), Z) & xAPAGARFreteNacional & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao01 & String(60 - Len(xIMPStrObservacao01), Z)
Linha(40) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao02 & String(60 - Len(xIMPStrObservacao02), Z)
Linha(41) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOAdValorem), Z) & xPAGOAdValorem & String(10, Y) & String(4 - Len(xIMPTipoADVAL), Z) & xIMPTipoADVAL & String(2, Y) & xAPAGARAdValorem & String(14 - Len(xAPAGARAdValorem), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao03 & String(60 - Len(xIMPStrObservacao03), Z)
Linha(42) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao04 & String(60 - Len(xIMPStrObservacao04), Z)
Linha(43) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxTerrOrig), Z) & xPAGOTxTerrOrig & String(16, Y) & String(14 - Len(xAPAGARTxTerrOrig), Z) & xAPAGARTxTerrOrig & EspacosMeio & Chr(27) & "!" & Chr(20)
Linha(44) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20)
Linha(45) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxTerrDest), Z) & xPAGOTxTerrDest & String(16, Y) & String(14 - Len(xAPAGARTxTerrDest), Z) & xAPAGARTxTerrDest & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsSeguro & String(60 - Len(xIMPObsSeguro), Z)
Linha(46) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsICMS & String(60 - Len(xIMPObsICMS), Z)
Linha(47) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxRedesp), Z) & xPAGOTxRedesp & String(16, Y) & String(14 - Len(xAPAGARTxRedesp), Z) & xAPAGARTxRedesp & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPObsPerecivel
Linha(48) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(42 - Len(Y), Y)
Linha(49) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxAgente), Z) & xPAGOTxAgente & String(16, Y) & String(14 - Len(xAPAGARTxAgente), Z) & xAPAGARTxAgente & EspacosMeio & Chr(27) & "!" & Chr(25) & String(17, Y) & xIMPStrRetiraSIM & String(3 - Len(xIMPStrRetiraSIM), Z) & String(5, Y) & xIMPStrRetiraNAO & String(3 - Len(xIMPStrRetiraNAO), Z)
Linha(50) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(51) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxDevTransp), Z) & xPAGOTxDevTransp & String(16, Y) & String(14 - Len(xAPAGARTxDevTransp), Z) & xAPAGARTxDevTransp & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPStrLocalRetira & String(42 - Len(xIMPStrLocalRetira), Z)
Linha(52) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(53) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxOutros1), Z) & xPAGOTxOutros1 & String(2, Y) & xIMPDescrTxOutros1 & String(12 - Len(xIMPDescrTxOutros1), Z) & String(2, Y) & String(14 - Len(xAPAGARTxOutros1), Z) & xAPAGARTxOutros1 & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(54) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(55) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxOutros2), Z) & xPAGOTxOutros2 & String(2, Y) & xIMPDescrTxOutros2 & String(12 - Len(xIMPDescrTxOutros2), Z) & String(2, Y) & String(14 - Len(xAPAGARTxOutros2), Z) & xAPAGARTxOutros2 & EspacosMeio & Chr(27) & "!" & Chr(25)
Linha(56) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(42 - Len(Y), Y)
Linha(57) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteTotal), Z) & xPAGOFreteTotal & String(16, Y) & String(14 - Len(xAPAGARFreteTotal), Z) & xAPAGARFreteTotal & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Y) & String(18 - Len(xIMPStrTotalServ), Z) & xIMPStrTotalServ
Linha(58) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Y) & String(18 - Len(xIMPStrBaseCalculo), Z) & xIMPStrBaseCalculo
Linha(59) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Z), Z) & Z & String(16, Y) & String(14 - Len(Z), Z) & Z & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Z) & xIMPStrAliquota & String(5 - Len(xIMPStrAliquota), Z) & String(7, Y) & String(18 - Len(Trim(xIMPStrICMS)), Z) & xIMPStrICMS
Linha(60) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(61) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPAgenteEmissor & String(38 - Len(xIMPAgenteEmissor), Z) & String(3, Y) & xIMPCodIATA & String(18 - Len(xIMPCodIATA), Z)
Linha(62) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(63) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPDtEmissao & String(22 - Len(xIMPDtEmissao), Z) & String(3, Y) & xIMPHoraEmissao & String(18 - Len(xIMPHoraEmissao), Z)
Linha(64) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(65) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(66) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPNaturezaOp & String(19 - Len(xIMPNaturezaOp), Z) & String(1, Y) & xIMPCFOP & String(6 - Len(xIMPCFOP), Z) & String(2, Y) & xIMPEmissor & String(42 - Len(xIMPEmissor), Z) & String(1, Y) & xIMPLocalidade & String(6 - Len(xIMPLocalidade), Z) & String(1, Y) & Z & String(11 - Len(Z), Z)



DoEvents
Open SETIMPImpressoraPadrao For Output As #1
Print #1, ""
    For e = 1 To 66
    Print #1, Linha(e)
    Next
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
'Print #1, ""

Close #1
End Sub

Private Sub MascaraAWBP8()

Dim Linha(1 To 67) As String
Dim EspacosESQ As String
Dim EspacosMeio As String
Dim EspacosAUX As String

Dim Z As String

Dim xPAGOFreteNacional As String
Dim xPAGOFreteRegional As String
Dim xPAGOAdValorem As String
Dim xPAGOTxTerrOrig As String
Dim xPAGOTxTerrDest As String
Dim xPAGOTxRedesp As String
Dim xPAGOTxAgente As String
Dim xPAGOTxDevTransp As String
Dim xPAGOTxOutros1 As String
Dim xPAGOTxOutros2 As String
Dim xPAGOFreteTotal As String
Dim xAPAGARFreteNacional As String
Dim xAPAGARFreteRegional As String
Dim xAPAGARAdValorem As String
Dim xAPAGARTxTerrOrig As String
Dim xAPAGARTxTerrDest As String
Dim xAPAGARTxRedesp As String
Dim xAPAGARTxAgente As String
Dim xAPAGARTxDevTransp As String
Dim xAPAGARTxOutros1 As String
Dim xAPAGARTxOutros2 As String
Dim xAPAGARFreteTotal As String


StringNF = Trim(StringNF)
xIMPStrNF01 = Trim(xIMPStrNF01)
xIMPStrNF02 = Trim(xIMPStrNF02)
xIMPStrNF03 = Trim(xIMPStrNF03)
xIMPStrNF04 = Trim(xIMPStrNF04)
xIMPStrNF05 = Trim(xIMPStrNF05)
xIMPStrNF06 = Trim(xIMPStrNF06)
xIMPStrNF07 = Trim(xIMPStrNF07)
xIMPStrNF08 = Trim(xIMPStrNF08)
xIMPStrNF09 = Trim(xIMPStrNF09)
xIMPStrNF10 = Trim(xIMPStrNF10)
xIMPStrNF11 = Trim(xIMPStrNF11)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPNomeEXP = Trim(xIMPNomeEXP)
xIMPCGCEXP = Trim(xIMPCGCEXP)
xIMPInscEstEXP = Trim(xIMPInscEstEXP)
xIMPEndEXP = Trim(xIMPEndEXP)
xIMPBairroEXP = Trim(xIMPBairroEXP)
xIMPCidadeEXP = Trim(xIMPCidadeEXP)
xIMPCepEXP = Trim(xIMPCepEXP)
xIMPUFEXP = Trim(xIMPUFEXP)
xIMPTelEXP = Trim(xIMPTelEXP)
xIMPFAXEXP = Trim(xIMPFAXEXP)
xIMPNomeDEST = Trim(xIMPNomeDEST)
xIMPCGCDEST = Trim(xIMPCGCDEST)
xIMPInscEstDEST = Trim(xIMPInscEstDEST)
xIMPEndDEST = Trim(xIMPEndDEST)
xIMPBairroDEST = Trim(xIMPBairroDEST)
xIMPCidadeDEST = Trim(xIMPCidadeDEST)
xIMPCepDEST = Trim(xIMPCepDEST)
xIMPUFDEST = Trim(xIMPUFDEST)
xIMPTelDEST = Trim(xIMPTelDEST)
xIMPFAXDEST = Trim(xIMPFAXDEST)
xIMPOrigem = Trim(xIMPOrigem)
xIMPVia = Trim(xIMPVia)
xIMPCidadeDESTINO = Trim(xIMPCidadeDESTINO)
xIMPSIGLA = Trim(xIMPSIGLA)
xIMPReqTranspMinuta = Trim(xIMPReqTranspMinuta)
xIMPNumControle = Trim(xIMPNumControle)
xIMPInscrEstCiaAerea = Trim(xIMPInscrEstCiaAerea)
xIMPCNPJCiaAerea = Trim(xIMPCNPJCiaAerea)
xIMPStrNF12 = Trim(xIMPStrNF12)
xIMPVlDecTRANSP = Trim(xIMPVlDecTRANSP)
xIMPVlDecSUFRAMA = Trim(xIMPVlDecSUFRAMA)
xIMPDescrEmbalagem = Trim(xIMPDescrEmbalagem)
xIMPQteVol = Trim(xIMPQteVol)
xIMPPesoReal = Trim(xIMPPesoReal)
xIMPPesoTax = Trim(xIMPPesoTax)
xIMPTrecho = Trim(xIMPTrecho)
xIMPCl = Trim(xIMPCl)
xIMPCodigo = Trim(xIMPCodigo)
xIMPKilo = Trim(xIMPKilo)
xIMPFreteNacEscopo = Trim(xIMPFreteNacEscopo)
xIMPNatureza = Trim(xIMPNatureza)
xIMPTxDescrDevAg = Trim(xIMPTxDescrDevAg)
xIMPTxDescrDevTransp = Trim(xIMPTxDescrDevTransp)
xIMPFreteNacional = Trim(xIMPFreteNacional)
xIMPFreteRegional = Trim(xIMPFreteRegional)
xIMPAdValorem = Trim(xIMPAdValorem)
xIMPTipoADVAL = Trim(xIMPTipoADVAL)
xIMPTxTerrOrig = Trim(xIMPTxTerrOrig)
xIMPTxTerrDest = Trim(xIMPTxTerrDest)
xIMPTxRedesp = Trim(xIMPTxRedesp)
xIMPTxAgente = Trim(xIMPTxAgente)
xIMPTxDevTransp = Trim(xIMPTxDevTransp)
xIMPDescrTxOutros1 = Trim(xIMPDescrTxOutros1)
xIMPTxOutros1 = Trim(xIMPTxOutros1)
xIMPDescrTxOutros2 = Trim(xIMPDescrTxOutros2)
xIMPTxOutros2 = Trim(xIMPTxOutros2)
xIMPFreteTotal = Trim(xIMPFreteTotal)
xIMPStrObservacao = Trim(xIMPStrObservacao)
xIMPStrObservacao01 = Trim(xIMPStrObservacao01)
xIMPStrObservacao02 = Trim(xIMPStrObservacao02)
xIMPStrObservacao03 = Trim(xIMPStrObservacao03)
xIMPStrObservacao04 = Trim(xIMPStrObservacao04)
xIMPObsICMS = Trim(xIMPObsICMS)
xIMPObsPerecivel = Trim(xIMPObsPerecivel)
xIMPObsSeguro = Trim(xIMPObsSeguro)
xIMPStrRetiraSIM = Trim(xIMPStrRetiraSIM)
xIMPStrRetiraNAO = Trim(xIMPStrRetiraNAO)
xIMPStrLocalRetira = Trim(xIMPStrLocalRetira)
xIMPHorarioAt = Trim(xIMPHorarioAt)
xIMPStrTelefone = Trim(xIMPStrTelefone)
xIMPStrTotalServ = Trim(xIMPStrTotalServ)
xIMPStrBaseCalculo = Trim(xIMPStrBaseCalculo)
xIMPStrAliquota = Trim(xIMPStrAliquota)
xIMPStrICMS = Trim(xIMPStrICMS)
xIMPAgenteEmissor = Trim(xIMPAgenteEmissor)
xIMPCodIATA = Trim(xIMPCodIATA)
'xIMPDtEmissao = DataHora("DATA")
'xIMPHoraEmissao = DataHora("HORA")
xIMPDtEmissao = Trim(xDataIMP)
xIMPHoraEmissao = Trim(xIMPHoraEmissao)
xIMPNaturezaOp = Trim(xIMPNaturezaOp)
xIMPCFOP = Trim(xIMPCFOP)
'xIMPEmissor = xUsuario
xIMPEmissor = Trim(xIMPEmissor)
xIMPLocalidade = Trim(xIMPLocalidade)
xIMPMatricula = Trim(xIMPMatricula)


xIMPStrNF01 = Mid(xIMPStrNF01, 1, 60)
xIMPStrNF02 = Mid(xIMPStrNF02, 1, 60)
xIMPStrNF03 = Mid(xIMPStrNF03, 1, 60)
xIMPStrNF04 = Mid(xIMPStrNF04, 1, 60)
xIMPStrNF05 = Mid(xIMPStrNF05, 1, 60)
xIMPStrNF06 = Mid(xIMPStrNF06, 1, 60)
xIMPStrNF07 = Mid(xIMPStrNF07, 1, 60)
xIMPStrNF08 = Mid(xIMPStrNF08, 1, 60)
xIMPStrNF09 = Mid(xIMPStrNF09, 1, 60)
xIMPStrNF10 = Mid(xIMPStrNF10, 1, 60)
xIMPStrNF11 = Mid(xIMPStrNF11, 1, 60)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 60)
xIMPNomeEXP = Mid(xIMPNomeEXP, 1, 40)
xIMPCGCEXP = Mid(xIMPCGCEXP, 1, 40)
xIMPInscEstEXP = Mid(xIMPInscEstEXP, 1, 40)
xIMPEndEXP = Mid(xIMPEndEXP, 1, 40)
xIMPBairroEXP = Mid(xIMPBairroEXP, 1, 23)
xIMPCidadeEXP = Mid(xIMPCidadeEXP, 1, 29)
xIMPCepEXP = Mid(xIMPCepEXP, 1, 15)
xIMPUFEXP = Mid(xIMPUFEXP, 1, 8)
xIMPTelEXP = Mid(xIMPTelEXP, 1, 18)
xIMPFAXEXP = Mid(xIMPFAXEXP, 1, 20)
xIMPNomeDEST = Mid(xIMPNomeDEST, 1, 40)
xIMPCGCDEST = Mid(xIMPCGCDEST, 1, 40)
xIMPInscEstDEST = Mid(xIMPInscEstDEST, 1, 40)
xIMPEndDEST = Mid(xIMPEndDEST, 1, 40)
xIMPBairroDEST = Mid(xIMPBairroDEST, 1, 23)
xIMPCidadeDEST = Mid(xIMPCidadeDEST, 1, 29)
xIMPCepDEST = Mid(xIMPCepDEST, 1, 15)
xIMPUFDEST = Mid(xIMPUFDEST, 1, 8)
xIMPTelDEST = Mid(xIMPTelDEST, 1, 18)
xIMPFAXDEST = Mid(xIMPFAXDEST, 1, 20)
xIMPOrigem = Mid(xIMPOrigem, 1, 8)
xIMPVia = Mid(xIMPVia, 1, 8)
xIMPCidadeDESTINO = Mid(xIMPCidadeDESTINO, 1, 18)
xIMPSIGLA = Mid(xIMPSIGLA, 1, 8)
xIMPReqTranspMinuta = Mid(xIMPReqTranspMinuta, 1, 15)
xIMPNumControle = Mid(xIMPNumControle, 1, 25)
xIMPInscrEstCiaAerea = Mid(xIMPInscrEstCiaAerea, 1, 20)
xIMPCNPJCiaAerea = Mid(xIMPCNPJCiaAerea, 1, 20)
xIMPStrNF12 = Mid(xIMPStrNF12, 1, 35)
xIMPVlDecTRANSP = Mid(xIMPVlDecTRANSP, 1, 28)
xIMPVlDecSUFRAMA = Mid(xIMPVlDecSUFRAMA, 1, 28)
xIMPDescrEmbalagem = Mid(xIMPDescrEmbalagem, 1, 60)
xIMPQteVol = Mid(xIMPQteVol, 1, 5)
xIMPPesoReal = Mid(xIMPPesoReal, 1, 7)
xIMPPesoTax = Mid(xIMPPesoTax, 1, 8)
xIMPTrecho = Mid(xIMPTrecho, 1, 8)
xIMPCl = Mid(xIMPCl, 1, 2)
xIMPCodigo = Mid(xIMPCodigo, 1, 3)
xIMPKilo = Mid(xIMPKilo, 1, 9)
xIMPFreteNacEscopo = Mid(xIMPFreteNacEscopo, 1, 13)
xIMPTxDescrDevAg = Mid(xIMPTxDescrDevAg, 1, 1)
xIMPTxDescrDevTransp = Mid(xIMPTxDescrDevTransp, 1, 1)
xIMPDescrTxOutros1 = Mid(xIMPDescrTxOutros1, 1, 12)
xIMPDescrTxOutros2 = Mid(xIMPDescrTxOutros2, 1, 12)
xIMPStrObservacao01 = Mid(xIMPStrObservacao01, 1, 60)
xIMPStrObservacao02 = Mid(xIMPStrObservacao02, 1, 60)
xIMPStrObservacao03 = Mid(xIMPStrObservacao03, 1, 60)
xIMPStrObservacao04 = Mid(xIMPStrObservacao04, 1, 60)
xIMPObsICMS = Mid(xIMPObsICMS, 1, 60)
xIMPObsSeguro = Mid(xIMPObsSeguro, 1, 60)
xIMPStrLocalRetira = Mid(xIMPStrLocalRetira, 1, 42)
xIMPAgenteEmissor = Mid(xIMPAgenteEmissor, 1, 28)
xIMPCodIATA = Mid(xIMPCodIATA, 1, 28)
xIMPDtEmissao = Mid(xIMPDtEmissao, 1, 22)
xIMPHoraEmissao = Mid(xIMPHoraEmissao, 1, 18)
xIMPNaturezaOp = Mid(xIMPNaturezaOp, 1, 19)
xIMPCFOP = Mid(xIMPCFOP, 1, 6)
xIMPEmissor = Mid(xIMPEmissor, 1, 42)
xIMPLocalidade = Mid(xIMPLocalidade, 1, 6)


'65 LINHAS - 9 LINHAS EM BRANCO NO COMECO

EspacosESQ = String(3, " ")
EspacosMeio = String(2, " ")
EspacosAUX = String(5, " ")
Z = " "
Y = " "

If OptPago.Value = True Then
xPAGOFreteNacional = xIMPFreteNacional
xPAGOFreteRegional = xIMPFreteRegional
xPAGOAdValorem = xIMPAdValorem
xPAGOTxTerrOrig = xIMPTxTerrOrig
xPAGOTxTerrDest = xIMPTxTerrDest
xPAGOTxRedesp = xIMPTxRedesp
xPAGOTxAgente = xIMPTxAgente
xPAGOTxDevTransp = xIMPTxDevTransp
xPAGOTxOutros1 = xIMPTxOutros1
xPAGOTxOutros2 = xIMPTxOutros2
xPAGOFreteTotal = xIMPFreteTotal
xAPAGARFreteNacional = ""
xAPAGARFreteRegional = ""
xAPAGARAdValorem = ""
xAPAGARTxTerrOrig = ""
xAPAGARTxTerrDest = ""
xAPAGARTxRedesp = ""
xAPAGARTxAgente = ""
xAPAGARTxDevTransp = ""
xAPAGARTxOutros1 = ""
xAPAGARTxOutros2 = ""
xAPAGARFreteTotal = ""
Else
xPAGOFreteNacional = ""
xPAGOFreteRegional = ""
xPAGOAdValorem = ""
xPAGOTxTerrOrig = ""
xPAGOTxTerrDest = ""
xPAGOTxRedesp = ""
xPAGOTxAgente = ""
xPAGOTxDevTransp = ""
xPAGOTxOutros1 = ""
xPAGOTxOutros2 = ""
xPAGOFreteTotal = ""
xAPAGARFreteNacional = xIMPFreteNacional
xAPAGARFreteRegional = xIMPFreteRegional
xAPAGARAdValorem = xIMPAdValorem
xAPAGARTxTerrOrig = xIMPTxTerrOrig
xAPAGARTxTerrDest = xIMPTxTerrDest
xAPAGARTxRedesp = xIMPTxRedesp
xAPAGARTxAgente = xIMPTxAgente
xAPAGARTxDevTransp = xIMPTxDevTransp
xAPAGARTxOutros1 = xIMPTxOutros1
xAPAGARTxOutros2 = xIMPTxOutros2
xAPAGARFreteTotal = xIMPFreteTotal
End If


Linha(1) = ""
Linha(2) = ""
Linha(3) = ""
Linha(4) = ""
Linha(5) = ""
Linha(6) = ""
Linha(7) = ""
Linha(8) = ""
Linha(9) = ""
Linha(10) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF01 & String(60 - Len(xIMPStrNF01), Z)
Linha(11) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(45, Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF02 & String(60 - Len(xIMPStrNF02), Z)
Linha(12) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeEXP & String(40 - Len(xIMPNomeEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF03 & String(60 - Len(xIMPStrNF03), Z)
Linha(13) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCEXP & String(40 - Len(xIMPCGCEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF04 & String(60 - Len(xIMPStrNF04), Z)
Linha(14) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstEXP & String(40 - Len(xIMPInscEstEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF05 & String(60 - Len(xIMPStrNF05), Z)
Linha(15) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndEXP & String(40 - Len(xIMPEndEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF06 & String(60 - Len(xIMPStrNF06), Z)
Linha(16) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroEXP & String(23 - Len(xIMPBairroEXP), Z) & String(2, Y) & xIMPCepEXP & String(15 - Len(xIMPCepEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF07 & String(60 - Len(xIMPStrNF07), Z)
Linha(17) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCidadeEXP & String(29 - Len(xIMPCidadeEXP), Z) & String(3, Y) & xIMPUFEXP & String(8 - Len(xIMPUFEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF08 & String(60 - Len(xIMPStrNF08), Z)
Linha(18) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPTelEXP & String(18 - Len(xIMPTelEXP), Z) & String(2, Y) & xIMPFAXEXP & String(20 - Len(xIMPFAXEXP), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF09 & String(60 - Len(xIMPStrNF09), Z)
Linha(19) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF10 & String(60 - Len(xIMPStrNF10), Z)
Linha(20) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrNF11 & String(60 - Len(xIMPStrNF11), Z)
Linha(21) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Z & String(60 - Len(Z), Z)
Linha(22) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPNomeDEST & String(40 - Len(xIMPNomeDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(72) & xIMPStrNF12 & String(35 - Len(xIMPStrNF12), Z)
Linha(23) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCGCDEST & String(40 - Len(xIMPCGCDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(24) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPInscEstDEST & String(40 - Len(xIMPInscEstDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPReqTranspMinuta & String(20 - Len(xIMPReqTranspMinuta), Z) & String(2, Y) & xIMPNumControle & String(20 - Len(xIMPNumControle), Z)
Linha(25) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPEndDEST & String(40 - Len(xIMPEndDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(26) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPBairroDEST & String(23 - Len(xIMPBairroDEST), Z) & String(2, Y) & xIMPCepDEST & String(15 - Len(xIMPCepDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPInscrEstCiaAerea & String(20 - Len(xIMPInscrEstCiaAerea), Z) & String(2, Y) & xIMPCNPJCiaAerea & String(20 - Len(xIMPCNPJCiaAerea), Z)
Linha(27) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPCidadeDEST & String(29 - Len(xIMPCidadeDEST), Z) & String(3, Y) & xIMPUFDEST & String(8 - Len(xIMPUFDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(28) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & xIMPTelDEST & String(18 - Len(xIMPTelDEST), Z) & String(2, Y) & xIMPFAXDEST & String(20 - Len(xIMPFAXDEST), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPVlDecTRANSP & String(28 - Len(xIMPVlDecTRANSP), Z) & String(2, Y) & xIMPVlDecSUFRAMA & String(28 - Len(xIMPVlDecSUFRAMA), Z)
Linha(29) = Chr(27) & "!" & Chr(25) & EspacosESQ & EspacosAUX & Y & String(40 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(20 - Len(Y), Y) & String(2, Y) & Y & String(20 - Len(Y), Y)
Linha(30) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPOrigem & String(8 - Len(xIMPOrigem), Z) & String(1, Y) & xIMPVia & String(8 - Len(xIMPVia), Z) & String(1, Y) & xIMPCidadeDESTINO & String(18 - Len(xIMPCidadeDESTINO), Z) & String(1, Y) & xIMPSIGLA & String(8 - Len(xIMPSIGLA), Z) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPDescrEmbalagem & String(60 - Len(xIMPDescrEmbalagem), Z)
Linha(31) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(32) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(33) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(34) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(5 - Len(xIMPQteVol), Z) & xIMPQteVol & String(2, Y) & String(7 - Len(xIMPPesoReal), Z) & xIMPPesoReal & String(1, Y) & String(8 - Len(xIMPPesoTax), Z) & xIMPPesoTax & String(2, Y) & xIMPTrecho & String(8 - Len(xIMPTrecho), Z) & String(1, Y) & xIMPCl & String(2 - Len(xIMPCl), Z) & String(2, Y) & xIMPCodigo & String(3 - Len(xIMPCodigo), Z) & String(2, Y) & String(9 - Len(xIMPKilo), Z) & xIMPKilo & String(1, Y) & String(13 - Len(xIMPFreteNacEscopo), Z) & xIMPFreteNacEscopo & String(2, Y) & Chr(27) & "!" & Chr(20) & xIMPNatureza
Linha(35) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(36) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(37) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(38) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & String(1, Y) & Y & String(18 - Len(Y), Y) & String(1, Y) & Y & String(8 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & Y & String(60 - Len(Y), Y)
Linha(39) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteNacional), Z) & xPAGOFreteNacional & String(16, Y) & String(14 - Len(xAPAGARFreteNacional), Z) & xAPAGARFreteNacional & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao01 & String(60 - Len(xIMPStrObservacao01), Z)
Linha(40) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao02 & String(60 - Len(xIMPStrObservacao02), Z)
Linha(41) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteRegional), Z) & xPAGOFreteRegional & String(16, Y) & String(14 - Len(xAPAGARFreteRegional), Z) & xAPAGARFreteRegional & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao03 & String(60 - Len(xIMPStrObservacao03), Z)
Linha(42) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPStrObservacao04 & String(60 - Len(xIMPStrObservacao04), Z)
Linha(43) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOAdValorem), Z) & xPAGOAdValorem & String(10, Y) & String(4 - Len(xIMPTipoADVAL), Z) & xIMPTipoADVAL & String(2, Y) & xAPAGARAdValorem & String(14 - Len(xAPAGARAdValorem), Z) & EspacosMeio & Chr(27) & "!" & Chr(20)
Linha(44) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20)
Linha(45) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxTerrOrig), Z) & xPAGOTxTerrOrig & String(16, Y) & String(14 - Len(xAPAGARTxTerrOrig), Z) & xAPAGARTxTerrOrig & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsSeguro & String(60 - Len(xIMPObsSeguro), Z)
Linha(46) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPObsICMS & String(60 - Len(xIMPObsICMS), Z)
Linha(47) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxTerrDest), Z) & xPAGOTxTerrDest & String(16, Y) & String(14 - Len(xAPAGARTxTerrDest), Z) & xAPAGARTxTerrDest & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPObsPerecivel
Linha(48) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(42 - Len(Y), Y)
Linha(49) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxRedesp), Z) & xPAGOTxRedesp & String(16, Y) & String(14 - Len(xAPAGARTxRedesp), Z) & xAPAGARTxRedesp & EspacosMeio & Chr(27) & "!" & Chr(25) & String(17, Y) & xIMPStrRetiraSIM & String(3 - Len(xIMPStrRetiraSIM), Z) & String(5, Y) & xIMPStrRetiraNAO & String(3 - Len(xIMPStrRetiraNAO), Z)
Linha(50) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(51) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxAgente), Z) & xPAGOTxAgente & String(16, Y) & String(14 - Len(xAPAGARTxAgente), Z) & xAPAGARTxAgente & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPStrLocalRetira & String(42 - Len(xIMPStrLocalRetira), Z)
Linha(52) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(53) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxDevTransp), Z) & xPAGOTxDevTransp & String(16, Y) & String(14 - Len(xAPAGARTxDevTransp), Z) & xAPAGARTxDevTransp & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(54) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Z & String(42 - Len(Z), Z)
Linha(55) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOTxOutros1), Z) & xPAGOTxOutros1 & String(2, Y) & xIMPDescrTxOutros1 & String(12 - Len(xIMPDescrTxOutros1), Z) & String(2, Y) & String(14 - Len(xAPAGARTxOutros1), Z) & xAPAGARTxOutros1 & EspacosMeio & Chr(27) & "!" & Chr(25)
Linha(56) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & Y & String(42 - Len(Y), Y)
Linha(57) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(xPAGOFreteTotal), Z) & xPAGOFreteTotal & String(16, Y) & String(14 - Len(xAPAGARFreteTotal), Z) & xAPAGARFreteTotal & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Y) & String(18 - Len(xIMPStrTotalServ), Z) & xIMPStrTotalServ
Linha(58) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(24, Y) & String(18 - Len(xIMPStrBaseCalculo), Z) & xIMPStrBaseCalculo
Linha(59) = Chr(27) & "!" & Chr(25) & EspacosESQ & String(15 - Len(Z), Z) & Z & String(16, Y) & String(14 - Len(Z), Z) & Z & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Z) & xIMPStrAliquota & String(5 - Len(xIMPStrAliquota), Z) & String(7, Y) & String(18 - Len(Trim(xIMPStrICMS)), Z) & xIMPStrICMS
Linha(60) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(61) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(20) & xIMPAgenteEmissor & String(38 - Len(xIMPAgenteEmissor), Z) & String(3, Y) & xIMPCodIATA & String(18 - Len(xIMPCodIATA), Z)
Linha(62) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(63) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & xIMPDtEmissao & String(22 - Len(xIMPDtEmissao), Z) & String(3, Y) & xIMPHoraEmissao & String(18 - Len(xIMPHoraEmissao), Z)
Linha(64) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(65) = Chr(27) & "!" & Chr(25) & EspacosESQ & Y & String(15 - Len(Y), Y) & String(16, Y) & Y & String(14 - Len(Y), Y) & EspacosMeio & Chr(27) & "!" & Chr(25) & String(12, Y) & Y & String(5 - Len(Y), Y) & String(7, Y) & String(18 - Len(Y), Y) & Y
Linha(66) = Chr(27) & "!" & Chr(25) & EspacosESQ & xIMPNaturezaOp & String(19 - Len(xIMPNaturezaOp), Z) & String(1, Y) & xIMPCFOP & String(6 - Len(xIMPCFOP), Z) & String(2, Y) & xIMPEmissor & String(42 - Len(xIMPEmissor), Z) & String(1, Y) & xIMPLocalidade & String(6 - Len(xIMPLocalidade), Z) & String(1, Y) & Z & String(11 - Len(Z), Z)

DoEvents
Open SETIMPImpressoraPadrao For Output As #1
    For e = 1 To 66
    Print #1, Linha(e)
    Next
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""

Close #1

End Sub


