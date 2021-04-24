VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaAWB 
   Caption         =   "Consulta à AWB"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   165
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Icon            =   "frmConsultaAWB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
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
      Height          =   675
      Left            =   120
      TabIndex        =   129
      Top             =   6600
      Width           =   11715
      Begin VB.TextBox TxtHora 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6420
         Locked          =   -1  'True
         TabIndex        =   135
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtEmissor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   240
         Width           =   2715
      End
      Begin VB.TextBox TxtEmissao 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   131
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
         TabIndex        =   130
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Left            =   6000
         TabIndex        =   136
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Emissor"
         Height          =   195
         Left            =   120
         TabIndex        =   134
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   3840
         TabIndex        =   133
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
      Left            =   4680
      TabIndex        =   104
      Top             =   60
      Width           =   7155
      Begin VB.TextBox TxtNome 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   117
         Top             =   240
         Width           =   5955
      End
      Begin VB.TextBox TxtCGC 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   116
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtInscrEst 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   115
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtEnd 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   114
         Top             =   1380
         Width           =   4395
      End
      Begin VB.TextBox TxtCEP 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   113
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtTel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3060
         TabIndex        =   112
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtFAX 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   111
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtApolice 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   110
         Top             =   1980
         Width           =   1755
      End
      Begin VB.TextBox TxtSeguradora 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   109
         Top             =   1980
         Width           =   3615
      End
      Begin VB.TextBox TxtBairro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   108
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
         Left            =   6720
         TabIndex        =   107
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox TxtCidade 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   106
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
         TabIndex        =   105
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   128
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
         TabIndex        =   127
         Top             =   2325
         Width           =   405
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   126
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
         TabIndex        =   125
         Top             =   2325
         Width           =   240
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "End."
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
         Top             =   780
         Width           =   2520
      End
   End
   Begin VB.Frame FraCiaAerea 
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
      TabIndex        =   100
      Top             =   1500
      Width           =   2895
      Begin VB.TextBox TxtFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         TabIndex        =   102
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox TxtSiglaCiaAerea 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   660
         TabIndex        =   101
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
         TabIndex        =   103
         Top             =   300
         Width           =   1635
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
      Left            =   4680
      TabIndex        =   18
      Top             =   780
      Width           =   7155
      Begin VB.TextBox TxtUF 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6540
         TabIndex        =   20
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TxtCidade 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   21
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
         Left            =   6720
         TabIndex        =   69
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox TxtBairro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   63
         Top             =   1380
         Width           =   1875
      End
      Begin VB.TextBox TxtSeguradora 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1020
         TabIndex        =   60
         Top             =   1980
         Width           =   3615
      End
      Begin VB.TextBox TxtApolice 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   59
         Top             =   1980
         Width           =   1755
      End
      Begin VB.TextBox TxtFAX 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   58
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtTel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3060
         TabIndex        =   57
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtCEP 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   56
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox TxtEnd 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   54
         Top             =   1380
         Width           =   4395
      End
      Begin VB.TextBox TxtInscrEst 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   52
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtCGC 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   22
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox TxtNome 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   5955
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   55
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
         TabIndex        =   53
         Top             =   2325
         Width           =   240
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   2325
         Width           =   405
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   23
         Top             =   285
         Width           =   420
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
      Height          =   1275
      Left            =   120
      TabIndex        =   35
      Top             =   5340
      Width           =   8355
      Begin VB.TextBox TxtKiloCob 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   71
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox TxtTipoADVAL 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   62
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox TxtTipoTaxa 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   50
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox TxtOutros2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   17
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox TxtOutros1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4980
         MaxLength       =   50
         TabIndex        =   15
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox TxtICMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   45
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox TxtDescrOutros2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6120
         MaxLength       =   12
         TabIndex        =   16
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox TxtDescrOutros1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3900
         MaxLength       =   12
         TabIndex        =   14
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox TxtAliquota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   46
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox TxtTXRedesp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4980
         MaxLength       =   50
         TabIndex        =   13
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox TxtTXDestino 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4980
         MaxLength       =   50
         TabIndex        =   38
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox TxtADValorem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   37
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox TxtFreteNacional 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   36
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox TxtFreteTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   44
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Kilo"
         Height          =   195
         Left            =   2520
         TabIndex        =   72
         Top             =   585
         Width           =   255
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Taxa"
         Height          =   195
         Left            =   75
         TabIndex        =   51
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   420
         TabIndex        =   47
         Top             =   885
         Width           =   390
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tx. Redesp."
         Height          =   195
         Left            =   4020
         TabIndex        =   42
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tx. Ter. Dest."
         Height          =   195
         Left            =   3960
         TabIndex        =   41
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ad. Val."
         Height          =   195
         Left            =   2220
         TabIndex        =   40
         Top             =   885
         Width           =   555
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Frete Nac."
         Height          =   195
         Left            =   2025
         TabIndex        =   39
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Alíquota"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Top             =   585
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Frete Total"
         Height          =   195
         Left            =   6360
         TabIndex        =   43
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Defina o Tipo de Procura"
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
      TabIndex        =   79
      Top             =   60
      Width           =   4515
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
         Left            =   1500
         TabIndex        =   80
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox TxtBuscaDig 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2340
            MaxLength       =   1
            TabIndex        =   4
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox TxtBuscaAWB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   960
            MaxLength       =   10
            TabIndex        =   3
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox TxtBuscaSiglaAWB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   540
            MaxLength       =   2
            TabIndex        =   2
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox TxtBuscaAWBFilial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   1
            Top             =   240
            Width           =   435
         End
      End
      Begin VB.Frame FraCTC 
         Caption         =   "Filial e CTC/CTR"
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
         Left            =   1500
         TabIndex        =   82
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox TxtBuscaFilial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtBuscaCTC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   660
            MaxLength       =   8
            TabIndex        =   9
            Top             =   240
            Width           =   2115
         End
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "Sair"
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         Top             =   900
         Width           =   1035
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   1500
         TabIndex        =   10
         Top             =   900
         Width           =   1815
      End
      Begin VB.OptionButton OptCTC 
         Caption         =   "Por CTC/CTR"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.OptionButton OptNF 
         Caption         =   "Por NF"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptAWB 
         Caption         =   "Por AWB"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1335
      End
      Begin VB.Frame FraNF 
         Caption         =   "Número de Nota Fiscal"
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
         Left            =   1500
         TabIndex        =   81
         Top             =   240
         Width           =   2895
         Begin VB.TextBox TxtBuscaNF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   6
            Top             =   240
            Width           =   2655
         End
      End
   End
   Begin VB.Frame FraSpot 
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
      Left            =   8520
      TabIndex        =   61
      Top             =   5340
      Width           =   3315
      Begin VB.TextBox TxtPesoCubado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   88
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtPesoReal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   87
         Top             =   780
         Width           =   735
      End
      Begin VB.TextBox TxtVolumes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   85
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox TxtTotalVM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   83
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Peso Cub."
         Height          =   195
         Left            =   1680
         TabIndex        =   90
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Peso Real"
         Height          =   195
         Left            =   1680
         TabIndex        =   89
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Volumes"
         Height          =   195
         Left            =   180
         TabIndex        =   86
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total VM."
         Height          =   195
         Left            =   90
         TabIndex        =   84
         Top             =   825
         Width           =   690
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
      Height          =   1035
      Left            =   120
      TabIndex        =   32
      Top             =   4260
      Width           =   11715
      Begin MSFlexGridLib.MSFlexGrid FlexGridVolumes 
         Height          =   675
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   1191
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         SelectionMode   =   1
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
      TabIndex        =   26
      Top             =   1500
      Width           =   8775
      Begin VB.TextBox TxtAeroportoVIA 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   27
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox TxtSiglaVIA 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   28
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox TxtAeroportoDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4140
         TabIndex        =   68
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox TxtSiglaDestinatario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3660
         TabIndex        =   67
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox TxtAeroportoExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   64
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox TxtSiglaExpedidor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   660
         TabIndex        =   65
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Left            =   3060
         TabIndex        =   70
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Origem"
         Height          =   195
         Left            =   105
         TabIndex        =   66
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "VIA"
         Height          =   195
         Left            =   6045
         TabIndex        =   29
         Top             =   345
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
      Height          =   675
      Left            =   120
      TabIndex        =   34
      Top             =   3540
      Width           =   11715
      Begin VB.TextBox TxtModal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   10260
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtClienteRetira 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8820
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtDescrIATA 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   240
         Width           =   3795
      End
      Begin VB.TextBox TxtPerecivel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtEspecie 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Modal Frete"
         Height          =   195
         Left            =   9360
         TabIndex        =   99
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Retira?"
         Height          =   195
         Left            =   7740
         TabIndex        =   97
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Perecível?"
         Height          =   195
         Left            =   6180
         TabIndex        =   94
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Espécie"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   300
         Width           =   570
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
      Height          =   915
      Left            =   120
      TabIndex        =   49
      Top             =   7260
      Width           =   11715
      Begin VB.TextBox TxtOBSEmissao 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   555
         Left            =   120
         MaxLength       =   239
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   11475
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
      Height          =   1275
      Left            =   120
      TabIndex        =   30
      Top             =   2220
      Width           =   11715
      Begin MSFlexGridLib.MSFlexGrid FlexGridNFs 
         Height          =   975
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         SelectionMode   =   1
      End
   End
End
Attribute VB_Name = "frmConsultaAWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public COMBODIGESP As Integer

Private Sub CmdBuscar_Click()
If de_informa.rsSelAWB_NF.State = 1 Then de_informa.rsSelAWB_NF.Close
If de_informa.rsSelAWB.State = 1 Then de_informa.rsSelAWB.Close
If de_informa.rsSelAWB_CTC.State = 1 Then de_informa.rsSelAWB_CTC.Close

    If OptNF.Value = True Then
        If Len(Trim(TxtBuscaNF.Text)) = 0 Then
        Exit Sub
        End If
        
        If de_informa.rsSelAWB_NF.State = 1 Then de_informa.rsSelAWB_NF.Close
        de_informa.SelAWB_NF TxtBuscaNF.Text
        
        If de_informa.rsSelAWB_NF.RecordCount = 1 Then
        
        FlexGridNFs.Rows = 0
        FlexGridVolumes.Rows = 0
        Call LimpaTela(Me)
        
        TxtFilial.Text = de_informa.rsSelAWB_NF.Fields("filial")
        TxtSiglaCiaAerea.Text = de_informa.rsSelAWB_NF.Fields("cia")
        TxtAWB.Caption = de_informa.rsSelAWB_NF.Fields("awb") & "-" & de_informa.rsSelAWB_NF.Fields("dig")
        TxtSiglaExpedidor.Text = de_informa.rsSelAWB_NF.Fields("siglaorigem")
        TxtSiglaVIA.Text = de_informa.rsSelAWB_NF.Fields("siglavia")
        TxtSiglaDestinatario.Text = de_informa.rsSelAWB_NF.Fields("siglades")
        
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cidadeorigem")) Then TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("cidadeorigem")) & " - " & de_informa.rsSelAWB_NF.Fields("uforigem") & " (" & PriMaiuscula(de_informa.rsSelAWB_NF.Fields("aeroportoorigem")) & ")"
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cidadevia")) Then TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("cidadevia")) & " - " & de_informa.rsSelAWB_NF.Fields("ufvia") & " (" & PriMaiuscula(de_informa.rsSelAWB_NF.Fields("aeroportovia")) & ")"
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cidadedestino")) Then TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("cidadedestino")) & " - " & de_informa.rsSelAWB_NF.Fields("ufdestino") & " (" & PriMaiuscula(de_informa.rsSelAWB_NF.Fields("aeroportodestino")) & ")"
        
        TxtEspecie.Text = de_informa.rsSelAWB_NF.Fields("especie")
        TxtDescrIATA.Text = de_informa.rsSelAWB_NF.Fields("descrprodsis")
        
            If de_informa.rsSelAWB_NF.Fields("perecivel") = "S" Then
            TxtPerecivel.Text = "S"
            Else
            TxtPerecivel.Text = "N"
            End If
            
            If de_informa.rsSelAWB_NF.Fields("retira") = "S" Then
            TxtClienteRetira.Text = "S"
            Else
            TxtClienteRetira.Text = "N"
            End If
            
        TxtModal.Text = de_informa.rsSelAWB_NF.Fields("modal")
        TxtTipoTaxa.Text = de_informa.rsSelAWB_NF.Fields("tipotaxa")
        TxtAliquota.Text = de_informa.rsSelAWB_NF.Fields("aliquota")
        TxtICMS.Text = de_informa.rsSelAWB_NF.Fields("icms")
        TxtFreteNacional.Text = de_informa.rsSelAWB_NF.Fields("fretenacional")
        TxtKiloCob.Text = de_informa.rsSelAWB_NF.Fields("kilo")
        TxtADValorem.Text = de_informa.rsSelAWB_NF.Fields("advalorem")
        TxtTipoADVAL.Text = de_informa.rsSelAWB_NF.Fields("tipoadval")
        TxtTXDestino.Text = de_informa.rsSelAWB_NF.Fields("txdestino")
        TxtTXRedesp.Text = de_informa.rsSelAWB_NF.Fields("txredesp")
        TxtDescrOutros1.Text = de_informa.rsSelAWB_NF.Fields("descrtxoutros1")
        TxtOutros1.Text = de_informa.rsSelAWB_NF.Fields("txoutros1")
        TxtDescrOutros2.Text = de_informa.rsSelAWB_NF.Fields("descrtxoutros2")
        TxtOutros2.Text = de_informa.rsSelAWB_NF.Fields("txoutros2")
        TxtFreteTotal.Text = de_informa.rsSelAWB_NF.Fields("fretetotal")
        TxtVolumes.Text = de_informa.rsSelAWB_NF.Fields("volumes")
        TxtPesoCubado.Text = de_informa.rsSelAWB_NF.Fields("pesocubado")
        TxtPesoReal.Text = de_informa.rsSelAWB_NF.Fields("pesoreal")
        TxtEmissor.Text = de_informa.rsSelAWB_NF.Fields("emissor")
        TxtEmissao.Text = de_informa.rsSelAWB_NF.Fields("data")
        TxtHora.Text = de_informa.rsSelAWB_NF.Fields("hora")
        
            If de_informa.rsSelAWB_NF.Fields("cancelado") = "X" Then
            TxtStatus.Text = "AWB Cancelado"
            Else
            TxtStatus.Text = ""
            End If
        
        TxtOBSEmissao.Text = de_informa.rsSelAWB_NF.Fields("obsemissor")
        
        If de_informa.rsSelAWB.State = 1 Then de_informa.rsSelAWB.Close
        de_informa.SelAWB de_informa.rsSelAWB_NF.Fields("codawb")
        
        FlexGridNFs.Clear
        FlexGridNFs.Rows = de_informa.rsSelAWB.RecordCount + 1
        FlexGridNFs.Cols = 6
        FlexGridNFs.FixedCols = 0
        FlexGridNFs.FixedRows = 1
        
        FlexGridNFs.TextMatrix(0, 0) = "NF"
        FlexGridNFs.TextMatrix(0, 1) = "Série"
        FlexGridNFs.TextMatrix(0, 2) = "Valor"
        FlexGridNFs.TextMatrix(0, 3) = "FilialCTC"
        FlexGridNFs.TextMatrix(0, 4) = "Remetente"
        FlexGridNFs.TextMatrix(0, 5) = "Destinatário"
        
        FlexGridNFs.ColWidth(0) = 700
        FlexGridNFs.ColWidth(1) = 500
        FlexGridNFs.ColWidth(2) = 1300
        FlexGridNFs.ColWidth(3) = 1200
        FlexGridNFs.ColWidth(4) = 3500
        FlexGridNFs.ColWidth(5) = 3500
        
        xCodAwb = de_informa.rsSelAWB.Fields("codawb")
        
        
        X = 0
        
            Do Until de_informa.rsSelAWB.EOF
            X = X + 1
            
            If Not IsNull(de_informa.rsSelAWB.Fields("nota")) Then FlexGridNFs.TextMatrix(X, 0) = de_informa.rsSelAWB.Fields("nota")
            If Not IsNull(de_informa.rsSelAWB.Fields("SERIE")) Then FlexGridNFs.TextMatrix(X, 1) = de_informa.rsSelAWB.Fields("serie")
            If Not IsNull(de_informa.rsSelAWB.Fields("VALOR")) Then FlexGridNFs.TextMatrix(X, 2) = Format(de_informa.rsSelAWB.Fields("VALOR"), "##,##0.00")
            If Not IsNull(de_informa.rsSelAWB.Fields("FILIALCTC")) Then FlexGridNFs.TextMatrix(X, 3) = de_informa.rsSelAWB.Fields("FILIALCTC")
            If Not IsNull(de_informa.rsSelAWB.Fields("REMET_NOME")) Then FlexGridNFs.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelAWB.Fields("REMET_NOME"))
            If Not IsNull(de_informa.rsSelAWB.Fields("DEST_NOME")) Then FlexGridNFs.TextMatrix(X, 5) = PriMaiuscula(de_informa.rsSelAWB.Fields("DEST_NOME"))
            
            de_informa.rsSelAWB.MoveNext
            Loop
            
        If de_informa.rsConsultaAWBVolume.State = 1 Then de_informa.rsConsultaAWBVolume.Close
        de_informa.ConsultaAWBVolume xCodAwb
        FlexGridVolumes.Clear
        FlexGridVolumes.Rows = de_informa.rsConsultaAWBVolume.RecordCount + 1
        FlexGridVolumes.Cols = 4
        FlexGridVolumes.FixedCols = 0
        FlexGridVolumes.FixedRows = 1
        
        FlexGridVolumes.TextMatrix(0, 0) = "Qtde."
        FlexGridVolumes.TextMatrix(0, 1) = "Comprimento"
        FlexGridVolumes.TextMatrix(0, 2) = "Largura"
        FlexGridVolumes.TextMatrix(0, 3) = "Altura"
        
        FlexGridVolumes.ColWidth(0) = 500
        FlexGridVolumes.ColWidth(1) = 1500
        FlexGridVolumes.ColWidth(2) = 1500
        FlexGridVolumes.ColWidth(3) = 1500
        
        X = 0
        
            Do Until de_informa.rsConsultaAWBVolume.EOF
            X = X + 1
            FlexGridVolumes.TextMatrix(X, 0) = de_informa.rsConsultaAWBVolume.Fields("volumes")
            FlexGridVolumes.TextMatrix(X, 1) = de_informa.rsConsultaAWBVolume.Fields("comprimento")
            FlexGridVolumes.TextMatrix(X, 2) = de_informa.rsConsultaAWBVolume.Fields("largura")
            FlexGridVolumes.TextMatrix(X, 3) = de_informa.rsConsultaAWBVolume.Fields("altura")
            
            de_informa.rsConsultaAWBVolume.MoveNext
            Loop
            
        TxtTotalVM.Text = de_informa.rsSelAWB_NF.Fields("valmerc")
        
        X = 1
        
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("nomeexp")) Then TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("nomeexp"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("endexp")) Then TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("endexp"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("bairroexp")) Then TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("bairroexp"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cidadexp")) Then TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("cidadexp"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("ufexp")) Then TxtUF(X).Text = de_informa.rsSelAWB_NF.Fields("ufexp")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("telexp")) Then TxtTel(X).Text = de_informa.rsSelAWB_NF.Fields("telexp")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("faxexp")) Then TxtFAX(X).Text = de_informa.rsSelAWB_NF.Fields("faxexp")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cnpjexp")) Then TxtCGC(X).Text = de_informa.rsSelAWB_NF.Fields("cnpjexp")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("inscrestexp")) Then TxtInscrEst(X).Text = de_informa.rsSelAWB_NF.Fields("inscrestexp")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cepexp")) Then TxtCEP(X).Text = de_informa.rsSelAWB_NF.Fields("cepexp")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("segexp")) Then TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("segexp"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("apoliceexp")) Then TxtApolice(X).Text = de_informa.rsSelAWB_NF.Fields("apoliceexp")
        
        
        X = 0
        
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("nomedes")) Then TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("nomedes"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("enddes")) Then TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("enddes"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("bairrodes")) Then TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("bairrodes"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cidadedes")) Then TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("cidadedes"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("ufdes")) Then TxtUF(X).Text = de_informa.rsSelAWB_NF.Fields("ufdes")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("teldes")) Then TxtTel(X).Text = de_informa.rsSelAWB_NF.Fields("teldes")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("faxdes")) Then TxtFAX(X).Text = de_informa.rsSelAWB_NF.Fields("faxdes")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cnpjdes")) Then TxtCGC(X).Text = de_informa.rsSelAWB_NF.Fields("cnpjdes")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("inscrestdes")) Then TxtInscrEst(X).Text = de_informa.rsSelAWB_NF.Fields("inscrestdes")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("cepdes")) Then TxtCEP(X).Text = de_informa.rsSelAWB_NF.Fields("cepdes")
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("segdes")) Then TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB_NF.Fields("segdes"))
        If Not IsNull(de_informa.rsSelAWB_NF.Fields("apolicedes")) Then TxtApolice(X).Text = de_informa.rsSelAWB_NF.Fields("apolicedes")
        ElseIf de_informa.rsSelAWB_NF.RecordCount > 1 Then
        FlexGridNFs.Rows = 0
        FlexGridVolumes.Rows = 0
        Call LimpaTela(Me)
        frmConsultaAWBFiltra.Show 1
        ElseIf de_informa.rsSelAWB_NF.RecordCount = 0 Then
        MsgBox "Nota Fiscal não encontrada!", vbCritical, ""
        Exit Sub
        End If
    ElseIf OptAWB.Value = True Then
        If Len(Trim(TxtBuscaAWBFilial.Text)) = 0 Or Len(Trim(TxtBuscaSiglaAWB.Text)) = 0 Or Len(Trim(TxtBuscaAWB.Text)) = 0 Or Len(Trim(TxtBuscaDig.Text)) = 0 Then
        Exit Sub
        End If
        
    xAWB = Trim(TxtBuscaAWB.Text)
    xDig = Trim(TxtBuscaDig.Text)
    xCodAwb = String(2 - Len(Trim(Str(Val(TxtBuscaAWBFilial.Text)))), "0") & Trim(Str(Val(TxtBuscaAWBFilial.Text))) & UCase(Trim(TxtBuscaSiglaAWB.Text)) & String(10 - Len(Trim(Str(Val(xAWB)))), "0") & Trim(Str(Val(xAWB))) & Trim(Str(Val(xDig)))
    
        If de_informa.rsSelAWB.State = 1 Then de_informa.rsSelAWB.Close
        FlexGridNFs.Rows = 0
        FlexGridVolumes.Rows = 0
        Call LimpaTela(Me)
        de_informa.SelAWB xCodAwb
        
        If de_informa.rsSelAWB.RecordCount > 0 Then
        
        TxtFilial.Text = de_informa.rsSelAWB.Fields("filial")
        TxtSiglaCiaAerea.Text = de_informa.rsSelAWB.Fields("cia")
        TxtAWB.Caption = de_informa.rsSelAWB.Fields("awb") & "-" & de_informa.rsSelAWB.Fields("dig")
        TxtSiglaExpedidor.Text = de_informa.rsSelAWB.Fields("siglaorigem")
        TxtSiglaVIA.Text = de_informa.rsSelAWB.Fields("siglavia")
        TxtSiglaDestinatario.Text = de_informa.rsSelAWB.Fields("siglades")
        
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadeorigem")) Then TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadeorigem")) & " - " & de_informa.rsSelAWB.Fields("uforigem") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportoorigem")) & ")"
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadevia")) Then TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadevia")) & " - " & de_informa.rsSelAWB.Fields("ufvia") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportovia")) & ")"
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadedestino")) Then TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadedestino")) & " - " & de_informa.rsSelAWB.Fields("ufdestino") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportodestino")) & ")"
        
        TxtEspecie.Text = de_informa.rsSelAWB.Fields("especie")
        TxtDescrIATA.Text = de_informa.rsSelAWB.Fields("descrprodsis")
        
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
        TxtTipoTaxa.Text = de_informa.rsSelAWB.Fields("tipotaxa")
        TxtAliquota.Text = de_informa.rsSelAWB.Fields("aliquota")
        TxtICMS.Text = de_informa.rsSelAWB.Fields("icms")
        TxtFreteNacional.Text = de_informa.rsSelAWB.Fields("fretenacional")
        TxtKiloCob.Text = de_informa.rsSelAWB.Fields("kilo")
        TxtADValorem.Text = de_informa.rsSelAWB.Fields("advalorem")
        TxtTipoADVAL.Text = de_informa.rsSelAWB.Fields("tipoadval")
        TxtTXDestino.Text = de_informa.rsSelAWB.Fields("txdestino")
        TxtTXRedesp.Text = de_informa.rsSelAWB.Fields("txredesp")
        TxtDescrOutros1.Text = de_informa.rsSelAWB.Fields("descrtxoutros1")
        TxtOutros1.Text = de_informa.rsSelAWB.Fields("txoutros1")
        TxtDescrOutros2.Text = de_informa.rsSelAWB.Fields("descrtxoutros2")
        TxtOutros2.Text = de_informa.rsSelAWB.Fields("txoutros2")
        TxtFreteTotal.Text = de_informa.rsSelAWB.Fields("fretetotal")
        TxtVolumes.Text = de_informa.rsSelAWB.Fields("volumes")
        TxtPesoCubado.Text = de_informa.rsSelAWB.Fields("pesocubado")
        TxtPesoReal.Text = de_informa.rsSelAWB.Fields("pesoreal")
        TxtEmissor.Text = de_informa.rsSelAWB.Fields("emissor")
        TxtEmissao.Text = de_informa.rsSelAWB.Fields("data")
        TxtHora.Text = de_informa.rsSelAWB.Fields("hora")
        
            If de_informa.rsSelAWB.Fields("cancelado") = "X" Then
            TxtStatus.Text = "AWB Cancelado"
            Else
            TxtStatus.Text = ""
            End If
        
        TxtOBSEmissao.Text = de_informa.rsSelAWB.Fields("obsemissor")
        
        
        FlexGridNFs.Clear
        FlexGridNFs.Rows = de_informa.rsSelAWB.RecordCount + 1
        FlexGridNFs.Cols = 6
        FlexGridNFs.FixedCols = 0
        FlexGridNFs.FixedRows = 1
        
        FlexGridNFs.TextMatrix(0, 0) = "NF"
        FlexGridNFs.TextMatrix(0, 1) = "Série"
        FlexGridNFs.TextMatrix(0, 2) = "Valor"
        FlexGridNFs.TextMatrix(0, 3) = "FilialCTC"
        FlexGridNFs.TextMatrix(0, 4) = "Remetente"
        FlexGridNFs.TextMatrix(0, 5) = "Destinatário"
        
        FlexGridNFs.ColWidth(0) = 700
        FlexGridNFs.ColWidth(1) = 500
        FlexGridNFs.ColWidth(2) = 1300
        FlexGridNFs.ColWidth(3) = 1200
        FlexGridNFs.ColWidth(4) = 3500
        FlexGridNFs.ColWidth(5) = 3500
        
        xCodAwb = de_informa.rsSelAWB.Fields("codawb")
        
        
        X = 0
        
            Do Until de_informa.rsSelAWB.EOF
            X = X + 1
            
            If Not IsNull(de_informa.rsSelAWB.Fields("nota")) Then FlexGridNFs.TextMatrix(X, 0) = de_informa.rsSelAWB.Fields("nota")
            If Not IsNull(de_informa.rsSelAWB.Fields("SERIE")) Then FlexGridNFs.TextMatrix(X, 1) = de_informa.rsSelAWB.Fields("serie")
            If Not IsNull(de_informa.rsSelAWB.Fields("VALOR")) Then FlexGridNFs.TextMatrix(X, 2) = Format(de_informa.rsSelAWB.Fields("VALOR"), "##,##0.00")
            If Not IsNull(de_informa.rsSelAWB.Fields("FILIALCTC")) Then FlexGridNFs.TextMatrix(X, 3) = de_informa.rsSelAWB.Fields("FILIALCTC")
            If Not IsNull(de_informa.rsSelAWB.Fields("REMET_NOME")) Then FlexGridNFs.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelAWB.Fields("REMET_NOME"))
            If Not IsNull(de_informa.rsSelAWB.Fields("DEST_NOME")) Then FlexGridNFs.TextMatrix(X, 5) = PriMaiuscula(de_informa.rsSelAWB.Fields("DEST_NOME"))
            
            de_informa.rsSelAWB.MoveNext
            Loop
            
        If de_informa.rsConsultaAWBVolume.State = 1 Then de_informa.rsConsultaAWBVolume.Close
        de_informa.ConsultaAWBVolume xCodAwb
        FlexGridVolumes.Clear
        FlexGridVolumes.Rows = de_informa.rsConsultaAWBVolume.RecordCount + 2
        FlexGridVolumes.Cols = 4
        FlexGridVolumes.FixedCols = 0
        FlexGridVolumes.FixedRows = 1
        
        FlexGridVolumes.TextMatrix(0, 0) = "Qtde."
        FlexGridVolumes.TextMatrix(0, 1) = "Comprimento"
        FlexGridVolumes.TextMatrix(0, 2) = "Largura"
        FlexGridVolumes.TextMatrix(0, 3) = "Altura"
        
        FlexGridVolumes.ColWidth(0) = 500
        FlexGridVolumes.ColWidth(1) = 1500
        FlexGridVolumes.ColWidth(2) = 1500
        FlexGridVolumes.ColWidth(3) = 1500
        
        X = 0
        
            Do Until de_informa.rsConsultaAWBVolume.EOF
            X = X + 1
            FlexGridVolumes.TextMatrix(X, 0) = de_informa.rsConsultaAWBVolume.Fields("volumes")
            FlexGridVolumes.TextMatrix(X, 1) = de_informa.rsConsultaAWBVolume.Fields("comprimento")
            FlexGridVolumes.TextMatrix(X, 2) = de_informa.rsConsultaAWBVolume.Fields("largura")
            FlexGridVolumes.TextMatrix(X, 3) = de_informa.rsConsultaAWBVolume.Fields("altura")
            
            de_informa.rsConsultaAWBVolume.MoveNext
            Loop
            
        de_informa.rsSelAWB.MoveFirst
            
        TxtTotalVM.Text = de_informa.rsSelAWB.Fields("valmerc")
        
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

        End If
    ElseIf OptCTC.Value = True Then
        
        If Len(Trim(TxtBuscaFilial.Text)) = 0 Or Len(Trim(TxtBuscaCTC.Text)) = 0 Then
        Exit Sub
        End If
        
        xFilialCTC = String(2 - Len(Trim(Str(Val(TxtBuscaFilial.Text)))), "0") & Trim(Str(Val(TxtBuscaFilial.Text))) & String(8 - Len(Trim(Str(Val(TxtBuscaCTC.Text)))), "0") & Trim(Str(Val(TxtBuscaCTC.Text)))
        
        If de_informa.rsSelAWB_CTC.State = 1 Then de_informa.rsSelAWB_CTC.Close
        de_informa.SelAWB_CTC xFilialCTC
        
        If de_informa.rsSelAWB_CTC.RecordCount = 1 Then
        FlexGridNFs.Rows = 0
        FlexGridVolumes.Rows = 0
        Call LimpaTela(Me)
        TxtFilial.Text = de_informa.rsSelAWB_CTC.Fields("filial")
        TxtSiglaCiaAerea.Text = de_informa.rsSelAWB_CTC.Fields("cia")
        TxtAWB.Caption = de_informa.rsSelAWB_CTC.Fields("awb") & "-" & de_informa.rsSelAWB_CTC.Fields("dig")
        TxtSiglaExpedidor.Text = de_informa.rsSelAWB_CTC.Fields("siglaorigem")
        TxtSiglaVIA.Text = de_informa.rsSelAWB_CTC.Fields("siglavia")
        TxtSiglaDestinatario.Text = de_informa.rsSelAWB_CTC.Fields("siglades")
        
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cidadeorigem")) Then TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("cidadeorigem")) & " - " & de_informa.rsSelAWB_CTC.Fields("uforigem") & " (" & PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("aeroportoorigem")) & ")"
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cidadevia")) Then TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("cidadevia")) & " - " & de_informa.rsSelAWB_CTC.Fields("ufvia") & " (" & PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("aeroportovia")) & ")"
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cidadedestino")) Then TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("cidadedestino")) & " - " & de_informa.rsSelAWB_CTC.Fields("ufdestino") & " (" & PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("aeroportodestino")) & ")"
        
        TxtEspecie.Text = de_informa.rsSelAWB_CTC.Fields("especie")
        TxtDescrIATA.Text = de_informa.rsSelAWB_CTC.Fields("descrprodsis")
        
            If de_informa.rsSelAWB_CTC.Fields("perecivel") = "S" Then
            TxtPerecivel.Text = "S"
            Else
            TxtPerecivel.Text = "N"
            End If
            
            If de_informa.rsSelAWB_CTC.Fields("retira") = "S" Then
            TxtClienteRetira.Text = "S"
            Else
            TxtClienteRetira.Text = "N"
            End If
            
        TxtModal.Text = de_informa.rsSelAWB_CTC.Fields("modal")
        TxtTipoTaxa.Text = de_informa.rsSelAWB_CTC.Fields("tipotaxa")
        TxtAliquota.Text = de_informa.rsSelAWB_CTC.Fields("aliquota")
        TxtICMS.Text = de_informa.rsSelAWB_CTC.Fields("icms")
        TxtFreteNacional.Text = de_informa.rsSelAWB_CTC.Fields("fretenacional")
        TxtKiloCob.Text = de_informa.rsSelAWB_CTC.Fields("kilo")
        TxtADValorem.Text = de_informa.rsSelAWB_CTC.Fields("advalorem")
        TxtTipoADVAL.Text = de_informa.rsSelAWB_CTC.Fields("tipoadval")
        TxtTXDestino.Text = de_informa.rsSelAWB_CTC.Fields("txdestino")
        TxtTXRedesp.Text = de_informa.rsSelAWB_CTC.Fields("txredesp")
        TxtDescrOutros1.Text = de_informa.rsSelAWB_CTC.Fields("descrtxoutros1")
        TxtOutros1.Text = de_informa.rsSelAWB_CTC.Fields("txoutros1")
        TxtDescrOutros2.Text = de_informa.rsSelAWB_CTC.Fields("descrtxoutros2")
        TxtOutros2.Text = de_informa.rsSelAWB_CTC.Fields("txoutros2")
        TxtFreteTotal.Text = de_informa.rsSelAWB_CTC.Fields("fretetotal")
        TxtVolumes.Text = de_informa.rsSelAWB_CTC.Fields("volumes")
        TxtPesoCubado.Text = de_informa.rsSelAWB_CTC.Fields("pesocubado")
        TxtPesoReal.Text = de_informa.rsSelAWB_CTC.Fields("pesoreal")
        TxtEmissor.Text = de_informa.rsSelAWB_CTC.Fields("emissor")
        TxtEmissao.Text = de_informa.rsSelAWB_CTC.Fields("data")
        TxtHora.Text = de_informa.rsSelAWB_CTC.Fields("hora")
        
            If de_informa.rsSelAWB_CTC.Fields("cancelado") = "X" Then
            TxtStatus.Text = "AWB Cancelado"
            Else
            TxtStatus.Text = ""
            End If
        
        TxtOBSEmissao.Text = de_informa.rsSelAWB_CTC.Fields("obsemissor")
        
        If de_informa.rsSelAWB.State = 1 Then de_informa.rsSelAWB.Close
        de_informa.SelAWB de_informa.rsSelAWB_CTC.Fields("codawb")
        
        FlexGridNFs.Clear
        FlexGridNFs.Rows = de_informa.rsSelAWB.RecordCount + 1
        FlexGridNFs.Cols = 6
        FlexGridNFs.FixedCols = 0
        FlexGridNFs.FixedRows = 1
        
        FlexGridNFs.TextMatrix(0, 0) = "NF"
        FlexGridNFs.TextMatrix(0, 1) = "Série"
        FlexGridNFs.TextMatrix(0, 2) = "Valor"
        FlexGridNFs.TextMatrix(0, 3) = "FilialCTC"
        FlexGridNFs.TextMatrix(0, 4) = "Remetente"
        FlexGridNFs.TextMatrix(0, 5) = "Destinatário"
        
        FlexGridNFs.ColWidth(0) = 700
        FlexGridNFs.ColWidth(1) = 500
        FlexGridNFs.ColWidth(2) = 1300
        FlexGridNFs.ColWidth(3) = 1200
        FlexGridNFs.ColWidth(4) = 3500
        FlexGridNFs.ColWidth(5) = 3500
        
        xCodAwb = de_informa.rsSelAWB.Fields("codawb")
        
        
        X = 0
        
            Do Until de_informa.rsSelAWB.EOF
            X = X + 1
            
            If Not IsNull(de_informa.rsSelAWB.Fields("nota")) Then FlexGridNFs.TextMatrix(X, 0) = de_informa.rsSelAWB.Fields("nota")
            If Not IsNull(de_informa.rsSelAWB.Fields("SERIE")) Then FlexGridNFs.TextMatrix(X, 1) = de_informa.rsSelAWB.Fields("serie")
            If Not IsNull(de_informa.rsSelAWB.Fields("VALOR")) Then FlexGridNFs.TextMatrix(X, 2) = Format(de_informa.rsSelAWB.Fields("VALOR"), "##,##0.00")
            If Not IsNull(de_informa.rsSelAWB.Fields("FILIALCTC")) Then FlexGridNFs.TextMatrix(X, 3) = de_informa.rsSelAWB.Fields("FILIALCTC")
            If Not IsNull(de_informa.rsSelAWB.Fields("REMET_NOME")) Then FlexGridNFs.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelAWB.Fields("REMET_NOME"))
            If Not IsNull(de_informa.rsSelAWB.Fields("DEST_NOME")) Then FlexGridNFs.TextMatrix(X, 5) = PriMaiuscula(de_informa.rsSelAWB.Fields("DEST_NOME"))
            
            de_informa.rsSelAWB.MoveNext
            Loop
            
        If de_informa.rsConsultaAWBVolume.State = 1 Then de_informa.rsConsultaAWBVolume.Close
        de_informa.ConsultaAWBVolume xCodAwb
        FlexGridVolumes.Clear
        FlexGridVolumes.Rows = de_informa.rsConsultaAWBVolume.RecordCount + 1
        FlexGridVolumes.Cols = 4
        FlexGridVolumes.FixedCols = 0
        FlexGridVolumes.FixedRows = 1
        
        FlexGridVolumes.TextMatrix(0, 0) = "Qtde."
        FlexGridVolumes.TextMatrix(0, 1) = "Comprimento"
        FlexGridVolumes.TextMatrix(0, 2) = "Largura"
        FlexGridVolumes.TextMatrix(0, 3) = "Altura"
        
        FlexGridVolumes.ColWidth(0) = 500
        FlexGridVolumes.ColWidth(1) = 1500
        FlexGridVolumes.ColWidth(2) = 1500
        FlexGridVolumes.ColWidth(3) = 1500
        
        X = 0
        
            Do Until de_informa.rsConsultaAWBVolume.EOF
            X = X + 1
            FlexGridVolumes.TextMatrix(X, 0) = de_informa.rsConsultaAWBVolume.Fields("volumes")
            FlexGridVolumes.TextMatrix(X, 1) = de_informa.rsConsultaAWBVolume.Fields("comprimento")
            FlexGridVolumes.TextMatrix(X, 2) = de_informa.rsConsultaAWBVolume.Fields("largura")
            FlexGridVolumes.TextMatrix(X, 3) = de_informa.rsConsultaAWBVolume.Fields("altura")
            
            de_informa.rsConsultaAWBVolume.MoveNext
            Loop
            
        TxtTotalVM.Text = de_informa.rsSelAWB_CTC.Fields("valmerc")
        
        X = 1
        
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("nomeexp")) Then TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("nomeexp"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("endexp")) Then TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("endexp"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("bairroexp")) Then TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("bairroexp"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cidadexp")) Then TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("cidadexp"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("ufexp")) Then TxtUF(X).Text = de_informa.rsSelAWB_CTC.Fields("ufexp")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("telexp")) Then TxtTel(X).Text = de_informa.rsSelAWB_CTC.Fields("telexp")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("faxexp")) Then TxtFAX(X).Text = de_informa.rsSelAWB_CTC.Fields("faxexp")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cnpjexp")) Then TxtCGC(X).Text = de_informa.rsSelAWB_CTC.Fields("cnpjexp")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("inscrestexp")) Then TxtInscrEst(X).Text = de_informa.rsSelAWB_CTC.Fields("inscrestexp")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cepexp")) Then TxtCEP(X).Text = de_informa.rsSelAWB_CTC.Fields("cepexp")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("segexp")) Then TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("segexp"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("apoliceexp")) Then TxtApolice(X).Text = de_informa.rsSelAWB_CTC.Fields("apoliceexp")
        
        
        X = 0
        
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("nomedes")) Then TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("nomedes"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("enddes")) Then TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("enddes"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("bairrodes")) Then TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("bairrodes"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cidadedes")) Then TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("cidadedes"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("ufdes")) Then TxtUF(X).Text = de_informa.rsSelAWB_CTC.Fields("ufdes")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("teldes")) Then TxtTel(X).Text = de_informa.rsSelAWB_CTC.Fields("teldes")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("faxdes")) Then TxtFAX(X).Text = de_informa.rsSelAWB_CTC.Fields("faxdes")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cnpjdes")) Then TxtCGC(X).Text = de_informa.rsSelAWB_CTC.Fields("cnpjdes")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("inscrestdes")) Then TxtInscrEst(X).Text = de_informa.rsSelAWB_CTC.Fields("inscrestdes")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("cepdes")) Then TxtCEP(X).Text = de_informa.rsSelAWB_CTC.Fields("cepdes")
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("segdes")) Then TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB_CTC.Fields("segdes"))
        If Not IsNull(de_informa.rsSelAWB_CTC.Fields("apolicedes")) Then TxtApolice(X).Text = de_informa.rsSelAWB_CTC.Fields("apolicedes")
        ElseIf de_informa.rsSelAWB_CTC.RecordCount > 1 Then
        FlexGridNFs.Rows = 0
        FlexGridVolumes.Rows = 0
        Call LimpaTela(Me)
        frmConsultaAWBFiltra.Show 1
        ElseIf de_informa.rsSelAWB_CTC.RecordCount = 0 Then
        MsgBox "CTC não encontrado!", vbCritical, ""
        Exit Sub
        End If
    
    End If
    
TxtICMS.Text = Format(TxtICMS.Text, "#,##0.00")
TxtFreteNacional.Text = Format(TxtFreteNacional.Text, "#,##0.00")
TxtKiloCob.Text = Format(TxtKiloCob.Text, "#,##0.00")
TxtADValorem.Text = Format(TxtADValorem.Text, "#,##0.00")
TxtTXDestino.Text = Format(TxtTXDestino.Text, "#,##0.00")
TxtTXRedesp.Text = Format(TxtTXRedesp.Text, "#,##0.00")
TxtOutros1.Text = Format(TxtOutros1.Text, "#,##0.00")
TxtOutros2.Text = Format(TxtOutros2.Text, "#,##0.00")
TxtFreteTotal.Text = Format(TxtFreteTotal.Text, "#,##0.00")
TxtTotalVM.Text = Format(TxtTotalVM.Text, "#,##0.00")
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



Private Sub CmdSair_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Call OptNF_Click
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

Private Sub OptCTC_Click()
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

Private Sub OptNF_Click()
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


