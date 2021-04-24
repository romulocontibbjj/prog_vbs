VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDinamico 
   Caption         =   "Lista Dinâmica"
   ClientHeight    =   8235
   ClientLeft      =   -3105
   ClientTop       =   1590
   ClientWidth     =   10155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10155
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Seleção de Dados"
      TabPicture(0)   =   "frmDinamico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraPeriodo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Resultados"
      TabPicture(1)   =   "frmDinamico.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         Height          =   1575
         Left            =   120
         TabIndex        =   62
         Top             =   6240
         Width           =   10455
         Begin VB.CheckBox Check19 
            Caption         =   "Status:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Rodoviário"
            Height          =   195
            Left            =   7920
            TabIndex        =   77
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Aéreo"
            Height          =   195
            Left            =   7080
            TabIndex        =   76
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Cancelados"
            Height          =   195
            Left            =   3840
            TabIndex        =   74
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Baixados (s/entr)"
            Height          =   195
            Left            =   3840
            TabIndex        =   73
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Entregues"
            Height          =   195
            Left            =   2520
            TabIndex        =   72
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Em Trânsito"
            Height          =   195
            Left            =   2520
            TabIndex        =   71
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Em Ocorrência"
            Height          =   195
            Left            =   1080
            TabIndex        =   70
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Sem Posição"
            Height          =   195
            Left            =   1080
            TabIndex        =   69
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   4800
            TabIndex        =   68
            Text            =   "Text13"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   2160
            TabIndex        =   66
            Text            =   "Text12"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   960
            TabIndex        =   64
            Text            =   "Text11"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   6480
            TabIndex        =   75
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Prod:"
            Height          =   195
            Left            =   3720
            TabIndex        =   67
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Emissor:"
            Height          =   195
            Left            =   1560
            TabIndex        =   65
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Filial Intec:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "** Localidade"
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
         Left            =   120
         TabIndex        =   50
         Top             =   4800
         Width           =   11655
         Begin VB.Frame Frame9 
            Caption         =   "Região"
            Height          =   855
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   4815
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   3840
               TabIndex        =   61
               Text            =   "Combo4"
               Top             =   360
               Width           =   855
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1080
               TabIndex        =   58
               Text            =   "Combo2"
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "Região SAC:"
               Height          =   255
               Left            =   2760
               TabIndex        =   60
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label4 
               Caption         =   "Geográfica:"
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Estado/UF"
            Height          =   855
            Left            =   5160
            TabIndex        =   51
            Top             =   360
            Width           =   6375
            Begin VB.TextBox Text10 
               Height          =   285
               Left            =   4680
               TabIndex        =   56
               Text            =   "Text10"
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox Check10 
               Caption         =   "Cidade:"
               Height          =   255
               Left            =   3720
               TabIndex        =   55
               Top             =   360
               Width           =   855
            End
            Begin VB.CheckBox Check9 
               Caption         =   "Interior"
               Height          =   255
               Left            =   2760
               TabIndex        =   54
               Top             =   360
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Capital/Metrop."
               Height          =   255
               Left            =   1200
               TabIndex        =   53
               Top             =   360
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   120
               TabIndex        =   52
               Text            =   "Combo2"
               Top             =   360
               Width           =   855
            End
         End
      End
      Begin VB.Frame FraPeriodo 
         Caption         =   "** Período de Emissão"
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
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   11655
         Begin VB.Frame Frame5 
            Caption         =   "Emissão nos Últimos..."
            Height          =   1095
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   2775
            Begin VB.OptionButton Option9 
               Caption         =   "20 dias"
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton Option8 
               Caption         =   "30 dias"
               Height          =   195
               Left            =   960
               TabIndex        =   44
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton Option7 
               Caption         =   "60 dias"
               Height          =   195
               Left            =   1800
               TabIndex        =   43
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton optPer15d 
               Caption         =   "05 dias"
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton opt30d 
               Caption         =   "10 dias"
               Height          =   195
               Left            =   960
               TabIndex        =   36
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton opt60d 
               Caption         =   "15 dias"
               Height          =   195
               Left            =   1800
               TabIndex        =   35
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Por Mês / Por Período (máx. 60 dias)"
            Height          =   1095
            Left            =   3000
            TabIndex        =   30
            Top             =   360
            Width           =   3015
            Begin VB.CommandButton Command6 
               Caption         =   "+"
               Height          =   255
               Left            =   2520
               TabIndex        =   49
               Top             =   360
               Width           =   255
            End
            Begin VB.CommandButton Command5 
               Caption         =   "-"
               Height          =   255
               Left            =   2280
               TabIndex        =   48
               Top             =   360
               Width           =   255
            End
            Begin VB.TextBox Text9 
               Height          =   285
               Left            =   1680
               TabIndex        =   47
               Text            =   "2002"
               Top             =   360
               Width           =   495
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   120
               TabIndex        =   46
               Text            =   "Por Mês"
               Top             =   360
               Width           =   1335
            End
            Begin MSMask.MaskEdBox mskPer2 
               Height          =   285
               Left            =   1680
               TabIndex        =   31
               Top             =   720
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
               TabIndex        =   32
               Top             =   720
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
               TabIndex        =   33
               Top             =   720
               Width           =   90
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Nos Dias da Semana / Horário de Emissão"
            Height          =   1095
            Left            =   6480
            TabIndex        =   22
            Top             =   360
            Width           =   5055
            Begin VB.TextBox Text8 
               Height          =   285
               Left            =   2760
               TabIndex        =   40
               Text            =   "Text8"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   1680
               TabIndex        =   39
               Text            =   "Text1"
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Seg"
               Height          =   195
               Left            =   840
               TabIndex        =   29
               Top             =   360
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Ter"
               Height          =   195
               Left            =   1560
               TabIndex        =   28
               Top             =   360
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Qua"
               Height          =   195
               Left            =   2160
               TabIndex        =   27
               Top             =   360
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Qui"
               Height          =   195
               Left            =   2880
               TabIndex        =   26
               Top             =   360
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Sex"
               Height          =   195
               Left            =   3480
               TabIndex        =   25
               Top             =   360
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Sab"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   4200
               TabIndex        =   24
               Top             =   360
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox Check7 
               Caption         =   "Dom"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   360
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "horas."
               Height          =   195
               Left            =   3480
               TabIndex        =   42
               Top             =   720
               Width           =   435
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "às"
               Height          =   195
               Left            =   2400
               TabIndex        =   41
               Top             =   720
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Horário de Emissão:"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   720
               Width           =   1410
            End
         End
         Begin VB.Line Line1 
            X1              =   6240
            X2              =   6240
            Y1              =   240
            Y2              =   1440
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "** Consignatário"
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
         TabIndex        =   15
         Top             =   2280
         Width           =   11655
         Begin VB.CommandButton Command4 
            Caption         =   "Busca CGC..."
            Height          =   255
            Left            =   3600
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   6960
            TabIndex        =   19
            Text            =   "Text3"
            Top             =   360
            Width           =   4215
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1560
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Por Nome:"
            Height          =   195
            Left            =   5760
            TabIndex        =   17
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Por CGC:"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "** Destinatário"
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
         TabIndex        =   9
         Top             =   1440
         Width           =   11655
         Begin VB.OptionButton Option4 
            Caption         =   "Por CGC:"
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Por Nome:"
            Height          =   195
            Left            =   5760
            TabIndex        =   13
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1560
            TabIndex        =   12
            Text            =   "Text2"
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   6960
            TabIndex        =   11
            Text            =   "Text3"
            Top             =   360
            Width           =   4215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Busca CGC..."
            Height          =   255
            Left            =   3600
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "** Remetente"
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
         TabIndex        =   3
         Top             =   600
         Width           =   11655
         Begin VB.CommandButton Command2 
            Caption         =   "Busca CGC..."
            Height          =   255
            Left            =   3600
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   6960
            TabIndex        =   7
            Text            =   "Text3"
            Top             =   360
            Width           =   4215
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1560
            TabIndex        =   6
            Text            =   "Text2"
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Por Nome:"
            Height          =   195
            Left            =   5760
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por CGC:"
            Height          =   195
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   1
         Top             =   600
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10186
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Dados Sintéticos - Totais"
         TabPicture(0)   =   "frmDinamico.frx":0038
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Dados Analíticos - Listagem"
         TabPicture(1)   =   "frmDinamico.frx":0054
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Command1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.CommandButton Command1 
            Caption         =   "Gerar Arquivo..."
            Height          =   375
            Left            =   8640
            TabIndex        =   2
            Top             =   480
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "frmDinamico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set frmDinamico = Nothing
End Sub
