VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCadUsu 
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   8325
   ClientLeft      =   465
   ClientTop       =   630
   ClientWidth     =   12075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   12075
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
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
      Height          =   2775
      Left            =   9480
      TabIndex        =   15
      Top             =   240
      Width           =   2415
      Begin VB.CommandButton cmdNovoUsu 
         Caption         =   "Novo Usuário"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdAltUsu 
         Caption         =   "Alterar Dados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Frame fraUsuarios 
      Caption         =   "Usuários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   2415
      Begin MSDataGridLib.DataGrid GridUsuario 
         Bindings        =   "frmCadUsu.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
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
         DataMember      =   "Sel_UsuariosTodos"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
            DataField       =   "senha"
            Caption         =   "senha"
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
            DataField       =   "Nome"
            Caption         =   "Nome"
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
         BeginProperty Column03 
            DataField       =   "Departamento"
            Caption         =   "Departamento"
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
         BeginProperty Column04 
            DataField       =   "DataCad"
            Caption         =   "DataCad"
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
         BeginProperty Column05 
            DataField       =   "status"
            Caption         =   "status"
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
         BeginProperty Column06 
            DataField       =   "stringdireitos"
            Caption         =   "stringdireitos"
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
            MarqueeStyle    =   3
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1620,284
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   464,882
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraDireitos 
      Caption         =   "Direitos do Usuário"
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
      Height          =   5175
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   11775
      Begin TabDlg.SSTab SSTabSistemas 
         Height          =   4815
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Informa"
         TabPicture(0)   =   "frmCadUsu.frx":0019
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FRADireitosInforma3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FRADireitosInforma2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FRADireitosInforma1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Faturamento"
         TabPicture(1)   =   "frmCadUsu.frx":0035
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FRADireitosFaturamento1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "FRADireitosFaturamento2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "FRADireitosFaturamento3"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Emissão"
         TabPicture(2)   =   "frmCadUsu.frx":0051
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label53"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Aéreo"
         TabPicture(3)   =   "frmCadUsu.frx":006D
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label93"
         Tab(3).ControlCount=   1
         Begin VB.Frame FRADireitosFaturamento3 
            Height          =   4215
            Left            =   -66960
            TabIndex        =   157
            Top             =   360
            Width           =   3375
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Sistema Faturamento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   435
               Index           =   86
               Left            =   360
               TabIndex        =   164
               Top             =   1440
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Impressão de Etiquetas"
               Height          =   195
               Index           =   85
               Left            =   360
               TabIndex        =   162
               Top             =   1200
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Movimentação não Faturado"
               Height          =   195
               Index           =   84
               Left            =   360
               TabIndex        =   161
               Top             =   960
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Faturas em Aberto"
               Height          =   195
               Index           =   83
               Left            =   360
               TabIndex        =   160
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Relatórios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   195
               Index           =   81
               Left            =   360
               TabIndex        =   159
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Faturamento"
               Height          =   195
               Index           =   82
               Left            =   360
               TabIndex        =   158
               Top             =   480
               Width           =   2535
            End
            Begin VB.Label Label92 
               AutoSize        =   -1  'True
               Caption         =   "81"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   203
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               Caption         =   "86"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   202
               Top             =   1560
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label90 
               AutoSize        =   -1  'True
               Caption         =   "85"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   201
               Top             =   1200
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label89 
               AutoSize        =   -1  'True
               Caption         =   "84"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   200
               Top             =   960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label88 
               AutoSize        =   -1  'True
               Caption         =   "83"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   199
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               Caption         =   "82"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   198
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosFaturamento2 
            Height          =   4215
            Left            =   -71160
            TabIndex        =   140
            Top             =   360
            Width           =   4095
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Imprimir Faturas"
               Height          =   195
               Index           =   80
               Left            =   360
               TabIndex        =   156
               Top             =   3840
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Impressão de Faturas"
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
               Index           =   79
               Left            =   360
               TabIndex        =   155
               Top             =   3600
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Efetuar Cancelar Faturas"
               Height          =   195
               Index           =   78
               Left            =   360
               TabIndex        =   154
               Top             =   3360
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cancelar Faturas"
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
               Index           =   77
               Left            =   360
               TabIndex        =   153
               Top             =   3120
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Visualizar Todas Pré-Faturas"
               Height          =   195
               Index           =   66
               Left            =   360
               TabIndex        =   152
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Pré-Faturas Pendentes / Gerar Fatura"
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
               Index           =   65
               Left            =   360
               TabIndex        =   151
               Top             =   240
               Width           =   3615
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Visualizar Somente as Minhas"
               Height          =   195
               Index           =   67
               Left            =   360
               TabIndex        =   150
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar Vencimento Pré Fatura"
               Height          =   195
               Index           =   68
               Left            =   360
               TabIndex        =   149
               Top             =   960
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Consultar e Alterar Pré Fatura"
               Height          =   195
               Index           =   69
               Left            =   360
               TabIndex        =   148
               Top             =   1200
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir Pré Fatura Pendente"
               Height          =   195
               Index           =   70
               Left            =   360
               TabIndex        =   147
               Top             =   1440
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar Fatura Final"
               Height          =   195
               Index           =   71
               Left            =   360
               TabIndex        =   146
               Top             =   1680
               Width           =   1815
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Prorrogar Vencimentos"
               Height          =   195
               Index           =   75
               Left            =   360
               TabIndex        =   145
               Top             =   2640
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Conceder descontos/Abatimentos"
               Height          =   195
               Index           =   74
               Left            =   360
               TabIndex        =   144
               Top             =   2400
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Consulta Fat. ( Prorr. / Quita. / Abat. )"
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
               Index           =   72
               Left            =   360
               TabIndex        =   143
               Top             =   1920
               Width           =   3615
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar endereços de cobrança"
               Height          =   195
               Index           =   73
               Left            =   360
               TabIndex        =   142
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Quitar Fatura"
               Height          =   195
               Index           =   76
               Left            =   360
               TabIndex        =   141
               Top             =   2880
               Width           =   2295
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "80"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   197
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label85 
               AutoSize        =   -1  'True
               Caption         =   "65"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   196
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "79"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   195
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label83 
               AutoSize        =   -1  'True
               Caption         =   "78"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   194
               Top             =   3360
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               Caption         =   "77"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   193
               Top             =   3120
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               Caption         =   "76"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   192
               Top             =   2880
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label80 
               AutoSize        =   -1  'True
               Caption         =   "75"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   191
               Top             =   2640
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label79 
               AutoSize        =   -1  'True
               Caption         =   "74"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   190
               Top             =   2400
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label78 
               AutoSize        =   -1  'True
               Caption         =   "73"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   189
               Top             =   2160
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label77 
               AutoSize        =   -1  'True
               Caption         =   "72"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   188
               Top             =   1920
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label76 
               AutoSize        =   -1  'True
               Caption         =   "71"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   187
               Top             =   1680
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label75 
               AutoSize        =   -1  'True
               Caption         =   "70"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   186
               Top             =   1440
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label74 
               AutoSize        =   -1  'True
               Caption         =   "69"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   185
               Top             =   1200
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label73 
               AutoSize        =   -1  'True
               Caption         =   "68"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   184
               Top             =   960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               Caption         =   "67"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   183
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               Caption         =   "66"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   182
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosFaturamento1 
            Height          =   4215
            Left            =   -74880
            TabIndex        =   123
            Top             =   360
            Width           =   3615
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir Tudo"
               Height          =   195
               Index           =   63
               Left            =   360
               TabIndex        =   139
               Top             =   3720
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar Fatura"
               Height          =   195
               Index           =   64
               Left            =   360
               TabIndex        =   138
               Top             =   3960
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar endereços de cobrança"
               Height          =   195
               Index           =   60
               Left            =   360
               TabIndex        =   137
               Top             =   3000
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Pré-Fatura Consulta / Alteração"
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
               Index           =   59
               Left            =   360
               TabIndex        =   136
               Top             =   2760
               Width           =   3135
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir CTC/NFS"
               Height          =   195
               Index           =   61
               Left            =   360
               TabIndex        =   135
               Top             =   3240
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir CTC"
               Height          =   195
               Index           =   62
               Left            =   360
               TabIndex        =   134
               Top             =   3480
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar Fatura"
               Height          =   195
               Index           =   58
               Left            =   360
               TabIndex        =   133
               Top             =   2520
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar Fatura Avulsa"
               Height          =   195
               Index           =   57
               Left            =   360
               TabIndex        =   132
               Top             =   2280
               Width           =   1815
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir Tudo"
               Height          =   195
               Index           =   56
               Left            =   360
               TabIndex        =   131
               Top             =   2040
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir CTC"
               Height          =   195
               Index           =   55
               Left            =   360
               TabIndex        =   130
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir CTC's ( Por Intervalo )"
               Height          =   195
               Index           =   54
               Left            =   360
               TabIndex        =   129
               Top             =   1560
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir CTC/NFS ( Manual )"
               Height          =   195
               Index           =   53
               Left            =   360
               TabIndex        =   128
               Top             =   1320
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova Pré-Fatura / Fatura Avulsa"
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
               Index           =   51
               Left            =   360
               TabIndex        =   127
               Top             =   840
               Width           =   3135
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar endereços de cobrança"
               Height          =   195
               Index           =   52
               Left            =   360
               TabIndex        =   126
               Top             =   1080
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Faturamento / Cobrança"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   195
               Index           =   50
               Left            =   360
               TabIndex        =   125
               Top             =   600
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Arquivos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   195
               Index           =   49
               Left            =   360
               TabIndex        =   124
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               Caption         =   "50"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   181
               Top             =   600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               Caption         =   "64"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   180
               Top             =   3960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               Caption         =   "63"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   179
               Top             =   3720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               Caption         =   "62"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   178
               Top             =   3480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               Caption         =   "61"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   177
               Top             =   3240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               Caption         =   "60"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   176
               Top             =   3000
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "59"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   175
               Top             =   2760
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               Caption         =   "58"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   174
               Top             =   2520
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               Caption         =   "57"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   173
               Top             =   2280
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               Caption         =   "56"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   172
               Top             =   2040
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "55"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   171
               Top             =   1800
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "54"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   170
               Top             =   1560
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               Caption         =   "53"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   169
               Top             =   1320
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "52"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   168
               Top             =   1080
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "51"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   167
               Top             =   840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   166
               Top             =   600
               Visible         =   0   'False
               Width           =   45
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "49"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   165
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosInforma1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   120
            TabIndex        =   90
            Top             =   480
            Width           =   3495
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - Import. EDI - Ocorr/POD"
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   106
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - Exportação para o SITLA"
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   105
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Altera Dados Clientes"
               Height          =   195
               Index           =   6
               Left            =   360
               TabIndex        =   104
               Top             =   1920
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Incl. / Alt. Feriados"
               Height          =   195
               Index           =   7
               Left            =   360
               TabIndex        =   103
               Top             =   2160
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Inclui Novos Clientes"
               Height          =   195
               Index           =   5
               Left            =   360
               TabIndex        =   102
               Top             =   1680
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Usuários"
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   101
               Top             =   1440
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Incl/Alt. Cod. Ocorrência"
               Height          =   195
               Index           =   3
               Left            =   360
               TabIndex        =   100
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Prazos de Entrega"
               Height          =   195
               Index           =   8
               Left            =   360
               TabIndex        =   99
               Top             =   2400
               Width           =   2655
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - LOG de Usuários"
               Height          =   195
               Index           =   26
               Left            =   360
               TabIndex        =   98
               Top             =   960
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - Export. EDI - Ocoren."
               Height          =   195
               Index           =   25
               Left            =   360
               TabIndex        =   97
               Top             =   480
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) Acompanhamento"
               Height          =   195
               Index           =   34
               Left            =   360
               TabIndex        =   96
               Top             =   2880
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   37
               Left            =   360
               TabIndex        =   95
               Top             =   3120
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   40
               Left            =   360
               TabIndex        =   94
               Top             =   3360
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   43
               Left            =   360
               TabIndex        =   93
               Top             =   3600
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   46
               Left            =   360
               TabIndex        =   92
               Top             =   3840
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   47
               Left            =   360
               TabIndex        =   91
               Top             =   2640
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "1"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   122
               Top             =   240
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "2"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   121
               Top             =   720
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "7"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   120
               Top             =   2160
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "3"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   119
               Top             =   1200
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "4"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   118
               Top             =   1440
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "5"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   117
               Top             =   1680
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "6"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   116
               Top             =   1920
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "8"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   115
               Top             =   2400
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "26"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   114
               Top             =   960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "25"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   113
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "34"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   112
               Top             =   2880
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "37"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   111
               Top             =   3120
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "40"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   110
               Top             =   3360
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "43"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   109
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "46"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   108
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               Caption         =   "47"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   107
               Top             =   2640
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosInforma2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   3720
            TabIndex        =   57
            Top             =   480
            Width           =   3495
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Ocorrências e POD"
               Height          =   195
               Index           =   11
               Left            =   360
               TabIndex        =   73
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Acomp. de Clientes"
               Height          =   195
               Index           =   15
               Left            =   360
               TabIndex        =   72
               Top             =   960
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Arq. Exclusivo Cliente"
               Height          =   195
               Index           =   17
               Left            =   360
               TabIndex        =   71
               Top             =   2400
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Consulta SAC"
               Height          =   195
               Index           =   10
               Left            =   360
               TabIndex        =   70
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Novos Prazos de Entr."
               Height          =   195
               Index           =   9
               Left            =   360
               TabIndex        =   69
               Top             =   240
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Acomp. Resumo"
               Height          =   195
               Index           =   27
               Left            =   360
               TabIndex        =   68
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Controle dos Canhotos"
               Height          =   195
               Index           =   28
               Left            =   360
               TabIndex        =   67
               Top             =   1440
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Controle de Devoluções"
               Height          =   195
               Index           =   29
               Left            =   360
               TabIndex        =   66
               Top             =   1680
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Alarme de Urgências"
               Height          =   195
               Index           =   24
               Left            =   360
               TabIndex        =   65
               Top             =   1920
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Informações via Email"
               Height          =   195
               Index           =   14
               Left            =   360
               TabIndex        =   64
               Top             =   2640
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) Ordem de Coleta"
               Height          =   195
               Index           =   32
               Left            =   360
               TabIndex        =   63
               Top             =   2880
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) Cancelamento"
               Height          =   195
               Index           =   35
               Left            =   360
               TabIndex        =   62
               Top             =   3120
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   38
               Left            =   360
               TabIndex        =   61
               Top             =   3360
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   41
               Left            =   360
               TabIndex        =   60
               Top             =   3600
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   44
               Left            =   360
               TabIndex        =   59
               Top             =   3840
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Alarme (GERENCIAL)"
               Height          =   195
               Index           =   30
               Left            =   360
               TabIndex        =   58
               Top             =   2160
               Width           =   3015
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "11"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   89
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "15"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   88
               Top             =   960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "17"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   87
               Top             =   2400
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "10"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   86
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "9"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   85
               Top             =   240
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "29"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   84
               Top             =   1680
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "28"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   1440
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "27"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   82
               Top             =   1200
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "24"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   81
               Top             =   1920
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "14"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   2640
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "32"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   79
               Top             =   2880
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "35"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   3120
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "38"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   77
               Top             =   3360
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "41"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   76
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               Caption         =   "44"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   75
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "30"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   74
               Top             =   2160
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosInforma3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   7320
            TabIndex        =   24
            Top             =   480
            Width           =   3495
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Excl. de Ocorrências"
               Height          =   195
               Index           =   22
               Left            =   360
               TabIndex        =   40
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Exclusão de PODs"
               Height          =   195
               Index           =   23
               Left            =   360
               TabIndex        =   39
               Top             =   480
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - Análise Estatística"
               Height          =   195
               Index           =   20
               Left            =   360
               TabIndex        =   38
               Top             =   1440
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - Análise de Ocorrências"
               Height          =   195
               Index           =   19
               Left            =   360
               TabIndex        =   37
               Top             =   2160
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - Análise de Entregas"
               Height          =   195
               Index           =   18
               Left            =   360
               TabIndex        =   36
               Top             =   1680
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Recalc. Prev. Entrega"
               Height          =   195
               Index           =   13
               Left            =   360
               TabIndex        =   35
               Top             =   960
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Recalc. Prazos de Entr."
               Height          =   195
               Index           =   12
               Left            =   360
               TabIndex        =   34
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Cancelar CTC"
               Height          =   195
               Index           =   16
               Left            =   360
               TabIndex        =   33
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Relatórios) - Protocolo para Arquivo"
               Height          =   195
               Index           =   21
               Left            =   360
               TabIndex        =   32
               Top             =   2400
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) POD"
               Height          =   195
               Index           =   33
               Left            =   360
               TabIndex        =   31
               Top             =   2640
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   36
               Left            =   360
               TabIndex        =   30
               Top             =   2880
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   39
               Left            =   360
               TabIndex        =   29
               Top             =   3120
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   42
               Left            =   360
               TabIndex        =   28
               Top             =   3360
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   45
               Left            =   360
               TabIndex        =   27
               Top             =   3600
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   48
               Left            =   360
               TabIndex        =   26
               Top             =   3840
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - An.Entregas - Abono"
               Height          =   195
               Index           =   31
               Left            =   360
               TabIndex        =   25
               Top             =   1920
               Width           =   3015
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "22"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "23"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   55
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "20"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   1440
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "19"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   53
               Top             =   2160
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "18"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   52
               Top             =   1680
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "13"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "12"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "16"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   49
               Top             =   1200
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "21"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   48
               Top             =   2400
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "33"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   2640
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "36"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   2880
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "39"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   3120
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "42"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   3360
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               Caption         =   "45"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "48"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               Caption         =   "31"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   1920
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Label Label93 
            Caption         =   "Em Análise"
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
            Left            =   -70320
            TabIndex        =   204
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label53 
            Caption         =   "Em Análise"
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
            Left            =   -70320
            TabIndex        =   163
            Top             =   2160
            Width           =   975
         End
      End
   End
   Begin VB.Frame FraDadosUsu 
      Caption         =   "Dados de Usuários"
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
      Height          =   2775
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      Begin VB.CheckBox Check1 
         Caption         =   "Enviar Email Informando Futuros Feriados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Frame Frame3 
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
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2415
         Begin VB.OptionButton optStatusInativo 
            Caption         =   "Inativo"
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optStatusAtivo 
            Caption         =   "Ativo"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtDepartamento 
         Height          =   285
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   8
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblOpcao 
         AutoSize        =   -1  'True
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
         Left            =   4200
         TabIndex        =   21
         Top             =   360
         Width           =   75
      End
      Begin VB.Label lblDataCad 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data Cadastro:"
         Height          =   195
         Left            =   3480
         TabIndex        =   4
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome Completo:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmCadUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAltUsu_Click()
    Dim X As Integer
    
    FRADireitosFaturamento1.Enabled = True
    FRADireitosFaturamento2.Enabled = True
    FRADireitosFaturamento3.Enabled = True
    
    FRADireitosInforma1.Enabled = True
    FRADireitosInforma2.Enabled = True
    FRADireitosInforma3.Enabled = True
    
    
    fraUsuarios.Enabled = False
    fraDireitos.Enabled = True
    SSTabSistemas.Enabled = True
    FraDadosUsu.Enabled = True
    cmdNovoUsu.Enabled = False
    cmdAltUsu.Enabled = False
    cmdGravar.Enabled = True
    CmdCancelar.Enabled = True
    cmdSair.Enabled = False
    TxtUsuario.Enabled = False
    txtNome.BackColor = &HC0FFFF     'AMARELO
    txtDepartamento.BackColor = &HC0FFFF     'AMARELO
    txtNome.SetFocus
    lblOpcao = "ALTERAÇÃO"
End Sub

Private Sub cmdCancelar_Click()
    TxtUsuario.Enabled = True
    fraUsuarios.Enabled = True
    fraDireitos.Enabled = True
    
    FRADireitosFaturamento1.Enabled = False
    FRADireitosFaturamento2.Enabled = False
    FRADireitosFaturamento3.Enabled = False
    
    FRADireitosInforma1.Enabled = False
    FRADireitosInforma2.Enabled = False
    FRADireitosInforma3.Enabled = False
    
    FraDadosUsu.Enabled = False
    cmdNovoUsu.Enabled = True
    cmdAltUsu.Enabled = False
    cmdGravar.Enabled = False
    CmdCancelar.Enabled = False
    cmdSair.Enabled = True

    TxtUsuario.BackColor = &H8000000E       'BRANCO
    txtNome.BackColor = &H8000000E       'BRANCO
    txtDepartamento.BackColor = &H8000000E      'BRANCO
    cmdNovoUsu.SetFocus
    lblOpcao = ""
End Sub

Private Sub cmdGravar_Click()
    Dim xStatus As String, xstrdireitos As String, X As Integer
    If Len(Trim$(TxtUsuario.Text)) < 6 Then
        MsgBox "O Nome do USUÁRIO deve ter no mínimo 6 caracteres."
        txtNome.SetFocus
        Exit Sub
    End If
    If Len(txtNome.Text) < 5 Then
        MsgBox "Nome Completo Inválido ! "
        txtNome.SetFocus
        Exit Sub
    End If
    If optStatusAtivo.Value = True Then
        xStatus = "1"
    Else
        xStatus = "0"
    End If
    xstrdireitos = ""
    For X = 1 To 86
        If chkDireitos(X).Value = 1 Then
            xstrdireitos = xstrdireitos + "1"
        Else
            xstrdireitos = xstrdireitos + "0"
        End If
    Next
    If lblOpcao = "INCLUSÃO" Then
        If de_informa.rsSel_Usuario.State = 1 Then de_informa.rsSel_Usuario.Close
        de_informa.Sel_Usuario Trim$(TxtUsuario)
        If de_informa.rsSel_Usuario.RecordCount > 0 Then
            MsgBox "USUÁRIO já cadastrado !"
            TxtUsuario.SetFocus
            Exit Sub
        Else
            de_informa.Ins_cadUsu TxtUsuario, TxtUsuario, txtNome, txtDepartamento, lblDataCad, xStatus, xstrdireitos, "S"
        End If
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "CAD. DE USUÁRIOS: " & TxtUsuario
        
    Else
        de_informa.alt_cadusu TxtUsuario, txtNome, txtDepartamento, xStatus, xstrdireitos
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "CAD. DE USUÁRIOS: " & TxtUsuario
        
    End If
    If de_informa.rsSel_UsuariosTodos.State = 1 Then de_informa.rsSel_UsuariosTodos.Close
    de_informa.Sel_UsuariosTodos
    GridUsuario.DataMember = "sel_usuariostodos"
    GridUsuario.Refresh
    lblOpcao = ""
    cmdCancelar_Click
End Sub
Private Sub cmdNovoUsu_Click()
    Dim X As Integer
    
    FRADireitosFaturamento1.Enabled = True
    FRADireitosFaturamento2.Enabled = True
    FRADireitosFaturamento3.Enabled = True
    
    FRADireitosInforma1.Enabled = True
    FRADireitosInforma2.Enabled = True
    FRADireitosInforma3.Enabled = True
    
    fraUsuarios.Enabled = False
    fraDireitos.Enabled = True
    FraDadosUsu.Enabled = True
    cmdNovoUsu.Enabled = False
    cmdAltUsu.Enabled = False
    cmdGravar.Enabled = True
    CmdCancelar.Enabled = True
    lblDataCad = datahora("data")
    cmdSair.Enabled = False
    TxtUsuario.BackColor = &HC0FFFF      'AMARELO
    txtNome.BackColor = &HC0FFFF     'AMARELO
    txtDepartamento.BackColor = &HC0FFFF     'AMARELO
    For X = 0 To frmCadUsu.Controls.Count - 1
        If TypeOf frmCadUsu.Controls(X) Is TextBox Then
            frmCadUsu.Controls(X).Text = ""
        ElseIf TypeOf frmCadUsu.Controls(X) Is CheckBox Then
            frmCadUsu.Controls(X).Value = 0
        End If
    Next
    optStatusAtivo = True
    TxtUsuario.SetFocus
    lblOpcao = "INCLUSÃO"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    cmdNovoUsu.SetFocus
    fraDireitos.Enabled = True
    
    FRADireitosFaturamento1.Enabled = False
    FRADireitosFaturamento2.Enabled = False
    FRADireitosFaturamento3.Enabled = False
    
    FRADireitosInforma1.Enabled = False
    FRADireitosInforma2.Enabled = False
    FRADireitosInforma3.Enabled = False
    
End Sub
Private Sub Form_Load()
    If de_informa.rsSel_UsuariosTodos.State = 1 Then de_informa.rsSel_UsuariosTodos.Close
    de_informa.Sel_UsuariosTodos
    GridUsuario.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCadUsu = Nothing
End Sub

Private Sub GridUsuario_Click()
    Dim X As Integer
    
    TxtUsuario = GridUsuario.Columns(0)
    txtNome = GridUsuario.Columns(2)
    txtDepartamento = GridUsuario.Columns(3)
    lblDataCad = GridUsuario.Columns(4)
    
    SSTabSistemas.Visible = True
    
    If GridUsuario.Columns(5) = "0" Then
        optStatusInativo = True
    Else
        optStatusAtivo = True
    End If
    
    For X = 1 To 86
        If Mid$(GridUsuario.Columns(6), X, 1) = "1" Then
            chkDireitos(X).Value = 1
        Else
            chkDireitos(X).Value = 0
        End If
    Next
    cmdAltUsu.Enabled = True
    SSTabSistemas.Visible = True
    
End Sub

Private Sub GridUsuario_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
GridUsuario_Click
End Sub

Private Sub SSTab1_DblClick()

End Sub


Private Sub txtDepartamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtDepartamento_LostFocus()
    txtDepartamento = UCase(txtDepartamento)
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtNome_LostFocus()
    txtNome = UCase(txtNome)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUsuario_LostFocus()
    TxtUsuario = UCase(TxtUsuario)
End Sub
