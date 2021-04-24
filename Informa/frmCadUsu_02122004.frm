VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadUsu 
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   8325
   ClientLeft      =   465
   ClientTop       =   630
   ClientWidth     =   12390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   12390
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
      TabIndex        =   73
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton cmdNovoUsu 
         Caption         =   "Novo Usuário"
         Height          =   375
         Left            =   480
         TabIndex        =   45
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdAltUsu 
         Caption         =   "Alterar Dados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   43
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   480
         TabIndex        =   47
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   44
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   46
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
      TabIndex        =   71
      Top             =   240
      Width           =   2295
      Begin MSDataGridLib.DataGrid GridUsuario 
         Bindings        =   "frmCadUsu.frx":0000
         Height          =   2295
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
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
         ColumnCount     =   9
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
            DataField       =   "Filial"
            Caption         =   "Filial"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "expirada"
            Caption         =   "expirada"
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
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   494,929
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   675,213
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
      TabIndex        =   69
      Top             =   3000
      Width           =   12135
      Begin VB.CommandButton cmd_limpar 
         Caption         =   "L i m p a r"
         Height          =   1215
         Left            =   11760
         TabIndex        =   445
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11760
         TabIndex        =   307
         Top             =   4320
         Width           =   255
      End
      Begin VB.CommandButton Command1 
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
         Height          =   315
         Left            =   11760
         TabIndex        =   306
         Top             =   4680
         Width           =   255
      End
      Begin TabDlg.SSTab SSTabSistemas 
         Height          =   4815
         Left            =   120
         TabIndex        =   36
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
         Tab(1).Control(0)=   "FRADireitosFaturamento3"
         Tab(1).Control(1)=   "FRADireitosFaturamento1"
         Tab(1).Control(2)=   "FRADireitosFaturamento2"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Emissão 1/2"
         TabPicture(2)   =   "frmCadUsu.frx":0051
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FRADireitosEmissao3"
         Tab(2).Control(1)=   "FRADireitosEmissao2"
         Tab(2).Control(2)=   "FRADireitosEmissao1"
         Tab(2).Control(3)=   "Label93"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Emissão 2/2"
         TabPicture(3)   =   "frmCadUsu.frx":006D
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FRADireitosEmissao6"
         Tab(3).Control(1)=   "FRADireitosEmissao5"
         Tab(3).Control(2)=   "FRADireitosEmissao4"
         Tab(3).ControlCount=   3
         Begin VB.Frame FRADireitosEmissao6 
            Height          =   4215
            Left            =   -67080
            TabIndex        =   290
            Top             =   480
            Width           =   3495
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Sistema Emissão"
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
               Index           =   166
               Left            =   480
               TabIndex        =   304
               Top             =   1080
               Width           =   2055
            End
            Begin VB.ComboBox cbPerfil 
               Height          =   315
               ItemData        =   "frmCadUsu.frx":0089
               Left            =   1320
               List            =   "frmCadUsu.frx":008B
               TabIndex        =   297
               Text            =   "Perfil de Usuário"
               Top             =   3720
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Impressão de CTC em Branco"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   164
               Left            =   480
               TabIndex        =   293
               Top             =   480
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Movimento de Emissão"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   163
               Left            =   480
               TabIndex        =   292
               Top             =   240
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Simulado de Frete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   165
               Left            =   480
               TabIndex        =   291
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label Label131 
               AutoSize        =   -1  'True
               Caption         =   "166"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   305
               Top             =   1080
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label176 
               AutoSize        =   -1  'True
               Caption         =   "163"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   296
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label177 
               AutoSize        =   -1  'True
               Caption         =   "164"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   295
               Top             =   480
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label178 
               AutoSize        =   -1  'True
               Caption         =   "165"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   294
               Top             =   720
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosEmissao5 
            Height          =   4215
            Left            =   -71520
            TabIndex        =   263
            Top             =   480
            Width           =   4335
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Produtividade dos Emissores"
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
               Index           =   162
               Left            =   480
               TabIndex        =   302
               Top             =   3840
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cancelar NFS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   156
               Left            =   480
               TabIndex        =   278
               Top             =   2400
               Width           =   1575
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Reimpressão de Notas Fiscais de Serv."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   155
               Left            =   480
               TabIndex        =   277
               Top             =   2160
               Width           =   3735
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar / Gerar NF Serv."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   153
               Left            =   480
               TabIndex        =   276
               Top             =   1680
               Width           =   2055
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Notas Fiscais de Serviços Emitidas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   154
               Left            =   480
               TabIndex        =   275
               Top             =   1920
               Width           =   3495
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar NFS de CTR Identificados"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   152
               Left            =   480
               TabIndex        =   274
               Top             =   1440
               Width           =   2655
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar NF de Serv. Manual(CTR) Avulso"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   151
               Left            =   480
               TabIndex        =   273
               Top             =   1200
               Width           =   3255
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emitir Notas Fiscais de Serv."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   149
               Left            =   480
               TabIndex        =   272
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar NF de Serv. de CTR Ident. Filial"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   150
               Left            =   480
               TabIndex        =   271
               Top             =   960
               Width           =   3135
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Reimpressão de Manifesto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   147
               Left            =   480
               TabIndex        =   270
               Top             =   240
               Width           =   2655
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cancelar Manifesto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   148
               Left            =   480
               TabIndex        =   269
               Top             =   480
               Width           =   2055
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Controle de Numeração"
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
               Index           =   157
               Left            =   480
               TabIndex        =   268
               Top             =   2640
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
               Index           =   158
               Left            =   480
               TabIndex        =   267
               Top             =   2880
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Minuta / CTC Emitidos"
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
               Index           =   159
               Left            =   480
               TabIndex        =   266
               Top             =   3120
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Manifestos Emitidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   160
               Left            =   480
               TabIndex        =   265
               Top             =   3360
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "NF Emitidas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   161
               Left            =   480
               TabIndex        =   264
               Top             =   3600
               Width           =   1575
            End
            Begin VB.Label Label175 
               AutoSize        =   -1  'True
               Caption         =   "162"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   303
               Top             =   3840
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label171 
               AutoSize        =   -1  'True
               Caption         =   "158"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   289
               Top             =   2880
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label170 
               AutoSize        =   -1  'True
               Caption         =   "157"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   288
               Top             =   2640
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label169 
               AutoSize        =   -1  'True
               Caption         =   "156"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   287
               Top             =   2400
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label168 
               AutoSize        =   -1  'True
               Caption         =   "155"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   286
               Top             =   2160
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label167 
               AutoSize        =   -1  'True
               Caption         =   "154"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   285
               Top             =   1920
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label162 
               AutoSize        =   -1  'True
               Caption         =   "149"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   284
               Top             =   720
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label161 
               AutoSize        =   -1  'True
               Caption         =   "148"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   283
               Top             =   480
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label160 
               AutoSize        =   -1  'True
               Caption         =   "147"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   282
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label158 
               AutoSize        =   -1  'True
               Caption         =   "159"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   281
               Top             =   3120
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label157 
               AutoSize        =   -1  'True
               Caption         =   "161"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   280
               Top             =   3600
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "160"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   279
               Top             =   3360
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosFaturamento3 
            Height          =   4095
            Left            =   -66960
            TabIndex        =   238
            Top             =   600
            Width           =   3375
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Faturamento"
               Height          =   195
               Index           =   82
               Left            =   360
               TabIndex        =   245
               Top             =   480
               Width           =   2535
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
               TabIndex        =   244
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Faturas em Aberto"
               Height          =   195
               Index           =   83
               Left            =   360
               TabIndex        =   243
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Movimentação não Faturado"
               Height          =   195
               Index           =   84
               Left            =   360
               TabIndex        =   242
               Top             =   960
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Impressão de Etiquetas"
               Height          =   195
               Index           =   85
               Left            =   360
               TabIndex        =   241
               Top             =   1200
               Width           =   2415
            End
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
               TabIndex        =   240
               Top             =   1440
               Width           =   2175
            End
            Begin VB.ComboBox cb_PerfilFat 
               Height          =   315
               ItemData        =   "frmCadUsu.frx":008D
               Left            =   1200
               List            =   "frmCadUsu.frx":008F
               TabIndex        =   239
               Text            =   "Perfil de Usuário"
               Top             =   3600
               Width           =   1935
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               Caption         =   "86"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   247
               Top             =   1560
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label92 
               AutoSize        =   -1  'True
               Caption         =   "81"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   246
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosFaturamento1 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   216
            Top             =   600
            Width           =   3615
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
               TabIndex        =   232
               Top             =   240
               Width           =   1095
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
               TabIndex        =   231
               Top             =   480
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar endereços de cobrança"
               Height          =   195
               Index           =   52
               Left            =   360
               TabIndex        =   230
               Top             =   960
               Width           =   2535
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
               TabIndex        =   229
               Top             =   720
               Width           =   3135
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir CTC/NFS ( Manual )"
               Height          =   195
               Index           =   53
               Left            =   360
               TabIndex        =   228
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir CTC's ( Por Intervalo )"
               Height          =   195
               Index           =   54
               Left            =   360
               TabIndex        =   227
               Top             =   1440
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir CTC"
               Height          =   195
               Index           =   55
               Left            =   360
               TabIndex        =   226
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir Tudo"
               Height          =   195
               Index           =   56
               Left            =   360
               TabIndex        =   225
               Top             =   1920
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar Fatura Avulsa"
               Height          =   195
               Index           =   57
               Left            =   360
               TabIndex        =   224
               Top             =   2160
               Width           =   1815
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar Fatura"
               Height          =   195
               Index           =   58
               Left            =   360
               TabIndex        =   223
               Top             =   2400
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir CTC"
               Height          =   195
               Index           =   62
               Left            =   360
               TabIndex        =   222
               Top             =   3360
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir CTC/NFS"
               Height          =   195
               Index           =   61
               Left            =   360
               TabIndex        =   221
               Top             =   3120
               Width           =   1695
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
               TabIndex        =   220
               Top             =   2640
               Width           =   3135
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar endereços de cobrança"
               Height          =   195
               Index           =   60
               Left            =   360
               TabIndex        =   219
               Top             =   2880
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar Fatura"
               Height          =   195
               Index           =   64
               Left            =   360
               TabIndex        =   218
               Top             =   3840
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir Tudo"
               Height          =   195
               Index           =   63
               Left            =   360
               TabIndex        =   217
               Top             =   3600
               Width           =   2295
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "49"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   237
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "51"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   236
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "59"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   235
               Top             =   2640
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               Caption         =   "64"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   234
               Top             =   3840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               Caption         =   "50"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   233
               Top             =   480
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosEmissao4 
            Height          =   4215
            Left            =   -74880
            TabIndex        =   210
            Top             =   480
            Width           =   3255
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Manifestos Emitidos"
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
               Index           =   146
               Left            =   480
               TabIndex        =   300
               Top             =   3840
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar Manifesto"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   145
               Left            =   480
               TabIndex        =   262
               Top             =   3600
               Width           =   1575
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir Linha"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   144
               Left            =   480
               TabIndex        =   261
               Top             =   3360
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir Minuta CTC"
               Height          =   195
               Index           =   143
               Left            =   480
               TabIndex        =   260
               Top             =   3120
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emitir Manifesto de Carga"
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
               Index           =   142
               Left            =   480
               TabIndex        =   259
               Top             =   2880
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cancelar CTR/CTC"
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
               Index           =   141
               Left            =   480
               TabIndex        =   258
               Top             =   2640
               Width           =   2055
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Previsão Entrega"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   132
               Left            =   480
               TabIndex        =   257
               Top             =   480
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Calcular"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   131
               Left            =   480
               TabIndex        =   256
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   134
               Left            =   480
               TabIndex        =   255
               Top             =   960
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "OBS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   133
               Left            =   480
               TabIndex        =   254
               Top             =   720
               Width           =   735
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "CTR / CTC Emitidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   135
               Left            =   480
               TabIndex        =   253
               Top             =   1200
               Width           =   2055
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   136
               Left            =   480
               TabIndex        =   252
               Top             =   1440
               Width           =   975
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "OBS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   138
               Left            =   480
               TabIndex        =   251
               Top             =   1920
               Width           =   735
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   137
               Left            =   480
               TabIndex        =   250
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   139
               Left            =   480
               TabIndex        =   249
               Top             =   2160
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Reimpressão de CTR/CTC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   140
               Left            =   480
               TabIndex        =   248
               Top             =   2400
               Width           =   2655
            End
            Begin VB.Label Label159 
               AutoSize        =   -1  'True
               Caption         =   "146"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   301
               Top             =   3840
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label145 
               AutoSize        =   -1  'True
               Caption         =   "131"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   215
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label149 
               AutoSize        =   -1  'True
               Caption         =   "135"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   214
               Top             =   1200
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label154 
               AutoSize        =   -1  'True
               Caption         =   "140"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   213
               Top             =   2400
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label155 
               AutoSize        =   -1  'True
               Caption         =   "141"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   212
               Top             =   2640
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label156 
               AutoSize        =   -1  'True
               Caption         =   "142"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   211
               Top             =   2880
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosEmissao3 
            Height          =   4215
            Left            =   -67200
            TabIndex        =   183
            Top             =   480
            Width           =   3615
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir NFS"
               Height          =   195
               Index           =   130
               Left            =   480
               TabIndex        =   298
               Top             =   3840
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Adicionar NFS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   129
               Left            =   480
               TabIndex        =   207
               Top             =   3600
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   128
               Left            =   480
               TabIndex        =   206
               Top             =   3360
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emissão CTR de COB"
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
               Index           =   127
               Left            =   480
               TabIndex        =   205
               Top             =   3120
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emitir CTC"
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
               Index           =   126
               Left            =   480
               TabIndex        =   204
               Top             =   2880
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emitir CTR"
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
               Index           =   125
               Left            =   480
               TabIndex        =   203
               Top             =   2640
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emissão"
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
               Index           =   124
               Left            =   480
               TabIndex        =   202
               Top             =   2400
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   123
               Left            =   480
               TabIndex        =   201
               Top             =   2160
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar Cadastro"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   122
               Left            =   480
               TabIndex        =   200
               Top             =   1920
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Filiais Intec"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   121
               Left            =   480
               TabIndex        =   196
               Top             =   1680
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Imprimir Tabela"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   120
               Left            =   480
               TabIndex        =   195
               Top             =   1440
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Reajustar Manual"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   118
               Left            =   480
               TabIndex        =   193
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Clonar Tabela"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   119
               Left            =   480
               TabIndex        =   192
               Top             =   1200
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Tabelas de Preços"
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
               Index           =   115
               Left            =   480
               TabIndex        =   186
               Top             =   240
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova Tabela"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   116
               Left            =   480
               TabIndex        =   185
               Top             =   480
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Reajuste Percentual"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   117
               Left            =   480
               TabIndex        =   184
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label144 
               AutoSize        =   -1  'True
               Caption         =   "130"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   299
               Top             =   3840
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label125 
               AutoSize        =   -1  'True
               Caption         =   "127"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   194
               Top             =   3120
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label139 
               AutoSize        =   -1  'True
               Caption         =   "115"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   191
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label133 
               AutoSize        =   -1  'True
               Caption         =   "121"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   190
               Top             =   1680
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label129 
               AutoSize        =   -1  'True
               Caption         =   "124"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   189
               Top             =   2400
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label128 
               AutoSize        =   -1  'True
               Caption         =   "125"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   188
               Top             =   2640
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label126 
               AutoSize        =   -1  'True
               Caption         =   "126"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   187
               Top             =   2880
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosEmissao2 
            Height          =   4215
            Left            =   -71280
            TabIndex        =   160
            Top             =   480
            Width           =   3855
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   113
               Left            =   480
               TabIndex        =   198
               Top             =   3600
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Bloquear"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   114
               Left            =   480
               TabIndex        =   197
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Subcontr. e Represent."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   110
               Left            =   480
               TabIndex        =   182
               Top             =   2880
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir"
               Height          =   195
               Index           =   111
               Left            =   480
               TabIndex        =   181
               Top             =   3120
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   112
               Left            =   480
               TabIndex        =   180
               Top             =   3360
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Rodov - Config. Prod / Cli. dest."
               Height          =   195
               Index           =   106
               Left            =   480
               TabIndex        =   179
               Top             =   1920
               Width           =   2655
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Aéreo - Config. Prod./ Cli. dest."
               Height          =   195
               Index           =   109
               Left            =   480
               TabIndex        =   178
               Top             =   2640
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Rodov - Config. Sub Contratado"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   107
               Left            =   480
               TabIndex        =   177
               Top             =   2160
               Width           =   2655
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Aéreo - Config. Prod. Local."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   108
               Left            =   480
               TabIndex        =   176
               Top             =   2400
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Detalhar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   104
               Left            =   480
               TabIndex        =   174
               Top             =   1200
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Desabilitar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   103
               Left            =   480
               TabIndex        =   173
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Rodov - Config. Prod./Local."
               Height          =   195
               Index           =   105
               Left            =   480
               TabIndex        =   172
               Top             =   1680
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Genérica"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   101
               Left            =   480
               TabIndex        =   163
               Top             =   480
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Aéreo"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   100
               Left            =   480
               TabIndex        =   162
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               Height          =   195
               Index           =   102
               Left            =   480
               TabIndex        =   161
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label140 
               AutoSize        =   -1  'True
               Caption         =   "114"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   199
               Top             =   3840
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label115 
               Caption         =   "Tab Regras de Emissão"
               Height          =   255
               Left            =   480
               TabIndex        =   175
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label123 
               AutoSize        =   -1  'True
               Caption         =   "100"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   165
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label121 
               AutoSize        =   -1  'True
               Caption         =   "110"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   164
               Top             =   2880
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosEmissao1 
            Height          =   4215
            Left            =   -74880
            TabIndex        =   145
            Top             =   480
            Width           =   3375
            Begin VB.CheckBox chkDireitos 
               Caption         =   "IncluirTabela"
               Height          =   195
               Index           =   98
               Left            =   360
               TabIndex        =   169
               Top             =   3720
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Rodoviário"
               Height          =   195
               Index           =   99
               Left            =   360
               TabIndex        =   168
               Top             =   3960
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir"
               Height          =   195
               Index           =   95
               Left            =   360
               TabIndex        =   156
               Top             =   2760
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Bloquear"
               Height          =   195
               Index           =   94
               Left            =   360
               TabIndex        =   155
               Top             =   2280
               Width           =   975
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               Height          =   195
               Index           =   96
               Left            =   360
               TabIndex        =   154
               Top             =   3000
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               Height          =   195
               Index           =   97
               Left            =   360
               TabIndex        =   153
               Top             =   3240
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Histórico"
               Height          =   195
               Index           =   93
               Left            =   360
               TabIndex        =   152
               Top             =   2040
               Width           =   1815
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               Height          =   195
               Index           =   92
               Left            =   360
               TabIndex        =   151
               Top             =   1800
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               Height          =   195
               Index           =   91
               Left            =   360
               TabIndex        =   150
               Top             =   1560
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir"
               Height          =   195
               Index           =   90
               Left            =   360
               TabIndex        =   149
               Top             =   1320
               Width           =   735
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Clientes"
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
               Index           =   89
               Left            =   360
               TabIndex        =   148
               Top             =   840
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Arquivo"
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
               Index           =   87
               Left            =   360
               TabIndex        =   147
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cadastros"
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
               Index           =   88
               Left            =   360
               TabIndex        =   146
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label97 
               AutoSize        =   -1  'True
               Caption         =   "99"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   171
               Top             =   3960
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label96 
               Caption         =   "Tab Tabela de Preços"
               Height          =   255
               Left            =   360
               TabIndex        =   170
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label Label95 
               Caption         =   "Tab Natureza / Produtos"
               Height          =   255
               Left            =   360
               TabIndex        =   167
               Top             =   2520
               Width           =   1815
            End
            Begin VB.Label Label53 
               Caption         =   "Tab Dados do Cliente"
               Height          =   255
               Left            =   360
               TabIndex        =   166
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "88"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   159
               Top             =   600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "89"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   158
               Top             =   840
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label94 
               AutoSize        =   -1  'True
               Caption         =   "87"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   157
               Top             =   240
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Frame FRADireitosFaturamento2 
            Height          =   4095
            Left            =   -71160
            TabIndex        =   139
            Top             =   600
            Width           =   4095
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Imprimir Faturas"
               Height          =   195
               Index           =   80
               Left            =   360
               TabIndex        =   63
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
               TabIndex        =   62
               Top             =   3600
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Efetuar Cancelar Faturas"
               Height          =   195
               Index           =   78
               Left            =   360
               TabIndex        =   61
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
               TabIndex        =   60
               Top             =   3120
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Visualizar Todas Pré-Faturas"
               Height          =   195
               Index           =   66
               Left            =   360
               TabIndex        =   49
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
               TabIndex        =   48
               Top             =   240
               Width           =   3615
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Visualizar Somente as Minhas"
               Height          =   195
               Index           =   67
               Left            =   360
               TabIndex        =   50
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar Vencimento Pré Fatura"
               Height          =   195
               Index           =   68
               Left            =   360
               TabIndex        =   51
               Top             =   960
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Consultar e Alterar Pré Fatura"
               Height          =   195
               Index           =   69
               Left            =   360
               TabIndex        =   52
               Top             =   1200
               Width           =   2415
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Excluir Pré Fatura Pendente"
               Height          =   195
               Index           =   70
               Left            =   360
               TabIndex        =   53
               Top             =   1440
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar Fatura Final"
               Height          =   195
               Index           =   71
               Left            =   360
               TabIndex        =   54
               Top             =   1680
               Width           =   1815
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Prorrogar Vencimentos"
               Height          =   195
               Index           =   75
               Left            =   360
               TabIndex        =   58
               Top             =   2640
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Conceder descontos/Abatimentos"
               Height          =   195
               Index           =   74
               Left            =   360
               TabIndex        =   57
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
               TabIndex        =   55
               Top             =   1920
               Width           =   3615
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar endereços de cobrança"
               Height          =   195
               Index           =   73
               Left            =   360
               TabIndex        =   56
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Quitar Fatura"
               Height          =   195
               Index           =   76
               Left            =   360
               TabIndex        =   59
               Top             =   2880
               Width           =   2295
            End
            Begin VB.Label Label77 
               AutoSize        =   -1  'True
               Caption         =   "72"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   209
               Top             =   1920
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "80"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   143
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
               TabIndex        =   142
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
               TabIndex        =   141
               Top             =   3600
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               Caption         =   "77"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   140
               Top             =   3120
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
            TabIndex        =   117
            Top             =   600
            Width           =   3495
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - Import. EDI - Ocorr/POD"
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   0
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - Exportação para o SITLA"
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   2
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Altera Dados Clientes"
               Height          =   195
               Index           =   6
               Left            =   360
               TabIndex        =   7
               Top             =   1920
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Incl. / Alt. Feriados"
               Height          =   195
               Index           =   7
               Left            =   360
               TabIndex        =   8
               Top             =   2160
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Inclui Novos Clientes"
               Height          =   195
               Index           =   5
               Left            =   360
               TabIndex        =   6
               Top             =   1680
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Usuários"
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   5
               Top             =   1440
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Incl/Alt. Cod. Ocorrência"
               Height          =   195
               Index           =   3
               Left            =   360
               TabIndex        =   4
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Prazos de Entrega"
               Height          =   195
               Index           =   8
               Left            =   360
               TabIndex        =   9
               Top             =   2400
               Width           =   2655
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - LOG de Usuários"
               Height          =   195
               Index           =   26
               Left            =   360
               TabIndex        =   3
               Top             =   960
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Arquivos) - Export. EDI - Ocoren."
               Height          =   195
               Index           =   25
               Left            =   360
               TabIndex        =   1
               Top             =   480
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) Acompanhamento"
               Height          =   195
               Index           =   34
               Left            =   360
               TabIndex        =   10
               Top             =   2880
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   37
               Left            =   360
               TabIndex        =   122
               Top             =   3120
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   40
               Left            =   360
               TabIndex        =   121
               Top             =   3360
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   43
               Left            =   360
               TabIndex        =   120
               Top             =   3600
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   46
               Left            =   360
               TabIndex        =   119
               Top             =   3840
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   47
               Left            =   360
               TabIndex        =   118
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
               TabIndex        =   138
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
               TabIndex        =   137
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
               TabIndex        =   136
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
               TabIndex        =   135
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
               TabIndex        =   134
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
               TabIndex        =   133
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
               TabIndex        =   132
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
               TabIndex        =   131
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
               TabIndex        =   130
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
               TabIndex        =   129
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
               TabIndex        =   128
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
               TabIndex        =   127
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
               TabIndex        =   126
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
               TabIndex        =   125
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
               TabIndex        =   124
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
               TabIndex        =   123
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
            Left            =   3840
            TabIndex        =   97
            Top             =   600
            Width           =   3615
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Ocorrências e POD"
               Height          =   195
               Index           =   11
               Left            =   360
               TabIndex        =   13
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Acomp. de Clientes"
               Height          =   195
               Index           =   15
               Left            =   360
               TabIndex        =   14
               Top             =   960
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Arq. Exclusivo Cliente"
               Height          =   195
               Index           =   17
               Left            =   360
               TabIndex        =   20
               Top             =   2400
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Consulta SAC"
               Height          =   195
               Index           =   10
               Left            =   360
               TabIndex        =   12
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Cadastros) - Novos Prazos de Entr."
               Height          =   195
               Index           =   9
               Left            =   360
               TabIndex        =   11
               Top             =   240
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Acomp. Resumo"
               Height          =   195
               Index           =   27
               Left            =   360
               TabIndex        =   15
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Controle dos Canhotos"
               Height          =   195
               Index           =   28
               Left            =   360
               TabIndex        =   16
               Top             =   1440
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Controle de Devoluções"
               Height          =   195
               Index           =   29
               Left            =   360
               TabIndex        =   17
               Top             =   1680
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Alarme de Urgências"
               Height          =   195
               Index           =   24
               Left            =   360
               TabIndex        =   18
               Top             =   1920
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Informações via Email"
               Height          =   195
               Index           =   14
               Left            =   360
               TabIndex        =   21
               Top             =   2640
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) Ordem de Coleta"
               Height          =   195
               Index           =   32
               Left            =   360
               TabIndex        =   22
               Top             =   2880
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) Cancelamento"
               Height          =   195
               Index           =   35
               Left            =   360
               TabIndex        =   23
               Top             =   3120
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   38
               Left            =   360
               TabIndex        =   100
               Top             =   3360
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   41
               Left            =   360
               TabIndex        =   99
               Top             =   3600
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   44
               Left            =   360
               TabIndex        =   98
               Top             =   3840
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Alarme (GERENCIAL)"
               Height          =   195
               Index           =   30
               Left            =   360
               TabIndex        =   19
               Top             =   2160
               Width           =   3015
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "11"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   116
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
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   110
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
               TabIndex        =   109
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
               TabIndex        =   108
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
               TabIndex        =   107
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
               TabIndex        =   106
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
               TabIndex        =   105
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
               TabIndex        =   104
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
               TabIndex        =   103
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
               TabIndex        =   102
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
               TabIndex        =   101
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
            Left            =   7680
            TabIndex        =   75
            Top             =   600
            Width           =   3495
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "frmCadUsu.frx":0091
               Left            =   1320
               List            =   "frmCadUsu.frx":0093
               TabIndex        =   144
               Text            =   "Perfil de Usuário"
               Top             =   3600
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Excl. de Ocorrências"
               Height          =   195
               Index           =   22
               Left            =   360
               TabIndex        =   24
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Exclusão de PODs"
               Height          =   195
               Index           =   23
               Left            =   360
               TabIndex        =   25
               Top             =   480
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - Análise Estatística"
               Height          =   195
               Index           =   20
               Left            =   360
               TabIndex        =   29
               Top             =   1440
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - Análise de Ocorrências"
               Height          =   195
               Index           =   19
               Left            =   360
               TabIndex        =   32
               Top             =   2160
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - Análise de Entregas"
               Height          =   195
               Index           =   18
               Left            =   360
               TabIndex        =   30
               Top             =   1680
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Recalc. Prev. Entrega"
               Height          =   195
               Index           =   13
               Left            =   360
               TabIndex        =   27
               Top             =   960
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Recalc. Prazos de Entr."
               Height          =   195
               Index           =   12
               Left            =   360
               TabIndex        =   26
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Processos) - Cancelar CTC"
               Height          =   195
               Index           =   16
               Left            =   360
               TabIndex        =   28
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Relatórios) - Protocolo para Arquivo"
               Height          =   195
               Index           =   21
               Left            =   360
               TabIndex        =   33
               Top             =   2400
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Coleta) POD"
               Height          =   195
               Index           =   33
               Left            =   360
               TabIndex        =   34
               Top             =   2640
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   36
               Left            =   360
               TabIndex        =   80
               Top             =   2880
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   39
               Left            =   360
               TabIndex        =   79
               Top             =   3120
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   42
               Left            =   360
               TabIndex        =   78
               Top             =   3360
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   45
               Left            =   360
               TabIndex        =   77
               Top             =   3600
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Height          =   195
               Index           =   48
               Left            =   360
               TabIndex        =   76
               Top             =   3840
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "(Informação) - An.Entregas - Abono"
               Height          =   195
               Index           =   31
               Left            =   360
               TabIndex        =   31
               Top             =   1920
               Width           =   3015
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "22"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   96
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
               TabIndex        =   95
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
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   92
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   88
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
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   83
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
               TabIndex        =   82
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
               TabIndex        =   81
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
            TabIndex        =   208
            Top             =   1860
            Width           =   975
         End
      End
      Begin TabDlg.SSTab SSTabSistemas1 
         Height          =   4815
         Left            =   120
         TabIndex        =   308
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8493
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Aéreo 1/2"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FRADireitosAereo1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FRADireitosAereo2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FRADireitosAereo3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Aéreo 2/2"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FRADireitosAereo4"
         Tab(1).Control(1)=   "FRADireitosAereo5"
         Tab(1).Control(2)=   "FRADireitosAereo6"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Ordem Serviço"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FRADireitosOrdem1"
         Tab(2).ControlCount=   1
         Begin VB.Frame FRADireitosOrdem1 
            Height          =   4335
            Left            =   -74880
            TabIndex        =   438
            Top             =   360
            Width           =   3615
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Ordem de Serviço"
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
               Index           =   255
               Left            =   480
               TabIndex        =   443
               Top             =   1080
               Width           =   2055
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Portaria"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   253
               Left            =   480
               TabIndex        =   441
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emitir Ordem ( Eng.)"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   252
               Left            =   480
               TabIndex        =   440
               Top             =   240
               Width           =   2055
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Manutenção"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   254
               Left            =   480
               TabIndex        =   439
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "255"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   24
               Left            =   120
               TabIndex        =   444
               Top             =   1080
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "252"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   21
               Left            =   120
               TabIndex        =   442
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosAereo6 
            Height          =   4335
            Left            =   -67320
            TabIndex        =   428
            Top             =   360
            Width           =   3615
            Begin VB.ComboBox cb_perfilaereo 
               Height          =   315
               ItemData        =   "frmCadUsu.frx":0095
               Left            =   1440
               List            =   "frmCadUsu.frx":0097
               TabIndex        =   435
               Text            =   "Perfil de Usuário"
               Top             =   3840
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar Arquivo de Movimentação"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   250
               Left            =   480
               TabIndex        =   431
               Top             =   480
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Processar Atualizar Dados"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   249
               Left            =   480
               TabIndex        =   430
               Top             =   240
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Sistema Aéreo"
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
               Index           =   251
               Left            =   480
               TabIndex        =   429
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "251"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   433
               Top             =   840
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "249"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   23
               Left            =   120
               TabIndex        =   432
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosAereo5 
            Height          =   4335
            Left            =   -71040
            TabIndex        =   398
            Top             =   360
            Width           =   3615
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Acompanhamento de Emissão"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   248
               Left            =   480
               TabIndex        =   416
               Top             =   4080
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gerar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   247
               Left            =   480
               TabIndex        =   415
               Top             =   3840
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Confirmar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   244
               Left            =   480
               TabIndex        =   414
               Top             =   3120
               Width           =   1095
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
               Index           =   245
               Left            =   480
               TabIndex        =   413
               Top             =   3360
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   243
               Left            =   480
               TabIndex        =   412
               Top             =   2880
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Etiquetas de Volume"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   242
               Left            =   480
               TabIndex        =   411
               Top             =   2640
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Confirmar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   240
               Left            =   480
               TabIndex        =   410
               Top             =   2160
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Etiquetas de Lote"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   241
               Left            =   480
               TabIndex        =   409
               Top             =   2400
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Manifestos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   238
               Left            =   480
               TabIndex        =   408
               Top             =   1680
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   239
               Left            =   480
               TabIndex        =   407
               Top             =   1920
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cancelar AWB"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   237
               Left            =   480
               TabIndex        =   406
               Top             =   1440
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar AWB para Imprimir "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   235
               Left            =   480
               TabIndex        =   405
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Usar Dados de AWB"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   236
               Left            =   480
               TabIndex        =   404
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar AWB"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   234
               Left            =   480
               TabIndex        =   403
               Top             =   720
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Calcular Tarifas"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   232
               Left            =   480
               TabIndex        =   402
               Top             =   240
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Continuar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   233
               Left            =   480
               TabIndex        =   401
               Top             =   480
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Relatório para Cia. Aérea"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   246
               Left            =   480
               TabIndex        =   400
               Top             =   3600
               Width           =   2535
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "242"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   19
               Left            =   120
               TabIndex        =   434
               Top             =   2640
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "248"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   427
               Top             =   4080
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "246"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   426
               Top             =   3600
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "245"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   425
               Top             =   3360
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "241"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   424
               Top             =   2400
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "238"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   423
               Top             =   1680
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "232"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   417
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosAereo4 
            Height          =   4335
            Left            =   -74880
            TabIndex        =   374
            Top             =   360
            Width           =   3735
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir Notas Fiscais"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   229
               Left            =   480
               TabIndex        =   391
               Top             =   3600
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Consultar AWB"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   216
               Left            =   480
               TabIndex        =   390
               Top             =   480
               Width           =   1695
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Processos"
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
               Index           =   215
               Left            =   480
               TabIndex        =   389
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   217
               Left            =   480
               TabIndex        =   388
               Top             =   720
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Processar / Atualizar Dados"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   219
               Left            =   480
               TabIndex        =   387
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Aconpanhamento de AWB's"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   218
               Left            =   480
               TabIndex        =   386
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Enviar e-mail para Representantes"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   220
               Left            =   480
               TabIndex        =   385
               Top             =   1440
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Geração de Arq. de Movimento"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   222
               Left            =   480
               TabIndex        =   384
               Top             =   1920
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Inserir / Alterar Vôo deste AWB"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   221
               Left            =   480
               TabIndex        =   383
               Top             =   1680
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Inserir / Alterar Vôos de AWB's"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   224
               Left            =   480
               TabIndex        =   382
               Top             =   2400
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir AWB já Informado por e-mail"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   223
               Left            =   480
               TabIndex        =   381
               Top             =   2160
               Width           =   2895
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   225
               Left            =   480
               TabIndex        =   380
               Top             =   2640
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   226
               Left            =   480
               TabIndex        =   379
               Top             =   2880
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Conhecimento Aéreo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   228
               Left            =   480
               TabIndex        =   378
               Top             =   3360
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Emissão"
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
               Index           =   227
               Left            =   480
               TabIndex        =   377
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Incluir Cubagens"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   230
               Left            =   480
               TabIndex        =   376
               Top             =   3840
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Tarifa Spot"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   231
               Left            =   480
               TabIndex        =   375
               Top             =   4080
               Width           =   1215
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "231"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   422
               Top             =   4080
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "228"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   421
               Top             =   3360
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "227"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   420
               Top             =   3120
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "224"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   419
               Top             =   2400
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "218"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   418
               Top             =   960
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "216"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   399
               Top             =   480
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "215"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   392
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosAereo3 
            Height          =   4335
            Left            =   7440
            TabIndex        =   355
            Top             =   360
            Width           =   3735
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   214
               Left            =   480
               TabIndex        =   373
               Top             =   4080
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   213
               Left            =   480
               TabIndex        =   372
               Top             =   3840
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   210
               Left            =   480
               TabIndex        =   370
               Top             =   3120
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Espécie embalagem"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   211
               Left            =   480
               TabIndex        =   369
               Top             =   3360
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   209
               Left            =   480
               TabIndex        =   368
               Top             =   2880
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   208
               Left            =   480
               TabIndex        =   367
               Top             =   2640
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   206
               Left            =   480
               TabIndex        =   366
               Top             =   2160
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Categoria Interna de Produtos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   207
               Left            =   480
               TabIndex        =   365
               Top             =   2400
               Width           =   3015
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   204
               Left            =   480
               TabIndex        =   364
               Top             =   1680
               Width           =   735
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   205
               Left            =   480
               TabIndex        =   363
               Top             =   1920
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Categoria IATA de Produtos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   203
               Left            =   480
               TabIndex        =   362
               Top             =   1440
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   201
               Left            =   480
               TabIndex        =   361
               Top             =   960
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Busca Banco"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   202
               Left            =   480
               TabIndex        =   360
               Top             =   1200
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   200
               Left            =   480
               TabIndex        =   359
               Top             =   720
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Representantes Intec"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   198
               Left            =   480
               TabIndex        =   358
               Top             =   240
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   199
               Left            =   480
               TabIndex        =   357
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Observação Padrão"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   212
               Left            =   480
               TabIndex        =   356
               Top             =   3600
               Width           =   2175
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "214"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   397
               Top             =   4080
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "212"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   396
               Top             =   3600
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "211"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   395
               Top             =   3360
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "207"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   394
               Top             =   2400
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "203"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   393
               Top             =   1440
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label210 
               AutoSize        =   -1  'True
               Caption         =   "198"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   152
               Left            =   120
               TabIndex        =   371
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosAereo2 
            Height          =   4335
            Left            =   3600
            TabIndex        =   333
            Top             =   360
            Width           =   3735
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Clientes Remet/Dest."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   197
               Left            =   480
               TabIndex        =   353
               Top             =   4080
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Remover"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   184
               Left            =   480
               TabIndex        =   347
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Adicionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   183
               Left            =   480
               TabIndex        =   346
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   185
               Left            =   480
               TabIndex        =   345
               Top             =   720
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Zerar Digitação"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   187
               Left            =   480
               TabIndex        =   344
               Top             =   1440
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Iniciar Digitação"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   186
               Left            =   480
               TabIndex        =   343
               Top             =   1200
               Width           =   1455
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cadastrar Nova Tabela"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   188
               Left            =   480
               TabIndex        =   342
               Top             =   1920
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Todas as Localidades"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   190
               Left            =   480
               TabIndex        =   341
               Top             =   2400
               Width           =   2175
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Reajustar Tabelas Cadastradas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   189
               Left            =   480
               TabIndex        =   340
               Top             =   2160
               Width           =   3135
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar Tabela Reajustada"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   192
               Left            =   480
               TabIndex        =   339
               Top             =   2880
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cadastro de Localidades"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   191
               Left            =   480
               TabIndex        =   338
               Top             =   2640
               Width           =   2055
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Localidades / Destinos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   193
               Left            =   480
               TabIndex        =   337
               Top             =   3120
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   194
               Left            =   480
               TabIndex        =   336
               Top             =   3360
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   196
               Left            =   480
               TabIndex        =   335
               Top             =   3840
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   195
               Left            =   480
               TabIndex        =   334
               Top             =   3600
               Width           =   855
            End
            Begin VB.Label Label188 
               AutoSize        =   -1  'True
               Caption         =   "197"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   354
               Top             =   4080
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label202 
               Caption         =   "Fase 2"
               Height          =   255
               Left            =   720
               TabIndex        =   352
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label Label209 
               Caption         =   "Fase 1"
               Height          =   255
               Left            =   720
               TabIndex        =   351
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label208 
               AutoSize        =   -1  'True
               Caption         =   "183"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   350
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label200 
               AutoSize        =   -1  'True
               Caption         =   "189"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   349
               Top             =   2160
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label197 
               AutoSize        =   -1  'True
               Caption         =   "193"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   348
               Top             =   3120
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame FRADireitosAereo1 
            Height          =   4335
            Left            =   120
            TabIndex        =   309
            Top             =   360
            Width           =   3375
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cadastro de Localidades"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   181
               Left            =   480
               TabIndex        =   330
               Top             =   3840
               Width           =   2295
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar Cliente"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   182
               Left            =   480
               TabIndex        =   329
               Top             =   4080
               Width           =   1335
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Todas as Localidades"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   180
               Left            =   480
               TabIndex        =   328
               Top             =   3600
               Width           =   1935
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cadastrar Novas Tabelas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   179
               Left            =   480
               TabIndex        =   324
               Top             =   3120
               Width           =   2535
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cancelar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   177
               Left            =   480
               TabIndex        =   323
               Top             =   2640
               Width           =   975
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Tabela de Preços Cia Aérea"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   178
               Left            =   480
               TabIndex        =   322
               Top             =   2880
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Confirmar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   175
               Left            =   480
               TabIndex        =   321
               Top             =   2160
               Width           =   1095
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Buscar AWB"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   176
               Left            =   480
               TabIndex        =   320
               Top             =   2400
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Inserir Formulário"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   173
               Left            =   480
               TabIndex        =   319
               Top             =   1680
               Width           =   1575
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cancelar Formulário"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   174
               Left            =   480
               TabIndex        =   318
               Top             =   1920
               Width           =   1815
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Gravar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   171
               Left            =   480
               TabIndex        =   317
               Top             =   1200
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Formulários AWB Cia Aérea"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   172
               Left            =   480
               TabIndex        =   316
               Top             =   1440
               Width           =   2775
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Nova"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   169
               Left            =   480
               TabIndex        =   313
               Top             =   720
               Width           =   855
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cadastros"
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
               Index           =   167
               Left            =   480
               TabIndex        =   312
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Cia. Aérea"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   168
               Left            =   480
               TabIndex        =   311
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox chkDireitos 
               Caption         =   "Alterar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   170
               Left            =   480
               TabIndex        =   310
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label193 
               Caption         =   "Fase 0"
               Height          =   255
               Left            =   720
               TabIndex        =   332
               Top             =   3360
               Width           =   615
            End
            Begin VB.Label Label191 
               AutoSize        =   -1  'True
               Caption         =   "182"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   182
               Left            =   120
               TabIndex        =   331
               Top             =   4080
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label189 
               AutoSize        =   -1  'True
               Caption         =   "179"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   179
               Left            =   120
               TabIndex        =   327
               Top             =   3120
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label184 
               AutoSize        =   -1  'True
               Caption         =   "178"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   178
               Left            =   120
               TabIndex        =   326
               Top             =   2880
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label182 
               AutoSize        =   -1  'True
               Caption         =   "172"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   172
               Left            =   120
               TabIndex        =   325
               Top             =   1440
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label174 
               AutoSize        =   -1  'True
               Caption         =   "168"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   168
               Left            =   120
               TabIndex        =   315
               Top             =   480
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label173 
               AutoSize        =   -1  'True
               Caption         =   "167"
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   167
               Left            =   120
               TabIndex        =   314
               Top             =   240
               Visible         =   0   'False
               Width           =   270
            End
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
      TabIndex        =   64
      Top             =   240
      Width           =   6735
      Begin VB.ComboBox cbFiliais 
         Height          =   315
         Left            =   3120
         TabIndex        =   436
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enviar Email Informando Futuros Feriados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   41
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
         TabIndex        =   70
         Top             =   1920
         Width           =   2415
         Begin VB.OptionButton optStatusInativo 
            Caption         =   "Inativo"
            Height          =   255
            Left            =   1320
            TabIndex        =   40
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optStatusAtivo 
            Caption         =   "Ativo"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtDepartamento 
         Height          =   285
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   38
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   37
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
         Height          =   345
         Left            =   720
         MaxLength       =   10
         TabIndex        =   35
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label56 
         Caption         =   "Filial:"
         Height          =   255
         Left            =   2760
         TabIndex        =   437
         Top             =   480
         Width           =   375
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
         TabIndex        =   74
         Top             =   360
         Width           =   75
      End
      Begin VB.Label lblDataCad 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4680
         TabIndex        =   42
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data Cadastro:"
         Height          =   195
         Left            =   3480
         TabIndex        =   68
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome Completo:"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   65
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
Private Sub cb_aereo_Change()

End Sub

Private Sub cb_PerfilFat_Click()
Dim I As Integer
    
    If MsgBox("Deseja Realmente alterar o Perfil de acesso ao sistema Emissão", vbYesNo, "Alteração de Perfil") = vbYes Then

        If cb_PerfilFat.Text = "Coordenador" Then
        
            'Acesso Completo
            For I = 49 To 86
                chkDireitos(I).Value = 1
            Next
    
        ElseIf cb_PerfilFat.Text = "Iniciante" Then
    
            'Acesso
            For I = 49 To 71
                chkDireitos(I).Value = 1
            Next
        
            'Tirar Acesso
            For I = 72 To 76
                chkDireitos(I).Value = 0
            Next
        
            'Acesso
            For I = 77 To 86
                chkDireitos(I).Value = 1
            Next
    
        ElseIf cb_PerfilFat.Text = "Aux. Administrativo" Then
    
            'Acesso Completo
            For I = 49 To 86
                chkDireitos(I).Value = 1
            Next
        End If
    Else
        cb_PerfilFat.Text = "Perfil de Usuário"
        cbPerfil.Text = "Perfil de Usuário"
        Exit Sub
    End If
End Sub

Private Sub cbPerfil_click()
Dim I As Integer

    If MsgBox("Deseja Realmente alterar o Perfil de acesso ao sistema Emissão", vbYesNo, "Alteração de Perfil") = vbYes Then
    
        If cbPerfil.Text = "Coordenador" Then
    
            'Acesso Completo
            For I = 87 To 166
                chkDireitos(I).Value = 1
            Next

        ElseIf cbPerfil.Text = "Emissor" Then
        
            'Acesso
            For I = 87 To 114
                chkDireitos(I).Value = 1
            Next
    
            'Tira acesso
            For I = 115 To 123
                chkDireitos(I).Value = 0
            Next
        
            'Acesso
            For I = 124 To 157
                chkDireitos(I).Value = 1
            Next

            'Acesso ao Menu Relatórios
            I = 158
            chkDireitos(I).Value = 1
        
            'Tiro acesso
            For I = 159 To 163
                chkDireitos(I).Value = 0
            Next
   
            'Acesso
            For I = 164 To 165
                chkDireitos(I).Value = 1
            Next
    
            'Acesso ao Módulo de Emissão
            I = 166
            chkDireitos(I).Value = 1
    
        ElseIf cbPerfil.Text = "Iniciante" Then
        
            'Acesso ao Menu Relatórios
            I = 87
            chkDireitos(I).Value = 1
        
            'Tiro acesso Menu Cadastro
            For I = 88 To 123
                chkDireitos(I).Value = 0
            Next

            'Acesso
            For I = 124 To 140
                chkDireitos(I).Value = 1
            Next

            'Tira acesso Cancelamentos
            I = 141
            chkDireitos(I).Value = 0
        
            'Acesso
            For I = 142 To 147
                chkDireitos(I).Value = 1
            Next

            'Tira acesso Cancelamentos
            I = 148
            chkDireitos(I).Value = 0
        
            'Acesso
            For I = 149 To 155
                chkDireitos(I).Value = 1
            Next

            'Tira acesso Cancelamentos
            I = 156
            chkDireitos(I).Value = 0
                
            'Tira acesso Cancelamentos
            I = 157
            chkDireitos(I).Value = 0
        
            'Tiro acesso Menu Relatórios
            For I = 158 To 165
                chkDireitos(I).Value = 0
            Next
        
            'Acesso ao Módulo de Emissão
            I = 166
            chkDireitos(I).Value = 1

        End If
    Else
        cb_PerfilFat.Text = "Perfil de Usuário"
        cbPerfil.Text = "Perfil de Usuário"
        Exit Sub
    End If

End Sub
Private Sub cmd_limpar_Click()
    Dim X As Integer
    
    If MsgBox("Deseja Realmente limpar todos os direitos do Usuário?", vbInformation + vbYesNo, "Tirar Todos os acessos") = vbYes Then
        
        For X = 1 To 255
            chkDireitos(X).Value = 0
        Next
    Else
        Exit Sub
    End If
        
    
End Sub
Private Sub cmdAltUsu_Click()
    Dim X As Integer
        
    FRADireitosEmissao1.Enabled = True
    FRADireitosEmissao2.Enabled = True
    FRADireitosEmissao3.Enabled = True
    FRADireitosEmissao4.Enabled = True
    FRADireitosEmissao5.Enabled = True
    FRADireitosEmissao6.Enabled = True
    
    
    FRADireitosFaturamento1.Enabled = True
    FRADireitosFaturamento2.Enabled = True
    FRADireitosFaturamento3.Enabled = True
    
    FRADireitosInforma1.Enabled = True
    FRADireitosInforma2.Enabled = True
    FRADireitosInforma3.Enabled = True
    
    FRADireitosAereo1.Enabled = True
    FRADireitosAereo2.Enabled = True
    FRADireitosAereo3.Enabled = True
    FRADireitosAereo4.Enabled = True
    FRADireitosAereo5.Enabled = True
    FRADireitosAereo6.Enabled = True
    
    FRADireitosOrdem1.Enabled = True
    
    cmd_limpar.Enabled = True
    
    
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
    cbFiliais.BackColor = &HC0FFFF     'AMARELO
    txtNome.BackColor = &HC0FFFF     'AMARELO
    txtDepartamento.BackColor = &HC0FFFF     'AMARELO
    txtNome.SetFocus
    lblOpcao = "ALTERAÇÃO"
End Sub

Private Sub cmdCancelar_Click()
    TxtUsuario.Enabled = True
    fraUsuarios.Enabled = True
    fraDireitos.Enabled = True
    
    FRADireitosEmissao1.Enabled = False
    FRADireitosEmissao2.Enabled = False
    FRADireitosEmissao3.Enabled = False
    FRADireitosEmissao4.Enabled = False
    FRADireitosEmissao5.Enabled = False
    FRADireitosEmissao6.Enabled = False
    
    FRADireitosFaturamento1.Enabled = False
    FRADireitosFaturamento2.Enabled = False
    FRADireitosFaturamento3.Enabled = False
    
    FRADireitosInforma1.Enabled = False
    FRADireitosInforma2.Enabled = False
    FRADireitosInforma3.Enabled = False
    
    FRADireitosAereo1.Enabled = False
    FRADireitosAereo2.Enabled = False
    FRADireitosAereo3.Enabled = False
    FRADireitosAereo4.Enabled = False
    FRADireitosAereo5.Enabled = False
    FRADireitosAereo6.Enabled = False
    
    FRADireitosOrdem1.Enabled = False
    
    cmd_limpar.Enabled = False
    
    FraDadosUsu.Enabled = False
    cmdNovoUsu.Enabled = True
    cmdAltUsu.Enabled = False
    cmdGravar.Enabled = False
    CmdCancelar.Enabled = False
    cmdSair.Enabled = True

    cbFiliais.BackColor = &H8000000E     'BRANCO
    TxtUsuario.BackColor = &H8000000E       'BRANCO
    txtNome.BackColor = &H8000000E       'BRANCO
    txtDepartamento.BackColor = &H8000000E      'BRANCO
    cmdNovoUsu.SetFocus
    lblOpcao = ""
    
    GridUsuario_Click
    
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
    For X = 1 To 255
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
            de_informa.Ins_cadUsu TxtUsuario, TxtUsuario, txtNome, Mid(cbFiliais, 1, 2), txtDepartamento, lblDataCad, xStatus, xstrdireitos, "S"
        End If
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "CAD. DE USUÁRIOS: " & TxtUsuario
        
    Else
    
        de_informa.alt_cadusu TxtUsuario, txtNome, Mid(cbFiliais, 1, 2), txtDepartamento, xStatus, xstrdireitos
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "CAD. DE USUÁRIOS: " & TxtUsuario
        
    End If
    
    If de_informa.rsSel_alluser.State = 1 Then de_informa.rsSel_alluser.Close
    de_informa.Sel_alluser
    GridUsuario.DataMember = "Sel_alluser"
    
    GridUsuario.Refresh
    lblOpcao = ""
    cmdCancelar_Click
    
    'Franklin
    cb_PerfilFat.Text = "Perfil de Usuário"
    cbPerfil.Text = "Perfil de Usuário"
    cb_perfilaereo = "Perfil de Usuário"

End Sub
Private Sub cmdNovoUsu_Click()
    Dim X As Integer
    
    FRADireitosEmissao1.Enabled = True
    FRADireitosEmissao2.Enabled = True
    FRADireitosEmissao3.Enabled = True
    FRADireitosEmissao4.Enabled = True
    FRADireitosEmissao5.Enabled = True
    FRADireitosEmissao6.Enabled = True
    
    FRADireitosFaturamento1.Enabled = True
    FRADireitosFaturamento2.Enabled = True
    FRADireitosFaturamento3.Enabled = True
    
    FRADireitosInforma1.Enabled = True
    FRADireitosInforma2.Enabled = True
    FRADireitosInforma3.Enabled = True
    
    FRADireitosAereo1.Enabled = True
    FRADireitosAereo2.Enabled = True
    FRADireitosAereo3.Enabled = True
    FRADireitosAereo4.Enabled = True
    FRADireitosAereo5.Enabled = True
    FRADireitosAereo6.Enabled = True
    
    FRADireitosOrdem1.Enabled = True
    
    cmd_limpar.Enabled = True
    
    fraUsuarios.Enabled = False
    fraDireitos.Enabled = True
    FraDadosUsu.Enabled = True
    cmdNovoUsu.Enabled = False
    cmdAltUsu.Enabled = False
    cmdGravar.Enabled = True
    CmdCancelar.Enabled = True
    lblDataCad = datahora("data")
    cmdSair.Enabled = False
    TxtUsuario.BackColor = &HC0FFFF         'AMARELO
    txtNome.BackColor = &HC0FFFF            'AMARELO
    txtDepartamento.BackColor = &HC0FFFF    'AMARELO
    cbFiliais.BackColor = &HC0FFFF          'AMARELO
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
    
    GridUsuario.Refresh
    
    cbFiliais.Text = "00"
    cb_PerfilFat.Text = "Perfil de Usuário"
    cbPerfil.Text = "Perfil de Usuário"
    cb_perfilaereo = "Perfil de Usuário"

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
SSTabSistemas.Visible = False
SSTabSistemas1.Visible = True
End Sub

Private Sub Command2_Click()
SSTabSistemas.Visible = True
SSTabSistemas1.Visible = False
End Sub

Private Sub Form_Activate()
    
    SSTabSistemas1.Visible = False
      
    
    cmdNovoUsu.SetFocus
    fraDireitos.Enabled = True
    
    FRADireitosEmissao1.Enabled = False
    FRADireitosEmissao2.Enabled = False
    FRADireitosEmissao3.Enabled = False
    FRADireitosEmissao4.Enabled = False
    FRADireitosEmissao5.Enabled = False
    FRADireitosEmissao6.Enabled = False
        
    FRADireitosFaturamento1.Enabled = False
    FRADireitosFaturamento2.Enabled = False
    FRADireitosFaturamento3.Enabled = False
    
    FRADireitosInforma1.Enabled = False
    FRADireitosInforma2.Enabled = False
    FRADireitosInforma3.Enabled = False
    
    FRADireitosAereo1.Enabled = False
    FRADireitosAereo2.Enabled = False
    FRADireitosAereo3.Enabled = False
    FRADireitosAereo4.Enabled = False
    FRADireitosAereo5.Enabled = False
    FRADireitosAereo6.Enabled = False
    
    FRADireitosOrdem1.Enabled = False
    
    cmd_limpar.Enabled = False
    
    If de_informa.rsSel_UsuarioFiliais.State = 1 Then de_informa.rsSel_UsuarioFiliais.Close
    
    de_informa.sel_Usuariofiliais
    If de_informa.rsSel_UsuarioFiliais.RecordCount <= 0 Then
        MsgBox ("Nenhuma Filial Cadastrada... Verificar!"), vbInformation, "Busca Filial"
    Else
        Do Until de_informa.rsSel_UsuarioFiliais.EOF
            cbFiliais.AddItem de_informa.rsSel_UsuarioFiliais.Fields("filial") & " -  " & de_informa.rsSel_UsuarioFiliais.Fields("nomefilial")
            de_informa.rsSel_UsuarioFiliais.MoveNext
        Loop
    End If
    
End Sub
Private Sub Form_Load()
    If de_informa.rsSel_alluser.State = 1 Then de_informa.rsSel_alluser.Close
    de_informa.Sel_alluser
    GridUsuario.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCadUsu = Nothing
End Sub

Private Sub GridUsuario_Click()
    Dim X As Integer
    
    TxtUsuario = GridUsuario.Columns(0)
    txtNome = GridUsuario.Columns(2)
    cbFiliais = GridUsuario.Columns(3)
    txtDepartamento = GridUsuario.Columns(4)
    lblDataCad = GridUsuario.Columns(5)
    
    SSTabSistemas.Visible = True
    
    If GridUsuario.Columns(5) = "0" Then
        optStatusInativo = True
    Else
        optStatusAtivo = True
    End If
    
    For X = 1 To 255
        If Mid$(GridUsuario.Columns(7), X, 1) = "1" Then
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


Private Sub Label122_Click()
End Sub

Private Sub Label110_Click()
End Sub

Private Sub Label201_Click()

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
