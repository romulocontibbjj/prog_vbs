VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadClientes 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   8055
   ClientLeft      =   1395
   ClientTop       =   1905
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12015
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   104
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Dados do Cliente"
      TabPicture(0)   =   "frmCadClientes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblcgc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraContatos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdIncluir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAlterar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdGravar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSair"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraDiversos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdHistorico"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdBloquear"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraConsig"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Natureza / Produtos"
      TabPicture(1)   =   "frmCadClientes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraGridProd"
      Tab(1).Control(1)=   "fraProd"
      Tab(1).Control(2)=   "cmdSairProd"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tabela de Preço"
      TabPicture(2)   =   "frmCadClientes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdtabSair"
      Tab(2).Control(1)=   "fraTabDetalhe"
      Tab(2).Control(2)=   "fraTabsPreco"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "Frame4"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Regras de Emissão"
      TabPicture(3)   =   "frmCadClientes.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame12"
      Tab(3).Control(2)=   "Frame14"
      Tab(3).Control(3)=   "Command16"
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdtabSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   -64680
         TabIndex        =   166
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Frame fraTabDetalhe 
         Caption         =   "Detalhe da Tabela"
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
         Left            =   -69120
         TabIndex        =   176
         Top             =   4200
         Width           =   5775
         Begin VB.CommandButton cmdTabDesabilitar 
            Caption         =   "Desabilitar ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   164
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton cmdTabDetalhar 
            Caption         =   "Detalhar ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   4440
            TabIndex        =   165
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton cmdTabGravar 
            Caption         =   "Gravar ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   4440
            TabIndex        =   163
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblTabDataIncl 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   186
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblTabUsuario 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   185
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   120
            TabIndex        =   184
            Top             =   2280
            Width           =   585
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Data Inclusão:"
            Height          =   195
            Left            =   120
            TabIndex        =   183
            Top             =   1920
            Width           =   1035
         End
         Begin VB.Label lblTabModalTab 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   182
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblTabDescrTab 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   765
            Left            =   1200
            TabIndex        =   181
            Top             =   720
            Width           =   4455
         End
         Begin VB.Label lblTabTabela 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   180
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   120
            TabIndex        =   179
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Descricao:"
            Height          =   195
            Left            =   120
            TabIndex        =   178
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Tabela:"
            Height          =   195
            Left            =   120
            TabIndex        =   177
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame fraTabsPreco 
         Caption         =   "Tabelas Cadastradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -74880
         TabIndex        =   175
         Top             =   4200
         Width           =   5655
         Begin MSDataGridLib.DataGrid gridTabTabelas 
            Bindings        =   "frmCadClientes.frx":0070
            Height          =   3015
            Left            =   120
            TabIndex        =   167
            Top             =   360
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   5318
            _Version        =   393216
            Enabled         =   0   'False
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
            DataMember      =   "Sel_CadCliProdsTAB"
            ColumnCount     =   10
            BeginProperty Column00 
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
            BeginProperty Column01 
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
            BeginProperty Column02 
               DataField       =   "remetente"
               Caption         =   "Remetente"
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
               DataField       =   "nomerem"
               Caption         =   "nomerem"
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
               DataField       =   "natproduto"
               Caption         =   "Natureza Produto"
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
               DataField       =   "tabelapreco"
               Caption         =   "Tabela"
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
               DataField       =   "descricaotab"
               Caption         =   "descricaotab"
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
               DataField       =   "modal"
               Caption         =   "Modal"
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
               DataField       =   "datacad"
               Caption         =   "datacad"
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
            BeginProperty Column09 
               DataField       =   "usuariocad"
               Caption         =   "usuariocad"
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
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   464,882
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1395,213
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2234,835
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   975,118
               EndProperty
               BeginProperty Column06 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   585,071
               EndProperty
               BeginProperty Column08 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column09 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1140,095
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Quando este Cliente for o ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -74880
         TabIndex        =   169
         Top             =   480
         Width           =   11535
         Begin VB.CommandButton cmdTabIncluirTab 
            Caption         =   "Incluir Tabela"
            Height          =   375
            Left            =   9120
            TabIndex        =   153
            Top             =   480
            Width           =   1935
         End
         Begin VB.Frame fraTabRemet 
            Caption         =   "... e o Cliente Remetente/Origem for o ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   240
            TabIndex        =   172
            Top             =   1200
            Width           =   8535
            Begin VB.TextBox txtTabRemet 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               MaxLength       =   8
               TabIndex        =   155
               Top             =   600
               Width           =   1695
            End
            Begin VB.Frame Frame8 
               Caption         =   "... e  o Produto / Natureza do Cliente Remetente/Origem for o ..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   120
               TabIndex        =   173
               Top             =   1080
               Width           =   8295
               Begin VB.CommandButton txtTabBuscaProd 
                  Caption         =   "?"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   3360
                  TabIndex        =   174
                  Top             =   600
                  Width           =   255
               End
               Begin VB.TextBox txtTabProd 
                  BackColor       =   &H8000000E&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   158
                  Top             =   600
                  Width           =   3135
               End
               Begin VB.CheckBox chkTabTodosProd 
                  Caption         =   "Todos"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  TabIndex        =   157
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   855
               End
               Begin VB.Label lblTabDescrProd 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   525
                  Left            =   3720
                  TabIndex        =   159
                  Top             =   360
                  Width           =   4455
               End
            End
            Begin VB.CheckBox chkTabTodosRemet 
               Caption         =   "Todos"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   154
               Top             =   360
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CommandButton txtTabBuscaRemet 
               Caption         =   "?"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2040
               TabIndex        =   168
               Top             =   600
               Width           =   255
            End
            Begin VB.Label lblTabDescrRemet 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2400
               TabIndex        =   156
               Top             =   600
               Width           =   4935
            End
         End
         Begin VB.Frame fraTabConsig 
            Caption         =   "... Consignatário ..."
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
            Left            =   240
            TabIndex        =   171
            Top             =   360
            Width           =   8535
            Begin VB.Label lblTabConsig 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   240
               TabIndex        =   151
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label lblTabDescrConsig 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2400
               TabIndex        =   152
               Top             =   360
               Width           =   4935
            End
         End
         Begin VB.Frame fraTabIncl 
            Caption         =   "Incluir Tabela de Preço"
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
            Height          =   2295
            Left            =   8880
            TabIndex        =   170
            Top             =   1200
            Width           =   2415
            Begin VB.CommandButton cmdTabGenIncl 
               Caption         =   "Genérica ..."
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               TabIndex        =   162
               Top             =   1680
               Width           =   1935
            End
            Begin VB.CommandButton cmdTabAirIncl 
               Caption         =   "Aérea ..."
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               TabIndex        =   161
               Top             =   1080
               Width           =   1935
            End
            Begin VB.CommandButton cmdTabRodoIncl 
               Caption         =   "Rodoviária ..."
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               TabIndex        =   160
               Top             =   480
               Width           =   1935
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tabelas de Preço de Frete"
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
         Left            =   -67320
         TabIndex        =   147
         Top             =   2760
         Width           =   3855
         Begin VB.CommandButton cmdVerTabPrecoProdAir 
            Caption         =   "Detalhar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1440
            TabIndex        =   150
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdVerTabPrecoProdGen 
            Caption         =   "Detalhar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2520
            TabIndex        =   149
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdVerTabPrecoProdRodo 
            Caption         =   "Detalhar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   360
            TabIndex        =   148
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Regra de Emissão: Modal Aéreo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   142
         Top             =   2880
         Width           =   11535
         Begin VB.CommandButton Command5 
            Caption         =   "Configurar: Produto / Localidade ..."
            Height          =   375
            Left            =   360
            TabIndex        =   146
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Configurar: Produto / Cliente Destino ..."
            Height          =   375
            Left            =   360
            TabIndex        =   145
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Frame Frame11 
            Caption         =   "Regras Cadastradas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   3960
            TabIndex        =   143
            Top             =   120
            Width           =   7455
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   1695
               Left            =   120
               TabIndex        =   144
               Top             =   240
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   2990
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
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
                  DataField       =   ""
                  Caption         =   ""
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
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Regra de Emissão: Modal Rodoviário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   136
         Top             =   480
         Width           =   11535
         Begin VB.Frame Frame13 
            Caption         =   "Regras Cadastradas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   3960
            TabIndex        =   140
            Top             =   120
            Width           =   7455
            Begin MSDataGridLib.DataGrid DataGrid3 
               Height          =   1695
               Left            =   120
               TabIndex        =   141
               Top             =   240
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   2990
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
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
                  DataField       =   ""
                  Caption         =   ""
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
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Configurar: Produto / Cliente Destino ..."
            Height          =   375
            Left            =   360
            TabIndex        =   139
            Top             =   840
            Width           =   3135
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Configurar: Produto / Localidade ..."
            Height          =   375
            Left            =   360
            TabIndex        =   138
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Configurar: SubContratado ..."
            Height          =   375
            Left            =   360
            TabIndex        =   137
            Top             =   1320
            Width           =   3135
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Instruções / Observações Padrão Para Este Cliente"
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
         Left            =   -74880
         TabIndex        =   129
         Top             =   5280
         Width           =   11535
         Begin VB.OptionButton Option3 
            Caption         =   "Como Remetente:"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Como Destinatário:"
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Text38 
            BackColor       =   &H8000000E&
            Height          =   495
            Left            =   1920
            TabIndex        =   133
            Top             =   480
            Width           =   8055
         End
         Begin VB.TextBox Text39 
            BackColor       =   &H8000000E&
            Height          =   495
            Left            =   1920
            TabIndex        =   132
            Top             =   1200
            Width           =   8055
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Gravar"
            Height          =   495
            Left            =   10320
            TabIndex        =   131
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Gravar"
            Height          =   495
            Left            =   10320
            TabIndex        =   130
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Sair"
         Height          =   375
         Left            =   -65040
         TabIndex        =   128
         Top             =   7320
         Width           =   1575
      End
      Begin VB.Frame fraConsig 
         Caption         =   "Consignatário Padrão"
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
         Height          =   1000
         Left            =   120
         TabIndex        =   109
         Top             =   5920
         Width           =   11535
         Begin VB.CommandButton cmdBuscaConsigEntrAir 
            Caption         =   "?"
            Height          =   285
            Left            =   2760
            TabIndex        =   188
            Top             =   630
            Width           =   255
         End
         Begin VB.TextBox txtConsigEntregaAir 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   14
            TabIndex        =   36
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkConsigEntrPropAir 
            Caption         =   "O Próprio"
            Height          =   195
            Left            =   3060
            TabIndex        =   187
            Top             =   630
            Width           =   975
         End
         Begin VB.CheckBox chkConsigDevolProp 
            Caption         =   "O Próprio"
            Height          =   195
            Left            =   8730
            TabIndex        =   122
            Top             =   320
            Width           =   975
         End
         Begin VB.CheckBox chkConsigTransfProp 
            Caption         =   "O Próprio"
            Height          =   195
            Left            =   8730
            TabIndex        =   121
            Top             =   630
            Width           =   975
         End
         Begin VB.CheckBox chkConsigEntrProp 
            Caption         =   "O Próprio"
            Height          =   195
            Left            =   3060
            TabIndex        =   120
            Top             =   320
            Width           =   975
         End
         Begin VB.CommandButton cmdBuscaConsigTransf 
            Caption         =   "?"
            Height          =   285
            Left            =   8400
            TabIndex        =   118
            Top             =   630
            Width           =   255
         End
         Begin VB.TextBox txtConsigEntrega 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   14
            TabIndex        =   35
            Top             =   320
            Width           =   1455
         End
         Begin VB.CommandButton cmdBuscaConsigDevol 
            Caption         =   "?"
            Height          =   285
            Left            =   8400
            TabIndex        =   116
            Top             =   320
            Width           =   255
         End
         Begin VB.TextBox txtConsigDevol 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   6960
            MaxLength       =   14
            TabIndex        =   37
            Top             =   320
            Width           =   1455
         End
         Begin VB.CommandButton cmdBuscaConsigEntr 
            Caption         =   "?"
            Height          =   285
            Left            =   2760
            TabIndex        =   114
            Top             =   320
            Width           =   255
         End
         Begin VB.TextBox txtConsigTransf 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   6960
            MaxLength       =   14
            TabIndex        =   38
            Top             =   630
            Width           =   1455
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Entrega Aéreo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   190
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label lblConsigEntregaNomeAir 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4080
            TabIndex        =   189
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label lblConsigEntregaNome 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4080
            TabIndex        =   119
            Top             =   320
            Width           =   1695
         End
         Begin VB.Label lblConsigDevolNome 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9720
            TabIndex        =   117
            Top             =   320
            Width           =   1695
         End
         Begin VB.Label lblConsigTransfNome 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9720
            TabIndex        =   115
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Devolução:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   112
            Top             =   320
            Width           =   825
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Transferência:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   111
            Top             =   630
            Width           =   1020
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Entrega Rodo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   320
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   1455
         Left            =   10440
         TabIndex        =   105
         Top             =   4440
         Width           =   1215
         Begin VB.CommandButton cmdDetalhaBloq 
            Caption         =   "Detalhe Bloqueio"
            Enabled         =   0   'False
            Height          =   735
            Left            =   120
            TabIndex        =   107
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblDataBloqueio 
            AutoSize        =   -1  'True
            Caption         =   "3"
            Height          =   195
            Left            =   600
            TabIndex        =   127
            Top             =   360
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblUsuBloqueio 
            AutoSize        =   -1  'True
            Caption         =   "2"
            Height          =   195
            Left            =   360
            TabIndex        =   126
            Top             =   360
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblDescrBloqueio 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   195
            Left            =   120
            TabIndex        =   125
            Top             =   360
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   540
            TabIndex        =   106
            Top             =   240
            Width           =   105
         End
      End
      Begin VB.CommandButton cmdSairProd 
         Caption         =   "Sair"
         Height          =   375
         Left            =   -65040
         TabIndex        =   85
         Top             =   7240
         Width           =   1575
      End
      Begin VB.CommandButton cmdBloquear 
         Caption         =   "Bloquear ..."
         Enabled         =   0   'False
         Height          =   495
         Left            =   10440
         TabIndex        =   84
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "Histórico ..."
         Enabled         =   0   'False
         Height          =   495
         Left            =   10440
         TabIndex        =   83
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Frame fraDiversos 
         Caption         =   "Diversos"
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
         Height          =   680
         Left            =   120
         TabIndex        =   79
         Top             =   6960
         Width           =   11535
         Begin VB.TextBox txtPrazo 
            Height          =   285
            Left            =   2760
            MaxLength       =   6
            TabIndex        =   40
            Top             =   285
            Width           =   855
         End
         Begin VB.TextBox txtAtend1 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   720
            MaxLength       =   10
            TabIndex        =   39
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Prazo:"
            Height          =   195
            Left            =   2160
            TabIndex        =   124
            Top             =   315
            Width           =   450
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Atend:"
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   315
            Width           =   465
         End
         Begin VB.Label lblUsuCad 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7320
            TabIndex        =   103
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label lblUltEmissao 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9600
            TabIndex        =   102
            Top             =   285
            Width           =   1815
         End
         Begin VB.Label lblDataCad 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4560
            TabIndex        =   101
            Top             =   285
            Width           =   1815
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro:"
            Height          =   195
            Left            =   3840
            TabIndex        =   82
            Top             =   315
            Width           =   675
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   6600
            TabIndex        =   81
            Top             =   315
            Width           =   585
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Ult. Emissão:"
            Height          =   195
            Left            =   8640
            TabIndex        =   80
            Top             =   315
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   495
         Left            =   10440
         TabIndex        =   42
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   10440
         TabIndex        =   41
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   10440
         TabIndex        =   78
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   495
         Left            =   10440
         TabIndex        =   77
         Top             =   600
         Width           =   1215
      End
      Begin VB.Frame fraProd 
         Caption         =   "Incluir Natureza de Produto do Cliente Origem"
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
         Left            =   -74880
         TabIndex        =   65
         Top             =   420
         Width           =   11535
         Begin VB.CommandButton cmdGravarProd 
            Caption         =   "Gravar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   10200
            TabIndex        =   69
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdAlterarProd 
            Caption         =   "Alterar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   10200
            TabIndex        =   70
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdIncluirProd 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   10200
            TabIndex        =   71
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtObsProd 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   68
            Top             =   1800
            Width           =   6855
         End
         Begin VB.CommandButton cmdBuscaIata 
            Caption         =   "?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1920
            TabIndex        =   74
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtCodIata 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   67
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtProdNatureza 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   66
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblAcaoProd 
            AutoSize        =   -1  'True
            Caption         =   "Consulta"
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
            Left            =   7800
            TabIndex        =   100
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblDescrIata 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            TabIndex        =   86
            Top             =   840
            Width           =   6495
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Observação de Emissão Padrão Para Este Produto:"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1560
            Width           =   3660
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Class. IATA:"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.Frame fraGridProd 
         Caption         =   "Natureza de Produtos do Cliente Origem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -74880
         TabIndex        =   63
         Top             =   3180
         Width           =   11535
         Begin MSDataGridLib.DataGrid gridProd 
            Bindings        =   "frmCadClientes.frx":0089
            Height          =   3405
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   6006
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
            DataMember      =   "Sel_CadCliProds"
            ColumnCount     =   7
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
               DataField       =   "natproduto"
               Caption         =   "Natureza Produto"
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
               DataField       =   "classiata"
               Caption         =   "Class.IATA"
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
               DataField       =   "obspadrao"
               Caption         =   "Observação de Emissão Padrão"
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
            BeginProperty Column05 
               DataField       =   "datacad"
               Caption         =   "Data Cadastro"
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
               DataField       =   "usucad"
               Caption         =   "Usuário Cad."
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
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3449,764
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   3644,788
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   464,882
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1200,189
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Cadastrais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   10215
         Begin VB.Frame fraDados 
            Enabled         =   0   'False
            Height          =   2040
            Left            =   120
            TabIndex        =   89
            Top             =   840
            Width           =   9975
            Begin VB.CommandButton cmdBuscaClasse 
               Caption         =   "Busca..."
               Height          =   255
               Left            =   6960
               TabIndex        =   194
               Top             =   1680
               Width           =   1095
            End
            Begin VB.OptionButton optPessoaJuridica 
               Caption         =   "Jurídica"
               Height          =   255
               Left            =   6480
               TabIndex        =   193
               Top             =   600
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optPessoaFisica 
               Caption         =   "Física"
               Height          =   255
               Left            =   7440
               TabIndex        =   191
               Top             =   600
               Width           =   855
            End
            Begin VB.CheckBox chkAlarme 
               Caption         =   "Tela de Alarme"
               Height          =   255
               Left            =   8520
               TabIndex        =   7
               Top             =   1680
               Width           =   1420
            End
            Begin VB.TextBox txtCep 
               BackColor       =   &H8000000E&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1080
               MaxLength       =   8
               TabIndex        =   11
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtRazao 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   1080
               MaxLength       =   40
               TabIndex        =   6
               Top             =   240
               Width           =   4455
            End
            Begin VB.TextBox txtFantasia 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   8
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtApelido 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   4440
               MaxLength       =   8
               TabIndex        =   9
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtEndereco 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   1080
               MaxLength       =   40
               TabIndex        =   10
               Top             =   960
               Width           =   4455
            End
            Begin VB.TextBox txtCidade 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   4440
               MaxLength       =   35
               TabIndex        =   12
               Top             =   1320
               Width           =   3615
            End
            Begin VB.TextBox txtIe 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   14
               Top             =   1680
               Width           =   1815
            End
            Begin VB.TextBox txtPabx 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   6120
               MaxLength       =   20
               TabIndex        =   15
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox txtFax 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   8520
               MaxLength       =   15
               TabIndex        =   16
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblClasseFiscal 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   4440
               TabIndex        =   196
               Top             =   1680
               Width           =   2415
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Classificação Fiscal:"
               Height          =   195
               Left            =   3000
               TabIndex        =   195
               Top             =   1680
               Width           =   1425
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Pessoa:"
               Height          =   195
               Left            =   5640
               TabIndex        =   192
               Top             =   600
               Width           =   570
            End
            Begin VB.Label lblAcao 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Consulta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   8760
               TabIndex        =   113
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblUf 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   9360
               TabIndex        =   13
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fantasia:"
               Height          =   195
               Left            =   120
               TabIndex        =   99
               Top             =   600
               Width           =   645
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Razão Soc:"
               Height          =   195
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Apelido:"
               Height          =   195
               Left            =   3840
               TabIndex        =   97
               Top             =   600
               Width           =   570
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "CEP:"
               Height          =   195
               Left            =   120
               TabIndex        =   96
               Top             =   1320
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   195
               Left            =   120
               TabIndex        =   95
               Top             =   960
               Width           =   735
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Inscr. Est.:"
               Height          =   195
               Left            =   120
               TabIndex        =   94
               Top             =   1680
               Width           =   750
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Pabx:"
               Height          =   195
               Left            =   5640
               TabIndex        =   93
               Top             =   960
               Width           =   405
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Height          =   195
               Left            =   3000
               TabIndex        =   92
               Top             =   1320
               Width           =   540
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Fax:"
               Height          =   195
               Left            =   8160
               TabIndex        =   91
               Top             =   960
               Width           =   300
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "UF:"
               Height          =   195
               Left            =   8880
               TabIndex        =   90
               Top             =   1320
               Width           =   255
            End
         End
         Begin VB.Frame fraCgc 
            Height          =   615
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   5175
            Begin VB.CommandButton cmdBusca 
               Caption         =   "Busca ..."
               Height          =   255
               Left            =   3720
               TabIndex        =   2
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton cmdGo 
               Caption         =   ">>"
               Enabled         =   0   'False
               Height          =   255
               Left            =   3120
               TabIndex        =   1
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtCgc 
               BackColor       =   &H0080FFFF&
               Height          =   285
               Left            =   1080
               MaxLength       =   14
               TabIndex        =   0
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CNPJ/CPF:"
               Height          =   195
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   825
            End
         End
         Begin VB.Frame fraCliente 
            Caption         =   "Cliente..."
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
            Height          =   615
            Left            =   5400
            TabIndex        =   75
            Top             =   240
            Width           =   4695
            Begin VB.OptionButton optLogistico 
               Caption         =   "Op. Logístico"
               Height          =   195
               Left            =   3240
               TabIndex        =   5
               Top             =   260
               Width           =   1335
            End
            Begin VB.OptionButton optDestinatario 
               Caption         =   "Destinatário"
               Height          =   195
               Left            =   1680
               TabIndex        =   4
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optRemetente 
               Caption         =   "Remetente"
               Height          =   195
               Left            =   120
               TabIndex        =   3
               Top             =   260
               Value           =   -1  'True
               Width           =   1215
            End
         End
      End
      Begin VB.Frame fraContatos 
         Caption         =   "Contatos"
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
         Height          =   2415
         Left            =   120
         TabIndex        =   43
         Top             =   3480
         Width           =   10215
         Begin VB.TextBox txtFoneCont2 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   25
            Top             =   1030
            Width           =   1935
         End
         Begin VB.TextBox txtContato2 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   23
            Top             =   1030
            Width           =   3615
         End
         Begin VB.TextBox txtFoneCont1 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   19
            Top             =   300
            Width           =   1935
         End
         Begin VB.TextBox txtContato1 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   17
            Top             =   300
            Width           =   3615
         End
         Begin VB.TextBox txtAvUsuCont3 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8760
            MaxLength       =   10
            TabIndex        =   34
            Top             =   2050
            Width           =   1215
         End
         Begin VB.CheckBox chkAvisarCont3 
            Caption         =   "Avisar Aniversário ao:"
            Height          =   195
            Left            =   8040
            TabIndex        =   33
            Top             =   1790
            Width           =   1935
         End
         Begin VB.TextBox txtFoneCont3 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   31
            Top             =   1760
            Width           =   1935
         End
         Begin VB.TextBox txtContato3 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   29
            Top             =   1760
            Width           =   3615
         End
         Begin VB.TextBox txtAniverCont3 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   6960
            MaxLength       =   5
            TabIndex        =   32
            Top             =   2050
            Width           =   735
         End
         Begin VB.TextBox txtEmailCont3 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   30
            Top             =   2050
            Width           =   3615
         End
         Begin VB.TextBox txtAvUsuCont2 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8760
            MaxLength       =   10
            TabIndex        =   28
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkAvisarCont2 
            Caption         =   "Avisar Aniversário ao:"
            Height          =   195
            Left            =   8040
            TabIndex        =   27
            Top             =   1060
            Width           =   1950
         End
         Begin VB.TextBox txtAniverCont2 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   6960
            MaxLength       =   5
            TabIndex        =   26
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtEmailCont2 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   24
            Top             =   1320
            Width           =   3615
         End
         Begin VB.CheckBox chkAvisarCont1 
            Caption         =   "Avisar Aniversário ao:"
            Height          =   195
            Left            =   8040
            TabIndex        =   21
            Top             =   330
            Width           =   1970
         End
         Begin VB.TextBox txtAniverCont1 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   6960
            MaxLength       =   5
            TabIndex        =   20
            Top             =   590
            Width           =   735
         End
         Begin VB.TextBox txtEmailCont1 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   18
            Top             =   590
            Width           =   3615
         End
         Begin VB.TextBox txtAvUsuCont1 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8760
            MaxLength       =   10
            TabIndex        =   22
            Top             =   590
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   10080
            Y1              =   1660
            Y2              =   1660
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   8040
            TabIndex        =   61
            Top             =   2085
            Width           =   585
         End
         Begin VB.Label Label31 
            Caption         =   "Fones:"
            Height          =   255
            Left            =   5160
            TabIndex        =   60
            Top             =   1755
            Width           =   495
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   480
            TabIndex        =   59
            Top             =   1755
            Width           =   465
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Aniversário (DD/MM):"
            Height          =   195
            Left            =   5160
            TabIndex        =   58
            Top             =   2085
            Width           =   1545
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Left            =   480
            TabIndex        =   57
            Top             =   2085
            Width           =   420
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   10080
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   8040
            TabIndex        =   56
            Top             =   1350
            Width           =   585
         End
         Begin VB.Label Label25 
            Caption         =   "Fones:"
            Height          =   255
            Left            =   5160
            TabIndex        =   55
            Top             =   1035
            Width           =   495
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   480
            TabIndex        =   54
            Top             =   1035
            Width           =   465
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Aniversário (DD/MM):"
            Height          =   195
            Left            =   5160
            TabIndex        =   53
            Top             =   1350
            Width           =   1545
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Left            =   480
            TabIndex        =   52
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   8040
            TabIndex        =   51
            Top             =   615
            Width           =   585
         End
         Begin VB.Label Label17 
            Caption         =   "Fones:"
            Height          =   255
            Left            =   5160
            TabIndex        =   50
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   480
            TabIndex        =   49
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "1:"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   300
            Width           =   135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "2:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   1035
            Width           =   135
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "3:"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   1755
            Width           =   135
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Aniversário (DD/MM):"
            Height          =   195
            Left            =   5160
            TabIndex        =   45
            Top             =   615
            Width           =   1545
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Left            =   480
            TabIndex        =   44
            Top             =   615
            Width           =   420
         End
      End
      Begin VB.Label lblcgc 
         AutoSize        =   -1  'True
         Caption         =   "CGC"
         Height          =   195
         Left            =   120
         TabIndex        =   123
         Top             =   240
         Visible         =   0   'False
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmCadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAlarme_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub chkAvisarCont1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub chkAvisarCont2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub chkAvisarCont3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub chkConsigDevolProp_Click()
    If chkConsigDevolProp.Value = 1 Then
        txtConsigDevol = txtCgc
        txtConsigDevol.Enabled = False
        lblConsigDevolNome = txtRazao
    ElseIf chkConsigDevolProp.Value = 0 Then
        txtConsigDevol.Enabled = True
    End If
End Sub

Private Sub chkConsigDevolProp_LostFocus()
    chkConsigDevolProp_Click
End Sub

Private Sub chkConsigEntrProp_Click()
    If chkConsigEntrProp.Value = 1 Then
        txtConsigEntrega = txtCgc
        txtConsigEntrega.Enabled = False
        lblConsigEntregaNome = txtRazao
    ElseIf chkConsigEntrProp.Value = 0 Then
        txtConsigEntrega.Enabled = True
    End If
End Sub

Private Sub chkConsigEntrProp_LostFocus()
    chkConsigEntrProp_Click
End Sub

Private Sub chkConsigEntrPropAir_Click()
    If chkConsigEntrPropAir.Value = 1 Then
        txtConsigEntregaAir = txtCgc
        txtConsigEntregaAir.Enabled = False
        lblConsigEntregaNomeAir = txtRazao
    ElseIf chkConsigEntrPropAir.Value = 0 Then
        txtConsigEntregaAir.Enabled = True
    End If
End Sub

Private Sub chkConsigTransfProp_Click()
    If chkConsigTransfProp.Value = 1 Then
        txtConsigTransf = txtCgc
        txtConsigTransf.Enabled = False
        lblConsigTransfNome = txtRazao
    ElseIf chkConsigTransfProp.Value = 0 Then
        txtConsigTransf.Enabled = True
    End If
End Sub

Private Sub chkConsigTransfProp_LostFocus()
    chkConsigTransfProp_Click
End Sub

Private Sub chkTabTodosProd_Click()
    If chkTabTodosProd.Value = 0 Then
        txtTabProd.Enabled = True
        txtTabProd.BackColor = xamarelo1
        txtTabBuscaProd.Enabled = True
    ElseIf chkTabTodosProd.Value = 1 Then
        txtTabProd.Enabled = False
        txtTabProd.Text = ""
        txtTabProd.BackColor = xbranco
        txtTabBuscaProd.Enabled = False
    End If
End Sub

Private Sub chkTabTodosProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub chkTabTodosRemet_Click()
    If chkTabTodosRemet.Value = 0 Then
        txtTabRemet.Enabled = True
        txtTabRemet.BackColor = xamarelo1
        txtTabBuscaRemet.Enabled = True
        txtTabRemet.SetFocus
    ElseIf chkTabTodosRemet.Value = 1 Then
        txtTabRemet.Enabled = False
        txtTabRemet.BackColor = xbranco
        txtTabRemet.Text = ""
        txtTabBuscaRemet.Enabled = False
        lblTabDescrRemet.Caption = ""
        chkTabTodosProd.Enabled = False
        txtTabProd.Enabled = False
        txtTabProd.BackColor = xbranco
        txtTabProd.Text = ""
        txtTabBuscaProd.Enabled = False
        lblTabDescrProd.Caption = ""

    End If
End Sub

Private Sub chkTabTodosRemet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdAlterar_Click()
    fraCliente.Enabled = True
    fraDados.Enabled = True
    fraContatos.Enabled = True
    fraConsig.Enabled = True
    fraDiversos.Enabled = True
    cmdDetalhaBloq.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdGravar.Enabled = True
    cmdHistorico.Enabled = False
    cmdBloquear.Enabled = False
    cmdGo.Enabled = False
    cmdBusca.Enabled = False
    txtCgc.Enabled = False
    txtCgc.BackColor = &H8000000E
    lblAcao = "Alteração"
    cmdSair.Caption = "Cancelar"
    TravaTela frmCadClientes, "D"
    txtFantasia.Enabled = False
    txtApelido.Enabled = False
    txtRazao.Enabled = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    txtEndereco.SetFocus
End Sub
Private Sub cmdAlterarProd_Click()
    If cmdAlterarProd.Caption = "Alterar" Then
    
        cmdAlterarProd.Caption = "Cancelar"
        
        cmdGravarProd.Enabled = True
        cmdBuscaIata.Enabled = True
        cmdIncluirProd.Enabled = False
        cmdSairProd.Enabled = False
        
        txtCodIata.Enabled = True
        txtObsProd.Enabled = True
        
        txtCodIata.BackColor = xamarelo1
        txtObsProd.BackColor = xamarelo1
        
        fraGridProd.Enabled = False
        
        lblAcaoProd = "Alteração"
        
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        
        txtCodIata.SetFocus
        
    ElseIf cmdAlterarProd.Caption = "Cancelar" Then
    
        cmdAlterarProd.Caption = "Alterar"
        
        cmdGravarProd.Enabled = False
        cmdBuscaIata.Enabled = False
        cmdIncluirProd.Enabled = True
        cmdSairProd.Enabled = True
        
        txtProdNatureza.Enabled = False
        txtCodIata.Enabled = False
        txtObsProd.Enabled = False
        
        txtProdNatureza.BackColor = xbranco
        txtCodIata.BackColor = xbranco
        txtObsProd.BackColor = xbranco
        
        fraGridProd.Enabled = True
        
        lblAcaoProd = "Consulta"
        
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
    
    End If

End Sub

Private Sub cmdBloquear_Click()
    If cmdBloquear.Caption = "Bloquear ..." Then
        frmBloqueioCliente.Caption = "Bloqueio deste Cliente"
        frmBloqueioCliente.cmdBloquear.Caption = "Bloquear"
        frmBloqueioCliente.lblUsuario = xusuario
        frmBloqueioCliente.lblData = datahora("DATAHORA")
        frmBloqueioCliente.Show 1
        xcgc = txtCgc
        cmdSair_Click
        txtCgc = xcgc
        cmdGo_Click
    ElseIf cmdBloquear.Caption = "Desbloquear..." Then
        frmBloqueioCliente.Caption = "Desbloqueio deste Cliente"
        frmBloqueioCliente.cmdBloquear.Caption = "Desbloquear"
        frmBloqueioCliente.lblUsuario = xusuario
        frmBloqueioCliente.lblData = datahora("DATAHORA")
        frmBloqueioCliente.Show 1
        xcgc = txtCgc
        cmdSair_Click
        txtCgc = xcgc
        cmdGo_Click
    End If
End Sub
Private Sub cmdBusca_Click()
    frmBuscaSubClientes.Caption = "Busca Cadastro de Clientes - Cadastro"
    frmBuscaSubClientes.Show 1
End Sub

Private Sub cmdBuscaClasse_Click()
    If Len(Trim$(lblUf)) = 0 Then
        MsgBox "Cadastre primeiro a Cidade/UF !"
        txtCidade.SetFocus
        Exit Sub
    Else
        frmBuscaClasseFiscal.Show 1
    End If
End Sub

Private Sub cmdBuscaConsigDevol_Click()
    frmBuscaSubClientes.Caption = "Busca Cadastro de Clientes/Consignatário (DEVOLUÇÃO)"
    frmBuscaSubClientes.Show 1
End Sub
Private Sub cmdBuscaConsigEntr_Click()
    frmBuscaSubClientes.Caption = "Busca Cadastro de Clientes/Consignatário (ENTREGA RODO)"
    frmBuscaSubClientes.Show 1
End Sub

Private Sub cmdBuscaConsigEntrAir_Click()
    frmBuscaSubClientes.Caption = "Busca Cadastro de Clientes/Consignatário (ENTREGA AÉREO)"
    frmBuscaSubClientes.Show 1
End Sub

Private Sub cmdBuscaConsigTransf_Click()
    frmBuscaSubClientes.Caption = "Busca Cadastro de Clientes/Consignatário (TRANSFERÊNCIA)"
    frmBuscaSubClientes.Show 1
End Sub

Private Sub cmdDetalhaBloq_Click()
    frmBloqueioCliente.txtMotivo.Enabled = False
    frmBloqueioCliente.cmdBloquear.Visible = False
    frmBloqueioCliente.cmdDesistir.Caption = "Voltar"
    frmBloqueioCliente.txtMotivo.Text = lblDescrBloqueio
    frmBloqueioCliente.lblUsuario = lblUsuBloqueio
    frmBloqueioCliente.lblData = lblDataBloqueio
    frmBloqueioCliente.Show 1
End Sub
Private Sub cmdGo_Click()
    If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
    de_informaEM.Sel_CadCliCGC Trim$(txtCgc)
    If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
        If de_informaEM.rsSel_CadCliCGC.Fields("rem_des_log") = "REM" Then
            optRemetente = True
        ElseIf de_informaEM.rsSel_CadCliCGC.Fields("rem_des_log") = "DES" Then
            optDestinatario = True
        ElseIf de_informaEM.rsSel_CadCliCGC.Fields("rem_des_log") = "LOG" Then
            optLogistico = True
        End If
        txtRazao = de_informaEM.rsSel_CadCliCGC.Fields("nome")
        chkAlarme.Value = 0
        If de_informaEM.rsSel_CadCliCGC.Fields("alarm_ger") = "S" Then
            chkAlarme.Value = 1
        End If
        lblcgc = txtCgc
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("fantasia")) Then
            txtFantasia = ""
        Else
            txtFantasia = de_informaEM.rsSel_CadCliCGC.Fields("fantasia")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("apelido")) Then
            txtApelido = ""
        Else
            txtApelido = de_informaEM.rsSel_CadCliCGC.Fields("apelido")
        End If
        txtEndereco = de_informaEM.rsSel_CadCliCGC.Fields("endereco")
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("cep")) Then
            txtCep = ""
        Else
            txtCep = de_informaEM.rsSel_CadCliCGC.Fields("cep")
        End If
        txtCidade = de_informaEM.rsSel_CadCliCGC.Fields("cidade")
        lblUf = de_informaEM.rsSel_CadCliCGC.Fields("uf")
        txtIe = de_informaEM.rsSel_CadCliCGC.Fields("ie")
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("pabx")) Then
            txtPabx = ""
        Else
            txtPabx = de_informaEM.rsSel_CadCliCGC.Fields("pabx")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("fax")) Then
            txtFax = ""
        Else
            txtFax = de_informaEM.rsSel_CadCliCGC.Fields("fax")
        End If
        
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("contato1")) Then
            txtContato1 = ""
        Else
            txtContato1 = de_informaEM.rsSel_CadCliCGC.Fields("contato1")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("fonecontato1")) Then
            txtFoneCont1 = ""
        Else
            txtFoneCont1 = de_informaEM.rsSel_CadCliCGC.Fields("fonecontato1")
        End If
        chkAvisarCont1.Value = 0
        If Not IsNull(de_informaEM.rsSel_CadCliCGC.Fields("avisarcontato1")) Then
            If de_informaEM.rsSel_CadCliCGC.Fields("avisarcontato1") = "S" Then
                chkAvisarCont1.Value = 1
            End If
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("emailcontato1")) Then
            txtEmailCont1 = ""
        Else
            txtEmailCont1 = de_informaEM.rsSel_CadCliCGC.Fields("emailcontato1")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("anivercontato1")) Then
            txtAniverCont1 = ""
        Else
            txtAniverCont1 = de_informaEM.rsSel_CadCliCGC.Fields("anivercontato1")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("avusucontato1")) Then
            txtAvUsuCont1 = ""
        Else
            txtAvUsuCont1 = de_informaEM.rsSel_CadCliCGC.Fields("avusucontato1")
        End If
        
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("contato2")) Then
            txtContato2 = ""
        Else
            txtContato2 = de_informaEM.rsSel_CadCliCGC.Fields("contato2")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("fonecontato2")) Then
            txtFoneCont2 = ""
        Else
            txtFoneCont2 = de_informaEM.rsSel_CadCliCGC.Fields("fonecontato2")
        End If
        chkAvisarCont2.Value = 0
        If Not IsNull(de_informaEM.rsSel_CadCliCGC.Fields("avisarcontato2")) Then
            If de_informaEM.rsSel_CadCliCGC.Fields("avisarcontato2") = "S" Then
                chkAvisarCont2.Value = 1
            End If
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("emailcontato2")) Then
            txtEmailCont2 = ""
        Else
            txtEmailCont2 = de_informaEM.rsSel_CadCliCGC.Fields("emailcontato2")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("anivercontato2")) Then
            txtAniverCont2 = ""
        Else
            txtAniverCont2 = de_informaEM.rsSel_CadCliCGC.Fields("anivercontato2")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("avusucontato2")) Then
            txtAvUsuCont2 = ""
        Else
            txtAvUsuCont2 = de_informaEM.rsSel_CadCliCGC.Fields("avusucontato2")
        End If
        
        
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("contato3")) Then
            txtContato3 = ""
        Else
            txtContato3 = de_informaEM.rsSel_CadCliCGC.Fields("contato3")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("fonecontato3")) Then
            txtFoneCont3 = ""
        Else
            txtFoneCont3 = de_informaEM.rsSel_CadCliCGC.Fields("fonecontato3")
        End If
        chkAvisarCont3.Value = 0
        If Not IsNull(de_informaEM.rsSel_CadCliCGC.Fields("avisarcontato3")) Then
            If de_informaEM.rsSel_CadCliCGC.Fields("avisarcontato3") = "S" Then
                chkAvisarCont3.Value = 1
            End If
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("emailcontato3")) Then
            txtEmailCont3 = ""
        Else
            txtEmailCont3 = de_informaEM.rsSel_CadCliCGC.Fields("emailcontato3")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("anivercontato3")) Then
            txtAniverCont3 = ""
        Else
            txtAniverCont3 = de_informaEM.rsSel_CadCliCGC.Fields("anivercontato3")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("avusucontato3")) Then
            txtAvUsuCont3 = ""
        Else
            txtAvUsuCont3 = de_informaEM.rsSel_CadCliCGC.Fields("avusucontato3")
        End If
        
        
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("consigentrega")) Then
            txtConsigEntrega = ""
        Else
            txtConsigEntrega = de_informaEM.rsSel_CadCliCGC.Fields("consigentrega")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("consigentregaair")) Then
            txtConsigEntregaAir = ""
        Else
            txtConsigEntregaAir = de_informaEM.rsSel_CadCliCGC.Fields("consigentregaair")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("consigtransf")) Then
            txtConsigTransf = ""
        Else
            txtConsigTransf = de_informaEM.rsSel_CadCliCGC.Fields("consigtransf")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("consigdevol")) Then
            txtConsigDevol = de_informaEM.rsSel_CadCliCGC.Fields("consigdevol")
        Else
            txtConsigDevol = de_informaEM.rsSel_CadCliCGC.Fields("consigdevol")
        End If

        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("atendusu")) Then
            txtAtend1 = ""
        Else
            txtAtend1 = de_informaEM.rsSel_CadCliCGC.Fields("atendusu")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("prazo")) Then
            txtPrazo = ""
        Else
            txtPrazo = de_informaEM.rsSel_CadCliCGC.Fields("prazo")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("datacad")) Then
            lblDataCad = ""
        Else
            lblDataCad = de_informaEM.rsSel_CadCliCGC.Fields("datacad")
        End If
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("usuariocad")) Then
            lblUsuCad = ""
        Else
            lblUsuCad = de_informaEM.rsSel_CadCliCGC.Fields("usuariocad")
        End If
        If Not IsNull(de_informaEM.rsSel_CadCliCGC.Fields("ultemissao")) Then
            lblUltEmissao = de_informaEM.rsSel_CadCliCGC.Fields("ultemissao")
        Else
            lblUltEmissao = ""
        End If
        
        If Not IsNull(de_informaEM.rsSel_CadCliCGC.Fields("pessoafj")) Then
            If de_informaEM.rsSel_CadCliCGC.Fields("pessoafj") = "" Then
                optPessoaFisica.Value = False
                optPessoaFisica.Value = False
            ElseIf de_informaEM.rsSel_CadCliCGC.Fields("pessoafj") = "F" Then
                optPessoaFisica = True
            ElseIf de_informaEM.rsSel_CadCliCGC.Fields("pessoafj") = "J" Then
                optPessoaJuridica = True
            End If
        Else
            optPessoaFisica.Value = False
            optPessoaFisica.Value = False
        End If
            
        If Not IsNull(de_informaEM.rsSel_CadCliCGC.Fields("cfop")) Then
            If de_informaEM.rsSel_CadCliCGC.Fields("cfop") = "" Then
                lblClasseFiscal = ""
            Else
                lblClasseFiscal = de_informaEM.rsSel_CadCliCGC.Fields("cfop") & "-" & de_informaEM.rsSel_CadCliCGC.Fields("classefiscal")
            End If
        Else
            lblClasseFiscal = ""
        End If
        
        If IsNull(de_informaEM.rsSel_CadCliCGC.Fields("status")) Then
            LblStatus = ""
            cmdDetalhaBloq.Enabled = False
            cmdBloquear.Caption = ""
            lblDescrBloqueio = ""
            lblUsuBloqueio = ""
            lblDataBloqueio = ""
            If de_informaEM.rsSel_CadCliCGC.Fields("status") = "1" Then
                LblStatus = "ATIVO"
                cmdDetalhaBloq.Enabled = False
                cmdBloquear.Caption = "Bloquear ..."
                lblDescrBloqueio = ""
                lblUsuBloqueio = ""
                lblDataBloqueio = ""
            Else
                LblStatus = "BLOQUEADO"
                cmdDetalhaBloq.Enabled = True
                cmdBloquear.Caption = "Desbloquear..."
                lblDescrBloqueio = de_informaEM.rsSel_CadCliCGC.Fields("descrbloqueio")
                lblUsuBloqueio = de_informaEM.rsSel_CadCliCGC.Fields("usubloqueio")
                lblDataBloqueio = de_informaEM.rsSel_CadCliCGC.Fields("databloqueio")
            End If
        End If
        cmdAlterar.Enabled = True
        cmdHistorico.Enabled = True
        cmdBloquear.Enabled = True
        
        If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
        de_informaEM.Sel_CadCliCGC txtConsigEntrega
        If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
            lblConsigEntregaNome = de_informaEM.rsSel_CadCliCGC.Fields("nome")
        Else
            lblConsigEntregaNome = ""
        End If
        
        If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
        de_informaEM.Sel_CadCliCGC txtConsigEntregaAir
        If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
            lblConsigEntregaNomeAir = de_informaEM.rsSel_CadCliCGC.Fields("nome")
        Else
            lblConsigEntregaNomeAir = ""
        End If
        
        If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
        de_informaEM.Sel_CadCliCGC txtConsigTransf
        If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
            lblConsigTransfNome = de_informaEM.rsSel_CadCliCGC.Fields("nome")
        Else
            lblConsigTransfNome = ""
        End If
        If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
        de_informaEM.Sel_CadCliCGC txtConsigDevol
        If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
            lblConsigDevolNome = de_informaEM.rsSel_CadCliCGC.Fields("nome")
        Else
            lblConsigDevolNome = ""
        End If
        
        'dados de produtos (natureza) aba 2
        
        If de_informaEM.rsSel_CadCliProds.State = 1 Then de_informaEM.rsSel_CadCliProds.Close
        de_informaEM.Sel_CadCliProds Mid(txtCgc, 1, 8), "%"
        
        If de_informaEM.rsSel_CadCliProds.RecordCount > 0 Then
            fraGridProd.Enabled = True
        Else
            fraGridProd.Enabled = False
        End If
        
        gridProd.DataMember = "sel_cadcliprods"
        gridProd.Refresh
        
        lblTabConsig.Caption = txtCgc
        lblTabDescrConsig.Caption = txtRazao
        
        If de_informaEM.rsSel_CadCliProdsTAB.State = 1 Then de_informaEM.rsSel_CadCliProdsTAB.Close
        de_informaEM.Sel_CadCliProdsTAB lblTabConsig.Caption
        
        gridTabTabelas.DataMember = "Sel_CadCliProdsTAB"
        gridTabTabelas.Refresh
        
        If de_informaEM.rsSel_CadCliProdsTAB.RecordCount > 0 Then
            gridTabTabelas.Enabled = True
        Else
            gridTabTabelas.Enabled = False
        End If
        
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
        
        If txtCgc.Enabled = True Then txtCgc.SetFocus
        
        
    Else
        MsgBox "Não Encontrado !"
        limpatela frmCadClientes
        cmdAlterar.Enabled = False
        cmdHistorico.Enabled = False
        cmdBloquear.Enabled = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        txtCgc.SetFocus
    End If
End Sub
Private Sub cmdGravar_Click()
    Dim xcli_dest As String, xcgc As String, xrem_des_log As String, xalarme As String, xavisar1 As String, xavisar2 As String, xavisar3 As String
    If Len(Trim$(txtRazao)) < 6 Then
        MsgBox "Razao Social Inválida !"
        txtRazao.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtEndereco)) < 6 Then
        MsgBox "Endereço Inválido !"
        txtEndereco.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtCidade)) < 1 Then
        MsgBox "Cidade Inválida !"
        txtCidade.SetFocus
        Exit Sub
    End If
    
    If de_informaEM.rsSel_CadCidadePorCidadeUF.State = 1 Then de_informaEM.rsSel_CadCidadePorCidadeUF.Close
    de_informaEM.Sel_CadCidadePorCidadeUF Trim$(txtCidade), lblUf
    If de_informaEM.rsSel_CadCidadePorCidadeUF.RecordCount < 1 Then
        MsgBox "Cidade / UF Inválidos ! Utilize o Cadastramento pelo CEP."
        txtCep.SetFocus
        Exit Sub
    End If
    
    If optRemetente = True Then
        xrem_des_log = "REM"
    ElseIf optDestinatario = True Then
        xrem_des_log = "DES"
    ElseIf optLogistico = True Then
        xrem_des_log = "LOG"
    End If
    
    If chkAlarme.Value = 1 Then
        xalarme = "S"
    Else
        xalarme = "N"
    End If
    
    If chkAvisarCont1 = 1 Then
        xavisar1 = "S"
    Else
        xavisar1 = "N"
    End If
    If chkAvisarCont2 = 1 Then
        xavisar2 = "S"
    Else
        xavisar2 = "N"
    End If
    If chkAvisarCont3 = 1 Then
        xavisar3 = "S"
    Else
        xavisar3 = "N"
    End If
    
    If optPessoaFisica.Value = True Then
        xpessoa = "F"
    Else
        xpessoa = "J"
    End If
    
    If lblAcao = "Inclusão" Then
        
        de_informaEM.Ins_CadClientes txtCgc, Trim$(txtRazao), Trim$(txtFantasia), Trim$(txtApelido), Trim$(txtEndereco), _
                              "", txtCep, Trim$(txtCidade), Trim$(lblUf), Trim$(txtIe), Trim$(txtPabx), _
                              Trim$(txtFax), Trim$(txtContato1), Trim$(txtFoneCont1), Trim$(txtEmailCont1), Trim$(txtAniverCont1), _
                              xavisar1, Trim$(txtAvUsuCont1), Trim$(txtContato2), Trim$(txtFoneCont2), Trim$(txtEmailCont2), Trim$(txtAniverCont2), _
                              xavisar2, Trim$(txtAvUsuCont2), Trim$(txtContato3), Trim$(txtFoneCont3), Trim$(txtEmailCont3), Trim$(txtAniverCont3), _
                              xavisar3, Trim$(txtAvUsuCont3), txtConsigEntrega, txtConsigEntregaAir, txtConsigTransf, txtConsigDevol, txtAtend1, txtPrazo, xusuario, xrem_des_log, xalarme, _
                              Trim$(Mid$(lblClasseFiscal, 1, 4)), Trim$(Mid$(lblClasseFiscal, 6, 20)), xpessoa
                              
        MsgBox "Registro Incluso no Banco de Dados !"
        
    ElseIf lblAcao = "Alteração" Then
    
        de_informaEM.Alt_CadCliente txtCgc, Trim$(txtRazao), Trim$(txtFantasia), Trim$(txtApelido), Trim$(txtEndereco), _
                                    "", txtCep, Trim$(txtCidade), Trim$(lblUf), Trim$(txtIe), Trim$(txtPabx), _
                                    Trim$(txtFax), Trim$(txtContato1), Trim$(txtFoneCont1), Trim$(txtEmailCont1), Trim$(txtAniverCont1), _
                                    xavisar1, Trim$(txtAvUsuCont1), Trim$(txtContato2), Trim$(txtFoneCont2), Trim$(txtEmailCont2), Trim$(txtAniverCont2), _
                                    xavisar2, Trim$(txtAvUsuCont2), Trim$(txtContato3), Trim$(txtFoneCont3), Trim$(txtEmailCont3), Trim$(txtAniverCont3), _
                                    xavisar3, Trim$(txtAvUsuCont3), txtConsigEntrega, txtConsigEntregaAir, txtConsigTransf, txtConsigDevol, txtAtend1, txtPrazo, xrem_des_log, xalarme, _
                                    Trim$(Mid$(lblClasseFiscal, 1, 4)), Trim$(Mid$(lblClasseFiscal, 6, 20)), xpessoa
                                    
        MsgBox "Registro Alterado no Banco de Dados !"
    
    End If
    
    xcgc = txtCgc
    cmdSair_Click
    txtCgc = xcgc
    cmdGo_Click
    

End Sub
Private Sub cmdGravarProd_Click()
    
    If Len(Trim$(txtProdNatureza)) < 3 Then
        MsgBox "Descricao da Natureza de Produto Inválida !"
        txtProdNatureza.SetFocus
        Exit Sub
    End If
    If Len(Trim$(lblDescrIata)) < 3 Then
        MsgBox "Classificação IATA Inválida !"
        txtCodIata.SetFocus
        Exit Sub
    End If
    
    If lblAcaoProd.Caption = "Inclusão" Then
    
        If de_informaEM.rsSel_CadCliProds.State = 1 Then de_informaEM.rsSel_CadCliProds.Close
        de_informaEM.Sel_CadCliProds Mid(txtCgc, 1, 8), Trim$(txtProdNatureza)
        
        If de_informaEM.rsSel_CadCliProds.RecordCount > 0 Then
            MsgBox "Esta Natureza de Produto já Está Cadastrado no Banco de Dados !"
        
            If de_informaEM.rsSel_CadCliProds.State = 1 Then de_informaEM.rsSel_CadCliProds.Close
            de_informaEM.Sel_CadCliProds Mid(txtCgc, 1, 8), "%"
            
            txtProdNatureza.SetFocus
            Exit Sub
        End If
    
        de_informaEM.Ins_CadCliProds Mid(txtCgc, 1, 8), txtProdNatureza, txtCodIata, txtObsProd, xusuario
        cmdIncluirProd.Caption = "Incluir"
        
    ElseIf lblAcaoProd.Caption = "Alteração" Then
    
        de_informaEM.alt_CadCliProds txtCodIata, txtObsProd, Mid(txtCgc, 1, 8), txtProdNatureza
        cmdAlterarProd.Caption = "Alterar"
    
    End If
    
    If de_informaEM.rsSel_CadCliProds.State = 1 Then de_informaEM.rsSel_CadCliProds.Close
    de_informaEM.Sel_CadCliProds Mid(txtCgc, 1, 8), "%"
    gridProd.DataMember = "sel_cadcliprods"
    gridProd.Refresh
    
    fraGridProd.Enabled = True
    cmdGravarProd.Enabled = False
    cmdBuscaIata.Enabled = False
    cmdIncluirProd.Enabled = True
    cmdAlterarProd.Enabled = True
    cmdSairProd.Enabled = True
    
    txtProdNatureza.Enabled = False
    txtCodIata.Enabled = False
    txtObsProd.Enabled = False
    
    txtProdNatureza.BackColor = xbranco
    txtCodIata.BackColor = xbranco
    txtObsProd.BackColor = xbranco
    
    lblAcaoProd = "Consulta"
    
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    
End Sub
Private Sub cmdHistorico_Click()
    
    If de_informaEM.rsSel_CadCliHistorico.State = 1 Then de_informaEM.rsSel_CadCliHistorico.Close
    de_informaEM.Sel_CadCliHistorico Mid(txtCgc, 1, 8)
    
    frmCadCliHistorico.gridHistorico.DataMember = "sel_cadclihistorico"
    frmCadCliHistorico.gridHistorico.Refresh
    frmCadCliHistorico.lblcgc = txtCgc
    frmCadCliHistorico.lblRazao = txtRazao
    frmCadCliHistorico.Show 1
    
End Sub
Private Sub cmdIncluir_Click()
    fraCliente.Enabled = True
    fraDados.Enabled = True
    fraContatos.Enabled = True
    fraConsig.Enabled = True
    fraDiversos.Enabled = True
    cmdDetalhaBloq.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdGravar.Enabled = True
    cmdHistorico.Enabled = False
    cmdBloquear.Enabled = False
    cmdSair.Caption = "Cancelar"
    cmdGo.Enabled = False
    cmdBusca.Enabled = False
    limpatela frmCadClientes
    TravaTela frmCadClientes, "D"
    lblAcao = "Inclusão"
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    txtCgc.SetFocus
End Sub
Private Sub cmdIncluirProd_Click()
    
    If cmdIncluirProd.Caption = "Incluir" Then
    
        cmdIncluirProd.Caption = "Cancelar"
        
        cmdGravarProd.Enabled = True
        cmdAlterarProd.Enabled = False
        cmdBuscaIata.Enabled = True
        cmdSairProd.Enabled = False
        
        txtProdNatureza.Enabled = True
        txtCodIata.Enabled = True
        txtObsProd.Enabled = True
        
        txtProdNatureza.BackColor = xamarelo1
        txtCodIata.BackColor = xamarelo1
        txtObsProd.BackColor = xamarelo1
        txtProdNatureza.Text = ""
        txtCodIata.Text = ""
        txtObsProd.Text = ""
        lblDescrIata.Caption = ""
        
        fraGridProd.Enabled = False
        
        lblAcaoProd = "Inclusão"
        
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        
        txtProdNatureza.SetFocus
        
    ElseIf cmdIncluirProd.Caption = "Cancelar" Then
    
        cmdIncluirProd.Caption = "Incluir"
        
        cmdGravarProd.Enabled = False
        cmdAlterarProd.Enabled = False
        cmdBuscaIata.Enabled = False
        cmdSairProd.Enabled = True
        
        txtProdNatureza.Enabled = False
        txtCodIata.Enabled = False
        txtObsProd.Enabled = False
        
        txtProdNatureza.BackColor = xbranco
        txtCodIata.BackColor = xbranco
        txtObsProd.BackColor = xbranco
        
        fraGridProd.Enabled = True
        
        lblAcaoProd = "Consulta"
        
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
    
    End If
End Sub

Private Sub cmdSair_Click()
    If cmdSair.Caption = "Cancelar" Then
        fraCliente.Enabled = False
        fraDados.Enabled = False
        fraContatos.Enabled = False
        fraConsig.Enabled = False
        fraDiversos.Enabled = False
        cmdDetalhaBloq.Enabled = False
        cmdIncluir.Enabled = True
        cmdAlterar.Enabled = False
        cmdGravar.Enabled = False
        cmdHistorico.Enabled = False
        cmdBloquear.Enabled = False
        If Len(Trim$(txtCgc)) >= 8 Then
            cmdGo.Enabled = True
        Else
            cmdGo.Enabled = False
        End If
        cmdBusca.Enabled = True
        cmdSair.Caption = "Sair"
        limpatela frmCadClientes
        TravaTela frmCadClientes, "T"
        txtCgc.Enabled = True
        txtFantasia.Enabled = False
        txtApelido.Enabled = False
        txtRazao.Enabled = False
        lblAcao = "Consulta"
        txtCgc.BackColor = &HC0FFFF
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
        txtCgc.SetFocus
    Else
        Unload Me
    End If
End Sub
Private Sub cmdSairProd_Click()
    Unload Me
End Sub

Private Sub cmdTabAirIncl_Click()
    frmBuscaTabPrecoAereo.Show 1
    If cmdTabGravar.Enabled = True Then
        cmdTabGravar.SetFocus
    End If
End Sub

Private Sub cmdTabGenIncl_Click()
    frmBuscaTabPrecoGener.Show 1
    If cmdTabGravar.Enabled = True Then
        cmdTabGravar.SetFocus
    End If
End Sub

Private Sub cmdTabGravar_Click()
    Dim xNatProd As String, xRemetCgc As String
    
    If chkTabTodosRemet.Value = 1 Then
        xRemetCgc = "TODOS"
    Else
        xRemetCgc = txtTabRemet
    End If
    
    If chkTabTodosProd.Value = 1 Then
        xNatProd = "TODOS"
    Else
        xNatProd = txtTabProd
    End If
    
    de_informaEM.Ins_CadCliProdTAB "1", lblTabConsig, xRemetCgc, lblTabDescrRemet, xNatProd, _
                                 lblTabTabela, lblTabDescrTab, lblTabDataIncl, lblTabUsuario, Mid$(lblTabModalTab, 1, 1)
    
    
    If de_informaEM.rsSel_CadCliProdsTAB.State = 1 Then de_informaEM.rsSel_CadCliProdsTAB.Close
    de_informaEM.Sel_CadCliProdsTAB lblTabConsig
    
    gridTabTabelas.DataMember = "Sel_CadCliProdsTab"
    gridTabTabelas.Refresh
    
    If de_informaEM.rsSel_CadCliProdsTAB.RecordCount > 0 Then
        gridTabTabelas.Enabled = True
    Else
        gridTabTabelas.Enabled = True
    End If
    
    cmdTabIncluirTab_Click
    
End Sub
Private Sub cmdTabGravar_GotFocus()
    If chkTabTodosRemet.Value = 0 And Len(Trim$(txtTabRemet)) = 0 Then chkTabTodosRemet.Value = 1
    If chkTabTodosProd.Value = 0 And Len(Trim$(txtTabProd)) = 0 Then chkTabTodosProd.Value = 1
End Sub

Private Sub cmdTabIncluirTab_Click()
    If cmdTabIncluirTab.Caption = "Cancelar" Then
        cmdTabIncluirTab.Caption = "Incluir Tabela"
        chkTabTodosRemet.Enabled = False
        chkTabTodosRemet.Value = 0
        chkTabTodosRemet.Enabled = False
        chkTabTodosProd.Value = 0
        txtTabRemet.Enabled = False
        txtTabRemet.BackColor = xbranco
        txtTabProd.Enabled = False
        txtTabProd.BackColor = xbranco
        fraTabIncl.Enabled = False
        cmdTabRodoIncl.Enabled = False
        cmdTabAirIncl.Enabled = False
        cmdTabGenIncl.Enabled = False
        fraTabsPreco.Enabled = True
        cmdTabGravar.Enabled = False
        cmdTabDesabilitar.Enabled = False
        cmdTabDetalhar.Enabled = False
        txtTabRemet = ""
        txtTabProd = ""
        lblTabDescrRemet = ""
        lblTabDescrProd = ""
        lblTabTabela = ""
        lblTabDescrTab = ""
        lblTabModalTab = ""
        lblTabDataIncl = ""
        lblTabUsuario = ""
    ElseIf cmdTabIncluirTab.Caption = "Incluir Tabela" Then
        cmdTabIncluirTab.Caption = "Cancelar"
        chkTabTodosRemet.Enabled = True
        chkTabTodosRemet.Value = 1
        chkTabTodosProd.Value = 1
        fraTabIncl.Enabled = True
        cmdTabRodoIncl.Enabled = True
        cmdTabAirIncl.Enabled = True
        cmdTabGenIncl.Enabled = True
        fraTabsPreco.Enabled = False
        cmdTabDesabilitar.Enabled = False
        cmdTabDetalhar.Enabled = False
        txtTabRemet = ""
        txtTabProd = ""
        lblTabDescrRemet = ""
        lblTabDescrProd = ""
        lblTabTabela = ""
        lblTabDescrTab = ""
        lblTabModalTab = ""
        lblTabDataIncl = ""
        lblTabUsuario = ""
    End If
End Sub
Private Sub cmdTabDetalhar_Click()
    
    frmDetalheTabAir.SSTab1.TabEnabled(0) = False
    frmDetalheTabAir.SSTab1.TabEnabled(1) = False
    frmDetalheTabAir.SSTab1.TabEnabled(2) = False
    
    If Mid$(gridProd.Columns(3), 1, 4) = "TA01" Then
        frmDetalheTabAir.SSTab1.Tab = 0
        frmDetalheTabAir.SSTab1.TabEnabled(0) = True
        If de_informaEM.rsSel_TA01Codigo.State = 1 Then de_informaEM.rsSel_TA01Codigo.Close
        de_informaEM.Sel_TA01Codigo gridProd.Columns(3)
        frmDetalheTabAir.gridDadosTabTA01.DataMember = "Sel_TA01Codigo"
        frmDetalheTabAir.gridDadosTabTA01.Refresh
        frmDetalheTabAir.Show 1
    ElseIf Mid$(gridProd.Columns(3), 1, 4) = "TA02" Then
        frmDetalheTabAir.SSTab1.Tab = 1
        frmDetalheTabAir.SSTab1.TabEnabled(1) = True
        If de_informaEM.rsSel_TA02Codigo.State = 1 Then de_informaEM.rsSel_TA02Codigo.Close
        de_informaEM.Sel_TA02Codigo gridProd.Columns(3)
        frmDetalheTabAir.gridDadosTabTA02.DataMember = "Sel_TA02Codigo"
        frmDetalheTabAir.gridDadosTabTA02.Refresh
        frmDetalheTabAir.Show 1
    ElseIf Mid$(gridProd.Columns(3), 1, 4) = "TA03" Then
        frmDetalheTabAir.SSTab1.TabEnabled(2) = True
        frmDetalheTabAir.SSTab1.Tab = 2
    End If
    
End Sub

Private Sub cmdTabRodoIncl_Click()
    frmBuscaTabPreco.Show 1
    If cmdTabGravar.Enabled = True Then
        cmdTabGravar.SetFocus
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdtabSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'mdiEmissao.toolMenu.Visible = False
    'mdiEmissao.StatusBar.Visible = False
    'mdiEmissao.mnuArquivo.Enabled = False
    'mdiEmissao.mnuCadastros.Enabled = False
    'mdiEmissao.mnuEmissao.Enabled = False
    'mdiEmissao.mnuRelat.Enabled = False
    'mdiEmissao.mnuSair.Enabled = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    mdiEmissao.toolMenu.Visible = True
'    mdiEmissao.StatusBar.Visible = True
'    mdiEmissao.mnuArquivo.Enabled = True
'    mdiEmissao.mnuCadastros.Enabled = True
'    mdiEmissao.mnuEmissao.Enabled = True
'    mdiEmissao.mnuRelat.Enabled = True
'    mdiEmissao.mnuSair.Enabled = True
End Sub

Private Sub gridProd_Click()
    txtProdNatureza.Text = gridProd.Columns(1)
    txtObsProd.Text = gridProd.Columns(3)
    txtCodIata.Text = gridProd.Columns(2)
    
    If de_informaEM.rsSel_ClassIATAPorCod.State = 1 Then de_informaEM.rsSel_ClassIATAPorCod.Close
    de_informaEM.Sel_ClassIATAPorCod txtCodIata
    lblDescrIata.Caption = de_informaEM.rsSel_ClassIATAPorCod.Fields("descricao")
    cmdAlterarProd.Enabled = True
    
End Sub
Private Sub gridProd_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gridProd_Click
End Sub

Private Sub gridTabTabelas_Click()
    txtTabRemet = gridTabTabelas.Columns(2)
    lblTabDescrRemet = gridTabTabelas.Columns(3)
    txtTabProd = gridTabTabelas.Columns(4)
    lblTabTabela = gridTabTabelas.Columns(5)
    lblTabDescrTab = gridTabTabelas.Columns(6)
    lblTabModalTab = gridTabTabelas.Columns(7)
    lblTabDataIncl = gridTabTabelas.Columns(8)
    lblTabUsuario = gridTabTabelas.Columns(9)
End Sub

Private Sub gridTabTabelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gridTabTabelas_Click
End Sub

Private Sub optDestinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub
Private Sub optLogistico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub optRemetente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtAniverCont1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtAniverCont1_LostFocus()
    If Len(Trim$(txtAniverCont1)) > 0 Then
        If Len(Trim$(txtAniverCont1)) = 4 Then txtAniverCont1 = Mid$(txtAniverCont1, 1, 2) & "/" & Mid$(txtAniverCont1, 3, 2)
        If IsDate("2000/" & Mid$(txtAniverCont1, 4, 2) & "/" & Mid$(txtAniverCont1, 1, 2)) = False Then
            MsgBox "Data Inválida. Digite no Formato DD/MM (DD=Dia , MM=Mês ambos com 2 dígitos) !"
            txtAniverCont1.SetFocus
        End If
    End If
End Sub

Private Sub txtAniverCont2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtAniverCont2_LostFocus()
    If Len(Trim$(txtAniverCont2)) > 0 Then
        If Len(Trim$(txtAniverCont2)) = 4 Then txtAniverCont2 = Mid$(txtAniverCont2, 1, 2) & "/" & Mid$(txtAniverCont2, 3, 2)
        If IsDate("2000/" & Mid$(txtAniverCont2, 4, 2) & "/" & Mid$(txtAniverCont2, 1, 2)) = False Then
            MsgBox "Data Inválida. Digite no Formato DD/MM (DD=Dia , MM=Mês ambos com 2 dígitos) !"
            txtAniverCont2.SetFocus
        End If
    End If
End Sub

Private Sub txtAniverCont3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAniverCont3_LostFocus()
    If Len(Trim$(txtAniverCont3)) > 0 Then
        If Len(Trim$(txtAniverCont3)) = 4 Then txtAniverCont3 = Mid$(txtAniverCont3, 1, 2) & "/" & Mid$(txtAniverCont3, 3, 2)
        If IsDate("2000/" & Mid$(txtAniverCont3, 4, 2) & "/" & Mid$(txtAniverCont3, 1, 2)) = False Then
            MsgBox "Data Inválida. Digite no Formato DD/MM (DD=Dia , MM=Mês ambos com 2 dígitos) !"
            txtAniverCont3.SetFocus
        End If
    End If
End Sub

Private Sub txtApelido_GotFocus()
    txtApelido.BackColor = xamarelo2
End Sub

Private Sub txtApelido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtApelido_LostFocus()
    txtApelido = UCase(txtApelido)
    txtApelido.BackColor = xamarelo1
End Sub

Private Sub txtAtend1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAtend1_LostFocus()
    txtAtend1 = UCase(txtAtend1)
    If Len(Trim$(txtAtend1)) > 0 Then
        If de_informaEM.rsSel_CadUsuarioPorUsu.State = 1 Then de_informaEM.rsSel_CadUsuarioPorUsu.Close
        de_informaEM.Sel_CadUsuarioPorUsu txtAtend1
        If de_informaEM.rsSel_CadUsuarioPorUsu.RecordCount < 1 Then
            MsgBox "Nome de Usuário não Encontrado !"
            txtAtend1.SetFocus
        Else
            If de_informaEM.rsSel_CadUsuarioPorUsu.Fields("status") = "0" Then
                MsgBox "Este Usuário Encontra-se Bloqueado de Acesso ao Sistema !"
                txtAtend1.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtAtend2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAtend2_LostFocus()
    txtAtend2 = UCase(txtAtend2)
End Sub

Private Sub txtAtend3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAtend3_LostFocus()
    txtAtend3 = UCase(txtAtend3)
End Sub

Private Sub txtAvUsuCont1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtAvUsuCont1_LostFocus()
    txtAvUsuCont1 = UCase(txtAvUsuCont1)
    If Len(Trim$(txtAvUsuCont1)) > 0 Then
        If de_informaEM.rsSel_CadUsuarioPorUsu.State = 1 Then de_informaEM.rsSel_CadUsuarioPorUsu.Close
        de_informaEM.Sel_CadUsuarioPorUsu txtAvUsuCont1
        If de_informaEM.rsSel_CadUsuarioPorUsu.RecordCount < 1 Then
            MsgBox "Nome de Usuário não Encontrado !"
            txtAvUsuCont1.SetFocus
        Else
            If de_informaEM.rsSel_CadUsuarioPorUsu.Fields("status") = "0" Then
                MsgBox "Este Usuário Encontra-se Bloqueado de Acesso ao Sistema !"
                txtAvUsuCont1.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtAvUsuCont2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtAvUsuCont2_LostFocus()
    txtAvUsuCont2 = UCase(txtAvUsuCont2)
    If Len(Trim$(txtAvUsuCont2)) > 0 Then
        If de_informaEM.rsSel_CadUsuarioPorUsu.State = 1 Then de_informaEM.rsSel_CadUsuarioPorUsu.Close
        de_informaEM.Sel_CadUsuarioPorUsu txtAvUsuCont2
        If de_informaEM.rsSel_CadUsuarioPorUsu.RecordCount < 1 Then
            MsgBox "Nome de Usuário não Encontrado !"
            txtAvUsuCont2.SetFocus
        Else
            If de_informaEM.rsSel_CadUsuarioPorUsu.Fields("status") = "0" Then
                MsgBox "Este Usuário Encontra-se Bloqueado de Acesso ao Sistema !"
                txtAvUsuCont2.SetFocus
            End If
        End If
    End If

End Sub

Private Sub txtAvUsuCont3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAvUsuCont3_LostFocus()
    txtAvUsuCont3 = UCase(txtAvUsuCont3)
    If Len(Trim$(txtAvUsuCont3)) > 0 Then
        If de_informaEM.rsSel_CadUsuarioPorUsu.State = 1 Then de_informaEM.rsSel_CadUsuarioPorUsu.Close
        de_informaEM.Sel_CadUsuarioPorUsu txtAvUsuCont3
        If de_informaEM.rsSel_CadUsuarioPorUsu.RecordCount < 1 Then
            MsgBox "Nome de Usuário não Encontrado !"
            txtAvUsuCont3.SetFocus
        Else
            If de_informaEM.rsSel_CadUsuarioPorUsu.Fields("status") = "0" Then
                MsgBox "Este Usuário Encontra-se Bloqueado de Acesso ao Sistema !"
                txtAvUsuCont3.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtCep_Change()
    If Not IsNumeric(txtCep) Then
        SendKeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtCep_LostFocus()
    If Len(Trim$(txtCep)) > 0 Then
        If de_informaEM.rsSel_CidadeCEP.State = 1 Then de_informaEM.rsSel_CidadeCEP.Close
        de_informaEM.Sel_CidadeCEP txtCep, txtCep
        If de_informaEM.rsSel_CidadeCEP.RecordCount > 0 Then
            txtCidade = de_informaEM.rsSel_CidadeCEP.Fields("cidade")
            lblUf = de_informaEM.rsSel_CidadeCEP.Fields("uf")
        Else
            MsgBox "CEP Não Encontrado. Se o Mesmo estiver Correto, Solicite a sua Inclusão no Banco de Dados !"
            txtCidade = ""
            lblUf = ""
        End If
    End If
End Sub

Private Sub txtCgc_Change()
    If Not IsNumeric(txtCgc) Then
        SendKeys "{BACKSPACE}"
    End If
    If Len(Trim$(txtCgc.Text)) > 7 And lblAcao = "Consulta" Then
        cmdGo.Enabled = True
    Else
        cmdGo.Enabled = False
    End If
End Sub

Private Sub txtCgc_GotFocus()
    txtCgc.BackColor = xamarelo2
End Sub

Private Sub txtCgc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtCgc_LostFocus()
    Dim xcgc As String
    txtCgc.BackColor = xamarelo1
    If Len(Trim$(txtCgc)) > 0 Then
        If (Len(Trim$(txtCgc)) = 11 Or Len(Trim$(txtCgc)) = 14) Then
            If Len(Trim$(txtCgc)) > 0 And lblAcao = "Inclusão" Then
                If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
                de_informaEM.Sel_CadCliCGC Trim$(txtCgc)
                If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
                    MsgBox "Este CNPJ/CPF Já Encontra-se Cadastrado no Banco de Dados !"
                    txtCgc.SetFocus
                Else
                    If isCNPJ(txtCgc) = False And isCPF(txtCgc) = False Then
                        MsgBox "Número de CNPJ ou CPF Inválido !"
                        'txtCgc.SetFocus
                    End If
                End If
            End If
            If Len(Trim$(txtCgc)) > 0 And lblAcao = "Consulta" Then
                If isCNPJ(txtCgc) = False And isCPF(txtCgc) = False Then
                    MsgBox "Número de CNPJ ou CPF Inválido !"
                    'txtCgc.SetFocus
                End If
                If lblcgc <> txtCgc And Len(Trim$(txtRazao)) > 2 Then
                    xcgc = txtCgc
                    limpatela frmCadClientes
                    txtCgc = xcgc
                    cmdGo.SetFocus
                End If
            End If
        Else
            MsgBox "Este Número deve Ter 11 (CPF) ou 14 (CNPJ) caracteres !"
            txtCgc.SetFocus
        End If
    End If
End Sub

Private Sub txtCidade_Change()
    lblClasseFiscal = ""
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtCidade_LostFocus()
    If Len(Trim$(txtCidade.Text)) > 0 Then
        If de_informaEM.rsSel_CidadesLike.State = 1 Then de_informaEM.rsSel_CidadesLike.Close
        de_informaEM.Sel_CidadesLike Trim$(txtCidade)
        If de_informaEM.rsSel_CidadesLike.RecordCount < 1 Then  'não encontrou a cidade
            frmBuscaCidades.Caption = "Busca Cidades - Cad. Clientes"
            frmBuscaCidades.txtBuscaCidade = Trim$(txtCidade)
            frmBuscaCidades.Show 1
        ElseIf de_informaEM.rsSel_CidadesLike.RecordCount > 1 Then 'encontrou mais de uma cidade
            frmBuscaCidades.Caption = "Busca Cidades - Cad. Clientes"
            frmBuscaCidades.txtBuscaCidade = Trim$(txtCidade)
            frmBuscaCidades.lblfoco = "GRID"
            frmBuscaCidades.Show 1
        Else 'encontrou só uma e busca o demais dados
            lblUf = de_informaEM.rsSel_CidadesLike.Fields("uf")
            If de_informaEM.rsSel_CadCidadePorCidadeUF.State = 1 Then de_informaEM.rsSel_CadCidadePorCidadeUF.Close
            de_informaEM.Sel_CadCidadePorCidadeUF Trim$(txtCidade), lblUf
            txtCep = de_informaEM.rsSel_CadCidadePorCidadeUF.Fields("cepi")
        End If
    Else
        txtCep = ""
        lblUf = ""
    End If
    txtCidade = UCase(txtCidade)
End Sub

Private Sub txtCodIata_Change()
    If Not IsNumeric(txtCodIata) Then
        SendKeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtCodIata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCodIata_LostFocus()
    If Len(Trim$(txtCodIata)) > 0 Then
        If de_informaEM.rsSel_ClassIATAPorCod.State = 1 Then de_informaEM.rsSel_ClassIATAPorCod.Close
        de_informaEM.Sel_ClassIATAPorCod txtCodIata
        If de_informaEM.rsSel_ClassIATAPorCod.RecordCount > 0 Then
            lblDescrIata = de_informaEM.rsSel_ClassIATAPorCod.Fields("descricao")
        Else
            lblDescrIata = ""
        End If
    Else
        lblDescrIata = ""
    End If
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtComplemento_LostFocus()
    txtComplemento = UCase(txtComplemento)
End Sub

Private Sub txtConsigDevol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtConsigDevol_LostFocus()
    If Len(Trim$(txtConsigDevol)) > 0 Then
        If isCNPJ(txtConsigDevol) = False And isCPF(txtConsigDevol) = False Then
            MsgBox "Número de CNPJ ou CPF Inválido !"
            'txtConsigDevol.SetFocus
        End If
        If txtConsigDevol = txtCgc Then
            chkConsigDevolProp.Value = 1
            chkConsigDevolProp_Click
        Else
            If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
            de_informaEM.Sel_CadCliCGC Trim$(txtConsigDevol)
            If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
                lblConsigDevolNome = de_informaEM.rsSel_CadCliCGC.Fields("nome")
            Else
                MsgBox "CGC Não Encontrado no Banco de Dados !"
                txtConsigDevol.SetFocus
            End If
        End If
    Else
        lblConsigDevolNome = ""
    End If
End Sub

Private Sub txtConsigEntrega_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtConsigEntrega_LostFocus()
    If Len(Trim$(txtConsigEntrega)) > 0 Then
        If isCNPJ(txtConsigEntrega) = False And isCPF(txtConsigEntrega) = False Then
            MsgBox "Número de CNPJ ou CPF Inválido !"
            'txtConsigEntrega.SetFocus
        End If
        If txtConsigEntrega = txtCgc Then
            chkConsigEntrProp.Value = 1
            chkConsigEntrProp_Click
        Else
            If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
            de_informaEM.Sel_CadCliCGC Trim$(txtConsigEntrega)
            If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
                lblConsigEntregaNome = de_informaEM.rsSel_CadCliCGC.Fields("nome")
            Else
                MsgBox "CGC Não Encontrado no Banco de Dados !"
                txtConsigEntrega.SetFocus
            End If
        End If
    Else
        lblConsigEntregaNome = ""
    End If
End Sub

Private Sub txtConsigEntregaAir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtConsigEntregaAir_LostFocus()
    If Len(Trim$(txtConsigEntregaAir)) > 0 Then
        If isCNPJ(txtConsigEntregaAir) = False And isCPF(txtConsigEntregaAir) = False Then
            MsgBox "Número de CNPJ ou CPF Inválido !"
            'txtConsigEntrega.SetFocus
        End If
        If txtConsigEntregaAir = txtCgc Then
            chkConsigEntrPropAir.Value = 1
            chkConsigEntrPropAir_Click
        Else
            If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
            de_informaEM.Sel_CadCliCGC Trim$(txtConsigEntregaAir)
            If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
                lblConsigEntregaNomeAir = de_informaEM.rsSel_CadCliCGC.Fields("nome")
            Else
                MsgBox "CGC Não Encontrado no Banco de Dados !"
                txtConsigEntregaAir.SetFocus
            End If
        End If
    Else
        lblConsigEntregaNomeAir = ""
    End If
End Sub

Private Sub txtConsigTransf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtConsigTransf_LostFocus()
    If Len(Trim$(txtConsigTransf)) > 0 Then
        If isCNPJ(txtConsigTransf) = False And isCPF(txtConsigTransf) = False Then
            MsgBox "Número de CNPJ ou CPF Inválido !"
            'txtConsigTransf.SetFocus
        End If
        If txtConsigTransf = txtCgc Then
            chkConsigTransfProp.Value = 1
            chkConsigTransfProp_Click
        Else
            If de_informaEM.rsSel_CadCliCGC.State = 1 Then de_informaEM.rsSel_CadCliCGC.Close
            de_informaEM.Sel_CadCliCGC Trim$(txtConsigTransf)
            If de_informaEM.rsSel_CadCliCGC.RecordCount > 0 Then
                lblConsigTransfNome = de_informaEM.rsSel_CadCliCGC.Fields("nome")
            Else
                MsgBox "CGC Não Encontrado no Banco de Dados !"
                txtConsigTransf.SetFocus
            End If
        End If
    Else
        lblConsigTransfNome = ""
    End If
End Sub

Private Sub txtContato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtContato1_LostFocus()
    txtContato1 = UCase(txtContato1)
End Sub

Private Sub txtContato2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtContato2_LostFocus()
    txtContato2 = UCase(txtContato2)
End Sub

Private Sub txtContato3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtContato3_LostFocus()
    txtContato3 = UCase(txtContato3)
End Sub

Private Sub txtEmailCont1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtEmailCont1_LostFocus()
    txtEmailCont1 = LCase(txtEmailCont1)
End Sub

Private Sub txtEmailCont2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtEmailCont2_LostFocus()
    txtEmailCont2 = LCase(txtEmailCont2)
End Sub

Private Sub txtEmailCont3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEmailCont3_LostFocus()
    txtEmailCont3 = LCase(txtEmailCont3)
End Sub

Private Sub txtEndereco_GotFocus()
    txtEndereco.BackColor = xamarelo2
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtEndereco_LostFocus()
    txtEndereco = UCase(txtEndereco)
    txtEndereco.BackColor = xamarelo1
End Sub

Private Sub txtFantasia_GotFocus()
    txtFantasia.BackColor = xamarelo2
End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtFantasia_LostFocus()
    txtFantasia = UCase(txtFantasia)
    txtFantasia.BackColor = xamarelo1
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtFax_LostFocus()
    txtFax = UCase(txtFax)
End Sub

Private Sub txtFoneCont1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtFoneCont1_LostFocus()
    txtFoneCont1 = UCase(txtFoneCont1)
End Sub

Private Sub txtFoneCont2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtFoneCont2_LostFocus()
    txtFoneCont2 = UCase(txtFoneCont2)
End Sub

Private Sub txtFoneCont3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFoneCont3_LostFocus()
    txtFoneCont3 = UCase(txtFoneCont3)
End Sub

Private Sub txtIe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtIe_LostFocus()
    txtIe = UCase(txtIe)
End Sub

Private Sub txtObsProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtObsProd_LostFocus()
    txtObsProd = UCase(txtObsProd)
End Sub

Private Sub txtPabx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtPabx_LostFocus()
    txtPabx = UCase(txtPabx)
End Sub

Private Sub txtPrazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPrazo_LostFocus()
    txtPrazo.Text = UCase(txtPrazo)
End Sub

Private Sub txtProdNatureza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtProdNatureza_LostFocus()
    txtProdNatureza = UCase(txtProdNatureza)
End Sub

Private Sub txtRazao_GotFocus()
    txtRazao.BackColor = xamarelo2
End Sub

Private Sub txtRazao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtRazao_LostFocus()
    txtRazao = UCase(txtRazao)
    txtRazao.BackColor = xamarelo1
End Sub

Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtUF_LostFocus()
    txtUf = UCase(txtUf)
End Sub

Private Sub txtTabBuscaProd_Click()
    If de_informaEM.rsSel_CadCliProds.State = 1 Then de_informaEM.rsSel_CadCliProds.Close
    de_informaEM.Sel_CadCliProds txtTabRemet, "%"
    
    frmBuscaProdEmissao.gridprods.DataMember = "Sel_CadCliProds"
    frmBuscaProdEmissao.gridprods.Refresh
        
    frmBuscaProdEmissao.Caption = "Produtos do Cliente - Tab. Preço"
    frmBuscaProdEmissao.Label2.Visible = False
    frmBuscaProdEmissao.Show 1
    txtTabProd.SetFocus
    
    DoEvents
    
    If Len(Trim$(lblTabDescrProd)) = 3 Then
        If de_informaEM.rsSel_ClassIATAPorCod.State = 1 Then de_informaEM.rsSel_ClassIATAPorCod.Close
        de_informaEM.Sel_ClassIATAPorCod Trim$(lblTabDescrProd)
        If de_informaEM.rsSel_ClassIATAPorCod.RecordCount > 0 Then
            lblTabDescrProd = Trim$(lblTabDescrProd) & " - " & de_informaEM.rsSel_ClassIATAPorCod.Fields("descricao")
        Else
            lblTabDescrProd = ""
        End If
    End If
    
End Sub

Private Sub txtTabBuscaRemet_Click()
    frmBuscaSubClientes.Caption = "Busca Cadastro de Clientes - Tab. Preço"
    frmBuscaSubClientes.Show 1
    txtTabRemet.SetFocus
End Sub

Private Sub txtTabProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtTabProd_LostFocus()
    If Len(Trim(txtTabProd)) > 0 Then
        txtTabProd = UCase(txtTabProd)
        If de_informaEM.rsSel_CadCliProds.State = 1 Then de_informaEM.rsSel_CadCliProds.Close
        de_informaEM.Sel_CadCliProds Trim$(txtTabRemet), "%"
        If de_informaEM.rsSel_CadCliProds.RecordCount > 0 Then
            If de_informaEM.rsSel_ClassIATAPorCod.State = 1 Then de_informaEM.rsSel_ClassIATAPorCod.Close
            de_informaEM.Sel_ClassIATAPorCod de_informaEM.rsSel_CadCliProds.Fields("classiata")
            lblTabDescrProd = de_informaEM.rsSel_ClassIATAPorCod.Fields("descricao")
        Else
            MsgBox "Produto Não Encontrado para Este Cliente !"
            txtTabProd.SetFocus
            Exit Sub
        End If
    Else
        lblTabDescrProd = ""
    End If
End Sub

Private Sub txtTabRemet_Change()
    If Not IsNumeric(txtTabRemet) Then
        SendKeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtTabRemet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub
Private Sub txtTabRemet_LostFocus()
    If Len(Trim$(txtTabRemet)) > 0 Then
        txtTabRemet = Trim$(txtTabRemet)
        If de_informaEM.rsSel_CadCliCGCLike.State = 1 Then de_informaEM.rsSel_CadCliCGCLike.Close
        de_informaEM.Sel_CadCliCGCLike txtTabRemet & "%"
        If de_informaEM.rsSel_CadCliCGCLike.RecordCount > 0 Then
            lblTabDescrRemet = de_informaEM.rsSel_CadCliCGCLike.Fields("nome")
            chkTabTodosProd.Enabled = True
            chkTabTodosProd.SetFocus
        Else
            MsgBox "CNPJ Base não Encontrado no Banco de Dados !"
            chkTabTodosProd.Enabled = False
            txtTabRemet.SetFocus
        End If
    End If
End Sub
