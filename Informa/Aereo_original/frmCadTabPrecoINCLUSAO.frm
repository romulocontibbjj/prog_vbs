VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadTabPrecoINCLUSAO 
   Caption         =   "Cadastramento de Tabela de Preço"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadTabPrecoINCLUSAO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid FlexGridImportacao 
      Height          =   375
      Left            =   180
      TabIndex        =   58
      Top             =   7740
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      Enabled         =   0   'False
   End
   Begin VB.CommandButton CmdProximaFase 
      Caption         =   "Avançar >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9900
      TabIndex        =   11
      Top             =   6780
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancelarTodoProcesso 
      Caption         =   "Cancelar Todo o Processo de Cadastramento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   6780
      Width           =   7635
   End
   Begin VB.CommandButton CmdFaseAnterior 
      Caption         =   "<< Voltar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   6780
      Width           =   1815
   End
   Begin TabDlg.SSTab TabFase 
      Height          =   6465
      Left            =   180
      TabIndex        =   20
      Top             =   180
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   11404
      _Version        =   393216
      TabsPerRow      =   10
      TabHeight       =   556
      TabCaption(0)   =   "Fase 0"
      TabPicture(0)   =   "frmCadTabPrecoINCLUSAO.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraOrigem"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraLocalidades"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FraTipoTabela"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraCiaAerea"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdPlanilha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Fase 1"
      TabPicture(1)   =   "frmCadTabPrecoINCLUSAO.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label27"
      Tab(1).Control(1)=   "Line5"
      Tab(1).Control(2)=   "OptNavLateral"
      Tab(1).Control(3)=   "OptNavVertical"
      Tab(1).Control(4)=   "CmdZerarDigitacao"
      Tab(1).Control(5)=   "CmdIniciarDigitacao"
      Tab(1).Control(6)=   "FraPanoramaNovaTabela"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Fase 2"
      TabPicture(2)   =   "frmCadTabPrecoINCLUSAO.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LblOrigem"
      Tab(2).Control(1)=   "Line4"
      Tab(2).Control(2)=   "Label20"
      Tab(2).Control(3)=   "FraNovaTabela"
      Tab(2).Control(4)=   "CmdCadastrarTabela"
      Tab(2).Control(5)=   "FraVigencia"
      Tab(2).Control(6)=   "TxtDescrSistema"
      Tab(2).ControlCount=   7
      Begin VB.CommandButton CmdPlanilha 
         Caption         =   "..."
         Height          =   315
         Left            =   10680
         TabIndex        =   57
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox TxtDescrSistema 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74820
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   5970
         Width           =   5235
      End
      Begin VB.Frame FraVigencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -69480
         TabIndex        =   43
         Top             =   5760
         Width           =   3375
         Begin MSMask.MaskEdBox MskVigencia 
            Height          =   285
            Left            =   1920
            TabIndex        =   44
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Entrará em Vigência em"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   225
            Width           =   1680
         End
      End
      Begin VB.CommandButton CmdCadastrarTabela 
         Caption         =   "Cadastrar Nova Tabela"
         Height          =   375
         Left            =   -66000
         TabIndex        =   42
         Top             =   5940
         Width           =   2415
      End
      Begin VB.Frame FraNovaTabela 
         Caption         =   "Nova Tabela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74760
         TabIndex        =   39
         Top             =   960
         Width           =   11235
         Begin VB.TextBox TxtOBS 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   120
            MaxLength       =   450
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   3840
            Width           =   10995
         End
         Begin MSFlexGridLib.MSFlexGrid FlexGridNovaTabela 
            Height          =   2175
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   3836
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid FlexGridOrigem2 
            Height          =   855
            Left            =   120
            TabIndex        =   51
            Top             =   2700
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1508
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid FlexGridDestino2 
            Height          =   855
            Left            =   8340
            TabIndex        =   52
            Top             =   2700
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1508
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observações sobre a Nova Tabela:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   3600
            Width           =   2535
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Origem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   120
            TabIndex        =   54
            Top             =   2400
            Width           =   1890
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Destino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   9000
            TabIndex        =   53
            Top             =   2400
            Width           =   1965
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Aguarde enquanto a Nova Tabela é Atualizada..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   1620
            TabIndex        =   40
            Top             =   1800
            Width           =   7710
         End
      End
      Begin VB.Frame FraPanoramaNovaTabela 
         Caption         =   "Panorama da Nova Tabela  a Ser Criada"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -74820
         TabIndex        =   37
         Top             =   1440
         Width           =   11235
         Begin MSFlexGridLib.MSFlexGrid FlexGridPanoramaNovaTabela 
            Height          =   3315
            Left            =   120
            TabIndex        =   18
            Top             =   180
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   5847
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid FlexGridOrigem 
            Height          =   855
            Left            =   180
            TabIndex        =   47
            Top             =   3840
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1508
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid FlexGridDestino 
            Height          =   855
            Left            =   8340
            TabIndex        =   48
            Top             =   3840
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1508
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Destino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   9060
            TabIndex        =   50
            Top             =   3540
            Width           =   1965
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Taxa de Origem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   180
            TabIndex        =   49
            Top             =   3540
            Width           =   1890
         End
      End
      Begin VB.CommandButton CmdIniciarDigitacao 
         Caption         =   "Iniciar Digitação"
         Height          =   375
         Left            =   -74820
         TabIndex        =   14
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton CmdZerarDigitacao 
         Caption         =   "Zerar Digitação"
         Height          =   375
         Left            =   -73140
         TabIndex        =   15
         Top             =   1020
         Width           =   1575
      End
      Begin VB.OptionButton OptNavVertical 
         Caption         =   "Navegação Vertical"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -69360
         TabIndex        =   17
         Top             =   1110
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton OptNavLateral 
         Caption         =   "Navegação Lateral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71280
         TabIndex        =   16
         Top             =   1110
         Width           =   1695
      End
      Begin VB.Frame FraCiaAerea 
         Caption         =   "Esta será uma Tabela de qual Cia. Aérea?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   6540
         TabIndex        =   31
         Top             =   3540
         Width           =   4875
         Begin VB.TextBox TxtSiglaCiaAerea 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TxtNomeCiaAerea 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   2280
            Width           =   3315
         End
         Begin MSDataGridLib.DataGrid GridCiaAerea 
            Bindings        =   "frmCadTabPrecoINCLUSAO.frx":0060
            Height          =   1635
            Left            =   120
            TabIndex        =   10
            Top             =   300
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   2884
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
            DataMember      =   "Sel_CiaAerea"
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "codcia"
               Caption         =   "Cod. Cia."
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
               DataField       =   "fantasia"
               Caption         =   "Fantasia"
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
               DataField       =   "descricao"
               Caption         =   "Descricao"
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
               DataField       =   "estoqueminimo"
               Caption         =   "estoqueminimo"
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
               DataField       =   "estoqueatual"
               Caption         =   "estoqueatual"
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
               DataField       =   "proximonum"
               Caption         =   "proximonum"
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
               DataField       =   "avisominimo"
               Caption         =   "avisominimo"
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
               DataField       =   "datacadastro"
               Caption         =   "datacadastro"
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
               BeginProperty Column00 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   3000,189
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1305,071
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   929,764
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sigla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   345
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cia. Aérea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   34
            Top             =   2040
            Width           =   735
         End
      End
      Begin VB.Frame FraTipoTabela 
         Caption         =   "Tipo da Tabela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   6540
         TabIndex        =   25
         Top             =   960
         Width           =   4875
         Begin VB.OptionButton OptTabelaOficial 
            Caption         =   "Tabela Oficial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   1995
         End
         Begin VB.OptionButton OptTabelaEspecifica 
            Caption         =   "Específica para Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1995
         End
         Begin VB.Frame FraDadosCliente 
            Caption         =   "Dados do Cliente"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            TabIndex        =   26
            Top             =   900
            Width           =   4635
            Begin VB.TextBox TxtCGCCliente 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2100
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   1065
               Width           =   2415
            End
            Begin VB.TextBox TxtNomeCliente 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   465
               Width           =   4395
            End
            Begin VB.CommandButton CmdBuscaCliente 
               Caption         =   "Buscar Cliente"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   9
               Top             =   1020
               Width           =   1875
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CGC do Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3435
               TabIndex        =   30
               Top             =   840
               Width           =   1080
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Nome do Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3345
               TabIndex        =   29
               Top             =   240
               Width           =   1170
            End
         End
      End
      Begin VB.Frame FraLocalidades 
         Caption         =   "Localidades de Atendimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   180
         TabIndex        =   24
         Top             =   1680
         Width           =   6255
         Begin VB.CommandButton CmdRemoveLocalidade 
            Caption         =   "REMOVER"
            Enabled         =   0   'False
            Height          =   1635
            Left            =   2940
            TabIndex        =   6
            Top             =   2820
            Width           =   315
         End
         Begin VB.ListBox ListLocalidadesDisponives 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3570
            Left            =   180
            MultiSelect     =   2  'Extended
            TabIndex        =   3
            Top             =   900
            Width           =   2715
         End
         Begin VB.ListBox ListLocalidadesSel 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3570
            Left            =   3300
            MultiSelect     =   2  'Extended
            TabIndex        =   5
            Top             =   900
            Width           =   2775
         End
         Begin VB.CommandButton CmdTodasLocalidades 
            Caption         =   "Todas as Localidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   1
            Top             =   300
            Width           =   2895
         End
         Begin VB.CommandButton CmdLocalidades 
            Caption         =   "Cadastro de Localidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3180
            TabIndex        =   2
            Top             =   300
            Width           =   2895
         End
         Begin VB.CommandButton CmdAdicionaLocalidade 
            Caption         =   "ADI C IONAR"
            Enabled         =   0   'False
            Height          =   1935
            Left            =   2940
            TabIndex        =   4
            Top             =   900
            Width           =   315
         End
      End
      Begin VB.Frame FraOrigem 
         Caption         =   "Selecione a Origem Desta Tabela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         TabIndex        =   23
         Top             =   960
         Width           =   6255
         Begin VB.ComboBox ComboOrigem 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   6030
         End
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Fase 4: Visualização, Confirmação e Cadastramento da Nova Tabela"
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
         Left            =   -74700
         TabIndex        =   41
         Top             =   540
         Width           =   7185
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -63650
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -63650
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Fase 3: Digitando os Valores da Nova Tabela"
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
         Left            =   -74700
         TabIndex        =   38
         Top             =   540
         Width           =   4785
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fase 1: Definições Gerais"
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
         Left            =   300
         TabIndex        =   36
         Top             =   540
         Width           =   2700
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   180
         X2              =   11350
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label LblOrigem 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   -63750
         TabIndex        =   22
         Top             =   1080
         Width           =   75
      End
   End
   Begin VB.Label LblTransferencia 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   8160
      Visible         =   0   'False
      Width           =   11415
   End
End
Attribute VB_Name = "frmCadTabPrecoINCLUSAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xPesoParte, xPesoTodo As String
Public xFaixaParte, xFaixaTodo, xCategoriaParte, xCategoriaTodo, xLocalidadeParte, xLocalidadeTodo As String

Private Sub CmdAdicionaLocalidade_Click()
CmdAdicionaLocalidade.Enabled = False
Call TransfereItemDeListBox(ListLocalidadesDisponives, ListLocalidadesSel)
Call OrdenaListBox(ListLocalidadesSel)
End Sub


Private Sub CmdBuscaCliente_Click()
Set xForm = Me
frm_busca_cliente.Show 1
End Sub

Private Sub CmdCadastrarTabela_Click()
Dim X, Y, xCont As Integer
Dim xCodTab, xCodItemGeral, xCodItemTETC As Integer
Dim xDescricao, xTipoTab, xStatus, xUsarGeral As String

    If MskVigencia.Text = "" Then
    MsgBox "Você deve informar quando esta Tabela entrará em vigor. Por favor, tente novamente...", vbInformation, ""
    Exit Sub
    End If
    
CmdCadastrarTabela.Enabled = False
CmdCancelarTodoProcesso.Enabled = False
CmdFaseAnterior.Enabled = False
frmCadTabPrecoINCLUSAO.MousePointer = 11
    de_informa.cn_informa.BeginTrans
    
    If de_informa.rsSel_CodCadTabPrecoEscopo.State = 1 Then de_informa.rsSel_CodCadTabPrecoEscopo.Close
    de_informa.Sel_CodCadTabPrecoEscopo
    
        If de_informa.rsSel_CodCadTabPrecoEscopo.RecordCount = 0 Then
        xCodTab = "1000"
        Else
        xCodTab = Val(de_informa.rsSel_CodCadTabPrecoEscopo.Fields("codtab")) + 1
        End If
        
        xDescricao = TxtDescrSistema.Text
        
        If OptTabelaOficial.Value = True Then
        xTipoTab = "OFICIAL"
        Else
        xTipoTab = "ESPECIFICA"
        End If
        
        If CDate(MskVigencia.Text) > Date Then
        xStatus = "AGUARDANDO"
        Else
        xStatus = "VIGORANDO"
        End If
        
        If xTipoTab = "OFICIAL" Then
            If JaExisteTabelaOficial(Trim(UCase(TxtSiglaCiaAerea.Text)), Trim(UCase(ComboOrigem.Text))) > 0 Then
                If MsgBox("Já existe uma tabela Oficial da Cia. " & Trim(UCase(TxtSiglaCiaAerea.Text)) & " e que tem sua origem em " & UCase(Mid(Trim(ComboOrigem.Text), 1, Len(Trim(ComboOrigem.Text)) - 6)) & ". Clique em SIM para subtituir a Tabela Oficial existente ou em NÃO para cancelar o processo.", vbYesNo + vbExclamation, "") = vbYes Then
                de_informa.Update_FimVigenciaTabelas CDate(MskVigencia.Text), JaExisteTabelaOficial(Trim(UCase(TxtSiglaCiaAerea.Text)), Trim(UCase(ComboOrigem.Text)))
                Else
                Exit Sub
                End If
            End If
        ElseIf xTipoTab = "ESPECIFICA" Then
            If JaExisteTabelaEspecifica(Trim(UCase(TxtSiglaCiaAerea.Text)), Trim(UCase(ComboOrigem.Text)), Trim(TxtCGCCliente.Text) & "%") > 0 Then
                If MsgBox("Já existe uma tabela da Cia. " & Trim(UCase(TxtSiglaCiaAerea.Text)) & " Específica para o Cliente " & PriMaiuscula(Trim(TxtNomeCliente.Text)) & " e que tem sua origem em " & UCase(Mid(Trim(ComboOrigem.Text), 1, Len(Trim(ComboOrigem.Text)) - 6)) & ". Clique em SIM para subtituir este Tabela Específica ou em NÃO para cancelar o processo.", vbYesNo + vbExclamation, "") = vbYes Then
                de_informa.Update_FimVigenciaTabelas CDate(MskVigencia.Text), JaExisteTabelaEspecifica(Trim(UCase(TxtSiglaCiaAerea.Text)), Trim(UCase(ComboOrigem.Text)), Trim(TxtCGCCliente.Text) & "%")
                Else
                Exit Sub
                End If
            End If
        End If
    
    de_informa.Ins_CadTabPrecoEscopo xCodTab, Trim(UCase(TxtSiglaCiaAerea.Text)), xDescricao, Trim(UCase(ComboOrigem.Text)), CDate(MskVigencia.Text), CDate("01/01/1900"), xTipoTab, Trim(TxtNomeCliente.Text), Trim(TxtCGCCliente.Text), Trim(TxtOBS.Text), xStatus, DataHora("DATA"), xUsuario, FlexGridOrigem2.TextMatrix(0, 1), FlexGridOrigem2.TextMatrix(1, 1), FlexGridOrigem2.TextMatrix(2, 1), FlexGridDestino2.TextMatrix(0, 1), FlexGridDestino2.TextMatrix(1, 1), FlexGridDestino2.TextMatrix(2, 1)
        
        For Y = 1 To FlexGridNovaTabela.Rows - 1
            If de_informa.rsSel_CodCadTabPrecoGeral.State = 1 Then de_informa.rsSel_CodCadTabPrecoGeral.Close
            de_informa.Sel_CodCadTabPrecoGeral
            
            If de_informa.rsSel_CodCadTabPrecoGeral.RecordCount = 0 Then
            xCodItemGeral = "1000"
            Else
            xCodItemGeral = Val(de_informa.rsSel_CodCadTabPrecoGeral.Fields("coditem")) + 1
            End If
        
        de_informa.Ins_CadTabPrecogeral xCodItemGeral, xCodTab, UCase(Trim(FlexGridNovaTabela.TextMatrix(Y, 0))), FlexGridNovaTabela.TextMatrix(Y, 1), FlexGridNovaTabela.TextMatrix(Y, 2), FlexGridNovaTabela.TextMatrix(Y, 3), FlexGridNovaTabela.TextMatrix(Y, 4), FlexGridNovaTabela.TextMatrix(Y, 5), FlexGridNovaTabela.TextMatrix(Y, 6), FlexGridNovaTabela.TextMatrix(Y, 7), FlexGridNovaTabela.TextMatrix(Y, (FlexGridNovaTabela.Cols - 1) - 4), FlexGridNovaTabela.TextMatrix(Y, (FlexGridNovaTabela.Cols - 1) - 3), (Val(SemPonto(FlexGridNovaTabela.TextMatrix(Y, 8))) / 100), Trim(FlexGridNovaTabela.TextMatrix(Y, (FlexGridNovaTabela.Cols - 1) - 0))
        Next
        
        For Y = 1 To FlexGridNovaTabela.Rows - 1
            For X = 1 To de_informa.rsSel_CadIATA.RecordCount * 2
                If de_informa.rsSel_CodCadTabPrecoTETC.State = 1 Then de_informa.rsSel_CodCadTabPrecoTETC.Close
                de_informa.Sel_CodCadTabPrecoTETC
                
                If de_informa.rsSel_CodCadTabPrecoTETC.RecordCount = 0 Then
                xCodItemTETC = "1000"
                Else
                xCodItemTETC = Val(de_informa.rsSel_CodCadTabPrecoTETC.Fields("coditem")) + 1
                End If
                
                If CDbl(FlexGridNovaTabela.TextMatrix(Y, (8 + X))) = 0 Then
                xUsarGeral = "S"
                Else
                xUsarGeral = ""
                End If
                
                de_informa.Ins_CadTabPrecoTETC xCodItemTETC, xCodTab, UCase(FlexGridNovaTabela.TextMatrix(Y, 0)), Trim(Mid(FlexGridNovaTabela.TextMatrix(0, (8 + X)), Len(FlexGridNovaTabela.TextMatrix(0, (8 + X))) - 3)), CDbl(Mid(FlexGridNovaTabela.TextMatrix(Y, (8 + X)), Len(FlexGridNovaTabela.TextMatrix(Y, (8 + X))) - 3)), xUsarGeral, CDbl(FlexGridNovaTabela.TextMatrix(Y, (8 + X + 1)))
                X = X + 1
            Next
        Next
    de_informa.cn_informa.CommitTrans

Call AtualizaStatusTabelas

frmCadTabPrecoINCLUSAO.MousePointer = 0
CmdCadastrarTabela.Enabled = True
MsgBox "Sua Tabela foi Cadastrada com sucesso. Automaticamente entrará nos cálculos do Sistema a partir da Data Informada", vbInformation, ""
mdiAereo.mnuArquivo.Enabled = True
mdiAereo.mnuCadastros.Enabled = True
mdiAereo.mnuEmissoes.Enabled = True
mdiAereo.mnuRelat.Enabled = True
mdiAereo.mnuSair.Enabled = True
Unload Me
End Sub

Private Sub CmdCancelarTodoProcesso_Click()

    If MsgBox("Você tem certeza de que quer Cancelar o Cadastramento? (Todos os Dados serão Perdidos!)", vbYesNo + vbQuestion, "") = vbYes Then
    mdiAereo.mnuArquivo.Enabled = True
    mdiAereo.mnuCadastros.Enabled = True
    mdiAereo.mnuEmissoes.Enabled = True
    mdiAereo.mnuRelat.Enabled = True
    mdiAereo.mnuSair.Enabled = True
    Unload Me
    End If
    
End Sub

Private Sub CmdDefineFaixaPesoAnterior_Click()

    If Val(xPesoParte) <= 1 Then
    Exit Sub
    End If

xPesoTodo = Trim(Str(Val(TxtQteFaixasPeso.Text)))
xPesoParte = Trim(Str(Val(xPesoParte) - 1))

If Len(Trim(xPesoTodo)) = 1 Then xPesoTodo = "0" & xPesoTodo
If Len(Trim(xPesoParte)) = 1 Then xPesoParte = "0" & xPesoParte

LblDefineFaixasPeso.Caption = "Faixa " & xPesoParte & " de " & xPesoTodo
DoEvents
    
FlexGridFaixasPeso.Row = Val(xPesoParte)
FlexGridFaixasPeso.Col = 1
TxtPesoInicial.Text = FlexGridFaixasPeso.Text
FlexGridFaixasPeso.Col = 2
TxtPesoFinal.Text = FlexGridFaixasPeso.Text
DoEvents
End Sub


Private Sub CmdFaseAnterior_Click()

Dim xCont As Integer

    If TabFase.Tab = 1 Then
        If MsgBox("ATENÇÃO! Ao retornar para a Fase Anterior, todo seu trabalho será perdido! Você tem certeza que deseja retornar?", vbYesNo + vbCritical, "ATENÇÃO! Confirmação para Zerar Tabela") = vbNo Then
        Exit Sub
        Else
        CmdIniciarDigitacao.Enabled = True
        FraPanoramaNovaTabela.Enabled = False
        End If
    End If
    
CmdProximaFase.Enabled = False
CmdCancelarTodoProcesso.Enabled = False
CmdFaseAnterior.Enabled = False

    If TabFase.Tab - 1 >= 0 Then
    TabFase.Tab = TabFase.Tab - 1
    TabFase.TabEnabled(TabFase.Tab) = True
        For xCont = 0 To TabFase.Tab - 1
        TabFase.TabEnabled(xCont) = False
        Next
        
        For xCont = TabFase.Tab + 1 To TabFase.Tabs - 1
        TabFase.TabEnabled(xCont) = False
        Next
    End If
    
CmdCancelarTodoProcesso.Enabled = True
        
        If TabFase.Tab = TabFase.Tabs - 1 Then
        CmdProximaFase.Enabled = False
        CmdFaseAnterior.Enabled = True
        ElseIf TabFase.Tab = 0 Then
        CmdProximaFase.Enabled = True
        CmdFaseAnterior.Enabled = False
        ElseIf TabFase.Tab = 3 Then
        CmdProximaFase.Enabled = False
        CmdFaseAnterior.Enabled = True
        Else
        CmdProximaFase.Enabled = True
        CmdFaseAnterior.Enabled = True
        End If

End Sub

Private Sub CmdGerarFaixasPeso_Click()

Dim xCont As Integer

    If Val(TxtQteFaixasPeso.Text) = 0 Then
    MsgBox "Não existe uma valor para a criação da Faixas de Peso. Por Favor tente novamente...", vbExclamation, ""
    TxtQteFaixasPeso.SetFocus
    Exit Sub
    ElseIf Val(TxtQteFaixasPeso.Text) = 1 Then
    MsgBox "O menor valor possível é 2. Por Favor tente novamente...", vbExclamation, ""
    TxtQteFaixasPeso.SetFocus
    Exit Sub
    End If
    
FraDefineFaixasPeso.Enabled = True
TxtPesoInicial.BackColor = xAmarelo
TxtPesoFinal.BackColor = xAmarelo
TxtPesoFinal.Enabled = True
LblDefineFaixasPeso.Enabled = True
CmdDefineFaixaPesoAnterior.Enabled = True
CmdDefineProximaFaixaPeso.Enabled = True

FlexGridFaixasPeso.Clear
FlexGridFaixasPeso.Cols = 3
FlexGridFaixasPeso.Rows = Val(TxtQteFaixasPeso.Text) + 1

    For xCont = 0 To Val(TxtQteFaixasPeso.Text)
        If xCont = 0 Then
        FlexGridFaixasPeso.Row = xCont
        FlexGridFaixasPeso.Col = 0
        FlexGridFaixasPeso.Text = "Faixa"
        FlexGridFaixasPeso.CellFontBold = True
        FlexGridFaixasPeso.Col = 1
        FlexGridFaixasPeso.Text = "Corte Inicial"
        FlexGridFaixasPeso.CellFontBold = True
        FlexGridFaixasPeso.Col = 2
        FlexGridFaixasPeso.Text = "Corte Final"
        FlexGridFaixasPeso.CellFontBold = True
        Else
            If xCont = Val(TxtQteFaixasPeso.Text) Then
            FlexGridFaixasPeso.Row = xCont
            FlexGridFaixasPeso.Col = 0
            FlexGridFaixasPeso.Text = "Acima de"
            FlexGridFaixasPeso.CellFontBold = True
            Else
            FlexGridFaixasPeso.Row = xCont
            FlexGridFaixasPeso.Col = 0
            FlexGridFaixasPeso.Text = Trim(Str(xCont)) & "ª"
            FlexGridFaixasPeso.CellFontBold = True
            End If
        End If
    Next

FlexGridFaixasPeso.ColWidth(1) = 1500
FlexGridFaixasPeso.ColWidth(2) = 1500

TxtPesoInicial.Text = "0.0"

TxtPesoFinal.SetFocus

xPesoTodo = Trim(Str(Val(TxtQteFaixasPeso.Text)))
xPesoParte = "1"

If Len(Trim(xPesoTodo)) = 1 Then xPesoTodo = "0" & xPesoTodo
If Len(Trim(xPesoParte)) = 1 Then xPesoParte = "0" & xPesoParte

LblDefineFaixasPeso.Caption = "Faixa " & xPesoParte & " de " & xPesoTodo
DoEvents

End Sub

Private Sub CmdIniciarDigitacao_Click()
Dim xCont As Integer
CmdIniciarDigitacao.Enabled = False
FraPanoramaNovaTabela.Enabled = True

FlexGridPanoramaNovaTabela.Row = 1
FlexGridPanoramaNovaTabela.Col = 1
FlexGridPanoramaNovaTabela.SetFocus

End Sub

Private Sub CmdLocalidades_Click()

Dim X, Y As Integer

    frmCadLocalidade.Show 1
DoEvents
ListLocalidadesDisponives.Clear
ComboOrigem.Clear

If de_informa.rsSel_CadLocalAirGROUP.State = 1 Then de_informa.rsSel_CadLocalAirGROUP.Close
de_informa.Sel_CadLocalAirgroup

    Do Until de_informa.rsSel_CadLocalAirGROUP.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAirGROUP.Fields("localidade")) '& " - " & de_informa.rsSel_CadLocalAirGROUP.Fields("SIGLA")
    ComboOrigem.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAirGROUP.Fields("localidade")) '& " - " & de_informa.rsSel_CadLocalAirGROUP.Fields("SIGLA")
    de_informa.rsSel_CadLocalAirGROUP.MoveNext
    Loop
    
    'INICIO DO TRECHO QUE AVERIGUA LIST BOX
    For Y = 0 To ListLocalidadesDisponives.ListCount - 1
    ListLocalidadesDisponives.Selected(Y) = False
    Next
    For X = 0 To ListLocalidadesSel.ListCount - 1
        For Y = 0 To ListLocalidadesDisponives.ListCount - 1
            If ListLocalidadesDisponives.List(Y) = ListLocalidadesSel.List(X) Then
            ListLocalidadesDisponives.Selected(Y) = True
            End If
        Next

    Next
    
    Y = 0
    Do While True
        If Y > ListLocalidadesDisponives.ListCount - 1 Then
        Exit Do
        End If
        If ListLocalidadesDisponives.Selected(Y) = True Then
        ListLocalidadesDisponives.RemoveItem (Y)
            If Y > 0 Then
            Y = Y - 1
            Else
            Y = 0
            End If
        Else
        Y = Y + 1
        End If
    Loop
    'FIM DO TRECHO QUE AVERIGUA LIST BOX

DoEvents
End Sub

Private Sub CmdProximaFaixaPeso_Click()
    If IsNumeric(TxtValorOficial.Text) = False Or IsNumeric(TxtDescINTEC.Text) = False Or IsNumeric(TxtCharter.Text) = False Then
    MsgBox "Para avançar, você precisa informar todos os campos corretamente.", vbInformation, ""
    TxtValorOficial.SetFocus
    Exit Sub
    ElseIf CDbl(TxtValorOficial.Text) = 0 Then
    MsgBox "Não é possível ter um Valor Oficial nulo.", vbInformation, ""
    TxtValorOficial.SetFocus
    Exit Sub
    End If
    
    FlexGridPanoramaNovaTabela.TextMatrix(Val(xLocalidadeParte) + 2, (Val(xFaixaParte) * 3) - 2) = TxtValorOficial.Text
    FlexGridPanoramaNovaTabela.TextMatrix(Val(xLocalidadeParte) + 2, (Val(xFaixaParte) * 3) - 1) = TxtCharter.Text
    FlexGridPanoramaNovaTabela.TextMatrix(Val(xLocalidadeParte) + 2, (Val(xFaixaParte) * 3) - 0) = TxtDescINTEC.Text
    
    xFaixaParte = Trim(Str(Val(xFaixaParte) + 1))
    
    If Len(Trim(xFaixaParte)) = 1 Then xFaixaParte = "0" & xFaixaParte
    
    LblTabelaGeral.Caption = "Faixa " & xFaixaParte & " de " & xFaixaTodo
    
    TxtValorOficial.Text = ""
    TxtCharter.Text = ""
    TxtDescINTEC.Text = ""
    
    TxtValorOficial.SetFocus
    
End Sub

Private Sub CmdPlanilha_Click()
FlexGridImportacao.Rows = 0
FrmCadTabPrecoImportacao.Show 1
    If FlexGridImportacao.Rows > 0 Then
        For X = 0 To ListLocalidadesDisponives.ListCount - 1
        ListLocalidadesDisponives.Selected(X) = False
        Next
    
        For Y = 1 To FlexGridImportacao.Rows - 1
            For X = 0 To ListLocalidadesDisponives.ListCount - 1
                If FlexGridImportacao.TextMatrix(Y, 0) = ListLocalidadesDisponives.List(X) Then
                ListLocalidadesDisponives.Selected(X) = True
                End If
            Next
        Next
    Call CmdAdicionaLocalidade_Click
    End If
DoEvents
End Sub

Private Sub CmdProximaFase_Click()

Dim X, Y, xCont As Integer

    If TabFase.Tab = 0 Then
        If ListLocalidadesSel.ListCount = 0 Then
        MsgBox "Você não escolheu Região alguma para atendimento. Por favor, tente novamente...", vbInformation, ""
        Exit Sub
        ElseIf OptTabelaEspecifica.Value = True And Len(Trim(TxtNomeCliente.Text)) = 0 Then
        MsgBox "Você informou que esta tabela seria espeífica para uma cliente porém não informou qual cliente. Por favor, tente novamente...", vbInformation, ""
        Exit Sub
        ElseIf OptTabelaOficial.Value = False And OptTabelaEspecifica.Value = False Then
        MsgBox "Você não informou se esta tabela seria Oficial ou Específica. Por favor, tente novamente...", vbInformation, ""
        Exit Sub
        ElseIf Len(Trim(TxtSiglaCiaAerea.Text)) = 0 Then
        MsgBox "Você não informou de qual Cia. Aérea esta tabela será. Por favor, tente novamente...", vbInformation, ""
        Exit Sub
        ElseIf Len(Trim(ComboOrigem.Text)) = 0 Then
        MsgBox "Você não informou a Origem de sua Tabela. Por favor, tente novamente...", vbInformation, ""
        Exit Sub
        End If
    ElseIf TabFase.Tab = 1 Then
        If Trim(FlexGridOrigem.TextMatrix(0, 1)) = "" Or Trim(FlexGridOrigem.TextMatrix(1, 1)) = "" Or Trim(FlexGridOrigem.TextMatrix(2, 1)) = "" Then
        MsgBox "Você não informou as Taxas de Origem corretamente. Por favor, tente novamente...", vbInformation, ""
        Exit Sub
        ElseIf Trim(FlexGridDestino.TextMatrix(0, 1)) = "" Or Trim(FlexGridDestino.TextMatrix(1, 1)) = "" Or Trim(FlexGridDestino.TextMatrix(2, 1)) = "" Then
        MsgBox "Você não informou as Taxas de Destino corretamente. Por favor, tente novamente...", vbInformation, ""
        Exit Sub
        End If
        
        For Y = 1 To FlexGridPanoramaNovaTabela.Rows - 1
        FlexGridPanoramaNovaTabela.Row = Y
            For X = 1 To FlexGridPanoramaNovaTabela.Cols - 2
            FlexGridPanoramaNovaTabela.Col = X
                If FlexGridPanoramaNovaTabela.Text = "" Then
                FlexGridPanoramaNovaTabela.Text = "0,00"
                X = X - 1
                'ElseIf CDbl(Val(SemPonto(FlexGridPanoramaNovaTabela.Text)) / 100) = 0 And X >= 1 And X <= 8 Then
                'MsgBox "Nenhum valor entre as Colunas 2 e 9 podem ser nulos. Para continuar, corrija este problema.", vbInformation, ""
                '    If FraPanoramaNovaTabela.Enabled = False Then Exit Sub
               '
               ' FlexGridPanoramaNovaTabela.SetFocus
               '     If X = FlexGridPanoramaNovaTabela.Cols - 1 Then
               '     SendKeys "{LEFT}"
               '     SendKeys "{RIGHT}"
               '     Else
               '     SendKeys "{RIGHT}"
               '     SendKeys "{LEFT}"
               '     End If
               '     If X = FlexGridPanoramaNovaTabela.Rows - 1 Then
               '     SendKeys "{UP}"
               '     SendKeys "{DOWN}"
               '     Else
               '     SendKeys "{DOWN}"
               '     SendKeys "{UP}"
               '     End If
               ' Exit Sub
                ElseIf CDbl(Val(SemPonto(FlexGridPanoramaNovaTabela.Text)) / 100) = 0 And FlexGridPanoramaNovaTabela.TextMatrix(0, X) = "Corte Charter" And CDbl(SemPonto(Val(FlexGridPanoramaNovaTabela.TextMatrix(Y, FlexGridPanoramaNovaTabela.Cols - 1))) / 100) > 0 Then
                MsgBox "Você informou um Valor Charter porém não informou qual será o corte de peso. Para continuar, corrija este problema.", vbInformation, ""
                FlexGridPanoramaNovaTabela.SetFocus
                    If X = FlexGridPanoramaNovaTabela.Cols - 1 Then
                    SendKeys "{LEFT}"
                    SendKeys "{RIGHT}"
                    Else
                    SendKeys "{RIGHT}"
                    SendKeys "{LEFT}"
                    End If
                    If X = FlexGridPanoramaNovaTabela.Rows - 1 Then
                    SendKeys "{UP}"
                    SendKeys "{DOWN}"
                    Else
                    SendKeys "{DOWN}"
                    SendKeys "{UP}"
                    End If
                Exit Sub
                End If
            Next
        Next
    End If
    
CmdProximaFase.Enabled = False
CmdCancelarTodoProcesso.Enabled = False
CmdFaseAnterior.Enabled = False

    If TabFase.Tab + 1 <= TabFase.Tabs - 1 Then
    TabFase.Tab = TabFase.Tab + 1
    TabFase.TabEnabled(TabFase.Tab) = True
        For xCont = 0 To TabFase.Tab - 1
        TabFase.TabEnabled(xCont) = False
        Next
        
        For xCont = TabFase.Tab + 1 To TabFase.Tabs - 1
        TabFase.TabEnabled(xCont) = False
        Next
    End If
        
        
        If TabFase.Tab = 1 Then
        
        FlexGridOrigem.Clear
        FlexGridOrigem.Rows = 3
        FlexGridOrigem.Cols = 2
        FlexGridOrigem.FixedRows = 0
        FlexGridOrigem.FixedCols = 1
        FlexGridOrigem.TextMatrix(0, 0) = "Valor"
        FlexGridOrigem.TextMatrix(1, 0) = "Até"
        FlexGridOrigem.TextMatrix(2, 0) = "Kg Exced."
        FlexGridOrigem.ColWidth(0) = 1300
        FlexGridOrigem.ColWidth(1) = 1300
        
        FlexGridDestino.Clear
        FlexGridDestino.Rows = 3
        FlexGridDestino.Cols = 2
        FlexGridDestino.FixedRows = 0
        FlexGridDestino.FixedCols = 1
        FlexGridDestino.TextMatrix(0, 0) = "Valor"
        FlexGridDestino.TextMatrix(1, 0) = "Até"
        FlexGridDestino.TextMatrix(2, 0) = "Kg Exced."
        FlexGridDestino.ColWidth(0) = 1300
        FlexGridDestino.ColWidth(1) = 1300
        
        FlexGridPanoramaNovaTabela.Clear
        
        LblOrigem.Caption = "Origem: " & ComboOrigem.Text
        
        If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
        de_informa.Sel_Cadiata "%"
                
        FlexGridPanoramaNovaTabela.Cols = 12 + ((de_informa.rsSel_CadIATA.RecordCount - 1) * 2)
        FlexGridPanoramaNovaTabela.Rows = ListLocalidadesSel.ListCount + 1
        
        FlexGridPanoramaNovaTabela.FixedRows = 1
        FlexGridPanoramaNovaTabela.FixedCols = 1
            
        FlexGridPanoramaNovaTabela.TextMatrix(0, 0) = "Localidades"
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 0
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        '  FlexGridPanoramaNovaTabela.TextMatrix(0, 1) = "Sigla"
        'FlexGridPanoramaNovaTabela.Row = 0
        'FlexGridPanoramaNovaTabela.Col = 1
        'FlexGridPanoramaNovaTabela.CellAlignment = 3
        'FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, 1) = "Taxa Mínima"
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 1
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, 2) = "Até 25,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 3) = "Até 50,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 4) = "Até 300,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 5) = "Até 500,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 6) = "Até 1000,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 7) = "Acima de 1000,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 8) = "Desc. Tab. Geral (%)"
        
        FlexGridPanoramaNovaTabela.Col = 8
            For Y = 1 To ListLocalidadesSel.ListCount
            FlexGridPanoramaNovaTabela.Row = Y
            FlexGridPanoramaNovaTabela.CellBackColor = xLaranja
            Next
        
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 2
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 3
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 4
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 5
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 6
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 7
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 8
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.ColWidth(0) = 2100
        FlexGridPanoramaNovaTabela.ColWidth(1) = 1300
        FlexGridPanoramaNovaTabela.ColWidth(2) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(3) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(4) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(5) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(6) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(7) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(8) = 2000
        
                    
        xCont = 9
            Do Until de_informa.rsSel_CadIATA.EOF
                If de_informa.rsSel_CadIATA.Fields("codigo") <> "000" Then
                FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Cód. " & de_informa.rsSel_CadIATA.Fields("codigo")
                FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1000
                
                FlexGridPanoramaNovaTabela.Row = 0
                FlexGridPanoramaNovaTabela.Col = xCont
                FlexGridPanoramaNovaTabela.CellAlignment = 3
                FlexGridPanoramaNovaTabela.CellFontBold = True
                
                xCont = xCont + 1
                
                FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Desc. Cód. " & de_informa.rsSel_CadIATA.Fields("codigo") & " (%)"
                FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1800
                
                FlexGridPanoramaNovaTabela.Row = 0
                FlexGridPanoramaNovaTabela.Col = xCont
                FlexGridPanoramaNovaTabela.CellAlignment = 3
                FlexGridPanoramaNovaTabela.CellFontBold = True
                
                    For Y = 1 To ListLocalidadesSel.ListCount
                    FlexGridPanoramaNovaTabela.Row = Y
                    FlexGridPanoramaNovaTabela.CellBackColor = xAmarelo
                    Next
                
                
                xCont = xCont + 1
                End If
            de_informa.rsSel_CadIATA.MoveNext
            Loop
        
        FlexGridPanoramaNovaTabela.Row = 0
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Corte Charter (Kg)"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1600
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
            For Y = 1 To ListLocalidadesSel.ListCount
            FlexGridPanoramaNovaTabela.Row = Y
            FlexGridPanoramaNovaTabela.CellBackColor = xCinzaClaro
            Next
        
        xCont = xCont + 1
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Valor Charter"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
            For Y = 1 To ListLocalidadesSel.ListCount
            FlexGridPanoramaNovaTabela.Row = Y
            FlexGridPanoramaNovaTabela.CellBackColor = xCinzaClaro
            Next
            
            
        ''xCont = xCont + 1
        ''FlexGridPanoramaNovaTabela.Row = 0
        ''FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Tx. Terrestre"
        ''FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        ''FlexGridPanoramaNovaTabela.Col = xCont
        ''FlexGridPanoramaNovaTabela.CellAlignment = 3
        ''FlexGridPanoramaNovaTabela.CellFontBold = True
        
        ''    For Y = 1 To ListLocalidadesSel.ListCount
        ''    FlexGridPanoramaNovaTabela.Row = Y
        ''    FlexGridPanoramaNovaTabela.CellBackColor = xBranco
        ''    Next
        
        ''xCont = xCont + 1
        ''FlexGridPanoramaNovaTabela.Row = 0
        ''FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Até (Kg)"
        ''FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        ''FlexGridPanoramaNovaTabela.Col = xCont
        ''FlexGridPanoramaNovaTabela.CellAlignment = 3
        ''FlexGridPanoramaNovaTabela.CellFontBold = True
        
        ''    For Y = 1 To ListLocalidadesSel.ListCount
        ''    FlexGridPanoramaNovaTabela.Row = Y
        ''    FlexGridPanoramaNovaTabela.CellBackColor = xBranco
        ''    Next
        
        ''xCont = xCont + 1
        ''FlexGridPanoramaNovaTabela.Row = 0
        ''FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Vl. Kg. Ex."
        ''FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        ''FlexGridPanoramaNovaTabela.Col = xCont
        ''FlexGridPanoramaNovaTabela.CellAlignment = 3
        ''FlexGridPanoramaNovaTabela.CellFontBold = True
        
        xCont = xCont + 1
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Tx. Terrestre"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        ''    For Y = 1 To ListLocalidadesSel.ListCount
        ''    FlexGridPanoramaNovaTabela.Row = Y
        ''    FlexGridPanoramaNovaTabela.CellBackColor = xBranco
        ''    Next
            
            
            For xCont = 0 To ListLocalidadesSel.ListCount - 1
            FlexGridPanoramaNovaTabela.TextMatrix(xCont + 1, 0) = ListLocalidadesSel.List(xCont)
            FlexGridPanoramaNovaTabela.Row = xCont + 1
            FlexGridPanoramaNovaTabela.Col = 0
            FlexGridPanoramaNovaTabela.CellAlignment = 1
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            'FlexGridPanoramaNovaTabela.TextMatrix(xCont + 1, 1) = Mid(ListLocalidadesSel.List(xCont), Len(ListLocalidadesSel.List(xCont)) - 3)
            'FlexGridPanoramaNovaTabela.Row = xCont + 1
            'FlexGridPanoramaNovaTabela.Col = 1
            'FlexGridPanoramaNovaTabela.CellAlignment = 1
            'FlexGridPanoramaNovaTabela.CellFontBold = True
            
            Next
            
            If FlexGridImportacao.Rows > 0 Then
                For X = 1 To FlexGridPanoramaNovaTabela.Rows - 1
                    For X2 = 1 To FlexGridImportacao.Rows - 1
                        If FlexGridPanoramaNovaTabela.TextMatrix(X, 0) = FlexGridImportacao.TextMatrix(X2, 0) Then
                            For Y = 1 To FlexGridPanoramaNovaTabela.Cols - 1
                                For Y2 = 1 To FlexGridImportacao.Cols - 1
                                    If FlexGridPanoramaNovaTabela.TextMatrix(0, Y) = FlexGridImportacao.TextMatrix(0, Y2) Then
                                    FlexGridPanoramaNovaTabela.TextMatrix(X, Y) = FlexGridImportacao.TextMatrix(X2, Y2)
                                    End If
                                Next
                            Next
                        End If
                    Next
                Next
            End If
        
        ElseIf TabFase.Tab = 2 Then
        FlexGridNovaTabela.Visible = False
        FlexGridOrigem2.Visible = False
        FlexGridDestino2.Visible = False
        DoEvents
        
        FlexGridOrigem2.Clear
        
        FlexGridOrigem2.Rows = FlexGridOrigem.Rows
        FlexGridOrigem2.Cols = FlexGridOrigem.Cols
        
        FlexGridOrigem2.FixedRows = FlexGridOrigem.FixedRows
        FlexGridOrigem2.FixedCols = FlexGridOrigem.FixedCols
        
            For Y = 0 To FlexGridOrigem2.Rows - 1
                For X = 0 To FlexGridOrigem2.Cols - 1
                FlexGridOrigem.Row = Y
                FlexGridOrigem2.Row = Y
                FlexGridOrigem2.Col = X
                FlexGridOrigem.Col = X
                
                    If FlexGridOrigem.CellFontBold = True Then
                    FlexGridOrigem2.CellFontBold = True
                    Else
                    FlexGridOrigem2.CellFontBold = False
                    End If
                
                FlexGridOrigem2.CellAlignment = FlexGridOrigem.CellAlignment
                FlexGridOrigem2.CellBackColor = FlexGridOrigem.CellBackColor
                FlexGridOrigem2.Text = FlexGridOrigem.Text
                FlexGridOrigem2.ColWidth(X) = FlexGridOrigem.ColWidth(X)
                Next
            Next
        
        FlexGridDestino2.Clear
        
        FlexGridDestino2.Rows = FlexGridDestino.Rows
        FlexGridDestino2.Cols = FlexGridDestino.Cols
        
        FlexGridDestino2.FixedRows = FlexGridDestino.FixedRows
        FlexGridDestino2.FixedCols = FlexGridDestino.FixedCols
        
            For Y = 0 To FlexGridDestino2.Rows - 1
                For X = 0 To FlexGridDestino2.Cols - 1
                FlexGridDestino.Row = Y
                FlexGridDestino2.Row = Y
                FlexGridDestino2.Col = X
                FlexGridDestino.Col = X
                
                    If FlexGridDestino.CellFontBold = True Then
                    FlexGridDestino2.CellFontBold = True
                    Else
                    FlexGridDestino2.CellFontBold = False
                    End If
                
                FlexGridDestino2.CellAlignment = FlexGridDestino.CellAlignment
                FlexGridDestino2.CellBackColor = FlexGridDestino.CellBackColor
                FlexGridDestino2.Text = FlexGridDestino.Text
                FlexGridDestino2.ColWidth(X) = FlexGridDestino.ColWidth(X)
                Next
            Next
        
        
        FlexGridNovaTabela.Clear
        
        FlexGridNovaTabela.Rows = FlexGridPanoramaNovaTabela.Rows
        FlexGridNovaTabela.Cols = FlexGridPanoramaNovaTabela.Cols
        
        FlexGridNovaTabela.FixedRows = FlexGridPanoramaNovaTabela.FixedRows
        FlexGridNovaTabela.FixedCols = FlexGridPanoramaNovaTabela.FixedCols
        
            For Y = 0 To FlexGridNovaTabela.Rows - 1
                For X = 0 To FlexGridNovaTabela.Cols - 1
                FlexGridPanoramaNovaTabela.Row = Y
                FlexGridNovaTabela.Row = Y
                FlexGridNovaTabela.Col = X
                FlexGridPanoramaNovaTabela.Col = X
                
                    If FlexGridPanoramaNovaTabela.CellFontBold = True Then
                    FlexGridNovaTabela.CellFontBold = True
                    Else
                    FlexGridNovaTabela.CellFontBold = False
                    End If
                
                FlexGridNovaTabela.CellAlignment = FlexGridPanoramaNovaTabela.CellAlignment
                FlexGridNovaTabela.CellBackColor = FlexGridPanoramaNovaTabela.CellBackColor
                FlexGridNovaTabela.Text = FlexGridPanoramaNovaTabela.Text
                FlexGridNovaTabela.ColWidth(X) = FlexGridPanoramaNovaTabela.ColWidth(X)
                Next
            Next
        FlexGridNovaTabela.Visible = True
        FlexGridOrigem2.Visible = True
        FlexGridDestino2.Visible = True
        DoEvents
        
            If OptTabelaOficial.Value = True Then
            TxtDescrSistema.Text = PriMaiuscula(TxtNomeCiaAerea.Text) & " (" & Trim(TxtSiglaCiaAerea.Text) & ") - Oficial - Origem " & Trim(ComboOrigem.Text)
            Else
            TxtDescrSistema.Text = PriMaiuscula(TxtNomeCiaAerea.Text) & " (" & Trim(TxtSiglaCiaAerea.Text) & ") - Específica - " & PriMaiuscula(TxtNomeCliente.Text) & " - " & Trim(LblOrigem.Caption)
            End If
        End If

CmdCancelarTodoProcesso.Enabled = True

        If TabFase.Tab = TabFase.Tabs - 1 Then
        CmdProximaFase.Enabled = False
        CmdFaseAnterior.Enabled = True
        ElseIf TabFase.Tab = 0 Then
        CmdProximaFase.Enabled = True
        CmdFaseAnterior.Enabled = False
        ElseIf TabFase.Tab = 3 Then
        CmdProximaFase.Enabled = False
        CmdFaseAnterior.Enabled = True
        Else
        CmdProximaFase.Enabled = True
        CmdFaseAnterior.Enabled = True
        End If

End Sub

Private Sub CmdRemoveLocalidade_Click()
CmdRemoveLocalidade.Enabled = False
Call TransfereItemDeListBox(ListLocalidadesSel, ListLocalidadesDisponives)
Call OrdenaListBox(ListLocalidadesDisponives)

End Sub

Private Sub CmdTodasLocalidades_Click()

CmdTodasLocalidades.Enabled = False
Dim xCont As Integer

    For xCont = 0 To ListLocalidadesDisponives.ListCount - 1
    ListLocalidadesDisponives.Selected(xCont) = True
    Next
    
Call TransfereItemDeListBox(ListLocalidadesDisponives, ListLocalidadesSel)
Call OrdenaListBox(ListLocalidadesSel)

CmdAdicionaLocalidade.Enabled = False
CmdRemoveLocalidade.Enabled = False

CmdTodasLocalidades.Enabled = True
    
End Sub


Private Sub CmdZerarDigitacao_Click()
Dim X, Y, xCont As Integer

If MsgBox("Confirma Zerar Nova Tabela?", vbYesNo + vbExclamation, "Confirmação para Zerar Tabela") = vbYes Then
    If MsgBox("ATENÇÃO! Ao zerar a digitação, todo seu trabalho será perdido! Você tem certeza que deseja zerar a tabela?", vbYesNo + vbCritical, "ATENÇÃO! Confirmação para Zerar Tabela") = vbYes Then
                
        FlexGridOrigem.Clear
        FlexGridOrigem.Rows = 3
        FlexGridOrigem.Cols = 2
        FlexGridOrigem.FixedRows = 0
        FlexGridOrigem.FixedCols = 1
        FlexGridOrigem.TextMatrix(0, 0) = "Valor"
        FlexGridOrigem.TextMatrix(1, 0) = "Até"
        FlexGridOrigem.TextMatrix(2, 0) = "Kg Exced."
        FlexGridOrigem.ColWidth(0) = 1300
        FlexGridOrigem.ColWidth(1) = 1300
        
        FlexGridDestino.Clear
        FlexGridDestino.Rows = 3
        FlexGridDestino.Cols = 2
        FlexGridDestino.FixedRows = 0
        FlexGridDestino.FixedCols = 1
        FlexGridDestino.TextMatrix(0, 0) = "Valor"
        FlexGridDestino.TextMatrix(1, 0) = "Até"
        FlexGridDestino.TextMatrix(2, 0) = "Kg Exced."
        FlexGridDestino.ColWidth(0) = 1300
        FlexGridDestino.ColWidth(1) = 1300
        
        FlexGridPanoramaNovaTabela.Clear
        
        LblOrigem.Caption = "Origem: " & ComboOrigem.Text
        
        If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
        de_informa.Sel_Cadiata "%"
                
        FlexGridPanoramaNovaTabela.Cols = 12 + ((de_informa.rsSel_CadIATA.RecordCount - 1) * 2)
        FlexGridPanoramaNovaTabela.Rows = ListLocalidadesSel.ListCount + 1
        
        FlexGridPanoramaNovaTabela.FixedRows = 1
        FlexGridPanoramaNovaTabela.FixedCols = 1
            
        FlexGridPanoramaNovaTabela.TextMatrix(0, 0) = "Localidades"
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 0
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        '  FlexGridPanoramaNovaTabela.TextMatrix(0, 1) = "Sigla"
        'FlexGridPanoramaNovaTabela.Row = 0
        'FlexGridPanoramaNovaTabela.Col = 1
        'FlexGridPanoramaNovaTabela.CellAlignment = 3
        'FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, 1) = "Taxa Mínima"
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 1
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, 2) = "Até 25,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 3) = "Até 50,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 4) = "Até 300,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 5) = "Até 500,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 6) = "Até 1000,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 7) = "Acima de 1000,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 8) = "Desc. Tab. Geral (%)"
        
        FlexGridPanoramaNovaTabela.Col = 8
            For Y = 1 To ListLocalidadesSel.ListCount
            FlexGridPanoramaNovaTabela.Row = Y
            FlexGridPanoramaNovaTabela.CellBackColor = xLaranja
            Next
        
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 2
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 3
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 4
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 5
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 6
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 7
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        FlexGridPanoramaNovaTabela.Col = 8
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.ColWidth(0) = 2100
        FlexGridPanoramaNovaTabela.ColWidth(1) = 1300
        FlexGridPanoramaNovaTabela.ColWidth(2) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(3) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(4) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(5) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(6) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(7) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(8) = 2000
        
                    
        xCont = 9
            Do Until de_informa.rsSel_CadIATA.EOF
                        
            FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Cód. " & de_informa.rsSel_CadIATA.Fields("codigo")
            FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1000
            
            FlexGridPanoramaNovaTabela.Row = 0
            FlexGridPanoramaNovaTabela.Col = xCont
            FlexGridPanoramaNovaTabela.CellAlignment = 3
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            xCont = xCont + 1
            
            FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Desc. Cód. " & de_informa.rsSel_CadIATA.Fields("codigo") & " (%)"
            FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1800
            
            FlexGridPanoramaNovaTabela.Row = 0
            FlexGridPanoramaNovaTabela.Col = xCont
            FlexGridPanoramaNovaTabela.CellAlignment = 3
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
                For Y = 1 To ListLocalidadesSel.ListCount
                FlexGridPanoramaNovaTabela.Row = Y
                FlexGridPanoramaNovaTabela.CellBackColor = xAmarelo
                Next
            
            
            xCont = xCont + 1
            de_informa.rsSel_CadIATA.MoveNext
            Loop
        
        FlexGridPanoramaNovaTabela.Row = 0
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Corte Charter (Kg)"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1600
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
            For Y = 1 To ListLocalidadesSel.ListCount
            FlexGridPanoramaNovaTabela.Row = Y
            FlexGridPanoramaNovaTabela.CellBackColor = xCinzaClaro
            Next
        
        xCont = xCont + 1
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Valor Charter"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
            For Y = 1 To ListLocalidadesSel.ListCount
            FlexGridPanoramaNovaTabela.Row = Y
            FlexGridPanoramaNovaTabela.CellBackColor = xCinzaClaro
            Next
            
            
        ''xCont = xCont + 1
        ''FlexGridPanoramaNovaTabela.Row = 0
        ''FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Tx. Terrestre"
        ''FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        ''FlexGridPanoramaNovaTabela.Col = xCont
        ''FlexGridPanoramaNovaTabela.CellAlignment = 3
        ''FlexGridPanoramaNovaTabela.CellFontBold = True
        
        ''    For Y = 1 To ListLocalidadesSel.ListCount
        ''    FlexGridPanoramaNovaTabela.Row = Y
        ''    FlexGridPanoramaNovaTabela.CellBackColor = xBranco
        ''    Next
        
        ''xCont = xCont + 1
        ''FlexGridPanoramaNovaTabela.Row = 0
        ''FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Até (Kg)"
        ''FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        ''FlexGridPanoramaNovaTabela.Col = xCont
        ''FlexGridPanoramaNovaTabela.CellAlignment = 3
        ''FlexGridPanoramaNovaTabela.CellFontBold = True
        
        ''    For Y = 1 To ListLocalidadesSel.ListCount
        ''    FlexGridPanoramaNovaTabela.Row = Y
        ''    FlexGridPanoramaNovaTabela.CellBackColor = xBranco
        ''    Next
        
        ''xCont = xCont + 1
        ''FlexGridPanoramaNovaTabela.Row = 0
        ''FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Vl. Kg. Ex."
        ''FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        ''FlexGridPanoramaNovaTabela.Col = xCont
        ''FlexGridPanoramaNovaTabela.CellAlignment = 3
        ''FlexGridPanoramaNovaTabela.CellFontBold = True
        
        xCont = xCont + 1
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Tx. Terrestre"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        ''    For Y = 1 To ListLocalidadesSel.ListCount
        ''    FlexGridPanoramaNovaTabela.Row = Y
        ''    FlexGridPanoramaNovaTabela.CellBackColor = xBranco
        ''    Next
            
            
            For xCont = 0 To ListLocalidadesSel.ListCount - 1
            FlexGridPanoramaNovaTabela.TextMatrix(xCont + 1, 0) = ListLocalidadesSel.List(xCont)
            FlexGridPanoramaNovaTabela.Row = xCont + 1
            FlexGridPanoramaNovaTabela.Col = 0
            FlexGridPanoramaNovaTabela.CellAlignment = 1
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            'FlexGridPanoramaNovaTabela.TextMatrix(xCont + 1, 1) = Mid(ListLocalidadesSel.List(xCont), Len(ListLocalidadesSel.List(xCont)) - 3)
            'FlexGridPanoramaNovaTabela.Row = xCont + 1
            'FlexGridPanoramaNovaTabela.Col = 1
            'FlexGridPanoramaNovaTabela.CellAlignment = 1
            'FlexGridPanoramaNovaTabela.CellFontBold = True
            
            Next
            

    CmdIniciarDigitacao.Enabled = True
    FraPanoramaNovaTabela.Enabled = False
    End If
End If
            
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub FlexGridAjustes_KeyPress(KeyAscii As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer


    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
            If FlexGridAjustes.Col = FlexGridAjustes.Cols - 1 Then
                If FlexGridAjustes.Row <> FlexGridAjustes.Rows - 1 Then
                FlexGridAjustes.Row = FlexGridAjustes.Row + 1
                FlexGridAjustes.Col = 2
                SendKeys "{LEFT}"
                Else
                CmdProximaFase.SetFocus
                End If
            Else
            SendKeys "{RIGHT}"
            End If
        ElseIf KeyAscii = 8 Then
            FlexGridAjustes.Text = Mid(FlexGridAjustes.Text, 1, Len(FlexGridAjustes.Text) - 1)
            FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
                For Y = 1 To FlexGridAjustar.Rows - 1
                FlexGridAjustar.Col = FlexGridAjustes.Col
                FlexGridAjustar.Row = Y
                
                FlexGridEspelho.Col = FlexGridAjustes.Col
                FlexGridEspelho.Row = Y
                
                FlexGridAjustar.Text = Format(CDbl(FlexGridEspelho.Text) + (CDbl(FlexGridEspelho.Text) * (CDbl(FlexGridAjustes.TextMatrix(1, FlexGridAjustes.Col)) / 100)), "###,##0.00")
                Next
        ElseIf KeyAscii = 3 Then
        LblTransferencia.Caption = FlexGridAjustes.Text
        ElseIf KeyAscii = 22 Then
            For Y1 = FlexGridAjustes.Row To FlexGridAjustes.RowSel
                For X1 = FlexGridAjustes.Col To FlexGridAjustes.ColSel
                FlexGridAjustes.TextMatrix(Y1, X1) = LblTransferencia.Caption
                
                'AQUI AQUI AQUI AQUI AQUI
                If Mid(Trim(FlexGridAjustes.TextMatrix(0, X1)), 1, 4) <> "Desc" And (Trim(FlexGridAjustes.TextMatrix(0, X1)) <> "Corte Charter (Kg)" And Trim(FlexGridAjustes.TextMatrix(0, X1)) <> "Valor Charter" And Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) <> "Tx. Terrestre" And Trim(FlexGridAjustes.TextMatrix(0, X1)) <> "Até (Kg)" And Trim(FlexGridAjustes.TextMatrix(0, X1)) <> "Vl. Kg. Ex.") Then
                'FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
                FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
                
                    For Y = 1 To FlexGridAjustar.Rows - 1
                    FlexGridAjustar.Col = X1
                    FlexGridAjustar.Row = Y
                    
                    FlexGridEspelho.Col = X1
                    FlexGridEspelho.Row = Y
                    
                    FlexGridAjustar.Text = Format(CDbl(FlexGridEspelho.Text) + (CDbl(FlexGridEspelho.Text) * (CDbl(FlexGridAjustes.TextMatrix(1, FlexGridAjustes.Col)) / 100)), "###,##0.00")
                    Next
                ElseIf Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Corte Charter (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Valor Charter" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Tx. Terrestre" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Até (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Vl. Kg. Ex." Then
                FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
                FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
                
                    For Y = 1 To FlexGridAjustar.Rows - 1
                    FlexGridAjustar.Col = X1
                    FlexGridAjustar.Row = Y
                    
                    FlexGridEspelho.Col = X1
                    FlexGridEspelho.Row = Y
                    
                    If CDbl(FlexGridAjustar.Text) > 0 Then
                    FlexGridAjustar.Text = Format(FlexGridAjustes.Text, "###,##0.00")
                    End If
                    Next
                Else
                FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
                FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
                
                    For Y = 1 To FlexGridAjustar.Rows - 1
                    FlexGridAjustar.Col = X1
                    FlexGridAjustar.Row = Y
                    
                    FlexGridEspelho.Col = X1
                    FlexGridEspelho.Row = Y
                    
                    FlexGridAjustar.Text = Format(FlexGridAjustes.Text, "###,##0.00")
                    Next
                End If
                
                
                Next
            Next
        
        Else
        KeyAscii = 0
        End If
    Else
        If Mid(Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)), 1, 4) <> "Desc" And (Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) <> "Corte Charter (Kg)" And Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) <> "Valor Charter" And Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) <> "Tx. Terrestre" And Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) <> "Até (Kg)" And Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) <> "Vl. Kg. Ex.") Then
        FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
        FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
        
            For Y = 1 To FlexGridAjustar.Rows - 1
            FlexGridAjustar.Col = FlexGridAjustes.Col
            FlexGridAjustar.Row = Y
            
            FlexGridEspelho.Col = FlexGridAjustes.Col
            FlexGridEspelho.Row = Y
            
            FlexGridAjustar.Text = Format(CDbl(FlexGridEspelho.Text) + (CDbl(FlexGridEspelho.Text) * (CDbl(FlexGridAjustes.TextMatrix(1, FlexGridAjustes.Col)) / 100)), "###,##0.00")
            Next
        ElseIf Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Corte Charter (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Valor Charter" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Tx. Terrestre" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Até (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Vl. Kg. Ex." Then
        FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
        FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
        
            For Y = 1 To FlexGridAjustar.Rows - 1
            FlexGridAjustar.Col = FlexGridAjustes.Col
            FlexGridAjustar.Row = Y
            
            FlexGridEspelho.Col = FlexGridAjustes.Col
            FlexGridEspelho.Row = Y
            
            If CDbl(FlexGridAjustar.Text) > 0 Then
            FlexGridAjustar.Text = Format(FlexGridAjustes.Text, "###,##0.00")
            End If
            Next
        Else
        FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
        FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
        
            For Y = 1 To FlexGridAjustar.Rows - 1
            FlexGridAjustar.Col = FlexGridAjustes.Col
            FlexGridAjustar.Row = Y
            
            FlexGridEspelho.Col = FlexGridAjustes.Col
            FlexGridEspelho.Row = Y
            
            FlexGridAjustar.Text = Format(FlexGridAjustes.Text, "###,##0.00")
            Next
        End If
    End If

End Sub


Private Sub FlexGridAjustes_Scroll()
FlexGridAjustar.TopRow = FlexGridAjustes.TopRow
FlexGridAjustar.LeftCol = FlexGridAjustes.LeftCol

FlexGridEspelho.TopRow = FlexGridAjustes.TopRow
FlexGridEspelho.LeftCol = FlexGridAjustes.LeftCol
End Sub

Private Sub FlexGridAjustar_Scroll()

FlexGridAjustes.LeftCol = FlexGridAjustes.LeftCol

FlexGridEspelho.TopRow = FlexGridAjustar.TopRow
FlexGridEspelho.LeftCol = FlexGridAjustar.LeftCol
End Sub

Private Sub FlexGridEspelho_Scroll()

FlexGridAjustes.LeftCol = FlexGridEspelho.LeftCol

FlexGridAjustar.TopRow = FlexGridEspelho.TopRow
FlexGridAjustar.LeftCol = FlexGridEspelho.LeftCol
End Sub



Private Sub Comboorigem_KeyPress(KeyAscii As Integer)
Dim xTextoVelho As String, xTextoNovo As String, Y As Integer

'Asc (UCase(Chr(KeyAscii)))

    If KeyAscii <> 13 And KeyAscii <> 8 Then
    xTextoVelho = Left(ComboOrigem.Text, ComboOrigem.SelStart) & Chr(KeyAscii)
    xTextoNovo = ""
        For Y = 0 To ComboOrigem.ListCount - 1
            If Len(xTextoVelho) <= Len(ComboOrigem.List(Y)) Then
                If UCase(Mid(ComboOrigem.List(Y), 1, Len(xTextoVelho))) = UCase(xTextoVelho) Then
                xTextoNovo = Mid(ComboOrigem.List(Y), Len(xTextoVelho) + 1)
                Y = ComboOrigem.ListCount
                End If
            End If
        Next
    ComboOrigem.Text = UCase(xTextoVelho) & xTextoNovo
    ComboOrigem.SelStart = Len(xTextoVelho)
    ComboOrigem.SelLength = 1000
    ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    ElseIf KeyAscii = 8 Then
        If Len(ComboOrigem.Text) > 0 Then
            If ComboOrigem.SelStart > 0 Then
            xTextoVelho = Mid(ComboOrigem.Text, 1, ComboOrigem.SelStart - 1)
            Else
            xTextoVelho = Mid(ComboOrigem.Text, 1, ComboOrigem.SelStart)
            End If
        ComboOrigem.Text = UCase(xTextoVelho)
        ComboOrigem.SelStart = Len(xTextoVelho)
        ComboOrigem.SelLength = 1000
        End If
    End If
KeyAscii = 0
End Sub


Private Sub Comboorigem_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 100
End Sub

Private Sub Comboorigem_LostFocus()
Dim Y As Integer, xTexto As String

xTexto = ""

        For Y = 0 To ComboOrigem.ListCount - 1
            If UCase(Trim((ComboOrigem.Text))) = UCase(Trim(ComboOrigem.List(Y))) Then
            xTexto = ComboOrigem.List(Y)
            Y = ComboOrigem.ListCount
            End If
        Next
ComboOrigem.Text = xTexto
End Sub

Private Sub FlexGriddestino_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer
        If KeyCode = 46 Then
            For Y1 = FlexGridDestino.Row To FlexGridDestino.RowSel
                For X1 = FlexGridDestino.Col To FlexGridDestino.ColSel
                FlexGridDestino.TextMatrix(Y1, X1) = ""
                Next
            Next
        End If
End Sub

Private Sub FlexGriddestino_KeyPress(KeyAscii As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer

    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
                If FlexGridDestino.Row = FlexGridDestino.Rows - 1 Then
                    If FlexGridDestino.Col <> FlexGridDestino.Cols - 1 Then
                    FlexGridDestino.Col = FlexGridDestino.Col + 1
                    FlexGridDestino.Row = 1
                    FlexGridDestino.TopRow = 1
                    'SendKeys "{DOWN}"
                    'SendKeys "{UP}"
                    Else
                    CmdProximaFase.SetFocus
                    End If
                Else
                SendKeys "{DOWN}"
                End If
        ElseIf KeyAscii = 8 Then
            FlexGridDestino.Text = Mid(FlexGridDestino.Text, 1, Len(FlexGridDestino.Text) - 1)
            FlexGridDestino.Text = Format((SemPonto(FlexGridDestino.Text) / 100), "###,##0.00")
        ElseIf KeyAscii = 3 Then
        LblTransferencia.Caption = FlexGridDestino.Text
        ElseIf KeyAscii = 22 Then
            For Y1 = FlexGridDestino.Row To FlexGridDestino.RowSel
                For X1 = FlexGridDestino.Col To FlexGridDestino.ColSel
                FlexGridDestino.TextMatrix(Y1, X1) = LblTransferencia.Caption
                Next
            Next
        Else
        KeyAscii = 0
        End If
    Else
        FlexGridDestino.Text = FlexGridDestino.Text & Chr(KeyAscii)
        FlexGridDestino.Text = Format((SemPonto(FlexGridDestino.Text) / 100), "###,##0.00")
    End If

End Sub

Private Sub FlexGridOrigem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer
        If KeyCode = 46 Then
            For Y1 = FlexGridOrigem.Row To FlexGridOrigem.RowSel
                For X1 = FlexGridOrigem.Col To FlexGridOrigem.ColSel
                FlexGridOrigem.TextMatrix(Y1, X1) = ""
                Next
            Next
        End If
End Sub

Private Sub FlexGridOrigem_KeyPress(KeyAscii As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer

    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
                If FlexGridOrigem.Row = FlexGridOrigem.Rows - 1 Then
                    If FlexGridOrigem.Col <> FlexGridOrigem.Cols - 1 Then
                    FlexGridOrigem.Col = FlexGridOrigem.Col + 1
                    FlexGridOrigem.Row = 1
                    FlexGridOrigem.TopRow = 1
                    'SendKeys "{DOWN}"
                    'SendKeys "{UP}"
                    Else
                    CmdProximaFase.SetFocus
                    End If
                Else
                SendKeys "{DOWN}"
                End If
        ElseIf KeyAscii = 8 Then
            FlexGridOrigem.Text = Mid(FlexGridOrigem.Text, 1, Len(FlexGridOrigem.Text) - 1)
            FlexGridOrigem.Text = Format((SemPonto(FlexGridOrigem.Text) / 100), "###,##0.00")
        ElseIf KeyAscii = 3 Then
        LblTransferencia.Caption = FlexGridOrigem.Text
        ElseIf KeyAscii = 22 Then
            For Y1 = FlexGridOrigem.Row To FlexGridOrigem.RowSel
                For X1 = FlexGridOrigem.Col To FlexGridOrigem.ColSel
                FlexGridOrigem.TextMatrix(Y1, X1) = LblTransferencia.Caption
                Next
            Next
        Else
        KeyAscii = 0
        End If
    Else
        FlexGridOrigem.Text = FlexGridOrigem.Text & Chr(KeyAscii)
        FlexGridOrigem.Text = Format((SemPonto(FlexGridOrigem.Text) / 100), "###,##0.00")
    End If
End Sub

Private Sub FlexGridPanoramaNovaTabela_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer
        If KeyCode = 46 Then
            For Y1 = FlexGridPanoramaNovaTabela.Row To FlexGridPanoramaNovaTabela.RowSel
                For X1 = FlexGridPanoramaNovaTabela.Col To FlexGridPanoramaNovaTabela.ColSel
                FlexGridPanoramaNovaTabela.TextMatrix(Y1, X1) = ""
                Next
            Next
        End If
End Sub

Private Sub FlexGridPanoramaNovaTabela_KeyPress(KeyAscii As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer

    If FlexGridPanoramaNovaTabela.Col <> FlexGridPanoramaNovaTabela.Cols - 1 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii = 13 Then
                If OptNavLateral.Value = True Then
                    If FlexGridPanoramaNovaTabela.Col = FlexGridPanoramaNovaTabela.Cols - 1 Then
                        If FlexGridPanoramaNovaTabela.Row <> FlexGridPanoramaNovaTabela.Rows - 1 Then
                        FlexGridPanoramaNovaTabela.Row = FlexGridPanoramaNovaTabela.Row + 1
                        FlexGridPanoramaNovaTabela.Col = 2
                        SendKeys "{LEFT}"
                        Else
                        CmdProximaFase.SetFocus
                        End If
                    Else
                    SendKeys "{RIGHT}"
                    End If
                Else
                    If FlexGridPanoramaNovaTabela.Row = FlexGridPanoramaNovaTabela.Rows - 1 Then
                        If FlexGridPanoramaNovaTabela.Col <> FlexGridPanoramaNovaTabela.Cols - 1 Then
                        FlexGridPanoramaNovaTabela.Col = FlexGridPanoramaNovaTabela.Col + 1
                        FlexGridPanoramaNovaTabela.Row = 1
                        FlexGridPanoramaNovaTabela.TopRow = 1
                        'SendKeys "{DOWN}"
                        'SendKeys "{UP}"
                        Else
                        CmdProximaFase.SetFocus
                        End If
                    Else
                    SendKeys "{DOWN}"
                    End If
                End If
                
            ElseIf KeyAscii = 8 Then
                FlexGridPanoramaNovaTabela.Text = Mid(FlexGridPanoramaNovaTabela.Text, 1, Len(FlexGridPanoramaNovaTabela.Text) - 1)
                FlexGridPanoramaNovaTabela.Text = Format((SemPonto(FlexGridPanoramaNovaTabela.Text) / 100), "###,##0.00")
            ElseIf KeyAscii = 3 Then
            LblTransferencia.Caption = FlexGridPanoramaNovaTabela.Text
            ElseIf KeyAscii = 22 Then
                For Y1 = FlexGridPanoramaNovaTabela.Row To FlexGridPanoramaNovaTabela.RowSel
                    For X1 = FlexGridPanoramaNovaTabela.Col To FlexGridPanoramaNovaTabela.ColSel
                        If X1 = FlexGridPanoramaNovaTabela.Cols - 1 And IsNumeric(LblTransferencia.Caption) = False Then
                        FlexGridPanoramaNovaTabela.TextMatrix(Y1, X1) = LblTransferencia.Caption
                        ElseIf X1 <> FlexGridPanoramaNovaTabela.Cols - 1 And IsNumeric(LblTransferencia.Caption) = True Then
                        FlexGridPanoramaNovaTabela.TextMatrix(Y1, X1) = LblTransferencia.Caption
                        End If
                    Next
                Next
            
            Else
            KeyAscii = 0
            End If
        Else
            FlexGridPanoramaNovaTabela.Text = FlexGridPanoramaNovaTabela.Text & Chr(KeyAscii)
            FlexGridPanoramaNovaTabela.Text = Format((SemPonto(FlexGridPanoramaNovaTabela.Text) / 100), "###,##0.00")
        End If
    Else
        If UCase(Chr(KeyAscii)) <> "S" Then
            If KeyAscii = 13 Then
                If OptNavLateral.Value = True Then
                    If FlexGridPanoramaNovaTabela.Col = FlexGridPanoramaNovaTabela.Cols - 1 Then
                        If FlexGridPanoramaNovaTabela.Row <> FlexGridPanoramaNovaTabela.Rows - 1 Then
                        FlexGridPanoramaNovaTabela.Row = FlexGridPanoramaNovaTabela.Row + 1
                        FlexGridPanoramaNovaTabela.Col = 2
                        SendKeys "{LEFT}"
                        Else
                        CmdProximaFase.SetFocus
                        End If
                    Else
                    SendKeys "{RIGHT}"
                    End If
                Else
                    If FlexGridPanoramaNovaTabela.Row = FlexGridPanoramaNovaTabela.Rows - 1 Then
                        If FlexGridPanoramaNovaTabela.Col <> FlexGridPanoramaNovaTabela.Cols - 1 Then
                        FlexGridPanoramaNovaTabela.Col = FlexGridPanoramaNovaTabela.Col + 1
                        FlexGridPanoramaNovaTabela.Row = 1
                        FlexGridPanoramaNovaTabela.TopRow = 1
                        'SendKeys "{DOWN}"
                        'SendKeys "{UP}"
                        Else
                        CmdProximaFase.SetFocus
                        End If
                    Else
                    SendKeys "{DOWN}"
                    End If
                End If
                
            ElseIf KeyAscii = 8 Then
                FlexGridPanoramaNovaTabela.Text = Mid(FlexGridPanoramaNovaTabela.Text, 1, Len(FlexGridPanoramaNovaTabela.Text) - 1)
            ElseIf KeyAscii = 3 Then
            LblTransferencia.Caption = FlexGridPanoramaNovaTabela.Text
            ElseIf KeyAscii = 22 Then
                For Y1 = FlexGridPanoramaNovaTabela.Row To FlexGridPanoramaNovaTabela.RowSel
                    For X1 = FlexGridPanoramaNovaTabela.Col To FlexGridPanoramaNovaTabela.ColSel
                    FlexGridPanoramaNovaTabela.TextMatrix(Y1, X1) = LblTransferencia.Caption
                    Next
                Next
            
            Else
            KeyAscii = 0
            End If
        Else
        FlexGridPanoramaNovaTabela.Text = "S"
        End If
    End If
        
End Sub






Private Sub Form_Load()

Dim xCont As Integer

    For xCont = 0 To TabFase.Tabs - 1
    TabFase.TabEnabled(xCont) = False
    Next

TabFase.Tab = 0
TabFase.TabEnabled(0) = True

If de_informa.rsSel_CadLocalAirGROUP.State = 1 Then de_informa.rsSel_CadLocalAirGROUP.Close
de_informa.Sel_CadLocalAirgroup

    Do Until de_informa.rsSel_CadLocalAirGROUP.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAirGROUP.Fields("localidade")) ' & " - " & de_informa.rsSel_CadLocalAirgroup.Fields("SIGLA")
    ComboOrigem.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAirGROUP.Fields("localidade")) ' & " - " & de_informa.rsSel_CadLocalAirgroup.Fields("SIGLA")
    de_informa.rsSel_CadLocalAirGROUP.MoveNext
    Loop
    
Call OrdenaListBox(ListLocalidadesDisponives)


End Sub

Private Sub gridCiaAerea_Click()
TxtSiglaCiaAerea.Text = GridCiaAerea.Columns(0)
TxtNomeCiaAerea.Text = GridCiaAerea.Columns(1)
End Sub

Private Sub gridCiaAerea_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
gridCiaAerea_Click
End Sub

Private Sub GridTabelas_Click()
TxtAçãoaRealizar.Text = "Ajustar " & GridTabelas.Columns(2)
End Sub

Private Sub GridTabelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
GridTabelas_Click
End Sub

Private Sub ListLocalidadesDisponives_Click()
CmdAdicionaLocalidade.Enabled = True
End Sub

Private Sub ListLocalidadesSel_Click()
CmdRemoveLocalidade.Enabled = True
End Sub

Private Sub MskVigencia_GotFocus()
Call Date_MskEdBox_GotFocus(MskVigencia)
End Sub

Private Sub MskVigencia_LostFocus()
Call Date_MskEdBox_LostFocus(MskVigencia)
End Sub



Private Sub OptAjustaTabela_Click()
    If OptAjustaTabela.Value = False Then
    FraAjusta.Enabled = False
    OptAjustaTabela.Value = False
    TxtAçãoaRealizar.Text = "Cadastrar uma Nova Tabela"
    Else
    FraAjusta.Enabled = True
    OptAjustaTabela.Value = True
    TxtAçãoaRealizar.Text = "Ajustar " & GridTabelas.Columns(2)
    End If
End Sub

Private Sub OptDescrSistema_Click()
    With OptDescrSistema
        If .Value = True Then
        TxtDescrSistema.Enabled = True
        TxtDescrSistema.BackColor = xAmarelo
        TxtDescrUsuario.Enabled = False
        TxtDescrUsuario.BackColor = xBranco
        Else
        TxtDescrSistema.Enabled = False
        TxtDescrSistema.BackColor = xBranco
        TxtDescrUsuario.Enabled = True
        TxtDescrUsuario.BackColor = xAmarelo
        End If
    End With
End Sub

Private Sub OptDescrUsuario_Click()
    With OptDescrUsuario
        If .Value = False Then
        TxtDescrSistema.Enabled = True
        TxtDescrSistema.BackColor = xAmarelo
        TxtDescrUsuario.Enabled = False
        TxtDescrUsuario.BackColor = xBranco
        Else
        TxtDescrSistema.Enabled = False
        TxtDescrSistema.BackColor = xBranco
        TxtDescrUsuario.Enabled = True
        TxtDescrUsuario.BackColor = xAmarelo
        End If
    End With
End Sub

Private Sub OptDescrSistemaALT_Click()
    With OptDescrSistemaALT
        If .Value = True Then
        TxtDescrSistemaALT.Enabled = True
        TxtDescrSistemaALT.BackColor = xAmarelo
        TxtDescrUsuarioALT.Enabled = False
        TxtDescrUsuarioALT.BackColor = xBranco
        Else
        TxtDescrSistemaALT.Enabled = False
        TxtDescrSistemaALT.BackColor = xBranco
        TxtDescrUsuarioALT.Enabled = True
        TxtDescrUsuarioALT.BackColor = xAmarelo
        End If
    End With
End Sub

Private Sub OptDescrUserALT_Click()
    With OptDescrUserALT
        If .Value = False Then
        TxtDescrSistemaALT.Enabled = True
        TxtDescrSistemaALT.BackColor = xAmarelo
        TxtDescrUsuarioALT.Enabled = False
        TxtDescrUsuarioALT.BackColor = xBranco
        Else
        TxtDescrSistemaALT.Enabled = False
        TxtDescrSistemaALT.BackColor = xBranco
        TxtDescrUsuarioALT.Enabled = True
        TxtDescrUsuarioALT.BackColor = xAmarelo
        End If
    End With
End Sub

Private Sub OptNovaTabela_Click()
    If OptNovaTabela.Value = True Then
    FraAjusta.Enabled = False
    OptAjustaTabela.Value = False
    TxtAçãoaRealizar.Text = "Cadastrar uma Nova Tabela"
    Else
    FraAjusta.Enabled = True
    OptAjustaTabela.Value = True
    TxtAçãoaRealizar.Text = "Ajustar " & GridTabelas.Columns(2)
    End If
End Sub

Private Sub OptTabelaEspecifica_Click()
    If OptTabelaEspecifica.Value = False Then
    CmdBuscaCliente.Enabled = False
    TxtCGCCliente.Enabled = False
    TxtNomeCliente.Enabled = False
    TxtCGCCliente.BackColor = xBranco
    TxtNomeCliente.BackColor = xBranco
    FraDadosCliente.Enabled = False
    Else
    CmdBuscaCliente.Enabled = True
    TxtCGCCliente.Enabled = True
    TxtNomeCliente.Enabled = True
    TxtCGCCliente.BackColor = xAmarelo
    TxtNomeCliente.BackColor = xAmarelo
    FraDadosCliente.Enabled = True
    End If
End Sub

Private Sub OptTabelaOficial_Click()
    If OptTabelaOficial.Value = True Then
    CmdBuscaCliente.Enabled = False
    TxtCGCCliente.Enabled = False
    TxtNomeCliente.Enabled = False
    TxtCGCCliente.BackColor = xBranco
    TxtNomeCliente.BackColor = xBranco
    FraDadosCliente.Enabled = False
    Else
    CmdBuscaCliente.Enabled = True
    TxtCGCCliente.Enabled = True
    TxtNomeCliente.Enabled = True
    TxtCGCCliente.BackColor = xAmarelo
    TxtNomeCliente.BackColor = xAmarelo
    FraDadosCliente.Enabled = True
    End If
End Sub


Private Sub TxtAjustaValor_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub



Private Sub TxtPesoInicial_KeyPress(KeyAscii As Integer)
Call TextPesoBox_KeyPress(KeyAscii)
End Sub



Private Sub TxtPesoFinal_KeyPress(KeyAscii As Integer)
Call TextPesoBox_KeyPress(KeyAscii)
End Sub

Private Sub TxtQteFaixasPeso_Change()
    If Val(TxtQteFaixasPeso.Text) > 20 Then
    MsgBox "Este não é um número válido para este campo. Por favor, tente novamente...", vbInformation, ""
    TxtQteFaixasPeso.Text = ""
    End If
End Sub

Private Sub TxtQteFaixasPeso_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        Else
            If KeyAscii <> 8 Then
            KeyAscii = 0
            End If
        End If
    End If
End Sub



Private Sub TxtValorOficial_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub



Private Sub TxtDESCintec_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub



Private Sub TxtCharter_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub



Private Sub TxtValoranvisa_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub

Private Sub TxtOBS_Change()
TxtOBS.Text = UCase(TxtOBS.Text)
TxtOBS.SelStart = Len(TxtOBS.Text)
End Sub
