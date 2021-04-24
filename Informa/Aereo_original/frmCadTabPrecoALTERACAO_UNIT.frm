VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadTabPrecoALTERACAO_UNIT 
   Caption         =   "Reajuste de Tabela de Preço"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   360
   ClientWidth     =   11895
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
   Icon            =   "frmCadTabPrecoALTERACAO_UNIT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
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
      Left            =   9960
      TabIndex        =   0
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
      TabIndex        =   10
      Top             =   6780
      Width           =   7635
   End
   Begin VB.CommandButton CmdFaseAnterior 
      Caption         =   "<< Voltar"
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
      TabIndex        =   9
      Top             =   6780
      Width           =   1815
   End
   Begin TabDlg.SSTab TabFase 
      Height          =   6465
      Left            =   180
      TabIndex        =   12
      Top             =   180
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   10
      TabHeight       =   556
      TabCaption(0)   =   "Fase 0"
      TabPicture(0)   =   "frmCadTabPrecoALTERACAO_UNIT.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GridTabelas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtAçãoaRealizar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraLocalidades"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtOBSTAB"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Fase 1"
      TabPicture(1)   =   "frmCadTabPrecoALTERACAO_UNIT.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Line2"
      Tab(1).Control(4)=   "Label17"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Fase 2"
      TabPicture(2)   =   "frmCadTabPrecoALTERACAO_UNIT.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "OptNavVertical"
      Tab(2).Control(1)=   "OptNavLateral"
      Tab(2).Control(2)=   "FlexGridAjustar"
      Tab(2).Control(3)=   "FlexGridOrigem"
      Tab(2).Control(4)=   "FlexGridDestino"
      Tab(2).Control(5)=   "Label10"
      Tab(2).Control(6)=   "Line3"
      Tab(2).Control(7)=   "Label6"
      Tab(2).Control(8)=   "Label7"
      Tab(2).Control(9)=   "LblOrigem"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Fase 3"
      TabPicture(3)   =   "frmCadTabPrecoALTERACAO_UNIT.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label11"
      Tab(3).Control(1)=   "Line6"
      Tab(3).Control(2)=   "CmdALTTab"
      Tab(3).Control(3)=   "TxtDescrSistemaALT"
      Tab(3).Control(4)=   "Frame4"
      Tab(3).Control(5)=   "Frame6"
      Tab(3).ControlCount=   6
      Begin VB.Frame Frame3 
         Caption         =   "Tabela Original"
         ForeColor       =   &H00800000&
         Height          =   1875
         Left            =   -74940
         TabIndex        =   49
         Top             =   6360
         Visible         =   0   'False
         Width           =   11235
         Begin MSFlexGridLib.MSFlexGrid FlexGridEspelho 
            Height          =   1515
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   2672
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Aplique os Reajustes por Coluna"
         ForeColor       =   &H00000080&
         Height          =   1275
         Left            =   -74820
         TabIndex        =   46
         Top             =   1020
         Width           =   11235
         Begin MSFlexGridLib.MSFlexGrid FlexGridAjustes 
            Height          =   915
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Visible         =   0   'False
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   1614
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tabela Reajustada"
         Height          =   3975
         Left            =   -74820
         TabIndex        =   43
         Top             =   2340
         Width           =   11235
         Begin MSFlexGridLib.MSFlexGrid FlexGridAjustarPERC 
            Height          =   3615
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Visible         =   0   'False
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   6376
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
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Aguarde enquanto a Visualização é Montada ..."
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
            Left            =   1920
            TabIndex        =   45
            Top             =   960
            Width           =   7500
         End
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
         Left            =   -65400
         TabIndex        =   36
         Top             =   960
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
         Left            =   -67320
         TabIndex        =   35
         Top             =   960
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tabela Reajustada"
         Height          =   4815
         Left            =   -74820
         TabIndex        =   22
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
            TabIndex        =   23
            Top             =   3960
            Width           =   10995
         End
         Begin MSFlexGridLib.MSFlexGrid FlexGridAlteraTAB 
            Height          =   2175
            Left            =   120
            TabIndex        =   24
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
            TabIndex        =   25
            Top             =   2760
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
            TabIndex        =   26
            Top             =   2760
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
            TabIndex        =   30
            Top             =   1800
            Width           =   7710
         End
         Begin VB.Label Label5 
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
            TabIndex        =   29
            Top             =   3720
            Width           =   2835
         End
         Begin VB.Label Label8 
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
            TabIndex        =   28
            Top             =   2460
            Width           =   1890
         End
         Begin VB.Label Label12 
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
            TabIndex        =   27
            Top             =   2460
            Width           =   1965
         End
      End
      Begin VB.Frame Frame4 
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
         TabIndex        =   31
         Top             =   5760
         Width           =   3375
         Begin MSMask.MaskEdBox MskVigenciaALT 
            Height          =   285
            Left            =   1920
            TabIndex        =   32
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
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
            TabIndex        =   33
            Top             =   225
            Width           =   1680
         End
      End
      Begin VB.TextBox TxtDescrSistemaALT 
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
         TabIndex        =   21
         Top             =   5940
         Width           =   5235
      End
      Begin VB.CommandButton CmdALTTab 
         Caption         =   "Gravar Tabela Reajustada"
         Height          =   375
         Left            =   -66000
         TabIndex        =   20
         Top             =   5880
         Width           =   2415
      End
      Begin VB.TextBox TxtOBSTAB 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1365
         Left            =   180
         MaxLength       =   450
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   4920
         Width           =   5175
      End
      Begin VB.Frame FraLocalidades 
         Caption         =   "Localidades de Atendimento da Tabela Selecionada"
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
         Left            =   5460
         TabIndex        =   16
         Top             =   1500
         Width           =   5955
         Begin VB.CommandButton CmdRemoveLocalidade 
            Caption         =   "REMOVER"
            Enabled         =   0   'False
            Height          =   1635
            Left            =   2820
            TabIndex        =   8
            Top             =   2940
            Width           =   315
         End
         Begin VB.CommandButton CmdAdicionaLocalidade 
            Caption         =   "ADI C IONAR"
            Enabled         =   0   'False
            Height          =   1935
            Left            =   2820
            TabIndex        =   7
            Top             =   1020
            Width           =   315
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
            Left            =   3240
            TabIndex        =   5
            Top             =   300
            Width           =   2535
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
            TabIndex        =   3
            Top             =   300
            Width           =   2535
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
            Left            =   3240
            MultiSelect     =   2  'Extended
            TabIndex        =   6
            Top             =   1020
            Width           =   2535
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
            TabIndex        =   4
            Top             =   1020
            Width           =   2535
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Aguarde..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3877
            TabIndex        =   52
            Top             =   2655
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Aguarde..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   817
            TabIndex        =   51
            Top             =   2655
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Localidades Disponíveis"
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
            Left            =   180
            TabIndex        =   18
            Top             =   780
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Localidades da Tabela"
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
            Left            =   3240
            TabIndex        =   17
            Top             =   780
            Width           =   1620
         End
      End
      Begin VB.TextBox TxtAçãoaRealizar 
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
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   1178
         Width           =   11235
      End
      Begin MSDataGridLib.DataGrid GridTabelas 
         Bindings        =   "frmCadTabPrecoALTERACAO_UNIT.frx":007C
         Height          =   3255
         Left            =   180
         TabIndex        =   2
         Top             =   1620
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
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
         DataMember      =   "Sel_TabsCadastradas"
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "codtab"
            Caption         =   "codtab"
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
            DataField       =   "codcia"
            Caption         =   "codcia"
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
            Caption         =   "descricao"
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
            DataField       =   "cidade_origem"
            Caption         =   "cidade_origem"
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
            DataField       =   "inicio"
            Caption         =   "inicio"
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
            DataField       =   "fim"
            Caption         =   "fim"
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
            DataField       =   "tipotab"
            Caption         =   "tipotab"
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
            DataField       =   "nome_cliente"
            Caption         =   "nome_cliente"
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
            DataField       =   "cgc_cliente"
            Caption         =   "cgc_cliente"
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
            DataField       =   "taxaorigem"
            Caption         =   "taxaorigem"
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
         BeginProperty Column10 
            DataField       =   "corteorigem"
            Caption         =   "corteorigem"
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
         BeginProperty Column11 
            DataField       =   "excedorigem"
            Caption         =   "excedorigem"
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
         BeginProperty Column12 
            DataField       =   "taxadestino"
            Caption         =   "taxadestino"
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
         BeginProperty Column13 
            DataField       =   "cortedestino"
            Caption         =   "cortedestino"
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
         BeginProperty Column14 
            DataField       =   "exceddestino"
            Caption         =   "exceddestino"
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
         BeginProperty Column15 
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
         BeginProperty Column16 
            DataField       =   "obs"
            Caption         =   "obs"
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
         BeginProperty Column17 
            DataField       =   "data_cadastro"
            Caption         =   "data_cadastro"
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
         BeginProperty Column18 
            DataField       =   "usuario"
            Caption         =   "usuario"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid FlexGridAjustar 
         Height          =   3795
         Left            =   -74820
         TabIndex        =   37
         Top             =   1260
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6694
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
         Left            =   -74820
         TabIndex        =   38
         Top             =   5400
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
         Left            =   -66420
         TabIndex        =   39
         Top             =   5400
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
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -63650
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fase 1: Reajustando em Percentuais"
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
         TabIndex        =   48
         Top             =   540
         Width           =   3825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fase 2: Reajustando Registro a Registro e Taxas de Origem e Destino"
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
         TabIndex        =   42
         Top             =   540
         Width           =   7335
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -63650
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label6 
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
         Left            =   -65700
         TabIndex        =   41
         Top             =   5100
         Width           =   1965
      End
      Begin VB.Label Label7 
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
         Left            =   -74820
         TabIndex        =   40
         Top             =   5100
         Width           =   1890
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -63650
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fase 3: Visualização e Confirmação da Alteração da Tabela"
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
         TabIndex        =   34
         Top             =   540
         Width           =   6270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fase 0: Escolha a Tabela que você gostaria de Reajustar"
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
         TabIndex        =   15
         Top             =   540
         Width           =   6015
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
         TabIndex        =   14
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tabela a ser Reajustada"
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1740
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   180
         X2              =   11350
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.Label LblTransferencia 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   8160
      Width           =   11415
   End
End
Attribute VB_Name = "frmCadTabPrecoALTERACAO_UNIT"
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

Private Sub CmdALTTab_Click()
Dim X, Y, xCont As Integer
Dim xCodTab, xCodItemGeral, xCodItemTETC As Integer
Dim xDescricao, xTipoTab, xStatus, xUsarGeral As String

    If MskVigenciaALT.Text = "" Then
    MsgBox "Você deve informar quando esta Tabela entrará em vigor. Por favor, tente novamente...", vbInformation, ""
    Exit Sub
    End If
    
CmdALTTab.Enabled = False
CmdCancelarTodoProcesso.Enabled = False
CmdFaseAnterior.Enabled = False

de_informa.cn_informa.BeginTrans
    
    If de_informa.rsSel_CodCadTabPrecoEscopo.State = 1 Then de_informa.rsSel_CodCadTabPrecoEscopo.Close
    de_informa.Sel_CodCadTabPrecoEscopo
    
        If de_informa.rsSel_CodCadTabPrecoEscopo.RecordCount = 0 Then
        xCodTab = "1000"
        Else
        xCodTab = Val(de_informa.rsSel_CodCadTabPrecoEscopo.Fields("codtab")) + 1
        End If
             
        xDescricao = TxtDescrSistemaALT.Text
        
        If GridTabelas.Columns(6) = "OFICIAL" Then
        xTipoTab = "OFICIAL"
        Else
        xTipoTab = "ESPECIFICA"
        End If
        
        If CDate(MskVigenciaALT.Text) > Date Then
        xStatus = "AGUARDANDO"
        Else
        xStatus = "VIGORANDO"
        End If
    
    de_informa.Ins_CadTabPrecoEscopo xCodTab, GridTabelas.Columns(1), xDescricao, UCase(GridTabelas.Columns(3)), CDate(MskVigenciaALT.Text), CDate("01/01/1900"), xTipoTab, Trim(GridTabelas.Columns(7)), Trim(GridTabelas.Columns(8)), Trim(TxtOBS.Text), xStatus, DataHora("DATA"), xUsuario, FlexGridOrigem2.TextMatrix(0, 1), FlexGridOrigem2.TextMatrix(1, 1), FlexGridOrigem2.TextMatrix(2, 1), FlexGridDestino2.TextMatrix(0, 1), FlexGridDestino2.TextMatrix(1, 1), FlexGridDestino2.TextMatrix(2, 1)
    de_informa.Update_VigenciaTAB CDate(MskVigenciaALT.Text), GridTabelas.Columns(0)
        
        For Y = 1 To FlexGridAlteraTAB.Rows - 1
            If de_informa.rsSel_CodCadTabPrecoGeral.State = 1 Then de_informa.rsSel_CodCadTabPrecoGeral.Close
            de_informa.Sel_CodCadTabPrecoGeral
            
            If de_informa.rsSel_CodCadTabPrecoGeral.RecordCount = 0 Then
            xCodItemGeral = "1000"
            Else
            xCodItemGeral = Val(de_informa.rsSel_CodCadTabPrecoGeral.Fields("coditem")) + 1
            End If
        
        de_informa.Ins_CadTabPrecogeral xCodItemGeral, xCodTab, UCase(Trim(FlexGridAlteraTAB.TextMatrix(Y, 0))), FlexGridAlteraTAB.TextMatrix(Y, 1), FlexGridAlteraTAB.TextMatrix(Y, 2), FlexGridAlteraTAB.TextMatrix(Y, 3), FlexGridAlteraTAB.TextMatrix(Y, 4), FlexGridAlteraTAB.TextMatrix(Y, 5), FlexGridAlteraTAB.TextMatrix(Y, 6), FlexGridAlteraTAB.TextMatrix(Y, 7), FlexGridAlteraTAB.TextMatrix(Y, (FlexGridAlteraTAB.Cols - 1) - 4), FlexGridAlteraTAB.TextMatrix(Y, (FlexGridAlteraTAB.Cols - 1) - 3), (Val(SemPonto(FlexGridAlteraTAB.TextMatrix(Y, 8))) / 100), Trim(FlexGridAlteraTAB.TextMatrix(Y, (FlexGridAlteraTAB.Cols - 1) - 0))
        Next
        
        For Y = 1 To FlexGridAlteraTAB.Rows - 1
            For X = 1 To de_informa.rsSel_TabelaTETCods.RecordCount * 2
                If de_informa.rsSel_CodCadTabPrecoTETC.State = 1 Then de_informa.rsSel_CodCadTabPrecoTETC.Close
                de_informa.Sel_CodCadTabPrecoTETC
                
                If de_informa.rsSel_CodCadTabPrecoTETC.RecordCount = 0 Then
                xCodItemTETC = "1000"
                Else
                xCodItemTETC = Val(de_informa.rsSel_CodCadTabPrecoTETC.Fields("coditem")) + 1
                End If
                
                If CDbl(FlexGridAlteraTAB.TextMatrix(Y, (9 + X))) = 0 Then
                xUsarGeral = "S"
                Else
                xUsarGeral = ""
                End If
                
                de_informa.Ins_CadTabPrecoTETC xCodItemTETC, xCodTab, UCase(FlexGridAlteraTAB.TextMatrix(Y, 0)), Trim(Mid(FlexGridAlteraTAB.TextMatrix(0, (8 + X)), Len(FlexGridAlteraTAB.TextMatrix(0, (8 + X))) - 3)), CDbl(Mid(FlexGridAlteraTAB.TextMatrix(Y, (8 + X)), Len(FlexGridAlteraTAB.TextMatrix(Y, (8 + X))) - 3)), xUsarGeral, CDbl(FlexGridAlteraTAB.TextMatrix(Y, (8 + X + 1)))
                X = X + 1
            Next
        Next
    de_informa.cn_informa.CommitTrans
Call AtualizaStatusTabelas
MsgBox "Sua Tabela foi Cadastrada com sucesso. Automaticamente entrará nos cálculos do Sistema a partir da Data Informada", vbInformation, ""
mdiAereo.mnuArquivo.Enabled = True
mdiAereo.mnuCadastros.Enabled = True
mdiAereo.mnuEmissoes.Enabled = True
mdiAereo.mnuRelat.Enabled = True
mdiAereo.mnuSair.Enabled = True
Unload Me
End Sub

Private Sub CmdBuscaCliente_Click()
Set xForm = Me
frm_busca_cliente.Show 1
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
FlexGridPanoramaNovaTabela.Col = 2
FlexGridPanoramaNovaTabela.SetFocus

End Sub

Private Sub CmdLocalidades_Click()
    frmCadLocalidade.Show 1
    
ListLocalidadesDisponives.Clear

If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
de_informa.Sel_CadLocalAir "%"

    Do Until de_informa.rsSel_CadLocalAir.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade"))
    de_informa.rsSel_CadLocalAir.MoveNext
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
        Exit Sub
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
CmdAdicionaLocalidade.Enabled = False
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

Private Sub CmdIncluiExcluiRegiao_Click()
frmCadTabPrecoLocalidades.Show
End Sub

Private Sub CmdProximaFase_Click()

Dim X, Y, xCont As Integer

    If TabFase.Tab = 0 Then
        If GridTabelas.Enabled = False Then
        MsgBox "Não existe nenhuma Tabela Cadastrada. Reajustes não são possíveis. Este processo será cancelado.", vbCritical, "Erro"
        mdiAereo.mnuArquivo.Enabled = True
        mdiAereo.mnuCadastros.Enabled = True
        mdiAereo.mnuEmissoes.Enabled = True
        mdiAereo.mnuRelat.Enabled = True
        mdiAereo.mnuSair.Enabled = True
        Unload Me
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
        
        For Y = 1 To FlexGridAjustar.Rows - 1
        FlexGridAjustar.Row = Y
            For X = 2 To FlexGridAjustar.Cols - 2
            FlexGridAjustar.Col = X
                If FlexGridAjustar.Text = "" Then
                FlexGridAjustar.Text = "0,00"
                X = X - 1
                ElseIf CDbl(Val(SemPonto(FlexGridAjustar.Text)) / 100) = 0 And X >= 1 And X <= 7 Then
                MsgBox "Nenhum valor entre as Colunas 2 e 8 podem ser nulos. Para continuar, corrija este problema.", vbInformation, ""
                    
                
                FlexGridAjustar.SetFocus
                    If X = FlexGridAjustar.Cols - 1 Then
                    SendKeys "{LEFT}"
                    SendKeys "{RIGHT}"
                    Else
                    SendKeys "{RIGHT}"
                    SendKeys "{LEFT}"
                    End If
                    If X = FlexGridAjustar.Rows - 1 Then
                    SendKeys "{UP}"
                    SendKeys "{DOWN}"
                    Else
                    SendKeys "{DOWN}"
                    SendKeys "{UP}"
                    End If
                Exit Sub
                ElseIf CDbl(Val(SemPonto(FlexGridAjustar.Text)) / 100) = 0 And FlexGridAjustar.TextMatrix(0, X) = "Corte Charter" And CDbl(SemPonto(Val(FlexGridAjustar.TextMatrix(Y, FlexGridAjustar.Cols - 1))) / 100) > 0 Then
                MsgBox "Você informou um Valor Charter porém não informou qual será o corte de peso. Para continuar, corrija este problema.", vbInformation, ""
                FlexGridAjustar.SetFocus
                    If X = FlexGridAjustar.Cols - 1 Then
                    SendKeys "{LEFT}"
                    SendKeys "{RIGHT}"
                    Else
                    SendKeys "{RIGHT}"
                    SendKeys "{LEFT}"
                    End If
                    If X = FlexGridAjustar.Rows - 1 Then
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
        
CmdProximaFase.Enabled = False
CmdCancelarTodoProcesso.Enabled = False
CmdFaseAnterior.Enabled = False

        
        If TabFase.Tab = 1 Then
        
        Call TransicaoParaFase1
        
        ElseIf TabFase.Tab = 2 Then
        
        FlexGridAjustar.Visible = False
        FlexGridOrigem.Visible = False
        FlexGridDestino.Visible = False
        DoEvents
        
        FlexGridAjustar.Clear
        
        FlexGridAjustar.Clear
        
        FlexGridAjustar.Rows = FlexGridAjustarPERC.Rows
        FlexGridAjustar.Cols = FlexGridAjustarPERC.Cols
        
        FlexGridAjustar.FixedRows = FlexGridAjustarPERC.FixedRows
        FlexGridAjustar.FixedCols = FlexGridAjustarPERC.FixedCols
        
            For Y = 0 To FlexGridAjustar.Rows - 1
                For X = 0 To FlexGridAjustar.Cols - 1
                FlexGridAjustarPERC.Row = Y
                FlexGridAjustar.Row = Y
                FlexGridAjustar.Col = X
                FlexGridAjustarPERC.Col = X
                
                    If FlexGridAjustarPERC.CellFontBold = True Then
                    FlexGridAjustar.CellFontBold = True
                    Else
                    FlexGridAjustar.CellFontBold = False
                    End If
                
                FlexGridAjustar.CellAlignment = FlexGridAjustarPERC.CellAlignment
                FlexGridAjustar.CellBackColor = FlexGridAjustarPERC.CellBackColor
                FlexGridAjustar.Text = FlexGridAjustarPERC.Text
                FlexGridAjustar.ColWidth(X) = FlexGridAjustarPERC.ColWidth(X)
                Next
            Next
        
        FlexGridAjustar.Visible = True
        FlexGridOrigem.Visible = True
        FlexGridDestino.Visible = True
        DoEvents
        
        ElseIf TabFase.Tab = 3 Then
        FlexGridAlteraTAB.Visible = False
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
            
        
        FlexGridAlteraTAB.Clear
        
        FlexGridAlteraTAB.Rows = FlexGridAjustar.Rows
        FlexGridAlteraTAB.Cols = FlexGridAjustar.Cols
        
        FlexGridAlteraTAB.FixedRows = FlexGridAjustar.FixedRows
        FlexGridAlteraTAB.FixedCols = FlexGridAjustar.FixedCols
        
            For Y = 0 To FlexGridAlteraTAB.Rows - 1
                For X = 0 To FlexGridAlteraTAB.Cols - 1
                FlexGridAjustar.Row = Y
                FlexGridAlteraTAB.Row = Y
                FlexGridAlteraTAB.Col = X
                FlexGridAjustar.Col = X
                
                    If FlexGridAjustar.CellFontBold = True Then
                    FlexGridAlteraTAB.CellFontBold = True
                    Else
                    FlexGridAlteraTAB.CellFontBold = False
                    End If
                
                FlexGridAlteraTAB.CellAlignment = FlexGridAjustar.CellAlignment
                FlexGridAlteraTAB.CellBackColor = FlexGridAjustar.CellBackColor
                FlexGridAlteraTAB.Text = FlexGridAjustar.Text
                FlexGridAlteraTAB.ColWidth(X) = FlexGridAjustar.ColWidth(X)
                Next
            Next
        FlexGridAlteraTAB.Visible = True
        FlexGridOrigem2.Visible = True
        FlexGridDestino2.Visible = True
        DoEvents
        
            If Trim(GridTabelas.Columns(6)) = "OFICIAL" Then
            TxtDescrSistemaALT.Text = GridTabelas.Columns(2)
            Else
            TxtDescrSistemaALT.Text = GridTabelas.Columns(2)
            End If
            
            TxtDescrSistemaALT.BackColor = xBranco
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
        FlexGridPanoramaNovaTabela.Clear
        
        
        
        If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
        de_informa.Sel_Cadiata "%"
                
        FlexGridPanoramaNovaTabela.Cols = 12 + ((de_informa.rsSel_CadIATA.RecordCount - 1) * 2)
        FlexGridPanoramaNovaTabela.Rows = ListLocalidadesSel.ListCount + 1
        
        FlexGridPanoramaNovaTabela.FixedRows = 1
        FlexGridPanoramaNovaTabela.FixedCols = 2
            
        FlexGridPanoramaNovaTabela.TextMatrix(0, 0) = "Localidades"
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 0
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, 1) = "Sigla"
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 1
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, 2) = "Taxa Mínima"
        FlexGridPanoramaNovaTabela.Row = 0
        FlexGridPanoramaNovaTabela.Col = 2
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, 3) = "Até 25,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 4) = "Até 50,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 5) = "Até 300,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 6) = "Até 500,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 7) = "Até 1000,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 8) = "Acima de 1000,5"
        FlexGridPanoramaNovaTabela.TextMatrix(0, 9) = "Desc. Tab. Geral"
        
        FlexGridPanoramaNovaTabela.Col = 9
            For Y = 1 To ListLocalidadesSel.ListCount
            FlexGridPanoramaNovaTabela.Row = Y
            FlexGridPanoramaNovaTabela.CellBackColor = xLaranja
            Next
        
        FlexGridPanoramaNovaTabela.Row = 0
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
        FlexGridPanoramaNovaTabela.Col = 9
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        FlexGridPanoramaNovaTabela.ColWidth(0) = 2100
        FlexGridPanoramaNovaTabela.ColWidth(1) = 800
        FlexGridPanoramaNovaTabela.ColWidth(2) = 1300
        FlexGridPanoramaNovaTabela.ColWidth(3) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(4) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(5) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(6) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(7) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(8) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(9) = 1500
        
                    
        xCont = 10
            Do Until de_informa.rsSel_CadIATA.EOF
                        
            FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Cód. " & de_informa.rsSel_CadIATA.Fields("codigo")
            FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1000
            
            FlexGridPanoramaNovaTabela.Row = 0
            FlexGridPanoramaNovaTabela.Col = xCont
            FlexGridPanoramaNovaTabela.CellAlignment = 3
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            xCont = xCont + 1
            
            FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Desc. Cód. " & de_informa.rsSel_CadIATA.Fields("codigo")
            FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1500
            
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
            
            
            For xCont = 0 To ListLocalidadesSel.ListCount - 1
            FlexGridPanoramaNovaTabela.TextMatrix(xCont + 1, 0) = Mid(ListLocalidadesSel.List(xCont), 1, Len(ListLocalidadesSel.List(xCont)) - 6)
            FlexGridPanoramaNovaTabela.Row = xCont + 1
            FlexGridPanoramaNovaTabela.Col = 0
            FlexGridPanoramaNovaTabela.CellAlignment = 1
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            FlexGridPanoramaNovaTabela.TextMatrix(xCont + 1, 1) = Mid(ListLocalidadesSel.List(xCont), Len(ListLocalidadesSel.List(xCont)) - 3)
            FlexGridPanoramaNovaTabela.Row = xCont + 1
            FlexGridPanoramaNovaTabela.Col = 1
            FlexGridPanoramaNovaTabela.CellAlignment = 1
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            Next
        
    CmdIniciarDigitacao.Enabled = True
    FraPanoramaNovaTabela.Enabled = False
    End If
End If
            
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub FlexGridAjustarPERC_Scroll()
FlexGridAjustes.LeftCol = FlexGridAjustes.LeftCol
FlexGridEspelho.TopRow = FlexGridAjustarPERC.TopRow
FlexGridEspelho.LeftCol = FlexGridAjustarPERC.LeftCol
End Sub

Private Sub FlexGridAjustes_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer
        
        If KeyCode = 46 Then
            For Y1 = FlexGridAjustes.Row To FlexGridAjustes.RowSel
                For X1 = FlexGridAjustes.Col To FlexGridAjustes.ColSel
                FlexGridAjustes.TextMatrix(Y1, X1) = ""
                
                    For Y2 = 1 To FlexGridAjustarPERC.Rows - 1
                    FlexGridAjustarPERC.Col = X1
                    FlexGridAjustarPERC.Row = Y2
                    
                    FlexGridEspelho.Col = X1
                    FlexGridEspelho.Row = Y2
                    
                    FlexGridAjustarPERC.Text = FlexGridEspelho.Text
                    Next
                
                Next
            Next
        End If
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
                For Y = 1 To FlexGridAjustarPERC.Rows - 1
                FlexGridAjustarPERC.Col = FlexGridAjustes.Col
                FlexGridAjustarPERC.Row = Y
                
                FlexGridEspelho.Col = FlexGridAjustes.Col
                FlexGridEspelho.Row = Y
                
                FlexGridAjustarPERC.Text = Format(CDbl(FlexGridEspelho.Text) + (CDbl(FlexGridEspelho.Text) * (CDbl(FlexGridAjustes.TextMatrix(1, FlexGridAjustes.Col)) / 100)), "###,##0.00")
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
                
                    For Y = 1 To FlexGridAjustarPERC.Rows - 1
                    FlexGridAjustarPERC.Col = X1
                    FlexGridAjustarPERC.Row = Y
                    
                    FlexGridEspelho.Col = X1
                    FlexGridEspelho.Row = Y
                    
                    FlexGridAjustarPERC.Text = Format(CDbl(FlexGridEspelho.Text) + (CDbl(FlexGridEspelho.Text) * (CDbl(FlexGridAjustes.TextMatrix(1, FlexGridAjustes.Col)) / 100)), "###,##0.00")
                    Next
                ElseIf Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Corte Charter (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Valor Charter" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Tx. Terrestre" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Até (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, X1)) = "Vl. Kg. Ex." Then
                FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
                FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
                
                    For Y = 1 To FlexGridAjustarPERC.Rows - 1
                    FlexGridAjustarPERC.Col = X1
                    FlexGridAjustarPERC.Row = Y
                    
                    FlexGridEspelho.Col = X1
                    FlexGridEspelho.Row = Y
                    
                    If CDbl(FlexGridAjustarPERC.Text) > 0 Then
                    FlexGridAjustarPERC.Text = Format(FlexGridAjustes.Text, "###,##0.00")
                    End If
                    Next
                Else
                FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
                FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
                
                    For Y = 1 To FlexGridAjustarPERC.Rows - 1
                    FlexGridAjustarPERC.Col = X1
                    FlexGridAjustarPERC.Row = Y
                    
                    FlexGridEspelho.Col = X1
                    FlexGridEspelho.Row = Y
                    
                    FlexGridAjustarPERC.Text = Format(FlexGridAjustes.Text, "###,##0.00")
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
        
            For Y = 1 To FlexGridAjustarPERC.Rows - 1
            FlexGridAjustarPERC.Col = FlexGridAjustes.Col
            FlexGridAjustarPERC.Row = Y
            
            FlexGridEspelho.Col = FlexGridAjustes.Col
            FlexGridEspelho.Row = Y
            
            FlexGridAjustarPERC.Text = Format(CDbl(FlexGridEspelho.Text) + (CDbl(FlexGridEspelho.Text) * (CDbl(FlexGridAjustes.TextMatrix(1, FlexGridAjustes.Col)) / 100)), "###,##0.00")
            Next
        ElseIf Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Corte Charter (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Valor Charter" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Tx. Terrestre" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Até (Kg)" Or Trim(FlexGridAjustes.TextMatrix(0, FlexGridAjustes.Col)) = "Vl. Kg. Ex." Then
        FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
        FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
        
            For Y = 1 To FlexGridAjustarPERC.Rows - 1
            FlexGridAjustarPERC.Col = FlexGridAjustes.Col
            FlexGridAjustarPERC.Row = Y
            
            FlexGridEspelho.Col = FlexGridAjustes.Col
            FlexGridEspelho.Row = Y
            
            If CDbl(FlexGridAjustarPERC.Text) > 0 Then
            FlexGridAjustarPERC.Text = Format(FlexGridAjustes.Text, "###,##0.00")
            End If
            Next
        Else
        FlexGridAjustes.Text = FlexGridAjustes.Text & Chr(KeyAscii)
        FlexGridAjustes.Text = Format((SemPonto(FlexGridAjustes.Text) / 100), "###,##0.00")
        
            For Y = 1 To FlexGridAjustarPERC.Rows - 1
            FlexGridAjustarPERC.Col = FlexGridAjustes.Col
            FlexGridAjustarPERC.Row = Y
            
            FlexGridEspelho.Col = FlexGridAjustes.Col
            FlexGridEspelho.Row = Y
            
            FlexGridAjustarPERC.Text = Format(FlexGridAjustes.Text, "###,##0.00")
            Next
        End If
    End If
End Sub


Private Sub FlexGridAjustes_Scroll()

FlexGridAjustarPERC.TopRow = FlexGridAjustes.TopRow
FlexGridAjustarPERC.LeftCol = FlexGridAjustes.LeftCol

FlexGridEspelho.TopRow = FlexGridAjustes.TopRow
FlexGridEspelho.LeftCol = FlexGridAjustes.LeftCol

End Sub

Private Sub FlexGridAjustar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer
        If KeyCode = 46 Then
            For Y1 = FlexGridAjustar.Row To FlexGridAjustar.RowSel
                For X1 = FlexGridAjustar.Col To FlexGridAjustar.ColSel
                FlexGridAjustar.TextMatrix(Y1, X1) = ""
                Next
            Next
        End If
End Sub

Private Sub FlexGridAjustar_KeyPress(KeyAscii As Integer)
Dim X1, Y1, X2, Y2, xCont As Integer

    If FlexGridAjustar.Col <> FlexGridAjustar.Cols - 1 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii = 13 Then
                If OptNavLateral.Value = True Then
                    If FlexGridAjustar.Col = FlexGridAjustar.Cols - 1 Then
                        If FlexGridAjustar.Row <> FlexGridAjustar.Rows - 1 Then
                        FlexGridAjustar.Row = FlexGridAjustar.Row + 1
                        FlexGridAjustar.Col = 2
                        SendKeys "{LEFT}"
                        Else
                        CmdProximaFase.SetFocus
                        End If
                    Else
                    SendKeys "{RIGHT}"
                    End If
                Else
                    If FlexGridAjustar.Row = FlexGridAjustar.Rows - 1 Then
                        If FlexGridAjustar.Col <> FlexGridAjustar.Cols - 1 Then
                        FlexGridAjustar.Col = FlexGridAjustar.Col + 1
                        FlexGridAjustar.Row = 1
                        FlexGridAjustar.TopRow = 1
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
                FlexGridAjustar.Text = Mid(FlexGridAjustar.Text, 1, Len(FlexGridAjustar.Text) - 1)
                FlexGridAjustar.Text = Format((SemPonto(FlexGridAjustar.Text) / 100), "###,##0.00")
            ElseIf KeyAscii = 3 Then
            LblTransferencia.Caption = FlexGridAjustar.Text
            ElseIf KeyAscii = 22 Then
                For Y1 = FlexGridAjustar.Row To FlexGridAjustar.RowSel
                    For X1 = FlexGridAjustar.Col To FlexGridAjustar.ColSel
                        If X1 = FlexGridAjustar.Cols - 1 And IsNumeric(LblTransferencia.Caption) = False Then
                        FlexGridAjustar.TextMatrix(Y1, X1) = LblTransferencia.Caption
                        ElseIf X1 <> FlexGridAjustar.Cols - 1 And IsNumeric(LblTransferencia.Caption) = True Then
                        FlexGridAjustar.TextMatrix(Y1, X1) = LblTransferencia.Caption
                        End If
                    Next
                Next
            
            Else
            KeyAscii = 0
            End If
        Else
            FlexGridAjustar.Text = FlexGridAjustar.Text & Chr(KeyAscii)
            FlexGridAjustar.Text = Format((SemPonto(FlexGridAjustar.Text) / 100), "###,##0.00")
        End If
    Else
        If UCase(Chr(KeyAscii)) <> "S" Then
            If KeyAscii = 13 Then
                If OptNavLateral.Value = True Then
                    If FlexGridAjustar.Col = FlexGridAjustar.Cols - 1 Then
                        If FlexGridAjustar.Row <> FlexGridAjustar.Rows - 1 Then
                        FlexGridAjustar.Row = FlexGridAjustar.Row + 1
                        FlexGridAjustar.Col = 2
                        SendKeys "{LEFT}"
                        Else
                        CmdProximaFase.SetFocus
                        End If
                    Else
                    SendKeys "{RIGHT}"
                    End If
                Else
                    If FlexGridAjustar.Row = FlexGridAjustar.Rows - 1 Then
                        If FlexGridAjustar.Col <> FlexGridAjustar.Cols - 1 Then
                        FlexGridAjustar.Col = FlexGridAjustar.Col + 1
                        FlexGridAjustar.Row = 1
                        FlexGridAjustar.TopRow = 1
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
                FlexGridAjustar.Text = Mid(FlexGridAjustar.Text, 1, Len(FlexGridAjustar.Text) - 1)
            ElseIf KeyAscii = 3 Then
            LblTransferencia.Caption = FlexGridAjustar.Text
            ElseIf KeyAscii = 22 Then
                For Y1 = FlexGridAjustar.Row To FlexGridAjustar.RowSel
                    For X1 = FlexGridAjustar.Col To FlexGridAjustar.ColSel
                    FlexGridAjustar.TextMatrix(Y1, X1) = LblTransferencia.Caption
                    Next
                Next
            
            Else
            KeyAscii = 0
            End If
        Else
        FlexGridAjustar.Text = "S"
        End If
    End If

End Sub

Private Sub FlexGridEspelho_Scroll()

FlexGridAjustes.LeftCol = FlexGridEspelho.LeftCol

FlexGridAjustarPERC.TopRow = FlexGridEspelho.TopRow
FlexGridAjustarPERC.LeftCol = FlexGridEspelho.LeftCol
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
                FlexGridPanoramaNovaTabela.TextMatrix(Y1, X1) = LblTransferencia.Caption
                Next
            Next
        
        Else
        KeyAscii = 0
        End If
    Else
        FlexGridPanoramaNovaTabela.Text = FlexGridPanoramaNovaTabela.Text & Chr(KeyAscii)
        FlexGridPanoramaNovaTabela.Text = Format((SemPonto(FlexGridPanoramaNovaTabela.Text) / 100), "###,##0.00")
    End If
End Sub

Private Sub Form_Activate()

If de_informa.rsSel_TabsCadastradas.State = 1 Then de_informa.rsSel_TabsCadastradas.Close
de_informa.Sel_TabsCadastradas "%"

GridTabelas.DataMember = "Sel_TabsCadastradas"
GridTabelas.Refresh


    If de_informa.rsSel_TabsCadastradas.RecordCount = 0 Then
    GridTabelas.Enabled = False
    Else
    GridTabelas_Click
    End If
End Sub

Private Sub Form_Load()

Dim xCont As Integer

    For xCont = 0 To TabFase.Tabs - 1
    TabFase.TabEnabled(xCont) = False
    Next

TabFase.Tab = 0
TabFase.TabEnabled(0) = True

GridTabelas.DataMember = "Sel_TabsCadastradas"
GridTabelas.Refresh

End Sub

Private Sub gridCiaAerea_Click()
TxtSiglaCiaAerea.Text = GridCiaAerea.Columns(0)
TxtNomeCiaAerea.Text = GridCiaAerea.Columns(1)
End Sub

Private Sub gridCiaAerea_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
gridCiaAerea_Click
End Sub

Private Sub GridTabelas_Click()
ListLocalidadesDisponives.Visible = False
ListLocalidadesSel.Visible = False
DoEvents

TxtAçãoaRealizar.Text = GridTabelas.Columns(2)
TxtOBSTAB.Text = GridTabelas.Columns(16)

Dim X, Y As Integer
FraLocalidades.Enabled = False
ListLocalidadesDisponives.Clear

If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
de_informa.Sel_CadLocalAir "%"

    Do Until de_informa.rsSel_CadLocalAir.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade"))
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop
    
If de_informa.rsSel_TabelaGeral.State = 1 Then de_informa.rsSel_TabelaGeral.Close
de_informa.Sel_TabelaGeral Val(Trim(GridTabelas.Columns(0))), "%"
 
ListLocalidadesSel.Clear

    Do Until de_informa.rsSel_TabelaGeral.EOF
    ListLocalidadesSel.AddItem PriMaiuscula(de_informa.rsSel_TabelaGeral.Fields("destino"))
    de_informa.rsSel_TabelaGeral.MoveNext
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
    DoEvents
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
    DoEvents
    Loop
    'FIM DO TRECHO QUE AVERIGUA LIST BOX
    
CmdAdicionaLocalidade.Enabled = False
FraLocalidades.Enabled = True
ListLocalidadesDisponives.Visible = True
ListLocalidadesSel.Visible = True
DoEvents
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

Private Sub MskVigenciaALT_GotFocus()
Call Date_MskEdBox_GotFocus(MskVigenciaALT)
End Sub

Private Sub MskVigenciaALT_LostFocus()
Call Date_MskEdBox_LostFocus(MskVigenciaALT)
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

Private Sub TransicaoParaFase1()
FlexGridOrigem.Visible = False
FlexGridDestino.Visible = False
FlexGridAjustar.Visible = False
FlexGridAjustes.Visible = False
DoEvents

FlexGridOrigem.Clear
FlexGridOrigem.Rows = 3
FlexGridOrigem.Cols = 2
FlexGridOrigem.FixedRows = 0
FlexGridOrigem.FixedCols = 1
FlexGridOrigem.TextMatrix(0, 0) = "Valor"
FlexGridOrigem.TextMatrix(0, 1) = Trim(GridTabelas.Columns(9))
FlexGridOrigem.TextMatrix(1, 0) = "Até"
FlexGridOrigem.TextMatrix(1, 1) = Trim(GridTabelas.Columns(10))
FlexGridOrigem.TextMatrix(2, 0) = "Kg Exced."
FlexGridOrigem.TextMatrix(2, 1) = Trim(GridTabelas.Columns(11))
FlexGridOrigem.ColWidth(0) = 1300
FlexGridOrigem.ColWidth(1) = 1300

FlexGridDestino.Clear
FlexGridDestino.Rows = 3
FlexGridDestino.Cols = 2
FlexGridDestino.FixedRows = 0
FlexGridDestino.FixedCols = 1
FlexGridDestino.TextMatrix(0, 0) = "Valor"
FlexGridDestino.TextMatrix(0, 1) = Trim(GridTabelas.Columns(12))
FlexGridDestino.TextMatrix(1, 0) = "Até"
FlexGridDestino.TextMatrix(1, 1) = Trim(GridTabelas.Columns(13))
FlexGridDestino.TextMatrix(2, 0) = "Kg Exced."
FlexGridDestino.TextMatrix(2, 1) = Trim(GridTabelas.Columns(14))
FlexGridDestino.ColWidth(0) = 1300
FlexGridDestino.ColWidth(1) = 1300
DoEvents
        
        FlexGridAjustarPERC.Clear
        FlexGridAjustes.Clear

        
        
        If de_informa.rsSel_TabelaTETC.State = 1 Then de_informa.rsSel_TabelaTETC.Close
        If de_informa.rsSel_TabelaTETCods.State = 1 Then de_informa.rsSel_TabelaTETCods.Close
                
        de_informa.Sel_TabelaTETCods Trim(GridTabelas.Columns(0))
        
        FlexGridAjustarPERC.Cols = 12 + (de_informa.rsSel_TabelaTETCods.RecordCount * 2)
        FlexGridAjustarPERC.Rows = ListLocalidadesSel.ListCount + 1
        FlexGridAjustarPERC.FixedCols = 1
        FlexGridAjustarPERC.FixedRows = 1
        
        FlexGridAjustes.Cols = FlexGridAjustarPERC.Cols - 3
        FlexGridAjustes.Rows = 2
        FlexGridAjustes.FixedCols = FlexGridAjustarPERC.FixedCols
        FlexGridAjustes.FixedRows = FlexGridAjustarPERC.FixedRows
        
        
        Y = 1
            For X = 0 To ListLocalidadesSel.ListCount - 1
            If de_informa.rsSel_TabelaGeral.State = 1 Then de_informa.rsSel_TabelaGeral.Close
            de_informa.Sel_TabelaGeral Val(Trim(GridTabelas.Columns(0))), ListLocalidadesSel.List(X)
            
            FlexGridAjustarPERC.TextMatrix(Y, 0) = ListLocalidadesSel.List(X)
            
                Do Until de_informa.rsSel_TabelaGeral.EOF
                    With de_informa.rsSel_TabelaGeral
                    
                    FlexGridAjustarPERC.TextMatrix(Y, 1) = (Trim(.Fields("taxaminima")))
                    FlexGridAjustarPERC.TextMatrix(Y, 2) = (Trim(.Fields("ate25")))
                    FlexGridAjustarPERC.TextMatrix(Y, 3) = (Trim(.Fields("ate50")))
                    FlexGridAjustarPERC.TextMatrix(Y, 4) = (Trim(.Fields("ate300")))
                    FlexGridAjustarPERC.TextMatrix(Y, 5) = (Trim(.Fields("ate500")))
                    FlexGridAjustarPERC.TextMatrix(Y, 6) = (Trim(.Fields("ate1000")))
                    FlexGridAjustarPERC.TextMatrix(Y, 7) = (Trim(.Fields("acima1000")))
                    FlexGridAjustarPERC.TextMatrix(Y, 8) = (Trim(.Fields("descontogeral")))
                    End With
                de_informa.rsSel_TabelaGeral.MoveNext
                Loop
                Y = Y + 1
            Next
        FlexGridAjustarPERC.TextMatrix(0, 0) = "Localidades"
        FlexGridAjustarPERC.Row = 0
        FlexGridAjustarPERC.Col = 0
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        
        FlexGridAjustarPERC.TextMatrix(0, 1) = "Taxa Mínima"
        FlexGridAjustarPERC.Row = 0
        FlexGridAjustarPERC.Col = 1
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        
        FlexGridAjustarPERC.TextMatrix(0, 2) = "Até 25,5"
        FlexGridAjustarPERC.TextMatrix(0, 3) = "Até 50,5"
        FlexGridAjustarPERC.TextMatrix(0, 4) = "Até 300,5"
        FlexGridAjustarPERC.TextMatrix(0, 5) = "Até 500,5"
        FlexGridAjustarPERC.TextMatrix(0, 6) = "Até 1000,5"
        FlexGridAjustarPERC.TextMatrix(0, 7) = "Acima de 1000,5"
        FlexGridAjustarPERC.TextMatrix(0, 8) = "Desc. Tab. Geral (%)"
        
        FlexGridAjustes.TextMatrix(1, 0) = "Aplique os Descontos (%)"
        FlexGridAjustes.Row = 0
        FlexGridAjustes.Col = 0
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        
        FlexGridAjustes.TextMatrix(0, 1) = "Taxa Mínima"
        FlexGridAjustes.Row = 0
        FlexGridAjustes.Col = 1
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        
        FlexGridAjustes.TextMatrix(0, 2) = "Até 25,5"
        FlexGridAjustes.TextMatrix(0, 3) = "Até 50,5"
        FlexGridAjustes.TextMatrix(0, 4) = "Até 300,5"
        FlexGridAjustes.TextMatrix(0, 5) = "Até 500,5"
        FlexGridAjustes.TextMatrix(0, 6) = "Até 1000,5"
        FlexGridAjustes.TextMatrix(0, 7) = "Acima de 1000,5"
        FlexGridAjustes.TextMatrix(0, 8) = "Desc. Tab. Geral (%)"
        
        
        FlexGridAjustarPERC.Col = 8
        FlexGridAjustes.Col = FlexGridAjustar.Col
            For Y = 1 To FlexGridAjustarPERC.Rows - 1
            FlexGridAjustarPERC.Row = Y
            FlexGridAjustarPERC.CellBackColor = xLaranja
            
            FlexGridAjustes.Row = 1
            FlexGridAjustes.CellBackColor = xLaranja
            Next
        
        FlexGridAjustarPERC.Row = 0
        FlexGridAjustarPERC.Col = 2
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        FlexGridAjustarPERC.Col = 3
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        FlexGridAjustarPERC.Col = 4
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        FlexGridAjustarPERC.Col = 5
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        FlexGridAjustarPERC.Col = 6
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        FlexGridAjustarPERC.Col = 7
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        FlexGridAjustarPERC.Col = 8
        FlexGridAjustarPERC.CellAlignment = 3
        FlexGridAjustarPERC.CellFontBold = True
        
        FlexGridAjustarPERC.ColWidth(0) = 2100
        FlexGridAjustarPERC.ColWidth(1) = 1300
        FlexGridAjustarPERC.ColWidth(2) = 1500
        FlexGridAjustarPERC.ColWidth(3) = 1500
        FlexGridAjustarPERC.ColWidth(4) = 1500
        FlexGridAjustarPERC.ColWidth(5) = 1500
        FlexGridAjustarPERC.ColWidth(6) = 1500
        FlexGridAjustarPERC.ColWidth(7) = 1500
        FlexGridAjustarPERC.ColWidth(8) = 2000
        
        FlexGridAjustes.Row = 0
        FlexGridAjustes.Col = 2
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        FlexGridAjustes.Col = 3
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        FlexGridAjustes.Col = 4
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        FlexGridAjustes.Col = 5
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        FlexGridAjustes.Col = 6
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        FlexGridAjustes.Col = 7
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        FlexGridAjustes.Col = 8
        FlexGridAjustes.CellAlignment = 3
        FlexGridAjustes.CellFontBold = True
        
        FlexGridAjustes.ColWidth(0) = 2100
        FlexGridAjustes.ColWidth(1) = 1300
        FlexGridAjustes.ColWidth(2) = 1500
        FlexGridAjustes.ColWidth(3) = 1500
        FlexGridAjustes.ColWidth(4) = 1500
        FlexGridAjustes.ColWidth(5) = 1500
        FlexGridAjustes.ColWidth(6) = 1500
        FlexGridAjustes.ColWidth(7) = 1500
        FlexGridAjustes.ColWidth(8) = 2000
        
        
        
        xCont = 9
            Do Until de_informa.rsSel_TabelaTETCods.EOF
                        
            FlexGridAjustarPERC.TextMatrix(0, xCont) = "Cód. " & de_informa.rsSel_TabelaTETCods.Fields("codtetc")
            FlexGridAjustarPERC.ColWidth(xCont) = 1000
            
            FlexGridAjustarPERC.Row = 0
            FlexGridAjustarPERC.Col = xCont
            FlexGridAjustarPERC.CellAlignment = 3
            FlexGridAjustarPERC.CellFontBold = True
            
            FlexGridAjustes.TextMatrix(0, xCont) = "Cód. " & de_informa.rsSel_TabelaTETCods.Fields("codtetc")
            FlexGridAjustes.ColWidth(xCont) = 1000
            
            FlexGridAjustes.Row = 0
            FlexGridAjustes.Col = xCont
            FlexGridAjustes.CellAlignment = 3
            FlexGridAjustes.CellFontBold = True
            
            
            xCont = xCont + 1
            
            FlexGridAjustarPERC.TextMatrix(0, xCont) = "Desc. Cód. " & de_informa.rsSel_TabelaTETCods.Fields("codtetc") & " (%)"
            FlexGridAjustarPERC.ColWidth(xCont) = 1800
            
            FlexGridAjustarPERC.Row = 0
            FlexGridAjustarPERC.Col = xCont
            FlexGridAjustarPERC.CellAlignment = 3
            FlexGridAjustarPERC.CellFontBold = True
            
            FlexGridAjustes.TextMatrix(0, xCont) = "Desc. Cód. " & de_informa.rsSel_TabelaTETCods.Fields("codtetc") & " (%)"
            FlexGridAjustes.ColWidth(xCont) = 1800
            
            FlexGridAjustes.Row = 0
            FlexGridAjustes.Col = xCont
            FlexGridAjustes.CellAlignment = 3
            FlexGridAjustes.CellFontBold = True
            
                For Y = 1 To FlexGridAjustarPERC.Rows - 1
                FlexGridAjustarPERC.Row = Y
                FlexGridAjustarPERC.CellBackColor = xAmarelo
                
                FlexGridAjustes.Row = 1
                FlexGridAjustes.CellBackColor = xAmarelo
                Next
            
            
                For Y = 1 To FlexGridAjustarPERC.Rows - 1
                If de_informa.rsSel_TabelaTETC.State = 1 Then de_informa.rsSel_TabelaTETC.Close
                de_informa.Sel_TabelaTETC Trim(GridTabelas.Columns(0)), Trim(FlexGridAjustarPERC.TextMatrix(Y, 0)), de_informa.rsSel_TabelaTETCods.Fields("codtetc")
                    If de_informa.rsSel_TabelaTETC.RecordCount > 0 Then
                    FlexGridAjustarPERC.TextMatrix(Y, xCont - 1) = de_informa.rsSel_TabelaTETC.Fields("porkilo")
                    FlexGridAjustarPERC.TextMatrix(Y, xCont) = de_informa.rsSel_TabelaTETC.Fields("desconto")
                    End If
                Next
                
            
            xCont = xCont + 1
            de_informa.rsSel_TabelaTETCods.MoveNext
            Loop
            
            
            
            Y = 1
            For X = 0 To ListLocalidadesSel.ListCount - 1
            If de_informa.rsSel_TabelaGeral.State = 1 Then de_informa.rsSel_TabelaGeral.Close
            de_informa.Sel_TabelaGeral Val(Trim(GridTabelas.Columns(0))), ListLocalidadesSel.List(X)
                Do Until de_informa.rsSel_TabelaGeral.EOF
                    With de_informa.rsSel_TabelaGeral
                    FlexGridAjustarPERC.TextMatrix(Y, xCont + 0) = (Trim(.Fields("cortecharter")))
                    FlexGridAjustarPERC.TextMatrix(Y, xCont + 1) = (Trim(.Fields("charter")))
                    FlexGridAjustarPERC.TextMatrix(Y, xCont + 2) = (Trim(.Fields("txterrestre")))
                    End With
                
                de_informa.rsSel_TabelaGeral.MoveNext
                Loop
            Y = Y + 1
            Next
            
            X = xCont
            
            FlexGridAjustarPERC.Row = 0
            
            FlexGridAjustarPERC.TextMatrix(0, xCont) = "Corte Charter (Kg)"
            FlexGridAjustarPERC.ColWidth(xCont) = 1600
            FlexGridAjustarPERC.Col = xCont
            FlexGridAjustarPERC.CellAlignment = 3
            FlexGridAjustarPERC.CellFontBold = True
            
                For Y = 1 To FlexGridAjustarPERC.Rows - 1
                FlexGridAjustarPERC.Row = Y
                FlexGridAjustarPERC.CellBackColor = xCinzaClaro
                Next
            
            xCont = xCont + 1
            FlexGridAjustarPERC.Row = 0
            FlexGridAjustarPERC.TextMatrix(0, xCont) = "Valor Charter"
            FlexGridAjustarPERC.ColWidth(xCont) = 1300
            FlexGridAjustarPERC.Col = xCont
            FlexGridAjustarPERC.CellAlignment = 3
            FlexGridAjustarPERC.CellFontBold = True
            
                For Y = 1 To FlexGridAjustarPERC.Rows - 1
                FlexGridAjustarPERC.Row = Y
                FlexGridAjustarPERC.CellBackColor = xCinzaClaro
                Next
                
                
            xCont = xCont + 1
            FlexGridAjustarPERC.Row = 0
            FlexGridAjustarPERC.TextMatrix(0, xCont) = "Tx. Terrestre"
            FlexGridAjustarPERC.ColWidth(xCont) = 1300
            FlexGridAjustarPERC.Col = xCont
            FlexGridAjustarPERC.CellAlignment = 3
            FlexGridAjustarPERC.CellFontBold = True
            
                For Y = 1 To FlexGridAjustarPERC.Rows - 1
                FlexGridAjustarPERC.Row = Y
                FlexGridAjustarPERC.CellBackColor = xBranco
                Next
            
                For Y = 1 To FlexGridAjustarPERC.Rows - 1
                    For X = 2 To FlexGridAjustarPERC.Cols - 2
                    FlexGridAjustarPERC.Row = Y
                    FlexGridAjustarPERC.Col = X
                        If Len(FlexGridAjustarPERC.TextMatrix(Y, X)) > 0 Then
                        FlexGridAjustarPERC.TextMatrix(Y, X) = Format(CDbl(FlexGridAjustarPERC.TextMatrix(Y, X)), "###,##0.00")
                        Else
                        FlexGridAjustarPERC.TextMatrix(Y, X) = "0,00"
                        End If
                    FlexGridAjustarPERC.CellAlignment = 7
                    DoEvents
                    Next
                Next
            
            With FlexGridAjustarPERC
            FlexGridEspelho.Clear
            FlexGridEspelho.Rows = .Rows
            FlexGridEspelho.Cols = .Cols
            FlexGridEspelho.FixedCols = .FixedCols
            FlexGridEspelho.FixedCols = .FixedCols
            End With
        
            For Y = 0 To FlexGridAjustarPERC.Rows - 1
                For X = 0 To FlexGridAjustarPERC.Cols - 1
                FlexGridEspelho.Col = X
                FlexGridEspelho.Row = Y
                
                FlexGridAjustarPERC.Col = X
                FlexGridAjustarPERC.Row = Y
                
                    With FlexGridAjustarPERC
                    FlexGridEspelho.CellBackColor = .CellBackColor
                    FlexGridEspelho.CellFontBold = .CellFontBold
                    FlexGridEspelho.CellAlignment = .CellAlignment
                    FlexGridEspelho.ColWidth(X) = .ColWidth(X)
                    FlexGridEspelho.Text = .Text
                    End With
                DoEvents
                Next
            Next
            
            
FlexGridOrigem.Visible = True
FlexGridDestino.Visible = True
FlexGridAjustarPERC.Visible = True
FlexGridAjustes.Visible = True
DoEvents
              
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


