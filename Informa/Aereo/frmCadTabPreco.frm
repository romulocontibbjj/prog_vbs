VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadTabPreco 
   Caption         =   "Cadastramento de Tabela de Preço"
   ClientHeight    =   6645
   ClientLeft      =   390
   ClientTop       =   1095
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadTabPreco.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   10815
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   6780
      Width           =   1815
   End
   Begin TabDlg.SSTab TabFase 
      Height          =   6465
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   10
      TabHeight       =   556
      TabCaption(0)   =   "Fase 0"
      TabPicture(0)   =   "frmCadTabPreco.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "OptAjustaTabela"
      Tab(0).Control(1)=   "FraAjusta"
      Tab(0).Control(2)=   "OptNovaTabela"
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(4)=   "Line1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Fase 1"
      TabPicture(1)   =   "frmCadTabPreco.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(2)=   "FraCiaAerea"
      Tab(1).Control(3)=   "FraTipoTabela"
      Tab(1).Control(4)=   "FraLocalidades"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Fase 2"
      TabPicture(2)   =   "frmCadTabPreco.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Line5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label27"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "FraPanoramaNovaTabela"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "CmdIniciarDigitacao"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "CmdZerarDigitacao"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Fase 3"
      TabPicture(3)   =   "frmCadTabPreco.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Fase 4"
      TabPicture(4)   =   "frmCadTabPreco.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label20"
      Tab(4).Control(1)=   "Line6"
      Tab(4).Control(2)=   "CmdCadastrarTabela"
      Tab(4).Control(3)=   "FraVigencia"
      Tab(4).Control(4)=   "FraDescrNovaTabela"
      Tab(4).Control(5)=   "FraNovaTabela"
      Tab(4).ControlCount=   6
      Begin VB.CommandButton CmdZerarDigitacao 
         Caption         =   "Zerar Digitação"
         Height          =   375
         Left            =   1860
         TabIndex        =   53
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton CmdIniciarDigitacao 
         Caption         =   "Iniciar Digitação"
         Height          =   375
         Left            =   180
         TabIndex        =   52
         Top             =   1020
         Width           =   1575
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
         Left            =   180
         TabIndex        =   50
         Top             =   1440
         Width           =   11235
         Begin MSFlexGridLib.MSFlexGrid FlexGridPanoramaNovaTabela 
            Height          =   4455
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   7858
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
      Begin VB.OptionButton OptAjustaTabela 
         Caption         =   "Ajustar uma Tabela"
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
         Left            =   -74700
         TabIndex        =   44
         Top             =   2220
         Width           =   2595
      End
      Begin VB.Frame FraAjusta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   -74820
         TabIndex        =   36
         Top             =   2400
         Width           =   10095
         Begin VB.TextBox TxtAjustaDesconto 
            Alignment       =   1  'Right Justify
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
            Left            =   3960
            TabIndex        =   41
            Top             =   1545
            Width           =   735
         End
         Begin VB.OptionButton OptAjustaDesconto 
            Caption         =   "Substituir TODOS os Descontos INTEC com "
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
            Left            =   300
            TabIndex        =   40
            Top             =   1590
            Width           =   3615
         End
         Begin VB.TextBox TxtAjustaValor 
            Alignment       =   1  'Right Justify
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
            Left            =   3180
            TabIndex        =   39
            Top             =   435
            Width           =   735
         End
         Begin VB.OptionButton OptAjustaManual 
            Caption         =   "Ajustar Tabela Manualmente"
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
            Left            =   300
            TabIndex        =   38
            Top             =   2700
            Width           =   3375
         End
         Begin VB.OptionButton OptAjustaValor 
            Caption         =   "Ajustar os Valores desta Tabela em"
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
            Left            =   300
            TabIndex        =   37
            Top             =   480
            Width           =   2835
         End
         Begin MSDataGridLib.DataGrid GridTabelas 
            Height          =   2475
            Left            =   5400
            TabIndex        =   42
            Top             =   540
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   4366
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tabelas Cadastradas"
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
            Left            =   5460
            TabIndex        =   43
            Top             =   300
            Width           =   1500
         End
      End
      Begin VB.OptionButton OptNovaTabela 
         Caption         =   "Cadastrar Nova Tabela"
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
         Left            =   -74700
         TabIndex        =   34
         Top             =   1260
         Width           =   2595
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
         Height          =   5355
         Left            =   -74820
         TabIndex        =   27
         Top             =   960
         Width           =   6255
         Begin VB.CommandButton CmdRemoveLocalidade 
            Caption         =   "REMOVER"
            Enabled         =   0   'False
            Height          =   2175
            Left            =   2970
            TabIndex        =   31
            Top             =   3015
            Width           =   315
         End
         Begin VB.CommandButton CmdAdicionaLocalidade 
            Caption         =   "ADI C IONAR"
            Enabled         =   0   'False
            Height          =   2175
            Left            =   2970
            TabIndex        =   30
            Top             =   840
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
            Left            =   3180
            TabIndex        =   32
            Top             =   300
            Width           =   2895
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
            TabIndex        =   45
            Top             =   300
            Width           =   2895
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
            Height          =   4350
            Left            =   3360
            MultiSelect     =   2  'Extended
            TabIndex        =   29
            Top             =   840
            Width           =   2715
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
            Height          =   4350
            Left            =   180
            MultiSelect     =   2  'Extended
            TabIndex        =   28
            Top             =   840
            Width           =   2715
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
         Left            =   -68460
         TabIndex        =   18
         Top             =   960
         Width           =   4875
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
            TabIndex        =   21
            Top             =   900
            Width           =   4635
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
               TabIndex        =   24
               Top             =   1020
               Width           =   1875
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
               TabIndex        =   23
               Top             =   465
               Width           =   4395
            End
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
               TabIndex        =   22
               Top             =   1065
               Width           =   2415
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
               TabIndex        =   26
               Top             =   240
               Width           =   1170
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
               TabIndex        =   25
               Top             =   840
               Width           =   1080
            End
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
            TabIndex        =   20
            Top             =   600
            Width           =   1995
         End
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
            TabIndex        =   19
            Top             =   300
            Width           =   1995
         End
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
         Left            =   -68460
         TabIndex        =   16
         Top             =   3540
         Width           =   4875
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
            TabIndex        =   47
            Top             =   2280
            Width           =   3315
         End
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
            TabIndex        =   46
            Top             =   2280
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid GridCiaAerea 
            Bindings        =   "frmCadTabPreco.frx":0098
            Height          =   1635
            Left            =   120
            TabIndex        =   17
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
            TabIndex        =   49
            Top             =   2040
            Width           =   735
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
            TabIndex        =   48
            Top             =   2040
            Width           =   345
         End
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
         Height          =   3195
         Left            =   -74820
         TabIndex        =   13
         Top             =   960
         Width           =   10095
         Begin MSFlexGridLib.MSFlexGrid FlexGridNovaTabela 
            Height          =   2715
            Left            =   120
            TabIndex        =   14
            Top             =   300
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   4789
            _Version        =   393216
         End
      End
      Begin VB.Frame FraDescrNovaTabela 
         Caption         =   "Descrição da Tabela"
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
         Left            =   -74820
         TabIndex        =   8
         Top             =   4200
         Width           =   6615
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
            Left            =   2220
            TabIndex        =   12
            Top             =   420
            Width           =   4215
         End
         Begin VB.TextBox TxtDescrUsuario 
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
            Left            =   2220
            TabIndex        =   11
            Top             =   960
            Width           =   4215
         End
         Begin VB.OptionButton OptDescrUsuario 
            Caption         =   "Descrição do Usuário"
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
            TabIndex        =   10
            Top             =   1020
            Width           =   1935
         End
         Begin VB.OptionButton OptDescrSistema 
            Caption         =   "Descrição do Sistema"
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
            TabIndex        =   9
            Top             =   480
            Width           =   1935
         End
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
         Height          =   735
         Left            =   -68100
         TabIndex        =   5
         Top             =   4200
         Width           =   3375
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   285
            Left            =   1860
            TabIndex        =   6
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
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
            TabIndex        =   7
            Top             =   345
            Width           =   1680
         End
      End
      Begin VB.CommandButton CmdCadastrarTabela 
         Caption         =   "Cadastrar Nova Tabela"
         Height          =   435
         Left            =   -68100
         TabIndex        =   4
         Top             =   5160
         Width           =   3375
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
         Left            =   300
         TabIndex        =   54
         Top             =   540
         Width           =   4785
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   180
         X2              =   11350
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fase 0: O que você deseja fazer?"
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
         TabIndex        =   35
         Top             =   540
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -63650
         Y1              =   840
         Y2              =   840
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
         Left            =   -74700
         TabIndex        =   33
         Top             =   540
         Width           =   2700
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -63650
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   -74820
         X2              =   -64740
         Y1              =   840
         Y2              =   840
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
         TabIndex        =   15
         Top             =   540
         Width           =   7185
      End
   End
   Begin VB.Label LblTransferencia 
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   7800
      Width           =   11415
   End
End
Attribute VB_Name = "frmCadTabPreco"
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

Private Sub CmdCancelarTodoProcesso_Click()

    If MsgBox("Você tem certeza de que quer Cancelar o Cadastramento? (Todos os Dados serão Perdidos!)", vbYesNo + vbQuestion, "") = vbYes Then
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

Private Sub CmdDefineProximaFaixaPeso_Click()

    If Val(xPesoParte) >= Val(xPesoTodo) Then
    Exit Sub
    Else
    
        If CDbl(TxtPesoInicial.Text) = CDbl(TxtPesoFinal.Text) Then
        MsgBox "Os dois valores são iguais. Por favor, tente novamente...", vbExclamation, ""
        TxtPesoFinal.SetFocus
        Exit Sub
        ElseIf CDbl(TxtPesoFinal.Text) = 0 Then
        MsgBox "O valor do Peso Final não foi informado corretamente. Por favor, tente novamente...", vbExclamation, ""
        TxtPesoFinal.SetFocus
        Exit Sub
        ElseIf CDbl(TxtPesoInicial.Text) = 0 And xPesoParte <> "01" Then
        MsgBox "Valor nulo para o Peso Inicial só é aceito na 1ª Faixa de Peso. Por favor, tente novamente...", vbExclamation, ""
        TxtPesoInicial.SetFocus
        Exit Sub
        ElseIf CDbl(TxtPesoInicial.Text) > CDbl(TxtPesoFinal.Text) Then
        MsgBox "O valor do Peso Inicial é MAIOR que o do Peso Final. Por favor, tente novamente...", vbExclamation, ""
        TxtPesoFinal.SetFocus
        Exit Sub
        End If
        
        
    FlexGridFaixasPeso.Row = xPesoParte
    FlexGridFaixasPeso.Col = 1
    FlexGridFaixasPeso.Text = TxtPesoInicial.Text
    FlexGridFaixasPeso.Col = 2
    FlexGridFaixasPeso.Text = TxtPesoFinal.Text
    
    TxtPesoInicial.Text = Format(((Val(SemPonto(TxtPesoFinal)) + 1) / 10), "###,##0.0")
    
    xPesoTodo = Trim(Str(Val(TxtQteFaixasPeso.Text)))
    xPesoParte = Trim(Str(Val(xPesoParte) + 1))
    
    If Len(Trim(xPesoTodo)) = 1 Then xPesoTodo = "0" & xPesoTodo
    If Len(Trim(xPesoParte)) = 1 Then xPesoParte = "0" & xPesoParte
    
    TxtPesoFinal.Text = "0,0"
    
    LblDefineFaixasPeso.Caption = "Faixa " & xPesoParte & " de " & xPesoTodo
    DoEvents
    
        If Val(xPesoParte) = Val(xPesoTodo) Then
        FlexGridFaixasPeso.Row = xPesoParte
        FlexGridFaixasPeso.Col = 1
        FlexGridFaixasPeso.Text = TxtPesoInicial.Text
                
        LblDefineFaixasPeso.Caption = "Faixa " & xPesoParte & " de " & xPesoTodo
        DoEvents
        End If
    End If
TxtPesoFinal.SetFocus
End Sub

Private Sub CmdFaseAnterior_Click()

Dim xCont As Integer

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
ListLocalidadesSel.Clear

If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
de_informa.Sel_CadLocalAir "%"

    Do Until de_informa.rsSel_CadLocalAir.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade")) & " - " & de_informa.rsSel_CadLocalAir.Fields("SIGLA")
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop
    
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

Private Sub CmdProximaFase_Click()

Dim x, y, xCont As Integer

    If TabFase.Tab = 1 Then
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
        End If
    ElseIf TabFase.Tab = 2 Then
        For y = 1 To FlexGridPanoramaNovaTabela.Rows - 1
        FlexGridPanoramaNovaTabela.Row = y
            For x = 2 To FlexGridPanoramaNovaTabela.Cols - 1
            FlexGridPanoramaNovaTabela.Col = x
                If FlexGridPanoramaNovaTabela.Text = "" Then
                FlexGridPanoramaNovaTabela.Text = "0,00"
                x = x - 1
                ElseIf CDbl(FlexGridPanoramaNovaTabela.Text) = 0 And x >= 2 And x <= 8 Then
                MsgBox "Nenhum valor entre as Colunas 3 e 9 podem ser nulos. Para continuar, corrija este problema.", vbInformation, ""
                FlexGridPanoramaNovaTabela.SetFocus
                    If x = FlexGridPanoramaNovaTabela.Cols - 1 Then
                    SendKeys "{LEFT}"
                    SendKeys "{RIGHT}"
                    Else
                    SendKeys "{RIGHT}"
                    SendKeys "{LEFT}"
                    End If
                    If x = FlexGridPanoramaNovaTabela.Rows - 1 Then
                    SendKeys "{UP}"
                    SendKeys "{DOWN}"
                    Else
                    SendKeys "{DOWN}"
                    SendKeys "{UP}"
                    End If
                Exit Sub
                ElseIf CDbl(FlexGridPanoramaNovaTabela.Text) = 0 And x = 8 + de_informa.rsSel_CadIATA.RecordCount + 1 And CDbl(FlexGridPanoramaNovaTabela.TextMatrix(y, de_informa.rsSel_CadIATA.RecordCount + 2)) > 0 Then
                MsgBox "Você informou um Valor Charter porém não informou qual será o corte de peso. Para continuar, corrija este problema.", vbInformation, ""
                FlexGridPanoramaNovaTabela.SetFocus
                    If x = FlexGridPanoramaNovaTabela.Cols - 1 Then
                    SendKeys "{LEFT}"
                    SendKeys "{RIGHT}"
                    Else
                    SendKeys "{RIGHT}"
                    SendKeys "{LEFT}"
                    End If
                    If x = FlexGridPanoramaNovaTabela.Rows - 1 Then
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
        
        
        If TabFase.Tab = 2 Then
        
        FlexGridPanoramaNovaTabela.Clear
        
        
        
        If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
        de_informa.Sel_Cadiata "%"
                
        FlexGridPanoramaNovaTabela.Cols = 11 + (de_informa.rsSel_CadIATA.RecordCount)
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
        
        FlexGridPanoramaNovaTabela.ColWidth(0) = 2100
        FlexGridPanoramaNovaTabela.ColWidth(1) = 800
        FlexGridPanoramaNovaTabela.ColWidth(2) = 1300
        FlexGridPanoramaNovaTabela.ColWidth(3) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(4) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(5) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(6) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(7) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(8) = 1500
        
                    
        xCont = 9
            Do Until de_informa.rsSel_CadIATA.EOF
                        
            FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Cód. " & de_informa.rsSel_CadIATA.Fields("codigo")
            FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1000
            
            FlexGridPanoramaNovaTabela.Row = 0
            FlexGridPanoramaNovaTabela.Col = xCont
            FlexGridPanoramaNovaTabela.CellAlignment = 3
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            xCont = xCont + 1
            de_informa.rsSel_CadIATA.MoveNext
            Loop
        
        FlexGridPanoramaNovaTabela.Row = 0
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Corte Charter (Kg)"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1600
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        xCont = xCont + 1
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Valor Charter"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
            
            
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
            
        End If
        
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
Dim x, y, xCont As Integer

If MsgBox("Confirma Zerar Nova Tabela?", vbYesNo + vbExclamation, "Confirmação para Zerar Tabela") = vbYes Then
    If MsgBox("ATENÇÃO! Ao zerar a digitação, todo seu trabalho será perdido! Você tem certeza que deseja zerar a tabela?", vbYesNo + vbCritical, "ATENÇÃO! Confirmação para Zerar Tabela") = vbYes Then
        FlexGridPanoramaNovaTabela.Clear
        
        If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
        de_informa.Sel_Cadiata "%"
                
        FlexGridPanoramaNovaTabela.Cols = 11 + (de_informa.rsSel_CadIATA.RecordCount)
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
        
        FlexGridPanoramaNovaTabela.ColWidth(0) = 2100
        FlexGridPanoramaNovaTabela.ColWidth(1) = 800
        FlexGridPanoramaNovaTabela.ColWidth(2) = 1300
        FlexGridPanoramaNovaTabela.ColWidth(3) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(4) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(5) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(6) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(7) = 1500
        FlexGridPanoramaNovaTabela.ColWidth(8) = 1500
        
                    
        xCont = 9
            Do Until de_informa.rsSel_CadIATA.EOF
                        
            FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Cód. " & de_informa.rsSel_CadIATA.Fields("codigo")
            FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1000
            
            FlexGridPanoramaNovaTabela.Row = 0
            FlexGridPanoramaNovaTabela.Col = xCont
            FlexGridPanoramaNovaTabela.CellAlignment = 3
            FlexGridPanoramaNovaTabela.CellFontBold = True
            
            xCont = xCont + 1
            de_informa.rsSel_CadIATA.MoveNext
            Loop
        
        FlexGridPanoramaNovaTabela.Row = 0
        
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Corte Charter (Kg)"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1600
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
        
        xCont = xCont + 1
        FlexGridPanoramaNovaTabela.TextMatrix(0, xCont) = "Valor Charter"
        FlexGridPanoramaNovaTabela.ColWidth(xCont) = 1300
        FlexGridPanoramaNovaTabela.Col = xCont
        FlexGridPanoramaNovaTabela.CellAlignment = 3
        FlexGridPanoramaNovaTabela.CellFontBold = True
            
            
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
        ElseIf KeyAscii = 8 Then
            If FlexGridPanoramaNovaTabela.Col = FlexGridPanoramaNovaTabela.Cols - 2 Then
            FlexGridPanoramaNovaTabela.Text = Mid(FlexGridPanoramaNovaTabela.Text, 1, Len(FlexGridPanoramaNovaTabela.Text) - 1)
            FlexGridPanoramaNovaTabela.Text = Format((SemPonto(FlexGridPanoramaNovaTabela.Text) / 10), "###,##0.0")
            Else
            FlexGridPanoramaNovaTabela.Text = Mid(FlexGridPanoramaNovaTabela.Text, 1, Len(FlexGridPanoramaNovaTabela.Text) - 1)
            FlexGridPanoramaNovaTabela.Text = Format((SemPonto(FlexGridPanoramaNovaTabela.Text) / 100), "###,##0.00")
            End If
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
        If FlexGridPanoramaNovaTabela.Col = FlexGridPanoramaNovaTabela.Cols - 2 Then
        FlexGridPanoramaNovaTabela.Text = FlexGridPanoramaNovaTabela.Text & Chr(KeyAscii)
        FlexGridPanoramaNovaTabela.Text = Format((SemPonto(FlexGridPanoramaNovaTabela.Text) / 10), "###,##0.0")
        Else
        FlexGridPanoramaNovaTabela.Text = FlexGridPanoramaNovaTabela.Text & Chr(KeyAscii)
        FlexGridPanoramaNovaTabela.Text = Format((SemPonto(FlexGridPanoramaNovaTabela.Text) / 100), "###,##0.00")
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

If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
de_informa.Sel_CadLocalAir "%"

    Do Until de_informa.rsSel_CadLocalAir.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade")) & " - " & de_informa.rsSel_CadLocalAir.Fields("SIGLA")
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop
    
Call OrdenaListBox(ListLocalidadesDisponives)


End Sub

Private Sub gridCiaAerea_Click()
TxtSiglaCiaAerea.Text = gridCiaAerea.Columns(0)
TxtNomeCiaAerea.Text = gridCiaAerea.Columns(1)
End Sub

Private Sub gridCiaAerea_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
gridCiaAerea_Click
End Sub

Private Sub ListLocalidadesDisponives_Click()
CmdAdicionaLocalidade.Enabled = True
End Sub

Private Sub ListLocalidadesSel_Click()
CmdRemoveLocalidade.Enabled = True
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

Private Sub TxtPesoInicial_Change()
Call TextPesoBox_Change(TxtPesoInicial)
End Sub

Private Sub TxtPesoInicial_GotFocus()
Call TextPesoBox_GotFocus(TxtPesoInicial)
End Sub

Private Sub TxtPesoInicial_KeyPress(KeyAscii As Integer)
Call TextPesoBox_KeyPress(KeyAscii)
End Sub

Private Sub TxtPesoFinal_Change()
Call TextPesoBox_Change(TxtPesoFinal)
End Sub

Private Sub TxtPesoFinal_GotFocus()
Call TextPesoBox_GotFocus(TxtPesoFinal)
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

Private Sub TxtValorOficial_Change()
Call TextMoneyBox_Change(TxtValorOficial)
End Sub

Private Sub TxtValorOficial_GotFocus()
Call TextMoneyBox_GotFocus(TxtValorOficial)
End Sub

Private Sub TxtValorOficial_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub

Private Sub TxtDESCintec_Change()
Call TextMoneyBox_Change(TxtDescINTEC)
End Sub

Private Sub TxtDESCintec_GotFocus()
Call TextMoneyBox_GotFocus(TxtDescINTEC)
End Sub

Private Sub TxtDESCintec_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub

Private Sub TxtCharter_Change()
Call TextMoneyBox_Change(TxtCharter)
End Sub

Private Sub TxtCharter_GotFocus()
Call TextMoneyBox_GotFocus(TxtCharter)
End Sub

Private Sub TxtCharter_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub

Private Sub TxtValoranvisa_Change()
Call TextMoneyBox_Change(TxtValorANVISA)
End Sub

Private Sub TxtValoranvisa_GotFocus()
Call TextMoneyBox_GotFocus(TxtValorANVISA)
End Sub

Private Sub TxtValoranvisa_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub

