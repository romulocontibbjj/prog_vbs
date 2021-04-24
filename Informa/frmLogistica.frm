VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogistica 
   Caption         =   "Informações de Logística"
   ClientHeight    =   8775
   ClientLeft      =   -45
   ClientTop       =   465
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   18441
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmLogistica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Pedido Cliente"
      TabPicture(1)   =   "frmLogistica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Plano Separação"
      TabPicture(2)   =   "frmLogistica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Notas Fiscais"
      TabPicture(3)   =   "frmLogistica.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).Control(1)=   "Frame10"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Importação"
      TabPicture(4)   =   "frmLogistica.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame12"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame12 
         Caption         =   "Importação das Interfaces (Logística)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   -72480
         TabIndex        =   68
         Top             =   720
         Width           =   9735
         Begin VB.OptionButton optImpNf 
            Caption         =   "Notas Fiscais"
            Height          =   255
            Left            =   8280
            TabIndex        =   80
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optImpPlano 
            Caption         =   "Plano de Separação"
            Height          =   255
            Left            =   6000
            TabIndex        =   79
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton optImpPed 
            Caption         =   "Pedidos"
            Height          =   255
            Left            =   4440
            TabIndex        =   78
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optImpDest 
            Caption         =   "Cliente Destinatário"
            Height          =   255
            Left            =   2280
            TabIndex        =   77
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton optImpRemet 
            Caption         =   "Cliente Remetente"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Frame Frame6 
            Caption         =   "Status de Importação"
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
            Left            =   6840
            TabIndex        =   73
            Top             =   2400
            Width           =   2655
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Registros Lidos...:"
               Height          =   195
               Left            =   120
               TabIndex        =   75
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lblLidos 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               Height          =   255
               Left            =   1920
               TabIndex        =   74
               Top             =   600
               Width           =   615
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Arquivo Escolhido"
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
            Left            =   6840
            TabIndex        =   72
            Top             =   960
            Width           =   2655
            Begin VB.Label lblArquivo 
               BackColor       =   &H80000009&
               BorderStyle     =   1  'Fixed Single
               Height          =   375
               Left            =   120
               TabIndex        =   81
               Top             =   480
               Width           =   2415
            End
         End
         Begin VB.CommandButton cmdImportar 
            Caption         =   "Importar..."
            Height          =   495
            Left            =   3600
            TabIndex        =   71
            Top             =   4440
            Width           =   2775
         End
         Begin VB.FileListBox File1 
            Height          =   2625
            Left            =   3840
            Pattern         =   "BC*.TXT"
            TabIndex        =   70
            Top             =   1080
            Width           =   2295
         End
         Begin VB.DirListBox Dir1 
            Height          =   2565
            Left            =   240
            TabIndex        =   69
            Top             =   1080
            Width           =   2895
         End
      End
      Begin VB.Frame Frame11 
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
         Height          =   8895
         Left            =   -74880
         TabIndex        =   63
         Top             =   1440
         Width           =   14895
         Begin MSDataGridLib.DataGrid DataGrid10 
            Height          =   8415
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   14843
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
      Begin VB.Frame Frame10 
         Caption         =   "Seleção dos Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   14895
         Begin VB.CommandButton cmdProcNf 
            Caption         =   "Consultar"
            Height          =   495
            Left            =   13560
            TabIndex        =   56
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdBuscaCliNf 
            Caption         =   "?"
            Height          =   255
            Left            =   6240
            TabIndex        =   55
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtClienteNf 
            Height          =   285
            Left            =   4440
            TabIndex        =   54
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdConsultarParaNf 
            Caption         =   "Consulta ..."
            Height          =   495
            Left            =   12240
            TabIndex        =   53
            Top             =   240
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   285
            Left            =   2400
            TabIndex        =   57
            Top             =   360
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
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   285
            Left            =   840
            TabIndex        =   58
            Top             =   360
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
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   3840
            TabIndex        =   62
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Período:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   2160
            TabIndex        =   60
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lblNomCliNf 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   59
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Plano de Separação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Left            =   -74880
         TabIndex        =   51
         Top             =   1440
         Width           =   14895
         Begin MSDataGridLib.DataGrid DataGrid9 
            Height          =   8415
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   14843
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
      Begin VB.Frame Frame2 
         Caption         =   "Seleção dos Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   14895
         Begin VB.CommandButton cmdProcPlano 
            Caption         =   "Consultar"
            Height          =   495
            Left            =   13560
            TabIndex        =   44
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdBuscaCliPlano 
            Caption         =   "?"
            Height          =   255
            Left            =   6240
            TabIndex        =   43
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtClientePlano 
            Height          =   285
            Left            =   4440
            TabIndex        =   42
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdConsultarParaPlano 
            Caption         =   "Consulta ..."
            Height          =   495
            Left            =   12240
            TabIndex        =   41
            Top             =   240
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   285
            Left            =   2400
            TabIndex        =   45
            Top             =   360
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
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   285
            Left            =   840
            TabIndex        =   46
            Top             =   360
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   3840
            TabIndex        =   50
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Período:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   2160
            TabIndex        =   48
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lblNomCliPlano 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   47
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados de Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   120
         TabIndex        =   31
         Top             =   6120
         Width           =   14895
         Begin VB.CommandButton Command13 
            Caption         =   "Consulta Informa..."
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid gridPod 
            Height          =   735
            Left            =   1680
            TabIndex        =   32
            Top             =   3360
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   1296
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
         Begin MSDataGridLib.DataGrid gridOcorr 
            Height          =   1095
            Left            =   1680
            TabIndex        =   33
            Top             =   2160
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   1931
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
         Begin MSDataGridLib.DataGrid gridManifesto 
            Height          =   855
            Left            =   1680
            TabIndex        =   34
            Top             =   1200
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   1508
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
         Begin MSDataGridLib.DataGrid gridCtc 
            Height          =   855
            Left            =   1680
            TabIndex        =   35
            Top             =   240
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   1508
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Conhecimento:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Manifesto:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrências:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   2160
            Width           =   900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Entrega (POD):"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   3360
            Width           =   1080
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   1680
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   1800
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   1680
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line7 
            X1              =   120
            X2              =   1680
            Y1              =   3600
            Y2              =   3600
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados de Logística"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   14895
         Begin MSDataGridLib.DataGrid gridPedido 
            Height          =   1575
            Left            =   1680
            TabIndex        =   25
            Top             =   240
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   2778
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
         Begin MSDataGridLib.DataGrid gridPlano 
            Height          =   855
            Left            =   1680
            TabIndex        =   26
            Top             =   1920
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   1508
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
         Begin MSDataGridLib.DataGrid gridNf 
            Height          =   1335
            Left            =   1680
            TabIndex        =   27
            Top             =   2880
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   2355
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
            Caption         =   "Pedido:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Plano de Separação:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   2040
            Width           =   1500
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nota Fiscal:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   3000
            Width           =   840
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   1680
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   1680
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   1680
            Y1              =   3240
            Y2              =   3240
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Left            =   -74880
         TabIndex        =   22
         Top             =   1440
         Width           =   14895
         Begin MSDataGridLib.DataGrid DataGrid8 
            Height          =   8415
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   14843
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
      Begin VB.Frame Frame7 
         Caption         =   "Seleção dos Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   14895
         Begin VB.CommandButton cmdConsultarParaPed 
            Caption         =   "Consulta ..."
            Height          =   495
            Left            =   12240
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtClientePed 
            Height          =   285
            Left            =   4440
            TabIndex        =   14
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdBuscaCliPed 
            Caption         =   "?"
            Height          =   255
            Left            =   6240
            TabIndex        =   13
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmdProcPed 
            Caption         =   "Consultar"
            Height          =   495
            Left            =   13560
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   2400
            TabIndex        =   18
            Top             =   360
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
            Left            =   840
            TabIndex        =   19
            Top             =   360
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
         Begin VB.Label lblNomCliPed 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   21
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   2160
            TabIndex        =   20
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Período:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   3840
            TabIndex        =   15
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame frame1 
         Caption         =   "Seleção de Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   14895
         Begin VB.OptionButton optWms 
            Caption         =   "Por Pedido WMS"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   530
            Width           =   1695
         End
         Begin VB.CommandButton cmdConsultar 
            Caption         =   "Consultar"
            Height          =   495
            Left            =   13560
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtNumPedNf 
            Height          =   285
            Left            =   12120
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdBuscaCliCons 
            Caption         =   "?"
            Height          =   255
            Left            =   6240
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtClienteCons 
            Height          =   285
            Left            =   4440
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optNf 
            Caption         =   "Por Nota Fiscal"
            Height          =   255
            Left            =   2160
            TabIndex        =   3
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optPedcli 
            Caption         =   "Por Pedido do Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.Label lblNomCliCons 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6720
            TabIndex        =   16
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Núm.Pedido Cliente:"
            Height          =   195
            Left            =   10560
            TabIndex        =   7
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   3840
            TabIndex        =   5
            Top             =   360
            Width           =   525
         End
      End
   End
End
Attribute VB_Name = "frmLogistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImportar_Click()

If optImpRemet = True Then
     Open Dir1.Path & "\" & lblArquivo For Input As #1
     
     Do Until EOF(1)
        Line Input #1, xlinha
        xid = Trim$(Mid$(xlinha, 1, 10))
        xcgc = zeros(Val(Trim$(Mid$(xlinha, 11, 30))), 14)
        xid_logistico = Trim$(Mid$(xlinha, 41, 10))
        xnome = Trim$(Mid$(xlinha, 51, 30))
        
        de_informa.Ins_BomiRemet xid, xcgc, xid_logistico, xnome
        'insere no banco de dados
       
     Loop
     Close #1
     MsgBox "Finalizado !"
ElseIf optImpDest = True Then
     Open Dir1.Path & "\" & lblArquivo For Input As #1

     Do Until EOF(1)
        Line Input #1, xlinha
        xid = Trim$(Mid$(xlinha, 1, 10))
        xcgc = zeros(Val(Trim$(Mid$(xlinha, 11, 30))), 14)
        xid_logistico = Trim$(Mid$(xlinha, 41, 10))
        xcgc_remet = zeros(Val(Trim$(Mid$(xlinha, 51, 30))), 14)
        xnome = Trim$(Mid$(xlinha, 81, 30))
        
        de_informa.Ins_BomiDest xid, xcgc, xid_logistico, xcgc_remet, xnome
       
     Loop
     
     Close #1
     
     MsgBox "Finalizado !"
     
ElseIf optImpPed = True Then
     Open Dir1.Path & "\" & lblArquivo For Input As #1
     
     Do Until EOF(1)
        Line Input #1, xlinha
        xpedidoremet = Trim(Mid$(xlinha, 1, 16))
        xid_remet = Trim$(Mid$(xlinha, 17, 10))
        xid_dest = Trim$(Mid$(xlinha, 27, 10))
        xid_logistico = Trim$(Mid$(xlinha, 37, 10))
        xpedidowms = Trim$(Mid$(xlinha, 47, 16))
        xdataenvped = CDate(Mid$(xlinha, 63, 10))
        xdatagerawms = CDate(Mid$(xlinha, 73, 10))
        xhoraenvped = Mid$(xlinha, 83, 6)
        xhoraenvped = Mid$(xhoraenvped, 1, 2) & ":" & Mid$(xhoraenvped, 3, 2)
        xhoragerawms = Mid$(xlinha, 89, 6)
        xhoragerawms = Mid$(xhoragerawms, 1, 2) & ":" & Mid$(xhoragerawms, 3, 2)
        xcoditem = Trim$(Mid$(xlinha, 95, 15))
        xarqmazem = Trim$(Mid$(xlinha, 110, 4))
        xlote = Trim$(Mid$(xlinha, 114, 15))
        xdataexpira = CDate(Mid$(xlinha, 129, 10))
        xquantidade = Val(Mid$(xlinha, 139, 9))
        xcategped = Trim$(Mid$(xlinha, 148, 2))
        xpedidodest = Trim$(Mid$(xlinha, 150, 16))
        
        de_informa.Ins_BomiPedido xpedidoremet, xid_remet, xid_dest, xid_logistico, xpedidowms, xdataenvped, _
                                  xdatagerawms, xhoraenvped, xhoragerawms, xcoditem, xarmazem, xlote, _
                                  xdataexpira, xquantidade, xcategped, xpedidodest
       
     Loop
     Close #1
     MsgBox "Finalizado !"

ElseIf optImpPlano = True Then
     Open Dir1.Path & "\" & lblArquivo For Input As #1
     
     Do Until EOF(1)
        Line Input #1, xlinha
        
        xnumplano = Val(Mid$(xlinha, 1, 8))
        xid_remet = Trim$(Mid$(xlinha, 9, 10))
        xid_logistico = Trim$(Mid$(xlinha, 19, 10))
        xpedidoremet = Trim$(Mid$(xlinha, 29, 16))
        xdatagera = CDate(Mid$(xlinha, 45, 10))
        xhoragera = Mid$(xlinha, 55, 6)
        xhoragera = Mid$(xhoragera, 1, 2) & ":" & Mid$(xhoragera, 3, 2)
        
        de_informa.Ins_BomiPlanoSep xnumplano, xid_logistico, xid_remet, xpedidoremet, xdatagera, xhoragera
       
     Loop
     Close #1
    MsgBox "ok"
ElseIf optImpNf = True Then
     Open Dir1.Path & "\" & lblArquivo For Input As #1
     
     Do Until EOF(1)
        Line Input #1, xlinha
        xid_dest = Trim$(Mid$(xlinha, 1, 10))
        xid_remet = Trim$(Mid$(xlinha, 11, 10))
        xid_logistico = Trim$(Mid$(xlinha, 21, 10))
        xnumnf = Trim$(Mid$(xlinha, 31, 15))
        xserie = Trim$(Mid$(xlinha, 46, 1))
        xemissao = CDate(Mid$(xlinha, 47, 10))
        xexpedicao = CDate(Mid$(xlinha, 57, 10))
        xhoraemi = Mid$(xlinha, 67, 6)
        xhoraemi = Mid$(xhoraemi, 1, 2) & ":" & Mid$(xhoraemi, 3, 2)
        xhoraexp = Mid$(xlinha, 73, 6)
        xhoraexp = Mid$(xhoraexp, 1, 2) & ":" & Mid$(xhoraexp, 3, 2)
        xcoditem = Trim$(Mid$(xlinha, 79, 15))
        xquantidade = Val(Mid$(xlinha, 94, 9))
        xlote = Trim$(Mid$(xlinha, 103, 15))
        xanvisa = Trim$(Mid$(xlinha, 118, 15))
        xvalorunit = Val(Trim$(Mid$(xlinha, 133, 15)))
        xvalortot = Val(SoNumeros(Mid$(xlinha, 148, 15))) / 100
        xpedidoremet = Trim$(Mid$(xlinha, 163, 16))
        If IsDate((Mid$(xlinha, 179, 10))) Then
            xdatafimpick = CDate(Mid$(xlinha, 179, 10))
        Else
            xdatafimpick = ""
        End If
        If IsDate((Mid$(xlinha, 189, 10))) Then
            xdatafimpack = CDate(Mid$(xlinha, 189, 10))
        Else
            xdatafimpack = ""
        End If
        xhorafimpick = Mid$(xlinha, 199, 6)
        xhorafimpick = Mid$(xhorafimpick, 1, 2) & ":" & Mid$(xhorafimpick, 3, 2)
        xhorafimpack = Mid$(xlinha, 205, 6)
        xhorafimpack = Mid$(xhorafimpack, 1, 2) & ":" & Mid$(xhorafimpack, 3, 2)
        xTransportador = Mid$(xlinha, 219, 2)
        
        de_informa.Ins_BomiNotFis xid_dest, xid_remet, xid_logistico, xnumnf, xserie, xemissao, xexpedicao, _
                                  xhoraemi, xhoraexp, xcoditem, xquantidade, xlote, xanvisa, xvalorunit, xvalortot, _
                                  xpedidoremet ' , xdatafimpick, xdatafimpack, xhorafimpick, xhorafimpack
    Loop
    Close #1

End If

End Sub

Private Sub File1_Click()
    lblArquivo.Caption = File1.FileName
End Sub

Private Sub Form_Load()
    Dir1.Path = "C:\INFORMA\BOMILUFT"
    File1.Path = Dir1.Path
    File1.Refresh
End Sub

Private Sub optImpDest_Click()
    If optImpDest = True Then
        File1.Path = Dir1.Path
        File1.Pattern = "BN*"
        File1.Refresh
    End If
End Sub

Private Sub optImpNf_Click()
    If optImpNf = True Then
        File1.Path = Dir1.Path
        File1.Pattern = "BF*"
        File1.Refresh
    End If
End Sub

Private Sub optImpPed_Click()
    If optImpPed = True Then
        File1.Path = Dir1.Path
        File1.Pattern = "BP*"
        File1.Refresh
    End If
End Sub

Private Sub optImpPlano_Click()
    If optImpPlano = True Then
        File1.Path = Dir1.Path
        File1.Pattern = "BS*"
        File1.Refresh
    End If

End Sub

Private Sub optImpRemet_Click()
    If optImpRemet = True Then
        File1.Path = Dir1.Path
        File1.Pattern = "BC*"
        File1.Refresh
    End If
End Sub

