VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_bonagura 
   Caption         =   "BONAGURA - CONTROLE"
   ClientHeight    =   9390
   ClientLeft      =   975
   ClientTop       =   1365
   ClientWidth     =   13860
   ControlBox      =   0   'False
   Icon            =   "frm_bonagura.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   13860
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&SAIR"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Príodo"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "&Buscar"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mas_inicio 
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mas_final 
         Height          =   300
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "até"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Final:"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio:"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "PERÍODO"
      TabPicture(0)   =   "frm_bonagura.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRD_BONA"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "COMPARAÇÕES"
      TabPicture(1)   =   "frm_bonagura.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd_compara"
      Tab(1).Control(1)=   "grd_tb_bona"
      Tab(1).Control(2)=   "grd_fatura_valor"
      Tab(1).Control(3)=   "Line1"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(6)=   "Label4"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "FATURA DIA"
      TabPicture(2)   =   "frm_bonagura.frx":12AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(3)=   "grd_intec"
      Tab(2).Control(4)=   "Line3"
      Tab(2).Control(5)=   "Line2"
      Tab(2).Control(6)=   "lab_total"
      Tab(2).Control(7)=   "Label11"
      Tab(2).Control(8)=   "lab_qtd"
      Tab(2).Control(9)=   "Label8"
      Tab(2).Control(10)=   "Label9"
      Tab(2).Control(11)=   "LAB_DIA"
      Tab(2).Control(12)=   "Label7"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frm_bonagura.frx":12C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame5 
         Height          =   1335
         Left            =   -71160
         TabIndex        =   42
         Top             =   3240
         Width           =   3375
         Begin VB.CommandButton cmd_limpa_bona 
            Caption         =   "Limpa Bona"
            Height          =   255
            Left            =   720
            TabIndex        =   50
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmd_deleta_bona 
            Caption         =   "Deleta"
            Height          =   255
            Left            =   2040
            TabIndex        =   49
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmd_altera_bona 
            Caption         =   "Altera"
            Height          =   255
            Left            =   2040
            TabIndex        =   48
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmd_insere_bona 
            Caption         =   "Inserir"
            Height          =   255
            Left            =   2040
            TabIndex        =   47
            Top             =   240
            Width           =   975
         End
         Begin MSMask.MaskEdBox Mas_valor 
            Height          =   300
            Left            =   720
            TabIndex        =   46
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "###.###,##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mas_data 
            Height          =   300
            Left            =   720
            TabIndex        =   45
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            Caption         =   "Valor:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Data:"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   -71160
         TabIndex        =   36
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txt_filial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            MaxLength       =   2
            TabIndex        =   39
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txt_fatura 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd_pesq_fat 
            Caption         =   "&Pesq"
            Height          =   255
            Left            =   960
            TabIndex        =   37
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "FILIAL:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "FATURA:"
            Height          =   255
            Left            =   1320
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados de Envio"
         Height          =   1455
         Left            =   -71160
         TabIndex        =   24
         Top             =   1560
         Width           =   7935
         Begin VB.Frame Frame4 
            Caption         =   "Enviado"
            Height          =   1215
            Left            =   6480
            TabIndex        =   25
            Top             =   120
            Width           =   1335
            Begin VB.OptionButton opt_s 
               Caption         =   "SIM"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton opt_n 
               Caption         =   "NÃO"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   720
               Width           =   735
            End
         End
         Begin VB.Label Label13 
            Caption         =   "Filial Fatura:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lab_filialfatura 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1080
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Valor:"
            Height          =   255
            Left            =   3600
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lab_valor_fat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   4320
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lab_cli 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1080
            TabIndex        =   30
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label18 
            Caption         =   "Arquivo:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lab_arq_nome 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1080
            TabIndex        =   28
            Top             =   960
            Width           =   1815
         End
      End
      Begin MSDataGridLib.DataGrid grd_intec 
         Bindings        =   "frm_bonagura.frx":12E2
         Height          =   5655
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   9975
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
         DataMember      =   "sel_com_dia"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "filialfatura"
            Caption         =   "filialfatura"
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
            DataField       =   "valorfatura"
            Caption         =   "valorfatura"
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
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grd_compara 
         Bindings        =   "frm_bonagura.frx":12F9
         Height          =   2895
         Left            =   -71760
         TabIndex        =   8
         Top             =   4320
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5106
         _Version        =   393216
         BackColor       =   8388608
         ForeColor       =   65535
         HeadLines       =   1
         RowHeight       =   18
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "sel_compara"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "DATA"
            Caption         =   "DATA"
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
            DataField       =   "VALOR_BONA"
            Caption         =   "VALOR_BONA"
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
            DataField       =   "VALOR_INTEC"
            Caption         =   "VALOR_INTEC"
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
            DataField       =   "DIFERENCA"
            Caption         =   "DIFERENCA"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grd_tb_bona 
         Bindings        =   "frm_bonagura.frx":1310
         Height          =   3135
         Left            =   -74880
         TabIndex        =   9
         Top             =   600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5530
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
         DataMember      =   "sel_bona"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "DATA"
            Caption         =   "DATA"
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
            DataField       =   "VALOR"
            Caption         =   "VALOR"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grd_fatura_valor 
         Bindings        =   "frm_bonagura.frx":1327
         Height          =   3135
         Left            =   -68040
         TabIndex        =   10
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
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
         DataMember      =   "sel_fatura_valor"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "EMISSAO"
            Caption         =   "EMISSAO"
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
            DataField       =   "FATURA"
            Caption         =   "FATURA"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GRD_BONA 
         Bindings        =   "frm_bonagura.frx":133E
         Height          =   6975
         Left            =   480
         TabIndex        =   23
         Top             =   600
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   12303
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
         DataMember      =   "sel_fatura_periodo"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "FATURA"
            Caption         =   "FATURA"
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
            DataField       =   "EMISSAO"
            Caption         =   "EMISSAO"
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
            DataField       =   "VENC"
            Caption         =   "VENC"
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
            DataField       =   "CLIENTE"
            Caption         =   "CLIENTE"
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
            DataField       =   "VLBRUTOICMS"
            Caption         =   "VLBRUTOICMS"
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
            DataField       =   "VLBRUTO"
            Caption         =   "VLBRUTO"
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
            DataField       =   "VALOR"
            Caption         =   "VALOR"
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
            DataField       =   "ENVIADO"
            Caption         =   "ENVIADO"
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
            DataField       =   "ARQUIVO"
            Caption         =   "ARQUIVO"
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
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
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
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Line Line3 
         X1              =   -71400
         X2              =   -61920
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line2 
         X1              =   -71520
         X2              =   -71520
         Y1              =   480
         Y2              =   7560
      End
      Begin VB.Label lab_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72720
         TabIndex        =   22
         Top             =   7200
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Valor Total:"
         Height          =   255
         Left            =   -73800
         TabIndex        =   21
         Top             =   7200
         Width           =   855
      End
      Begin VB.Label lab_qtd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72720
         TabIndex        =   20
         Top             =   6840
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Qtd. Faturas:"
         Height          =   255
         Left            =   -73800
         TabIndex        =   19
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "BANCO INTEC"
         Height          =   255
         Left            =   -74040
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label LAB_DIA 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -74400
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Data:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   -68280
         X2              =   -68280
         Y1              =   600
         Y2              =   3720
      End
      Begin VB.Label Label6 
         Caption         =   "DIAS COM PROBLEMAS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   -69240
         TabIndex        =   15
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "FATURAS - BONAGURA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -73320
         TabIndex        =   14
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "FATURAS - INTEC"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -65880
         TabIndex        =   13
         Top             =   3840
         Width           =   1935
      End
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   375
      Left            =   6360
      TabIndex        =   51
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frm_bonagura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public grid As Integer

Private Sub cmd_altera_bona_Click()

deb_bona.up_bona Mas_valor, mas_data

MsgBox "Dia: " & mas_data & " Alterada", vbInformation, "ALTERAÇÃO"


mas_data.Mask = Empty
mas_data.Text = Empty
mas_data.Mask = "##/##/####"
Mas_valor.Mask = "Empty"
Mas_valor.Mask = "###.###,##"
mas_data.SetFocus



End Sub

Private Sub cmd_buscar_Click()

If deb_bona.rssel_fatura_periodo.State = 1 Then deb_bona.rssel_fatura_periodo.Close
    deb_bona.sel_fatura_periodo mas_inicio, mas_final
    
    If deb_bona.rssel_fatura_periodo.RecordCount < 1 Then
        MsgBox "Não há Faturas", vbInformation, "FAURAS"
        grid = 1
        Exit Sub
    Else
        grid = 0
        GRD_BONA.DataMember = "sel_fatura_periodo"
        GRD_BONA.Refresh
        
        
           
    End If
    
If deb_bona.rssel_fatura_valor.State = 1 Then deb_bona.rssel_fatura_valor.Close
   deb_bona.sel_fatura_valor mas_inicio, mas_final
    
    grd_fatura_valor.DataMember = "sel_fatura_valor"
    grd_fatura_valor.Refresh
    
If deb_bona.rssel_bona.State = 1 Then deb_bona.rssel_bona.Close
    deb_bona.sel_bona mas_inicio, mas_final
    
    grd_tb_bona.DataMember = "sel_bona"
    grd_tb_bona.Refresh
    
    
    
If deb_bona.rssel_compara.State = 1 Then deb_bona.rssel_compara.Close
    deb_bona.sel_compara mas_inicio, mas_final
    
    grd_compara.DataMember = "sel_compara"
    grd_compara.Refresh
    


    
    



End Sub

Private Sub cmd_deleta_bona_Click()

deb_bona.deL_data_bona mas_data

mas_data.Mask = Empty
mas_data.Text = Empty

End Sub

Private Sub cmd_insere_bona_Click()

deb_bona.in_tb_bona mas_data, Mas_valor

mas_data.Mask = Empty
mas_data.Text = Empty
mas_data.Mask = "##/##/####"
Mas_valor.Mask = "Empty"
Mas_valor.Text = Empty
Mas_valor.Mask = "###.###,##"
mas_data.SetFocus


End Sub

Private Sub cmd_limpa_bona_Click()
deb_bona.del_bona

MsgBox "TB_BONA Limpa", vbInformation, "BONAGURA"


End Sub

Private Sub cmd_pesq_fat_Click()
Dim xfilialfatura As String
Dim xdatahora As String

xfilialfatura = txt_filial.Text & String(6 - Len(Trim$(txt_fatura.Text)), "0") & Trim$(txt_fatura.Text)

If deb_bona.rssel_pesq_fatura.State = 1 Then deb_bona.rssel_pesq_fatura.Close
    deb_bona.sel_pesq_fatura xfilialfatura
    
    With deb_bona.rssel_pesq_fatura
    
    If .RecordCount < 1 Then
        MsgBox "Fatura Não localizada", vbInformation, "FATURA - " & xfilialfatura
        Exit Sub
    Else
        lab_filialfatura.Caption = xfilialfatura
        lab_cli.Caption = .Fields("CLIENTE")
        lab_valor_fat.Caption = .Fields("VALOR")
        lab_valor_fat.Caption = Format(.Fields("VALOR"), "#,##0.00")
        If .Fields("ENVIADO") = "S" Then
            opt_s.Value = True
            opt_n.Enabled = False
        Else
            opt_n.Value = True
            opt_s.Value = False
            lab_arq_nome.Caption = "NÃO ENVIADO"
            Exit Sub
        End If
        xdatahora = .Fields("ARQUIVO")
        lab_arq_nome.Caption = "M5FATURA" & Mid(.Fields("ARQUIVO"), 1, 2) & Mid(.Fields("ARQUIVO"), 4, 2) & Mid(.Fields("ARQUIVO"), 12, 2) & Mid(.Fields("ARQUIVO"), 15, 2)
        
        
    End If
    End With
    

End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
grid = 1

End Sub

Private Sub grd_compara_Click()
If grid = 0 Then

If deb_bona.rssel_com_dia.State = 1 Then deb_bona.rssel_com_dia.Close
    deb_bona.sel_com_dia deb_bona.rssel_compara.Fields("DATA")
    
    LAB_DIA.Caption = deb_bona.rssel_compara.Fields("DATA")
    grd_intec.DataMember = "sel_com_dia"
    grd_intec.Refresh
    
    If deb_bona.rssel_qtd_valor.State = 1 Then deb_bona.rssel_qtd_valor.Close
        deb_bona.sel_qtd_valor deb_bona.rssel_compara.Fields("DATA")
        
        lab_qtd.Caption = deb_bona.rssel_qtd_valor.Fields("QTD")
        LAB_TOTAL.Caption = deb_bona.rssel_qtd_valor.Fields("FATURA")
        LAB_TOTAL.Caption = Format(LAB_TOTAL, "#,##0.00")
        
Else

    MsgBox "NÃO DIAS PENDENTES", vbInformation, "SEM FATURAS"
    
End If

If deb_bona.rssel_fatura_periodo.State = 1 Then deb_bona.rssel_fatura_periodo.Close
    deb_bona.sel_fatura_periodo deb_bona.rssel_compara.Fields("DATA"), deb_bona.rssel_compara.Fields("DATA")
    
    If deb_bona.rssel_fatura_periodo.RecordCount < 1 Then
        MsgBox "Não há Faturas", vbInformation, "FAURAS"
        grid = 1
        Exit Sub
    Else
        grid = 0
        GRD_BONA.DataMember = "sel_fatura_periodo"
        GRD_BONA.Refresh
        
        
           
    End If




End Sub

Private Sub grd_intec_Click()
Label19.Caption = deb_bona.rssel_com_dia.Fields("filialfatura")
End Sub
